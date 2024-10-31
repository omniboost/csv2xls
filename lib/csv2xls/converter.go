package csv2xls

import (
	"bufio"
	"bytes"
	"encoding/csv"
	"errors"
	"fmt"
	"io"
	"math"
	"os"
	"strconv"
	"strings"
	"time"
	"unicode/utf8"

	"github.com/omniboost/csv2xls/lib/goxls"
)

const (
	olePpsTypeRoot   = 5
	olePpsTypeDir    = 1
	olePpsTypeFile   = 2
	oleDataSizeSmall = 0x1000
	oleLongIntSize   = 4
	olePpsSize       = 0x80
)

// Csv2XlsConverter ...
type Csv2XlsConverter struct {
	csvFileName    string
	xlsFileName    string
	csvDelimiter   rune
	title          string
	subject        string
	creator        string
	keywords       string
	description    string
	lastModifiedBy string
}

type dataSectionItem struct {
	summary    uint32
	offset     uint32
	sType      uint32
	dataInt    uint32
	dataString string
	dataLength uint32
}

// NewCsv2XlsConverter ...
func NewCsv2XlsConverter(csvFileName string, xlsFileName string, csvDelimiter string) (*Csv2XlsConverter, error) {
	if utf8.RuneCountInString(csvDelimiter) > 1 {
		return nil, errors.New("csv delimiter must be one character string")
	}

	csvDelimiterDecoded, _ := utf8.DecodeRuneInString(csvDelimiter)

	return &Csv2XlsConverter{
		csvFileName:  csvFileName,
		xlsFileName:  xlsFileName,
		csvDelimiter: csvDelimiterDecoded,
	}, nil
}

// Convert....
func (c *Csv2XlsConverter) Convert() error {
	sc, err := GetStringCollectionFromCSVFile(c.csvFileName, c.csvDelimiter)
	if err != nil {
		return err
	}

	buf, err := c.FromStringCollectionToXLS(&sc)
	if err != nil {
		return err
	}

	f, err := os.Create(c.xlsFileName)
	if err != nil {
		return err
	}
	defer f.Close()

	// Issue a `Sync` to flush writes to stable storage.
	err3 := f.Sync()
	if err3 != nil {
		return err
	}

	w := bufio.NewWriter(f)

	_, err2 := w.Write(buf)
	if err2 != nil {
		return err
	}

	// Use `Flush` to ensure all buffered operations have
	err4 := w.Flush()
	if err4 != nil {
		return err
	}

	return nil
}

// From CSV Reader to XLS ...
func (c *Csv2XlsConverter) FromStringCollectionToXLS(stringCollection *goxls.StringCollection) ([]byte, error) {
	var CreatedAtInt int64 = time.Now().Unix()
	var ModifiedAtInt int64 = time.Now().Unix()

	columnWidths := make(map[int]int, 0)
	//columnWidths[1] = 40 // parameter todo

	wsArr := make([]goxls.Worksheet, 0)
	n := 0
	for i := 0; i < len(stringCollection.StringGrid); i += 65535 {
		wsName := "worksheet"
		if n > 0 {
			wsName += strconv.Itoa(n)
		}

		last := i + 65535
		if last > len(stringCollection.StringGrid) {
			last = len(stringCollection.StringGrid)
		}
		wsArr = append(wsArr, goxls.Worksheet{
			Name:         wsName,
			Grid:         stringCollection.StringGrid[i:last],
			ColumnWidths: columnWidths,
		})
		n++
	}

	worksheetDatas := make([]string, 0)
	worksheetNames := make([]string, 0)
	for _, ws := range wsArr {
		worksheetDatas = append(worksheetDatas, ws.GetData(stringCollection))
		worksheetNames = append(worksheetNames, ws.Name)
	}

	worksheetSizes := make([]int, 0)
	for _, wsd := range worksheetDatas {
		worksheetSizes = append(worksheetSizes, len(wsd))
	}

	workbook := goxls.Workbook{
		WorksheetSizes:   worksheetSizes,
		WorksheetNames:   worksheetNames,
		StringCollection: stringCollection,
	}

	var data strings.Builder
	data.WriteString(workbook.GetWorksheetSizesData())

	for _, wsd := range worksheetDatas {
		data.WriteString(wsd)
	}

	rootPps := goxls.PPS{
		No:         0,
		Name:       goxls.AsciiToUcs("Root Entry"),
		PpsType:    olePpsTypeRoot,
		PrevPps:    0xFFFFFFFF,
		NextPps:    0xFFFFFFFF,
		DirPps:     1,
		Data:       "",
		Size:       0,
		StartBlock: 0,
	}

	workbookPps := goxls.PPS{
		No:         1,
		Name:       goxls.AsciiToUcs("workbook"),
		PpsType:    olePpsTypeFile,
		PrevPps:    2,
		NextPps:    3,
		DirPps:     0xFFFFFFFF,
		Data:       data.String(),
		Size:       0,
		StartBlock: 0,
	}

	// TODO
	//documentSummaryInformationPps := pps{2, fmt.Sprintf("%c%s", rune(5), ascToUcs("DocumentSummaryInformation")), olePpsTypeFile, 0xFFFFFFFF, 0xFFFFFFFF, 0xFFFFFFFF, getDocumentSummaryInformation(), 0, 0}

	summaryInformation := getSummaryInformation(c.title, c.subject, c.creator, c.keywords, c.description, c.lastModifiedBy, CreatedAtInt, ModifiedAtInt)
	summaryInformationPps := goxls.PPS{
		No:         2,
		Name:       goxls.AsciiToUcs(fmt.Sprintf("%c%s", rune(5), "SummaryInformation")),
		PpsType:    olePpsTypeFile,
		PrevPps:    0xFFFFFFFF,
		NextPps:    0xFFFFFFFF,
		DirPps:     0xFFFFFFFF,
		Data:       summaryInformation,
		Size:       0,
		StartBlock: 0,
	}

	aList := []goxls.PPS{rootPps, workbookPps /*, TODO documentSummaryInformationPps*/, summaryInformationPps}

	iSBDcnt, iBBcnt, iPPScnt := calcSize(aList) // change types to uint32 TODO

	// Content of this buffer is result xls file
	resultBuffer := new(bytes.Buffer)

	saveHeader(resultBuffer, iSBDcnt, iBBcnt, iPPScnt)

	smallData := makeSmallData(resultBuffer, aList)
	aList[0].Data = smallData

	// Write BB
	saveBigData(resultBuffer, iSBDcnt, aList)

	// Write PPS
	savePps(resultBuffer, aList)

	// Write Big Block Depot and BDList and Adding Header information
	saveBbd(resultBuffer, iSBDcnt, iBBcnt, iPPScnt)

	return resultBuffer.Bytes(), nil
}

// WithTitle ...
func (c *Csv2XlsConverter) WithTitle(title string) *Csv2XlsConverter {
	c.title = title
	return c
}

// WithSubject ...
func (c *Csv2XlsConverter) WithSubject(subject string) *Csv2XlsConverter {
	c.subject = subject
	return c
}

// WithCreator ...
func (c *Csv2XlsConverter) WithCreator(creator string) *Csv2XlsConverter {
	c.creator = creator
	return c
}

// WithKeywords ...
func (c *Csv2XlsConverter) WithKeywords(keywords string) *Csv2XlsConverter {
	c.keywords = keywords
	return c
}

// WithDescription ...
func (c *Csv2XlsConverter) WithDescription(description string) *Csv2XlsConverter {
	c.description = description
	return c
}

// WithLastModifiedBy ...
func (c *Csv2XlsConverter) WithLastModifiedBy(lastModifiedBy string) *Csv2XlsConverter {
	c.lastModifiedBy = lastModifiedBy
	return c
}

func saveBbd(buffer *bytes.Buffer, iSbdSize, iBsize, iPpsCnt uint32) {
	// Calculate Basic Setting
	var iBbCnt uint32 = 512 / oleLongIntSize
	var i1stBdL uint32 = (512 - 0x4C) / oleLongIntSize

	var iBdExL uint32 = 0
	iAll := iBsize + iPpsCnt + iSbdSize
	iAllW := iAll
	iBdCntW := uint32(math.Floor(float64(iAllW) / float64(iBbCnt)))
	if iAllW%iBbCnt > 0 {
		iBdCntW++
	}
	iBdCnt := uint32(math.Floor(float64(iAll+iBdCntW) / float64(iBbCnt)))
	if (iAllW+iBdCntW)%iBbCnt > 0 {
		iBdCnt++
	}
	// Calculate BD count
	if iBdCnt > i1stBdL {
		for {
			iBdExL++
			iAllW++
			iBdCntW = uint32(math.Floor(float64(iAllW) / float64(iBbCnt)))
			if iAllW%iBbCnt > 0 {
				iBdCntW++
			}
			iBdCnt = uint32(math.Floor(float64(iAllW+iBdCntW) / float64(iBbCnt)))
			if (iAllW+iBdCntW)%iBbCnt > 0 {
				iBdCnt++
			}
			if iBdCnt <= (iBdExL*iBbCnt + i1stBdL) {
				break
			}
		}
	}

	// Making BD
	// Set for SBD
	if iSbdSize > 0 {
		var i uint32
		for i = 0; i < (iSbdSize - 1); i++ {
			goxls.PutVar(buffer, i+1)
		}
		goxls.PutVar(buffer, []byte("\xFE\xFF\xFF\xFF")) // uint32(-2)
	}

	// Set for B
	var i uint32
	for i = 0; i < (iBsize - 1); i++ {
		goxls.PutVar(buffer, i+iSbdSize+1)
	}
	goxls.PutVar(buffer, []byte("\xFE\xFF\xFF\xFF"))

	// Set for PPS
	for i = 0; i < (iPpsCnt - 1); i++ {
		goxls.PutVar(buffer, i+iSbdSize+iBsize+1)
	}
	goxls.PutVar(buffer, []byte("\xFE\xFF\xFF\xFF"))

	// Set for BBD itself ( 0xFFFFFFFD : BBD)
	for i = 0; i < iBdCnt; i++ {
		goxls.PutVar(buffer, uint32(0xFFFFFFFD))
	}

	// Set for ExtraBDList
	for i = 0; i < iBdExL; i++ {
		goxls.PutVar(buffer, uint32(0xFFFFFFFC))
	}

	// Adjust for Block
	if (iAllW+iBdCnt)%iBbCnt > 0 {
		iBlock := iBbCnt - ((iAllW + iBdCnt) % iBbCnt)
		for i = 0; i < iBlock; i++ {
			goxls.PutVar(buffer, []byte("\xFF\xFF\xFF\xFF"))
		}
	}

	// Extra BDList
	if iBdCnt > i1stBdL {
		var iN, iNb uint32
		for i = i1stBdL; i < iBdCnt; i++ {
			if iN >= (iBbCnt - 1) {
				iN = 0
				iNb++
				goxls.PutVar(buffer, iAll+iBdCnt+iNb)
			}
			goxls.PutVar(buffer, iBsize+iSbdSize+iPpsCnt+i)
			iN++
		}
		if (iBdCnt-i1stBdL)%(iBbCnt-1) > 0 {
			iB := (iBbCnt - 1) - ((iBdCnt - i1stBdL) % (iBbCnt - 1))
			for i = 0; i < iB; i++ {
				goxls.PutVar(buffer, []byte("\xFF\xFF\xFF\xFF"))
			}
		}
		goxls.PutVar(buffer, []byte("\xFE\xFF\xFF\xFF"))
	}
}

func savePps(buffer *bytes.Buffer, raList []goxls.PPS) {
	// Save each PPS WK
	for _, pps := range raList {
		goxls.PutVar(buffer, []byte(pps.GetPpsWk())) // maybe it'll be better to change return type to []byte
	}
	// Adjust for Block
	iCnt := len(raList)
	iBCnt := 512 / olePpsSize
	if iCnt%iBCnt > 0 {
		goxls.PutVar(buffer, []byte(strings.Repeat("\x00", (iBCnt-(iCnt%iBCnt))*olePpsSize)))
	}
}

func saveBigData(buffer *bytes.Buffer, iStBlk uint32, raList []goxls.PPS) {
	// cycle through PPS's
	for i := range raList {
		if raList[i].PpsType != olePpsTypeDir {
			raList[i].Size = uint32(len(raList[i].Data))
			if raList[i].Size >= oleDataSizeSmall || (raList[i].PpsType == olePpsTypeRoot && len(raList[i].Data) != 0) {
				goxls.PutVar(buffer, []byte(raList[i].Data))

				if raList[i].Size%512 > 0 {
					goxls.PutVar(buffer, []byte(strings.Repeat("\x00", 512-int(raList[i].Size)%512)))
				}
				// Set For PPS
				raList[i].StartBlock = iStBlk
				iStBlk += uint32(math.Floor(float64(raList[i].Size) / 512))
				if raList[i].Size%512 > 0 {
					iStBlk++
				}
			}
		}
	}
}

func makeSmallData(buffer *bytes.Buffer, raList []goxls.PPS) string {
	var smallData strings.Builder
	var iSmBlk uint32 = 0

	for i := range raList {
		// Make SBD, small data string
		if raList[i].PpsType == olePpsTypeFile {
			if raList[i].Size <= 0 {
				continue
			}

			if raList[i].Size < oleDataSizeSmall {
				iSmbCnt := uint32(math.Floor(float64(raList[i].Size) / 64))
				if raList[i].Size%64 > 0 {
					iSmbCnt++
				}
				jB := iSmbCnt - 1
				var j uint32
				for j = 0; j < jB; j++ {
					goxls.PutVar(buffer, j+iSmBlk+1)
				}
				goxls.PutVar(buffer, []byte("\xFE\xFF\xFF\xFF")) // uint32(-2)

				smallData.WriteString(raList[i].Data)
				if raList[i].Size%64 > 0 {
					smallData.WriteString(strings.Repeat("\x00", 64-int(raList[i].Size%64)))
				}
				// Set for PPS
				raList[i].StartBlock = iSmBlk
				iSmBlk += iSmbCnt
			}
		}
	}

	iSbCnt := uint32(math.Floor(512.0 / oleLongIntSize))
	if iSmBlk%iSbCnt > 0 {
		iB := iSbCnt - (iSmBlk % iSbCnt)
		var i uint32
		for i = 0; i < iB; i++ {
			goxls.PutVar(buffer, []byte("\xFF\xFF\xFF\xFF"))
		}
	}

	return smallData.String()
}

func saveHeader(buffer *bytes.Buffer, iSBDcnt, iBBcnt, iPPScnt uint32) {
	// Calculate Basic Setting
	var iBlCnt uint32 = 512 / oleLongIntSize
	var i1stBdL uint32 = (512 - 0x4C) / oleLongIntSize

	var iBdExL uint32 = 0
	iAll := uint32(iBBcnt + iPPScnt + iSBDcnt)
	iAllW := iAll
	iBdCntW := uint32(math.Floor(float64(iAllW) / float64(iBlCnt)))
	if iAllW%iBlCnt > 0 {
		iBdCntW++
	}
	iBdCnt := uint32(math.Floor(float64(iAll+iBdCntW) / float64(iBlCnt)))
	if (iAllW+iBdCntW)%iBlCnt > 0 {
		iBdCnt++
	}

	// Calculate BD count
	if iBdCnt > i1stBdL {
		for {
			iBdExL++
			iAllW++
			iBdCntW = uint32(math.Floor(float64(iAllW) / float64(iBlCnt)))
			if iAllW%iBlCnt > 0 {
				iBdCntW++
			}
			iBdCnt = uint32(math.Floor(float64(iAllW+iBdCntW) / float64(iBlCnt)))
			if (iAllW+iBdCntW)%iBlCnt > 0 {
				iBdCnt++
			}
			if iBdCnt <= (iBdExL*iBlCnt + i1stBdL) {
				break
			}
		}
	}

	// Save Header
	goxls.PutVar(buffer,
		[]byte("\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"),
		[]byte("\x00\x00\x00\x00"),
		[]byte("\x00\x00\x00\x00"),
		[]byte("\x00\x00\x00\x00"),
		[]byte("\x00\x00\x00\x00"),
		uint16(0x3b),
		uint16(0x03),
		[]byte("\xFE\xFF"), // uint16(-2),
		uint16(9),
		uint16(6),
		uint16(0),
		[]byte("\x00\x00\x00\x00"),
		[]byte("\x00\x00\x00\x00"),
		iBdCnt,
		iBBcnt+iSBDcnt,
		uint32(0),
		uint32(0x1000),
	)
	if iSBDcnt > 0 {
		goxls.PutVar(buffer, uint32(0))
	} else {
		goxls.PutVar(buffer, []byte("\xFE\xFF\xFF\xFF"))
	}
	goxls.PutVar(buffer, iSBDcnt)

	// Extra BDList Start, Count
	if iBdCnt < i1stBdL {
		goxls.PutVar(buffer,
			[]byte("\xFE\xFF\xFF\xFF"), // Extra BDList Start
			uint32(0),                  // Extra BDList Count
		)
	} else {
		goxls.PutVar(buffer, iAll+iBdCnt, iBdExL)
	}

	// BDList
	var i uint32
	for i = 0; i < i1stBdL && i < iBdCnt; i++ {
		goxls.PutVar(buffer, iAll+i)
	}
	if i < i1stBdL {
		jB := i1stBdL - i
		var j uint32
		for j = 0; j < jB; j++ {
			goxls.PutVar(buffer, []byte("\xFF\xFF\xFF\xFF"))
		}
	}
}

func calcSize(aList []goxls.PPS) (uint32, uint32, uint32) {
	var iSBDcnt, iBBcnt, iPPScnt uint32 = 0, 0, 0

	iSBcnt := 0
	iCount := len(aList)
	for i := 0; i < iCount; i++ {
		if aList[i].PpsType == olePpsTypeFile {
			aList[i].Size = uint32(len(aList[i].Data))

			if aList[i].Size < oleDataSizeSmall {
				iSBcnt += int(math.Floor(float64(aList[i].Size) / 64))
				if aList[i].Size%64 > 0 {
					iSBcnt++
				}
			} else {
				iBBcnt += uint32(math.Floor(float64(aList[i].Size) / 512))
				if aList[i].Size%512 > 0 {
					iBBcnt++
				}
			}
		}
	}

	iSlCnt := int(math.Floor(512 / oleLongIntSize))
	if (math.Floor(float64(iSBcnt)/float64(iSlCnt)) + float64(iSBcnt%iSlCnt)) > 0 {
		iSBDcnt = 1
	}

	iSmallLen := float64(iSBcnt) * 64
	iBBcnt += uint32(math.Floor(iSmallLen / 512))
	if int(iSmallLen)%512 > 0 {
		iBBcnt++
	}
	iCnt := len(aList)
	iBdCnt := float64(512) / olePpsSize
	iPPScnt = uint32(math.Floor(float64(iCnt) / iBdCnt))
	if iCnt%int(iBdCnt) > 0 {
		iPPScnt++
	}

	return iSBDcnt, iBBcnt, iPPScnt
}

func getSummaryInformation(title, subject, creator, keywords, description, lastModifiedBy string, created, modified int64) string {
	buffer := new(bytes.Buffer)

	// offset: 0; size: 2; must be 0xFE 0xFF (UTF-16 LE byte order mark)
	goxls.PutVar(buffer, uint16(0xFFFE))
	// offset: 2; size: 2;
	goxls.PutVar(buffer, uint16(0x0000))
	// offset: 4; size: 2; OS version
	goxls.PutVar(buffer, uint16(0x0106))
	// offset: 6; size: 2; OS indicator
	goxls.PutVar(buffer, uint16(0x0002))
	// offset: 8; size: 16
	goxls.PutVar(buffer, uint32(0x00), uint32(0x00), uint32(0x00), uint32(0x00))
	// offset: 24; size: 4; section count
	goxls.PutVar(buffer, uint32(0x0001))

	// offset: 28; size: 16; first section's class id: 02 d5 cd d5 9c 2e 1b 10 93 97 08 00 2b 2c f9 ae
	goxls.PutVar(buffer, uint16(0x85E0), uint16(0xF29F), uint16(0x4FF9), uint16(0x1068), uint16(0x91AB), uint16(0x0008), uint16(0x272B), uint16(0xD9B3))
	// offset: 44; size: 4; offset of the start
	goxls.PutVar(buffer, uint32(0x30))

	var dataSectionNumProps uint32 = 0
	dataSections := make([]dataSectionItem, 0)

	// CodePage : CP-1252
	dataSections = append(dataSections, dataSectionItem{0x01, 0, 0x02, 1252, "", 0})
	dataSectionNumProps++

	// Title
	if title != "" {
		dataSections = append(dataSections, dataSectionItem{0x02, 0, 0x1E, 0, title, uint32(len(title))})
		dataSectionNumProps++
	}

	// Subject
	if subject != "" {
		dataSections = append(dataSections, dataSectionItem{0x03, 0, 0x1E, 0, subject, uint32(len(subject))})
		dataSectionNumProps++
	}

	// Author (Creator)
	if creator != "" {
		dataSections = append(dataSections, dataSectionItem{0x04, 0, 0x1E, 0, creator, uint32(len(creator))})
		dataSectionNumProps++
	}

	// Keywords
	if keywords != "" {
		dataSections = append(dataSections, dataSectionItem{0x05, 0, 0x1E, 0, keywords, uint32(len(keywords))})
		dataSectionNumProps++
	}

	// Comments (Description)
	if description != "" {
		dataSections = append(dataSections, dataSectionItem{0x06, 0, 0x1E, 0, description, uint32(len(description))})
		dataSectionNumProps++
	}

	// Last Saved By (LastModifiedBy)
	if lastModifiedBy != "" {
		dataSections = append(dataSections, dataSectionItem{0x08, 0, 0x1E, 0, lastModifiedBy, uint32(len(lastModifiedBy))})
		dataSectionNumProps++
	}

	// Created Date/Time
	if created != 0 {
		dataSections = append(dataSections, dataSectionItem{0x0C, 0, 0x40, 0, goxls.LocalDateToOLE(created), 0})
		dataSectionNumProps++
	}

	// Modified Date/Time
	if modified != 0 {
		dataSections = append(dataSections, dataSectionItem{0x0D, 0, 0x40, 0, goxls.LocalDateToOLE(modified), 0})
		dataSectionNumProps++
	}

	// Security
	dataSections = append(dataSections, dataSectionItem{0x13, 0, 0x03, 0x00, "", 0})
	dataSectionNumProps++

	dataSectionSummary := new(bytes.Buffer)
	dataSectionContent := new(bytes.Buffer)
	dataSectionContentOffset := 8 + dataSectionNumProps*8

	for _, dataSection := range dataSections {
		// Summary
		goxls.PutVar(dataSectionSummary, dataSection.summary)
		// Offset
		goxls.PutVar(dataSectionSummary, dataSectionContentOffset)
		// DataType
		goxls.PutVar(dataSectionContent, dataSection.sType)
		// Data
		if dataSection.sType == 0x02 { // 2 byte signed integer
			goxls.PutVar(dataSectionContent, dataSection.dataInt)
			dataSectionContentOffset += 8
		} else if dataSection.sType == 0x03 { // 4 byte signed integer
			goxls.PutVar(dataSectionContent, dataSection.dataInt)
			dataSectionContentOffset += 8
		} else if dataSection.sType == 0x1E { // null-terminated string prepended by dword string length
			// Null-terminated string
			dataSection.dataString += "\x00"
			dataSection.dataLength++

			// Complete the string with null string for being a %4
			if (4 - dataSection.dataLength%4) != 4 {
				dataSection.dataLength += 4 - dataSection.dataLength%4
			}

			dataSection.dataString = dataSection.dataString + strings.Repeat("\x00", int(dataSection.dataLength)-len(dataSection.dataString))

			goxls.PutVar(dataSectionContent, dataSection.dataLength)
			goxls.PutVar(dataSectionContent, []byte(dataSection.dataString))

			dataSectionContentOffset += 8 + uint32(len(dataSection.dataString))
		} else if dataSection.sType == 0x40 { // Filetime (64-bit value representing the number of 100-nanosecond intervals since January 1, 1601)
			goxls.PutVar(dataSectionContent, []byte(dataSection.dataString))
			dataSectionContentOffset += 4 + 8
		}
		// Data Type Not Used at the moment
	}
	// Now dataSectionContentOffset contains the size of the content

	// section header
	// offset: $secOffset; size: 4; section length
	//         + x  Size of the content (summary + content)
	goxls.PutVar(buffer, dataSectionContentOffset)

	// offset: $secOffset+4; size: 4; property count
	goxls.PutVar(buffer, dataSectionNumProps)

	// Section Summary
	goxls.PutVar(buffer, dataSectionSummary.Bytes())

	// Section Content
	goxls.PutVar(buffer, dataSectionContent.Bytes())

	return buffer.String()
}

func GetStringCollectionFromCSVFile(csvFileName string, delimiter rune) (goxls.StringCollection, error) {
	f, err := os.Open(csvFileName)
	if err != nil {
		return goxls.StringCollection{}, fmt.Errorf(`cannot read csv file "%s"`, csvFileName)
	}
	defer f.Close()

	sc, err := GetStringCollectionFromCSVReader(f, delimiter)
	return sc, err
}

func GetStringCollectionFromCSVReader(reader io.Reader, delimiter rune) (goxls.StringCollection, error) {
	sc := goxls.StringCollection{
		StringGrid:   make([][]string, 0),
		StringMap:    make(map[string]int, 0),
		StringList:   make([]string, 0),
		StringTotal:  0,
		StringUnique: 0,
	}

	r := csv.NewReader(reader)
	r.FieldsPerRecord = -1
	r.Comma = delimiter
	r.LazyQuotes = true

	for {
		record, err := r.Read()
		// Stop at EOF.
		if err == io.EOF {
			break
		}

		if err != nil {
			return sc, err
		}

		sc.AddRow(record)
	}

	return sc, nil
}
