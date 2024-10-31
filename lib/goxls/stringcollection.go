package goxls

// stringCollection ...
type StringCollection struct {
	StringGrid   [][]string
	StringMap    map[string]int
	StringList   []string
	StringTotal  int
	StringUnique int
}

func (sc *StringCollection) AddRow(row []string) {
	sc.StringGrid = append(sc.StringGrid, row)
	for _, str := range row {
		strToSave := Utf8toBIFF8UnicodeLong(str)
		if _, ok := sc.StringMap[strToSave]; !ok {
			sc.StringMap[strToSave] = sc.StringUnique
			sc.StringList = append(sc.StringList, strToSave)
			sc.StringUnique++
		}

		sc.StringTotal++
	}
}
