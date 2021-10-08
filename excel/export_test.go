package excel

import "testing"

type Record struct {
	Name        string `xlsx:"A-姓名"`
	Age         int    `xlsx:"B-年龄"`
	IgnoreField int
	Score       float64 `xlsx:"C-分数"`
}

func TestExportExcel(t *testing.T) {
	records := make([]*Record, 0)
	records = append(records, &Record{
		Name:        "Frank",
		Age:         21,
		IgnoreField: 1,
		Score:       98.2,
	})
	records = append(records, &Record{
		Name:        "Alice",
		Age:         22,
		IgnoreField: 2,
		Score:       61.3,
	})
	ExportExcel(records, "test.xlsx")
}
