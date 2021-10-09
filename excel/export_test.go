package excel

import (
	"testing"
	"time"
)

type Record struct {
	Name        string `xlsx:"A-姓名"`
	Age         int    `xlsx:"B-年龄"`
	IgnoreField int
	Score       float64   `xlsx:"C-分数"`
	UpdatedAt   time.Time `xlsx:"D-更新时间"`
	TestPointer *int      `xlsx:"E-测试"`
}

func TestExportExcel(t *testing.T) {
	records := make([]*Record, 0)
	i := 1
	records = append(records, &Record{
		Name:        "Frank",
		Age:         21,
		IgnoreField: 1,
		Score:       98.2,
		UpdatedAt:   time.Now(),
		TestPointer: &i,
	})
	j := 2
	records = append(records, &Record{
		Name:        "Alice",
		Age:         22,
		IgnoreField: 2,
		Score:       61.3,
		UpdatedAt:   time.Now(),
		TestPointer: &j,
	})
	ExportExcelFromSlice(records, &ExportConfig{
		Mode:       ExportTaggedField,
		OutputPath: "TestExcel.xlsx",
	})

	ExportExcelFromSlice(records, &ExportConfig{
		Mode:       ExportAllField,
		OutputPath: "TestExcel2.xlsx",
	})

	ExportExcelFromSlice(records, &ExportConfig{
		Mode:       ExportByHeaders,
		OutputPath: "TestExcel3.xlsx",
		Headers:    []string{"姓名", "分数"},
		Header2FieldMap: map[string]string{
			"姓名": "Name",
			"分数": "Score",
		},
	})
}

func Test_convertToTitle(t *testing.T) {
	type args struct {
		columnNumber int
	}
	tests := []struct {
		name string
		args args
		want string
	}{
		{name: "1", args: args{columnNumber: 1}, want: "A"},
		{name: "27", args: args{columnNumber: 27}, want: "AA"},
		{name: "2", args: args{columnNumber: 2}, want: "B"},
		{name: "28", args: args{columnNumber: 28}, want: "AB"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := convertToTitle(tt.args.columnNumber); got != tt.want {
				t.Errorf("convertToTitle() = %v, want %v", got, tt.want)
			}
		})
	}
}
