package excel

import (
	"os"
	"testing"
	"time"
)

type Record struct {
	Name         string `xlsx:"A-姓名"`
	Age          int    `xlsx:"B-年龄"`
	IgnoredField int
	Score        float64   `xlsx:"C-分数"`
	UpdatedAt    time.Time `xlsx:"D-更新时间"`
	TestPointer  *int      `xlsx:"E-测试"`
	SubRecords   []*SubRecord
	Remark       string `xlsx:"H-备注"`
}

type SubRecord struct {
	QuestionNum int    `xlsx:"F-题号"`
	Answer      string `xlsx:"G-答案"`
}

func TestExportExcel(t *testing.T) {
	subRecords := make([]*SubRecord, 0)
	subRecords = append(subRecords, &SubRecord{
		QuestionNum: 1,
		Answer:      "A",
	})
	subRecords = append(subRecords, &SubRecord{
		QuestionNum: 2,
		Answer:      "B",
	})
	subRecords = append(subRecords, &SubRecord{
		QuestionNum: 3,
		Answer:      "B",
	})
	records := make([]*Record, 0)
	i := 1
	records = append(records, &Record{
		Name:         "Frank",
		Age:          21,
		IgnoredField: 1,
		Score:        98.2,
		UpdatedAt:    time.Now(),
		TestPointer:  &i,
		SubRecords:   subRecords,
		Remark:       "Good job!",
	})

	subRecordsB := make([]*SubRecord, 0)
	subRecordsB = append(subRecordsB, &SubRecord{
		QuestionNum: 1,
		Answer:      "C",
	})
	subRecordsB = append(subRecordsB, &SubRecord{
		QuestionNum: 2,
		Answer:      "A",
	})
	subRecordsB = append(subRecordsB, &SubRecord{
		QuestionNum: 3,
		Answer:      "B",
	})
	j := 2
	records = append(records, &Record{
		Name:         "Alice",
		Age:          22,
		IgnoredField: 2,
		Score:        61.3,
		UpdatedAt:    time.Now(),
		TestPointer:  &j,
		SubRecords:   subRecordsB,
		Remark:       "Keep it up!",
	})

	err := ExportExcelFromSlice(records, &ExportConfig{
		Mode:              ExportTaggedField,
		OutputPath:        "TestExcel.xlsx",
		SubSliceFieldName: "SubRecords",
	})
	if err != nil {
		panic(err)
	}

	file, err := os.OpenFile("TestExcel22.xlsx", os.O_WRONLY|os.O_TRUNC|os.O_CREATE, 0666)
	if err != nil {
		panic(err)
	}
	defer file.Close()
	err = ExportExcelFromSlice(records, &ExportConfig{
		Mode: ExportAllField,
		//OutputPath:        "TestExcel2.xlsx",
		OutputWriter:      file,
		SubSliceFieldName: "SubRecords",
	})
	if err != nil {
		panic(err)
	}

	err = ExportExcelFromSlice(records, &ExportConfig{
		Mode:       ExportByHeaders,
		OutputPath: "TestExcel3.xlsx",
		Headers:    []string{"姓名", "分数"},
		Header2FieldMap: map[string]string{
			"姓名": "Name",
			"分数": "Score",
		},
		SubSliceFieldName: "SubRecords",
		SubHeader2FieldMap: map[string]string{
			"题号": "QuestionNum",
			"答案": "Answer",
		},
	})
	if err != nil {
		panic(err)
	}

	err = ExportExcelFromSlice(records, &ExportConfig{
		Mode:       ExportByHeaders,
		OutputPath: "TestExcel4.xlsx",
		Headers:    []string{"姓名", "分数"},
		Header2FieldMap: map[string]string{
			"姓名": "Name",
			"分数": "Score",
		},
		SubSliceFieldName: "SubRecords",
	})
	if err != nil {
		panic(err)
	}
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
