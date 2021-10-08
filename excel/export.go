package excel

import (
	"errors"
	"fmt"
	"reflect"
	"strings"

	jsoniter "github.com/json-iterator/go"

	"github.com/360EntSecGroup-Skylar/excelize"
)

const XlsxTagName = "xlsx"

var json = jsoniter.ConfigCompatibleWithStandardLibrary

type ExportExcelErr error

var (
	RecordTypeErr   ExportExcelErr = errors.New("records is not slice")
	ConfigErr       ExportExcelErr = errors.New("export config err")
	SliceEmptyErr   ExportExcelErr = errors.New("slice is empty")
	HeaderConfigErr ExportExcelErr = errors.New("header config err")
)

// ExportExcelFromSlice 从结构体数组中导出
func ExportExcelFromSlice(records interface{}, exportConfig *ExportConfig) error {
	if !isSlice(records) {
		return RecordTypeErr
	}
	sl := reflect.ValueOf(records)
	if sl.Len() == 0 {
		return SliceEmptyErr
	}

	if checkConfig(exportConfig) {
		return ConfigErr
	}

	xlsx := excelize.NewFile()
	sheetName := "Sheet1"
	index := xlsx.NewSheet(sheetName)

	s := reflect.ValueOf(records)

	if exportConfig.Mode == ExportTaggedField {
		for row := 0; row < s.Len(); row++ {
			item := sl.Index(row)
			itemType := item.Type().Elem()
			for i := 0; i < itemType.NumField(); i++ {
				field := itemType.Field(i)
				xlsxTag := field.Tag.Get(XlsxTagName)
				if len(xlsxTag) == 0 || !strings.Contains(xlsxTag, "-") {
					// 忽略没有设置tag的字段
					continue
				}

				ss := strings.Split(xlsxTag, "-")
				columnAxis := ss[0]
				// 设置表头
				if row == 0 {
					header := ss[1]
					xlsx.SetCellValue(sheetName, fmt.Sprintf("%s%d", columnAxis, row+1), header)
				}
				// 设置内容
				setCellValue(sheetName, row+2, columnAxis, xlsx, field, item, field.Name)
			}
		}
	} else if exportConfig.Mode == ExportAllField {
		for row := 0; row < s.Len(); row++ {
			item := sl.Index(row)
			itemType := item.Type().Elem()
			if itemType.NumField() > 26*27 {

			}
			for i := 0; i < itemType.NumField(); i++ {
				field := itemType.Field(i)
				columnAxis := convertToTitle(i + 1)
				// 设置表头
				if row == 0 {
					header := field.Name
					xlsx.SetCellValue(sheetName, fmt.Sprintf("%s%d", columnAxis, row+1), header)
				}
				// 设置内容
				setCellValue(sheetName, row+2, columnAxis, xlsx, field, item, field.Name)
			}
		}
	} else if exportConfig.Mode == ExportByHeaders {
		headers := exportConfig.Headers
		header2FieldMap := exportConfig.Header2FieldMap
		if len(headers) == 0 || len(header2FieldMap) != len(headers) {
			return HeaderConfigErr
		}
		for row := 0; row < s.Len(); row++ {
			t := s.Index(row)
			d := t.Type().Elem()
			for col, header := range headers {
				alpha := convertToTitle(col + 1)
				colName := header2FieldMap[header]
				if len(colName) == 0 {
					continue
				}
				field, exist := d.FieldByName(colName)
				if !exist {
					continue
				}
				// 设置表头
				if row == 0 {
					xlsx.SetCellValue(sheetName, fmt.Sprintf("%s%d", alpha, row+1), header)
				}
				setCellValue(sheetName, row+2, alpha, xlsx, field, t, colName)
			}
		}
	}

	xlsx.SetActiveSheet(index)
	err := xlsx.SaveAs(exportConfig.OutputPath)
	if err != nil {
		return err
	}
	return nil
}

func setCellValue(sheetName string, row int, alpha string, xlsx *excelize.File, field reflect.StructField, t reflect.Value, colName string) {
	switch field.Type.Kind() {
	case reflect.String:
		xlsx.SetCellValue(sheetName, fmt.Sprintf("%s%d", alpha, row), t.Elem().FieldByName(colName).String())
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		xlsx.SetCellValue(sheetName, fmt.Sprintf("%s%d", alpha, row), t.Elem().FieldByName(colName).Int())
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		xlsx.SetCellValue(sheetName, fmt.Sprintf("%s%d", alpha, row), t.Elem().FieldByName(colName).Uint())
	case reflect.Bool:
		xlsx.SetCellValue(sheetName, fmt.Sprintf("%s%d", alpha, row), t.Elem().FieldByName(colName).Bool())
	case reflect.Float32, reflect.Float64:
		xlsx.SetCellValue(sheetName, fmt.Sprintf("%s%d", alpha, row), t.Elem().FieldByName(colName).Float())
	default:
		item := reflect.ValueOf(t).Elem().FieldByName(colName).Interface()
		bytes, _ := json.Marshal(&item)
		xlsx.SetCellValue(sheetName, fmt.Sprintf("%s%d", alpha, row), string(bytes))
	}
}

// 将数字转化成 Excel Header的 ColumnAxis
// index 从1开始计数
// index:1 => A; 27 => AA;
func convertToTitle(columnNumber int) string {
	var ans []byte
	for columnNumber > 0 {
		columnNumber--
		ans = append(ans, 'A'+byte(columnNumber%26))
		columnNumber /= 26
	}
	for i, n := 0, len(ans); i < n/2; i++ {
		ans[i], ans[n-1-i] = ans[n-1-i], ans[i]
	}
	return string(ans)
}

func checkConfig(config *ExportConfig) bool {
	return config == nil || len(config.OutputPath) == 0
}

func isSlice(records interface{}) bool {
	return reflect.TypeOf(records).Kind() == reflect.Slice
}

type ExportMode int

const (
	ExportTaggedField ExportMode = 0 // 仅导出有 xlsx tag的字段
	ExportAllField    ExportMode = 1 // 导出全部字段
	ExportByHeaders   ExportMode = 2 // 根据 ExportConfig 的 Headers 来决定导出字段
)

type ExportConfig struct {
	Mode            ExportMode        // 导出模式
	OutputPath      string            // 文件导出路径
	Headers         []string          // excel 列表头
	Header2FieldMap map[string]string // excel 列表头到结构体字段名的映射
}
