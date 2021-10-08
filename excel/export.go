package excel

import (
	"errors"
	"fmt"
	"reflect"
	"strings"
	"unicode"

	jsoniter "github.com/json-iterator/go"

	"github.com/360EntSecGroup-Skylar/excelize"
)

const XlsxTagName = "xlsx"

// ExportExcel 从结构体数组中导出
func ExportExcel(records interface{}, filePath string) (exportErr error) {
	if reflect.TypeOf(records).Kind() != reflect.Slice {
		return errors.New("records is not slice")
	}
	xlsx := excelize.NewFile()
	sheetName := "Sheet1"
	index := xlsx.NewSheet(sheetName)

	sl := reflect.ValueOf(records)
	if sl.Len() == 0 {
		return nil
	}

	// 设置内容
	var json = jsoniter.ConfigCompatibleWithStandardLibrary
	s := reflect.ValueOf(records)
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

			fieldValue := item.Elem().Field(i)
			switch field.Type.Kind() {
			case reflect.String:
				xlsx.SetCellValue(sheetName, fmt.Sprintf("%s%d", columnAxis, row+2), fieldValue.String())
			case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
				xlsx.SetCellValue(sheetName, fmt.Sprintf("%s%d", columnAxis, row+2), fieldValue.Int())
			case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
				xlsx.SetCellValue(sheetName, fmt.Sprintf("%s%d", columnAxis, row+2), fieldValue.Uint())
			case reflect.Bool:
				xlsx.SetCellValue(sheetName, fmt.Sprintf("%s%d", columnAxis, row+2), fieldValue.Bool())
			case reflect.Float32, reflect.Float64:
				xlsx.SetCellValue(sheetName, fmt.Sprintf("%s%d", columnAxis, row+2), fieldValue.Float())
			default:
				item := fieldValue.Interface()
				bytes, _ := json.Marshal(&item)
				xlsx.SetCellValue(sheetName, fmt.Sprintf("%s%d", columnAxis, row+2), string(bytes))
			}
		}
	}

	xlsx.SetActiveSheet(index)
	err := xlsx.SaveAs(filePath)
	if err != nil {
		return err
	}
	return nil
}

func RefactorWriteV2(records interface{}, headers []string, header2FieldMap map[string]string, filePath string) error {
	if reflect.TypeOf(records).Kind() != reflect.Slice {
		return errors.New("records is not slice")
	}
	xlsx := excelize.NewFile()
	index := xlsx.NewSheet("Sheet1")
	sheet := fmt.Sprintf("Sheet%d", index)

	alphabet := make([]string, 0, 26)
	for r := 'a'; r < 'z'; r++ {
		R := unicode.ToUpper(r)
		alphabet = append(alphabet, string(R))
	}

	// 设置表头
	for i, header := range headers {
		alpha := alphabet[i%26]
		xlsx.SetCellValue(sheet, fmt.Sprintf("%s%d", alpha, 1), header)
	}

	// 设置内容
	var json = jsoniter.ConfigCompatibleWithStandardLibrary
	s := reflect.ValueOf(records)
	for row := 0; row < s.Len(); row++ {
		t := s.Index(row)
		d := t.Type().Elem()
		for col, header := range headers {
			alpha := alphabet[col%26]
			colName := header2FieldMap[header]
			if len(colName) == 0 {
				continue
			}
			field, exist := d.FieldByName(colName)
			if !exist {
				continue
			}
			switch field.Type.Kind() {
			case reflect.String:
				xlsx.SetCellValue(sheet, fmt.Sprintf("%s%d", alpha, row+2), t.Elem().FieldByName(colName).String())
			case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
				xlsx.SetCellValue(sheet, fmt.Sprintf("%s%d", alpha, row+2), t.Elem().FieldByName(colName).Int())
			case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
				xlsx.SetCellValue(sheet, fmt.Sprintf("%s%d", alpha, row+2), t.Elem().FieldByName(colName).Uint())
			case reflect.Bool:
				xlsx.SetCellValue(sheet, fmt.Sprintf("%s%d", alpha, row+2), t.Elem().FieldByName(colName).Bool())
			case reflect.Float32, reflect.Float64:
				xlsx.SetCellValue(sheet, fmt.Sprintf("%s%d", alpha, row+2), t.Elem().FieldByName(colName).Float())
			default:
				item := reflect.ValueOf(t).Elem().FieldByName(colName).Interface()
				bytes, _ := json.Marshal(&item)
				xlsx.SetCellValue(sheet, fmt.Sprintf("%s%d", alpha, row+2), string(bytes))
			}
		}
	}

	xlsx.SetActiveSheet(index)
	// 保存到xlsx中
	err := xlsx.SaveAs(filePath)
	if err != nil {
		return err
	}
	return nil
}
