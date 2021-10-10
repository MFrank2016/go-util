package excel

import (
	"errors"
	"fmt"
	"io"
	"reflect"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
	jsoniter "github.com/json-iterator/go"
)

const XlsxTagName = "xlsx"

var json = jsoniter.ConfigCompatibleWithStandardLibrary

type ExportExcelErr error

var (
	RecordTypeErr         ExportExcelErr = errors.New("records is not slice")
	SubSliceTypeErr       ExportExcelErr = errors.New("sub slice's field type is not slice")
	ConfigErr             ExportExcelErr = errors.New("export config err")
	SliceEmptyErr         ExportExcelErr = errors.New("slice is empty")
	SubSliceEmptyErr      ExportExcelErr = errors.New("sub slice is empty")
	HeaderConfigErr       ExportExcelErr = errors.New("header config err")
	FieldNotExistErr      ExportExcelErr = errors.New("header field not exist")
	SubFieldNotExistErr   ExportExcelErr = errors.New("sub field not exist")
	SubSliceNotExistErr   ExportExcelErr = errors.New("header sub slice not exist")
	ExportModeNotExistErr ExportExcelErr = errors.New("export mode not exist")
	NoXlsxTagFoundErr     ExportExcelErr = errors.New("xlsx tag not found")
	AxisOutOfIndexErr     ExportExcelErr = errors.New("column axis out of index")
)

// ExportExcelFromSlice 从结构体切片中导出
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

	exportStrategy := getExportStrategy(exportConfig.Mode)
	if exportStrategy == nil {
		return ExportModeNotExistErr
	}
	strategyInitErr := exportStrategy.Init(sl.Index(0), exportConfig)
	if strategyInitErr != nil {
		return strategyInitErr
	}
	headers, getHeadersErr := exportStrategy.GetHeaders()
	if getHeadersErr != nil {
		return getHeadersErr
	}

	row := 0
	for i := 0; i < sl.Len(); i++ {
		item := sl.Index(i)
		itemType := item.Type().Elem()
		for i, header := range headers {
			fieldName, getFieldNameErr := exportStrategy.GetFieldNameByHeader(header)
			if getFieldNameErr != nil {
				return getFieldNameErr
			}
			field, exist := itemType.FieldByName(fieldName)
			if !exist {
				return FieldNotExistErr
			}
			if field.Name == exportConfig.SubSliceFieldName {
				if field.Type.Kind() != reflect.Slice {
					return SubSliceTypeErr
				}
				subSliceFieldErr := dealWithSubSliceField(item, xlsx, exportConfig, exportStrategy, &row, sheetName, i)
				if subSliceFieldErr != nil {
					return subSliceFieldErr
				}
			} else {
				colAxis, getColErr := exportStrategy.GetColAxis(i)
				if getColErr != nil {
					return getColErr
				}

				// 设置表头
				if row == 0 {
					xlsx.SetCellValue(sheetName, fmt.Sprintf("%s%d", colAxis, row+1), header)
				}
				// 设置内容
				setCellValue(sheetName, row+2, colAxis, xlsx, field, item, fieldName)
			}
		}
		row++
	}

	xlsx.SetActiveSheet(index)
	var ioErr error
	if len(exportConfig.OutputPath) > 0 {
		ioErr = xlsx.SaveAs(exportConfig.OutputPath)
	} else {
		ioErr = xlsx.Write(exportConfig.OutputWriter)
	}
	if ioErr != nil {
		return ioErr
	}
	return nil
}

func dealWithSubSliceField(item reflect.Value, xlsx *excelize.File, exportConfig *ExportConfig, exportStrategy ExportStrategy, row *int, sheetName string, colIndexBegin int) error {
	subSlice := item.Elem().FieldByName(exportConfig.SubSliceFieldName)
	for i := 0; i < subSlice.Len(); i++ {
		subItem := subSlice.Index(i)
		subItemType := subItem.Type().Elem()
		subHeaders, getSubHeaderErr := exportStrategy.GetSubHeaders()
		if getSubHeaderErr != nil {
			return getSubHeaderErr
		}
		for j, subHeader := range subHeaders {
			subFieldName, getSubFieldNameErr := exportStrategy.GetSubFieldNameByHeader(subHeader)
			if getSubFieldNameErr != nil {
				return getSubFieldNameErr
			}
			subField, exist := subItemType.FieldByName(subFieldName)
			if !exist {
				return FieldNotExistErr
			}
			colAxis, getColErr := exportStrategy.GetColAxis(colIndexBegin + j)
			if getColErr != nil {
				return getColErr
			}
			// 设置表头
			if *row == 0 {
				xlsx.SetCellValue(sheetName, fmt.Sprintf("%s%d", colAxis, *row+1), subHeader)
			}

			// 设置内容
			setCellValue(sheetName, *row+2, colAxis, xlsx, subField, subItem, subFieldName)
		}
		*row++
	}
	*row--
	return nil
}

func getExportStrategy(mode ExportMode) ExportStrategy {
	var exportStrategy ExportStrategy
	switch mode {
	case ExportTaggedField:
		exportStrategy = &exportTaggedFieldStrategy{}
	case ExportAllField:
		exportStrategy = &exportAllFieldStrategy{}
	case ExportByHeaders:
		exportStrategy = &exportByHeadersStrategy{}
	}
	return exportStrategy
}

func setCellValue(sheetName string, row int, alpha string, xlsx *excelize.File, field reflect.StructField, t reflect.Value, fieldName string) {
	itemFieldValue := t.Elem().FieldByName(fieldName)
	switch field.Type.Kind() {
	case reflect.String:
		xlsx.SetCellValue(sheetName, fmt.Sprintf("%s%d", alpha, row), itemFieldValue.String())
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		xlsx.SetCellValue(sheetName, fmt.Sprintf("%s%d", alpha, row), itemFieldValue.Int())
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		xlsx.SetCellValue(sheetName, fmt.Sprintf("%s%d", alpha, row), itemFieldValue.Uint())
	case reflect.Bool:
		xlsx.SetCellValue(sheetName, fmt.Sprintf("%s%d", alpha, row), itemFieldValue.Bool())
	case reflect.Float32, reflect.Float64:
		xlsx.SetCellValue(sheetName, fmt.Sprintf("%s%d", alpha, row), itemFieldValue.Float())
	default:
		item := itemFieldValue.Interface()
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
	return config == nil || (len(config.OutputPath) == 0 && config.OutputWriter == nil)
}

func isSlice(records interface{}) bool {
	return reflect.TypeOf(records).Kind() == reflect.Slice
}

type ExportConfig struct {
	Mode               ExportMode        // 导出模式
	OutputPath         string            // 文件导出路径
	OutputWriter       io.Writer         // 文件输出流，优先使用导出路径
	Headers            []string          // excel 列表头
	Header2FieldMap    map[string]string // excel 列表头到结构体字段名的映射
	SubSliceFieldName  string            // 附属子列表字段名，设置该字段则会进行导出，不设置则不导出
	SubHeader2FieldMap map[string]string // excel 子列表头到结构体字段名的映射
}

type ExportMode int

const (
	ExportTaggedField ExportMode = 0 // 仅导出有 xlsx tag的字段
	ExportAllField    ExportMode = 1 // 导出全部字段
	ExportByHeaders   ExportMode = 2 // 根据 ExportConfig 的 Headers 来决定导出字段
)

type ExportStrategy interface {
	Init(item reflect.Value, config *ExportConfig) error
	GetHeaders() (headers []string, err error)
	GetSubHeaders() (subHeaders []string, err error)
	GetFieldNameByHeader(header string) (fieldName string, err error)
	GetSubFieldNameByHeader(header string) (fieldName string, err error)
	GetColAxis(columnNum int) (axis string, err error)
}

type exportByHeadersStrategy struct {
	headers                   []string
	header2FieldNameMap       map[string]string
	subHeaders                []string
	subHeader2SubFieldNameMap map[string]string
}

func (s *exportByHeadersStrategy) Init(item reflect.Value, config *ExportConfig) error {
	if len(config.Headers) == 0 || len(config.Header2FieldMap) != len(config.Headers) {
		return HeaderConfigErr
	}
	fieldName2HeaderMap := make(map[string]string)
	for header, fieldName := range config.Header2FieldMap {
		fieldName2HeaderMap[fieldName] = header
	}
	subFieldName2HeaderMap := make(map[string]string)
	for header, fieldName := range config.SubHeader2FieldMap {
		subFieldName2HeaderMap[fieldName] = header
	}

	s.headers = make([]string, 0)
	s.subHeaders = make([]string, 0)
	s.header2FieldNameMap = config.Header2FieldMap
	s.subHeader2SubFieldNameMap = config.SubHeader2FieldMap
	itemType := item.Type().Elem()
	for i := 0; i < itemType.NumField(); i++ {
		fieldName := itemType.Field(i).Name
		if header, exist := fieldName2HeaderMap[fieldName]; exist {
			s.headers = append(s.headers, header)
		}
	}
	if len(config.SubSliceFieldName) > 0 {
		s.header2FieldNameMap[config.SubSliceFieldName] = config.SubSliceFieldName
		s.headers = append(s.headers, config.SubSliceFieldName)
		subSliceField, exist := itemType.FieldByName(config.SubSliceFieldName)
		if !exist {
			return SubSliceNotExistErr
		}
		if subSliceField.Type.Kind() != reflect.Slice {
			return SubSliceTypeErr
		}

		subSlice := item.Elem().FieldByName(config.SubSliceFieldName)
		if subSlice.Len() == 0 {
			return SubSliceEmptyErr
		}
		subItem := subSlice.Index(0)
		subItemType := subItem.Type().Elem()
		for i := 0; i < subItemType.NumField(); i++ {
			subField := subItemType.Field(i)
			if header, exist := subFieldName2HeaderMap[subField.Name]; exist {
				s.subHeaders = append(s.subHeaders, header)
			}
		}
	}
	return nil
}

func (s *exportByHeadersStrategy) GetHeaders() (headers []string, err error) {
	return s.headers, nil
}

func (s *exportByHeadersStrategy) GetSubHeaders() (headers []string, err error) {
	return s.subHeaders, nil
}

func (s *exportByHeadersStrategy) GetFieldNameByHeader(header string) (fieldName string, err error) {
	name, exist := s.header2FieldNameMap[header]
	if !exist {
		return "", FieldNotExistErr
	}
	return name, nil
}

func (s *exportByHeadersStrategy) GetSubFieldNameByHeader(header string) (fieldName string, err error) {
	name, exist := s.subHeader2SubFieldNameMap[header]
	if !exist {
		return "", FieldNotExistErr
	}
	return name, nil
}

func (s *exportByHeadersStrategy) GetColAxis(columnNum int) (axis string, err error) {
	columnAxis := convertToTitle(columnNum + 1)
	return columnAxis, nil
}

type exportAllFieldStrategy struct {
	headers    []string
	subHeaders []string
}

func (s *exportAllFieldStrategy) Init(item reflect.Value, config *ExportConfig) error {
	s.headers = make([]string, 0)
	s.subHeaders = make([]string, 0)
	itemType := item.Type().Elem()
	for i := 0; i < itemType.NumField(); i++ {
		field := itemType.Field(i)
		if field.Name == config.SubSliceFieldName {
			continue
		}
		s.headers = append(s.headers, field.Name)
	}
	if len(s.headers) == 0 {
		return NoXlsxTagFoundErr
	}
	if len(config.SubSliceFieldName) > 0 {
		s.headers = append(s.headers, config.SubSliceFieldName)
		subSliceField, exist := itemType.FieldByName(config.SubSliceFieldName)
		if !exist {
			return SubSliceNotExistErr
		}
		if subSliceField.Type.Kind() != reflect.Slice {
			return SubSliceTypeErr
		}

		subSlice := item.Elem().FieldByName(config.SubSliceFieldName)
		if subSlice.Len() == 0 {
			return SubSliceEmptyErr
		}
		subItem := subSlice.Index(0)
		subItemType := subItem.Type().Elem()
		for i := 0; i < subItemType.NumField(); i++ {
			subField := subItemType.Field(i)
			s.subHeaders = append(s.subHeaders, subField.Name)
		}
	}
	return nil
}

func (s *exportAllFieldStrategy) GetHeaders() (headers []string, err error) {
	return s.headers, nil
}

func (s *exportAllFieldStrategy) GetSubHeaders() (subHeaders []string, err error) {
	return s.subHeaders, nil
}

func (s *exportAllFieldStrategy) GetFieldNameByHeader(header string) (fieldName string, err error) {
	return header, nil
}

func (s *exportAllFieldStrategy) GetSubFieldNameByHeader(header string) (fieldName string, err error) {
	return header, nil
}

func (s *exportAllFieldStrategy) GetColAxis(columnNum int) (axis string, err error) {
	columnAxis := convertToTitle(columnNum + 1)
	return columnAxis, nil
}

type exportTaggedFieldStrategy struct {
	columnAxis                []string
	headers                   []string
	header2FieldNameMap       map[string]string
	subHeaders                []string
	subHeader2SubFieldNameMap map[string]string
}

func (s *exportTaggedFieldStrategy) Init(item reflect.Value, config *ExportConfig) error {
	s.columnAxis = make([]string, 0)
	s.headers = make([]string, 0)
	s.header2FieldNameMap = make(map[string]string)
	s.subHeader2SubFieldNameMap = make(map[string]string)
	itemType := item.Type().Elem()
	for i := 0; i < itemType.NumField(); i++ {
		field := itemType.Field(i)
		xlsxTag := field.Tag.Get(XlsxTagName)
		if len(xlsxTag) == 0 {
			// 忽略没有设置tag的字段
			continue
		}
		ss := strings.Split(xlsxTag, "-")
		if len(ss) != 2 {
			continue
		}
		s.columnAxis = append(s.columnAxis, ss[0])
		s.headers = append(s.headers, ss[1])
		s.header2FieldNameMap[ss[1]] = field.Name
	}
	if len(s.headers) == 0 {
		return NoXlsxTagFoundErr
	}

	if len(config.SubSliceFieldName) > 0 {
		s.headers = append(s.headers, config.SubSliceFieldName)
		s.header2FieldNameMap[config.SubSliceFieldName] = config.SubSliceFieldName
		subSliceField, exist := itemType.FieldByName(config.SubSliceFieldName)
		if !exist {
			return SubSliceNotExistErr
		}
		if subSliceField.Type.Kind() != reflect.Slice {
			return SubSliceTypeErr
		}

		subSlice := item.Elem().FieldByName(config.SubSliceFieldName)
		if subSlice.Len() == 0 {
			return SubSliceEmptyErr
		}
		subItem := subSlice.Index(0)
		subItemType := subItem.Type().Elem()
		for i := 0; i < subItemType.NumField(); i++ {
			subField := subItemType.Field(i)
			xlsxTag := subField.Tag.Get(XlsxTagName)
			if len(xlsxTag) == 0 {
				// 忽略没有设置tag的字段
				continue
			}
			ss := strings.Split(xlsxTag, "-")
			if len(ss) != 2 {
				continue
			}
			s.columnAxis = append(s.columnAxis, ss[0])
			s.subHeaders = append(s.subHeaders, ss[1])
			s.subHeader2SubFieldNameMap[ss[1]] = subField.Name
		}
	}
	return nil
}

func (s *exportTaggedFieldStrategy) GetHeaders() (headers []string, err error) {
	return s.headers, nil
}

func (s *exportTaggedFieldStrategy) GetSubHeaders() (subHeaders []string, err error) {
	return s.subHeaders, nil
}

func (s *exportTaggedFieldStrategy) GetFieldNameByHeader(header string) (fieldName string, err error) {
	name, exist := s.header2FieldNameMap[header]
	if !exist {
		return "", FieldNotExistErr
	}
	return name, nil
}

func (s *exportTaggedFieldStrategy) GetSubFieldNameByHeader(header string) (fieldName string, err error) {
	name, exist := s.subHeader2SubFieldNameMap[header]
	if !exist {
		return "", SubFieldNotExistErr
	}
	return name, nil
}

func (s *exportTaggedFieldStrategy) GetColAxis(columnNum int) (axis string, err error) {
	if columnNum < len(s.columnAxis) {
		return s.columnAxis[columnNum], nil
	}
	return "", AxisOutOfIndexErr
}
