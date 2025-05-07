package excel_template

import (
	"fmt"
	"strings"
	"text/template"

	"github.com/mzzya/excel_template/constant"
	"github.com/tiendc/go-deepcopy"

	"github.com/samber/lo"
	"github.com/xuri/excelize/v2"
)

type ColumnCell struct {
	_key    int
	Formula string
	// 数据单元格样式Id
	StyleId int
	// 数据单元格样式
	Style *excelize.Style

	Data string
}
type Column struct {
	_key int
	// 模板中的列号
	ColNum  int
	ColName string
	// 渲染时的列号=ColNum-1
	RenderColNum  int
	RenderColName string

	Header string
	// 数据字段
	DataField string
	// 是不是模板字段
	IsTemplate          bool
	BackgroundColorExpr string
	FontColorExpr       string

	CellList []*ColumnCell
}

type SheetCache struct {
	Config        map[string][][]string
	ColumnList    []*Column
	StartRowNum   int
	FillData      map[string]any
	DataRowHeight float64
}

// ExcelTemplate 表示Excel模板渲染器
type ExcelTemplate struct {
	TemplatePath  string
	File          *excelize.File
	SheetCache    map[string]*SheetCache
	FormulaEngine FormulaEngine
	FuncMap       template.FuncMap
	ListField     string
}

var configKeys = []string{constant.Header, constant.Data, constant.DataField, constant.BackgroundColor, constant.FontColor, constant.Subtotal}

// var formulaEngine FormulaEngine

// func SetFormulaEngine(fe FormulaEngine) {
// 	formulaEngine = fe
// }
// func init() {
// 	SetFormulaEngine(NewFormulaEnginePool(10, 20, NewSimpleFormulaEngine))
// }

// OpenFile 创建一个新的Excel模板渲染器
func OpenFile(templatePath string) (*ExcelTemplate, error) {
	f, err := excelize.OpenFile(templatePath)
	if err != nil {
		fmt.Println(err)
		return nil, fmt.Errorf("OpenFile: failed to open Excel file [path=%s]: %w", templatePath, err)
	}
	et := &ExcelTemplate{
		TemplatePath:  templatePath,
		File:          f,
		SheetCache:    make(map[string]*SheetCache),
		FormulaEngine: NewSimpleFormulaEngine(),
		ListField:     "table",
	}
	return et, nil
}

type RenderOptions struct {
	Data map[string]any
}

// Render 渲染Excel模板
func (et *ExcelTemplate) Render(data any) (*excelize.File, error) {
	//遍历所有的sheet
	for i, sheet := range et.File.GetSheetList() {
		et.SheetCache[sheet] = &SheetCache{
			Config:      make(map[string][][]string),
			ColumnList:  make([]*Column, 0, 3),
			StartRowNum: 0,
		}
		var ok bool
		et.SheetCache[sheet].FillData, ok = data.(map[string]any)
		if ok {
			listData, ok := data.([]map[string]any)
			if ok && len(listData) > i {
				et.SheetCache[sheet].FillData = listData[i]
			}
		}
		err := et.processSheet(sheet)
		if err != nil {
			return nil, fmt.Errorf("Render: failed to process sheet [sheet=%s]: %w", sheet, err)
		}
	}

	//更新公式缓存
	et.File.UpdateLinkedValue()
	return et.File, nil
}

// getFormulaResult 获取公式计算结果
func (et *ExcelTemplate) getFormulaResult(formulaResultCache map[string]any, listIndex int, formulaExpr string, data map[string]any) (any, error) {
	// 构造缓存key
	cacheKey := fmt.Sprintf("%d_%s", listIndex, formulaExpr)
	if cachedResult, ok := formulaResultCache[cacheKey]; ok {
		// 直接使用缓存的结果
		return cachedResult, nil
	}
	value, _, err := et.FormulaEngine.EvalFormula(formulaExpr, data)
	if err != nil {
		return nil, fmt.Errorf("getFormulaResult: failed to evaluate formula [expr=%s]: %w", formulaExpr, err)
	}
	formulaResultCache[cacheKey] = value
	return value, nil
}

func (et *ExcelTemplate) fillRows(mergeCells []excelize.MergeCell, rows [][]string) [][]string {
	for _, merge := range mergeCells {
		rangeStr := merge[0]
		value := merge[1]

		startCell, endCell := "", ""
		if coords := strings.Split(rangeStr, ":"); len(coords) == 2 {
			startCell, endCell = coords[0], coords[1]
		} else {
			continue
		}

		startCol, startRow, err1 := excelize.CellNameToCoordinates(startCell)
		endCol, endRow, err2 := excelize.CellNameToCoordinates(endCell)
		if err1 != nil || err2 != nil {
			continue
		}
		// 填充 value 到 rows 中对应的区域（行和列都从 1 开始，所以要减 1）
		for row := startRow - 1; row <= endRow-1; row++ {
			// 若当前行超出 rows 范围，先扩展 rows 行数
			for len(rows) <= row {
				rows = append(rows, []string{})
			}

			// 当前行的最大列数
			neededCols := endCol
			if len(rows[row]) < neededCols {
				newRow := make([]string, neededCols)
				copy(newRow, rows[row])
				rows[row] = newRow
			}

			// 写入合并区域的每个单元格
			for col := startCol - 1; col <= endCol-1; col++ {
				rows[row][col] = value
			}
		}
	}
	return rows
}

// processSheet 处理单个sheet的数据
func (et *ExcelTemplate) processSheet(sheet string) error {
	// 获取基础数据
	rows, mergeCells, err := et.getSheetData(sheet)
	if err != nil {
		return fmt.Errorf("processSheet: failed to get sheet data [sheet=%s]: %w", sheet, err)
	}

	// 处理模板语法
	err = et.processTemplates(sheet, rows)
	if err != nil {
		return fmt.Errorf("processSheet: failed to process templates [sheet=%s]: %w", sheet, err)
	}

	rows = et.fillRows(mergeCells, rows)

	// 处理配置和列信息
	config := make(map[string][][]string)
	columns := make([]*Column, 0, 3)
	fillRowNum := 0
	configRowNums := make([]int, 0, 1)

	for rowIndex, row := range rows {
		//如果不是配置列，则跳过
		if len(row) == 0 || row[0] == "" || !lo.Contains(configKeys, row[0]) {
			continue
		}

		configName := row[0]
		if config[configName] == nil {
			config[configName] = make([][]string, 0, 1)
		}
		config[configName] = append(config[configName], row)
		rowNum := rowIndex + 1
		configRowNums = append(configRowNums, rowNum)

		if configName == constant.Header {
			fillRowNum = rowNum + 1
		}

		for colIndex, col := range row {
			if colIndex == 0 {
				continue
			}
			colNum := colIndex + 1
			value := col
			cellName, err := excelize.CoordinatesToCellName(colNum, rowNum)
			if err != nil {
				return fmt.Errorf("processSheet: failed to convert coordinates to cell name [sheet=%s, row=%d, col=%d]: %w", sheet, rowNum, colNum, err)
			}

			// if ContainsGoTemplateSyntax(value) {
			// 	value, err = RenderTemplate(value, et.SheetCache[sheet].FillData, et.FuncMap)
			// 	if err != nil {
			// 		fmt.Println(cellName, value)
			// 		panic(err)
			// 	}
			// 	et.File.SetCellValue(sheet, cellName, value)
			// }

			column, ok := lo.Find(columns, func(column *Column) bool {
				return column != nil && column._key == colNum
			})

			if ok {
				switch configName {
				case constant.DataField:
					column.DataField = value
					column.IsTemplate = ContainsGoTemplateSyntax(value)
				case constant.BackgroundColor:
					column.BackgroundColorExpr = value
				case constant.FontColor:
					column.FontColorExpr = value
				case constant.Data:
					if column.CellList == nil {
						column.CellList = make([]*ColumnCell, 0, 1)
					}
					columnCell := ColumnCell{}
					et.SheetCache[sheet].DataRowHeight, err = et.File.GetRowHeight(sheet, rowNum)
					if err != nil {
						return fmt.Errorf("processSheet: failed to get row height [sheet=%s, row=%d]: %w", sheet, rowNum, err)
					}

					columnCell._key = rowNum
					columnCell.Formula, err = et.File.GetCellFormula(sheet, cellName)
					if err != nil {
						return fmt.Errorf("processSheet: failed to get cell formula [sheet=%s, cell=%s]: %w", sheet, cellName, err)
					}

					//设置样式
					columnCell.StyleId, err = et.File.GetCellStyle(sheet, cellName)
					if err != nil {
						return fmt.Errorf("processSheet: failed to get cell style [sheet=%s, cell=%s]: %w", sheet, cellName, err)
					}
					columnCell.Style, err = et.File.GetStyle(columnCell.StyleId)
					if err != nil {
						return fmt.Errorf("processSheet: failed to get style details [sheet=%s, styleId=%d]: %w", sheet, columnCell.StyleId, err)
					}
					//清空公式，因为带公式的话后续RemoveRow会报错
					et.File.SetCellFormula(sheet, cellName, "")
					column.CellList = append(column.CellList, &columnCell)
				}
			} else {
				if value == "" {
					continue
				}

				column := Column{_key: colNum, ColNum: colNum, RenderColNum: colNum - 1}
				column.Header = value

				colName, err := excelize.ColumnNumberToName(colNum)
				if err != nil {
					return fmt.Errorf("processSheet: failed to convert column number to name [sheet=%s, col=%d]: %w", sheet, colNum, err)
				}
				column.ColName = colName

				renderColName, err := excelize.ColumnNumberToName(column.RenderColNum)
				if err != nil {
					return fmt.Errorf("processSheet: failed to convert render column number to name [sheet=%s, col=%d]: %w", sheet, column.RenderColNum, err)
				}
				column.RenderColName = renderColName

				columns = append(columns, &column)
			}
		}

		// 缓存配置
		et.SheetCache[sheet].Config = config
		et.SheetCache[sheet].ColumnList = columns
		et.SheetCache[sheet].StartRowNum = fillRowNum
	}

	if len(columns) == 0 {
		return nil
	}

	//表头留1行 数据留2行 这样如果有公式的话会自动更新
	for i := len(configRowNums) - 1; i > 2; i-- {
		// fmt.Println("remove row", configRowNums[i])
		et.File.RemoveRow(sheet, configRowNums[i])
	}
	et.File.RemoveCol(sheet, "A")

	table, ok := et.SheetCache[sheet].FillData[et.ListField]
	if !ok {
		return nil
	}
	list, ok := table.([]map[string]any)
	if !ok {
		tableList, ok := table.([]any)
		if ok {
			list = lo.Map(tableList, func(item any, index int) map[string]any {
				return item.(map[string]any)
			})
		}
	}

	if list == nil || len(list) == 0 {
		return nil
	}

	// 处理分类汇总
	list = et.handleSubtotal(config, list, fillRowNum)

	// 插入数据行
	et.File.InsertRows(sheet, fillRowNum+1, len(list)-2)

	// 处理数据填充
	err = et.processData(sheet, list)
	if err != nil {
		return fmt.Errorf("processSheet: failed to process data [sheet=%s]: %w", sheet, err)
	}

	// 设置自动筛选
	et.setAutoFilter(sheet, len(list))
	return nil
}

// getSheetData 获取sheet的基础数据
func (et *ExcelTemplate) getSheetData(sheet string) ([][]string, []excelize.MergeCell, error) {
	rows, err := et.File.GetRows(sheet)
	if err != nil {
		return nil, nil, fmt.Errorf("getSheetData: failed to get sheet rows [sheet=%s]: %w", sheet, err)
	}
	mergeCells, err := et.File.GetMergeCells(sheet)
	if err != nil {
		return nil, nil, fmt.Errorf("getSheetData: failed to get merge cells [sheet=%s]: %w", sheet, err)
	}
	return rows, mergeCells, nil
}

// processTemplates 处理模板语法
func (et *ExcelTemplate) processTemplates(sheet string, rows [][]string) error {
	fillData := et.SheetCache[sheet].FillData
	for i, row := range rows {
		if len(row) > 0 && row[0] == constant.DataField {
			continue
		}
		for j, col := range row {
			if ContainsGoTemplateSyntax(col) {
				value, err := RenderTemplate(col, fillData, et.FuncMap)
				if err != nil {
					return fmt.Errorf("processTemplates: failed to render template [sheet=%s, row=%d, col=%d]: %w", sheet, i+1, j+1, err)
				}
				cellName, err := excelize.CoordinatesToCellName(j+1, i+1)
				if err != nil {
					return fmt.Errorf("processTemplates: failed to convert coordinates to cell name [sheet=%s, row=%d, col=%d]: %w", sheet, i+1, j+1, err)
				}
				et.File.SetCellValue(sheet, cellName, value)
			}
		}
	}
	return nil
}

// processData 处理数据填充
func (et *ExcelTemplate) processData(sheet string, list []map[string]any) error {
	for i := range list {
		err := et.processDataRow(sheet, i, list[i])
		if err != nil {
			return fmt.Errorf("processData: failed to process data row [sheet=%s, row=%d]: %w", sheet, i, err)
		}
	}
	return nil
}

func (et *ExcelTemplate) processDataRow(sheet string, listIndex int, rowData map[string]any) error {
	columns := et.SheetCache[sheet].ColumnList
	fillRowNum := et.SheetCache[sheet].StartRowNum
	rowNum := fillRowNum + listIndex
	formulaResultCache := make(map[string]any)
	styleIdCache := make(map[string]int)
	et.File.SetRowHeight(sheet, rowNum, et.SheetCache[sheet].DataRowHeight)
	isSubtotal := rowData["_row_type"] == "subtotal"
	_listIndex := listIndex
	if _, ok := rowData["_row_index"]; ok {
		_listIndex = rowData["_row_index"].(int)
	}
	for _, column := range columns {
		cellName := fmt.Sprintf("%s%d", column.RenderColName, rowNum)
		err := et.setCellValue(sheet, cellName, column, _listIndex, rowNum, rowData, isSubtotal)
		if err != nil {
			return fmt.Errorf("processDataRow: failed to set cell value [sheet=%s, cell=%s]: %w", sheet, cellName, err)
		}

		if isSubtotal {
			et.File.SetCellStyle(sheet, cellName, cellName, 0)
			continue
		}

		err = et.applyCellStyle(sheet, formulaResultCache, styleIdCache, cellName, column, _listIndex, rowData)
		if err != nil {
			return fmt.Errorf("processDataRow: failed to apply cell style [sheet=%s, cell=%s]: %w", sheet, cellName, err)
		}
	}
	return nil
}

// setAutoFilter 设置自动筛选
func (et *ExcelTemplate) setAutoFilter(sheet string, listLen int) {
	columns := et.SheetCache[sheet].ColumnList
	fillRowNum := et.SheetCache[sheet].StartRowNum
	if len(columns) == 0 {
		return
	}

	startCell, serr := excelize.CoordinatesToCellName(columns[0].RenderColNum, fillRowNum-1)
	endCell, eerr := excelize.CoordinatesToCellName(columns[len(columns)-1].RenderColNum, fillRowNum+listLen-1)
	if serr == nil && eerr == nil {
		et.File.AutoFilter(sheet, startCell+":"+endCell, []excelize.AutoFilterOptions{})
	}
}

// applyCellStyle 处理单元格样式设置，包括背景色和字体颜色
func (et *ExcelTemplate) applyCellStyle(sheet string, formulaResultCache map[string]any, styleIdCache map[string]int, cellName string, column *Column, listIndex int, rowData map[string]any) error {
	idx := listIndex % len(column.CellList)
	dataProp := column.CellList[idx]
	et.File.SetCellStyle(sheet, cellName, cellName, dataProp.StyleId)

	if len(column.BackgroundColorExpr) == 0 && len(column.FontColorExpr) == 0 {
		return nil
	}

	var bgColor = ""
	if column.BackgroundColorExpr != "" && column.BackgroundColorExpr[0] == '=' {
		result, err := et.getFormulaResult(formulaResultCache, listIndex, column.BackgroundColorExpr, rowData)
		if err != nil {
			return fmt.Errorf("applyCellStyle: failed to calculate background color formula [sheet=%s, cell=%s, expr=%s]: %w", sheet, cellName, column.BackgroundColorExpr, err)
		}
		if result != nil {
			bgColor, _ = result.(string)
		}
	}
	var fontColor = ""
	if column.FontColorExpr != "" && column.FontColorExpr[0] == '=' {
		result, err := et.getFormulaResult(formulaResultCache, listIndex, column.FontColorExpr, rowData)
		if err != nil {
			return fmt.Errorf("applyCellStyle: failed to calculate font color formula [sheet=%s, cell=%s, expr=%s]: %w", sheet, cellName, column.FontColorExpr, err)
		}
		if result != nil {
			fontColor, _ = result.(string)
		}
	}
	if bgColor == "" && fontColor == "" {
		return nil
	}

	styleKey := fmt.Sprintf("%d-%s-%s", dataProp.StyleId, bgColor, fontColor)
	styleId, ok := styleIdCache[styleKey]
	if ok {
		et.File.SetCellStyle(sheet, cellName, cellName, styleId)
	} else {
		style := &excelize.Style{}
		deepcopy.Copy(style, dataProp.Style)
		if bgColor != "" {
			style.Fill.Type = "pattern"
			style.Fill.Pattern = 1
			style.Fill.Color = []string{bgColor}
		}
		if fontColor != "" {
			style.Font.Color = fontColor
			style.Font.ColorTheme = nil
		}
		styleId, err := et.File.NewStyle(style)
		if err != nil {
			return fmt.Errorf("applyCellStyle: failed to create new style [sheet=%s, cell=%s]: %w", sheet, cellName, err)
		}
		styleIdCache[styleKey] = styleId
		et.File.SetCellStyle(sheet, cellName, cellName, styleId)
	}
	return nil
}

// setCellValue 处理单元格值设置，包括小计行和普通数据行
func (et *ExcelTemplate) setCellValue(sheet string, cellName string, column *Column, listIndex int, rowNum int, rowData map[string]any, isSubtotal bool) error {
	idx := listIndex % len(column.CellList)
	itemData, ok := rowData[column.DataField]
	//如果是分类汇总字段
	if isSubtotal {
		et.File.SetCellValue(sheet, cellName, "")
		if ok {
			valueStr, ok := itemData.(string)
			// fmt.Println("valueStr", itemData, valueStr, column.DataField)
			if ok && len(valueStr) > 1 {
				if valueStr[0] == '=' {
					et.File.SetCellFormula(sheet, cellName, valueStr[1:])
				} else {
					et.File.SetCellValue(sheet, cellName, itemData)
				}
			}
		}
		return nil
	}

	dataProp := column.CellList[idx]
	//如果是公式
	if dataProp.Formula != "" {
		var newFormula = ReplaceFormulaRow(dataProp.Formula, rowNum, -1)
		et.File.SetCellFormula(sheet, cellName, newFormula)
		return nil
	}
	//如果字段使用了模板语法
	if column.IsTemplate {
		value, err := RenderTemplate(column.DataField, rowData, et.FuncMap)
		if err != nil {
			return fmt.Errorf("setCellValue: failed to render template field [sheet=%s, cell=%s, field=%s]: %w", sheet, cellName, column.DataField, err)
		}
		et.File.SetCellValue(sheet, cellName, value)
		return nil
	}
	// dataConfigValue := column.CellList[idx].Data
	// if dataConfigValue != "" {
	// 	fmt.Println("dataConfigValue", cellName, dataConfigValue)
	// 	if len(dataConfigValue) > 0 && dataConfigValue[0] == '=' {
	// 		value, _, err := et.FormulaEngine.EvalFormula(dataConfigValue, rowData)
	// 		if err != nil {
	// 			panic(err)
	// 		}
	// 		et.File.SetCellValue(sheet, cellName, value)
	// 	} else {
	// 		et.File.SetCellValue(sheet, cellName, dataConfigValue)
	// 	}
	// }
	et.File.SetCellValue(sheet, cellName, itemData)
	return nil
}

// handleSubtotal 处理数据的分类汇总
func (et *ExcelTemplate) handleSubtotal(config map[string][][]string, list []map[string]any, fillRowNum int) []map[string]any {
	if rows, ok := config[constant.Subtotal]; ok {
		for _, row := range rows {
			_, groupByIdex, gOk := lo.FindIndexOf(row, func(item string) bool {
				return item == "分类"
			})
			_, subtotalIndex, sOk := lo.FindIndexOf(row, func(item string) bool {
				return item == "求和"
			})
			if gOk && sOk {
				groupKey := config[constant.DataField][0][groupByIdex]
				subtotalKey := config[constant.DataField][0][subtotalIndex]
				subtotalCellLetter, _ := excelize.ColumnNumberToName(subtotalIndex)
				subtotal, _ := lo.Find(Subtotals, func(item Subtotal) bool {
					return item.Func == "求和"
				})
				list = GroupAndSubtotal(list, groupKey, subtotalKey, subtotalCellLetter, fillRowNum, subtotal)
			}
		}
	}
	return list
}
