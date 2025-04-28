package excel_template

import "fmt"

type Subtotal struct {
	Code             int
	Func             string
	GroupFieldSuffix string
	TotalField       string
}

var Subtotals = []Subtotal{
	{Code: 9, Func: "求和", GroupFieldSuffix: "汇总", TotalField: "总计"},
}

func GroupAndSubtotal(data []map[string]any, groupField string, sumField string, subtotalCellLetter string, dataStartRow int, subtotal Subtotal) []map[string]any {
	if len(data) == 0 {
		return nil
	}
	// 存储结果
	result := make([]map[string]any, 0, len(data))
	// 按 groupField 分组
	grouped := make(map[string][]map[string]any)
	order := make([]string, 0)

	for _, item := range data {
		key := fmt.Sprintf("%v", item[groupField])
		if _, ok := grouped[key]; !ok {
			order = append(order, key)
		}
		grouped[key] = append(grouped[key], item)
	}

	// 组内处理
	for _, key := range order {
		items := grouped[key]
		startRow := dataStartRow + len(result)

		for index, item := range items {
			item["_row_index"] = index
			result = append(result, item)
		}

		endRow := dataStartRow + len(result) - 1
		formula := fmt.Sprintf("=SUBTOTAL(%d,%s%d:%s%d)",
			subtotal.Code,
			subtotalCellLetter, startRow,
			subtotalCellLetter, endRow)

		result = append(result, map[string]any{
			groupField:  fmt.Sprintf("%s %s", key, subtotal.GroupFieldSuffix),
			sumField:    formula,
			"_row_type": "subtotal",
		})
	}

	// 整体总计
	firstDataRow := dataStartRow
	lastDataRow := dataStartRow + len(result) - 1
	totalFormula := fmt.Sprintf("=SUBTOTAL(%d,%s%d:%s%d)",
		subtotal.Code,
		subtotalCellLetter, firstDataRow,
		subtotalCellLetter, lastDataRow)

	result = append(result, map[string]any{
		groupField:  subtotal.TotalField,
		sumField:    totalFormula,
		"_row_type": "subtotal",
	})

	return result
}
