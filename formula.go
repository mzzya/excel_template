package excel_template

import (
	"fmt"
	"regexp"

	"github.com/xuri/excelize/v2"
)

// 替换公式中的行号为目标行号，并偏移列数
func ReplaceFormulaRow(formula string, targetRow int, moveColNum int) string {
	re := regexp.MustCompile(`([$]?)([A-Z]+)([$]?)(\d+)`)
	result := re.ReplaceAllStringFunc(formula, func(match string) string {
		parts := re.FindStringSubmatch(match)
		if len(parts) < 5 {
			return match
		}
		dollarCol := parts[1]
		colLetters := parts[2]
		dollarRow := parts[3]
		// rowNum := parts[4] // 可用于需要验证的场景
		colNum, err := excelize.ColumnNameToNumber(colLetters)
		if err != nil {
			// return match
			panic(fmt.Sprintf("解析列号出错: %v", err))
		}
		newCol, err := excelize.ColumnNumberToName(colNum + moveColNum)
		return fmt.Sprintf("%s%s%s%d", dollarCol, newCol, dollarRow, targetRow)
	})
	return result
}

func ReplaceCellRange(s string, replacement string) string {
	// 匹配类似 H5:H5、A1:B10 的格式
	re := regexp.MustCompile(`[A-Z]+\d+:[A-Z]+\d+`)
	return re.ReplaceAllString(s, replacement)
}
