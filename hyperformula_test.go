package excel_template

import (
	"testing"
)

func TestEvalFormula(t *testing.T) {
	engine := NewSimpleFormulaEngine()

	result, _, err := engine.EvalFormula(`=IF(是否签收="是","ffff00","")`, map[string]any{
		"下单时间": "2025-04-24 15:19:20", "公司名称": "鑫航王五据服务有限公司", "含税金额": 5722.399260710206, "客户名称": "李四", "成本中心": "CC000", "数量": 7, "是否签收": "是", "未税金额": 1352.4176198876264, "签收时间": "2025-04-24", "订单号": "MB0887584",
	})
	if err != nil {
		t.Fatalf("EvalFormula error: %v", err)
	}
	if result != "ffff00" {
		t.Errorf("期望结果 ffff00，实际结果 %s", result)
	}

	// 简单加法测试
	result, _, err = engine.EvalFormula("A+B", map[string]any{
		"A": 2,
		"B": 3,
	})
	if err != nil {
		t.Fatalf("EvalFormula error: %v", err)
	}
	if result != "5" {
		t.Errorf("期望结果 5，实际结果 %s", result)
	}

	// 浮点数测试
	result, _, err = engine.EvalFormula("A*B", map[string]any{
		"A": 2.5,
		"B": 4,
	})
	if err != nil {
		t.Fatalf("EvalFormula error: %v", err)
	}
	if result != "10" && result != "10.0" {
		t.Errorf("期望结果 10 或 10.0，实际结果 %s", result)
	}

	// 中文变量名测试
	result, _, err = engine.EvalFormula("数量+单价", map[string]any{
		"数量": 3,
		"单价": 7,
	})
	if err != nil {
		t.Fatalf("EvalFormula error: %v", err)
	}
	if result != "10" {
		t.Errorf("期望结果 10，实际结果 %s", result)
	}
}
