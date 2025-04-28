package excel_template

import (
	_ "embed"
	"fmt"
	"regexp"
	"sync"
	"time"

	"github.com/xuri/excelize/v2"
)

// FormulaEngine 是 hyperformula 虚拟机的接口
type FormulaEngine interface {
	EvalFormula(formulaExpr string, data map[string]any) (string, any, error)
}

type SimpleFormulaEngine struct {
	File *excelize.File
}

func (s SimpleFormulaEngine) EvalFormula(formulaExpr string, data map[string]any) (string, any, error) {
	cellNames := make([]string, 0, 2)
	defer func() {
		for _, cell := range cellNames {
			s.File.SetCellValue("Sheet1", cell, "")
		}
		s.File.SetCellValue("Sheet1", "B1", "")
	}()
	keyMap := make(map[string]string)
	row := 1

	cellNames = append(cellNames, "B1")
	for k, v := range data {
		cellName := fmt.Sprintf("A%d", row)
		cellNames = append(cellNames, cellName)
		keyMap[k] = cellName
		s.File.SetCellValue("Sheet1", cellName, v)
		row++
	}
	// 替换公式中的变量为单元格名
	newExpr := replaceVarsWithCells(formulaExpr, keyMap)
	s.File.SetCellFormula("Sheet1", "B1", newExpr)
	value, err := s.File.CalcCellValue("Sheet1", "B1")
	return value, value, err
}

// 正确的工厂类型
type CreateEngine = func() FormulaEngine

func NewSimpleFormulaEngine() FormulaEngine {
	return SimpleFormulaEngine{File: excelize.NewFile()}
}

type Item struct {
	FormulaEngine FormulaEngine
	timestamp     int64 // 归还时间
}

type FormulaEnginePool struct {
	minPoolSize int
	maxPoolSize int
	mu          sync.Mutex
	items       []Item
	cleanTicker *time.Ticker
}

func NewFormulaEnginePool(minPoolSize, maxPoolSize int, create CreateEngine) *FormulaEnginePool {
	pool := &FormulaEnginePool{
		minPoolSize: minPoolSize,
		maxPoolSize: maxPoolSize,
		items:       make([]Item, 0, maxPoolSize),
		cleanTicker: time.NewTicker(5 * time.Minute),
	}
	// 初始化最小池
	for range minPoolSize {
		pool.items = append(pool.items, Item{FormulaEngine: create(), timestamp: time.Now().Unix()})
	}
	// 启动定时清理协程
	go pool.cleaner()
	return pool
}

func (e *FormulaEnginePool) getFormulaEngine() FormulaEngine {
	e.mu.Lock()
	defer e.mu.Unlock()
	if len(e.items) > 0 {
		pf := e.items[len(e.items)-1]
		e.items = e.items[:len(e.items)-1]
		return pf.FormulaEngine
	}
	return NewSimpleFormulaEngine()
}

// 回收 FormulaEngine 到池中
func (e *FormulaEnginePool) putEngine(engine FormulaEngine) {
	e.mu.Lock()
	defer e.mu.Unlock()
	if len(e.items) < e.maxPoolSize {
		e.items = append(e.items, Item{FormulaEngine: engine, timestamp: time.Now().Unix()})
	}
	// 超出最大池直接丢弃
}

func (e *FormulaEnginePool) cleaner() {
	for range e.cleanTicker.C {
		e.mu.Lock()
		now := time.Now().Unix()
		newPool := make([]Item, 0, e.maxPoolSize)
		for _, pf := range e.items {
			if len(newPool) < e.minPoolSize || now-pf.timestamp < 600 {
				newPool = append(newPool, pf)
			}
			// 超过10分钟未用且池大于min，丢弃
		}
		e.items = newPool
		e.mu.Unlock()
	}
}

// EvalFormula 方法内部调用 getFormulaEngine/putEngine
func (e *FormulaEnginePool) EvalFormula(formulaExpr string, data map[string]any) (string, any, error) {
	engine := e.getFormulaEngine()
	defer e.putEngine(engine)
	return engine.EvalFormula(formulaExpr, data)
}

// 将变量名替换为单元格名
func replaceVarsWithCells(expr string, keyMap map[string]string) string {
	// 匹配变量名（支持中文、英文、数字、下划线）
	re := regexp.MustCompile(`[\p{Han}\w]+`)
	return re.ReplaceAllStringFunc(expr, func(s string) string {
		if cell, ok := keyMap[s]; ok {
			return cell
		}
		return s
	})
}
