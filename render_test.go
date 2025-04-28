package excel_template

import (
	"fmt"
	"math/rand"
	"os"
	"strings"
	"sync"
	"testing"
	"text/template"
	"time"

	"runtime/pprof"

	"github.com/xuri/excelize/v2"
)

func generateRandomData(i int) map[string]any {
	companName := []string{"恒张三世信息科技有限公司", "宏李四网络技术有限公司", "鑫航王五据服务有限公司"}
	names := []string{"张三", "李四", "王五"}
	YN := []string{"是", "否"}
	return map[string]any{
		"订单号":  fmt.Sprintf("order%07d_%d", rand.Intn(1000000), i),
		"客户名称": names[rand.Intn(len(names))],
		"成本中心": fmt.Sprintf("CC%03d", rand.Intn(3)),
		"公司名称": companName[rand.Intn(len(companName))],
		"含税金额": float64(rand.Intn(10000)) + rand.Float64(),
		"未税金额": float64(rand.Intn(10000)) + rand.Float64(),
		"是否签收": YN[rand.Intn(len(YN))],
		"数量":   rand.Intn(100),
		"下单时间": time.Now().Format("2006-01-02 15:04:05"),
		"签收时间": time.Now().Format("2006-01-02"),
	}
}
func TestRender(t *testing.T) {

	// pool, err := NewGojaPool(20)
	// if err != nil {
	// 	log.Fatal().Err(err).Msg("创建公式引擎失败")
	// 	return
	// }

	// SetFormulaEngine(pool)

	var fillData = make(map[string]any)
	data := make([]map[string]any, 0, 100)

	for i := range 20 {
		data = append(data, generateRandomData(i+1))
	}
	fillData["table"] = data
	fillData["总金额"] = 10000000
	fillData["对账日期"] = time.Now().Format("2006年01月02日")
	fillData["生成日期"] = time.Now().Format("2006-01-02")
	imageBase64, _ := ImageToBase64WithMime("barcode.png")
	fillData["条形码"] = imageBase64

	// 创建 CPU profile 文件
	pf, err := os.Create("cpu.pprof")
	if err != nil {
		t.Fatal("无法创建 pprof 文件: ", err)
	}
	pprof.StartCPUProfile(pf)
	defer pprof.StopCPUProfile()

	startTime := time.Now()
	// frmulaPool, err := NewGojaPool(10)
	et, err := OpenFile("template/template.xlsx")
	et.FuncMap = template.FuncMap{
		"toUpper": strings.ToUpper, // 直接用标准库方法
		"repeat": func(s string, count int) string {
			return strings.Repeat(s, count)
		},
	}
	f, err := et.Render(fillData)
	if err != nil {
		t.Fail()
	}
	f.SaveAs("dist/output.xlsx")
	t.Logf("程序运行时间：%s", time.Since(startTime))
}

func TestExcelize(t *testing.T) {
	f := excelize.NewFile()
	startTime := time.Now()
	var wg sync.WaitGroup
	taskCh := make(chan int, 100)

	workerNum := 20
	for range workerNum {
		wg.Add(1)
		go func() {
			defer wg.Done()
			for i := range taskCh {
				err := f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i+1), fmt.Sprintf("Value %d", i+1))
				if err != nil {
					t.Fatal(err)
				}
			}
		}()
	}

	for i := range 100000 {
		taskCh <- i
	}
	close(taskCh)
	wg.Wait()
	err := f.SaveAs("output.xlsx")
	if err != nil {
		t.Fatal(err)
	}
	t.Logf("程序运行时间：%s", time.Since(startTime))
}
