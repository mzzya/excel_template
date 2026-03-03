# Excel模板引擎

基于[excelize](https://github.com/qax-os/excelize)的Excel模板渲染引擎，支持数据绑定、公式计算、颜色设置和分类汇总等功能。

## 功能特性

- **模板渲染**: 支持Go模板语法，动态填充Excel数据
- **公式计算**: 集成HyperFormula引擎进行公式计算
- **数据绑定**: 将数据映射到Excel模板中的指定位置
- **样式控制**: 支持背景色和字体颜色的动态设置
- **分类汇总**: 支持按字段对数据进行分组和统计
- **图片处理**: 支持图片转Base64和解析功能
- **多语言支持**: 支持中文和英文界面

## 安装

```bash
go mod tidy
```

## 使用说明

### 基本用法

有关具体使用方法，请参见 [render_test.go](./render_test.go) 和 [hyperformula_test.go](./hyperformula_test.go) 文件中的测试示例。

### 模板语法

在Excel模板中可以使用以下特殊标识符：

- `{{.FieldName}}`: 基本数据绑定
- `Header`: 表头定义
- `DataField`: 数据字段映射
- `Data`: 数据内容
- `BackgroundColor`: 背景色表达式
- `FontColor`: 字体颜色表达式
- `Subtotal`: 分类汇总标记

### 颜色设置

支持通过表达式动态设置单元格颜色。

### 分类汇总

使用分类汇总功能对数据进行分组统计，详情请参考 [render_test.go](./render_test.go) 文件中的示例。

### 公式处理

支持Excel公式的动态处理和行号替换，具体用法请查看 [formula.go](./formula.go) 文件。

### 图片处理

支持图片与Base64数据URI之间的转换，详情请参考 [image.go](./image.go) 文件。

## 开发说明

### 项目结构

```
.
├── constant/              # 常量定义
│   └── language.go        # 语言相关的常量
├── formula.go             # 公式处理相关函数
├── hyperformula.go        # HyperFormula引擎实现
├── image.go               # 图片处理功能
├── render.go              # 核心渲染逻辑
├── render_test.go         # 渲染功能测试
├── hyperformula_test.go   # HyperFormula引擎测试
├── subtotal.go            # 分类汇总功能
├── template.go            # 模板处理基础函数
├── README.md              # 项目说明文档
└── package.json           # Node.js依赖（可能用于前端集成）
```

### 核心组件

#### ExcelTemplate 结构

主渲染器结构，负责整个Excel模板的渲染过程：

- `TemplatePath`: 模板文件路径
- [File](./render.go#L45-L45): Excel文件对象
- `SheetCache`: 工作表缓存
- `FormulaEngine`: 公式引擎
- `FuncMap`: 模板函数映射
- `ListField`: 列表字段名称

#### FormulaEngine 接口

公式计算引擎接口，目前使用HyperFormula实现：

```go
type FormulaEngine interface {
    EvalFormula(formulaExpr string, data map[string]any) (string, any, error)
}
```

#### SheetCache 结构

工作表缓存结构，优化渲染性能：

- `Config`: 配置信息
- `ColumnList`: 列定义列表
- `StartRowNum`: 起始行号
- `FillData`: 填充数据
- `DataRowHeight`: 数据行高度

### 测试

运行单元测试：

```bash
go test ./...
```

### 架构设计

此项目采用模块化设计，各功能组件职责明确：

1. **模板解析**: 解析Excel模板中的特殊标记和配置
2. **数据绑定**: 将输入数据映射到模板中的占位符
3. **公式计算**: 使用外部引擎计算Excel公式
4. **样式应用**: 根据条件动态设置单元格样式
5. **输出生成**: 生成最终的Excel文件

### 性能优化

- 使用FormulaEngine池管理公式计算资源
- 缓存公式计算结果避免重复计算
- 对大数据集进行分批处理

## 依赖库

- [excelize](https://github.com/xuri/excelize): Excel操作核心库
- [go-deepcopy](https://github.com/tiendc/go-deepcopy): 深度复制工具
- [lo (Lodash-style)](https://github.com/samber/lo): 实用函数集合
- [HyperFormula](https://hyperformula.handsontable.com/): 公式计算引擎

## 许可证

请参阅项目许可证文件。