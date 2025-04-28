package constant

var (
	Header          = "Header"
	DataField       = "DataField"
	Data            = "Data"
	BackgroundColor = "BackgroundColor"
	FontColor       = "FontColor"
	Subtotal        = "Subtotal"
)

var languageData = map[string]map[string]string{
	"en": {
		"Header":          "Header",
		"DataField":       "DataField",
		"Data":            "Data",
		"BackgroundColor": "BackgroundColor",
		"FontColor":       "FontColor",
		"Subtotal":        "Subtotal",
	},
	"zh": {
		"Header":          "表头",
		"DataField":       "数据字段",
		"Data":            "数据",
		"BackgroundColor": "背景色",
		"FontColor":       "字体色",
		"Subtotal":        "分类汇总",
	},
}

func setLanguage(language string) {
	if m, ok := languageData[language]; ok {
		Header = m["Header"]
		DataField = m["DataField"]
		Data = m["Data"]
		BackgroundColor = m["BackgroundColor"]
		FontColor = m["FontColor"]
		Subtotal = m["Subtotal"]
	}
}

func init() {
	setLanguage("zh")
}
