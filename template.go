package excel_template

import (
	"bytes"
	"regexp"
	"text/template"
)

func ContainsGoTemplateSyntax(s string) bool {
	re := regexp.MustCompile(`{{[^{}]+}}`)
	return re.MatchString(s)
}

func RenderTemplate(tmplStr string, data any, funcMap template.FuncMap) (string, error) {
	tmpl := template.New("template").Funcs(funcMap)

	tmpl, err := tmpl.Parse(tmplStr)
	if err != nil {
		return "", err
	}

	var buf bytes.Buffer
	err = tmpl.Execute(&buf, data)
	if err != nil {
		return "", err
	}

	return buf.String(), nil
}
