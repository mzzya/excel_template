package excel_template

import (
	"encoding/base64"
	"errors"
	"fmt"
	"os"
	"regexp"
	"strings"

	"github.com/gabriel-vasile/mimetype"
)

func ImageToBase64WithMime(filePath string) (string, error) {
	// 读取文件内容
	data, err := os.ReadFile(filePath)
	if err != nil {
		return "", err
	}

	// 自动检测 MIME 类型
	mime := mimetype.Detect(data)

	// 编码 base64
	encoded := base64.StdEncoding.EncodeToString(data)

	// 拼接带 MIME 的 data URI
	dataURI := fmt.Sprintf("data:%s;base64,%s", mime.String(), encoded)

	return dataURI, nil
}

// 提取图片类型和字节内容
func ParseBase64Image(dataURI string) (mimeType string, data []byte, err error) {
	// 示例格式：data:image/png;base64,xxx
	if !strings.HasPrefix(dataURI, "data:image") {
		return "", nil, errors.New("不是 base64 图片字符串")
	}

	re := regexp.MustCompile(`^data:image/(\w+);base64,(.+)$`)
	matches := re.FindStringSubmatch(dataURI)
	if len(matches) != 3 {
		return "", nil, errors.New("无法解析 base64 图片数据")
	}

	mimeType = matches[1]
	base64Data := matches[2]

	data, err = base64.StdEncoding.DecodeString(base64Data)
	if err != nil {
		return "", nil, err
	}

	return mimeType, data, nil
}
