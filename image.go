package excel_template

import (
	"bytes"
	"encoding/base64"
	"errors"
	"fmt"
	"image"
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

// IsBase64Image 检查字符串是否为base64图片格式
func IsBase64Image(value string) bool {
	return strings.HasPrefix(value, "data:image/") && strings.Contains(value, ";base64,")
}

// GetImageExtension 根据MIME类型获取图片扩展名
func GetImageExtension(mimeType string) string {
	switch mimeType {
	case "jpeg":
		return ".jpeg"
	case "jpg":
		return ".jpg"
	case "png":
		return ".png"
	case "gif":
		return ".gif"
	case "bmp":
		return ".bmp"
	default:
		return ".png" // 默认使用 png 格式
	}
}

// ProcessImageData 处理base64图片数据，返回扩展名、解码后的图片数据以及图片的配置信息
func ProcessImageData(value string) (extension string, imageData []byte, config image.Config, err error) {
	if !IsBase64Image(value) {
		return "", nil, image.Config{}, errors.New("不是base64图片格式")
	}

	mimeType, data, err := ParseBase64Image(value)
	if err != nil {
		return "", nil, image.Config{}, err
	}

	extension = GetImageExtension(mimeType)
	imageData = data

	// 获取图片配置信息
	config, _, err = image.DecodeConfig(bytes.NewReader(data))
	if err != nil {
		return "", nil, image.Config{}, fmt.Errorf("无法获取图片配置信息: %w", err)
	}

	return extension, imageData, config, nil
}