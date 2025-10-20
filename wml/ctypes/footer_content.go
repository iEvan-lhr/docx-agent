package ctypes

import (
	"encoding/xml"
)

// Footer 代表一个页脚 (ftr.xml) 的内容
type Footer struct {
	XMLName    xml.Name     `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main ftr"`
	Paragraphs []*Paragraph `xml:"p"`   // 引用 wml/ctypes/para.go 中的 Paragraph
	Tables     []*Table     `xml:"tbl"` // 引用 wml/ctypes/table.go 中的 Table
}
