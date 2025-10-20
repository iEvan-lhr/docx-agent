package ctypes

import (
	"encoding/xml"
)

// Header 代表一个页眉 (hdr.xml) 的内容
type Header struct {
	XMLName    xml.Name     `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main hdr"`
	Paragraphs []*Paragraph `xml:"p"`   // 引用 wml/ctypes/para.go 中的 Paragraph
	Tables     []*Table     `xml:"tbl"` // 引用 wml/ctypes/table.go 中的 Table
}
