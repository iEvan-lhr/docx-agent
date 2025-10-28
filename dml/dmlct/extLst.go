package dmlct

import (
	"encoding/xml"
	"io" // 导入 io 包

	"github.com/iEvan-lhr/docx-agent/common/constants" // 导入
)

// --- 新增结构体：A14UseLocalDpi ---
// 对应 <a14:useLocalDpi ...>
type A14UseLocalDpi struct {
	XmlnsA14 string `xml:"xmlns:a14,omitempty"`
	Val      string `xml:"val,omitempty"`
}

// MarshalXML 实现了 <a14:useLocalDpi> 的手动序列化
func (u A14UseLocalDpi) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "a14:useLocalDpi"
	start.Attr = []xml.Attr{
		// 显式定义命名空间
		{Name: xml.Name{Local: "xmlns:a14"}, Value: "http://schemas.microsoft.com/office/drawing/2010/main"},
		{Name: xml.Name{Local: "val"}, Value: u.Val},
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

// UnmarshalXML 实现了 <a14:useLocalDpi> 的手动反序列化
func (u *A14UseLocalDpi) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "val" {
			u.Val = attr.Value
		}
		// 存储 xmlns:a14 的值 (尽管在 Marshal 时我们是硬编码的)
		if attr.Name.Local == "a14" && attr.Name.Space == "xmlns" {
			u.XmlnsA14 = attr.Value
		}
	}

	// 这是一个空元素，循环直到找到它的结束标签
	for {
		token, err := d.Token()
		if err != nil {
			if err == io.EOF {
				return nil
			} // io.EOF 也意味着结束
			return err
		}
		if elem, ok := token.(xml.EndElement); ok && elem.Name == start.Name {
			return nil
		}
	}
}

// --- 新增结构体：AExt ---
// 对应 <a:ext ...>
type Ext struct {
	URI         string          `xml:"uri,omitempty"`
	UseLocalDpi *A14UseLocalDpi `xml:"a14:useLocalDpi,omitempty"`
}

// MarshalXML 实现了 <a:ext> 的手动序列化
func (a Ext) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "a:ext"
	start.Attr = []xml.Attr{
		{Name: xml.Name{Local: "uri"}, Value: a.URI},
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// 序列化子元素
	if a.UseLocalDpi != nil {
		if err := a.UseLocalDpi.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

// UnmarshalXML 实现了 <a:ext> 的手动反序列化
func (a *Ext) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "uri" {
			a.URI = attr.Value
		}
	}

loop:
	for {
		token, err := d.Token()
		if err != nil {
			if err == io.EOF {
				break loop
			}
			return err
		}

		switch elem := token.(type) {
		case xml.StartElement:
			// 检查 <a14:useLocalDpi>
			// 注意：Go 的 XML 解析器会自动处理命名空间 URI
			if elem.Name.Local == "useLocalDpi" && elem.Name.Space == "http://schemas.microsoft.com/office/drawing/2010/main" {
				a.UseLocalDpi = new(A14UseLocalDpi)
				if err := a.UseLocalDpi.UnmarshalXML(d, elem); err != nil {
					return err
				}
			} else {
				// 跳过其他不认识的子元素
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			if elem.Name == start.Name {
				break loop // 到达 </a:ext>
			}
		}
	}
	return nil
}

// --- 新增结构体：AExtLst ---
// 对应 <a:extLst ...>
type ExtLst struct {
	Exts []Ext `xml:"ext,omitempty"`
}

// MarshalXML 实现了 <a:extLst> 的手动序列化
func (l ExtLst) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "a:extLst"
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	for _, ext := range l.Exts {
		if err := ext.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

// UnmarshalXML 实现了 <a:extLst> 的手动反序列化
func (l *ExtLst) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	l.Exts = []Ext{}
loop:
	for {
		token, err := d.Token()
		if err != nil {
			if err == io.EOF {
				break loop
			}
			return err
		}

		switch elem := token.(type) {
		case xml.StartElement:
			if elem.Name.Local == "ext" && elem.Name.Space == constants.DrawingMLMainNS {
				ext := Ext{}
				if err := ext.UnmarshalXML(d, elem); err != nil {
					return err
				}
				l.Exts = append(l.Exts, ext)
			} else {
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			if elem.Name == start.Name {
				break loop // 到达 </a:extLst>
			}
		}
	}
	return nil
}
