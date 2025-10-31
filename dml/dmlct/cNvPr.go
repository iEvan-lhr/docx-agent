package dmlct

import (
	"encoding/xml"
	"io"
	"strconv"
)

// Non-Visual Drawing Properties
type CNvPr struct {
	ID   uint   `xml:"id,attr,omitempty"`
	Name string `xml:"name,attr,omitempty"`

	//Alternative Text for Object - Default value is "".
	Description string `xml:"descr,attr,omitempty"`

	// Hidden - Default value is "false".
	Hidden *bool `xml:"hidden,attr,omitempty"`

	//TODO: implement child elements
	// Sequence [1..1]
	// a:hlinkClick [0..1]    Drawing Element On Click Hyperlink
	// a:hlinkHover [0..1]    Hyperlink for Hover
	// a:extLst [0..1]    Extension List
}

func NewNonVisProp(id uint, name string) *CNvPr {
	return &CNvPr{
		ID:   id,
		Name: name,
	}
}

func (c CNvPr) MarshalXML(e *xml.Encoder, start xml.StartElement) error {

	// ! NOTE: Disabling the empty name check for the Picture
	//  since popular docx tools allow them
	// if c.Name == "" {
	// 	return fmt.Errorf("invalid Name for Non-Visual Drawing Properties when marshaling")
	// }

	start.Attr = []xml.Attr{
		{Name: xml.Name{Local: "id"}, Value: strconv.FormatUint(uint64(c.ID), 10)},
		{Name: xml.Name{Local: "name"}, Value: c.Name},
	}
	if c.Description != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "descr"}, Value: c.Description})
	}

	if c.Hidden != nil {
		if *c.Hidden {
			start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "hidden"}, Value: "true"})
		} else {
			start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "hidden"}, Value: "false"})
		}
	}

	err := e.EncodeToken(start)
	if err != nil {
		return err
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

// UnmarshalXML 为 CNvPr 实现 xml.Unmarshaler
func (c *CNvPr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// 1. 读取属性
	for _, attr := range start.Attr {
		var err error
		switch attr.Name.Local {
		case "id":
			var idVal uint64
			idVal, err = strconv.ParseUint(attr.Value, 10, 0)
			c.ID = uint(idVal) // 注意 uint 转换可能在非常大的 ID 上溢出
		case "name":
			c.Name = attr.Value
		case "descr":
			c.Description = attr.Value
		case "hidden":
			var hiddenVal bool
			hiddenVal, err = strconv.ParseBool(attr.Value)
			c.Hidden = &hiddenVal
		}
		if err != nil {
			// 如果属性解析失败，可以选择返回错误或记录日志
			// return fmt.Errorf("error parsing attribute %s: %w", attr.Name.Local, err)
		}
	}

	// 2. 循环读取子元素 (当前你的结构体没有子元素字段，所以直接跳过)
	// 如果未来添加了 HLinkClick, HLinkHover, ExtLst 字段，需要在这里处理
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
			// 例如:
			// if elem.Name.Local == "hlinkClick" && elem.Name.Space == constants.DrawingMLMainNS {
			//    c.HLinkClick = new(HLink)
			//    if err := c.HLinkClick.UnmarshalXML(d, elem); err != nil { return err }
			// } else ...
			// 否则，跳过
			if err := d.Skip(); err != nil {
				return err
			}
		case xml.EndElement:
			if elem.Name == start.Name {
				break loop // 到达 </pic:cNvPr> 或 </wps:cNvPr>
			}
		}
	}
	return nil
}
