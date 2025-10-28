package dmlct

import (
	"encoding/xml"
	"io"
	"strconv"

	"github.com/iEvan-lhr/docx-agent/common/units"
)

// Complex Type: CT_PositiveSize2D
type PSize2D struct {
	Width  uint64 `xml:"cx,attr,omitempty"`
	Height uint64 `xml:"cy,attr,omitempty"`
}

func NewPostvSz2D(width units.Emu, height units.Emu) *PSize2D {
	return &PSize2D{
		Height: uint64(height),
		Width:  uint64(width),
	}
}

func (p PSize2D) MarshalXML(e *xml.Encoder, start xml.StartElement) error {

	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "cx"}, Value: strconv.FormatUint(p.Width, 10)})
	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "cy"}, Value: strconv.FormatUint(p.Height, 10)})

	return e.EncodeElement("", start)
}

// UnmarshalXML 为 PSize2D 实现 xml.Unmarshaler
func (p *PSize2D) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	var err error
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "cx":
			p.Width, err = strconv.ParseUint(attr.Value, 10, 64)
		case "cy":
			p.Height, err = strconv.ParseUint(attr.Value, 10, 64)
		}
		if err != nil {
			//可以选择记录或返回错误
			// return fmt.Errorf("error parsing attribute %s for PSize2D: %w", attr.Name.Local, err)
		}
	}

	// <a:ext> 是空元素，消耗掉 token 直到找到结束标签
	for {
		token, err := d.Token()
		if err != nil {
			if err == io.EOF {
				return nil
			}
			return err
		}
		if elem, ok := token.(xml.EndElement); ok && elem.Name == start.Name {
			return nil
		}
	}
}
