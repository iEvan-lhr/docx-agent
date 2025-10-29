package ctypes

import (
	"encoding/xml"
	"github.com/iEvan-lhr/docx-agent/internal"
	"strconv"
)

// PageMargin represents the page margins of a Word document.
type PageMargin struct {
	Left   *int `xml:"left,attr,omitempty"`
	Right  *int `xml:"right,attr,omitempty"`
	Gutter *int `xml:"gutter,attr,omitempty"`
	Header *int `xml:"header,attr,omitempty"`
	Top    *int `xml:"top,attr,omitempty"`
	Footer *int `xml:"footer,attr,omitempty"`
	Bottom *int `xml:"bottom,attr,omitempty"`
}

// MarshalXML implements the xml.Marshaler interface for the PageMargin type.
// It encodes the PageMargin to its corresponding XML representation.
func (p PageMargin) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "w:pgMar"

	start.Attr = []xml.Attr{}

	attrs := []struct {
		local string
		value *int
	}{
		{"w:left", p.Left},
		{"w:right", p.Right},
		{"w:gutter", p.Gutter},
		{"w:header", p.Header},
		{"w:top", p.Top},
		{"w:footer", p.Footer},
		{"w:bottom", p.Bottom},
	}

	for _, attr := range attrs {
		if attr.value == nil {
			continue
		}
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: attr.local}, Value: strconv.Itoa(*attr.value)})
	}

	return e.EncodeElement("", start)
}

type PgBorders struct {
	// 1. Table Cell Top Border
	Top *Border `xml:"top,omitempty"`

	// 2. Table Cell Left Border
	Left *Border `xml:"left,omitempty"`

	// 3. Table Cell Bottom Border
	Bottom *Border `xml:"bottom,omitempty"`

	// 4. Table Cell Right Border
	Right *Border `xml:"right,omitempty"`
}

func (p PgBorders) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "w:pgBorders"

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// 1. Top
	if p.Top != nil {
		if err := p.Top.MarshalXML(e, xml.StartElement{
			Name: xml.Name{Local: "w:top"},
		}); err != nil {
			return err
		}
	}

	// 2. Left
	if p.Left != nil {
		if err := p.Left.MarshalXML(e, xml.StartElement{
			Name: xml.Name{Local: "w:left"},
		}); err != nil {
			return err
		}
	}

	// 3. Bottom
	if p.Bottom != nil {
		if err := p.Bottom.MarshalXML(e, xml.StartElement{
			Name: xml.Name{Local: "w:bottom"},
		}); err != nil {
			return err
		}
	}

	// 4. Right
	if p.Right != nil {
		if err := p.Right.MarshalXML(e, xml.StartElement{
			Name: xml.Name{Local: "w:right"},
		}); err != nil {
			return err
		}
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})

}
func (p *PgBorders) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {

	for {
		currentToken, err := d.Token()
		if err != nil {
			return err
		}

		switch elem := currentToken.(type) {
		case xml.StartElement:
			switch elem.Name.Local {
			case "top":
				p.Top = &Border{}
				if err := d.DecodeElement(p.Top, &elem); err != nil {
					return err
				}
			case "left":
				p.Left = &Border{}
				if err := d.DecodeElement(p.Left, &elem); err != nil {
					return err
				}
			case "right":
				p.Right = &Border{}
				if err := d.DecodeElement(p.Right, &elem); err != nil {
					return err
				}
			case "bottom":
				p.Bottom = &Border{}
				if err := d.DecodeElement(p.Bottom, &elem); err != nil {
					return err
				}

			default:
				// 跳过未知元素，避免解析错误
				if err = d.Skip(); err != nil {
					return err
				}
			}

		case xml.EndElement:
			// 遇到 sectPr 的结束标签，返回
			return nil
		}
	}
}

type Cols struct {
	Space *string `xml:"space,attr,omitempty"`
}

func (c Cols) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "w:cols"

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if c.Space != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "w:space"}, Value: *c.Space})
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})

}
func (c *Cols) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "id" {
			c.Space = internal.ToPtr(attr.Value)
		}

	}
	return d.Skip() // 空元素
}
