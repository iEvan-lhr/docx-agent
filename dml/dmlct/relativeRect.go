package dmlct

import (
	"encoding/xml"
	"errors"
	"io"
	"strconv"
)

// RelativeRect represents a Relative Rectangle structure with abbreviated attributes.
type RelativeRect struct {
	Top    *int `xml:"t,attr,omitempty"` // Top margin
	Left   *int `xml:"l,attr,omitempty"` // Left margin
	Bottom *int `xml:"b,attr,omitempty"` // Bottom margin
	Right  *int `xml:"r,attr,omitempty"` // Right margin
}

// MarshalXML implements the xml.Marshaler interface for RelativeRect.
func (r RelativeRect) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Attr = []xml.Attr{}

	if r.Top != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "t"}, Value: strconv.Itoa(*r.Top)})
	}

	if r.Left != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "l"}, Value: strconv.Itoa(*r.Left)})
	}

	if r.Bottom != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "b"}, Value: strconv.Itoa(*r.Bottom)})
	}

	if r.Right != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "r"}, Value: strconv.Itoa(*r.Right)})
	}

	return e.EncodeElement("", start)
}

func getInt(val string) (int, error) {
	if val == "" {
		return 0, nil
	}
	i, err := strconv.ParseInt(val, 10, 64)
	return int(i), err
}

// UnmarshalXML 实现了 xml.Unmarshaler 接口
func (r *RelativeRect) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		var i int
		i, err = getInt(attr.Value)
		if err != nil {
			return err
		}

		switch attr.Name.Local {
		case "t":
			r.Top = &i
		case "l":
			r.Left = &i
		case "b":
			r.Bottom = &i
		case "r":
			r.Right = &i
		}
	}

	// 这是一个空元素 (e.g., <a:srcRect l="1000" />), 我们需要消耗掉它的 EndElement
	for {
		tok, err := d.Token()
		if err != nil {
			if err == io.EOF {
				return errors.New("unexpected EOF in RelativeRect")
			}
			return err
		}
		if se, ok := tok.(xml.EndElement); ok {
			if se.Name == start.Name {
				break // 找到结束标签
			}
		}
	}
	return nil
}
