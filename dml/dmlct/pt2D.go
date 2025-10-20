package dmlct

import (
	"encoding/xml"
	"errors"
	"io"
	"strconv"
)

// Wrapping Polygon Point2D
type Point2D struct {
	XAxis uint64 `xml:"x,attr,omitempty"`
	YAxis uint64 `xml:"y,attr,omitempty"`
}

func getUint64(val string) (uint64, error) {
	if val == "" {
		return 0, nil
	}
	return strconv.ParseUint(val, 10, 64)
}
func NewPoint2D(x, y uint64) Point2D {
	return Point2D{
		XAxis: uint64(x),
		YAxis: uint64(y),
	}
}

func (p Point2D) MarshalXML(e *xml.Encoder, start xml.StartElement) error {

	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "x"}, Value: strconv.FormatUint(p.XAxis, 10)})
	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "y"}, Value: strconv.FormatUint(p.YAxis, 10)})

	return e.EncodeElement("", start)
}

func (p *Point2D) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "x":
			p.XAxis, err = getUint64(attr.Value)
		case "y":
			p.YAxis, err = getUint64(attr.Value)
		}
		if err != nil {
			return err
		}
	}

	// 跳过任何可能的内部元素, 以防万一
	for {
		tok, err := d.Token()
		if err != nil {
			if err == io.EOF {
				return errors.New("unexpected EOF in Point2D")
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
