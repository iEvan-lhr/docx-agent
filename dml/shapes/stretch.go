package shapes

import (
	"encoding/xml"
	"github.com/iEvan-lhr/docx-agent/common/constants"
	"io"

	"github.com/iEvan-lhr/docx-agent/dml/dmlct"
)

type Stretch struct {
	FillRect *dmlct.RelativeRect `xml:"fillRect,omitempty"`
}

func (s Stretch) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "a:stretch"

	err := e.EncodeToken(start)
	if err != nil {
		return err
	}

	if s.FillRect != nil {
		if err := s.FillRect.MarshalXML(e, xml.StartElement{Name: xml.Name{Local: "a:fillRect"}}); err != nil {
			return err
		}
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

// UnmarshalXML 实现了 xml.Unmarshaler 接口
func (s *Stretch) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
loop:
	for {
		currentToken, err := d.Token()
		if err != nil {
			if err == io.EOF {
				break loop
			}
			return err
		}

		switch elem := currentToken.(type) {
		case xml.StartElement:
			switch elem.Name {
			// 这就是“图片位置”
			case xml.Name{Space: constants.DrawingMLMainNS, Local: "fillRect"}:
				s.FillRect = new(dmlct.RelativeRect)
				if err = d.DecodeElement(s.FillRect, &elem); err != nil {
					return err
				}
			default:
				if err = d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			if elem.Name == start.Name {
				break loop
			}
		}
	}
	return nil
}
