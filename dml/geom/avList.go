package geom

import (
	"encoding/xml"
	"fmt"
	"github.com/iEvan-lhr/docx-agent/common/constants"
	"io"
)

// List of Shape Adjust Values
type AdjustValues struct {
	ShapeGuides []ShapeGuide `xml:"gd,omitempty"`
}

func (a AdjustValues) MarshalXML(e *xml.Encoder, start xml.StartElement) (err error) {
	start.Name.Local = "a:avLst"

	err = e.EncodeToken(start)
	if err != nil {
		return err
	}

	for _, data := range a.ShapeGuides {
		err := data.MarshalXML(e, xml.StartElement{})
		if err != nil {
			return err
		}
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

func (a *AdjustValues) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	a.ShapeGuides = []ShapeGuide{} // 初始化切片

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
			// 检查子元素 <a:gd>
			if elem.Name.Local == "gd" && elem.Name.Space == constants.DrawingMLMainNS {
				var guide ShapeGuide
				// 假设 ShapeGuide 也有 UnmarshalXML
				if err := guide.UnmarshalXML(d, elem); err != nil {
					return fmt.Errorf("unmarshalling ShapeGuide (gd): %w", err)
				}
				a.ShapeGuides = append(a.ShapeGuides, guide)
			} else {
				// 跳过其他不认识的子元素
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			if elem.Name == start.Name {
				break loop // 到达 </a:avLst>
			}
		}
	}
	return nil
}
