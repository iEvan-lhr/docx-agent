package dmlpic

import (
	"encoding/xml"
	"fmt"
	"github.com/iEvan-lhr/docx-agent/common/constants"
	"io"

	"github.com/iEvan-lhr/docx-agent/dml/dmlct"
	"github.com/iEvan-lhr/docx-agent/dml/shapes"
)

type BlipFill struct {
	// 1. Blip
	Blip *Blip `xml:"blip,omitempty"`

	//2.Source Rectangle
	SrcRect *dmlct.RelativeRect `xml:"srcRect,omitempty"`

	// 3. Choice of a:EG_FillModeProperties
	Stretch *shapes.Stretch `xml:"stretch,omitempty"`
	Tile    *shapes.Tile    `xml:"tile,omitempty"`

	//Attributes:
	DPI          *uint32 `xml:"dpi,attr,omitempty"`          //DPI Setting
	RotWithShape *bool   `xml:"rotWithShape,attr,omitempty"` //Rotate With Shape
}

// NewBlipFill creates a new BlipFill with the given relationship ID (rID)
// The rID is used to reference the image in the presentation.
func NewBlipFill(rID string) BlipFill {
	return BlipFill{
		Blip: &Blip{
			EmbedID: rID,
		},
	}
}

func (b BlipFill) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "pic:blipFill"

	if b.DPI != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "dpi"}, Value: fmt.Sprintf("%d", *b.DPI)})
	}

	if b.RotWithShape != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "rotWithShape"}, Value: fmt.Sprintf("%t", *b.RotWithShape)})
	}

	err := e.EncodeToken(start)
	if err != nil {
		return err
	}

	// 1. Blip
	if b.Blip != nil {
		if err := b.Blip.MarshalXML(e, xml.StartElement{Name: xml.Name{Local: "a:blip"}}); err != nil {
			return err
		}
	}

	// 2. SrcRect
	if b.SrcRect != nil {
		if err = b.SrcRect.MarshalXML(e, xml.StartElement{Name: xml.Name{Local: "a:SrcRect"}}); err != nil {
			return err
		}
	}

	// 3. Choice: FillModProperties
	if err = b.Stretch.MarshalXML(e, xml.StartElement{Name: xml.Name{Local: "a:stretch"}}); err != nil {
		return err
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

//type FillModeProps struct {
//	Stretch *shapes.Stretch `xml:"stretch,omitempty"`
//	Tile    *shapes.Tile    `xml:"tile,omitempty"`
//}
//
//func (f FillModeProps) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
//
//	if f.Stretch != nil {
//		return f.Stretch.MarshalXML(e, xml.StartElement{})
//	}
//
//	if f.Tile != nil {
//		return f.Tile.MarshalXML(e, xml.StartElement{})
//	}
//
//	return nil
//}

// UnmarshalXML 实现了 xml.Unmarshaler 接口
func (bf *BlipFill) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			case xml.Name{Space: constants.DrawingMLMainNS, Local: "blip"}:
				bf.Blip = new(Blip)
				if err = d.DecodeElement(bf.Blip, &elem); err != nil {
					return err
				}
			// 这就是“裁剪位置”
			case xml.Name{Space: constants.DrawingMLMainNS, Local: "srcRect"}:
				bf.SrcRect = new(dmlct.RelativeRect)
				if err = d.DecodeElement(bf.SrcRect, &elem); err != nil {
					return err
				}
			// 这里包含了“图片位置”
			case xml.Name{Space: constants.DrawingMLMainNS, Local: "stretch"}:
				bf.Stretch = new(shapes.Stretch)
				if err = d.DecodeElement(bf.Stretch, &elem); err != nil {
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
