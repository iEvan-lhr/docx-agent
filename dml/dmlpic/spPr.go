package dmlpic

import (
	"encoding/xml"
	"fmt"
	"io"
)

const (
	BlackWhiteModeClr        = "clr"
	BlackWhiteModeAuto       = "auto"
	BlackWhiteModeGray       = "gray"
	BlackWhiteModeLtGray     = "ltGray"
	BlackWhiteModeInvGray    = "invGray"
	BlackWhiteModeGrayWhite  = "grayWhite"
	BlackWhiteModeBlackGray  = "blackGray"
	BlackWhiteModeBlackWhite = "blackWhite"
	BlackWhiteModeBlack      = "black"
	BlackWhiteModeWhite      = "white"
	BlackWhiteModeHidden     = "hidden"
)

type PicShapeProp struct {
	// -- Attributes --
	//Black and White Mode
	BwMode *string `xml:"bwMode,attr,omitempty"`

	// -- Child Elements --
	//1.2D Transform for Individual Objects
	TransformGroup *TransformGroup `xml:"xfrm,omitempty"`

	// 2. Choice
	//TODO: Modify it as Geometry choice
	PresetGeometry *PresetGeometry `xml:"prstGeom,omitempty"`

	//TODO: Remaining sequcence of elements
	NoFill *NoFill `xml:"noFill,omitempty"`
}

type PicShapePropOption func(*PicShapeProp)

func WithTransformGroup(options ...TFGroupOption) PicShapePropOption {
	return func(p *PicShapeProp) {
		p.TransformGroup = NewTransformGroup(options...)
	}
}

func WithPrstGeom(preset string) PicShapePropOption {
	return func(p *PicShapeProp) {
		p.PresetGeometry = NewPresetGeom(preset)
	}
}

func NewPicShapeProp(options ...PicShapePropOption) *PicShapeProp {
	p := &PicShapeProp{}

	for _, opt := range options {
		opt(p)
	}

	return p
}

func (p PicShapeProp) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "pic:spPr"

	if p.BwMode != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "bwMode"}, Value: *p.BwMode})
	}

	err := e.EncodeToken(start)
	if err != nil {
		return err
	}

	//1. Transform
	if p.TransformGroup != nil {
		if err = p.TransformGroup.MarshalXML(e, xml.StartElement{
			Name: xml.Name{Local: "a:xfrm"},
		}); err != nil {
			return fmt.Errorf("marshalling TransformGroup: %w", err)
		}
	}

	//2. Geometry
	if p.PresetGeometry != nil {

		if err = p.PresetGeometry.MarshalXML(e, xml.StartElement{
			Name: xml.Name{Local: "a:prstGeom"},
		}); err != nil {
			return fmt.Errorf("marshalling PresetGeometry: %w", err)
		}
	}
	// 3. VVVV 新增序列化逻辑 VVVV
	if p.NoFill != nil {
		if err = p.NoFill.MarshalXML(e, xml.StartElement{}); err != nil {
			return fmt.Errorf("marshalling NoFill: %w", err)
		}
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

type NoFill struct{}

// MarshalXML 序列化 <a:noFill/>
func (n NoFill) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "a:noFill"
	// 这是一个空元素，所以立即编码开始和结束
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

// UnmarshalXML 反序列化 <a:noFill/>
func (n *NoFill) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// 这是一个空元素，只需消耗掉 token 直到找到它的结束标签
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
