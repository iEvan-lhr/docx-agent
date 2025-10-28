package dmlpic

import (
	"encoding/xml"
	"fmt"
	"io"
	"strconv"

	"github.com/iEvan-lhr/docx-agent/common/constants"
	"github.com/iEvan-lhr/docx-agent/common/units"
	"github.com/iEvan-lhr/docx-agent/dml/dmlct"
	"github.com/iEvan-lhr/docx-agent/dml/geom"
	"github.com/iEvan-lhr/docx-agent/dml/shapes"
)

type Pic struct {
	// 1. Non-Visual Picture Properties
	NonVisualPicProp NonVisualPicProp `xml:"nvPicPr,omitempty"`

	// 2.Picture Fill
	BlipFill BlipFill `xml:"blipFill,omitempty"`

	// 3.Shape Properties
	PicShapeProp PicShapeProp `xml:"spPr,omitempty"`
}

func NewPic(rID string, imgCount uint, width units.Emu, height units.Emu) *Pic {
	shapeProp := NewPicShapeProp(
		WithTransformGroup(
			WithTFExtent(width, height),
		),
		WithPrstGeom("rect"),
	)

	nvPicProp := DefaultNVPicProp(imgCount, fmt.Sprintf("Image%v", imgCount))

	blipFill := NewBlipFill(rID)

	blipFill.Stretch = &shapes.Stretch{
		FillRect: &dmlct.RelativeRect{},
	}

	return &Pic{
		BlipFill:         blipFill,
		NonVisualPicProp: nvPicProp,
		PicShapeProp:     *shapeProp,
	}
}

func (p Pic) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "pic:pic"

	start.Attr = []xml.Attr{
		{Name: xml.Name{Local: "xmlns:pic"}, Value: constants.DrawingMLPicNS},
	}

	err := e.EncodeToken(start)
	if err != nil {
		return err
	}

	// 1. nvPicPr
	if err = p.NonVisualPicProp.MarshalXML(e, xml.StartElement{
		Name: xml.Name{Local: "pic:nvPicPr"},
	}); err != nil {
		return fmt.Errorf("marshalling NonVisualPicProp: %w", err)
	}

	// 2. BlipFill
	if err = p.BlipFill.MarshalXML(e, xml.StartElement{
		Name: xml.Name{Local: "pic:blipFill"},
	}); err != nil {
		return fmt.Errorf("marshalling BlipFill: %w", err)
	}

	// 3. spPr
	if err = p.PicShapeProp.MarshalXML(e, xml.StartElement{
		Name: xml.Name{Local: "pic:spPr"},
	}); err != nil {
		return fmt.Errorf("marshalling PicShapeProp: %w", err)
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

// UnmarshalXML implements the xml.Unmarshaler interface for Pic
func (p *Pic) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// Pic 元素本身没有属性需要读取

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
			// 检查子元素的命名空间和本地名
			ns := constants.DrawingMLPicNS // pic 命名空间
			if elem.Name.Local == "nvPicPr" && elem.Name.Space == ns {
				// NonVisualPicProp 结构体需要有自己的 UnmarshalXML 方法
				if err := p.NonVisualPicProp.UnmarshalXML(d, elem); err != nil {
					return err
				}
			} else if elem.Name.Local == "blipFill" && elem.Name.Space == ns {
				// BlipFill 结构体需要有自己的 UnmarshalXML 方法
				if err := p.BlipFill.UnmarshalXML(d, elem); err != nil {
					return err
				}
			} else if elem.Name.Local == "spPr" && elem.Name.Space == ns {
				// PicShapeProp 结构体需要有自己的 UnmarshalXML 方法
				if err := p.PicShapeProp.UnmarshalXML(d, elem); err != nil {
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
				break loop // 到达 </pic:pic>
			}
		}
	}
	return nil
}

type TransformGroup struct {
	Extent   *dmlct.PSize2D `xml:"ext,omitempty"`
	Offset   *Offset        `xml:"off,omitempty"`
	CHExtent *dmlct.PSize2D `xml:"chExt,omitempty"`
	CHOffset *Offset        `xml:"chOff,omitempty"`
}

type TFGroupOption func(*TransformGroup)

func NewTransformGroup(options ...TFGroupOption) *TransformGroup {
	tf := &TransformGroup{}

	for _, opt := range options {
		opt(tf)
	}

	return tf
}

func WithTFExtent(width units.Emu, height units.Emu) TFGroupOption {
	return func(tf *TransformGroup) {
		tf.Extent = dmlct.NewPostvSz2D(width, height)
	}
}

func (t TransformGroup) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "a:xfrm"

	err := e.EncodeToken(start)
	if err != nil {
		return err
	}

	if t.Offset != nil {
		if err := e.EncodeElement(t.Offset, xml.StartElement{Name: xml.Name{Local: "a:off"}}); err != nil {
			return err
		}
	}

	if t.Extent != nil {
		if err := e.EncodeElement(t.Extent, xml.StartElement{Name: xml.Name{Local: "a:ext"}}); err != nil {
			return err
		}
	}
	if t.CHOffset != nil {
		if err := e.EncodeElement(t.CHOffset, xml.StartElement{Name: xml.Name{Local: "a:chOff"}}); err != nil {
			return err
		}
	}

	if t.CHExtent != nil {
		if err := e.EncodeElement(t.CHExtent, xml.StartElement{Name: xml.Name{Local: "a:chExt"}}); err != nil {
			return err
		}
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

type Offset struct {
	X uint64 `xml:"x,attr,omitempty"`
	Y uint64 `xml:"y,attr,omitempty"`
}

func (o Offset) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	//start.Name.Local = "a:off"
	start.Attr = []xml.Attr{
		{Name: xml.Name{Local: "x"}, Value: strconv.FormatUint(o.X, 10)},
		{Name: xml.Name{Local: "y"}, Value: strconv.FormatUint(o.Y, 10)},
	}

	err := e.EncodeToken(start)
	if err != nil {
		return err
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

// UnmarshalXML 为 Offset 实现 xml.Unmarshaler
func (o *Offset) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	var err error
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "x":
			o.X, err = strconv.ParseUint(attr.Value, 10, 64)
		case "y":
			o.Y, err = strconv.ParseUint(attr.Value, 10, 64)
		}
		if err != nil {
			//可以选择记录或返回错误
			// return fmt.Errorf("error parsing attribute %s for Offset: %w", attr.Name.Local, err)
		}
	}

	// <a:off> 是空元素，消耗掉 token 直到找到结束标签
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

type PresetGeometry struct {
	Preset       string             `xml:"prst,attr,omitempty"`
	AdjustValues *geom.AdjustValues `xml:"avLst,omitempty"`
}

func NewPresetGeom(preset string) *PresetGeometry {
	return &PresetGeometry{
		Preset: preset,
	}
}

func (p PresetGeometry) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "a:prstGeom"
	start.Attr = []xml.Attr{}

	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "prst"}, Value: p.Preset})

	err := e.EncodeToken(start)
	if err != nil {
		return err
	}

	if p.AdjustValues != nil {
		if err := e.EncodeElement(p.AdjustValues, xml.StartElement{Name: xml.Name{Local: "a:avLst"}}); err != nil {
			return err
		}
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

func (g *PresetGeometry) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "prst" {
			g.Preset = attr.Value
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
			if elem.Name.Local == "avLst" && elem.Name.Space == constants.DrawingMLMainNS {
				g.AdjustValues = new(geom.AdjustValues)
				if err := g.AdjustValues.UnmarshalXML(d, elem); err != nil {
					return err
				} // AvLst 需要 UnmarshalXML
			} else {
				if err := d.Skip(); err != nil {
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
func (t *TransformGroup) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			ns := constants.DrawingMLMainNS
			if elem.Name.Local == "off" && elem.Name.Space == ns {
				t.Offset = new(Offset)
				if err := t.Offset.UnmarshalXML(d, elem); err != nil {
					return err
				}
			} else if elem.Name.Local == "ext" && elem.Name.Space == ns {
				t.Extent = new(dmlct.PSize2D)
				if err := t.Extent.UnmarshalXML(d, elem); err != nil {
					return err
				}
			} else if elem.Name.Local == "chOff" && elem.Name.Space == ns {
				t.CHOffset = new(Offset)
				if err := t.CHOffset.UnmarshalXML(d, elem); err != nil {
					return err
				}
			} else if elem.Name.Local == "chExt" && elem.Name.Space == ns {
				t.CHExtent = new(dmlct.PSize2D)
				if err := t.CHExtent.UnmarshalXML(d, elem); err != nil {
					return err
				}
			} else {
				if err := d.Skip(); err != nil {
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

// UnmarshalXML 为 PicShapeProp 实现 xml.Unmarshaler
func (p *PicShapeProp) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// 1. 读取属性
	for _, attr := range start.Attr {
		if attr.Name.Local == "bwMode" {
			p.BwMode = &attr.Value
		}
	}

	// 2. 循环读取子元素
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
			// 检查子元素的命名空间和本地名
			ns := constants.DrawingMLMainNS // "a:" 命名空间
			if elem.Name.Local == "xfrm" && elem.Name.Space == ns {
				// 假设 TransformGroup 有 UnmarshalXML
				p.TransformGroup = NewTransformGroup() // 初始化
				if err := p.TransformGroup.UnmarshalXML(d, elem); err != nil {
					return fmt.Errorf("unmarshalling TransformGroup: %w", err)
				}
			} else if elem.Name.Local == "prstGeom" && elem.Name.Space == ns {
				// 假设 PresetGeometry 有 UnmarshalXML
				p.PresetGeometry = NewPresetGeom("") // 初始化
				if err := p.PresetGeometry.UnmarshalXML(d, elem); err != nil {
					return fmt.Errorf("unmarshalling PresetGeometry: %w", err)
				}
			} else if elem.Name.Local == "noFill" && elem.Name.Space == ns {
				p.NoFill = new(NoFill)
				if err := p.NoFill.UnmarshalXML(d, elem); err != nil { // NoFill 的 UnmarshalXML 已提供
					return fmt.Errorf("unmarshalling NoFill: %w", err)
				}
			} else {
				// 跳过其他不认识的子元素
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			if elem.Name == start.Name {
				break loop // 到达 </pic:spPr>
			}
		}
	}
	return nil
}
