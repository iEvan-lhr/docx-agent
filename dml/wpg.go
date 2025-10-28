// in: dml/wpg.go (新文件)
package dml

import (
	"encoding/xml"
	"fmt"
	"io"

	"github.com/iEvan-lhr/docx-agent/common/constants"
	"github.com/iEvan-lhr/docx-agent/dml/dmlct"
	"github.com/iEvan-lhr/docx-agent/dml/dmlpic"
)

// ==========================================================
// <wpg:wgp> (Wordprocessing Group)
// ==========================================================
type WPGGroup struct {
	CNvGrpSpPr *WPGNonVisualGroupShapeProps `xml:"cnvGrpSpPr,omitempty"`
	GrpSpPr    *WPGGroupShapeProperties     `xml:"grpSpPr,omitempty"`
	Wsp        *WPSWordprocessingShape      `xml:"wsp,omitempty"`
	Pic        *dmlpic.Pic                  `xml:"pic,omitempty"` // 复用 dmlpic 中的 Pic 结构
}

func (g *WPGGroup) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "wpg:wgp"
	// 确保定义 wpg 命名空间 (通常在父元素 a:graphicData 中定义，这里可能不需要重复)
	// start.Attr = append(start.Attr, xml.Attr{
	// 	Name:  xml.Name{Local: "xmlns:wpg"},
	// 	Value: constants.WPGNamespace, // 假设你有这个常量
	// })

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if g.CNvGrpSpPr != nil {
		if err := g.CNvGrpSpPr.MarshalXML(e, xml.StartElement{}); err != nil {
			return fmt.Errorf("marshalling CNvGrpSpPr: %w", err)
		}
	}
	if g.GrpSpPr != nil {
		if err := g.GrpSpPr.MarshalXML(e, xml.StartElement{}); err != nil {
			return fmt.Errorf("marshalling GrpSpPr: %w", err)
		}
	}
	if g.Wsp != nil {
		if err := g.Wsp.MarshalXML(e, xml.StartElement{}); err != nil {
			return fmt.Errorf("marshalling Wsp: %w", err)
		}
	}
	if g.Pic != nil {
		// Pic 可能也有自己的 MarshalXML
		if err := g.Pic.MarshalXML(e, xml.StartElement{}); err != nil {
			return fmt.Errorf("marshalling Pic: %w", err)
		}
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

func (g *WPGGroup) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			// 注意：Go 的 XML 解析器会自动处理命名空间 URI
			// "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
			wpgNS := constants.WPGNamespace // 假设你有这个常量
			wpsNS := constants.WPSNamespace // 假设你有这个常量
			if elem.Name.Local == "cNvGrpSpPr" && elem.Name.Space == wpgNS {
				g.CNvGrpSpPr = new(WPGNonVisualGroupShapeProps)
				if err := g.CNvGrpSpPr.UnmarshalXML(d, elem); err != nil {
					return err
				}
			} else if elem.Name.Local == "grpSpPr" && elem.Name.Space == wpgNS {
				g.GrpSpPr = new(WPGGroupShapeProperties)
				if err := g.GrpSpPr.UnmarshalXML(d, elem); err != nil {
					return err
				}
			} else if elem.Name.Local == "wsp" && elem.Name.Space == wpsNS {
				g.Wsp = new(WPSWordprocessingShape)
				if err := g.Wsp.UnmarshalXML(d, elem); err != nil {

					return err
				}
			} else if elem.Name.Local == "pic" && elem.Name.Space == constants.DrawingMLPicNS { //
				g.Pic = new(dmlpic.Pic)
				// 假设 Pic 也有 UnmarshalXML
				if err := g.Pic.UnmarshalXML(d, elem); err != nil {
					return err
				}
			} else {
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			if elem.Name == start.Name {
				break loop // 到达 </wpg:wgp>
			}
		}
	}
	return nil
}

// ==========================================================
// <wpg:cNvGrpSpPr> (Non-Visual Properties for Group Shape)
// ==========================================================
type WPGNonVisualGroupShapeProps struct{} // 空元素

func (p *WPGNonVisualGroupShapeProps) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "wpg:cNvGrpSpPr"
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

func (p *WPGNonVisualGroupShapeProps) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	return d.Skip() // 空元素，直接跳过内部内容（如果有的话）
}

// ==========================================================
// <wpg:grpSpPr> (Group Shape Properties)
// ==========================================================
type WPGGroupShapeProperties struct {
	Xfrm *dmlpic.TransformGroup `xml:"xfrm,omitempty"`
}

func (p *WPGGroupShapeProperties) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "wpg:grpSpPr"
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if p.Xfrm != nil {
		if err := p.Xfrm.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

func (p *WPGGroupShapeProperties) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			if elem.Name.Local == "xfrm" && elem.Name.Space == constants.DrawingMLMainNS {
				p.Xfrm = new(dmlpic.TransformGroup)
				if err := p.Xfrm.UnmarshalXML(d, elem); err != nil {
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

// ==========================================================
// <a:chOff> (Child Offset)
// ==========================================================
type ChildOffset struct {
	X string `xml:"x,attr"`
	Y string `xml:"y,attr"`
}

func (o *ChildOffset) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "a:chOff"
	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "x"}, Value: o.X})
	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "y"}, Value: o.Y})
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}
func (o *ChildOffset) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "x" {
			o.X = attr.Value
		}
		if attr.Name.Local == "y" {
			o.Y = attr.Value
		}
	}
	return d.Skip() // 空元素
}

// ==========================================================
// <a:chExt> (Child Extents)
// ==========================================================
type AChildExtents struct {
	Cx string `xml:"cx,attr"`
	Cy string `xml:"cy,attr"`
}

func (o *AChildExtents) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "a:chExt"
	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "cx"}, Value: o.Cx})
	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "cy"}, Value: o.Cy})
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}
func (o *AChildExtents) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "cx" {
			o.Cx = attr.Value
		}
		if attr.Name.Local == "cy" {
			o.Cy = attr.Value
		}
	}
	return d.Skip() // 空元素
}

// ==========================================================
// <wps:wsp> (Wordprocessing Shape)
// ==========================================================
type WPSWordprocessingShape struct {
	CNvPr   *dmlct.CNvPr
	CNvSpPr *WPSNonVisualShapeDrawingProps
	SpPr    *WPSShapeProperties
	BodyPr  *ABodyProperties
}

func (s *WPSWordprocessingShape) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "wps:wsp"
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if s.CNvPr != nil {
		if err := s.CNvPr.MarshalXML(e, xml.StartElement{
			Name: xml.Name{Local: "wps:cNvPr"},
		}); err != nil {
			return err
		}
	}
	if s.CNvSpPr != nil {
		if err := s.CNvSpPr.MarshalXML(e, xml.StartElement{
			Name: xml.Name{Local: "wps:cNvSpPr"},
		}); err != nil {
			return err
		}
	}
	if s.SpPr != nil {
		if err := s.SpPr.MarshalXML(e, xml.StartElement{
			Name: xml.Name{Local: "wps:cSpPr"},
		}); err != nil {
			return err
		}
	}
	if s.BodyPr != nil {
		if err := s.BodyPr.MarshalXML(e, xml.StartElement{
			Name: xml.Name{Local: "wps:cBodyPr"},
		}); err != nil {
			return err
		}
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}
func (s *WPSWordprocessingShape) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			ns := constants.WPSNamespace // 假设有
			if elem.Name.Local == "cNvPr" && elem.Name.Space == ns {
				s.CNvPr = new(dmlct.CNvPr)
				if err := s.CNvPr.UnmarshalXML(d, elem); err != nil {
					return err
				}
			} else if elem.Name.Local == "cNvSpPr" && elem.Name.Space == ns {
				s.CNvSpPr = new(WPSNonVisualShapeDrawingProps)
				if err := s.CNvSpPr.UnmarshalXML(d, elem); err != nil {
					return err
				}
			} else if elem.Name.Local == "spPr" && elem.Name.Space == ns {
				s.SpPr = new(WPSShapeProperties)
				if err := s.SpPr.UnmarshalXML(d, elem); err != nil {
					return err
				}
			} else if elem.Name.Local == "bodyPr" && elem.Name.Space == ns {
				s.BodyPr = new(ABodyProperties)
				if err := s.BodyPr.UnmarshalXML(d, elem); err != nil {
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

// ==========================================================
// <wps:cNvPr> (Non-Visual Shape Properties)
// ==========================================================
type WPSNonVisualShapeProps struct {
	ID   string `xml:"id,attr"`
	Name string `xml:"name,attr"`
}

func (p *WPSNonVisualShapeProps) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "wps:cNvPr"
	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "id"}, Value: p.ID})
	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "name"}, Value: p.Name})
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}
func (p *WPSNonVisualShapeProps) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "id" {
			p.ID = attr.Value
		}
		if attr.Name.Local == "name" {
			p.Name = attr.Value
		}
	}
	return d.Skip() // 空元素
}

// ==========================================================
// <wps:cNvSpPr> (Non-Visual Shape Drawing Properties)
// ==========================================================
type WPSNonVisualShapeDrawingProps struct {
	SpLocks *ShapeLocking `xml:"spLocks"`
}

func (p *WPSNonVisualShapeDrawingProps) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "wps:cNvSpPr"
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if p.SpLocks != nil {
		if err := p.SpLocks.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}
func (p *WPSNonVisualShapeDrawingProps) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			if elem.Name.Local == "spLocks" && elem.Name.Space == constants.DrawingMLMainNS {
				p.SpLocks = new(ShapeLocking)
				if err := p.SpLocks.UnmarshalXML(d, elem); err != nil {
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

// ==========================================================
// <a:spLocks> (Shape Locking)
// ==========================================================
type ShapeLocking struct {
	NoChangeAspect     string `xml:"noChangeAspect,attr,omitempty"`
	NoChangeArrowheads string `xml:"noChangeArrowheads,attr,omitempty"`
}

func (l *ShapeLocking) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "a:spLocks"
	if l.NoChangeAspect != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "noChangeAspect"}, Value: l.NoChangeAspect})
	}
	if l.NoChangeArrowheads != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "noChangeArrowheads"}, Value: l.NoChangeArrowheads})
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}
func (l *ShapeLocking) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "noChangeAspect" {
			l.NoChangeAspect = attr.Value
		}
		if attr.Name.Local == "noChangeArrowheads" {
			l.NoChangeArrowheads = attr.Value
		}
	}
	return d.Skip() // 空元素
}

// ==========================================================
// <wps:spPr> (Shape Properties)
// ==========================================================
type WPSShapeProperties struct {
	BwMode   *string                `xml:"bwMode,attr,omitempty"`
	Xfrm     *dmlpic.TransformGroup `xml:"xfrm,omitempty"`
	PrstGeom *dmlpic.PresetGeometry `xml:"prstGeom,omitempty"`
	GradFill *GradientFill          `xml:"gradFill,omitempty"`
	Ln       *LineProperties        `xml:"ln,omitempty"`
}

func (p *WPSShapeProperties) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "wps:spPr"
	if p.BwMode != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "bwMode"}, Value: *p.BwMode})
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if p.Xfrm != nil {
		if err := p.Xfrm.MarshalXML(e, xml.StartElement{Name: xml.Name{Local: "a:xfrm"}}); err != nil {
			return err
		}
	}

	if p.PrstGeom != nil {
		if err := p.PrstGeom.MarshalXML(e, xml.StartElement{Name: xml.Name{Local: "a:prstGeom"}}); err != nil {
			return err
		}
	}
	if p.GradFill != nil {
		if err := p.GradFill.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	if p.Ln != nil {
		if err := p.Ln.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

func (p *WPSShapeProperties) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "bwMode" {
			p.BwMode = &attr.Value
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
			ns := constants.DrawingMLMainNS
			if elem.Name.Local == "xfrm" && elem.Name.Space == ns {
				p.Xfrm = new(dmlpic.TransformGroup)
				// 假设 TransformGroup 也有 UnmarshalXML
				if err := p.Xfrm.UnmarshalXML(d, elem); err != nil {
					return err
				}
			} else if elem.Name.Local == "prstGeom" && elem.Name.Space == ns {
				p.PrstGeom = new(dmlpic.PresetGeometry)
				// 假设 PresetGeometry 也有 UnmarshalXML
				if err := p.PrstGeom.UnmarshalXML(d, elem); err != nil {
					return err
				}
			} else if elem.Name.Local == "gradFill" && elem.Name.Space == ns {
				p.GradFill = new(GradientFill)
				if err := p.GradFill.UnmarshalXML(d, elem); err != nil {
					return err
				}
			} else if elem.Name.Local == "ln" && elem.Name.Space == ns {
				p.Ln = new(LineProperties)
				if err := p.Ln.UnmarshalXML(d, elem); err != nil {
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

// ==========================================================
// <a:gradFill> (Gradient Fill)
// ==========================================================
type GradientFill struct {
	RotWithShape string `xml:"rotWithShape,attr,omitempty"`
	GsLst        *AGradientStopList
	Lin          *ALinearGradient
}

func (f *GradientFill) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "a:gradFill"
	if f.RotWithShape != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "rotWithShape"}, Value: f.RotWithShape})
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if f.GsLst != nil {
		if err := f.GsLst.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	if f.Lin != nil {
		if err := f.Lin.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}
func (f *GradientFill) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "rotWithShape" {
			f.RotWithShape = attr.Value
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
			ns := constants.DrawingMLMainNS
			if elem.Name.Local == "gsLst" && elem.Name.Space == ns {
				f.GsLst = new(AGradientStopList)
				if err := f.GsLst.UnmarshalXML(d, elem); err != nil {
					return err
				}
			} else if elem.Name.Local == "lin" && elem.Name.Space == ns {
				f.Lin = new(ALinearGradient)
				if err := f.Lin.UnmarshalXML(d, elem); err != nil {
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

// ==========================================================
// <a:gsLst> (Gradient Stop List)
// ==========================================================
type AGradientStopList struct {
	Gs []*AGradientStop
}

func (l *AGradientStopList) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "a:gsLst"
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	for _, gs := range l.Gs {
		if err := gs.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}
func (l *AGradientStopList) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	l.Gs = []*AGradientStop{}
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
			if elem.Name.Local == "gs" && elem.Name.Space == constants.DrawingMLMainNS {
				gs := new(AGradientStop)
				if err := gs.UnmarshalXML(d, elem); err != nil {
					return err
				}
				l.Gs = append(l.Gs, gs)
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

// ==========================================================
// <a:gs> (Gradient Stop)
// ==========================================================
type AGradientStop struct {
	Pos     string `xml:"pos,attr"`
	SrgbClr *ASrgbColor
}

func (s *AGradientStop) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "a:gs"
	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "pos"}, Value: s.Pos})
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if s.SrgbClr != nil {
		if err := s.SrgbClr.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}
func (s *AGradientStop) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "pos" {
			s.Pos = attr.Value
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
			if elem.Name.Local == "srgbClr" && elem.Name.Space == constants.DrawingMLMainNS {
				s.SrgbClr = new(ASrgbColor)
				if err := s.SrgbClr.UnmarshalXML(d, elem); err != nil {
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

// ==========================================================
// <a:srgbClr> (SRGB Color)
// ==========================================================
type ASrgbColor struct {
	Val string `xml:"val,attr"`
}

func (c *ASrgbColor) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "a:srgbClr"
	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "val"}, Value: c.Val})
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}
func (c *ASrgbColor) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "val" {
			c.Val = attr.Value
		}
	}
	return d.Skip() // 空元素
}

// ==========================================================
// <a:lin> (Linear Gradient)
// ==========================================================
type ALinearGradient struct {
	Ang    string `xml:"ang,attr,omitempty"`
	Scaled string `xml:"scaled,attr,omitempty"`
}

func (l *ALinearGradient) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "a:lin"
	if l.Ang != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "ang"}, Value: l.Ang})
	}
	if l.Scaled != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "scaled"}, Value: l.Scaled})
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}
func (l *ALinearGradient) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "ang" {
			l.Ang = attr.Value
		}
		if attr.Name.Local == "scaled" {
			l.Scaled = attr.Value
		}
	}
	return d.Skip() // 空元素
}

// ==========================================================
// <a:ln> (Line Properties)
// ==========================================================
type LineProperties struct {
	NoFill *dmlpic.NoFill // 复用
}

func (l *LineProperties) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "a:ln"
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if l.NoFill != nil {
		if err := l.NoFill.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}
func (l *LineProperties) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			if elem.Name.Local == "noFill" && elem.Name.Space == constants.DrawingMLMainNS {
				l.NoFill = new(dmlpic.NoFill)
				if err := l.NoFill.UnmarshalXML(d, elem); err != nil {
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

// ==========================================================
// <wps:bodyPr> (Body Properties)
// ==========================================================
type ABodyProperties struct {
	Rot       string `xml:"rot,attr,omitempty"`
	Vert      string `xml:"vert,attr,omitempty"`
	Wrap      string `xml:"wrap,attr,omitempty"`
	LIns      string `xml:"lIns,attr,omitempty"`
	TIns      string `xml:"tIns,attr,omitempty"`
	RIns      string `xml:"rIns,attr,omitempty"`
	BIns      string `xml:"bIns,attr,omitempty"`
	Anchor    string `xml:"anchor,attr,omitempty"`
	AnchorCtr string `xml:"anchorCtr,attr,omitempty"`
	Upright   string `xml:"upright,attr,omitempty"`
	NoAutofit *ANoAutofit
}

func (p *ABodyProperties) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "wps:bodyPr"
	if p.Rot != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "rot"}, Value: p.Rot})
	}
	if p.Vert != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "vert"}, Value: p.Vert})
	}
	if p.Wrap != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "wrap"}, Value: p.Wrap})
	}
	if p.LIns != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "lIns"}, Value: p.LIns})
	}
	if p.TIns != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "tIns"}, Value: p.TIns})
	}
	if p.RIns != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "rIns"}, Value: p.RIns})
	}
	if p.BIns != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "bIns"}, Value: p.BIns})
	}
	if p.Anchor != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "anchor"}, Value: p.Anchor})
	}
	if p.AnchorCtr != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "anchorCtr"}, Value: p.AnchorCtr})
	}
	if p.Upright != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "upright"}, Value: p.Upright})
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if p.NoAutofit != nil {
		if err := p.NoAutofit.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}
func (p *ABodyProperties) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "rot":
			p.Rot = attr.Value
		case "vert":
			p.Vert = attr.Value
		case "wrap":
			p.Wrap = attr.Value
		case "lIns":
			p.LIns = attr.Value
		case "tIns":
			p.TIns = attr.Value
		case "rIns":
			p.RIns = attr.Value
		case "bIns":
			p.BIns = attr.Value
		case "anchor":
			p.Anchor = attr.Value
		case "anchorCtr":
			p.AnchorCtr = attr.Value
		case "upright":
			p.Upright = attr.Value
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
			if elem.Name.Local == "noAutofit" && elem.Name.Space == constants.DrawingMLMainNS {
				p.NoAutofit = new(ANoAutofit)
				if err := p.NoAutofit.UnmarshalXML(d, elem); err != nil {
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

// ==========================================================
// <a:noAutofit> (No Autofit)
// ==========================================================
type ANoAutofit struct{}

func (a *ANoAutofit) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "a:noAutofit"
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}
func (a *ANoAutofit) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	return d.Skip() // 空元素
}
