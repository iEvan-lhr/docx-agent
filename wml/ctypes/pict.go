package ctypes

import (
	"encoding/xml"
	"fmt"
	"github.com/iEvan-lhr/docx-agent/common/constants"
	"io"
)

type Pict struct {
	Shape *Shape `xml:"shape,omitempty"`
	Group *Group `xml:"group,omitempty"` // <v:group>
}

// MarshalXML implements the xml.Marshaler interface.
func (b Pict) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "w:pict"

	err := e.EncodeToken(start)
	if err != nil {
		return err
	}

	if b.Shape != nil {
		if err = b.Shape.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}

	if b.Group != nil {
		if err = b.Group.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

// UnmarshalXML 为 Pict 实现 xml.Unmarshaler
func (p *Pict) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			// 检查子元素 <v:group>
			// VML 命名空间: "urn:schemas-microsoft-com:vml"
			if elem.Name.Local == "group" && elem.Name.Space == constants.XMLNS_V { // 假设有常量
				// --- VVVV 修改 VVVV ---
				p.Group = new(Group)
				if err := p.Group.UnmarshalXML(d, elem); err != nil {
					return fmt.Errorf("unmarshalling Group: %w", err)
				}
				// --- ^^^^ 修改 ^^^^ ---
			} else {
				// 跳过其他不认识的子元素
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			if elem.Name == start.Name {
				break loop // 到达 </w:pict>
			}
		}
	}
	return nil
}

type Group struct {
	AnchorId    string `xml:"anchorId,attr,omitempty"` // w14:anchorId
	ID          string `xml:"id,attr,omitempty"`
	Spid        string `xml:"spid,attr,omitempty"` // o:spid
	Style       string `xml:"style,attr,omitempty"`
	CoordOrigin string `xml:"coordorigin,attr,omitempty"`
	CoordSize   string `xml:"coordsize,attr,omitempty"`
	GfxData     string `xml:"gfxdata,attr,omitempty"` // o:gfxdata

	// 子元素 (根据你的 XML 示例，可以包含 Rect 和 Shape)
	Rect      *Rect      `xml:"rect,omitempty"`  // <v:rect>
	Shape     *VShape    `xml:"shape,omitempty"` // <v:shape> (注意：与你之前的 Shape 不同，这是 VML Shape)
	ShapeType *ShapeType `xml:"shapetype,omitempty"`
}

type Shape struct {
	Type  string `xml:"type,attr,omitempty"`
	Style string `xml:"style,attr,omitempty"`

	ImageData *ImageData `xml:"imagedata,omitempty"`
}

func (g *Group) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "v:group"
	start.Attr = []xml.Attr{}
	// 添加属性
	if g.AnchorId != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "w14:anchorId"}, Value: g.AnchorId})
	}
	if g.ID != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "id"}, Value: g.ID})
	}
	if g.Spid != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "o:spid"}, Value: g.Spid})
	}
	if g.Style != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "style"}, Value: g.Style})
	}
	if g.CoordOrigin != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "coordorigin"}, Value: g.CoordOrigin})
	}
	if g.CoordSize != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "coordsize"}, Value: g.CoordSize})
	}
	if g.GfxData != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "o:gfxdata"}, Value: g.GfxData})
	}

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// 序列化子元素
	if g.Rect != nil {
		if err := g.Rect.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	if g.ShapeType != nil {
		if err := g.ShapeType.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	if g.Shape != nil {
		if err := g.Shape.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

func (g *Group) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// 1. 读取属性
	for _, attr := range start.Attr {
		prefix := attr.Name.Space // Go xml 会自动处理前缀并放入 Space
		local := attr.Name.Local

		// VML 命名空间 (v:) 是默认的，所以 prefix 为空
		// Office 命名空间 (o:)
		// Word 2010 命名空间 (w14:)

		switch {
		case prefix == "w14" && local == "anchorId":
			g.AnchorId = attr.Value
		case prefix == "" && local == "id":
			g.ID = attr.Value
		case prefix == constants.XMLNS_O && local == "spid":
			g.Spid = attr.Value
		case prefix == "" && local == "style":
			g.Style = attr.Value
		case prefix == "" && local == "coordorigin":
			g.CoordOrigin = attr.Value
		case prefix == "" && local == "coordsize":
			g.CoordSize = attr.Value
		case prefix == constants.XMLNS_O && local == "gfxdata":
			g.GfxData = attr.Value
		}
	}

	// 2. 读取子元素
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
			// VML 命名空间: "urn:schemas-microsoft-com:vml"
			if elem.Name.Local == "rect" && elem.Name.Space == constants.XMLNS_V {
				g.Rect = new(Rect)
				if err := g.Rect.UnmarshalXML(d, elem); err != nil {
					return err
				}
			} else if elem.Name.Local == "shape" && elem.Name.Space == constants.XMLNS_V {
				g.Shape = new(VShape)
				if err := g.Shape.UnmarshalXML(d, elem); err != nil {
					return err
				}
			} else if elem.Name.Local == "shapetype" && elem.Name.Space == constants.XMLNS_V {
				g.ShapeType = new(ShapeType)
				if err := g.ShapeType.UnmarshalXML(d, elem); err != nil {
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
			} // 到达 </v:group>
		}
	}
	return nil
}

type Rect struct {
	ID        string `xml:"id,attr,omitempty"`
	Spid      string `xml:"spid,attr,omitempty"` // o:spid
	Style     string `xml:"style,attr,omitempty"`
	FillColor string `xml:"fillcolor,attr,omitempty"`
	Stroked   string `xml:"stroked,attr,omitempty"`
	GfxData   string `xml:"gfxdata,attr,omitempty"`
	Fill      *Fill  `xml:"fill,omitempty"` // <v:fill>
}

func (r *Rect) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "v:rect"
	start.Attr = []xml.Attr{}
	if r.ID != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "id"}, Value: r.ID})
	}
	if r.Spid != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "o:spid"}, Value: r.Spid})
	}
	if r.Style != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "style"}, Value: r.Style})
	}
	if r.FillColor != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "fillcolor"}, Value: r.FillColor})
	}
	if r.Stroked != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "stroked"}, Value: r.Stroked})
	}
	if r.GfxData != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "o:gfxdata"}, Value: r.GfxData})
	}

	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if r.Fill != nil {
		if err := r.Fill.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}
func (r *Rect) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		prefix := attr.Name.Space
		local := attr.Name.Local
		switch {
		case prefix == "" && local == "id":
			r.ID = attr.Value
		case prefix == constants.XMLNS_O && local == "spid":
			r.Spid = attr.Value
		case prefix == constants.XMLNS_O && local == "gfxdata":
			r.GfxData = attr.Value
		case prefix == "" && local == "style":
			r.Style = attr.Value
		case prefix == "" && local == "fillcolor":
			r.FillColor = attr.Value
		case prefix == "" && local == "stroked":
			r.Stroked = attr.Value
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
			if elem.Name.Local == "fill" && elem.Name.Space == constants.XMLNS_V {
				r.Fill = new(Fill)
				if err := r.Fill.UnmarshalXML(d, elem); err != nil {
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

type Fill struct {
	Color2 string `xml:"color2,attr,omitempty"`
	Rotate string `xml:"rotate,attr,omitempty"`
	Angle  string `xml:"angle,attr,omitempty"`
	Focus  string `xml:"focus,attr,omitempty"`
	Type   string `xml:"type,attr,omitempty"`
}

func (f *Fill) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "v:fill"
	start.Attr = []xml.Attr{}
	if f.Color2 != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "color2"}, Value: f.Color2})
	}
	if f.Rotate != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "rotate"}, Value: f.Rotate})
	}
	if f.Angle != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "angle"}, Value: f.Angle})
	}
	if f.Focus != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "focus"}, Value: f.Focus})
	}
	if f.Type != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "type"}, Value: f.Type})
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}
func (f *Fill) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "color2":
			f.Color2 = attr.Value
		case "rotate":
			f.Rotate = attr.Value
		case "angle":
			f.Angle = attr.Value
		case "focus":
			f.Focus = attr.Value
		case "type":
			f.Type = attr.Value
		}
	}
	return d.Skip() // 空元素
}

// ==========================================================
// 新增 VShape 结构体 (对应 <v:shape>)
// ==========================================================
type VShape struct {
	ID        string      `xml:"id,attr,omitempty"`
	Spid      string      `xml:"spid,attr,omitempty"` // o:spid
	Type      string      `xml:"type,attr,omitempty"`
	Alt       string      `xml:"alt,attr,omitempty"`
	Style     string      `xml:"style,attr,omitempty"`
	Gfxdata   string      `xml:"gfxdata,attr,omitempty"`
	ImageData *VImageData `xml:"imageData,omitempty"` // <v:imagedata>
}

func (s *VShape) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "v:shape"
	start.Attr = []xml.Attr{}
	if s.ID != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "id"}, Value: s.ID})
	}
	if s.Spid != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "o:spid"}, Value: s.Spid})
	}
	if s.Type != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "type"}, Value: s.Type})
	}
	if s.Alt != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "alt"}, Value: s.Alt})
	}
	if s.Style != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "style"}, Value: s.Style})
	}
	if s.Gfxdata != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "o:gfxdata"}, Value: s.Gfxdata})
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if s.ImageData != nil {
		if err := s.ImageData.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

func (s *VShape) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		prefix := attr.Name.Space
		local := attr.Name.Local
		switch {
		case prefix == "" && local == "id":
			s.ID = attr.Value
		case prefix == constants.XMLNS_O && local == "spid":
			s.Spid = attr.Value
		case prefix == "" && local == "type":
			s.Type = attr.Value
		case prefix == "" && local == "alt":
			s.Alt = attr.Value
		case prefix == "" && local == "style":
			s.Style = attr.Value
		case prefix == constants.XMLNS_O && local == "gfxdata":
			s.Gfxdata = attr.Value
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
			if elem.Name.Local == "imagedata" && elem.Name.Space == constants.XMLNS_V {
				s.ImageData = new(VImageData)
				if err := s.ImageData.UnmarshalXML(d, elem); err != nil {
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
// 新增 VImageData 结构体 (对应 <v:imagedata>)
// ==========================================================
type VImageData struct {
	RId   string `xml:"id,attr,omitempty"`    // r:id
	Title string `xml:"title,attr,omitempty"` // o:title
}

func (i *VImageData) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "v:imagedata"
	start.Attr = []xml.Attr{}
	if i.RId != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "r:id"}, Value: i.RId})
	}
	if i.Title != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "o:title"}, Value: i.Title})
	}

	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

func (i *VImageData) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		//prefix := attr.Name.Space
		local := attr.Name.Local
		switch {
		case local == "id":
			i.RId = attr.Value
		case local == "title":
			i.Title = attr.Value
		}
	}
	return d.Skip() // 空元素
}

// MarshalXML implements the xml.Marshaler interface.
func (b Shape) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "v:shape"
	start.Attr = []xml.Attr{
		{Name: xml.Name{Local: "type"}, Value: b.Type},
		{Name: xml.Name{Local: "style"}, Value: b.Style},
	}

	err := e.EncodeToken(start)
	if err != nil {
		return err
	}

	if b.ImageData != nil {
		if err := b.ImageData.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

type ImageData struct {
	RId   string `xml:"id,attr,omitempty"`
	Title string `xml:"title,attr,omitempty"`
}

// MarshalXML implements the xml.Marshaler interface.
func (b ImageData) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "v:imagedata"
	start.Attr = []xml.Attr{
		{Name: xml.Name{Local: "r:id"}, Value: b.RId},
		{Name: xml.Name{Local: "o:title"}, Value: b.Title},
	}

	err := e.EncodeToken(start)
	if err != nil {
		return err
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}
