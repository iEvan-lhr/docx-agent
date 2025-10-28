package ctypes

import (
	"encoding/xml"
	"github.com/iEvan-lhr/docx-agent/common/constants"
	"io"
)

type ShapeType struct {
	ID             string    `xml:"id,attr,omitempty"`
	Coordsize      string    `xml:"coordsize,attr,omitempty"` // o:spid
	Spt            string    `xml:"spt,attr,omitempty"`
	Preferrelative string    `xml:"preferrelative,attr,omitempty"`
	Path           string    `xml:"path,attr,omitempty"`
	Filled         string    `xml:"filled,attr,omitempty"`
	Stroked        string    `xml:"stroked,attr,omitempty"`
	Stroke         *Stroke   `xml:"stroke,omitempty"`
	VPath          *VPath    `xml:"path,omitempty"`
	Lock           *Lock     `xml:"lock,omitempty"`
	Formulas       *Formulas `xml:"formulas,omitempty"`
}

func (st *ShapeType) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "v:shapetype"
	start.Attr = []xml.Attr{}
	// 添加属性
	if st.Coordsize != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "coordsize"}, Value: st.Coordsize})
	}
	if st.ID != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "id"}, Value: st.ID})
	}
	if st.Preferrelative != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "o:preferrelative"}, Value: st.Preferrelative})
	}
	if st.Spt != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "o:spt"}, Value: st.Spt})
	}
	if st.Path != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "path"}, Value: st.Path})
	}
	if st.Filled != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "filled"}, Value: st.Filled})
	}
	if st.Stroked != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "stroked"}, Value: st.Stroked})
	}

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// 序列化子元素
	if st.Stroke != nil {
		if err := st.Stroke.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	if st.Formulas != nil {
		if err := st.Formulas.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	if st.VPath != nil {
		if err := st.VPath.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	if st.Lock != nil {
		if err := st.Lock.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

func (st *ShapeType) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// 1. 读取属性
	for _, attr := range start.Attr {
		prefix := attr.Name.Space // Go xml 会自动处理前缀并放入 Space
		local := attr.Name.Local

		// VML 命名空间 (v:) 是默认的，所以 prefix 为空
		// Office 命名空间 (o:)
		// Word 2010 命名空间 (w14:)

		switch {
		case prefix == "" && local == "id":
			st.ID = attr.Value
		case prefix == constants.XMLNS_O && local == "spt":
			st.Spt = attr.Value
		case prefix == "" && local == "coordsize":
			st.Coordsize = attr.Value
		case prefix == "" && local == "path":
			st.Path = attr.Value
		case prefix == "" && local == "filled":
			st.Filled = attr.Value
		case prefix == "" && local == "stroked":
			st.Stroked = attr.Value
		case prefix == constants.XMLNS_O && local == "preferrelative":
			st.Preferrelative = attr.Value
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
			if elem.Name.Local == "path" && elem.Name.Space == constants.XMLNS_V {
				st.VPath = new(VPath)
				if err := st.VPath.UnmarshalXML(d, elem); err != nil {
					return err
				}
			} else if elem.Name.Local == "stroke" && elem.Name.Space == constants.XMLNS_V {
				st.Stroke = new(Stroke)
				if err := st.Stroke.UnmarshalXML(d, elem); err != nil {
					return err
				}
			} else if elem.Name.Local == "formulas" && elem.Name.Space == constants.XMLNS_V {
				st.Formulas = new(Formulas)
				if err := st.Formulas.UnmarshalXML(d, elem); err != nil {
					return err
				}
			} else if elem.Name.Local == "lock" && elem.Name.Space == constants.XMLNS_O {
				st.Lock = new(Lock)
				if err := st.Lock.UnmarshalXML(d, elem); err != nil {
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

type Stroke struct {
	JoinStyle string `xml:"joinstyle,attr,omitempty"`
}

func (s *Stroke) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "v:stroke"
	start.Attr = []xml.Attr{}

	if s.JoinStyle != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "joinstyle"}, Value: s.JoinStyle})
	}

	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

func (s *Stroke) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "joinstyle" {
			s.JoinStyle = attr.Value
		}
	}
	return d.Skip() // 空元素
}

type Formulas struct {
	F []*F `xml:"f,omitempty"`
}

func (f *Formulas) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "v:formulas"
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	for _, fs := range f.F {
		if err := fs.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}
func (f *Formulas) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	f.F = []*F{}
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
			if elem.Name.Local == "f" && elem.Name.Space == constants.XMLNS_V {
				gs := new(F)
				if err := gs.UnmarshalXML(d, elem); err != nil {
					return err
				}
				f.F = append(f.F, gs)
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

type F struct {
	Eqn string `xml:"eqn,attr,omitempty"`
}

func (f0 *F) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "v:f"
	start.Attr = []xml.Attr{}

	if f0.Eqn != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "eqn"}, Value: f0.Eqn})
	}

	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

func (f0 *F) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "eqn" {
			f0.Eqn = attr.Value
		}
	}
	return d.Skip() // 空元素
}

type VPath struct {
	Extrusionok     string `xml:"extrusionok,attr,omitempty"`
	Qradientshapeok string `xml:"gradientshapeok,attr,omitempty"`
	Connecttype     string `xml:"connecttype,attr,omitempty"`
}

func (vp *VPath) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "v:path"
	start.Attr = []xml.Attr{}

	if vp.Connecttype != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "o:connecttype"}, Value: vp.Connecttype})
	}
	if vp.Qradientshapeok != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "gradientshapeok"}, Value: vp.Qradientshapeok})
	}
	if vp.Extrusionok != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "o:extrusionok"}, Value: vp.Extrusionok})
	}

	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

func (vp *VPath) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "connecttype" {
			vp.Connecttype = attr.Value
		}
		if attr.Name.Local == "gradientshapeok" {
			vp.Qradientshapeok = attr.Value
		}
		if attr.Name.Local == "extrusionok" {
			vp.Extrusionok = attr.Value
		}
	}
	return d.Skip() // 空元素
}

type Lock struct {
	Ext         string `xml:"ext,attr,omitempty"`
	Aspectratio string `xml:"aspectratio,attr,omitempty"`
}

func (lo *Lock) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "o:lock"
	start.Attr = []xml.Attr{}

	if lo.Ext != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "v:ext"}, Value: lo.Ext})
	}
	if lo.Aspectratio != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "aspectratio"}, Value: lo.Aspectratio})
	}

	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

func (lo *Lock) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "ext" {
			lo.Ext = attr.Value
		}
		if attr.Name.Local == "aspectratio" {
			lo.Aspectratio = attr.Value
		}

	}
	return d.Skip() // 空元素
}
