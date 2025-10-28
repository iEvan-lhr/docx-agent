package geom

import "encoding/xml"

type ShapeGuide struct {
	Name    string `xml:"name,attr,omitempty"`
	Formula string `xml:"fmla,attr,omitempty"`
}

func (s ShapeGuide) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "a:gd"
	start.Attr = []xml.Attr{
		{Name: xml.Name{Local: "name"}, Value: s.Name},
		{Name: xml.Name{Local: "fmla"}, Value: s.Formula},
	}

	err := e.EncodeToken(start)
	if err != nil {
		return err
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

// UnmarshalXML for ShapeGuide
func (g *ShapeGuide) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "name" {
			g.Name = attr.Value
		}
		if attr.Name.Local == "fmla" {
			g.Formula = attr.Value
		}
	}
	// <a:gd> 是空元素，直接跳过
	return d.Skip()
}
