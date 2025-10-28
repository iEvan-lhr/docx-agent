// in: wml/ctypes/mc.go (新文件)
package ctypes

import (
	"encoding/xml"
	"github.com/iEvan-lhr/docx-agent/dml" // 引用 dml 包
)

// AlternateContent 支持 mc:AlternateContent 的解析（仅解析，按需取其内的 drawing/pict）
type AlternateContent struct {
	Choice   *ACChoice `xml:"Choice"`   // mc:Choice
	Fallback *Fallback `xml:"Fallback"` // mc:Fallback
}

type ACChoice struct {
	Requires string       `xml:"Requires,attr,omitempty"`
	Drawing  *dml.Drawing `xml:"drawing,omitempty"`
	Pict     *Pict        `xml:"pict,omitempty"`
}

type Fallback struct {
	Drawing *dml.Drawing `xml:"drawing,omitempty"`
	Pict    *Pict        `xml:"pict,omitempty"`
}

func (ac AlternateContent) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "mc:AlternateContent"
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if ac.Choice != nil && (ac.Choice.Drawing != nil || ac.Choice.Pict != nil) {
		ch := xml.StartElement{Name: xml.Name{Local: "mc:Choice"}}
		if ac.Choice.Requires != "" {
			ch.Attr = append(ch.Attr, xml.Attr{Name: xml.Name{Local: "Requires"}, Value: ac.Choice.Requires})
		}
		if err := e.EncodeToken(ch); err != nil {
			return err
		}
		if ac.Choice.Drawing != nil {
			if err := ac.Choice.Drawing.MarshalXML(e, xml.StartElement{Name: xml.Name{Local: "w:drawing"}}); err != nil {
				return err
			}
		}
		if ac.Choice.Pict != nil {
			if err := ac.Choice.Pict.MarshalXML(e, xml.StartElement{Name: xml.Name{Local: "w:pict"}}); err != nil {
				return err
			}
		}
		if err := e.EncodeToken(xml.EndElement{Name: ch.Name}); err != nil {
			return err
		}
	}
	if ac.Fallback != nil && (ac.Fallback.Drawing != nil || ac.Fallback.Pict != nil) {
		fb := xml.StartElement{Name: xml.Name{Local: "mc:Fallback"}}
		if err := e.EncodeToken(fb); err != nil {
			return err
		}
		if ac.Fallback.Drawing != nil {
			if err := ac.Fallback.Drawing.MarshalXML(e, xml.StartElement{Name: xml.Name{Local: "w:drawing"}}); err != nil {
				return err
			}
		}
		if ac.Fallback.Pict != nil {
			if err := ac.Fallback.Pict.MarshalXML(e, xml.StartElement{Name: xml.Name{Local: "w:pict"}}); err != nil {
				return err
			}
		}
		if err := e.EncodeToken(xml.EndElement{Name: fb.Name}); err != nil {
			return err
		}
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}
