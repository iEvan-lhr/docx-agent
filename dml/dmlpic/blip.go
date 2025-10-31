package dmlpic

import (
	"encoding/xml"
	"github.com/iEvan-lhr/docx-agent/dml/dmlct"
)

// Binary large image or picture
type Blip struct {
	EmbedID string        `xml:"embed,attr,omitempty"`
	ExtLst  *dmlct.ExtLst `xml:"extLst,omitempty"`
}

func (b Blip) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "a:blip"

	start.Attr = []xml.Attr{
		{Name: xml.Name{Local: "r:embed"}, Value: b.EmbedID},
	}

	err := e.EncodeToken(start)
	if err != nil {
		return err
	}
	if b.ExtLst != nil {
		if err := b.ExtLst.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}
