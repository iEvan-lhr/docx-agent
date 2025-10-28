package dml

import (
	"encoding/xml"
	"io"

	"github.com/iEvan-lhr/docx-agent/common/constants"
	"github.com/iEvan-lhr/docx-agent/dml/dmlpic"
)

type Graphic struct {
	Data *GraphicData `xml:"graphicData,omitempty"`
}

func NewGraphic(data *GraphicData) *Graphic {
	return &Graphic{Data: data}
}

func DefaultGraphic() *Graphic {
	return &Graphic{}
}

type GraphicData struct {
	URI      string      `xml:"uri,attr,omitempty"`
	WPGGroup *WPGGroup   `xml:"wgp,omitempty"`
	Pic      *dmlpic.Pic `xml:"pic,omitempty"`
}

func NewPicGraphic(pic *dmlpic.Pic) *Graphic {
	return &Graphic{
		Data: &GraphicData{
			URI: constants.DrawingMLPicNS,
			Pic: pic,
		},
	}
}

func (g Graphic) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "a:graphic"
	start.Attr = []xml.Attr{
		{Name: xml.Name{Local: "xmlns:a"}, Value: constants.DrawingMLMainNS},
	}

	err := e.EncodeToken(start)
	if err != nil {
		return err
	}

	if g.Data != nil {
		if err = g.Data.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

func (gd GraphicData) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "a:graphicData"
	start.Attr = []xml.Attr{
		{Name: xml.Name{Local: "uri"}, Value: constants.DrawingMLPicNS},
	}

	err := e.EncodeToken(start)
	if err != nil {
		return err
	}

	if gd.Pic != nil {
		if err := e.EncodeElement(gd.Pic, xml.StartElement{Name: xml.Name{Local: "pic:pic"}}); err != nil {
			return err
		}
	}
	if gd.WPGGroup != nil {
		if err := e.EncodeElement(gd.WPGGroup, xml.StartElement{Name: xml.Name{Local: "wpg:wgp"}}); err != nil {
			return err
		}
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

func (g *Graphic) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			case xml.Name{Space: constants.DrawingMLMainNS, Local: "graphicData"}:
				g.Data = new(GraphicData)
				if err = d.DecodeElement(g.Data, &elem); err != nil {
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

func (gd *GraphicData) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "uri" {
			gd.URI = attr.Value
		}
	}

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
			case xml.Name{Space: constants.DrawingMLPicNS, Local: "pic"}:
				gd.Pic = new(dmlpic.Pic)
				if err = d.DecodeElement(gd.Pic, &elem); err != nil {
					return err
				}
			case xml.Name{Space: constants.WPGNamespace, Local: "wgp"}:
				gd.WPGGroup = new(WPGGroup)
				if err = d.DecodeElement(gd.WPGGroup, &elem); err != nil {
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
