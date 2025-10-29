package ctypes

import (
	"encoding/xml"
	"github.com/iEvan-lhr/docx-agent/internal"
)

type Hyperlink struct {
	XMLName xml.Name `xml:"hyperlink,omitempty"`
	ID      string   `xml:"id,attr"`
	Tooltip *string  `xml:"tooltip,attr"`
	History *string  `xml:"history,attr"`
	Run     *Run
	//Children []ParagraphChild
}

func (h Hyperlink) MarshalXML(e *xml.Encoder, start xml.StartElement) (err error) {
	start.Name.Local = "w:hyperlink"

	if &h.ID != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "w:id"}, Value: h.ID})
	}
	if h.Tooltip != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "w:tooltip"}, Value: *h.Tooltip})
	}
	if h.History != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "w:history"}, Value: *h.History})
	}

	if err = e.EncodeToken(start); err != nil {
		return err
	}

	if h.Run != nil {
		if err = h.Run.MarshalXML(e, xml.StartElement{
			Name: xml.Name{Local: "w:r"},
		}); err != nil {
			return err
		}
	}

	// Closing </w:p> element
	return e.EncodeToken(start.End())
}

func (p *Hyperlink) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	// Decode attributes
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "id":
			p.ID = attr.Value
		case "tooltip":
			p.Tooltip = internal.ToPtr(attr.Value)
		case "history":
			p.History = internal.ToPtr(attr.Value)

		}
	}

loop:
	for {
		currentToken, err := d.Token()
		if err != nil {
			return err
		}

		switch elem := currentToken.(type) {
		case xml.StartElement:
			switch elem.Name.Local {
			case "r":
				p.Run = &Run{}
				if err := d.DecodeElement(p.Run, &elem); err != nil {
					return err
				}
			default:
				if err = d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			break loop
		}
	}

	return nil
}
