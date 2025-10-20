package dml

import (
	"encoding/xml"
	"errors"
	"github.com/iEvan-lhr/docx-agent/common/constants"
	"io"
	"strconv"

	"github.com/iEvan-lhr/docx-agent/dml/dmlst"
)

type PoistionH struct {
	RelativeFrom dmlst.RelFromH `xml:"relativeFrom,attr"`
	PosOffset    int            `xml:"posOffset"`
}

type PoistionV struct {
	RelativeFrom dmlst.RelFromV `xml:"relativeFrom,attr"`
	PosOffset    int            `xml:"posOffset"`
}

func getInt(val string) (int, error) {
	if val == "" {
		return 0, nil
	}
	i, err := strconv.ParseInt(val, 10, 64)
	return int(i), err
}

func (p *PoistionH) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "relativeFrom" {
			p.RelativeFrom = dmlst.RelFromH(attr.Value)
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
			case xml.Name{Space: constants.WMLDrawingNS, Local: "posOffset"}:
				if err = d.DecodeElement(&p.PosOffset, &elem); err != nil {
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

func (p *PoistionV) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "relativeFrom" {
			p.RelativeFrom = dmlst.RelFromV(attr.Value)
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
			case xml.Name{Space: constants.WMLDrawingNS, Local: "posOffset"}:
				if err = d.DecodeElement(&p.PosOffset, &elem); err != nil {
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
func (p PoistionH) MarshalXML(e *xml.Encoder, start xml.StartElement) error {

	if p.RelativeFrom == "" {
		return errors.New("Invalid RelativeFrom in PoistionH")
	}

	start.Name.Local = "wp:positionH"

	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "relativeFrom"}, Value: string(p.RelativeFrom)})

	err := e.EncodeToken(start)
	if err != nil {
		return err
	}

	offsetElem := xml.StartElement{Name: xml.Name{Local: "wp:posOffset"}}
	if err = e.EncodeElement(p.PosOffset, offsetElem); err != nil {
		return err
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

func (p PoistionV) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if p.RelativeFrom == "" {
		return errors.New("Invalid RelativeFrom in PoistionV")
	}

	start.Name.Local = "wp:positionV"

	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "relativeFrom"}, Value: string(p.RelativeFrom)})

	err := e.EncodeToken(start)
	if err != nil {
		return err
	}

	offsetElem := xml.StartElement{Name: xml.Name{Local: "wp:posOffset"}}
	if err = e.EncodeElement(p.PosOffset, offsetElem); err != nil {
		return err
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}
