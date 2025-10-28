package ctypes

import (
	"bytes"
	"encoding/xml"
)

type FldChar struct {
	FldCharType string `xml:"fldCharType,attr"`
}

func (f *FldChar) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "w:fldChar"
	start.Attr = []xml.Attr{}

	if f.FldCharType != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "w:fldCharType"}, Value: f.FldCharType})
	}

	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

func (f *FldChar) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "fldCharType" {
			f.FldCharType = attr.Value
		}
	}
	return d.Skip() // 空元素
}

type InStrText struct {
	Text  string
	Space string `xml:"space,attr"`
}

func (i *InStrText) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "w:instrText"
	start.Attr = []xml.Attr{}

	if i.Space != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "xml:space"}, Value: i.Space})
	}

	if err := e.EncodeElement(i.Text, start); err != nil {
		return err
	}

	return nil
}

func (i *InStrText) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	var buf bytes.Buffer

	for _, attr := range start.Attr {
		if attr.Name.Local == "space" {
			i.Space = attr.Value
		}
	}
	for {
		token, err := d.Token()
		if err != nil {
			return err
		}

		switch tokenElem := token.(type) {
		case xml.CharData:
			buf.Write([]byte(tokenElem))
		case xml.EndElement:
			if tokenElem == start.End() {
				i.Text = buf.String()
				return nil
			}
		}
	}
	//return d.Skip() // 空元素
}

type BookmarkStart struct {
	ID   string `xml:"id,attr"`
	Name string `xml:"name,attr"`
}

func (b *BookmarkStart) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "w:bookmarkStart"
	start.Attr = []xml.Attr{}

	if b.ID != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "w:id"}, Value: b.ID})
	}
	if b.Name != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "w:name"}, Value: b.Name})
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

func (b *BookmarkStart) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "id" {
			b.ID = attr.Value
		}
		if attr.Name.Local == "name" {
			b.Name = attr.Value
		}
	}
	return d.Skip() // 空元素
}
