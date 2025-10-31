package docx

import (
	"encoding/xml"
	"github.com/iEvan-lhr/docx-agent/wml/ctypes"
)

// Header represents a document header
type Header struct {
	root *RootDoc
	// XMLName 字段在 Unmarshal 时用于匹配
	//XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main hdr"`
	Children     []DocumentChild
	SectPr       *ctypes.SectionProp
	RelativePath string     // 文件在 zip 包中的路径，如 word/header1.xml
	ID           string     // relationship ID
	Attrs        []xml.Attr // <--- 用于存储根元素的属性
}

// NewHeader creates a new Header instance
func NewHeader(root *RootDoc) *Header {
	return &Header{
		root: root,
	}
}

// MarshalXML implements the xml.Marshaler interface for the Body type.
// It encodes the Body to its corresponding XML representation.
func (h Header) MarshalXML(e *xml.Encoder, start xml.StartElement) (err error) {
	start.Name.Local = "w:hdr"
	start.Attr = h.Attrs

	err = e.EncodeToken(start)
	if err != nil {
		return err
	}

	if h.Children != nil {
		for _, child := range h.Children {
			if child.Para != nil {
				if err = child.Para.ct.MarshalXML(e, xml.StartElement{}); err != nil {
					return err
				}
			}

			if child.Table != nil {
				if err = child.Table.ct.MarshalXML(e, xml.StartElement{}); err != nil {
					return err
				}
			}
		}
	}

	if h.SectPr != nil {
		if err = h.SectPr.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

// UnmarshalXML implements the xml.Unmarshaler interface for the Body type.
// It decodes the XML representation of the Body.
func (h *Header) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	h.Attrs = make([]xml.Attr, len(start.Attr))
	for i := range h.Attrs {
		if start.Attr[i].Name.Local != "Ignorable" {
			h.Attrs[i].Name.Local = start.Attr[i].Name.Space + ":" + start.Attr[i].Name.Local
			h.Attrs[i].Value = start.Attr[i].Value
		} else {
			h.Attrs[i].Name.Local = "mc:" + start.Attr[i].Name.Local
			h.Attrs[i].Value = start.Attr[i].Value
		}
	}

	for {
		currentToken, err := d.Token()
		if err != nil {
			return err
		}

		switch elem := currentToken.(type) {
		case xml.StartElement:
			switch elem.Name.Local {
			case "p":
				para := newParagraph(h.root)
				if err := para.unmarshalXML(d, elem); err != nil {
					return err
				}
				h.Children = append(h.Children, DocumentChild{Para: para})
			case "tbl":
				tbl := NewTable(h.root)
				if err := tbl.unmarshalXML(d, elem); err != nil {
					return err
				}
				h.Children = append(h.Children, DocumentChild{Table: tbl})
			case "sectPr":
				h.SectPr = ctypes.NewSectionProper()
				if err := d.DecodeElement(h.SectPr, &elem); err != nil {
					return err
				}
			default:
				if err = d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			return nil
		}
	}
}

// LoadHeaderXml loads header from XML bytes
func LoadHeaderXml(rd *RootDoc, fileName string, fileBytes []byte) (*Header, error) {
	header := NewHeader(rd)
	err := xml.Unmarshal(fileBytes, header)
	if err != nil {
		return nil, err
	}
	header.RelativePath = fileName
	return header, nil
}
