package dml

import (
	"encoding/xml"
	"fmt"
	"io"
	"strconv"

	"github.com/iEvan-lhr/docx-agent/common/constants"
	"github.com/iEvan-lhr/docx-agent/dml/dmlct"
	"github.com/iEvan-lhr/docx-agent/dml/dmlst"
)

// This element specifies that the DrawingML object located at this position in the document is an inline object. Within a WordprocessingML document, drawing objects can exist in two states:
//
//• Inline - The drawing object is in line with the text, and affects the line height and layout of its line (like a character glyph of similar size).

type Inline struct {
	/// Specifies the minimum distance which shall be maintained between the top edge of this drawing object and any subsequent text within the document when this graphical object is displayed within the document's contents.,
	/// The distance shall be measured in EMUs (English Mektric Units).,
	//
	// NOTE!: As per http://www.datypic.com/sc/ooxml/e-wp_inline.html, Dist* attributes is optional
	// But MS Word requires them to be there

	//Distance From Text on Top Edge
	DistT uint `xml:"distT,attr,omitempty"`

	//Distance From Text on Bottom Edge
	DistB uint `xml:"distB,attr,omitempty"`

	//Distance From Text on Left Edge
	DistL uint `xml:"distL,attr,omitempty"`

	//Distance From Text on Right Edge
	DistR uint `xml:"distR,attr,omitempty"`

	// Child elements:
	// 1. Drawing Object Size
	Extent dmlct.PSize2D `xml:"extent,omitempty"`

	// 2. Inline Wrapping Extent
	EffectExtent *EffectExtent `xml:"effectExtent,omitempty"`

	// 3. Drawing Object Non-Visual Properties
	DocProp DocProp `xml:"docPr,omitempty"`

	//4.Common DrawingML Non-Visual Properties
	CNvGraphicFramePr *NonVisualGraphicFrameProp `xml:"cNvGraphicFramePr,omitempty"`

	//5.Graphic Object
	Graphic Graphic `xml:"graphic,omitempty"`
}

func NewInline(extent dmlct.PSize2D, docProp DocProp, graphic Graphic) Inline {
	return Inline{
		Extent:  extent,
		DocProp: docProp,
		Graphic: graphic,
		CNvGraphicFramePr: &NonVisualGraphicFrameProp{
			GraphicFrameLocks: &GraphicFrameLocks{
				NoChangeAspect: dmlst.NewOptBool(true),
			},
		},
	}
}

func (i Inline) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "wp:inline"
	start.Attr = []xml.Attr{
		{Name: xml.Name{Local: "xmlns:a"}, Value: constants.DrawingMLMainNS},
		{Name: xml.Name{Local: "xmlns:pic"}, Value: constants.DrawingMLPicNS},
	}

	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "distT"}, Value: strconv.FormatUint(uint64(i.DistT), 10)})
	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "distB"}, Value: strconv.FormatUint(uint64(i.DistB), 10)})
	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "distL"}, Value: strconv.FormatUint(uint64(i.DistL), 10)})
	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "distR"}, Value: strconv.FormatUint(uint64(i.DistR), 10)})

	err := e.EncodeToken(start)
	if err != nil {
		return err
	}

	// 1.Extent
	if err := i.Extent.MarshalXML(e, xml.StartElement{Name: xml.Name{Local: "wp:extent"}}); err != nil {
		return fmt.Errorf("marshalling Extent: %w", err)
	}

	// 2. EffectExtent
	if i.EffectExtent != nil {
		if err := i.EffectExtent.MarshalXML(e, xml.StartElement{Name: xml.Name{Local: "wp:effectExtent"}}); err != nil {
			return fmt.Errorf("EffectExtent: %v", err)
		}
	}

	// 3. docPr
	if err = i.DocProp.MarshalXML(e, xml.StartElement{}); err != nil {
		return fmt.Errorf("marshalling DocProp: %w", err)
	}

	// 4. cNvGraphicFramePr
	if i.CNvGraphicFramePr != nil {
		if err = i.CNvGraphicFramePr.MarshalXML(e, xml.StartElement{}); err != nil {
			return fmt.Errorf("marshalling cNvGraphicFramePr: %w", err)
		}
	}

	// 5. graphic
	if err = i.Graphic.MarshalXML(e, xml.StartElement{}); err != nil {
		return fmt.Errorf("marshalling Graphic: %w", err)
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

func getUint64(val string) (uint64, error) {
	if val == "" {
		return 0, nil
	}
	return strconv.ParseUint(val, 10, 64)
}

func getUint(val string) (uint, error) {
	u, err := getUint64(val)
	return uint(u), err
}

func (i *Inline) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "distT":
			i.DistT, err = getUint(attr.Value)
		case "distB":
			i.DistB, err = getUint(attr.Value)
		case "distL":
			i.DistL, err = getUint(attr.Value)
		case "distR":
			i.DistR, err = getUint(attr.Value)
		}
		if err != nil {
			return err
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
			case xml.Name{Space: constants.WMLDrawingNS, Local: "extent"}:
				if err = d.DecodeElement(&i.Extent, &elem); err != nil {
					return err
				}
			case xml.Name{Space: constants.WMLDrawingNS, Local: "effectExtent"}:
				i.EffectExtent = new(EffectExtent)
				if err = d.DecodeElement(i.EffectExtent, &elem); err != nil {
					return err
				}
			case xml.Name{Space: constants.WMLDrawingNS, Local: "docPr"}:
				if err = d.DecodeElement(&i.DocProp, &elem); err != nil {
					return err
				}
			case xml.Name{Space: constants.WMLDrawingNS, Local: "cNvGraphicFramePr"}:
				i.CNvGraphicFramePr = new(NonVisualGraphicFrameProp)
				if err = d.DecodeElement(i.CNvGraphicFramePr, &elem); err != nil {
					return err
				}
			case xml.Name{Space: constants.DrawingMLMainNS, Local: "graphic"}:
				if err = d.DecodeElement(&i.Graphic, &elem); err != nil {
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
