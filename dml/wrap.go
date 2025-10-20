package dml

import (
	"encoding/xml"
	"github.com/iEvan-lhr/docx-agent/common/constants"
	"io"
	"strconv"

	"github.com/iEvan-lhr/docx-agent/dml/dmlct"
	"github.com/iEvan-lhr/docx-agent/dml/dmlst"
)

type WrapNone struct {
	XMLName xml.Name `xml:"wrapNone"`
}

func (w WrapNone) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "wp:wrapNone"
	return e.EncodeElement("", start)
}

// // Dummy implementation to ensure only these types are allowed in the wrap type
// func (w WrapNone) getWrapName()    {}
// func (w WrapSquare) getWrapName()  {}
// func (w WrapThrough) getWrapName() {}
// func (w WrapTopBtm) getWrapName()  {}

type WrapSquare struct {
	XMLName xml.Name `xml:"wrapSquare"`

	//Text Wrapping Location
	WrapText dmlst.WrapText `xml:"wrapText,attr"`

	//Distance From Text (Top)
	DistT *uint `xml:"distT,attr,omitempty"`

	//Distance From Text on Bottom Edge
	DistB *uint `xml:"distB,attr,omitempty"`

	//Distance From Text on Left Edge
	DistL *uint `xml:"distL,attr,omitempty"`

	//Distance From Text on Right Edge
	DistR *uint `xml:"distR,attr,omitempty"`

	EffectExtent *EffectExtent `xml:"effectExtent,omitempty"`
}

func (ws WrapSquare) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "wp:wrapSquare"

	start.Attr = []xml.Attr{
		{Name: xml.Name{Local: "wrapText"}, Value: string(ws.WrapText)},
	}

	if ws.DistT != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "distT"}, Value: strconv.FormatUint(uint64(*ws.DistT), 10)})
	}
	if ws.DistB != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "distB"}, Value: strconv.FormatUint(uint64(*ws.DistB), 10)})
	}
	if ws.DistL != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "distL"}, Value: strconv.FormatUint(uint64(*ws.DistL), 10)})
	}
	if ws.DistR != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "distR"}, Value: strconv.FormatUint(uint64(*ws.DistR), 10)})
	}

	err := e.EncodeToken(start)
	if err != nil {
		return err
	}

	if ws.EffectExtent != nil {
		if err := ws.EffectExtent.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

type WrapPolygon struct {
	Start  dmlct.Point2D   `xml:"start"`
	LineTo []dmlct.Point2D `xml:"lineTo"`
	Edited *bool           `xml:"edited,attr,omitempty"`
}

func (wp WrapPolygon) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "wp:wrapPolygon"

	start.Attr = []xml.Attr{}

	if wp.Edited != nil {
		if *wp.Edited {
			start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "edited"}, Value: "true"})
		} else {
			start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "edited"}, Value: "false"})
		}
	}

	err := e.EncodeToken(start)
	if err != nil {
		return err
	}

	if err := wp.Start.MarshalXML(e, xml.StartElement{
		Name: xml.Name{Local: "wp:start"},
	}); err != nil {
		return err
	}

	for _, lineTo := range wp.LineTo {
		if err := lineTo.MarshalXML(e, xml.StartElement{
			Name: xml.Name{Local: "wp:lineTo"},
		}); err != nil {
			return err
		}
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

// Tight Wrapping
type WrapTight struct {
	XMLName xml.Name `xml:"wrapTight"`

	//Tight Wrapping Extents Polygon
	WrapPolygon WrapPolygon `xml:"wrapPolygon"`

	// Text Wrapping Location
	WrapText dmlst.WrapText `xml:"wrapText,attr"`

	// Distance From Text on Left Edge
	DistL *uint `xml:"distL,attr,omitempty"`

	// Distance From Text on Right Edge
	DistR *uint `xml:"distR,attr,omitempty"`
}

func (w WrapTight) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "wp:wrapTight"

	start.Attr = []xml.Attr{
		{Name: xml.Name{Local: "wrapText"}, Value: string(w.WrapText)},
	}

	if w.DistL != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "distL"}, Value: strconv.FormatUint(uint64(*w.DistL), 10)})
	}
	if w.DistR != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "distR"}, Value: strconv.FormatUint(uint64(*w.DistR), 10)})
	}

	err := e.EncodeToken(start)
	if err != nil {
		return err
	}

	if err := w.WrapPolygon.MarshalXML(e, xml.StartElement{}); err != nil {
		return err
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

// Through Wrapping
type WrapThrough struct {
	XMLName xml.Name `xml:"wrapThrough"`

	//Tight Wrapping Extents Polygon
	WrapPolygon WrapPolygon `xml:"wrapPolygon"`

	// Text Wrapping Location
	WrapText dmlst.WrapText `xml:"wrapText,attr"`

	// Distance From Text on Left Edge
	DistL *uint `xml:"distL,attr,omitempty"`

	// Distance From Text on Right Edge
	DistR *uint `xml:"distR,attr,omitempty"`
}

func (w WrapThrough) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "wp:wrapThrough"

	start.Attr = []xml.Attr{
		{Name: xml.Name{Local: "wrapText"}, Value: string(w.WrapText)},
	}

	if w.DistL != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "distL"}, Value: strconv.FormatUint(uint64(*w.DistL), 10)})
	}
	if w.DistR != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "distR"}, Value: strconv.FormatUint(uint64(*w.DistR), 10)})
	}

	err := e.EncodeToken(start)
	if err != nil {
		return err
	}

	if err := w.WrapPolygon.MarshalXML(e, xml.StartElement{}); err != nil {
		return err
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

// Top and Bottom Wrapping
type WrapTopBtm struct {
	XMLName xml.Name `xml:"wrapTopAndBottom"`

	//Distance From Text (Top)
	DistT *uint `xml:"distT,attr,omitempty"`

	//Distance From Text on Bottom Edge
	DistB *uint `xml:"distB,attr,omitempty"`

	//Wrapping Boundaries
	EffectExtent *EffectExtent `xml:"effectExtent,omitempty"`
}

func (w WrapTopBtm) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "wp:wrapTopAndBottom"

	start.Attr = []xml.Attr{}

	if w.DistT != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "distT"}, Value: strconv.FormatUint(uint64(*w.DistT), 10)})
	}
	if w.DistB != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "distB"}, Value: strconv.FormatUint(uint64(*w.DistB), 10)})
	}

	err := e.EncodeToken(start)
	if err != nil {
		return err
	}

	if w.EffectExtent != nil {
		if err := w.EffectExtent.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

func (w *WrapNone) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// 这是一个空元素, 跳过它
	return d.Skip()
}

func (ws *WrapSquare) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "wrapText":
			ws.WrapText = dmlst.WrapText(attr.Value)
		case "distT":
			var u uint
			u, err = getUint(attr.Value)
			ws.DistT = &u
		case "distB":
			var u uint
			u, err = getUint(attr.Value)
			ws.DistB = &u
		case "distL":
			var u uint
			u, err = getUint(attr.Value)
			ws.DistL = &u
		case "distR":
			var u uint
			u, err = getUint(attr.Value)
			ws.DistR = &u
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
			case xml.Name{Space: constants.WMLDrawingNS, Local: "effectExtent"}:
				ws.EffectExtent = new(EffectExtent)
				if err = d.DecodeElement(ws.EffectExtent, &elem); err != nil {
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

func (wp *WrapPolygon) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		if attr.Name.Local == "edited" {
			var b bool
			if attr.Value == "true" || attr.Value == "1" {
				b = true
			}
			wp.Edited = &b
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
			case xml.Name{Space: constants.WMLDrawingNS, Local: "start"}:
				if err = d.DecodeElement(&wp.Start, &elem); err != nil {
					return err
				}
			case xml.Name{Space: constants.WMLDrawingNS, Local: "lineTo"}:
				var pt dmlct.Point2D
				if err = d.DecodeElement(&pt, &elem); err != nil {
					return err
				}
				wp.LineTo = append(wp.LineTo, pt)
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

func (w *WrapTight) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "wrapText":
			w.WrapText = dmlst.WrapText(attr.Value)
		case "distL":
			var u uint
			u, err = getUint(attr.Value)
			w.DistL = &u
		case "distR":
			var u uint
			u, err = getUint(attr.Value)
			w.DistR = &u
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
			case xml.Name{Space: constants.WMLDrawingNS, Local: "wrapPolygon"}:
				if err = d.DecodeElement(&w.WrapPolygon, &elem); err != nil {
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

func (w *WrapThrough) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "wrapText":
			w.WrapText = dmlst.WrapText(attr.Value)
		case "distL":
			var u uint
			u, err = getUint(attr.Value)
			w.DistL = &u
		case "distR":
			var u uint
			u, err = getUint(attr.Value)
			w.DistR = &u
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
			case xml.Name{Space: constants.WMLDrawingNS, Local: "wrapPolygon"}:
				if err = d.DecodeElement(&w.WrapPolygon, &elem); err != nil {
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

func (w *WrapTopBtm) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "distT":
			var u uint
			u, err = getUint(attr.Value)
			w.DistT = &u
		case "distB":
			var u uint
			u, err = getUint(attr.Value)
			w.DistB = &u
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
			case xml.Name{Space: constants.WMLDrawingNS, Local: "effectExtent"}:
				w.EffectExtent = new(EffectExtent)
				if err = d.DecodeElement(w.EffectExtent, &elem); err != nil {
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
