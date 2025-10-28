package dml

import (
	"encoding/xml"
	"fmt"
	"github.com/iEvan-lhr/docx-agent/common/constants"
	"github.com/iEvan-lhr/docx-agent/internal"
	"github.com/iEvan-lhr/docx-agent/wml/stypes"
	"io"
	"strconv"

	"github.com/iEvan-lhr/docx-agent/dml/dmlct"
)

type Anchor struct {
	/// Specifies that this object shall be positioned using the positioning information in the
	/// simplePos child element (§20.4.2.13). This positioning, when specified, positions the
	/// object on the page by placing its top left point at the x-y coordinates specified by that
	/// element.
	/// Reference: http://officeopenxml.com/drwPicFloating-position.php
	//Page Positioning
	SimplePosAttr *int `xml:"simplePos,attr,omitempty"`

	/// Specifies the minimum distance which shall be maintained between the top edge of this drawing object and any subsequent text within the document when this graphical object is displayed within the document's contents.,
	/// The distance shall be measured in EMUs (English Mektric Units).,
	//Distance From Text on Top Edge
	DistT uint `xml:"distT,attr,omitempty"`
	//Distance From Text on Bottom Edge
	DistB uint `xml:"distB,attr,omitempty"`
	//Distance From Text on Left Edge
	DistL uint `xml:"distL,attr,omitempty"`
	//Distance From Text on Right Edge
	DistR uint `xml:"distR,attr,omitempty"`

	//Relative Z-Ordering Position
	RelativeHeight int `xml:"relativeHeight,attr"`

	//Layout In Table Cell
	LayoutInCell int `xml:"layoutInCell,attr"`

	//Display Behind Document Text
	BehindDoc int `xml:"behindDoc,attr"`

	//Lock Anchor
	Locked int `xml:"locked,attr"`

	//Allow Objects to Overlap
	AllowOverlap int `xml:"allowOverlap,attr"`

	Hidden *int `xml:"hidden,attr,omitempty"`

	// wp14:anchorId – Unique identifier for the anchor (Office 2010+)
	AnchorId *stypes.LongHexNum `xml:"wp14:anchorId,attr,omitempty"`

	// wp14:editId – Edit session identifier (Office 2010+)
	EditId *stypes.LongHexNum `xml:"wp14:editId,attr,omitempty"`
	// Child elements:
	// 1. Simple Positioning Coordinates
	SimplePos dmlct.Point2D `xml:"simplePos"`

	// 2. Horizontal Positioning
	PositionH PoistionH `xml:"positionH"`

	// 3. Vertical Positioning
	PositionV PoistionV `xml:"positionV"`

	// 4. Inline Drawing Object Extents
	Extent dmlct.PSize2D `xml:"extent"`

	// 5. EffectExtent
	EffectExtent *EffectExtent `xml:"effectExtent,omitempty"`

	// 6. Wrapping
	WrapTight *WrapTight `xml:"wrapTight,omitempty"`
	// 6.1 .wrapNone
	WrapNone *WrapNone `xml:"wrapNone,omitempty"`

	// 6.2. wrapSquare
	WrapSquare *WrapSquare `xml:"wrapSquare,omitempty"`

	// 6.3. wrapThrough
	WrapThrough *WrapThrough `xml:"wrapThrough,omitempty"`

	// 6.4. wrapTopAndBottom
	WrapTopBtm *WrapTopBtm `xml:"wrapTopAndBottom,omitempty"`

	// 7. Drawing Object Non-Visual Properties
	DocProp DocProp `xml:"docPr"`

	// 8. Common DrawingML Non-Visual Properties
	CNvGraphicFramePr *NonVisualGraphicFrameProp `xml:"cNvGraphicFramePr,omitempty"`

	// 9. Graphic Object
	Graphic Graphic `xml:"graphic"`
}

func NewAnchor() *Anchor {
	return &Anchor{}
}

func (a Anchor) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "wp:anchor"

	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "behindDoc"}, Value: strconv.Itoa(a.BehindDoc)})

	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "distT"}, Value: strconv.FormatUint(uint64(a.DistT), 10)})
	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "distB"}, Value: strconv.FormatUint(uint64(a.DistB), 10)})
	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "distL"}, Value: strconv.FormatUint(uint64(a.DistL), 10)})
	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "distR"}, Value: strconv.FormatUint(uint64(a.DistR), 10)})
	if a.AnchorId != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "wp14:anchorId"}, Value: string(*a.AnchorId)})
	}
	if a.EditId != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "wp14:editId"}, Value: string(*a.EditId)})
	}

	if a.SimplePosAttr != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "simplePos"}, Value: strconv.Itoa(*a.SimplePosAttr)})
	}

	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "locked"}, Value: strconv.Itoa(a.Locked)})
	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "layoutInCell"}, Value: strconv.Itoa(a.LayoutInCell)})
	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "allowOverlap"}, Value: strconv.Itoa(a.AllowOverlap)})
	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "relativeHeight"}, Value: strconv.Itoa(a.RelativeHeight)})
	if a.Hidden != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "hidden"}, Value: strconv.Itoa(*a.Hidden)})
	}

	err := e.EncodeToken(start)
	if err != nil {
		return err
	}

	// The sequence (order) of these element is important

	// 1. SimplePos
	if err := a.SimplePos.MarshalXML(e, xml.StartElement{
		Name: xml.Name{Local: "wp:simplePos"},
	}); err != nil {
		return fmt.Errorf("simplePos: %v", err)
	}

	// 2. PositionH
	if err := a.PositionH.MarshalXML(e, xml.StartElement{}); err != nil {
		return fmt.Errorf("PositionH: %v", err)
	}

	// 3. PositionH
	if err := a.PositionV.MarshalXML(e, xml.StartElement{}); err != nil {
		return fmt.Errorf("PositionV: %v", err)
	}

	// 4. Extent
	if err := a.Extent.MarshalXML(e, xml.StartElement{Name: xml.Name{Local: "wp:extent"}}); err != nil {
		return fmt.Errorf("Extent: %v", err)
	}

	// 5. EffectExtent
	if err := a.EffectExtent.MarshalXML(e, xml.StartElement{}); err != nil {
		return fmt.Errorf("EffectExtent: %v", err)
	}
	// 6. New:WrapTight
	//if err := a.WrapTight.MarshalXML(e, xml.StartElement{}); err != nil {
	//	return fmt.Errorf("WrapTight: %v", err)
	//}
	// 6. Wrap Choice
	if err := a.MarshalWrap(e); err != nil {
		return err
	}

	// 7. DocProp
	if err := a.DocProp.MarshalXML(e, xml.StartElement{}); err != nil {
		return err
	}

	// 8. CNvGraphicFramePr
	if a.CNvGraphicFramePr != nil {
		if err := a.CNvGraphicFramePr.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}

	// 9. Graphic
	if err := a.Graphic.MarshalXML(e, xml.StartElement{}); err != nil {
		return err
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

func (a *Anchor) MarshalWrap(e *xml.Encoder) error {
	if a.WrapNone != nil {
		return a.WrapNone.MarshalXML(e, xml.StartElement{})
	} else if a.WrapSquare != nil {
		return a.WrapSquare.MarshalXML(e, xml.StartElement{})
	} else if a.WrapThrough != nil {
		return a.WrapThrough.MarshalXML(e, xml.StartElement{})
	} else if a.WrapTopBtm != nil {
		return a.WrapTopBtm.MarshalXML(e, xml.StartElement{})
	} else if a.WrapTight != nil { // <--- VVVV 添加这个 else if VVVV
		// 确保使用正确的标签名
		return a.WrapTight.MarshalXML(e, xml.StartElement{
			Name: xml.Name{Local: "wrapTight"},
		})
	}
	return nil
}

func (a *Anchor) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "simplePos":
			var i int
			i, err = getInt(attr.Value)
			a.SimplePosAttr = &i
		case "distT":
			a.DistT, err = getUint(attr.Value)
		case "distB":
			a.DistB, err = getUint(attr.Value)
		case "distL":
			a.DistL, err = getUint(attr.Value)
		case "distR":
			a.DistR, err = getUint(attr.Value)
		case "relativeHeight":
			a.RelativeHeight, err = getInt(attr.Value)
		case "layoutInCell":
			a.LayoutInCell, err = getInt(attr.Value)
		case "behindDoc":
			a.BehindDoc, err = getInt(attr.Value)
		case "locked":
			a.Locked, err = getInt(attr.Value)
		case "anchorId":
			a.AnchorId = internal.ToPtr(stypes.LongHexNum(attr.Value))
		case "editId":
			a.EditId = internal.ToPtr(stypes.LongHexNum(attr.Value))
		case "allowOverlap":
			a.AllowOverlap, err = getInt(attr.Value)
		case "hidden":
			var i int
			i, err = getInt(attr.Value)
			a.Hidden = &i
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
			case xml.Name{Space: constants.WMLDrawingNS, Local: "simplePos"}:
				if err = d.DecodeElement(&a.SimplePos, &elem); err != nil {
					return err
				}
			case xml.Name{Space: constants.WMLDrawingNS, Local: "positionH"}:
				if err = d.DecodeElement(&a.PositionH, &elem); err != nil {
					return err
				}
			case xml.Name{Space: constants.WMLDrawingNS, Local: "positionV"}:
				if err = d.DecodeElement(&a.PositionV, &elem); err != nil {
					return err
				}
			case xml.Name{Space: constants.WMLDrawingNS, Local: "extent"}:
				if err = d.DecodeElement(&a.Extent, &elem); err != nil {
					return err
				}
			case xml.Name{Space: constants.WMLDrawingNS, Local: "effectExtent"}:
				a.EffectExtent = new(EffectExtent)
				if err = d.DecodeElement(a.EffectExtent, &elem); err != nil {
					return err
				}
			case xml.Name{Space: constants.WMLDrawingNS, Local: "wrapNone"}:
				a.WrapNone = new(WrapNone)
				if err = d.DecodeElement(a.WrapNone, &elem); err != nil {
					return err
				}
			case xml.Name{Space: constants.WMLDrawingNS, Local: "wrapTight"}:
				a.WrapTight = new(WrapTight)
				if err = d.DecodeElement(a.WrapTight, &elem); err != nil {
					return err
				}
			case xml.Name{Space: constants.WMLDrawingNS, Local: "wrapSquare"}:
				a.WrapSquare = new(WrapSquare)
				if err = d.DecodeElement(a.WrapSquare, &elem); err != nil {
					return err
				}
			case xml.Name{Space: constants.WMLDrawingNS, Local: "wrapThrough"}:
				a.WrapThrough = new(WrapThrough)
				if err = d.DecodeElement(a.WrapThrough, &elem); err != nil {
					return err
				}
			case xml.Name{Space: constants.WMLDrawingNS, Local: "wrapTopAndBottom"}:
				a.WrapTopBtm = new(WrapTopBtm)
				if err = d.DecodeElement(a.WrapTopBtm, &elem); err != nil {
					return err
				}
			case xml.Name{Space: constants.WMLDrawingNS, Local: "docPr"}:
				if err = d.DecodeElement(&a.DocProp, &elem); err != nil {
					return err
				}
			case xml.Name{Space: constants.WMLDrawingNS, Local: "cNvGraphicFramePr"}:
				a.CNvGraphicFramePr = new(NonVisualGraphicFrameProp)
				if err = d.DecodeElement(a.CNvGraphicFramePr, &elem); err != nil {
					return err
				}
			case xml.Name{Space: constants.DrawingMLMainNS, Local: "graphic"}:
				if err = d.DecodeElement(&a.Graphic, &elem); err != nil {
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
