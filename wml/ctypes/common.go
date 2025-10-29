package ctypes

import (
	"encoding/xml"
	"strconv"
)

// Generic Element with Single Val attribute
type GenSingleStrVal[T ~string] struct {
	Val T `xml:"val,attr"`
}

func (g GenSingleStrVal[T]) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "w:val"}, Value: string(g.Val)})
	return e.EncodeElement("", start)
}

func NewGenSingleStrVal[T ~string](val T) *GenSingleStrVal[T] {
	return &GenSingleStrVal[T]{
		Val: val,
	}
}

// Generic Element with Optional Single Val attribute
type GenOptStrVal[T ~string] struct {
	Val *T `xml:"val,attr"`
}

func (g GenOptStrVal[T]) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if g.Val != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "w:val"}, Value: string(*g.Val)})
	}
	return e.EncodeElement("", start)
}

func NewGenOptStrVal[T ~string](val T) *GenOptStrVal[T] {
	return &GenOptStrVal[T]{
		Val: &val,
	}
}

// CTString - Generic Element that has only one string-type attribute
// And the String type does not have validation
// dont use this if the element requires validation
type CTString struct {
	Val         string  `xml:"val,attr"`
	FirstRow    *string `xml:"firstRow,attr"`
	LastRow     *string `xml:"lastRow,attr"`
	FirstColumn *string `xml:"firstColumn,attr"`
	LastColumn  *string `xml:"lastColumn,attr"`
	NoHBand     *string `xml:"noHBand,attr"`
	NoVBand     *string `xml:"noVBand,attr"`
}

func NewCTString(value string) *CTString {
	return &CTString{
		Val: value,
	}
}

// MarshalXML implements the xml.Marshaler interface for the CTString type.
// It encodes the instance into XML using the "w:ELEMENT_NAME" element with a "w:val" attribute.
func (s CTString) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if &s.Val != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "w:val"}, Value: s.Val})
	}
	if s.FirstRow != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "w:firstRow"}, Value: *s.FirstRow})
	}
	if s.LastRow != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "w:lastRow"}, Value: *s.LastRow})
	}
	if s.FirstColumn != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "w:firstColumn"}, Value: *s.FirstColumn})
	}
	if s.LastColumn != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "w:lastColumn"}, Value: *s.LastColumn})
	}
	if s.NoHBand != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "w:noHBand"}, Value: *s.NoHBand})
	}
	if s.NoVBand != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "w:noVBand"}, Value: *s.NoVBand})
	}
	err := e.EncodeElement("", start)

	return err
}

// UnmarshalXML implements the xml.Marshaler interface for the CTString type.
// It encodes the instance into XML using the "w:ELEMENT_NAME" element with a "w:val" attribute.
func (s *CTString) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "val" {
			s.Val = attr.Value
		}
		if attr.Name.Local == "firstRow" {
			s.FirstRow = &attr.Value
		}
		if attr.Name.Local == "lastRow" {
			s.LastRow = &attr.Value
		}
		if attr.Name.Local == "firstColumn" {
			s.FirstColumn = &attr.Value
		}
		if attr.Name.Local == "lastColumn" {
			s.LastColumn = &attr.Value
		}
		if attr.Name.Local == "noHBand" {
			s.NoHBand = &attr.Value
		}
		if attr.Name.Local == "noVBand" {
			s.NoVBand = &attr.Value
		}
	}
	return d.Skip() // 空元素
}

type DecimalNum struct {
	Val int `xml:"val,attr"`
}

func NewDecimalNum(value int) *DecimalNum {
	return &DecimalNum{
		Val: value,
	}
}

func (s DecimalNum) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "w:val"}, Value: strconv.Itoa(s.Val)})
	err := e.EncodeElement("", start)

	return err
}

// !--- DecimalNum ends here---!

// !--- Uint64Elem starts---!

// Uint64Elem - Gomplex type that contains single val attribute which is type of uint64
// can be used where w:ST_UnsignedDecimalNumber is applicable
// example: ST_HpsMeasure
type Uint64Elem struct {
	Val uint64 `xml:"val,attr"`
}

func NewUint64Elem(value uint64) *Uint64Elem {
	return &Uint64Elem{
		Val: value,
	}
}

// MarshalXML implements the xml.Marshaler interface for the Uint64Elem type.
// It encodes the instance into XML using the "w:ELEMENT_NAME" element with a "w:val" attribute.
func (s Uint64Elem) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "w:val"}, Value: strconv.FormatUint(s.Val, 10)})
	err := e.EncodeElement("", start)

	return err
}

// !--- Uint64Elem ends here---!

type Empty struct {
}

func (s Empty) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	return e.EncodeElement("", start)
}

type Markup struct {
	ID int `xml:"id,attr"`
}

func (m Markup) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "w:id"}, Value: strconv.Itoa(m.ID)})
	return e.EncodeElement("", start)
}
