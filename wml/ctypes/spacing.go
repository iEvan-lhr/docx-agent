package ctypes

import (
	"encoding/xml"
	"strconv"

	"github.com/iEvan-lhr/docx-agent/wml/stypes"
)

// Spacing Between Lines and Above/Below Paragraph
type Spacing struct {
	//Spacing Above Paragraph
	Before *uint64 `xml:"before,attr,omitempty"`

	//Spacing Above Paragraph IN Line Units
	BeforeLines *int `xml:"beforeLines,attr,omitempty"`

	//Spacing Below Paragraph
	After *uint64 `xml:"after,attr,omitempty"`

	// Automatically Determine Spacing Above Paragraph
	BeforeAutospacing *stypes.OnOff `xml:"beforeAutospacing,attr,omitempty"`

	// Automatically Determine Spacing Below Paragraph
	AfterAutospacing *stypes.OnOff `xml:"afterAutospacing,attr,omitempty"`

	//Spacing Between Lines in Paragraph
	Line *int `xml:"line,omitempty"`

	//Type of Spacing Between Lines
	LineRule *stypes.LineSpacingRule `xml:"lineRule,attr,omitempty"`
}

func NewParagraphSpacing(before uint64, after uint64) *Spacing {
	return &Spacing{
		Before: &before,
		After:  &after,
	}
}

func (s Spacing) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "w:spacing"

	start.Attr = []xml.Attr{}

	if s.Before != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "w:before"}, Value: strconv.FormatUint(*s.Before, 10)})
	}

	if s.After != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "w:after"}, Value: strconv.FormatUint(*s.After, 10)})
	}

	if s.BeforeLines != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "w:beforeLines"}, Value: strconv.Itoa(*s.BeforeLines)})
	}

	if s.BeforeAutospacing != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "w:beforeAutospacing"}, Value: string(*s.BeforeAutospacing)})
	}

	if s.AfterAutospacing != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "w:afterAutospacing"}, Value: string(*s.AfterAutospacing)})
	}

	if s.Line != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "w:line"}, Value: strconv.Itoa(*s.Line)})
	}

	if s.LineRule != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "w:lineRule"}, Value: string(*s.LineRule)})
	}

	return e.EncodeElement("", start)
}

func (s *Spacing) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "before" {
			atoi, err := strconv.Atoi(attr.Value)
			if err == nil {
				u := uint64(atoi)
				s.Before = &u
			}
		}
		if attr.Name.Local == "after" {
			atoi, err := strconv.Atoi(attr.Value)
			if err == nil {
				u := uint64(atoi)
				s.After = &u
			}
		}
		if attr.Name.Local == "beforeLines" {
			atoi, err := strconv.Atoi(attr.Value)
			if err == nil {
				s.BeforeLines = &atoi
			}
		}
		if attr.Name.Local == "beforeAutospacing" {
			s2 := stypes.OnOff(attr.Value)
			s.BeforeAutospacing = &s2
		}
		if attr.Name.Local == "afterAutospacing" {
			s2 := stypes.OnOff(attr.Value)
			s.AfterAutospacing = &s2
		}
		if attr.Name.Local == "line" {
			atoi, err := strconv.Atoi(attr.Value)
			if err == nil {
				s.Line = &atoi
			}
		}
		if attr.Name.Local == "lineRule" {
			s2 := stypes.LineSpacingRule(attr.Value)
			s.LineRule = &s2
		}
	}
	return d.Skip() // 空元素
}
