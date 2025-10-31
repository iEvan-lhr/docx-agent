package ctypes

import (
	"encoding/xml"
	"github.com/iEvan-lhr/docx-agent/wml/stypes"
)

// Document Final Section Properties : w:sectPr
type SectionProp struct {
	// 支持多个 header/footer references (first, odd, even)
	HeaderReferences []*HeaderReference                     `xml:"headerReference,omitempty"`
	FooterReferences []*FooterReference                     `xml:"footerReference,omitempty"`
	PageSize         *PageSize                              `xml:"pgSz,omitempty"`
	Type             *GenSingleStrVal[stypes.SectionMark]   `xml:"type,omitempty"`
	PageMargin       *PageMargin                            `xml:"pgMar,omitempty"`
	PageNum          *PageNumbering                         `xml:"pgNumType,omitempty"`
	FormProt         *GenSingleStrVal[stypes.OnOff]         `xml:"formProt,omitempty"`
	TitlePg          *GenSingleStrVal[stypes.OnOff]         `xml:"titlePg,omitempty"`
	TextDir          *GenSingleStrVal[stypes.TextDirection] `xml:"textDirection,omitempty"`
	DocGrid          *DocGrid                               `xml:"docGrid,omitempty"`
	PgBorders        *PgBorders                             `xml:"pgBorders,omitempty"`
	Cols             *Cols                                  `xml:"cols,omitempty"`
}

func NewSectionProper() *SectionProp {
	return &SectionProp{}
}

func (s SectionProp) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "w:sectPr"

	err := e.EncodeToken(start)
	if err != nil {
		return err
	}

	// 序列化所有 header references
	for _, headerRef := range s.HeaderReferences {
		if err := headerRef.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}

	// 序列化所有 footer references
	for _, footerRef := range s.FooterReferences {
		if err := footerRef.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}

	if s.Type != nil {
		if err := s.Type.MarshalXML(e, xml.StartElement{
			Name: xml.Name{Local: "w:type"},
		}); err != nil {
			return err
		}
	}

	if s.PageSize != nil {
		if err := s.PageSize.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}

	if s.PageMargin != nil {
		if err = s.PageMargin.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	if s.PgBorders != nil {
		if err = s.PgBorders.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	if s.Cols != nil {
		if err = s.Cols.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	if s.PageNum != nil {
		if err = s.PageNum.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}

	if s.FormProt != nil {
		if err = s.FormProt.MarshalXML(e, xml.StartElement{
			Name: xml.Name{Local: "w:formProt"},
		}); err != nil {
			return err
		}
	}

	if s.TitlePg != nil {
		if err = s.TitlePg.MarshalXML(e, xml.StartElement{
			Name: xml.Name{Local: "w:titlePg"},
		}); err != nil {
			return err
		}
	}

	if s.TextDir != nil {
		if s.TextDir.MarshalXML(e, xml.StartElement{
			Name: xml.Name{Local: "w:textDirection"},
		}); err != nil {
			return err
		}
	}

	if s.DocGrid != nil {
		if s.DocGrid.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

// UnmarshalXML 实现自定义的 XML 解析
func (s *SectionProp) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// 初始化切片
	s.HeaderReferences = []*HeaderReference{}
	s.FooterReferences = []*FooterReference{}

	for {
		currentToken, err := d.Token()
		if err != nil {
			return err
		}

		switch elem := currentToken.(type) {
		case xml.StartElement:
			switch elem.Name.Local {
			case "headerReference":
				headerRef := &HeaderReference{}
				// 手动解析属性以处理命名空间
				for _, attr := range elem.Attr {
					switch attr.Name.Local {
					case "type":
						headerRef.Type = stypes.HdrFtrType(attr.Value)
					case "id":
						headerRef.ID = attr.Value
					}
				}
				// 消费结束标签
				if err := d.Skip(); err != nil {
					return err
				}
				s.HeaderReferences = append(s.HeaderReferences, headerRef)

			case "footerReference":
				footerRef := &FooterReference{}
				// 手动解析属性以处理命名空间
				for _, attr := range elem.Attr {
					switch attr.Name.Local {
					case "type":
						footerRef.Type = stypes.HdrFtrType(attr.Value)
					case "id":
						footerRef.ID = attr.Value
					}
				}
				// 消费结束标签
				if err := d.Skip(); err != nil {
					return err
				}
				s.FooterReferences = append(s.FooterReferences, footerRef)

			case "pgSz":
				s.PageSize = &PageSize{}
				if err := d.DecodeElement(s.PageSize, &elem); err != nil {
					return err
				}

			case "type":
				s.Type = &GenSingleStrVal[stypes.SectionMark]{}
				// 解析 type 元素的 val 属性
				for _, attr := range elem.Attr {
					if attr.Name.Local == "val" {
						s.Type.Val = stypes.SectionMark(attr.Value)
					}
				}
				if err := d.Skip(); err != nil {
					return err
				}

			case "pgMar":
				s.PageMargin = &PageMargin{}
				if err := d.DecodeElement(s.PageMargin, &elem); err != nil {
					return err
				}
			case "pgBorders":
				s.PgBorders = &PgBorders{}
				if err := d.DecodeElement(s.PgBorders, &elem); err != nil {
					return err
				}
			case "cols":
				s.Cols = &Cols{}
				if err := d.DecodeElement(s.Cols, &elem); err != nil {
					return err
				}
			case "pgNumType":
				s.PageNum = &PageNumbering{}
				if err := d.DecodeElement(s.PageNum, &elem); err != nil {
					return err
				}

			case "formProt":
				s.FormProt = &GenSingleStrVal[stypes.OnOff]{}
				// 解析 formProt 元素的 val 属性
				for _, attr := range elem.Attr {
					if attr.Name.Local == "val" {
						s.FormProt.Val = stypes.OnOff(attr.Value)
					}
				}
				if err := d.Skip(); err != nil {
					return err
				}

			case "titlePg":
				s.TitlePg = &GenSingleStrVal[stypes.OnOff]{}
				// 解析 titlePg 元素的 val 属性
				hasVal := false
				for _, attr := range elem.Attr {
					if attr.Name.Local == "val" {
						s.TitlePg.Val = stypes.OnOff(attr.Value)
						hasVal = true
					}
				}
				// 如果没有 val 属性，titlePg 的存在本身就表示 true
				if !hasVal {
					s.TitlePg.Val = stypes.OnOff("true")
				}
				if err := d.Skip(); err != nil {
					return err
				}

			case "textDirection":
				s.TextDir = &GenSingleStrVal[stypes.TextDirection]{}
				// 解析 textDirection 元素的 val 属性
				for _, attr := range elem.Attr {
					if attr.Name.Local == "val" {
						s.TextDir.Val = stypes.TextDirection(attr.Value)
					}
				}
				if err := d.Skip(); err != nil {
					return err
				}

			case "docGrid":
				s.DocGrid = &DocGrid{}
				if err := d.DecodeElement(s.DocGrid, &elem); err != nil {
					return err
				}

			default:
				// 跳过未知元素，避免解析错误
				if err = d.Skip(); err != nil {
					return err
				}
			}

		case xml.EndElement:
			// 遇到 sectPr 的结束标签，返回
			return nil
		}
	}
}

// 辅助方法：获取第一个 header reference（保持向后兼容）
func (s *SectionProp) GetHeaderReference() *HeaderReference {
	if len(s.HeaderReferences) > 0 {
		return s.HeaderReferences[0]
	}
	return nil
}

// 辅助方法：获取第一个 footer reference（保持向后兼容）
func (s *SectionProp) GetFooterReference() *FooterReference {
	if len(s.FooterReferences) > 0 {
		return s.FooterReferences[0]
	}
	return nil
}

// GetHeaderReferenceByType 辅助方法：根据类型获取 header reference
func (s *SectionProp) GetHeaderReferenceByType(hdrType stypes.HdrFtrType) *HeaderReference {
	for _, ref := range s.HeaderReferences {
		if ref.Type == hdrType {
			return ref
		}
	}
	return nil
}

// GetFooterReferenceByType 辅助方法：根据类型获取 footer reference
func (s *SectionProp) GetFooterReferenceByType(ftrType stypes.HdrFtrType) *FooterReference {
	for _, ref := range s.FooterReferences {
		if ref.Type == ftrType {
			return ref
		}
	}
	return nil
}
