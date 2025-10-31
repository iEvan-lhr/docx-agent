package dmlprops

import (
	"encoding/xml"
	"io"

	"github.com/iEvan-lhr/docx-agent/dml/dmlst"
)

// Picture Locks
type PicLocks struct {
	DisallowShadowGrouping dmlst.OptBool `xml:"noGrp,attr,omitempty"`
	NoSelect               dmlst.OptBool `xml:"noSelect,attr,omitempty"`
	NoRot                  dmlst.OptBool `xml:"noRot,attr,omitempty"`
	NoChangeAspect         dmlst.OptBool `xml:"noChangeAspect,attr,omitempty"`
	NoMove                 dmlst.OptBool `xml:"noMove,attr,omitempty"`
	NoResize               dmlst.OptBool `xml:"noResize,attr,omitempty"`
	NoEditPoints           dmlst.OptBool `xml:"noEditPoints,attr,omitempty"`
	NoAdjustHandles        dmlst.OptBool `xml:"noAdjustHandles,attr,omitempty"`
	NoChangeArrowheads     dmlst.OptBool `xml:"noChangeArrowheads,attr,omitempty"`
	NoChangeShapeType      dmlst.OptBool `xml:"noChangeShapeType,attr,omitempty"`
	NoCrop                 dmlst.OptBool `xml:"noCrop,attr,omitempty"`
}

func DefaultPicLocks() *PicLocks {
	return &PicLocks{
		NoChangeAspect:     dmlst.NewOptBool(true),
		NoChangeArrowheads: dmlst.NewOptBool(true),
	}
}
func (p PicLocks) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "a:picLocks"
	start.Attr = []xml.Attr{}

	if p.DisallowShadowGrouping.Valid {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "noGrp"}, Value: p.DisallowShadowGrouping.ToStringFlag()})
	}

	if p.NoSelect.Valid {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "noSelect"}, Value: p.NoSelect.ToStringFlag()})
	}
	if p.NoRot.Valid {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "noRot"}, Value: p.NoRot.ToStringFlag()})
	}
	if p.NoChangeAspect.Valid {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "noChangeAspect"}, Value: p.NoChangeAspect.ToStringFlag()})
	}
	if p.NoMove.Valid {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "noMove"}, Value: p.NoMove.ToStringFlag()})
	}
	if p.NoResize.Valid {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "noResize"}, Value: p.NoResize.ToStringFlag()})
	}
	if p.NoEditPoints.Valid {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "noEditPoints"}, Value: p.NoEditPoints.ToStringFlag()})
	}
	if p.NoAdjustHandles.Valid {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "noAdjustHandles"}, Value: p.NoAdjustHandles.ToStringFlag()})
	}
	if p.NoChangeArrowheads.Valid {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "noChangeArrowheads"}, Value: p.NoChangeArrowheads.ToStringFlag()})
	}
	if p.NoChangeShapeType.Valid {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "noChangeShapeType"}, Value: p.NoChangeShapeType.ToStringFlag()})
	}
	if p.NoCrop.Valid {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "noCrop"}, Value: p.NoCrop.ToStringFlag()})
	}

	err := e.EncodeToken(start)
	if err != nil {
		return err
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

// UnmarshalXML 为 PicLocks 实现 xml.Unmarshaler
func (p *PicLocks) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// 1. 遍历并解析所有属性
	for _, attr := range start.Attr {
		var target *dmlst.OptBool
		switch attr.Name.Local {
		case "noGrp":
			target = &p.DisallowShadowGrouping
		case "noSelect":
			target = &p.NoSelect
		case "noRot":
			target = &p.NoRot
		case "noChangeAspect":
			target = &p.NoChangeAspect
		case "noMove":
			target = &p.NoMove
		case "noResize":
			target = &p.NoResize
		case "noEditPoints":
			target = &p.NoEditPoints
		case "noAdjustHandles":
			target = &p.NoAdjustHandles
		case "noChangeArrowheads":
			target = &p.NoChangeArrowheads
		case "noChangeShapeType":
			target = &p.NoChangeShapeType
		case "noCrop":
			target = &p.NoCrop
		default:
			continue // 跳过不认识的属性
		}

		// 使用 dmlst.OptBool 的 UnmarshalXMLAttr 方法来解析布尔值
		// 它应该能处理 "1", "true", "0", "false" 等情况
		if target != nil {
			// UnmarshalXMLAttr 需要 xml.Attr 类型
			if err := target.UnmarshalXMLAttr(attr); err != nil {
				// 可以选择记录或返回属性解析错误
				// return fmt.Errorf("error parsing attribute %s: %w", attr.Name.Local, err)
			}
		}
	}

	// 2. <a:picLocks> 是空元素，消耗掉 token 直到找到结束标签
	for {
		token, err := d.Token()
		if err != nil {
			if err == io.EOF {
				return nil
			} // io.EOF 也意味着结束
			return err
		}
		if elem, ok := token.(xml.EndElement); ok && elem.Name == start.Name {
			return nil // 找到结束标签，成功返回
		}
		// 如果在结束标签前遇到其他 token (理论上不应该)，忽略它们
	}
}
