package dmlpic

import (
	"encoding/xml"
	"fmt"
	"github.com/iEvan-lhr/docx-agent/common/constants"
	"io"
	"strconv"

	"github.com/iEvan-lhr/docx-agent/dml/dmlct"
	"github.com/iEvan-lhr/docx-agent/dml/dmlprops"
)

// Non-Visual Picture Drawing Properties
type CNvPicPr struct {
	//Relative Resize Preferred	- Default value is "true"(i.e when attr not specified).
	PreferRelativeResize *bool `xml:"preferRelativeResize,attr,omitempty"`

	//1. Picture Locks
	PicLocks *dmlprops.PicLocks `xml:"picLocks,omitempty"`

	//TODO:
	// 2. Extension List

}

func NewCNvPicPr() CNvPicPr {
	return CNvPicPr{}
}

func (c CNvPicPr) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "pic:cNvPicPr"

	if c.PreferRelativeResize != nil {
		if *c.PreferRelativeResize {
			start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "preferRelativeResize"}, Value: "true"})
		} else {
			start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "preferRelativeResize"}, Value: "false"})
		}
	}

	err := e.EncodeToken(start)
	if err != nil {
		return err
	}

	// 1. PicLocks
	if c.PicLocks != nil {
		if err := e.EncodeElement(c.PicLocks, xml.StartElement{Name: xml.Name{Local: "a:picLocks"}}); err != nil {
			return err
		}
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

// UnmarshalXML 为 CNvPicPr 实现 xml.Unmarshaler
func (c *CNvPicPr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// 1. 读取属性
	for _, attr := range start.Attr {
		if attr.Name.Local == "preferRelativeResize" {
			var prefVal bool
			prefVal, err := strconv.ParseBool(attr.Value)
			if err == nil { // 只在解析成功时设置
				c.PreferRelativeResize = &prefVal
			}
			// 可以选择记录解析错误
		}
	}

	// 2. 循环读取子元素
loop:
	for {
		token, err := d.Token()
		if err != nil {
			if err == io.EOF {
				break loop
			}
			return err
		}

		switch elem := token.(type) {
		case xml.StartElement:
			// 检查 <a:picLocks>
			if elem.Name.Local == "picLocks" && elem.Name.Space == constants.DrawingMLMainNS {
				// 假设 dmlprops.PicLocks 也有 UnmarshalXML
				c.PicLocks = new(dmlprops.PicLocks)
				if err := c.PicLocks.UnmarshalXML(d, elem); err != nil {
					return fmt.Errorf("unmarshalling PicLocks: %w", err)
				}
			} else {
				// 跳过其他不认识的子元素 (如未来的 a:extLst)
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			if elem.Name == start.Name {
				break loop // 到达 </pic:cNvPicPr>
			}
		}
	}
	return nil
}

// Non-Visual Picture Properties
type NonVisualPicProp struct {
	// 1. Non-Visual Drawing Properties
	CNvPr *dmlct.CNvPr `xml:"cNvPr,omitempty"`

	// 2.Non-Visual Picture Drawing Properties
	CNvPicPr *CNvPicPr `xml:"cNvPicPr,omitempty"`
}

func NewNVPicProp(cNvPr *dmlct.CNvPr, cNvPicPr *CNvPicPr) NonVisualPicProp {
	return NonVisualPicProp{
		CNvPr:    cNvPr,
		CNvPicPr: cNvPicPr,
	}
}

func DefaultNVPicProp(id uint, name string) NonVisualPicProp {
	cnvPicPr := NewCNvPicPr()
	cnvPicPr.PicLocks = dmlprops.DefaultPicLocks()
	return NonVisualPicProp{
		CNvPr:    dmlct.NewNonVisProp(id, name),
		CNvPicPr: &cnvPicPr,
	}
}

func (n NonVisualPicProp) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "pic:nvPicPr"

	err := e.EncodeToken(start)
	if err != nil {
		return err
	}

	// 1. cNvPr
	if err = n.CNvPr.MarshalXML(e, xml.StartElement{
		Name: xml.Name{Local: "pic:cNvPr"},
	}); err != nil {
		return fmt.Errorf("marshalling cNvPr: %w", err)
	}

	// 2. cNvPicPr
	if err = n.CNvPicPr.MarshalXML(e, xml.StartElement{
		Name: xml.Name{Local: "pic:cNvPicPr"},
	}); err != nil {
		return fmt.Errorf("marshalling cNvPicPr: %w", err)
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

// UnmarshalXML 为 NonVisualPicProp 实现 xml.Unmarshaler
func (p *NonVisualPicProp) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
loop:
	for {
		token, err := d.Token()
		if err != nil {
			if err == io.EOF {
				break loop
			}
			return err
		}

		switch elem := token.(type) {
		case xml.StartElement:
			picNS := constants.DrawingMLPicNS // pic 命名空间
			if elem.Name.Local == "cNvPr" && elem.Name.Space == picNS {
				// 假设 dmlct.CNvPr 有 UnmarshalXML
				p.CNvPr = new(dmlct.CNvPr)
				if err := p.CNvPr.UnmarshalXML(d, elem); err != nil {
					return fmt.Errorf("unmarshalling CNvPr: %w", err)
				}
			} else if elem.Name.Local == "cNvPicPr" && elem.Name.Space == picNS {
				// 假设本包的 CNvPicPr 有 UnmarshalXML
				p.CNvPicPr = new(CNvPicPr)
				if err := p.CNvPicPr.UnmarshalXML(d, elem); err != nil {
					return fmt.Errorf("unmarshalling CNvPicPr: %w", err)
				}
			} else {
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			if elem.Name == start.Name {
				break loop // 到达 </pic:nvPicPr>
			}
		}
	}
	return nil
}
