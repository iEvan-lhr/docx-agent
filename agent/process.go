package agent

import (
	"github.com/iEvan-lhr/docx-agent/docx"
	"github.com/iEvan-lhr/docx-agent/wml/ctypes"
	"strings"
)

// processParagraph low-level function to process a single paragraph element.
// 这是一个底层处理函数，专门用来解析一个段落(Paragraph)元素
func (a *Agent) processParagraph(para *ctypes.Paragraph) *Text {
	// 防止传入空的段落指针
	if para == nil {
		return nil
	}

	currentText := &Text{
		Body: []*TextBody{},
	}
	var contentBuilder strings.Builder
	addSpace := true
	// 遍历段落的所有子元素 (Children)
	// 根据您的逻辑，这里的子元素主要是 Run 元素
	for _, paraChild := range para.Children {
		// 确保这个子元素是一个 Run (w:r)
		if paraChild.Run == nil {
			continue
		}
		// 遍历 Run 元素的所有子元素 (Children)
		// 这里的子元素主要是 Text (w:t) 元素
		for i, runChild := range paraChild.Run.Children {
			// 确保这个子元素是 Text (w:t) 元素，并且不是 nil
			if runChild.Text == nil {
				if runChild.Tab != nil || runChild.PTab != nil || runChild.CarrRtn != nil {
					addSpace = false
				}
				continue
			}
			if strings.TrimSpace(runChild.Text.Text) == "" {
				continue
			}
			// 提取文本内容
			textValue := runChild.Text.Text
			// 创建 TextBody
			textBody := &TextBody{
				Text:     textValue,
				addr:     &paraChild.Run.Children[i],
				run:      paraChild.Run,
				addSpace: addSpace,
			}
			if paraChild.Run.Property != nil {
				textBody.oSize = paraChild.Run.Property.Size
				textBody.oSizeCs = paraChild.Run.Property.SizeCs
			}
			addSpace = true

			// 检查并设置 Space 属性 (xml:space)
			if runChild.Text.Space != nil {
				textBody.Space = runChild.Text.Space
			}

			currentText.Body = append(currentText.Body, textBody)
			contentBuilder.WriteString(textValue)
		}
	}

	// 设置段落的完整拼接内容
	currentText.Content = contentBuilder.String()

	// 如果这个段落没有任何文本内容，我们就不需要它
	if len(currentText.Body) == 0 {
		return nil
	}

	return currentText
}

func (a *Agent) GetTable(table *docx.Table) (allTexts []*Text) {
	for _, row := range table.GetCT().RowContents {
		if row.Row == nil {
			continue
		}
		// 遍历行的所有单元格 (w:tc)
		for _, cell := range row.Row.Contents {
			if cell.Cell == nil {
				continue
			}
			// 遍历单元格的所有内容，这些内容也是段落 (w:p)
			for _, cellContent := range cell.Cell.Contents {
				if cellContent.Paragraph != nil {
					if parsedText := a.processParagraph(cellContent.Paragraph); parsedText != nil {

						allTexts = append(allTexts, parsedText)
					}
				}
				if cellContent.Table != nil {
					if parsedText := a.CtypeTable(cellContent.Table); parsedText != nil {
						t := new(Text)
						t.rootDoc = a.root
						t.document = a.root.Document
						t.body = a.root.Document.Body
						t.table = cellContent.Table
						for _, text := range parsedText {
							text.rootDoc = a.root
							text.document = a.root.Document
							text.body = a.root.Document.Body
							t.Next = append(t.Next, text)
						}
						allTexts = append(allTexts, t)
					}
				}
			}
		}
	}
	return
}

func (a *Agent) CtypeTable(cTable *ctypes.Table) (allTexts []*Text) {
	for _, row := range cTable.RowContents {
		if row.Row == nil {
			continue
		}
		// 遍历行的所有单元格 (w:tc)
		for _, cell := range row.Row.Contents {
			if cell.Cell == nil {
				continue
			}
			// 遍历单元格的所有内容，这些内容也是段落 (w:p)
			for _, cellContent := range cell.Cell.Contents {
				if cellContent.Paragraph != nil {
					if parsedText := a.processParagraph(cellContent.Paragraph); parsedText != nil {
						allTexts = append(allTexts, parsedText)
					}
				}
				if cellContent.Table != nil {
					if parsedText := a.CtypeTable(cellContent.Table); parsedText != nil {
						t := new(Text)
						t.rootDoc = a.root
						t.document = a.root.Document
						t.body = a.root.Document.Body
						t.table = cellContent.Table
						for _, text := range parsedText {
							text.rootDoc = a.root
							text.document = a.root.Document
							text.body = a.root.Document.Body
							t.Next = append(t.Next, text)
						}
						allTexts = append(allTexts, t)
					}
				}
			}
		}
	}
	return
}
