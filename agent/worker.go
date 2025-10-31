package agent

func (a *Agent) BuildingText() {
	var allTexts []*Text
	// 核心逻辑: 直接遍历 Body 的 Children
	// Body 的 Children 可能是段落(Para)，也可能是表格(Tbl)等其他元素
	for _, bodyChild := range a.root.Document.Body.Children {

		// 情况一：如果这个 Child 是一个段落 (w:p)
		if bodyChild.Para != nil {
			if parsedText := a.processParagraph(bodyChild.Para.GetCT()); parsedText != nil {
				parsedText.rootDoc = a.root
				parsedText.document = a.root.Document
				parsedText.body = a.root.Document.Body
				parsedText.paragraph = bodyChild.Para
				allTexts = append(allTexts, parsedText)
			}
		}

		// 情况二：如果这个 Child 是一个表格 (w:tbl)
		// 我们需要深入表格内部，遍历它的行(tr)和单元格(tc)，因为单元格里也包含段落
		if bodyChild.Table != nil {
			// 遍历表格的所有行 (w:tr)
			t := new(Text)
			t.rootDoc = a.root
			t.document = a.root.Document
			t.body = a.root.Document.Body
			t.table = bodyChild.Table.GetCT()
			for _, text := range a.GetTable(bodyChild.Table) {
				text.rootDoc = a.root
				text.document = a.root.Document
				text.body = a.root.Document.Body
				text.table = bodyChild.Table.GetCT()
				t.Next = append(t.Next, text)
			}
			allTexts = append(allTexts, t)

		}
	}
	for i := range a.root.Document.Headers {
		for j := range a.root.Document.Headers[i].Children {
			if a.root.Document.Headers[i].Children[j].Para != nil {
				if parsedText := a.processParagraph(a.root.Document.Headers[i].Children[j].Para.GetCT()); parsedText != nil {
					parsedText.rootDoc = a.root
					parsedText.document = a.root.Document
					parsedText.body = a.root.Document.Body
					parsedText.paragraph = a.root.Document.Headers[i].Children[j].Para
					allTexts = append(allTexts, parsedText)
				}
			}
			if a.root.Document.Headers[i].Children[j].Table != nil {
				t := new(Text)
				t.rootDoc = a.root
				t.document = a.root.Document
				t.body = a.root.Document.Body
				t.table = a.root.Document.Headers[i].Children[j].Table.GetCT()
				for _, text := range a.GetTable(a.root.Document.Headers[i].Children[j].Table) {
					text.rootDoc = a.root
					text.document = a.root.Document
					text.body = a.root.Document.Body
					text.table = a.root.Document.Headers[i].Children[j].Table.GetCT()
					t.Next = append(t.Next, text)
				}
				allTexts = append(allTexts, t)
			}
		}
	}
	for i := range a.root.Document.Footers {
		for j := range a.root.Document.Footers[i].Children {
			if a.root.Document.Footers[i].Children[j].Para != nil {
				if parsedText := a.processParagraph(a.root.Document.Footers[i].Children[j].Para.GetCT()); parsedText != nil {
					parsedText.rootDoc = a.root
					parsedText.document = a.root.Document
					parsedText.body = a.root.Document.Body
					parsedText.paragraph = a.root.Document.Footers[i].Children[j].Para
					allTexts = append(allTexts, parsedText)
				}
			}
			if a.root.Document.Footers[i].Children[j].Table != nil {
				t := new(Text)
				t.rootDoc = a.root
				t.document = a.root.Document
				t.body = a.root.Document.Body
				t.table = a.root.Document.Footers[i].Children[j].Table.GetCT()
				for _, text := range a.GetTable(a.root.Document.Footers[i].Children[j].Table) {
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
