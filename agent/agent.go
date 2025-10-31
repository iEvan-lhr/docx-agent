package agent

import (
	"errors"
	godocx "github.com/iEvan-lhr/docx-agent"
	"github.com/iEvan-lhr/docx-agent/docx"
	"github.com/ievan-lhr/go-llm-client"
	"strings"
)

type Agent struct {
	Text    []*Text
	inStart bool
	root    *docx.RootDoc
}

func BuildAgent(llm *llm.Client) *Agent {
	return &Agent{}
}

func (a *Agent) Start() {
	doc := docx.NewRootDoc()
	firstT := new(Text)
	firstT.rootDoc = doc
	a.Text = append(a.Text, firstT)
}

func (a *Agent) StartWithFile(filename string) {
	document, err := godocx.OpenDocument(filename)
	if err != nil {
		panic(err)
	}
	a.root = document
	a.BuildingText()
}

func (a *Agent) Stop() {
	a.status()
}

func (a *Agent) SaveAndStop(filename string) {
	a.status()
}
func (a *Agent) status() {
	if !a.inStart {
		panic(errors.New("请先初始化Agent或重新启动Agent"))
	}
}
func (a *Agent) AddText(text string) {
	a.status()
}

func (a *Agent) FindText(text string) *Text {
	for i := range a.Text {
		findText, t := a.findText(a.Text[i], text)
		if findText {
			return t
		}
	}
	return nil
}

func (a *Agent) findText(text *Text, textCon string) (bool, *Text) {
	for i := range a.Text {
		if a.Text[i] != nil {
			if strings.Contains(a.Text[i].Content, textCon) {
				return true, a.Text[i]
			}
		}
		if a.Text[i].Next != nil {
			for j := range a.Text[i].Next {
				text, t := a.findText(a.Text[i].Next[j], textCon)
				if text {
					return true, t
				}
			}
		}
	}
	return false, nil
}
