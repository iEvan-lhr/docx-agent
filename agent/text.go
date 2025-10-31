package agent

import (
	"encoding/json"
	"github.com/iEvan-lhr/docx-agent/docx"
	"github.com/iEvan-lhr/docx-agent/wml/ctypes"
	"reflect"
)

type TextBody struct {
	Text     string  `json:"text"`
	Space    *string `json:"space,omitempty"`
	addSpace bool
	oSize    *ctypes.FontSize
	oSizeCs  *ctypes.FontSizeCS
	addr     *ctypes.RunChild
	run      *ctypes.Run
}

type Text struct {
	Body      []*TextBody `json:"text_body"`
	Content   string      `json:"content"`
	rootDoc   *docx.RootDoc
	document  *docx.Document
	body      *docx.Body
	paragraph *docx.Paragraph
	table     *ctypes.Table
	Next      []*Text
}

func (t *Text) Get() string {
	t.tryMergeText()
	marshal, err := json.Marshal(t)
	if err != nil {
		return ""
	}
	return string(marshal)
}

func Set(text string) *Text {
	t := new(Text)
	err := json.Unmarshal([]byte(text), t)
	if err != nil {
		return nil
	}
	return t
}

// tryMergeText 尝试合并 t.Body 中具有兼容 RunChild 属性的相邻 TextBody 元素。
func (t *Text) tryMergeText() {
	if len(t.Body) < 2 {
		return // 至少需要两个元素才能合并
	}

	// 我们从 i = 1 开始，比较 i 和 i-1
	// 注意：当从切片中删除元素时，需要调整索引 i
	for i := 1; i < len(t.Body); i++ {
		curr := t.Body[i]
		prev := t.Body[i-1]

		// 如果当前元素标记为需要加空格，则它不应与前一个元素合并
		if curr.addSpace {
			continue
		}

		// 检查 RunChild 属性是否兼容
		if areRunChildrenCompatible(prev.addr, curr.addr) {
			// 1. 合并文本
			prev.Text += curr.Text

			// 2. 合并 Space 属性：后一个元素的 Space 会覆盖前一个的
			// (如果 curr.Space 为 nil, prev.Space 也会变为 nil)
			prev.Space = curr.Space
			curr.addr.Text.Text = ""
			// 3. 从 Body 切片中移除当前元素
			// t.Body[:i]       -> 从开头到 i (不包括 i)
			// t.Body[i+1:]...  -> 从 i+1 (i 之后的元素) 到末尾
			t.Body = append(t.Body[:i], t.Body[i+1:]...)

			// 4. 关键步骤：
			// 因为我们删除了索引 i 处的元素,
			// 原本在 i+1 的元素现在在索引 i 处。
			// 我们需要将 i 减 1，以便下一次循环 (i++ 之后)
			// 重新比较 'prev' (已合并) 和这个新移动过来的元素。
			i--
		}
	}
}

// areRunChildrenCompatible 检查两个 RunChild 指针是否在*除了* 'Text' 字段之外
// 的所有其他字段上都深度相等。
func areRunChildrenCompatible(r1, r2 *ctypes.RunChild) bool {
	// 如果两者都为 nil，则它们是 "相同" 的
	if r1 == nil && r2 == nil {
		return true
	}
	// 如果只有一个为 nil，则它们不同
	if r1 == nil || r2 == nil {
		return false
	}

	// 创建结构体的副本，以避免修改原始指针数据
	rc1 := *r1
	rc2 := *r2

	// **关键逻辑：**
	// 将我们想要忽略比较的 'Text' 字段在副本中设置-为 nil。
	rc1.Text = nil
	rc2.Text = nil

	// 使用 reflect.DeepEqual 比较两个修改后的副本。
	// DeepEqual 会正确地递归比较所有其他字段，包括指针和嵌套结构。
	return reflect.DeepEqual(rc1, rc2)
}
