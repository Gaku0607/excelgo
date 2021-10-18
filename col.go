package excelgo

import (
	"errors"
	"strconv"
)

type Numice struct {
	IsNumice bool `json:"is_numice"` // 是否轉為數字
	IsSum    bool `json:"is_sum"`    //是否加總
	Total    int  `json:"-"`         //總和
}

func (n *Numice) toNumice(val string) (interface{}, error) {
	if n.IsNumice {
		return strconv.Atoi(val)
	}
	return val, nil
}

type Impurity struct {
	IsSplit  bool     `json:"is_split"` // 是否切割去除關鍵字
	Contains []string `json:"contains"` // 關鍵字符（去除用）
}

//最小單位列
type Col struct {
	Span string `json:"span"` //欄號名稱

	Col int `json:"-"` //以數字顯示第幾欄位

	ColStr string `json:"-"` //以字符顯示第幾欄位

	TCol []*TargetCol `json:"tcol"` //合併時的目標欄位

	Impurity `json:"impurity"`

	Filter []string //篩選

	Numice `json:"numice"`

	Info []interface{} `json:"-"` //該欄位下內容  （愈被切割內容）
}

func NewCol(name string) *Col {
	s := &Col{}
	s.Span = name
	return s
}

//初始化col以及包含的tcol
func (c *Col) initCol() {
	c.Col = TwentysixToTen(c.ColStr)
	for _, tcol := range c.TCol {
		tcol.TCol = TwentysixToTen(tcol.TColStr)
		tcol.FatherCol = c
		tcol.InitFormat()
	}
}

//將該列下的內容進行指定轉換
func (c *Col) TransferFormat(val string) (interface{}, error) {
	val = c.removecharacters(val)
	data, err := c.toNumice(val)
	if err != nil {
		return nil, err
	}
	return data, nil
}

//刪除關鍵字
func (c *Col) removecharacters(val string) string {
	if c.IsSplit {
		for _, k := range c.Contains {
			val = removecharacters(val, k)
		}
	}
	return val
}

func (c *Col) Sum(data interface{}) error {
	if c.IsSum {
		var total int
		switch data.(type) {
		case [][]interface{}:
			for _, row := range data.([][]interface{}) {
				total += row[c.Col].(int)
			}
			c.Numice.Total = total

		case [][]string:
			for _, row := range data.([][]string) {
				num, _ := strconv.Atoi(row[c.Col])
				total += num
			}
			c.Numice.Total = total

		default:
			return errors.New("Does not match the sumfunc format")
		}
	}
	return nil
}
