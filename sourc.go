package excelgo

import (
	"fmt"
	"strings"
)

//源頭文件
type Sourc struct {
	Path string `json:"-"`

	Rows [][]string `json:"-"`

	Headers []string `json:"-"` //標頭

	SheetName string `json:"sheet_name"`

	SortOrder map[string]Order `json:"sort_order"` //排序順序

	Cols []*Col `json:"cols"` //關鍵欄位名稱

	Formulas Formulas `json:"formulas"`
}

func NewSourc(name string, tarspans ...*Col) *Sourc {

	s := &Sourc{}
	s.SheetName = name
	s.Cols = append(s.Cols, tarspans...)

	return s
}

func (s *Sourc) Init(file FormFile) error {
	s.Path = file.Path()

	rows, err := file.Rows(s.SheetName)
	if err != nil {
		return err
	}

	s.Rows = s.deleteHeaderSpace(rows)
	s.Headers = s.Rows[0]
	s.Rows = s.Rows[1:]

	if err := s.setCol(); err != nil {
		return err
	}
	s.Formulas.initFormula()
	s.initCols()
	return nil
}

//找尋所有指定的列並進行儲存
//當查無指定的列時返回錯誤
func (s *Sourc) setCol() error {

	var forget []string
	var flag bool

	//i為列的位置 span為該躝的HEADER
	for _, ts := range s.Cols {

		flag = true

		for i, header := range s.Headers {
			if ts.Span == strings.Trim(header, " ") {
				ts.Col = i
				flag = false
				break
			}
		}

		if flag {
			forget = append(forget, ts.Span)
		}

	}

	if len(forget) > 0 {
		errmes := ""
		for _, f := range forget {
			errmes += f + " "
		}
		return fmt.Errorf("No such field found, %s", errmes)
	}

	return nil
}

//初始化所有col
func (s *Sourc) initCols() {
	for _, col := range s.Cols {
		col.InitCol()
	}
}

//依據傳入的Headers重設Col
func (s *Sourc) ResetCol(headers []string) error {
	s.Headers = headers
	if err := s.setCol(); err != nil {
		return err
	}
	s.initCols()
	return nil
}

//將該列下的所有數值填入Info中
func (s *Sourc) SetColInfo(Rows [][]string) error {

	var (
		data interface{}
		err  error
	)

	for _, col := range s.Cols {
		col.Info = make([]interface{}, len(Rows))
		for i, row := range Rows {
			data = row[col.Col]
			data, err = col.TransferFormat(row[col.Col])
			if err != nil {
				return err
			}
			col.Info[i] = data
		}
	}

	return nil
}

//將Sheet的所有內容進行指定的格式轉換別且以interface返回
func (s *Sourc) Transform(originldata [][]string) ([][]interface{}, error) {

	var (
		rows [][]interface{} = make([][]interface{}, len(originldata))
		row  []interface{}
		data interface{}
		err  error
	)

	for index, strrow := range originldata {
		row = make([]interface{}, len(strrow))
		for i, val := range strrow {
			data = val
			for _, col := range s.Cols {
				if col.Col == i {
					data, err = col.TransferFormat(val)
					if err != nil {
						return nil, err
					}
				}
			}
			row[i] = data
		}
		rows[index] = row
	}
	return rows, nil
}

func (s *Sourc) FilterAll(originldata [][]string) [][]string {
	return s.FilterByCol(originldata, s.Cols...)
}

func (s *Sourc) FilterByCol(originldata [][]string, cols ...*Col) [][]string {
	for _, col := range cols {
		originldata = col.filter(originldata)
	}
	return originldata
}

func (s *Sourc) Sum(data interface{}) error {
	for _, col := range s.Cols {
		if _, err := col.Sum(data); err != nil {
			return err
		}
	}
	return nil
}

func (s *Sourc) GetColTotal() map[string]int {
	var m map[string]int = make(map[string]int)
	for _, col := range s.Cols {
		if col.IsNumice {
			m[col.Span] = col.Total
		}
	}
	return m
}

func (s *Sourc) SetHeader(data [][]interface{}) [][]interface{} {

	old := data
	h := make([]interface{}, len(s.Headers))

	for i, val := range s.Headers {
		h[i] = val
	}
	data = make([][]interface{}, 0)
	data = append(data, h)
	data = append(data, old...)

	return data
}

//返回不含列頭的所有內容
func (s *Sourc) GetRows() [][]string {
	return s.Rows
}

//根據該列的名稱查詢
func (s *Sourc) GetCol(span string) *Col {
	for _, col := range s.Cols {
		if col.Span == span {
			return col
		}
	}
	return nil
}

func (s *Sourc) Sort(data [][]interface{}) {

	for orderspan, order := range s.SortOrder {
		for i, span := range s.Headers {
			if span == orderspan {
				Sort(data, i, order)
			}
		}
	}
	return
}

//Col迭代器
func (s *Sourc) IteratorByCol() func() (*Col, bool) {
	index := 0
	return func() (*Col, bool) {
		if index >= len(s.Cols) {
			return nil, false
		}
		col := s.Cols[index]
		index++
		return col, true
	}
}

//TCol迭代器
func (s *Sourc) IteratorByTCol() func() (*TargetCol, bool) {
	var col *Col
	var tcol *TargetCol
	colfn := s.IteratorByCol()
	tcolindex := 0
	return func() (*TargetCol, bool) {

		if col != nil && tcolindex < len(col.TCol) {
			tcol = col.TCol[tcolindex]
			tcolindex++
			return tcol, true
		}

		for {
			c, b := colfn()
			if !b {
				break
			}

			tcolindex = 0

			if len(c.TCol) > 0 {
				col = c
				tcol = c.TCol[tcolindex]
				tcolindex++
				return tcol, true
			}
		}
		return nil, false
	}
}

func (s *Sourc) deleteHeaderSpace(data [][]string) [][]string {
	index := 0
	for i, row := range data {
		if len(row) == 0 {
			index = i + 1
		} else {
			break
		}
	}
	return data[index:]
}
