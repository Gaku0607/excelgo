package excelgo

import "fmt"

type FormulaFunc func(int, string) string

var DefaultFormulaFunc FormulaFunc = func(index int, formula string) string { return fmt.Sprintf(formula, index) }

type SheetFormula map[string]map[string]FormulaFunc

func (v SheetFormula) SetFormulaCategory(servicename, sheet, tcolstr string, f FormulaFunc) {
	if v == nil {
		v = map[string]map[string]FormulaFunc{}
	}

	if v[servicename] == nil {
		v[servicename] = map[string]FormulaFunc{}
	}
	//sheet+tcolstr為key
	v[servicename][sheet+tcolstr] = f
}

func (v SheetFormula) GetFormulaCategory(servicename, sheet, tcolstr string) FormulaFunc {
	if services, exist := v[servicename]; exist {
		if format, exist := services[sheet+tcolstr]; exist {
			return format
		}
	}
	return DefaultFormulaFunc
}

var FormulaCategory SheetFormula = make(SheetFormula)

type Formulas []*Formula

func (fs Formulas) initFormula() {
	for _, f := range fs {
		f.initFormula()
	}
}

type Formula struct {
	FormulaStr  string      `json:"formula_str"`
	TSheet      string      `json:"tsheet"`
	TColStr     string      `json:"tcol_str"`
	TCol        int         `json:"-"`
	formulafunc FormulaFunc `json:"-"`
	ServiceName string      `json:"service_name"`
}

func NewFormula(service, formula, sheet, tcolstr string) *Formula {
	return &Formula{ServiceName: service, FormulaStr: formula, TSheet: sheet, TColStr: tcolstr}
}

func (f *Formula) initFormula() {
	f.TCol = TwentysixToTen(f.TColStr)
	FormulaCategory.GetFormulaCategory(f.ServiceName, f.TSheet, f.TColStr)
}
