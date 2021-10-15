package excelgo

type Formulas []*Formula

func (fs Formulas) initFormula() {
	for _, f := range fs {
		f.initFormula()
	}
}

type Formula struct {
	FormulaStr string `json:"formula_str"`
	TSheet     string `json:"tsheet"`
	TColStr    string `json:"tcol_str"`
	TCol       int    `json:"-"`
}

func NewFormula(formula string, sheet string, tcolstr string) *Formula {
	return &Formula{FormulaStr: formula, TSheet: sheet, TColStr: tcolstr}
}

func (f *Formula) initFormula() {
	f.TCol = TwentysixToTen(f.TColStr)
}
