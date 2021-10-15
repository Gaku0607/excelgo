package excelgo

type TcolFormatFunc func(interface{}) interface{}

var DefaultTcolFormat TcolFormatFunc = func(i interface{}) interface{} { return i }

var FormatCategory map[string]TcolFormatFunc = make(map[string]TcolFormatFunc)

func SetFormatCategory(sheet string, tcolstr string, f TcolFormatFunc) {
	if FormatCategory == nil {
		FormatCategory = make(map[string]TcolFormatFunc)
	}
	//sheet+tcolstr為key
	FormatCategory[sheet+tcolstr] = f
}

func GetFormatCategory(sheet string, tcolstr string) TcolFormatFunc {
	if f, exist := FormatCategory[sheet+tcolstr]; exist {
		return f
	} else {
		return DefaultTcolFormat
	}
}

type TargetCol struct {
	Sheet   string         `json:"sheet"`
	TCol    int            `json:"-"`
	TColStr string         `json:"tcol_str"`
	Format  TcolFormatFunc `json:"-"`
}

func NewTCol(sheet string, TColStr string) *TargetCol {
	return &TargetCol{Sheet: sheet, TColStr: TColStr}
}

func (t *TargetCol) InitFormat() {
	t.Format = GetFormatCategory(t.Sheet, t.TColStr)
}
