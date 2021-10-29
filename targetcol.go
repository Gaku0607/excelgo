package excelgo

type TcolFormatFunc func(interface{}) interface{}

var DefaultTcolFormat TcolFormatFunc = func(i interface{}) interface{} { return i }

type ValueForamt map[string]map[string]TcolFormatFunc

func (v ValueForamt) SetFormatCategory(servicename, sheet, tcolstr string, f TcolFormatFunc) {
	if v == nil {
		v = map[string]map[string]TcolFormatFunc{}
	}

	if v[servicename] == nil {
		v[servicename] = map[string]TcolFormatFunc{}
	}
	//sheet+tcolstr為key
	v[servicename][sheet+tcolstr] = f
}

func (v ValueForamt) GetFormatCategory(servicename, sheet, tcolstr string) TcolFormatFunc {
	if services, exist := v[servicename]; exist {
		if format, exist := services[sheet+tcolstr]; exist {
			return format
		}
	}
	return DefaultTcolFormat
}

var FormatCategory ValueForamt = make(ValueForamt)

type TargetCol struct {
	ParentCol   *Col           `json:"-"`
	ServiceName string         `json:"service_name"`
	Sheet       string         `json:"sheet"`
	TCol        int            `json:"-"`
	TColStr     string         `json:"tcol_str"`
	Format      TcolFormatFunc `json:"-"`
}

func NewTCol(servicename, sheet, TColStr string) *TargetCol {
	return &TargetCol{ServiceName: servicename, Sheet: sheet, TColStr: TColStr}
}

func (t *TargetCol) InitFormat() {
	t.Format = FormatCategory.GetFormatCategory(t.ServiceName, t.Sheet, t.TColStr)
}
