package excelgo

type Filter struct {
	col      int      `json:"-"`
	IsTarget bool     `json:"is_target"`
	Target   []string `josn:"target"`
}

func (f *Filter) filter(rows [][]string) [][]string {

	if len(f.Target) == 0 {
		return rows
	}

	var assertion func(string) bool

	if f.IsTarget {
		assertion = func(val string) bool {
			for _, t := range f.Target {
				if t == val {
					return true
				}
			}
			return false
		}
	} else {
		assertion = func(val string) bool {
			for _, t := range f.Target {
				if t == val {
					return false
				}
			}
			return true
		}
	}

	var newrows [][]string

	for _, row := range rows {
		if assertion(row[f.col]) {
			newrows = append(newrows, row)
		}
	}

	return newrows
}
