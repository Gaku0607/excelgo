package excelgo

import "github.com/xuri/excelize/v2"

//ORIGIN_STYLE
var (
	Top              = excelize.Border{Type: "top", Style: 1, Color: "DADEE0"}
	Left             = excelize.Border{Type: "left", Style: 1, Color: "DADEE0"}
	Right            = excelize.Border{Type: "right", Style: 1, Color: "DADEE0"}
	Bottom           = excelize.Border{Type: "bottom", Style: 1, Color: "DADEE0"}
	Fill             = excelize.Fill{Type: "pattern", Pattern: 1}
	Font             = &excelize.Font{Size: 12, Family: "Microsoft JhengHei"}
	Alignment_Center = &excelize.Alignment{Horizontal: "center", Vertical: "center"}
	Alignment_Right  = &excelize.Alignment{Horizontal: "right", Vertical: "center"}
	Alignment_Left   = &excelize.Alignment{Horizontal: "left", Vertical: "center"}
	Origin_Style     = excelize.Style{
		Border:    []excelize.Border{Top, Left, Right, Bottom},
		Font:      Font,
		Fill:      Fill,
		Alignment: Alignment_Center,
	}
)
