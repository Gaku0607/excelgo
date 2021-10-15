package main

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"os"
	"path/filepath"

	"github.com/Gaku0607/excelgo"
)

var Environment SystemParms

type SystemParms struct {
	TC TotalCalculation `json:"total_calculation"`
}
type TotalCalculation struct {
	ManageSourc    `json:"mange_sourc"`
	InventorySourc `json:"inventory_sourc"`
	TotalPCSTCol   int `json:"total_pcs_tcol"`
	DifferenceTCol int `json:"difference_tcol"`
}

//管理
type ManageSourc struct {
	excelgo.Sourc
	CodeSpan string `json:"code_span"`
	PSCSpan  string `json:"psc_span"`
}

//庫存
type InventorySourc struct {
	excelgo.Sourc
	CodeSpan      string `json:"code_span"`
	InventorySpan string `json:"inventory_span"`
}

const jsonpath = "./examles.json"

func main() {
	if err := manage_goods_parms(&SystemParms{}); err != nil {
		panic(err.Error())
	}

	if err := manage_goods_parms_read(); err != nil {
		panic(err.Error())
	}
	for _, col := range Environment.TC.ManageSourc.Cols {
		fmt.Println(col)
	}
	fmt.Println(Environment.TC.ManageSourc.Cols)
}

//寫入
func manage_goods_parms(s *SystemParms) error {

	const (
		manage_sheet    = "基準在庫値(商品+販促物)"
		inventory_sheet = "庫存彙總表"
	)

	s.TC.TotalPCSTCol = 7
	s.TC.DifferenceTCol = 8

	//inventorySourc
	{
		s.TC.InventorySourc.InventorySpan = "結存"
		inventory := excelgo.NewCol(s.TC.InventorySourc.InventorySpan)
		inventory.Numice = excelgo.Numice{IsNumice: true}

		s.TC.InventorySourc.CodeSpan = "商品型號"
		code := excelgo.NewCol(s.TC.InventorySourc.CodeSpan)

		book := excelgo.NewCol("帳面")
		book.Filter = []string{"P100/S", "P104/S"}
		sourc := excelgo.NewSourc(
			inventory_sheet,
			book,
			inventory,
			code,
		)
		s.TC.InventorySourc.Sourc = *sourc
	}

	//ManageSourc
	{

		s.TC.ManageSourc.CodeSpan = "品番"
		codecol := excelgo.NewCol(s.TC.ManageSourc.CodeSpan)
		codecol.Impurity = excelgo.Impurity{IsSplit: true, Contains: []string{"'"}}
		codecol.TCol = []*excelgo.TargetCol{
			excelgo.NewTCol(inventory_sheet, "AA"),
			excelgo.NewTCol(inventory_sheet, "BZ"),
		}

		s.TC.ManageSourc.PSCSpan = "閾値(pcs)"
		psccol := excelgo.NewCol(s.TC.ManageSourc.PSCSpan)

		sourc := excelgo.NewSourc(
			manage_sheet,
			codecol,
			psccol,
		)

		sourc.Formulas = []*excelgo.Formula{excelgo.NewFormula("=ROUNDUP((G%d-F%d)/D%d,0)", manage_sheet, "T")}

		s.TC.ManageSourc.Sourc = *sourc

	}

	data, err := json.Marshal(&s.TC)
	if err != nil {
		return err
	}

	path, err := filepath.Abs(jsonpath)
	if err != nil {
		return err
	}

	return ioutil.WriteFile(path, data, 0777)
}

//讀取
func manage_goods_parms_read() error {

	file, err := os.OpenFile(jsonpath, os.O_RDWR|os.O_CREATE, os.ModeAppend|os.ModePerm)
	if err != nil {
		return err
	}

	defer file.Close()
	data, err := ioutil.ReadAll(file)
	if err != nil {
		return err
	}
	return json.Unmarshal(data, &Environment.TC)
}
