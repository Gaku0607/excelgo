package main

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"os"

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

const (
	testfile        = "/Users/gaku/Documents/GitHub/excelgo/examples/filter/20211001.xlsx"
	testoutfilepath = "/Users/gaku/Documents/GitHub/excelgo/examples/filter/20211001_test.xlsx"
	configpath      = "../examles.json"
)

func main() {

	if err := manage_goods_parms_read(); err != nil {
		panic(err.Error())
	}

	f, err := excelgo.OpenFile(testfile)
	if err != nil {
		panic(err.Error())
	}

	s := Environment.TC.InventorySourc.Sourc

	if err := s.Init(f); err != nil {
		panic(err.Error())
	}

	rows := s.FilterAll(s.Rows)

	for _, row := range rows {
		fmt.Println(row)
	}
}

//讀取
func manage_goods_parms_read() error {

	file, err := os.OpenFile(configpath, os.O_RDWR|os.O_CREATE, os.ModeAppend|os.ModePerm)
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
