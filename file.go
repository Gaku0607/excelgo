package excelgo

import (
	"encoding/csv"

	"github.com/xuri/excelize/v2"
)

type FormFile interface {
	Path() string
	Rows(string) ([][]string, error)
}

type XlsxFile struct {
	File *excelize.File
}

func NewXlsxFile(f *excelize.File) FormFile {
	x := &XlsxFile{}
	x.File = f
	return x
}

func (x *XlsxFile) Path() string {
	return x.File.Path
}

func (x *XlsxFile) Rows(sheet string) ([][]string, error) {
	return x.File.GetRows(sheet)
}

type CsvFile struct {
	File *csv.Reader
	path string
}

func NewCsvFile(f *csv.Reader, path string) FormFile {
	c := &CsvFile{}
	c.File = f
	c.path = path
	return c
}

func (c *CsvFile) Path() string {
	return c.path
}

func (c *CsvFile) Rows(sheet string) ([][]string, error) {
	return c.File.ReadAll()
}
