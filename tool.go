package excelgo

import (
	"encoding/csv"
	"errors"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func removecharacters(str, sep string) string {
	sli := strings.Split(str, sep)
	return strings.Join(sli, "")
}

// func ConvertToLetter(col int) string {

// 	if col == 0 {
// 		return ""
// 	}

// 	p := col / 27

// 	if p == 0 {
// 		return string((col - 1) + 65)
// 	}

// 	return ""
// }

func calculate1(col int) int {

	i := 0

	for p := col / 26; p == 0; i++ {
		p /= 26
	}

	return i
}

//計算字符欄位實際的整字
func TwentysixToTen(colstr string) int {

	sum := 0

	for i := len(colstr) - 1; i >= 0; i-- {

		// 確定了每一個字元所對應的整數

		p := int(colstr[i])

		//A的整數型為65 當整數小於Ａ大於Ｚ時返回0
		if (p-64) > 26 || (p-64) < 1 {
			return 0
		}

		m := calculate2(len(colstr) - 1 - i)

		sum += (p - 64) * m

	}
	return sum
}

//
func calculate2(n int) int {
	sum := 1
	for i := 0; i < n; i++ {
		sum *= 26
	}
	return sum
}

//確認是否有相同檔名有則變更
func CheckFileName(path string) string {
	base := ".xlsx"
	format := base
	for i := 0; ; i++ {
		if !isFileExist(path) {
			return path
		}
		path = strings.Replace(path, format, "("+strconv.Itoa(i+1)+")"+base, 1)
		format = "(" + strconv.Itoa(i+1) + ")" + base
	}
}

func isFileExist(filename string) bool {
	if _, err := os.Stat(filename); err != nil {
		if os.IsExist(err) {
			return true
		}
		return false
	}
	return true
}

//插入排序依據指定的行列進行排序
func sort(rows [][]interface{}, col int) {

	for i := 1; i < len(rows); i++ {
		ii := rows[i]
		index := i - 1
		insertVal := rows[i][col]
		targetVal := rows[index][col]

		switch insertVal.(type) {
		case string:

			for targetVal.(string) > insertVal.(string) {
				rows[index+1] = rows[index]
				index--
				if index < 0 {
					break
				}
				targetVal = rows[index][col].(string)
			}
			if index != i-1 {
				rows[index+1] = ii
			}

		case int:
			for targetVal.(int) > insertVal.(int) {
				rows[index+1] = rows[index]
				index--
				if index < 0 {
					break
				}
				targetVal = rows[index][col].(int)
			}
			if index != i-1 {
				rows[index+1] = ii
			}

		default:
			panic("Does not match the sorting format")
		}

	}
}

//************************************************
//****************  FormFile  ********************
//************************************************

//開啟多個文件
func OpenFiles(paths []string) ([]FormFile, error) {

	var files []FormFile

	for _, path := range paths {
		file, err := OpenFile(path)
		if err != nil {
			return nil, err
		}
		files = append(files, file)
	}
	return files, nil
}

//開啟文件
func OpenFile(path string) (FormFile, error) {

	base := filepath.Base(path)

	s := strings.Split(base, ".")

	switch s[len(s)-1] {
	case "xlsx":
		f, err := excelize.OpenFile(path)
		if err != nil {
			return nil, err
		}
		return NewXlsxFile(f), nil
	case "csv":
		f, err := os.Open(path)
		if err != nil {
			return nil, err
		}
		return NewCsvFile(csv.NewReader(f), path), nil
	default:
		return nil, errors.New("File format does not match")
	}
}

//初始化源文件
func initSourcFile(file FormFile, s *Sourc) error {

	s.Path = file.Path()
	rows, err := file.Rows(s.SheetName)
	if err != nil {
		return err
	}
	s.Rows = rows

	if err := s.Init(); err != nil {
		return err
	}

	return nil
}
