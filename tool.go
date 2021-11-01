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

//欄位數字轉字母欄位
func ConvertToLetter(num int) string {

	if num == 0 {
		panic("headercol is zero")
	}

	const letter = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

	var (
		str  string
		i, n int
	)

	for {
		m := 1
		n, i = calculate1(num)

		str += string(letter[n-1])

		if i == 0 {
			break
		}

		for ; i > 0; i-- {
			m *= 26
		}
		num -= m * n
	}

	return str
}

func calculate1(num int) (int, int) {

	if num <= 26 {
		return num, 0
	}

	i := 1
	for {
		num = num / 26

		if num <= 26 {
			break
		}
		i++
	}
	return num, i
}

//字母欄位轉數字
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

func calculate2(n int) int {
	sum := 1
	for i := 0; i < n; i++ {
		sum *= 26
	}
	return sum
}

//確認是否有相同檔名有則變更
func CheckFileName(path string) string {
	base := filepath.Ext(path)
	format := base
	for i := 1; ; i++ {
		if !isFileExist(path) {
			return path
		}
		path = strings.Replace(path, format, "("+strconv.Itoa(i)+")"+base, 1)
		format = "(" + strconv.Itoa(i) + ")" + base
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

//************************************************
//****************  OrderSort  *******************
//************************************************

const (
	ReverseOrder  Order = 0
	PositiveOrder Order = 1
)

type Order uint8

func (o Order) String() string {

	switch o {
	case ReverseOrder:
		return "ReverseOrder"
	case PositiveOrder:
		return "PositiverOrder"
	default:
		return "null"
	}
}

//插入排序依據指定的行列進行排序
func Sort(rows [][]interface{}, col int, order Order) {

	if len(rows) == 0 {
		panic("rows of data is nil")
	}
	switch order {
	case ReverseOrder:
		reverseSort(rows, col)
	case PositiveOrder:
		positiveSort(rows, col)
	default:
		panic("No such sorting method")
	}
}

func reverseSort(rows [][]interface{}, col int) {
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

func positiveSort(rows [][]interface{}, col int) {
	for i := 1; i < len(rows); i++ {
		ii := rows[i]
		index := i - 1
		insertVal := rows[i][col]
		targetVal := rows[index][col]

		switch insertVal.(type) {
		case string:

			for targetVal.(string) < insertVal.(string) {
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
			for targetVal.(int) < insertVal.(int) {
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
