package main

import (
	"crypto/md5"
	"encoding/hex"
	"errors"
	"fmt"
	"regexp"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/szyhf/go-excel"
)

type Operators struct {
	ID         string `xlsx:"column(id)"`
	FirstName  string `xlsx:"column(Имя)"`
	MiddleName string `xlsx:"column(Отчество)"`
	LastName   string `xlsx:"column(Фамилия)"`
	Org        string `xlsx:"column(Отдел)"`
}

func main() {
	conn := excel.NewConnecter()
	err := conn.Open("./filename.xlsx")
	if err != nil {
		panic(err)
	}
	defer conn.Close()

	rd, err := conn.NewReader("Sheet1")
	if err != nil {
		panic(err)
	}
	defer rd.Close()

	data := make([]interface{}, 0)
	for rd.Next() {
		var s Operators
		// Read a row into a struct.
		err := rd.Read(&s)
		if err != nil {
			panic(err)
		}
		s.ID = GenerateNewHash(&s)

		ff := []interface{}{s.ID, s.FirstName, s.MiddleName, s.LastName, s.Org}
		data = append(data, ff)
	}
	var field = []string{"id", "Имя", "Отчество", "Фамилия", "Отдел"}
	createNewFile(field, data)
}

func GenerateNewHash(op *Operators) string {
	operatorString := op.Org + op.LastName + op.FirstName + op.MiddleName
	editingString := EditingStrings(operatorString)
	hash := md5.Sum([]byte(editingString))
	newHash := hex.EncodeToString(hash[:])
	return newHash
}

func createNewFile(str []string, data []interface{}) {
	f := excelize.NewFile()

	//Write the first line and write the field name
	for col := 0; col < len(str); col++ {
		if colName, err := getColName(col); err == nil {
			_ = f.SetCellValue("Sheet1", fmt.Sprintf("%s%d", colName, 1), str[col])
		}
	}
	//Write data
	for row := 0; row < len(data); row++ {
		for col := 0; col < len(str); col++ {
			if colName, err := getColName(col); err == nil {
				_ = f.SetCellValue("Sheet1", fmt.Sprintf("%s%d", colName, 2+row), data[row].([]interface{})[col])
			}
		}
	}
	_ = f.SaveAs("respFileName.xlsx")
}

func getColName(length int) (string, error) {
	const asciiLength int = 26 //Number of letters A-Z
	//701 column
	if length >= asciiLength*asciiLength+asciiLength {
		return "", errors.New("column out of bounds")
	}
	Ascii := make([]string, 0)
	for i := 97; i < 97+26; i++ {
		Ascii = append(Ascii, strings.ToUpper(string(i)))
	}
	if length < asciiLength {
		return Ascii[length], nil
	} else {
		colName := Ascii[(length/asciiLength)-1] //Take the head
		colName += Ascii[length%asciiLength]     //Take the remainder as the second place
		return colName, nil
	}
}

func EditingStrings(operatorString string) string {
	if strings.HasPrefix(operatorString, "OP:") {
		operatorString = strings.TrimPrefix(operatorString, "OP:")
	}

	if strings.HasPrefix(operatorString, " ") {
		operatorString = strings.TrimPrefix(operatorString, " ")
	}

	re3 := regexp.MustCompile(`[[:punct:]]|[[:space:]]`)
	operatorString = re3.ReplaceAllString(operatorString, "")

	re := regexp.MustCompile("[0-9]+")
	operatorString = re.ReplaceAllString(operatorString, "")

	re2 := regexp.MustCompile(`[а-яА-ЯёЁ-]+`)
	operatorString = re2.FindString(operatorString)

	operatorString = strings.ToLower(operatorString)

	return operatorString
}
