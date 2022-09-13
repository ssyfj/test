package main

import (
	"fmt"

	"github.com/shakinm/xlsReader/xls"
)

func TestGetWorkBook() {

	wb, _ := xls.OpenFile("./testfie/urltest.xls")

	s, _ := wb.GetSheet(0)
	cells1, _ := s.GetRow(0)
	cells2, _ := s.GetRow(1)

	for idx, _ := range cells1.GetCols() {
		c1, _ := cells1.GetCol(idx)
		c2, _ := cells2.GetCol(idx)
		fmt.Println(c1, c2)
		val1 := c1.GetString()
		val2 := c2.GetString()
		fmt.Println(val1, val2)
	}
}

func main() {
	TestGetWorkBook()
}
