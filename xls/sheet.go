package xls

import (
	"bytes"
	"fmt"

	"github.com/shakinm/xlsReader/helpers"
	"github.com/shakinm/xlsReader/xls/record"
	"github.com/shakinm/xlsReader/xls/structure"
)

type rw struct {
	cols map[int]structure.CellData
}

type Sheet struct {
	boundSheet    *record.BoundSheet
	rows          map[int]*rw
	wb            *Workbook
	maxCol        int // maxCol index, countCol=maxCol+1
	maxRow        int // maxRow index, countRow=maxRow+1
	hasAutofilter bool
}

func (s *Sheet) GetName() string {
	return s.boundSheet.GetName()
}

// Get row by index

func (s *Sheet) GetRow(index int) (row *rw, err error) {

	if row, ok := s.rows[index]; ok {
		return row, err
	} else {
		r := new(rw)
		r.cols = make(map[int]structure.CellData)
		return r, nil
	}
}

func (rw *rw) GetCol(index int) (c structure.CellData, err error) {

	if col, ok := rw.cols[index]; ok {
		return col, err
	} else {
		c = new(record.FakeBlank)
		return c, nil
	}

}

func (rw *rw) GetCols() (cols []structure.CellData) {

	var maxColKey int

	for k, _ := range rw.cols {
		if k > maxColKey {
			maxColKey = k
		}
	}

	for i := 0; i <= maxColKey; i++ {
		if rw.cols[i] == nil {
			cols = append(cols, new(record.FakeBlank))
		} else {
			cols = append(cols, rw.cols[i])
		}
	}

	return cols
}

// Get all rows
func (s *Sheet) GetRows() (rows []*rw) {
	for i := 0; i <= s.GetNumberRows()-1; i++ {
		if s.rows[i] == nil {
			r := new(rw)
			r.cols = make(map[int]structure.CellData)
			rows = append(rows, r)
		} else {
			rows = append(rows, s.rows[i])
		}
	}

	return rows
}

// Get number of rows
func (s *Sheet) GetNumberRows() (n int) {

	var maxRowKey int

	for k, _ := range s.rows {
		if k > maxRowKey {
			maxRowKey = k
		}
	}

	return maxRowKey + 1
}

func (s *Sheet) read(stream []byte) (err error) { // nolint: gocyclo	很重要的！！！！

	var point int64
	point = int64(helpers.BytesToUint32(s.boundSheet.LbPlyPos[:])) //这里是数据的绝对偏移量！！！！，可以进行数据的读取
	var sPoint int64
	eof := false
	records := make(map[string]string)
Next:

	recordNumber := stream[point : point+2] //找到对应流的前4字节，获取类型和长度信息
	recordDataLength := int64(helpers.BytesToUint16(stream[point+2 : point+4]))
	sPoint = point + 4
	records[fmt.Sprintf("%x", recordNumber)] = fmt.Sprintf("%x", recordNumber) //赋值个寂寞
	fmt.Println(recordNumber)
	//下拉箭头计数
	if bytes.Compare(recordNumber, record.AutofilterInfoRecord[:]) == 0 {
		c := new(record.AutofilterInfo)
		c.Read(stream[sPoint : sPoint+recordDataLength])
		if c.GetCountEntries() > 0 {
			s.hasAutofilter = true
		} else {
			s.hasAutofilter = false
		}
		goto EIF

	}

	//----------单元格可以包括多种类的数据，对于每类的数据，我们单独进行处理------LabelSStRecord（保存在SST中的字符串常量），LabelRecord（字符串常量），

	//LABELSST - String constant that uses BIFF8 shared string table (new to BIFF8)	//单元格值，字符串常量/SST，开始读取单元格数据了！！！！
	if bytes.Compare(recordNumber, record.LabelSStRecord[:]) == 0 { //14bytes，针对字符串值已经在SST中保存，这里只保存其对应的序号（实际的数据在sst中，这里只保留索引过去的序号）
		fmt.Println(66661)
		c := new(record.LabelSSt)                                 //第几行、第几列、XFRecord索引值、共享字符串表索引值
		c.Read(stream[sPoint:sPoint+recordDataLength], &s.wb.sst) //保存sst，到LabelSSt中，并获取stream前面10字节，为索引信息放入LabelSSt中
		s.addCell(c, c.GetRow(), c.GetCol())                      //获取了行列信息，添加到sheet中去！！！！
		fmt.Println(c.GetString())
		goto EIF //读取完成基本就退出了
	}

	//LABEL - Cell Value, String Constant	字符串常量
	if bytes.Compare(recordNumber, record.LabelRecord[:]) == 0 {
		fmt.Println(66662)
		if bytes.Compare(s.wb.vers[:], record.FlagBIFF8) == 0 {
			c := new(record.LabelBIFF8)
			c.Read(stream[sPoint : sPoint+recordDataLength])
			s.addCell(c, c.GetRow(), c.GetCol())
		} else {
			c := new(record.LabelBIFF5)
			c.Read(stream[sPoint : sPoint+recordDataLength])
			s.addCell(c, c.GetRow(), c.GetCol())
		}

		goto EIF
	}

	if bytes.Compare(recordNumber, []byte{0xFD, 0x00}) == 0 { //多余，就是上面的LabelSStRecord
		//todo: сделать
		goto EIF
	}

	//ARRAY - An array-entered formula  数组记录描述了数组输入到一系列单元格中的公式。
	if bytes.Compare(recordNumber, record.ArrayRecord[:]) == 0 {
		//todo: сделать
		goto EIF
	}
	//BLANK - An empty col	空白列单元格
	if bytes.Compare(recordNumber, record.BlankRecord[:]) == 0 {
		fmt.Println(66663)
		c := new(record.Blank)
		c.Read(stream[sPoint : sPoint+recordDataLength])
		s.addCell(c, c.GetRow(), c.GetCol())
		goto EIF
	}

	//BOOLERR - A Boolean or error value	布尔值单元格
	if bytes.Compare(recordNumber, record.BoolErrRecord[:]) == 0 {
		fmt.Println(66664)
		c := new(record.BoolErr)
		c.Read(stream[sPoint : sPoint+recordDataLength])
		s.addCell(c, c.GetRow(), c.GetCol())
		goto EIF
	}

	//FORMULA - A col formula, stored as parse tokens	描述包含公式的单元格。
	if bytes.Compare(recordNumber, record.FormulaRecord[:]) == 0 {
		//todo: сделать
		goto EIF
	}

	//NUMBER  - An IEEE floating-point number	浮点数信息
	if bytes.Compare(recordNumber, record.NumberRecord[:]) == 0 {
		fmt.Println(66665)
		c := new(record.Number)
		c.Read(stream[sPoint : sPoint+recordDataLength])
		s.addCell(c, c.GetRow(), c.GetCol())
		goto EIF
	}

	//MULBLANK - Multiple empty rows (new to BIFF5) 多空白记录（没看）
	if bytes.Compare(recordNumber, record.MulBlankRecord[:]) == 0 {
		fmt.Println(66667)
		c := new(record.MulBlank)
		c.Read(stream[sPoint : sPoint+recordDataLength])
		blRecords := c.GetArrayBlRecord()
		for i := 0; i <= len(blRecords)-1; i++ {
			s.addCell(blRecords[i].Get(), blRecords[i].GetRow(), blRecords[i].GetCol())
		}
		goto EIF
	}

	//RK - An RK number  Excel使用称为RK数字的内部数字类型来节省内存和磁盘空间
	if bytes.Compare(recordNumber, record.RkRecord[:]) == 0 {
		fmt.Println(66668)
		c := new(record.Rk)
		c.Read(stream[sPoint : sPoint+recordDataLength])
		s.addCell(c, c.GetRow(), c.GetCol())
		goto EIF
	}

	//MULRK - Multiple RK numbers (new to BIFF5)
	if bytes.Compare(recordNumber, record.MulRKRecord[:]) == 0 {
		fmt.Println(66669)
		c := new(record.MulRk)
		c.Read(stream[sPoint : sPoint+recordDataLength])
		rkRecords := c.GetArrayRKRecord()
		for i := 0; i <= len(rkRecords)-1; i++ {
			s.addCell(rkRecords[i].Get(), rkRecords[i].GetRow(), rkRecords[i].GetCol())
		}
		goto EIF

	}

	//RSTRING - Cell with character formatting
	if bytes.Compare(recordNumber, record.RStringRecord[:]) == 0 {
		//todo: сделать
		goto EIF
	}

	//SHRFMLA - A shared formula (new to BIFF5)
	if bytes.Compare(recordNumber, record.SharedFormulaRecord[:]) == 0 {
		//todo: сделать
		goto EIF
	}

	//STRING - A string that represents the result of a formula
	if bytes.Compare(recordNumber, record.StringRecord[:]) == 0 {
		//todo: сделать
		goto EIF
	}

	if bytes.Compare(recordNumber, record.RowRecord[:]) == 0 { //描述一行，多少个0x0208就有多少行
		//todo: сделать
		goto EIF
	}

	//EOF
	if bytes.Compare(recordNumber, record.EOFRecord[:]) == 0 && recordDataLength == 0 {
		eof = true
	}
EIF:
	point = point + recordDataLength + 4
	if !eof {
		goto Next
	}

	return

}

func (s *Sheet) addCell(cd structure.CellData, row [2]byte, column [2]byte) {

	r := int(helpers.BytesToUint16(row[:]))
	c := int(helpers.BytesToUint16(column[:]))

	if s.rows == nil {
		s.rows = map[int]*rw{}
	}
	if _, ok := s.rows[r]; !ok {
		s.rows[r] = new(rw)

		if _, ok := s.rows[r].cols[c]; !ok {

			colVal := map[int]structure.CellData{}
			colVal[c] = cd

			s.rows[r].cols = colVal
		}

	}
	fmt.Println("addCell")
	s.rows[r].cols[c] = cd

}
