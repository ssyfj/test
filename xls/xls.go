package xls

import (
	"encoding/binary"
	"io"

	"github.com/shakinm/xlsReader/cfb"
)

// OpenFile - Open document from the file
func OpenFile(fileName string) (workbook Workbook, err error) {

	adaptor, err := cfb.OpenFile(fileName)

	if err != nil {
		return workbook, err
	}
	return openCfb(adaptor) //干啥呢
}

// OpenReader - Open document from the file reader
func OpenReader(fileReader io.ReadSeeker) (workbook Workbook, err error) {

	adaptor, err := cfb.OpenReader(fileReader)

	if err != nil {
		return workbook, err
	}
	return openCfb(adaptor)
}

// OpenFile - Open document from the file
func openCfb(adaptor cfb.Cfb) (workbook Workbook, err error) {
	var book *cfb.Directory
	var root *cfb.Directory
	for _, dir := range adaptor.GetDirs() { //获取文件目录流扇区数据（4个），解析文件名等数据
		fn := dir.Name() //获取name：RootEntry、Workbook、SummaryInformation、DocumentSummaryInformation（后面两个是摘要信息）

		if fn == "Workbook" { //一个wook book里面包含多个sheet
			if book == nil {
				book = dir
			}
		}
		if fn == "Book" {
			book = dir

		}
		if fn == "Root Entry" {
			root = dir
		}

	}

	if book != nil { //读取WorkBook子流数据
		size := binary.LittleEndian.Uint32(book.StreamSize[:]) //获取数据大小

		reader, err := adaptor.OpenObject(book, root) //获取了workbook子流的所有数据

		if err != nil {
			return workbook, err
		}

		return readStream(reader, size)

	}

	return workbook, err
}

func readStream(reader io.ReadSeeker, streamSize uint32) (workbook Workbook, err error) {

	stream := make([]byte, streamSize)

	_, err = reader.Read(stream) //读取数据到stream

	if err != nil {
		return workbook, nil
	}

	if err != nil {
		return workbook, nil
	}

	err = workbook.read(stream) //读取到workbook中去（初始化为空的一个结构体）

	if err != nil {
		return workbook, nil
	}

	for k := range workbook.sheets { //重要：存放了sheet name和对应数据的偏移量！！！！
		sheet, err := workbook.GetSheet(k) //获取一个sheet

		if err != nil {
			return workbook, nil
		}

		err = sheet.read(stream) //读取数据到sheet中（注意stream就是上面的完整的stream）

		if err != nil {
			return workbook, nil
		}
	}

	return
}
