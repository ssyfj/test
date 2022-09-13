package xls

import (
	"bytes"
	"errors"

	"github.com/shakinm/xlsReader/helpers"
	"github.com/shakinm/xlsReader/xls/record"
)

// Workbook struct
type Workbook struct {
	sheets   []Sheet
	codepage record.CodePage
	sst      record.SST
	xf       []record.XF
	formats  map[int]record.Format
	vers     [2]byte
}

// GetNumberSheets - Number of sheets in the workbook
func (wb *Workbook) GetNumberSheets() int {
	return len(wb.sheets)
}

// GetSheets - Get sheets in the workbook
func (wb *Workbook) GetSheets() []Sheet {
	return wb.sheets
}

// GetSheet - Get Sheet by ID
func (wb *Workbook) GetSheet(sheetID int) (sheet *Sheet, err error) { // nolint: golint

	if len(wb.sheets) >= 1 && len(wb.sheets) >= sheetID {
		return &wb.sheets[sheetID], err
	}

	return nil, errors.New("error. Sheet not found")
}

// GetXF -  Return Extended Format Record by index
func (wb *Workbook) GetXFbyIndex(index int) record.XF {
	if len(wb.xf)-1 < index {
		return wb.xf[15]
	}
	return wb.xf[index]
}

// GetXF -  Return FORMAT record describes a number format in the workbook
func (wb *Workbook) GetFormatByIndex(index int) record.Format {
	return wb.formats[index]
}

// GetCodePage - codepage
func (wb *Workbook) GetCodePage() record.CodePage {
	return wb.codepage
}

// GetVersionBIFF - version BIFF
func (wb *Workbook) GetVersionBIFF() []byte {
	return wb.vers[:]
}

func (wb *Workbook) addSheet(bs *record.BoundSheet) (sheet Sheet) { // nolint: golint
	sheet.boundSheet = bs
	sheet.wb = wb
	wb.sheets = append(wb.sheets, sheet)
	return sheet
}

func (wb *Workbook) read(stream []byte) (err error) { // nolint: gocyclo

	var point int32
	var SSTContinue = false
	var sPoint, prevLen int32
	var readType string
	var grbit byte
	var grbitOffset int32

	eof := false

Next:
	//好像循环读取4个字节（前两个，和后两个），其中前两个一般是对应了各个结构（比如BOF、EOF、BoundSheet...)，后两个对应了其结构长度信息！！！！！
	recordNumber := stream[point : point+2]
	recordDataLength := int32(helpers.BytesToUint16(stream[point+2 : point+4]))
	sPoint = point + 4 //跳转到后面数据部分

	//IndexRecord (20Bh)存放索引信息，包括第一行和最后一行的序号，以及默认的行单元数和每隔32行的数据单元数（？看完代码再解释）
	if bytes.Compare(recordNumber, record.IndexRecord[:]) == 0 {
		_ = new(record.LabelSSt) //跳过？？？我擦，干啥呢
		goto EIF
	}

	//BoundSheet 0x0085 即Sheet指针区，N个Sheet则有N个0x0085，包含每个Sheet的名称、sheet数据内容在xls文件中的偏移量。 主要就是存放了sheetname和数据内容偏移量
	if bytes.Compare(recordNumber, record.BoundSheetRecord[:]) == 0 { //Sheet指针区，N个Sheet则有N个0x0085，包含每个Sheet的名称、sheet数据内容在xls文件中的偏移量。
		var bs record.BoundSheet
		//stream[sPoint+grbitOffset:sPoint+recordDataLength]存放了包括sheet偏移量、文件类型标识和sheet name长度、以及最后的sheet name信息
		bs.Read(stream[sPoint+grbitOffset:sPoint+recordDataLength], wb.vers[:]) //开始wb是空的，BOF里面会为wb.vers赋值，用于存放XLS文件的版本类型BIFF8/BIFF8X-->0x0600 BIFF7-->0x0500
		_ = wb.addSheet(&bs)                                                    //加入了sheet
		goto EIF
	}

	//Continue，记录数据的延续（与下面的sst数据切分有关）
	if bytes.Compare(recordNumber, record.ContinueRecord[:]) == 0 {

		if SSTContinue {
			readType = "continue"

			if len(wb.sst.RgbSrc) == 0 {
				grbitOffset = 0
			} else {
				grbitOffset = 1
			}

			grbit = stream[sPoint]

			wb.sst.RgbSrc = append(wb.sst.RgbSrc, stream[sPoint+grbitOffset:sPoint+recordDataLength]...)
			wb.sst.Read(readType, grbit, prevLen)
		}
		goto EIF
	}

	//SST	SST内容（Sharing String Table 用来存放字符串，目的是为了让各个sheet都能够共享该SST中字符串内容）
	if bytes.Compare(recordNumber, record.SSTRecord[:]) == 0 { //SharingStringTable用来存放字符串，目的是为了让各个sheet都能够共享该SST中字符串内容）（Excel表数据，所有Sheet数据均放于此）
		wb.sst.NewSST(stream[sPoint : sPoint+recordDataLength]) //记录该段数据

		wb.sst.Read(readType, grbit, prevLen)
		totalSSt := helpers.BytesToUint32(wb.sst.CstTotal[:])
		if recordDataLength >= 8224 || uint32(len(wb.sst.Rgb)) < totalSSt-1 { //被切分了
			SSTContinue = true
		}
		goto EIF
	}

	if bytes.Compare(recordNumber, record.XFRecord[:]) == 0 { //extend format；也循环了很多次
		xf := new(record.XF) //font、format、ttype
		xf.Read(stream[sPoint : sPoint+recordDataLength])
		wb.xf = append(wb.xf, *xf)
		goto EIF
	}

	if bytes.Compare(recordNumber, record.FormatRecord[:]) == 0 { //循环了很多次，为啥？感觉挺重要的
		format := new(record.Format) //ifmt、cch、grbit、rgb、vers、stFormat

		format.Read(stream[sPoint:sPoint+recordDataLength], wb.vers[:]) //按照0x0600格式读取xls数据

		if wb.formats == nil {
			wb.formats = make(map[int]record.Format, 0)
		}
		wb.formats[format.GetIndex()] = *format
		goto EIF
	}

	//CodePage
	if bytes.Compare(recordNumber, record.CodePageRecord[:]) == 0 {
		wb.codepage.Read(stream[sPoint : sPoint+recordDataLength]) //wb.codepage.cv赋值，干啥用的？？
		goto EIF
	}

	//EOF
	if bytes.Compare(recordNumber, record.EOFRecord[:]) == 0 && recordDataLength == 0 { //直到读取到EOF才退出循环
		eof = true
	}

	//BOF
	if bytes.Compare(recordNumber, record.BOFMARKS[:]) == 0 { //var BOFMARKS = []byte{0x09, 0x08} //(809h)
		copy(wb.vers[:], stream[sPoint:sPoint+2]) //获取前2字节版本信息，BIFF8/BIFF8X 0x0600
		goto EIF
	}

EIF:
	point = point + recordDataLength + 4 //+4跳过原来的头部，recordNumber+recordDataLength部分
	if !eof {
		goto Next
	}

	return err
}
