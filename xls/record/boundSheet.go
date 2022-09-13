package record

import (
	"bytes"
	"strings"

	"github.com/shakinm/xlsReader/xls/structure"
)

// This record stores the sheet name, sheet type, and stream position
var BoundSheetRecord = [2]byte{0x85, 0x00} //(85h)

/*
Offset		 Field Name		 Size		 Contents
-------------------------------------------------
4			lbPlyPos		4			Stream position of the start of the BOF record for the sheet
8			grbit			2			Option flags
10			cch				1			Sheet name ( grbit / rgb fields of Unicode String)
11			rgch			var			Sheet name ( grbit / rgb fields of Unicode String)
*/

/*
The grbit structure contains the following options:（就两个字节）
Bits（16位，占了那些位） Mask：对应的值
Bits	Mask	Option Name		Contents
----------------------------------------
1–0 	0003h 	hsState 		Hidden state:	//隐藏状态
									00h = visible
									01h = hidden
									02h = very hidden (see text)
7–2 	00FCh 						(Reserved)
15–8	FF00h 	dt				Sheet type:
									00h = worksheet or dialog sheet 工作表或对话框
									01h = Excel 4.0 macro sheet Excel 4.0宏表
									02h = chart	图表
									06h = Visual Basic module Visual Basic模块
*/

type BoundSheet struct {
	LbPlyPos [4]byte                               //sheet数据内容在xls文件中的绝对偏移量
	Grbit    [2]byte                               //选项标志，如上所示
	Cch      [1]byte                               //工作表名称长度
	Rgch     []byte                                //工作表名称（Unicode字符串的grbit/rgb字段）对于非FlagBIFF8类型，在cch和rgch获取name信息
	stFormat structure.XLUnicodeRichExtendedString //对于FlagBIFF8，需要到这里面去获取sheetname
	vers     []byte                                //xls文件版本类型
}

func (r *BoundSheet) Read(stream []byte, vers []byte) {

	r.vers = vers

	copy(r.LbPlyPos[:], stream[0:4]) //获取偏移量
	copy(r.Grbit[:], stream[4:6])    //获取选项标志
	copy(r.Cch[:], stream[6:7])      //sheetname长度信息

	if bytes.Compare(vers, FlagBIFF8) == 0 { //对比xls文件版本类型，一般都是BIFF8。

		fixedStream := []byte{r.Cch[0], 0x00}            //包含了sheetname长度
		fixedStream = append(fixedStream, stream[7:]...) //追加了sheet name信息
		_ = r.stFormat.Read(fixedStream)

	} else { //对于非BIFF8类型，直接获取name信息
		r.Rgch = make([]byte, int(r.Cch[0]))
		copy(r.Rgch[:], stream[7:]) //获取名称信息
	}
}
func (r *BoundSheet) GetName() string {
	if bytes.Compare(r.vers, FlagBIFF8) == 0 { //可以知道sheet name存在两个方式
		return r.stFormat.String()
	}
	strLen := int(r.Cch[0])
	return strings.TrimSpace(string(decodeWindows1251(bytes.Trim(r.Rgch[:int(strLen)], "\x00"))))
}
