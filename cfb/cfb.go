package cfb

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"io"
	"os"
	"path/filepath"

	"github.com/shakinm/xlsReader/helpers"
)

// Cfb - Compound File Binary
type Cfb struct {
	header           Header
	file             io.ReadSeeker
	difatPositions   []uint32
	miniFatPositions []uint32
	dirs             []*Directory
}

// EntrySize - Directory array entry length
var EntrySize = 128

// DefaultDIFATEntries -Number FAT locations in DIFAT
var DefaultDIFATEntries = uint32(109)

// GetDirs - Get a list of directories
func (cfb *Cfb) GetDirs() []*Directory {
	return cfb.dirs
}

// OpenFile - Open document from the file
func OpenFile(filename string) (cfb Cfb, err error) {
	fmt.Println(filepath.Clean(filename))
	cfb.file, err = os.Open(filepath.Clean(filename)) //打开document文档，放入cfb.file中

	if err != nil {
		return cfb, err
	}

	err = open(&cfb) //主要获取了difatPostions和dirs扇区数据

	return cfb, err
}

// OpenReader - Open document from the reader
func OpenReader(reader io.ReadSeeker) (cfb Cfb, err error) {

	cfb.file = reader

	if err != nil {
		return
	}

	err = open(&cfb)

	return
}

func open(cfb *Cfb) (err error) {

	err = cfb.getHeader() //获取4096字节，解析复合文档头部信息

	if err != nil {
		return err
	}

	err = cfb.getMiniFATSectors() //获取mini fat 扇区大小，但是里面啥也没干！！！

	if err != nil {
		return err
	}

	err = cfb.getFatSectors() //获取所有扇区数据转int32到cfb.difatPositions中去

	if err != nil {
		return err
	}

	err = cfb.getDirectories() //读取目录流扇区数据

	return err
}

func (cfb *Cfb) getHeader() (err error) {

	var bHeader = make([]byte, 4096)

	_, err = cfb.file.Read(bHeader) //从文件中读取4096字节，放入bHeader。不应该是4096啊，为什么呢？？？
	// n, err := cfb.file.Read(bHeader) //从文件中读取4096字节，放入bHeader。不应该是4096啊，为什么呢？？？

	// fmt.Println(n)
	if err != nil {
		return
	}

	err = binary.Read(bytes.NewBuffer(bHeader), binary.LittleEndian, &cfb.header) //二进制读取，转小端，存放cfg.header中去
	// for idx, difat := range cfb.header.DIFAT {
	// 	fmt.Println(idx, difat)
	// }
	if err != nil {
		return
	}

	err = cfb.header.validate() //header字段的简单验证

	return
}

func (cfb *Cfb) getDirectories() (err error) {

	stream, err := cfb.getDataFromFatChain(helpers.BytesToUint32(cfb.header.FirstDirectorySectorLocation[:])) //文件目录流的起始扇区编号，读取下面所有的文件扇区链（直到链末端）并返回

	if err != nil {
		return err
	}
	var section = make([]byte, 0)

	for _, value := range stream { //读取扇区数据
		section = append(section, value)
		if len(section) == EntrySize {
			var dir Directory                                                      //是一个记录了复合文档中的所有内容的目录，红黑树
			err = binary.Read(bytes.NewBuffer(section), binary.LittleEndian, &dir) //复合文档目录项读取
			if err == nil && dir.ObjectType != 0x00 {
				cfb.dirs = append(cfb.dirs, &dir) //添加到cfb.dirs中去
			}

			section = make([]byte, 0)
		}

	}

	return

}

func (cfb *Cfb) getMiniFATSectors() (err error) {

	var section = make([]byte, 0)

	position := cfb.calculateOffset(cfb.header.FirstMiniFATSectorLocation[:]) //先找到位置，mini FAT的起始扇区编号

	for i := uint32(0); i < helpers.BytesToUint32(cfb.header.NumberMiniFATSectors[:]); i++ { //mini FAT扇区的数量，进行遍历，但是这里是0
		sector := NewSector(&cfb.header)
		err := cfb.getData(position, &sector.Data)

		if err != nil {
			return err
		}

		for _, value := range sector.getMiniFatFATSectorLocations() {
			section = append(section, value)
			if len(section) == 4 {
				cfb.miniFatPositions = append(cfb.miniFatPositions, helpers.BytesToUint32(section))
				section = make([]byte, 0)
			}
		}
		position = position + sector.SectorSize
	}

	return
}

func (cfb *Cfb) getFatSectors() (err error) { // nolint: gocyclo其实这里的cfb.difatPositions就是对应的全局FAT表，可能名字写的有歧义而已

	entries := DefaultDIFATEntries //默认复合文档是109个DIFAT

	if helpers.BytesToUint32(cfb.header.NumberFATSectors[:]) < DefaultDIFATEntries { //NumberFATSectors原本为1，FAT扇区的数量
		entries = helpers.BytesToUint32(cfb.header.NumberFATSectors[:])
	}

	for i := uint32(0); i < entries; i++ { //循环读取扇区信息

		position := cfb.calculateOffset(cfb.header.getDIFATEntry(i)) //getDIFATEntry返回cfb.header.DIFAT中对应扇区数据，[i*4:(i*4)+4],每个DIFAT占用4个字节
		sector := NewSector(&cfb.header)                             //版本3，扇区大小为512，空间分配，未填充
		//偏移量是指对应扇区的数据在文件数据中的偏移位置
		err := cfb.getData(position, &sector.Data) //读取一个扇区数据到新的sector扇区中去

		if err != nil {
			return err
		}
		//将512字节的扇区数据，转为128个int32的数组数据，存放在cfb.difatPositions中去
		cfb.difatPositions = append(cfb.difatPositions, sector.values(EntrySize)...) //EntrySize目录数组条目长度，128

	}

	if bytes.Compare(cfb.header.FirstDIFATSectorLocation[:], ENDOFCHAIN) == 0 { //DIFAT的起始扇区编号
		return
	}

	position := cfb.calculateOffset(cfb.header.FirstDIFATSectorLocation[:])
	var section = make([]byte, 0)
	for i := uint32(0); i < helpers.BytesToUint32(cfb.header.NumberDIFATSectors[:]); i++ {
		sector := NewSector(&cfb.header)
		err := cfb.getData(position, &sector.Data)

		if err != nil {
			return err
		}

		for _, value := range sector.getFATSectorLocations() {
			section = append(section, value)
			if len(section) == 4 {

				position = cfb.calculateOffset(section)
				sectorF := NewSector(&cfb.header)
				err := cfb.getData(position, &sectorF.Data)

				if err != nil {
					return err
				}
				cfb.difatPositions = append(cfb.difatPositions, sectorF.values(EntrySize)...)

				section = make([]byte, 0)
			}

		}

		position = cfb.calculateOffset(sector.getNextDIFATSectorLocation())

	}

	return
}
func (cfb *Cfb) getDataFromMiniFat(miniFatSectorLocation uint32, offset uint32) (data []byte, err error) {

	sPoint := cfb.sectorOffset(miniFatSectorLocation)    //根据root entry获取数据的偏移位置------------------这一步为什么是在root entry
	point := sPoint + cfb.calculateMiniFatOffset(offset) //根据book的偏移位置，找到对应的minifat偏移地址

	for {

		sector := NewMiniFatSector(&cfb.header)

		err = cfb.getData(point, &sector.Data)

		if err != nil {
			return data, err
		}

		data = append(data, sector.Data...)

		if cfb.miniFatPositions[offset] == helpers.BytesToUint32(ENDOFCHAIN) {
			break
		}

		offset = cfb.miniFatPositions[offset] //更新偏移位置

		point = sPoint + cfb.calculateMiniFatOffset(offset)

	}

	return data, err
}

func (cfb *Cfb) getDataFromFatChain(offset uint32) (data []byte, err error) {

	for {
		sector := NewSector(&cfb.header)
		point := cfb.sectorOffset(offset) //offset就是文件目录流的起始扇区编号：47；（47+1）*512；+1跳过了目录流吧？？？

		err = cfb.getData(point, &sector.Data) //读取扇区数据

		if err != nil {
			return data, err
		}

		data = append(data, sector.Data...)
		offset = cfb.difatPositions[offset]              //找到对应47索引数据，4294967294偏移
		if offset == helpers.BytesToUint32(ENDOFCHAIN) { //连接的扇区链的末端
			break
		}
	}

	return data, err
}

// OpenObject - Get object stream
func (cfb *Cfb) OpenObject(object *Directory, root *Directory) (reader io.ReadSeeker, err error) {

	if helpers.BytesToUint32(object.StreamSize[:]) < uint32(helpers.BytesToUint16(cfb.header.MiniStreamCutoffSize[:])) {

		data, err := cfb.getDataFromMiniFat(root.GetStartingSectorLocation(), object.GetStartingSectorLocation()) //minifat位置，取决于root entry和work book--->mini fat表两者

		if err != nil {
			return reader, err
		}

		reader = bytes.NewReader(data)
	} else {

		data, err := cfb.getDataFromFatChain(object.GetStartingSectorLocation()) //读取WorkBook子流的所有数据，只取决于work book--->fat表

		if err != nil {
			return reader, err
		}

		reader = bytes.NewReader(data)

	}

	return reader, err
}

func (cfb *Cfb) getData(offset uint32, data *[]byte) (err error) {

	_, err = cfb.file.Seek(int64(offset), 0) //原来偏移量是指文件数据的偏移啊

	if err != nil {
		return
	}

	_, err = cfb.file.Read(*data)

	if err != nil {
		return
	}
	return

}

func (cfb *Cfb) sectorOffset(sid uint32) uint32 {
	return (sid + 1) * cfb.header.sectorSize() //+1是表示跳过文件头（文件头占512，一个sector大小）
}

func (cfb *Cfb) calculateMiniFatOffset(sid uint32) (n uint32) {

	return sid * 64
}

func (cfb *Cfb) calculateOffset(sectorID []byte) (n uint32) { //传入cfg.header.DIFAT扇区对应的4字节数据

	if len(sectorID) == 4 {
		n = helpers.BytesToUint32(sectorID) //转32int类型，46
	}
	if len(sectorID) == 2 {
		n = uint32(binary.LittleEndian.Uint16(sectorID))
	}
	return (n * cfb.header.sectorSize()) + cfb.header.sectorSize() //版本3，sectorSize对应512，返回46*512+512
}
