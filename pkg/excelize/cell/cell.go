package main

import (
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"

	"github.com/xuri/excelize/v2"
)

func main() {
	srcFile := "test_files/test.xlsx"
	// 打开一个 Excel 文件
	opts := excelize.Options{}
	f, err := excelize.OpenFile(srcFile, opts)
	if err != nil {
		return
	}

	// 设置行高。三个参数分别为：Sheet 名，行号，高度
	err = f.SetRowHeight("Sheet1", 1, 45)
	if err != nil {
		panic(err)
	}

	// 设置列宽。四个参数分别为：Sheet 名，起始列号，结束列号，宽度
	err = f.SetColWidth("Sheet1", "A", "H", 5.57)
	if err != nil {
		panic(err)
	}

	// 插入图片
	// 设置图片格式
	format := `{
		"autofit": true,
		"lock_aspect_ratio": true
	}`
	err = f.AddPicture("Sheet1", "A1", "test_files/BT1-001R.png", format)
	if err != nil {
		panic(err)
	}

	f.Save()
}
