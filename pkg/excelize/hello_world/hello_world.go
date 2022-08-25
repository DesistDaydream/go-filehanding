package main

import (
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"log"

	"github.com/xuri/excelize/v2"
)

func main() {
	// 实例化一个 excelize.File，这是一个电子表格文件结构。
	// 也可以说 excelize.File 的实例就是一个 Excel 文件
	f := excelize.NewFile()
	// 创建一个新的 Sheet 页，即创建一个新的 Workbook(工作簿)
	index := f.NewSheet("Sheet2")
	// 设置 Cell(单元格) 的值
	f.SetCellValue("Sheet2", "A2", "Hello world.")
	f.SetCellValue("Sheet1", "B2", 100)
	// 将指定的 Sheet 页设为活跃状态，即打开文档后，首先看到的 Sheet 页
	f.SetActiveSheet(index)
	// 另存为到指定文件
	if err := f.SaveAs("test_files/test.xlsx"); err != nil {
		log.Fatalln(err)
	}
}
