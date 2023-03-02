package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"time"
	"zabbixDemo01/excel"
	"zabbixDemo01/zabbix"
)

func main() {

	timeStart := time.Date(2023, 2, 1, 0, 0, 0, 0, time.UTC)
	timeEnd := time.Date(2023, 2, 23, 0, 0, 0, 0, time.UTC)
	count := zabbix.TimeCount(timeStart, timeEnd)

	a := []zabbix.IspItem{
		{"Mobile", 43998, 43993},
		{"Unicom", 43994, 43989},
		{"Telecom", 43995, 43990},
		{"Mpls", 43987, 43984},
		{"CDtoBJ", 43986, 43983},
	}

	data := zabbix.InitDayData(a, count)
	fmt.Println(data)
	// 生成xlsx文件
	excel.CreateXlsx(data)

	// 打开一个文件
	f, err := excelize.OpenFile("../zabbixDemo01Files/book1.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// 写入数据
	for i := 0; i < len(data); i++ {
		excel.WriteExcel((data)[i].Isp, (data)[i].AveResult, f)
	}

	// 生成线路月流量平均值表格

	if err := f.SaveAs("../zabbixDemo01Files/book1.xlsx"); err != nil {
		fmt.Println(err)
	}
}
