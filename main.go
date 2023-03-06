package main

import (
	"time"
	"zabbixDemo01/zabbix"
)

func boolCust(b bool) *bool { return &b }

func main() {

	timeStart := time.Date(2023, 2, 1, 0, 0, 0, 0, time.UTC)
	timeEnd := time.Date(2023, 2, 28, 0, 0, 0, 0, time.UTC)
	count := zabbix.TimeCount(timeStart, timeEnd)

	a := []zabbix.IspItem{
		{"Mobile", 43998, 43993},
		{"Unicom", 43994, 43989},
		{"Telecom", 43995, 43990},
		{"Mpls", 43987, 43984},
		{"CDtoBJ", 43986, 43983},
	}

	data := zabbix.InitDayData(a, count)
	//fmt.Println(data)
	zabbix.MonthTrafficHandle(data)
	//// 生成xlsx文件
	//excel.CreateXlsx(data)
	//
	//// 打开一个文件
	//f, err := excelize.OpenFile("../zabbixDemo01Files/book1.xlsx")
	//if err != nil {
	//	fmt.Println(err)
	//	return
	//}
	//defer func() {
	//	if err := f.Close(); err != nil {
	//		fmt.Println(err)
	//	}
	//}()
	//
	//// 不显示网格线
	//if err = f.SetSheetView("月报", 0, &excelize.ViewOptions{ShowGridLines: boolCust(false)}); err != nil {
	//	fmt.Println(err)
	//	return
	//}
	//
	//// 写入数据
	//for i := 0; i < len(data); i++ {
	//	n := excel.WriteExcel((data)[i].Isp, (data)[i].AveResult, f)
	//	if i == len(data)-1 {
	//		excel.LineChart(f, n)
	//		//excel.PieChart(f, n)
	//	}
	//}
	//
	//if err := f.SaveAs("../zabbixDemo01Files/book1.xlsx"); err != nil {
	//	fmt.Println(err)
	//}
}
