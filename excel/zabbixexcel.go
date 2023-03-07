package excel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"zabbixDemo01/zabbix"
)

// WriteExcel 将数据写入Excel
func WriteExcel(sheet string, das []zabbix.DayAveSlice, f *excelize.File) int {

	_ = f.SetCellValue(sheet, "A1", "日期")
	_ = f.SetCellValue(sheet, "B1", "上传")
	_ = f.SetCellValue(sheet, "C1", "下载")
	var n int
	for i := 0; i < len(das); i++ {
		n = i + 2
		_ = f.SetCellValue(sheet, fmt.Sprintf("%s%d", "A", n), das[i].Clock)
		//if err != nil {
		//	fmt.Println(err)
		//}
		//exp := "[$-380A]yyyy\"-\"m\"-\"d\"\";@"
		//style, err := f.NewStyle(&excelize.Style{CustomNumFmt: &exp})
		//if err != nil {
		//	panic(err)
		//}
		//_ = f.SetCellStyle(sheet, fmt.Sprintf("%s%d", "A", n), fmt.Sprintf("%s%d", "A", n), style)
		_ = f.SetCellFloat(sheet, fmt.Sprintf("%s%d", "B", n), (das)[i].UpAve, 3, 64)
		_ = f.SetCellFloat(sheet, fmt.Sprintf("%s%d", "C", n), (das)[i].DownAve, 3, 64)

		if i == len(das)-1 {
			_ = f.SetCellValue(sheet, fmt.Sprintf("%s%d", "A", n+1), "平均值")

			style, err := f.NewStyle(&excelize.Style{DecimalPlaces: 2})
			if err != nil {
				panic(err)
			}

			_ = f.SetCellFormula(sheet, fmt.Sprintf("%s%d", "B", n+1), fmt.Sprintf("=AVERAGE(B2:B%d)", n))
			//_ = f.SetCellStyle(sheet, fmt.Sprintf("%s%d", "C", n+1), fmt.Sprintf("%s%d", "C", n+1), style)
			_ = f.SetCellFormula(sheet, fmt.Sprintf("%s%d", "C", n+1), fmt.Sprintf("=AVERAGE(C2:C%d)", n))
			_ = f.SetCellStyle(sheet, fmt.Sprintf("%s%d", "B", n+1), fmt.Sprintf("%s%d", "C", n+1), style)

		}

	}
	return n
}

// AveTable 制作月平均值表格
func AveTable(f *excelize.File, mas []zabbix.MonthAveSlice) {
	err := f.SetCellValue("月报", "A63", "三月线路平均使用率")
	style3, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Size: 12,
			Bold: true,
		},
	})
	if err != nil {
		fmt.Println(err)
		return
	}
	_ = f.SetCellStyle("月报", "A63", "A63", style3)
	err = f.SetCellValue("月报", "A64", "运营商")
	if err != nil {
		return
	}
	err = f.SetCellValue("月报", "B64", "上传使用率")
	if err != nil {
		return
	}
	err = f.SetCellValue("月报", "C64", "下载使用率")
	if err != nil {
		return
	}
	tabelCellRanges := [][]string{{"A65", "B65", "C65"}, {"A66", "B66", "C66"}, {"A67", "B67", "C67"}, {"A68", "B68", "C68"}, {"A69", "B69", "C69"}}
	for i, ranges := range tabelCellRanges {
		err := f.SetCellFormula("月报", ranges[0], fmt.Sprintf("=A%d", i+4))
		if err != nil {
			return
		}
		err = f.SetCellFloat("月报", ranges[1], mas[i].Result.UpAve, 2, 64)
		if err != nil {
			return
		}
		err = f.SetCellFloat("月报", ranges[2], mas[i].Result.DownAve, 2, 64)
		if err != nil {
			return
		}
	}

	style, _ := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
		NumFmt: 10,
	})

	err = f.SetCellStyle("月报", "B65", "C69", style)
	if err != nil {
		return
	}

	err = f.AddTable("月报", "A64:C69", &excelize.TableOptions{
		Name:      "table",
		StyleName: "TableStyleMedium2",
	})

	err = f.SetColWidth("月报", "C", "C", 17)
	if err != nil {
		return
	}
}

// LineChart 制作折线图
func LineChart(f *excelize.File, n int) {
	// 上传
	if err := f.AddChart("月报", "A19", &excelize.Chart{
		Type: "line",
		Series: []excelize.ChartSeries{
			// 移动
			{
				Name:       "月报!$A$4",
				Categories: fmt.Sprintf("%s!$A$2:$A$%d", "Mobile", n),
				Values:     fmt.Sprintf("%s!$B$2:$B$%d", "Mobile", n),
				Line: excelize.ChartLine{
					Smooth: true,
					Width:  2,
				},
				Marker: excelize.ChartMarker{
					Symbol: "none",
				},
			},
			// 联通
			{
				Name:       "月报!$A$5",
				Categories: fmt.Sprintf("%s!$A$2:$A$%d", "Unicom", n),
				Values:     fmt.Sprintf("%s!$B$2:$B$%d", "Unicom", n),
				Line: excelize.ChartLine{
					Smooth: true,
					Width:  2,
				},
				Marker: excelize.ChartMarker{
					Symbol: "none",
				},
			},
			// 电信
			{
				Name:       "月报!$A$6",
				Categories: fmt.Sprintf("%s!$A$2:$A$%d", "Telecom", n),
				Values:     fmt.Sprintf("%s!$B$2:$B$%d", "Telecom", n),
				Line: excelize.ChartLine{
					Smooth: true,
					Width:  2,
				},
				Marker: excelize.ChartMarker{
					Symbol: "none",
				},
			},
			// MPLS
			{
				Name:       "月报!$A$7",
				Categories: fmt.Sprintf("%s!$A$2:$A$%d", "Mpls", n),
				Values:     fmt.Sprintf("%s!$B$2:$B$%d", "Mpls", n),
				Line: excelize.ChartLine{
					Smooth: true,
					Width:  2,
				},
				Marker: excelize.ChartMarker{
					Symbol: "none",
				},
			},
			// 专线
			{
				Name:       "月报!$A$8",
				Categories: fmt.Sprintf("%s!$A$2:$A$%d", "CDtoBJ", n),
				Values:     fmt.Sprintf("%s!$B$2:$B$%d", "CDtoBJ", n),
				Line: excelize.ChartLine{
					Smooth: true,
					Width:  2,
				},
				Marker: excelize.ChartMarker{
					Symbol: "none",
				},
			},
		},
		Format: excelize.GraphicOptions{
			OffsetX: 15,
			OffsetY: 10,
		},
		YAxis: excelize.ChartAxis{
			MajorGridLines: true,
		},
		Legend: excelize.ChartLegend{
			Position:      "top",
			ShowLegendKey: true,
		},
		Title: excelize.ChartTitle{
			Name: "上传",
		},
		PlotArea: excelize.ChartPlotArea{
			ShowCatName:     false,
			ShowLeaderLines: false,
			ShowPercent:     true,
			ShowSerName:     false,
			ShowVal:         false,
		},
		Dimension: excelize.ChartDimension{
			Height: 400,
			Width:  700,
		},
		ShowBlanksAs: "zero",
	}); err != nil {
		fmt.Println(err)
		return
	}

	// 下载
	if err := f.AddChart("月报", "A40", &excelize.Chart{
		Type: "line",
		Series: []excelize.ChartSeries{
			// 移动
			{
				Name:       "月报!$A$4",
				Categories: fmt.Sprintf("%s!$A$2:$A$%d", "Mobile", n),
				Values:     fmt.Sprintf("%s!$C$2:$C$%d", "Mobile", n),
				Line: excelize.ChartLine{
					Smooth: true,
					Width:  2,
				},
				Marker: excelize.ChartMarker{
					Symbol: "none",
				},
			},
			// 联通
			{
				Name:       "月报!$A$5",
				Categories: fmt.Sprintf("%s!$A$2:$A$%d", "Unicom", n),
				Values:     fmt.Sprintf("%s!$C$2:$C$%d", "Unicom", n),
				Line: excelize.ChartLine{
					Smooth: true,
					Width:  2,
				},
				Marker: excelize.ChartMarker{
					Symbol: "none",
				},
			},
			// 电信
			{
				Name:       "月报!$A$6",
				Categories: fmt.Sprintf("%s!$A$2:$A$%d", "Telecom", n),
				Values:     fmt.Sprintf("%s!$C$2:$C$%d", "Telecom", n),
				Line: excelize.ChartLine{
					Smooth: true,
					Width:  2,
				},
				Marker: excelize.ChartMarker{
					Symbol: "none",
				},
			},
			// MPLS
			{
				Name:       "月报!$A$7",
				Categories: fmt.Sprintf("%s!$A$2:$A$%d", "Mpls", n),
				Values:     fmt.Sprintf("%s!$C$2:$C$%d", "Mpls", n),
				Line: excelize.ChartLine{
					Smooth: true,
					Width:  2,
				},
				Marker: excelize.ChartMarker{
					Symbol: "none",
				},
			},
			// 专线
			{
				Name:       "月报!$A$8",
				Categories: fmt.Sprintf("%s!$A$2:$A$%d", "CDtoBJ", n),
				Values:     fmt.Sprintf("%s!$C$2:$C$%d", "CDtoBJ", n),
				Line: excelize.ChartLine{
					Smooth: true,
					Width:  2,
				},
				Marker: excelize.ChartMarker{
					Symbol: "none",
				},
			},
		},
		Format: excelize.GraphicOptions{
			OffsetX: 15,
			OffsetY: 10,
		},
		YAxis: excelize.ChartAxis{
			MajorGridLines: true,
		},
		Legend: excelize.ChartLegend{
			Position:      "top",
			ShowLegendKey: true,
		},
		Title: excelize.ChartTitle{
			Name: "下载",
		},
		PlotArea: excelize.ChartPlotArea{
			ShowCatName:     false,
			ShowLeaderLines: false,
			ShowPercent:     true,
			ShowSerName:     false,
			ShowVal:         false,
		},
		Dimension: excelize.ChartDimension{
			Height: 400,
			Width:  700,
		},
		ShowBlanksAs: "zero",
	}); err != nil {
		fmt.Println(err)
		return
	}
}

// CreateXlsx 创建xlsx文件
func CreateXlsx(d []zabbix.DayData) {
	// 删除同名的xlsx文件
	err := os.Remove("../zabbixDemo01Files/book1.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	// 创建一个xlsx文件
	f := excelize.NewFile()
	sheetName := "月报"
	_ = f.SetSheetName("Sheet1", sheetName)
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	data1 := [][]interface{}{
		{"三月月报"},
		{"成都公共互联网线路"},
		{"运营商", "带宽大小"},
		{"移动", "1000Mb"},
		{"联通", "250Mb"},
		{"电信", "100Mb"},
		{"MPLS", "80Mb"},
		{"成都-北京 专线", "50Mb"},
	}

	for i, row := range data1 {
		startCell, err := excelize.JoinCellName("A", i+1)
		if err != nil {
			fmt.Println(err)
			return
		}
		err = f.SetSheetRow(sheetName, startCell, &row)
		if err != nil {
			fmt.Println(err)
			return
		}
	}

	mergeCellRanges := [][]string{{"A1", "K1"}, {"A2", "F2"}}
	for _, ranges := range mergeCellRanges {
		if err := f.MergeCell(sheetName, ranges[0], ranges[1]); err != nil {
			fmt.Println(err)
			return
		}
	}

	style1, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "center"},
		Font:      &excelize.Font{Size: 20, Bold: true},
	})
	if err != nil {
		fmt.Println(err)
		return
	}
	err = f.SetCellStyle(sheetName, "A1", "A1", style1)
	if err != nil {
		fmt.Println(err)
		return
	}

	style2, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "center"},
	})
	if err != nil {
		fmt.Println(err)
		return
	}
	centerRowRanges := [][]string{{"A3", "A8"}, {"B3", "B8"}}
	for _, ranges := range centerRowRanges {
		err = f.SetCellStyle(sheetName, ranges[0], ranges[1], style2)
		if err != nil {
			fmt.Println(err)
			return
		}
	}

	style3, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Size: 12,
			Bold: true,
		},
	})
	if err != nil {
		fmt.Println(err)
		return
	}
	err = f.SetCellStyle(sheetName, "A2", "A2", style3)
	if err != nil {
		fmt.Println(err)
		return
	}

	err = f.SetColWidth(sheetName, "A", "B", 17)
	if err != nil {
		fmt.Println(err)
		return
	}

	err = f.AddTable(sheetName, "A3:B8", &excelize.TableOptions{
		Name:      "table1",
		StyleName: "TableStyleLight2",
	})
	if err != nil {
		fmt.Println(err)
		return
	}

	// 添加注释
	err = f.AddComment(sheetName, excelize.Comment{
		Cell:   "B8",
		Author: "备注",
		Runs: []excelize.RichTextRun{
			{Text: "备注:", Font: &excelize.Font{Bold: true}},
			{Text: "计划2023年5月降至10Mb"},
		},
	})

	data2 := [][]interface{}{
		{"线路功能"},
		{"移动:", "互联网访问、成都-北京(IPSec VPN)数据同步以及互访;"},
		{"联通:", "移动线路备份、集团没有MPLS线路的site(IPSec VPN)数据同步以及互访;"},
		{"电信:", "移动线路备份;"},
		{"MPLS:", "集团各site内网数据同步以及互访;"},
		{"成都-北京 专线:", "夏普项目."},
	}
	for i, row := range data2 {
		startCell, err := excelize.JoinCellName("A", i+10)
		if err != nil {
			fmt.Println(err)
			return
		}
		err = f.SetSheetRow(sheetName, startCell, &row)
		if err != nil {
			fmt.Println(err)
			return
		}
	}

	err = f.SetCellStyle(sheetName, "A10", "A10", style3)
	if err != nil {
		fmt.Println(err)
		return
	}

	err = f.SetCellValue(sheetName, "A17", "三月各线路平均流量图")
	if err != nil {
		fmt.Println(err)
		return
	}

	err = f.SetCellStyle(sheetName, "A17", "A17", style3)
	if err != nil {
		fmt.Println(err)
		return
	}

	// 新建多张表
	for i := 0; i < len(d); i++ {
		_, err := f.NewSheet((d)[i].Isp)
		if err != nil {
			fmt.Println(err)
		}
	}

	// 根据指定路径保存文件
	if err := f.SaveAs("../zabbixDemo01Files/book1.xlsx"); err != nil {
		fmt.Println(err)
	}
}
