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
func AveTable(f *excelize.File) {
	style, _ := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#FFE4E1"},
			Pattern: 1,
		},
	})
	f.SetCellStyle("sheet1", "B24", "C24", style)
	f.SetCellStyle("sheet1", "A25", "A29", style)

	f.SetCellFormula("sheet1", "B25", "=AVERAGE(移动!B3:B19)")
	f.SetCellFormula("sheet1", "B26", "=AVERAGE(联通!B3:B19)")
	f.SetCellFormula("sheet1", "B27", "=AVERAGE(电信!B3:B19)")
	f.SetCellFormula("sheet1", "B28", "=AVERAGE(MPLS!B3:B19)")
	f.SetCellFormula("sheet1", "B29", "=AVERAGE(专线!B3:B19)")
	f.SetCellFormula("sheet1", "C25", "=AVERAGE(移动!C3:C19)")
	f.SetCellFormula("sheet1", "C26", "=AVERAGE(联通!C3:C19)")
	f.SetCellFormula("sheet1", "C27", "=AVERAGE(电信!C3:C19)")
	f.SetCellFormula("sheet1", "C28", "=AVERAGE(MPLS!C3:C19)")
	f.SetCellFormula("sheet1", "C29", "=AVERAGE(专线!C3:C19)")
}

// LineChart 制作折线图
func LineChart(f *excelize.File, n int) {
	// 上传
	if err := f.AddChart("月报", "A1", &excelize.Chart{
		Type: "line",
		Series: []excelize.ChartSeries{
			// 移动
			{
				Name:       "Mobile",
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
				Name:       "Unicom",
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
				Name:       "Telecom",
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
				Name:       "Mpls",
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
				Name:       "CDtoBJ",
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
	if err := f.AddChart("月报", "M1", &excelize.Chart{
		Type: "line",
		Series: []excelize.ChartSeries{
			// 移动
			{
				Name:       "Mobile",
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
				Name:       "Unicom",
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
				Name:       "Telecom",
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
				Name:       "Mpls",
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
				Name:       "CDtoBJ",
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

// PieChart 饼图
func PieChart(f *excelize.File, n int) {
	if err := f.AddChart("月报", "A23", &excelize.Chart{
		Type: "pie",
		Series: []excelize.ChartSeries{
			{
				Name: "数量",
				//Categories: "Sheet1!$A$1:$C$1",
				Values: fmt.Sprintf("%s!$B$%d", "Mobile", n+1),
			},
		},
		Format: excelize.GraphicOptions{
			OffsetX: 15,
			OffsetY: 10,
		},
		Title: excelize.ChartTitle{
			Name: "三维饼图",
		},
		PlotArea: excelize.ChartPlotArea{
			ShowPercent:     true,
			ShowCatName:     true,
			ShowLeaderLines: true,
		},
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
