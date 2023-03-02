package excel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"zabbixDemo01/zabbix"
)

// WriteExcel 将数据写入Excel
func WriteExcel(sheet string, das []zabbix.DayAveSlice, f *excelize.File) {

	for i := 0; i < len(das); i++ {

		f.SetCellValue(sheet, fmt.Sprintf("%s%d", "A", i+3), (das)[i].Clock)
		exp := "[$-380A]yyyy\"年\"m\"月\"d\"日\";@"
		style, err := f.NewStyle(&excelize.Style{CustomNumFmt: &exp})
		if err != nil {
			panic(err)
		}
		f.SetCellStyle(sheet, fmt.Sprintf("%s%d", "A", i+3), fmt.Sprintf("%s%d", "A", i+3), style)

		f.SetCellFloat(sheet, fmt.Sprintf("%s%d", "B", i+3), (das)[i].UpAve, 3, 64)
		f.SetCellFloat(sheet, fmt.Sprintf("%s%d", "C", i+3), (das)[i].DownAve, 3, 64)
	}
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

// AveChart 制作折线图
func AveChart(f *excelize.File) {
	// 上传
	if err := f.AddChart("sheet1", "A1", &excelize.Chart{
		Type: "line",
		Series: []excelize.ChartSeries{
			// 移动
			{
				Name:       "移动!$A$1",
				Categories: "移动!$A$3:$A$19",
				Values:     "移动!$B$3:$B$19",
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
				Name:       "联通!$A$1",
				Categories: "联通!$A$3:$A$19",
				Values:     "联通!$B$3:$B$19",
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
				Name:       "电信!$A$1",
				Categories: "电信!$A$3:$A$19",
				Values:     "电信!$B$3:$B$19",
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
				Name:       "MPLS!$A$1",
				Categories: "MPLS!$A$3:$A$19",
				Values:     "MPLS!$B$3:$B$19",
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
				Name:       "专线!$A$1",
				Categories: "专线!$A$3:$A$19",
				Values:     "专线!$B$3:$B$19",
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
	if err := f.AddChart("sheet1", "M1", &excelize.Chart{
		Type: "line",
		Series: []excelize.ChartSeries{
			// 移动
			{
				Name:       "移动!$A$1",
				Categories: "移动!$A$3:$A$19",
				Values:     "移动!$C$3:$C$19",
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
				Name:       "联通!$A$1",
				Categories: "联通!$A$3:$A$19",
				Values:     "联通!$C$3:$C$19",
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
				Name:       "电信!$A$1",
				Categories: "电信!$A$3:$A$19",
				Values:     "电信!$C$3:$C$19",
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
				Name:       "MPLS!$A$1",
				Categories: "MPLS!$A$3:$A$19",
				Values:     "MPLS!$C$3:$C$19",
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
				Name:       "专线!$A$1",
				Categories: "专线!$A$3:$A$19",
				Values:     "专线!$C$3:$C$19",
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
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// 新建多张表
	for i := 0; i < len(d); i++ {
		f.NewSheet((d)[i].Isp)
		f.SetCellValue((d)[i].Isp, "A1", (d)[i].Isp)
		f.MergeCell((d)[i].Isp, "A1", "C1")
		f.SetColWidth((d)[i].Isp, fmt.Sprintf("%s", "A"), fmt.Sprintf("%s", "A"), 15)
		f.SetCellValue((d)[i].Isp, "A2", "日期")
		f.SetCellValue((d)[i].Isp, "B2", "上传")
		f.SetCellValue((d)[i].Isp, "C2", "下载")
	}

	// 根据指定路径保存文件
	if err := f.SaveAs("../zabbixDemo01Files/book1.xlsx"); err != nil {
		fmt.Println(err)
	}
}
