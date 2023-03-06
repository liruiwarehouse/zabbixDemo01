package zabbix

import (
	"bytes"
	"encoding/json"
	"log"
	"net/http"
	"strconv"
	"time"
)

// IspItem 运营商上传下载对应的itemid
type IspItem struct {
	Isp      string
	UpItem   int
	DownItem int
}

// BodyRaw history.Get 访问zabbix api时携带的JSON数据结构体
type BodyRaw struct {
	Jsonrpc string        `json:"jsonrpc"`
	Method  string        `json:"method"`
	Params  BodyRawParams `json:"params"`
	Auth    string        `json:"auth"`
	ID      int           `json:"id"`
}

// BodyRawParams 是 BodyRaw 结构体 params 选项的值 的结构体
type BodyRawParams struct {
	Output    string `json:"output"`
	History   int    `json:"history"`
	Itemids   int    `json:"itemids"`
	TimeFrom  int64  `json:"time_from"`
	TimeTill  int64  `json:"time_till"`
	Sortfield string `json:"sortfield"`
}

// ZabbixData JSON数据存储
type ZabbixData struct {
	Jsonrpc string   `json:"jsonrpc"`
	Result  []Result `json:"result"`
	ID      int      `json:"id"`
}

// Result zabbixData result数据
type Result struct {
	Itemid string `json:"itemid"`
	Clock  string `json:"clock"`
	Value  string `json:"value"`
	Ns     string `json:"ns"`
}

// TimeSlice 时间的结构体
type TimeSlice struct {
	Am, Pm time.Time
}

// DayAveSlice 保存当月上班时间以及当天流量平均值
type DayAveSlice struct {
	Clock          string
	UpAve, DownAve float64
}

// MonthAveSlice 保存月流量平均值
type MonthAveSlice struct {
	Isp    string
	Result DayAveSlice
}

// TrafficInterface 接口
type TrafficInterface interface {
	TrafficUpload(am, pm int64) ZabbixData
	TrafficDownload(am, pm int64) ZabbixData
}

// NewIspItem 构造函数
func NewIspItem(isp string, up, down int) *IspItem {
	return &IspItem{
		Isp:      isp,
		UpItem:   up,
		DownItem: down,
	}
}

// TrafficUpload 是 IspItem 结构体的方法
func (i *IspItem) TrafficUpload(am, pm int64) ZabbixData {
	return ZabbixGet(i.UpItem, am, pm)
}

// TrafficDownload 是 IspItem 结构体的方法
func (i *IspItem) TrafficDownload(am, pm int64) ZabbixData {
	return ZabbixGet(i.DownItem, am, pm)
}

// TimeCount 计算时间
func TimeCount(timeStart, timeEnd time.Time) []TimeSlice {
	var timeslice []TimeSlice
	for i := 0; i < timeEnd.Day(); i++ {
		if timeStart.AddDate(0, 0, i).Weekday() != time.Sunday && timeStart.AddDate(0, 0, i).Weekday() != time.Saturday {
			am := timeStart.AddDate(0, 0, i).Add(time.Hour * 9).Add(time.Minute * 30)
			pm := timeStart.AddDate(0, 0, i).Add(time.Hour * 18).Add(time.Minute * 30)
			m := TimeSlice{Am: am, Pm: pm}
			timeslice = append(timeslice, m)
		}
	}
	return timeslice
}

// ZabbixGet 通过history.get查询历史流量
func ZabbixGet(item int, am, pm int64) ZabbixData {
	// Body Raw Json 结构体实例化
	var brj = &BodyRaw{
		Jsonrpc: "2.0",
		Method:  "history.get",
		Params: BodyRawParams{
			Output:    "extend",
			History:   3,
			Itemids:   item,
			TimeFrom:  am,
			TimeTill:  pm,
			Sortfield: "clock",
		},
		Auth: "945d969a4147f120eb72235adb65a3f2",
		ID:   1,
	}

	// 将结构体实例转换为json
	jm, err3 := json.Marshal(brj)
	if err3 != nil {
		log.Fatal(err3)
	}
	// 将转换后的json数据转换成byte类型
	//b := []byte(jm)

	// zabbix api查询
	reader := bytes.NewReader(jm)
	req, err2 := http.NewRequest("GET", "http://192.168.87.45/api_jsonrpc.php", reader)
	if err2 != nil {
		log.Fatal(err2)
	}
	req.Header.Set("Content-Type", "application/json-rpc")

	client := &http.Client{}
	resp, err := client.Do(req)
	if err != nil {
		log.Fatal(err)
	}

	defer resp.Body.Close()

	// 将json数据转化为结构体
	var info ZabbixData
	json.NewDecoder(resp.Body).Decode(&info)

	return info
}

// MonthQuery 根据传入的时间切片生成天流量平均值并保存至DayAveSlice类型切片中并返回
func MonthQuery(t TrafficInterface, timeslice []TimeSlice) []DayAveSlice {
	// 初始化一个DayAveSlice类型切片
	var dayaveslice []DayAveSlice

	for i := 0; i < len(timeslice); i++ {
		// 这里(*timeslice)[i].Am如果写成*timeslice[i].Am会报错"Invalid operation: 'timeslice[i]' (type '*[]TimeSlice' does not support indexing)"
		// 因为Go会把*timeslice[i].Am当作(*timeslice[i].Am)
		upload := t.TrafficUpload((timeslice)[i].Am.Unix(), (timeslice)[i].Pm.Unix())
		download := t.TrafficDownload((timeslice)[i].Am.Unix(), (timeslice)[i].Pm.Unix())

		u := DayTrafficHandle(upload)
		d := DayTrafficHandle(download)

		ck := timeslice[i].Am.Format("01-02")
		c := DayAveSlice{Clock: ck, UpAve: u, DownAve: d}
		dayaveslice = append(dayaveslice, c)
	}
	return dayaveslice
}

// DayTrafficHandle 天流量平均值数据处理
func DayTrafficHandle(z ZabbixData) float64 {
	var num float64 = 0
	// 判断查询结果中result选项如果没有值，则ZabbixData结构体中的[]Result结构体切片的长度为0，
	// 返回这个0值，不然会显示为"NaN"，影响后面的计算
	if len(z.Result) == 0 {
		return 0
	}
	for i := 0; i < len(z.Result); i++ {
		p, err := strconv.ParseFloat(z.Result[i].Value, 32)
		if err != nil {
			log.Fatal(err)
		}
		num += p
	}
	return num / float64(len(z.Result)) / 1024 / 1024
}

// MonthTrafficHandle 月流量平均值数据处理
func MonthTrafficHandle(d []DayData) []MonthAveSlice {
	// 初始化一个MonthAveSlice类型切片
	var MontHave []MonthAveSlice

	//v := make([]interface{}, 3, 3)
	//MonthUsage := make([]interface{}, 5, 5)

	for i := 0; i < len(d); i++ {
		var u float64 = 0
		var n float64 = 0
		for j := 0; j < len(d[i].AveResult); j++ {
			u += d[i].AveResult[j].UpAve
			n += d[i].AveResult[j].DownAve
		}
		//s1 := fmt.Sprintf("%v: 上传: %.2f 下载: %.2f", d[i].Isp, u/float64(len(d[i].AveResult)), n/float64(len(d[i].AveResult)))
		//fmt.Println(s1)

		switch d[i].Isp {
		case "Mobile":
			v := MonthAveSlice{Isp: d[i].Isp, Result: DayAveSlice{UpAve: u / float64(len(d[i].AveResult)) / 1000, DownAve: n / float64(len(d[i].AveResult)) / 1000}}
			MontHave = append(MontHave, v)
		case "Unicom":
			v := MonthAveSlice{Isp: d[i].Isp, Result: DayAveSlice{UpAve: u / float64(len(d[i].AveResult)) / 250, DownAve: n / float64(len(d[i].AveResult)) / 250}}
			MontHave = append(MontHave, v)
		case "Telecom":
			v := MonthAveSlice{Isp: d[i].Isp, Result: DayAveSlice{UpAve: u / float64(len(d[i].AveResult)) / 100, DownAve: n / float64(len(d[i].AveResult)) / 100}}
			MontHave = append(MontHave, v)
		case "Mpls":
			v := MonthAveSlice{Isp: d[i].Isp, Result: DayAveSlice{UpAve: u / float64(len(d[i].AveResult)) / 80, DownAve: n / float64(len(d[i].AveResult)) / 80}}
			MontHave = append(MontHave, v)
		case "CDtoBJ":
			v := MonthAveSlice{Isp: d[i].Isp, Result: DayAveSlice{UpAve: u / float64(len(d[i].AveResult)) / 50, DownAve: n / float64(len(d[i].AveResult)) / 50}}
			MontHave = append(MontHave, v)
		}

	}

	return MontHave

}

type DayData struct {
	Isp       string
	AveResult []DayAveSlice
}

func InitDayData(item []IspItem, ts []TimeSlice) []DayData {
	var b []DayData
	for i := 0; i < len(item); i++ {
		it := NewIspItem(item[i].Isp, item[i].UpItem, item[i].DownItem)
		res := MonthQuery(it, ts)
		bb := DayData{Isp: item[i].Isp, AveResult: res}
		b = append(b, bb)
	}
	return b
}
