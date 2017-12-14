package main

import (
	"net/http"
	"io/ioutil"
	"fmt"
	"encoding/json"
	"net/url"
	"bytes"
	"github.com/tealeg/xlsx"
	"strconv"
)

type resulta struct {
	Id        int    `json:"id"`
	Name      string `json:"name"`
	Ordertime string `json:"ordertime"`
}
type gamelist struct {
	Ret    int       `json:"ret"`
	Msg    string    `json:"msg"`
	Result []resulta `json:"result"`
}

type details struct {
	Msg    string        `json:"msg"`
	Result detailsResult `json:"result"`
}

type detailsResult struct {
	Rows []detailsRows `json:"rows"`
}

type detailsRows struct {
	ARPUZ     string `json:"ARPU"`
	Channelid int    `json:"channelid"`
	Date      int    `json:"date"`
	FeeNum    int    `json:"feeNum"`
	Gameappid int    `json:"gameappid"`
	LoginNum  int    `json:"loginNum"`
	PayARPU   string `json:"payARPU"`
	PayNum    int    `json:"payNum"`
	PayRate   int    `json:"payRate"`
	RegNum    int    `json:"regNum"`
	Res3Rate  string `json:"res3Rate"`
	Res7Rate  string `json:"res7Rate"`
	ResRate   string `json:"resRate"`
	Spstatus  string `json:"spstatus"`
	Tips      string `json:"tips"`
}

type gamenamelist struct {
	Msg    string   `json:"msg"`
	Result []gameid `json:"result"`
	Ret    int      `json:"ret"`
}
type gameid struct {
	Id        int    `json:"id"`
	Name      string `json:"name"`
	Ordertime string `json:"ordertime"`
}

func main() {
	client := &http.Client{}
	//声明游戏列表的结构体
	//gamelist := gamelist{}
	gamenamelist := gamenamelist{}
	dataCentre := details{}
	gamemap := make(map[int] string)

	requse, _ := http.NewRequest("POST", "http://s.qq.com/service/datacenter/gamelist", nil)
	requse.Header.Set("Cookie", "RK=zH9eob16GS; _qpsvr_localtk=0.21118394657969475; pgv_pvi=1420285952; pgv_si=s580742144; ptui_loginuin=2880997838; ptisp=ctc; ptcz=30176a2feaae43b130ff1d3266e0537af749354108426df819640dfa5a95448f; uin=o2880997838; skey=@UmeQnuT8W; pt2gguin=o2880997838; isuser=1; key=2880997838-9b251f70-dfdc-11e7-964e-d39c5844c4b4; iCPID=488; isregist=1; isaudit=1; issign=1; ilevel=0; iHead=http%3A//thirdqq.qlogo.cn/g%3Fb%3Dsdk%26k%3DfgFdOq0Tzuwt6micsbQF7bA%26s%3D140%26t%3D1493197858")
	response, _ := client.Do(requse)
	if response.StatusCode == 200 {
		//获取游戏名称
		body, _ := ioutil.ReadAll(response.Body)
		bodystr := string(body)
		data := []byte(bodystr)
		json.Unmarshal(data, &gamenamelist)
		for _,v:=range gamenamelist.Result {
			gamemap[v.Id]=v.Name
		}

		//获取数据中心列表
		//
		postValues := url.Values{}
		postValues.Set("tStartTime", "2017-12-07")
		postValues.Set("tEndTime", "2017-12-13")
		postValues.Set("vChannelId", `["10024328"]`)
		postValues.Set("vGameId", `["1106520","11056012372","1106453097521"]`)
		postValues.Set("iPageNo", `0`)
		postValues.Set("iPageSize", `500`)
		postValues.Set("iType", `0`)
		postdatastr := postValues.Encode()
		postDataBytes := []byte(postdatastr)
		postBytesReader := bytes.NewReader(postDataBytes)

		dataCenter, _ := http.NewRequest("POST", "http://s.qq.com/service/datacenter/details", postBytesReader)
		dataCenter.Header.Set("Cookie", "RK=zH9eob16GS; _qpsvr_localtk=0.21118394657969475; pgv_pvi=1420285952; pgv_si=s580742144; ptui_loginuin=2880997838; ptisp=ctc; ptcz=30176a2feaae43b130ff1d3266e0537af749354108426df819640dfa5a95448f; uin=o2880997838; skey=@UmeQnuT8W; pt2gguin=o2880997838; isuser=1; key=2880997838-9b251f70-dfdc-11e7-964e-d39c5844c4b4; iCPID=488; isregist=1; isaudit=1; issign=1; ilevel=0; iHead=http%3A//thirdqq.qlogo.cn/g%3Fb%3Dsdk%26k%3DfgFdOq0Tzuwt6micsbQF7bA%26s%3D140%26t%3D1493197858")
		dataCenter.Header.Add("Content-Type", "application/x-www-form-urlencoded")
		dataCenterResponse, _ := client.Do(dataCenter)
		if dataCenterResponse.StatusCode == 200 {
			bodyDetails, _ := ioutil.ReadAll(dataCenterResponse.Body)
			bodyDetailsStr := string(bodyDetails)
			bbs := []byte(bodyDetailsStr)
			json.Unmarshal(bbs, &dataCentre)
			dd := dataCentre.Result.Rows

			file := xlsx.NewFile()
			sheet, _ := file.AddSheet("Sheet1")
			//设置第一行的名字
			row := sheet.AddRow()
			row.SetHeightCM(1) //设置每行的高度
			cell := row.AddCell()
			cell.Value = "日期"
			cell = row.AddCell()
			cell.Value = "游戏名称"
			cell = row.AddCell()
			cell.Value = "渠道号"
			cell = row.AddCell()
			cell.Value = "新增用户"
			cell = row.AddCell()
			cell.Value = "活跃用户"
			cell = row.AddCell()
			cell.Value = "付费用户"
			cell = row.AddCell()
			cell.Value = "运营流水（元）"
			cell = row.AddCell()
			cell.Value = "付费ARPU"
			cell = row.AddCell()
			cell.Value = "活跃ARPU"
			cell = row.AddCell()
			cell.Value = "付费率"
			cell = row.AddCell()
			cell.Value = "次日留存率"
			cell = row.AddCell()
			cell.Value = "三日留存率"
			cell = row.AddCell()
			cell.Value = "七日留存率"

			for _, v := range dd {
				row := sheet.AddRow()
				row.SetHeightCM(1) //设置每行的高度
				cell := row.AddCell()
				cell.Value = strconv.Itoa(v.Date)
				cell = row.AddCell()
				cell.Value = gamemap[v.Gameappid]
				cell = row.AddCell()
				cell.Value = strconv.Itoa(v.Channelid)
				cell = row.AddCell()
				cell.Value = strconv.Itoa(v.RegNum)
				cell = row.AddCell()
				cell.Value = strconv.Itoa(v.LoginNum)
				cell = row.AddCell()
				cell.Value = strconv.Itoa(v.PayNum)
				cell = row.AddCell()
				cell.Value = v.PayARPU
				cell = row.AddCell()
				cell.Value = v.ARPUZ
				cell = row.AddCell()
				cell.Value = fmt.Sprintf("%v", v.PayRate)
				cell = row.AddCell()
				cell.Value = v.ResRate
				cell = row.AddCell()
				cell.Value = v.Res3Rate
				cell = row.AddCell()
				cell.Value = v.Res7Rate
			}

			err := file.Save("data_center.xlsx")
			if err != nil {
				panic(err)
			}

		}

	}
}
