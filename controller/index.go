package controller

import (
	"encoding/json"
	"github.com/360EntSecGroup-Skylar/excelize"
	"io/ioutil"
	"log"
	"net/http"
	"strconv"
)

//  通用样式，水平垂直居中
const styleConfigJson = "{\"alignment\": {\"horizontal\": \"center\",\"vertical\": \"center\"}}"

//  标题样式，水平垂直居中，字体16号 颜色 红色
const fontJson = "{\"alignment\": {\"horizontal\": \"center\",\"vertical\": \"center\"},\"font\":{\"size\":16,\"color\":\"#ff4d4d\"}}"
const importFileUrl = "controller/Book1.xlsx"
const outPutFileUrl = "controller/Book1.xlsx"

type RowConfig struct {
	RowTitle []struct {
		SheetName    string `json:"sheetName"`
		RowTitleName string `json:"rowTitleName"`
		IsMergeCell  string `json:"isMergeCell"`
	} `json:"rowTitleName"`
}

func GetExcel(w http.ResponseWriter, r *http.Request) {
	result, _ := ioutil.ReadAll(r.Body)
	var rowConfig RowConfig
	json.Unmarshal(result, &rowConfig)

	f, err := excelize.OpenFile(importFileUrl)
	if err != nil {
		log.Fatal("文件读取失败" + err.Error())
		return
	}

	// 遍历Sheet页
	for _, sheetName := range f.GetSheetMap() {
		rows := f.GetRows(sheetName)
		// 遍历所有行
		for row, rowname := range rows {
			if row == 0 {
				// 控制单元格样式
				CellStyle(f, sheetName, rows, rowname)
				// 控制合并单元格
				mergeCellContent(f, sheetName, rows, rowname, rowConfig)
			}
		}
	}

	// 根据指定路径保存文件
	if err := f.SaveAs(outPutFileUrl); err != nil {
		println(err.Error())
	}

}

func switchCode(row, line int) string {
	c1 := string(rune(line + 65))
	c2 := strconv.Itoa(row + 1)
	c3 := c1 + c2
	return c3
}

// 控制单元样式
func CellStyle(f *excelize.File, sheetName string, rows [][]string, rowname []string) {
	// 通用样式
	style1, _ := f.NewStyle(styleConfigJson)
	// 标题样式
	style2, _ := f.NewStyle(fontJson)
	for index, _ := range rowname {
		for index2, _ := range rows {
			if index2 == 0 {
				f.SetCellStyle(sheetName, switchCode(index2, index), switchCode(index2, index), style2)
			} else {
				f.SetCellStyle(sheetName, switchCode(index2, index), switchCode(index2, index), style1)
			}
		}
	}
}

func mergeCellContent(f *excelize.File, sheetName string, rows [][]string, rowname []string, rowConfig RowConfig) {

	for index, _ := range rowname {
		toCoordinate := ""
		endCoordinates := ""
		for index2, _ := range rows {
			if !isMergeCell(rowConfig, rows[index2][index]) {
				break
			}
			if index2 == 0 {
				continue
			}
			if rows[index2][index] == rows[index2-1][index] {
				if toCoordinate == "" {
					toCoordinate = switchCode(index2-1, index)
				}
			} else {
				if toCoordinate != "" {
					endCoordinates = switchCode(index2-1, index)
					f.MergeCell(sheetName, toCoordinate, endCoordinates)
					toCoordinate = ""
					endCoordinates = ""
				}
			}
			if index2 == len(rows)-1 {
				if toCoordinate != "" {
					f.MergeCell(sheetName, toCoordinate, switchCode(len(rows)-1, index))
					toCoordinate = ""
				}
			}
		}
	}
}

// 根据配置文件,决定合并列
func isMergeCell(rowConfig RowConfig, cellValue string) bool {
	for _, value := range rowConfig.RowTitle {
		if cellValue == value.RowTitleName && value.IsMergeCell == "N" {
			return false
		}
	}
	return true
}
