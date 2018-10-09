package main

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"excel_to_json/parseConfig"
	"fmt"
	"github.com/tealeg/xlsx"
	"io/ioutil"
	"os"
	"strings"
	"unsafe"
)

func main() {
	readPath()
}
func readPath() {
	println("**** START ****")
	var config = parseConfig.New("config.json")
	var itemList = config.Get("data").([]interface{})
	for _, v := range itemList {
		var plat = v.(map[string]interface{})
		name := plat["name"].(string)
		inPath := plat["inPath"].(string)
		serverOutPath := plat["serverOutPath"].(string)
		clientOutPath := plat["clientOutPath"].(string)

		println("\n **** PROCESS " + name + "**** \n")
		processAll(inPath, serverOutPath, clientOutPath)
	}
	fmt.Println("\n **** DONE ****")
	fmt.Print("\n Press 'Enter' to continue...\n")
	fmt.Scanln()
}

func processAll(inpath string, serverPath string, clientPath string) {
	files, _ := ioutil.ReadDir(inpath)
	var buf bytes.Buffer
	buf.WriteString("<?xml version=\"1.0\" encoding=\"utf-8\"?>\n")
	buf.WriteString("<mysql>\n")
	buf.WriteString("<database name=\"txtgame\">\n")

	for _, file := range files {
		itemBytes := excelOp(inpath, file.Name(), serverPath, clientPath)
		if itemBytes == "" {
			continue
		}
		buf.WriteString(itemBytes)
	}
	buf.WriteString("</database>")
	buf.WriteString("</mysql>")

	createSdata(serverPath, buf.String())
}

func createSdata(path string, filecontent string) {
	println("create sdata.zip")
	fzip, _ := os.Create(path + "sdata.zip")
	w := zip.NewWriter(fzip)

	defer fzip.Close()
	defer w.Close()

	fw, _ := w.Create("sdata.xml")
	fw.Write([]byte(filecontent))
}

func excelOp(path string, fileName string, serverPath string, clientPath string) string {
	if strings.HasPrefix(fileName, "~") {
		return ""
	}
	if !strings.HasSuffix(fileName, "xlsx") {
		return ""
	}
	println("process " + path + "" + fileName)
	xlFile, err := xlsx.OpenFile(path + fileName)
	if err != nil {
		fmt.Println("open file error")
	}
	sheet := xlFile.Sheets[0]
	rowLen, s := 0, 0

	celLen := len(sheet.Cols)
	var field = make([]string, celLen)
	var types = make([]string, celLen)

	var fieldClient = make([]interface{}, celLen)

	for _, row := range sheet.Rows {
		if row.Cells[0].String() != "" {
			rowLen++
		}
	}

	var cbody = make([][]interface{}, rowLen-1)
	cbody[0] = fieldClient

	fineNameArr := strings.Split(fileName, ".")

	var buffer bytes.Buffer
	buffer.WriteString("<table name=\"" + fineNameArr[0] + "\">\n")

	for idxRow, row := range sheet.Rows {
		if idxRow == 0 || idxRow == 1 {
			for cellIdx, cell := range row.Cells {
				text := strings.TrimSpace(cell.String())
				if idxRow == 0 {
					field[cellIdx] = text
					fieldClient[cellIdx] = text
					continue
				}
				if idxRow == 1 {
					types[cellIdx] = text
					continue
				}
			}
			continue
		}
		zeroV, _ := row.Cells[0].Int()
		if field[0] == "id" && zeroV == -1 {
			continue
		}
		fieldContent := []string{}

		var cValue = make([]interface{}, celLen)
		for cellIdx, cellName := range field {
			if types[cellIdx] == "int" {
				cValue[cellIdx] = 0
				if cellIdx < len(row.Cells) {
					v, _ := row.Cells[cellIdx].Int64()
					if v < 0 {
						v = 0
					}
					cValue[cellIdx] = v
				}
			} else {
				cValue[cellIdx] = ""
				if cellIdx < len(row.Cells) {
					v := row.Cells[cellIdx].String()
					cValue[cellIdx] = v

				}
			}

			tempV := ""
			//tempV := InterfaceToJsonString(cValue[cellIdx])
			tempValue, ok := cValue[cellIdx].(string)
			if ok {
				tempV = tempValue
			} else {
				tempValue, ok := cValue[cellIdx].(int)
				if !ok {
					tempValue2 := cValue[cellIdx].(int64)
					tempV = IntToString(tempValue2)
				} else {
					tempV = IntToString(tempValue)
				}
			}
			if tempV == "nil" {
				tempV = ""
			}
			itemV := "<field name=\"" + cellName + "\">"
			if len(tempV) > 0 {
				itemV += tempV
			}
			itemV += "</field>"
			fieldContent = append(fieldContent, itemV)
		}

		if len(fieldContent) > 0 {
			buffer.WriteString("<row>")
			for _, item := range fieldContent {
				buffer.WriteString(item)
			}
			buffer.WriteString("</row>\n")
		}

		cbody[s+1] = cValue
		s++
	}
	buffer.WriteString("</table>\n")
	cbyte, _ := json.Marshal(cbody)

	ioutil.WriteFile(clientPath+getOutputFileName(fileName), cbyte, 0666)
	return buffer.String()
}

func InterfaceToJsonString(d interface{}) string {
	data, err := json.Marshal(d)
	if err != nil {
		return ""
	}
	if len(data) == 0 {
		return ""
	}
	return BytesToString(data)
}
func BytesToString(b []byte) string {
	return *(*string)(unsafe.Pointer(&b))
}

func getOutputFileName(excelName string) string {
	arr := strings.Split(excelName, "-")
	var len = len(arr)
	if len == 1 {
		r := strings.Split(excelName, ".")
		return r[0] + ".json"
	}
	r := strings.Split(arr[1], ".")
	return r[0] + ".json"
}
func IntToString(n interface{}) string {
	return fmt.Sprintf("%d", n)
}
