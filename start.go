package main
import (
	"fmt"
	"encoding/json"
	"github.com/tealeg/xlsx"
	"os"
	"io/ioutil"
	"excel_to_json/parseConfig"
	"strings"
	"archive/zip"
	"bytes"
)
func main() {
	readPath()
}
func readPath(){
	println("**** START ****")
	var config = parseConfig.New("config.json")
	var itemList = config.Get("data").([]interface{})
	for _,v := range itemList {
		var plat = v.(map[string]interface{})
		name := plat["name"].(string)
		inPath := plat["inPath"].(string)
		serverOutPath := plat["serverOutPath"].(string)
		clientOutPath := plat["clientOutPath"].(string)



		println("\n **** PROCESS "+name+"**** \n")
		processAll(inPath,serverOutPath,clientOutPath)
	}
	fmt.Println("\n **** DONE ****")
	fmt.Print("\n Press 'Enter' to continue...\n")
	fmt.Scanln()
}

func processAll(inpath string,serverPath string,clientPath string){
	files,_:=ioutil.ReadDir(inpath)
	var buf bytes.Buffer
	buf.WriteString("<?xml version=\"1.0\" encoding=\"utf-8\"?>\n")
	buf.WriteString("<mysql>\n")
	buf.WriteString("<database name=\"txtgame\">\n")

	for _,file := range files{
		itemBytes:=excelOp(inpath,file.Name(),serverPath,clientPath)
		if itemBytes==""{
			continue
		}
		buf.WriteString(itemBytes)
	}
	buf.WriteString("</database>")
	buf.WriteString("</mysql>")

	createSdata(serverPath,buf.String())

}



func createSdata(path string,filecontent string){
	println("create sdata.zip")
	fzip, _ := os.Create(path+"sdata.zip")
	w := zip.NewWriter(fzip)

	defer fzip.Close()
	defer w.Close()

	fw, _ := w.Create("sdata.xml")
	fw.Write([]byte(filecontent))

}



func excelOp(path string,fileName string,serverPath string,clientPath string)string {
	if strings.HasPrefix(fileName,"~"){
		return ""
	}
	println("process "+path+""+fileName)
	xlFile, err := xlsx.OpenFile(path+fileName)
	if err != nil {
		fmt.Println("open file error")
	}
	sheet := xlFile.Sheets[0]
	rowLen := len(sheet.Rows)

	celLen := len(sheet.Cols)
	var field= make([]string, celLen)
	var types= make([]string, celLen)

	var fieldClient= make([]interface{}, celLen)

	s := 0

	var cbody = make([][]interface{}, rowLen-1)
	cbody[0] = fieldClient


	fineNameArr := strings.Split(fileName,".")

	var buffer bytes.Buffer
	buffer.WriteString("<table name=\""+fineNameArr[0]+"\">\n")

	for idxRow, row := range sheet.Rows {
		if idxRow == 0 || idxRow == 1 {
			for cellIdx, cell := range row.Cells {
				text := strings.TrimSpace(cell.String())
				if idxRow == 0 {
					field[cellIdx] = text
					fieldClient[cellIdx] = text
					continue
				}
				if (idxRow == 1) {
					types[cellIdx] = text
					continue
				}
			}
			continue
		}
		buffer.WriteString("<row>")

		var cValue = make([]interface{}, celLen)
		for cellIdx, cell := range row.Cells {
			if types[cellIdx] == "int" {
				v, _ := cell.Int64()

				cValue[cellIdx] = v

			} else{
				itemCell:=strings.TrimSpace(cell.String())

				cValue[cellIdx] = itemCell
			}
			buffer.WriteString("<field name=\""+field[cellIdx]+"\">"+cell.String()+"</field>")
		}
		buffer.WriteString("</row>\n")
		cbody[s+1] = cValue

		s++
	}
	buffer.WriteString("</table>\n")
	cbyte,_ := json.Marshal(cbody)

	ioutil.WriteFile(clientPath+getOutputFileName(fileName),cbyte,0666)
	return buffer.String()
}

func getOutputFileName(excelName string) string{
	arr := strings.Split(excelName,"-")
	var len = len(arr)
	if(len==1){
		r := strings.Split(excelName,".")
		return r[0]+".json"
	}
	r := strings.Split(arr[1],".")
	return r[0]+".json"
}