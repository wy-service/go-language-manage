package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"os"
	"strconv"

	"github.com/xuri/excelize/v2"
)

func main() {
	excelToJson()
}

func log(ars ...interface{}) {
	fmt.Println(ars...)
}

const FILE_NAME = "./lang.xlsx"
const SAVE_PATH = "./json/"

func excelToJson() {
	f, err := excelize.OpenFile(FILE_NAME)
	if err != nil {
		log("open excel error:", err)
	}
	defer func() {
		if err := f.Close(); err != nil {
			log("close excel error:", err)
		}
	}()

	// get sheet
	sheetName := f.GetSheetName(0)
	rows, err := f.GetRows(sheetName)
	if err != nil {
		log("get rows error:", err)
	}
	langList := rows[0]

	// to map
	m := make(map[string]map[string]string)
	for i := 1; i < len(rows); i++ {
		row := rows[i]
		key := row[0]

		for j := 1; j < len(row); j++ {
			colCell := row[j]
			jsonMap := make(map[string]string, 0)
			lang := langList[j]
			v, ok := m[lang]
			if ok {
				jsonMap = v
			}
			// 如果希望有值，则需要判断
			// if len(colCell) > 0 {
			// }
			s2, err := strconv.Unquote(`"` + colCell + `"`) // 转义
			if err != nil {
				log("strconv.Unquote error:", err)
			}
			jsonMap[key] = s2
			m[lang] = jsonMap

		}
	}

	// save json
	err = saveJson(langList, m)
	if err != nil {
		log("save json error:", err)
	}
	log("save json success")

}

func saveJson(langList []string, m map[string]map[string]string) error {
	for i := 1; i < len(langList); i++ {
		lang := langList[i]
		mapData := m[lang]

		// 保留转义
		// jsonData, err := json.Marshal(mapData) //MarshalIndent可以美化格式
		// 原样输出
		jsonData, err := JSONMarshal(mapData)
		if err != nil {
			return err
		}
		fileName := fmt.Sprintf("%s.json", lang)
		err = PathFix(SAVE_PATH)
		if err != nil {
			return err
		}
		err = os.WriteFile(fmt.Sprintf("%s/%s", SAVE_PATH, fileName), jsonData, 0644)
		if err != nil {
			return err
		}
	}

	return nil
}

// golang标准包encoding/json 默认地 将&<>这三个字符进行转义。所以自己重写了这个方法，不转义
func JSONMarshal(t interface{}) ([]byte, error) {
	buffer := &bytes.Buffer{}
	encoder := json.NewEncoder(buffer)
	encoder.SetEscapeHTML(false)
	err := encoder.Encode(t)
	return buffer.Bytes(), err
}

// 判断路径是否存在，不存在就创建，返回路径
func PathFix(path string) error {
	_, err := os.Stat(path)
	if err == nil {
		return nil
	}
	if os.IsNotExist(err) {
		err := os.MkdirAll(path, os.ModePerm)
		return err
	}
	return err
}
