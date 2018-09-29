package main

import (
	"flag"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"io"
	"io/ioutil"
	"os"
	"path/filepath"
	"strings"
	"time"
)

func main() {
	findDir := flag.String("f", "", "findDir 要搜索的文件夹目录")
	outDir := flag.String("o", "", "outDir 要输出的文件夹目录")
	xlsxFile := flag.String("x", "", "xlsxFile 目标xlsx")
	month := flag.Int("m", 1, "month 要查找的月份")
	flag.Parse()
	if findDir != nil && outDir != nil && xlsxFile != nil && month != nil {
		//findDir := "C:\\Users\\Administrator\\Desktop\\findDir"
		//outDir := "C:\\Users\\Administrator\\Desktop\\outDir"
		if !PathExists(*findDir) {
			fmt.Println("找不到finddir：", *findDir)
			return
		}
		if !PathExists(*outDir) {
			os.Mkdir(*outDir, os.ModePerm)
		}
		file, err := excelize.OpenFile(*xlsxFile)
		if err != nil {
			fmt.Println(err)
			return
		}
		rows := file.GetRows("资产.人力.二手车")
		dataList := make([][2]string, 0, len(rows))
		for i, row := range rows {
			if i > 0 {
				t, _ := time.Parse("2006-01-02", row[13])
				if int(t.Month()) == *month {
					dataList = append(dataList, [2]string{row[5], row[11]})
				}
			}
		}
		findFileInfoList, err := ioutil.ReadDir(*findDir)
		if err != nil {
			return
		}
		for _, data := range dataList {
			fileInfo := search(data[1], findFileInfoList)
			if fileInfo != nil {
				fileName := (*fileInfo).Name()
				CopyFile(filepath.Join(*outDir, fileName), filepath.Join(*findDir, fileName))
				os.Rename(filepath.Join(*outDir, fileName), filepath.Join(*outDir, data[0]+fileName))
			}
		}
	}
}

func search(searchText string, findFileInfoList []os.FileInfo) *os.FileInfo {
	for _, fileInfo := range findFileInfoList {
		if strings.Contains(fileInfo.Name(), searchText) {
			return &fileInfo
		}
	}
	return nil
}

func CopyFile(dstName, srcName string) (written int64, err error) {
	src, err := os.Open(srcName)
	if err != nil {
		return
	}
	defer src.Close()

	dst, err := os.OpenFile(dstName, os.O_WRONLY|os.O_CREATE, 0644)
	if err != nil {
		return
	}
	defer dst.Close()
	return io.Copy(dst, src)
}

func PathExists(path string) bool {
	_, err := os.Stat(path)
	if err == nil {
		return true
	}
	if os.IsNotExist(err) {
		return false
	}
	return false
}
