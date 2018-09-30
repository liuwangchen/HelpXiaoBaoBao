package main

import (
	"flag"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"io"
	"io/ioutil"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"sync"
	"time"
)

type SearchResult struct {
	tel    string
	result os.FileInfo
}

type NameCount struct {
	name  string
	count int
}

var dataMap *sync.Map = new(sync.Map)

func main() {
	findDir := flag.String("f", "", "findDir 要搜索的文件夹目录")
	outDir := flag.String("o", "", "outDir 要输出的文件夹目录")
	xlsxFile := flag.String("x", "", "xlsxFile 目标xlsx")
	month := flag.Int("m", 1, "month 要查找的月份")
	flag.Parse()
	if findDir != nil && outDir != nil && xlsxFile != nil && month != nil {
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
		findFileInfoList, err := ioutil.ReadDir(*findDir)
		if err != nil {
			return
		}
		inChan := make(chan string)
		outChan := make(chan SearchResult)
		go InitInChan(file, inChan, *month)
		go DoSearch(4, findFileInfoList, inChan, outChan)
		for value := range outChan {
			v, _ := dataMap.Load(value.tel)
			nc := v.(*NameCount)
			nc.count++
			handleResult(*outDir, *findDir, value.result.Name(), nc.name+strconv.Itoa(nc.count)+"_")
		}
	}
}

func InitInChan(file *excelize.File, inChan chan string, month int) {
	rows := file.GetRows("资产.人力.二手车")
	for i, row := range rows {
		if i > 0 {
			t, _ := time.Parse("2006-01-02", row[13])
			if int(t.Month()) == month {
				dataMap.Store(row[11], &NameCount{
					name:  row[5],
					count: 0,
				})
				inChan <- row[11]
			}
		}
	}
	close(inChan)
}

func DoSearch(workerCount int, sourceList []os.FileInfo, inChan <-chan string, outChan chan<- SearchResult) {
	wg := new(sync.WaitGroup)
	for i := 0; i < workerCount; i++ {
		wg.Add(1)
		go search(inChan, outChan, sourceList, wg)
	}
	wg.Wait()
	close(outChan)
}

func search(inChan <-chan string, outChan chan<- SearchResult, findFileInfoList []os.FileInfo, wg *sync.WaitGroup) {
	defer wg.Done()
	for tel := range inChan {
		for _, fileInfo := range findFileInfoList {
			if strings.Contains(fileInfo.Name(), tel) {
				outChan <- SearchResult{
					tel:    tel,
					result: fileInfo,
				}
			}
		}
	}
}

func handleResult(outDir string, inDir string, fileName string, addString string) {
	CopyFile(filepath.Join(outDir, fileName), filepath.Join(inDir, fileName))
	os.Rename(filepath.Join(outDir, fileName), filepath.Join(outDir, addString+fileName))
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
