package main

import (
	"errors"
	"fmt"
	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/storage"
	"fyne.io/fyne/v2/theme"
	"fyne.io/fyne/v2/widget"
	"github.com/flopp/go-findfont"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
	"reflect"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"time"
)

func main() {
	//新建一个app
	a := app.New()
	a.Settings().SetTheme(theme.DarkTheme())
	//新建一个窗口
	w := a.NewWindow("Excel智能抽取程序V1.0")
	//主界面框架布局
	MainShow(w)
	//尺寸
	w.Resize(fyne.Size{Width: 800, Height: 300})
	//w居中显示
	w.CenterOnScreen()
	//循环运行
	w.ShowAndRun()
	err := os.Unsetenv("FYNE_FONT")
	if err != nil {
		return
	}
}

// MainShow 主界面函数
func MainShow(w fyne.Window) {
	//获取模板目录
	title := widget.NewLabel("Excel智能抽取程序v1.0")
	templatePath := widget.NewLabel("模板文件路径:")
	templateEntry := widget.NewEntry()
	templatButton := widget.NewButton("打开", func() {
		templateButton := dialog.NewFileOpen(func(reader fyne.URIReadCloser, err error) {
			if err != nil {
				dialog.ShowError(err, w)
				return
			}
			if reader == nil {
				log.Println("Cancelled")
				return
			}
			templateEntry.SetText(reader.URI().Path())
		}, w)
		templateButton.SetFilter(storage.NewExtensionFileFilter([]string{".xlsx"}))
		templateButton.Show()
	})
	//获取目标目录
	targetPath := widget.NewLabel("抽取文件目录:")
	targetEntry := widget.NewEntry()
	targetButton := widget.NewButton("打开", func() {
		dialog.ShowFolderOpen(func(list fyne.ListableURI, err error) {
			if err != nil {
				dialog.ShowError(err, w)
				return
			}
			if list == nil {
				log.Println("Cancelled")
				return
			}
			targetEntry.SetText(list.Path())
		}, w)
	})

	begin := widget.NewEntry()
	end := widget.NewEntry()
	datec := widget.NewEntry()

	form := &widget.Form{
		Items: []*widget.FormItem{ // we can specify items in the constructor
			{Text: "起始行数(输入数字)", Widget: begin},
			{Text: "结尾行数(输入数字)", Widget: end},
			{Text: "日期格式年-月-日的列", Widget: datec},
		}, /*,
		OnSubmit: func() { // optional, handle form submission
			log.Println("Form submitted:", begin.Text)
			log.Println("Form submitted:", end.Text)
		},*/
	}
	isComplementselected := ""
	isComplementLabel := widget.NewLabel("是否开启向下补全:")
	isComplement := widget.NewSelect([]string{"是", "否"}, func(s string) {
		isComplementselected = s
	})
	isComplement.SetSelectedIndex(0)

	isNoSelected := ""
	isNoLabel := widget.NewLabel("第一列是否为序号列:")
	isNo := widget.NewSelect([]string{"是", "否"}, func(s string) {
		isNoSelected = s
	})
	isNo.SetSelectedIndex(0)

	isComplementFromc := widget.NewEntry()
	isComplementFrom := &widget.Form{
		Items: []*widget.FormItem{ // we can specify items in the constructor
			{Text: "需要补齐的列(输入大写字母使用,分割例如: B,C,D)", Widget: isComplementFromc},
		}, /*,
		OnSubmit: func() { // optional, handle form submission
			log.Println("Form submitted:", begin.Text)
			log.Println("Form submitted:", end.Text)
		},*/
	}
	text := widget.NewMultiLineEntry()
	text.Disable()

	save := widget.NewButton("保存", func() {
		log.Println("校验入参后开始调用逻辑代码")
		text.SetText("")
		text.Refresh()
		//调用逻辑代码 func(){ ...}
		err := verifyParameter(begin.Text, templateEntry.Text, end.Text, targetEntry.Text, "A", isNoSelected, isComplementselected, isComplementFromc.Text, datec.Text)
		if err != nil {
			text.SetText(err.Error())
			text.Refresh()
		} else {
			err, msg := start(begin.Text, templateEntry.Text, end.Text, targetEntry.Text, "A", isNoSelected, isComplementselected, isComplementFromc.Text, datec.Text)
			if err != nil {
				text.SetText(err.Error())
				text.Refresh()
			} else {
				text.SetText(msg)
				//text.SetText("本次成功抽取文件24个,共用时3ms,请检查模板...")
				text.Refresh()
			}

		}
	})
	head := container.NewCenter(title)
	templateBox := container.NewBorder(layout.NewSpacer(), layout.NewSpacer(), templatePath, templatButton, templateEntry)
	targetPathBox := container.NewBorder(layout.NewSpacer(), layout.NewSpacer(), targetPath, targetButton, targetEntry)
	isComplementBox := container.NewHBox(isComplementLabel, isComplement, isNoLabel, isNo)
	ctxPath := container.NewVBox(head, templateBox, targetPathBox, form, isComplementBox, isComplementFrom, save, text)
	w.SetContent(ctxPath)
}

func start(begin, tempPath, end, tigerPath, colO, t, isComplement, abcCompletionstr, abcStr string) (err error, msg string) {
	unix := time.Now().Unix()
	abcCompletion := strings.Split(abcCompletionstr, ",")
	abc := strings.Split(abcStr, ",")
	beginInt, err := strconv.Atoi(begin)
	endInt, err := strconv.Atoi(end)
	tb := true

	if t == "是" {
		tb = true
	} else {
		tb = false
	}
	isComplementBool := true
	if isComplement == "是" {
		isComplementBool = true
	} else {
		isComplementBool = false
	}

	//进行复制
	templateEmptyRows, endRows, sheet, err, i := startCopy(tempPath, tigerPath, beginInt, endInt, tb, colO)
	if err != nil {
		return
	}
	//数据清理补齐序号
	startCleanOff(templateEmptyRows, endRows, tempPath, beginInt, sheet, tb, endInt, colO, isComplementBool, abc, abcCompletion)

	u := time.Now().Unix() - unix
	msg = fmt.Sprintf("本次抽取文件共%d个,共用时%d秒,请检查模板", i, u)
	return
}
func print(w fyne.Window) {
	infinite := widget.NewProgressBarInfinite()
	w.SetContent(container.NewVBox(infinite))
}
func verifyParameter(begin, tempPath, end, tigerPath, colO, t, isComplement, abcCompletion, abc string) error {
	_, err := strconv.Atoi(begin)
	if err != nil {
		error := errors.New("起始行数需为数字 ")
		return error
	}
	_, err = os.ReadFile(tempPath)
	if err != nil {
		error := errors.New("未找到模板文件 ")
		return error
	}
	_, err = strconv.Atoi(end)
	if err != nil {
		error := errors.New("末尾行数需为数字 ")
		return error
	}
	_, err = os.ReadDir(tigerPath)
	if err != nil {
		error := errors.New("未找到目标目录 ")
		return error
	}
	if isComplement == "是" {
		reg := regexp.MustCompile(`^[A-Z,]+$`)
		if !reg.MatchString(abcCompletion) {
			error := errors.New("需要补齐的列请参考这种写法(大写字母与大写字母之间使用英文逗号分割):B,C,D,E,F")
			return error
		}
	}
	if abc != "" {
		reg := regexp.MustCompile(`^[A-Z,]+$`)
		if !reg.MatchString(abc) {
			error := errors.New("日期列请参考这种写法(大写字母与大写字母之间使用英文逗号分割):B,C,D,E,F")
			return error
		}
	}

	return nil
}
func init() {
	fontPaths := findfont.List()
	for _, path := range fontPaths {
		//楷体:simkai.ttf
		//黑体:simhei.ttf
		if strings.Contains(path, "simkai.ttf") {
			os.Setenv("FYNE_FONT", path)
			break
		}
	}
}

func startCleanOff(templateEmptyRows [][]string, endRows []string, path string, begin int, sheet string, t bool, end int, colO string, isComplement bool, abc []string, abcCompletion []string) {
	var clearExcel clearExcel
	err := clearExcel.newClearExcel(templateEmptyRows, endRows, path, begin, sheet, t, end, colO)
	if err != nil {
		return
	}
	//首先执行清理
	var ints []int
	count := 0
	rows := clearExcel.clearRows
	for rows.Next() {
		row, err := rows.Columns()
		if err != nil {
			fmt.Println(err)
		}
		//逐行比较是否存在
		if count > clearExcel.begin {
			b, err := clearExcel.excel.compare(clearExcel.templateEmptyRows, row)
			if err != nil {
				return
			}
			if b {
				ints = append(ints, count+1)
			}
		}
		count++
	}
	sort.Sort(sort.Reverse(sort.IntSlice(ints)))

	//清理空行完毕
	for _, i := range ints {
		clearExcel.clearFile.RemoveRow(sheet, i)
	}
	//重新赋值
	clearExcel.clearRows, err = clearExcel.clearFile.Rows(sheet)
	if err != nil {
		return
	}
	//获取当前文档总共行数
	clearFileRows, err := clearExcel.clearFile.Rows(sheet)
	if err != nil {
		fmt.Println(err)
		return
	}
	rowCount := 0
	for clearFileRows.Next() {
		_, err := clearFileRows.Columns()
		if err != nil {
			fmt.Println(err)
		}
		rowCount++
	}
	if err = rows.Close(); err != nil {
		fmt.Println(err)
	}
	//默认序列
	var ic []int
	for i := 0; i < rowCount-begin; i++ {
		ic = append(ic, i+1)
	}

	length := 0
	//如果第一列是序号列
	if clearExcel.excel.isNo {
		f := clearExcel.clearFile
		cols, err := f.Cols(sheet)
		if err != nil {
			fmt.Println(err)
			return
		}
		c := 0
		for cols.Next() {
			_, err := cols.Rows()
			if err != nil {
				fmt.Println(err)
			}
			if c == 0 && clearExcel.excel.end != 0 {
				if clearExcel.excel.end > clearExcel.excel.begin {
					var iu []int
					for i := 0; i < len(ic); i++ {
						iu = append(iu, i+1)
					}
					//取消合并单元格
					clearExcel.clearFile.UnmergeCell(sheet, fmt.Sprint(clearExcel.excel.colO, begin), fmt.Sprint(clearExcel.excel.colO, end))
					//clearExcel.clearFile.SetSheetCol(sheet, fmt.Sprint(clearExcel.excel.colO, begin), &iu)
					clearExcel.setSheetCol(sheet, iu)
					length = len(iu)
				}
				if c == 0 && clearExcel.excel.end == 0 {
					clearExcel.clearFile.UnmergeCell(sheet, fmt.Sprint(clearExcel.excel.colO, begin), fmt.Sprint(clearExcel.excel.colO, rowCount))
					//clearExcel.clearFile.SetSheetRow(sheet, fmt.Sprint(clearExcel.excel.colO, begin), &ic)
					clearExcel.setSheetCol(sheet, ic)
					length = len(ic)
				}
			}
			c++
			break
		}
	}
	//时间格式转换
	clearExcel.autoDate(sheet, abc, length)
	//列补偿
	if isComplement {
		//开启列向下补齐
		clearExcel.autoCompletion(sheet, abcCompletion, length)
	}

	clearExcel.clearExcelClose()
}
func (e *clearExcel) autoCompletion(sheet string, abcCompletion []string, length int) {
	//k开启自动补全
	//遍历补齐列
	for _, abc := range abcCompletion {
		//遍历补全次数
		prev := ""
		for i := 0; i < length; i++ {
			//值位置
			sprint := fmt.Sprint(abc, i+e.begin)
			value, err := e.clearFile.GetCellValue(sheet, sprint)
			if err != nil {
				continue
			}
			if value != "" {
				prev = value
			} else {
				e.clearFile.SetCellValue(sheet, sprint, prev)
			}

		}
	}
}
func (e *clearExcel) autoDate(sheet string, strs []string, length int) {
	//获取带有时间或者年月日的列
	for _, str := range strs {
		for i := 0; i < length; i++ {
			sprint := fmt.Sprint(str, i+e.begin)
			value, err := e.clearFile.GetCellValue(sheet, sprint)
			if err != nil {
				continue
			}
			if value != "" {
				day := convertToFormatDay(value)
				if day == "" {
					continue
				}
				e.clearFile.SetCellValue(sheet, sprint, day)
			}
		}
	}

}
func (e *clearExcel) setSheetCol(sheet string, iu []int) {
	for i, item := range iu {
		sprint := fmt.Sprint(e.excel.colO, (e.begin)+i)
		e.clearFile.SetCellValue(sheet, sprint, item)
	}
}
func (e *clearExcel) clearExcelClose() {
	e.clearFile.Save()
	e.clearRows.Close()
	e.clearFile.Close()
}
func (e *clearExcel) newClearExcel(templateEmptyRows [][]string, endRows []string, path string, begin int, sheet string, isNo bool, end int,
	colO string) (err error) {
	e.templateEmptyRows = templateEmptyRows
	e.endRows = endRows
	e.begin = begin
	//打开文件
	e.clearFile, err = excelize.OpenFile(path)
	if err != nil {
		return
	}
	e.clearRows, err = e.clearFile.Rows(sheet)
	if err != nil {
		return err
	}
	var excel excel
	excel.isNo = isNo
	excel.end = end
	excel.colO = colO
	e.excel = excel
	return
}

type clearExcel struct {
	//模板文件
	clearFile *excelize.File
	clearRows *excelize.Rows
	//空行集合
	templateEmptyRows [][]string
	//末行集合
	endRows []string
	//正文index
	begin int
	excel excel
}

func startCopy(tempPath string, tigerPath string, begin int, end int, b bool, colO string) ([][]string, []string, string, error, int) {
	var excel excel
	err := excel.newExcel(tempPath, tigerPath, begin, end, b, colO)
	if err != nil {
		printlnErr(err)
		return nil, nil, "", err, 0
	}
	//清理目标行
	excel.targetClear()
	//目标行复制到模板行
	excel.copy()

	err = excel.finish()
	if err != nil {
		printlnErr(err)
		return nil, nil, "", err, 0
	}

	return excel.templateEmptyRows, excel.endRows, excel.templateSheet, nil, len(excel.targetFile)
}
func printlnErr(err error) {
	fmt.Println("startCopy err = ", err)
}
func isIdenticalExclude0(once []string, twos []string) (b bool) {
	//比较两个集合除去第一个元素是否一致
	for i := 0; i < len(once); i++ {
		if i > 0 {
			if once[i] != twos[i] {
				return false
			}
		}
	}
	return true
}

type excel struct {
	colO string
	//表头起始行数
	begin int
	//表尾起始行数
	end int
	//末行模板
	endRows []string
	//第一列是否为序号
	isNo bool
	//模板文件
	templateFile *excelize.File
	//目标文件
	targetFile []*excelize.File
	//Sheet
	templateSheet string
	//模板行集合
	templateRows *excelize.Rows
	//模板空行集合
	templateEmptyRows [][]string
	//目标行集合
	targetRows [][]string
}

func (e *excel) newExcel(templatePath string, targetPath string, begin int, end int, isNo bool, colO string) error {
	var err error
	e.begin = begin
	e.end = end
	e.isNo = isNo
	e.colO = colO
	//打开模板
	e.templateFile, err = excelize.OpenFile(templatePath)
	//获取第一个sheet的名字
	e.templateSheet = e.templateFile.GetSheetName(0)
	//打开模板数组
	targets, err := os.ReadDir(targetPath)
	if err != nil {
		fmt.Println("打开目标目录错误.", err)
		return err
	}
	for _, target := range targets {
		tigerPath := targetPath + "\\" + target.Name()
		tigerFile, err := excelize.OpenFile(tigerPath)
		if err != nil {
			return err
		}
		e.targetFile = append(e.targetFile, tigerFile)
	}
	//获取行集合
	e.templateRows, err = e.templateFile.Rows(e.templateSheet)
	if err != nil {
		return err
	}

	//目标行集合
	for _, f := range e.targetFile {
		countTar := 0
		rows, err := f.Rows(e.templateSheet)
		if err != nil {
			fmt.Println(err)
			return nil
		}
		for rows.Next() {
			if countTar >= e.begin-1 {
				row, err := rows.Columns()
				if err != nil {
					fmt.Println(err)
				}
				e.targetRows = append(e.targetRows, row)
			}
			countTar++
		}

	}
	//提取空行模板及末行模板
	count := 0
	for e.templateRows.Next() {
		row, err := e.templateRows.Columns()
		if err != nil {
			fmt.Println(err)
			return err
		}
		if count >= begin {
			e.templateEmptyRows = append(e.templateEmptyRows, row)
		}
		if end > begin && count == end {
			e.endRows = row
		}
		count++
	}
	return err
}
func (e *excel) finish() (err error) {
	//关闭目标文件及行资源
	err = e.templateRows.Close()
	if err != nil {
		return
	}
	err = e.templateFile.Close()
	if err != nil {
		return
	}
	//保存文件
	err = e.templateFile.Save()
	if err != nil {
		return err
	}
	for _, file := range e.targetFile {
		err = file.Close()
		if err != nil {
			return err
		}
	}
	return nil
}

// 比较数组是否与空行模板相同,相同则返回true 否则false
func (e *excel) compare(rows [][]string, slice []string) (b bool, err error) {
	//rows := e.templateEmptyRows
	if slice == nil || len(slice) <= 0 {
		err = errors.New("compare input is empty")
		return
	}

	b = true
	if e.isNo {
		//剔除第一列进行比较
		for _, emptyList := range rows {
			//
			if len(slice) != len(emptyList) {
				return false, nil
			}
			//比较第一个数组是否和入参数组一致
			if isIdenticalExclude0(emptyList, slice) {
				return true, nil
			}
		}
		return false, nil
	} else {
		//第一列加入比较
		for _, row := range e.templateEmptyRows {
			if !reflect.DeepEqual(row, slice) {
				return false, nil
			}
		}
	}
	return
}

// 执行复制
func (e *excel) copy() {
	//目标集合
	rows := e.targetRows
	//遍历目标进行插入
	for i, row := range rows {
		//复制一个空行为下面做准备
		e.templateFile.DuplicateRow(e.templateSheet, e.begin+i)
		//插入新的行
		e.templateFile.SetSheetRow(e.templateSheet, fmt.Sprint(e.colO, e.begin+i), &row)
	}
}

// 清理目标空行
func (e *excel) targetClear() {
	rows := e.templateEmptyRows
	var newSlice [][]string
	for _, row := range e.targetRows {
		b, _ := e.compare(rows, row)
		if !b {
			newSlice = append(newSlice, row)
		}
	}
	e.targetRows = newSlice
}

func (e *excel) templateClear(beginClear int) {
	//获取模板行
	rows := e.templateRows
	empty := e.templateEmptyRows
	count := 0
	for rows.Next() {
		row, err := rows.Columns()
		if err != nil {
			fmt.Println(err)
		}
		if count > beginClear {
			//比较是否存在模板空行
			b, _ := e.compare(empty, row)
			if b {
				//删除当前位置的行
				e.templateFile.RemoveRow(e.templateSheet, count)
			}
		}
		count++
	}
}

// excel日期字段格式化 yyyy-mm-dd
func convertToFormatDay(excelDaysString string) string {
	// 2006-01-02 距离 1900-01-01的天数
	baseDiffDay := 38719 //在网上工具计算的天数需要加2天，什么原因没弄清楚
	curDiffDay := excelDaysString
	b, err := strconv.Atoi(curDiffDay)
	if err != nil {
		return ""
	}
	// 获取excel的日期距离2006-01-02的天数
	realDiffDay := b - baseDiffDay
	//fmt.Println("realDiffDay:",realDiffDay)
	// 距离2006-01-02 秒数
	realDiffSecond := realDiffDay * 24 * 3600
	//fmt.Println("realDiffSecond:",realDiffSecond)
	// 2006-01-02 15:04:05距离1970-01-01 08:00:00的秒数 网上工具可查出
	baseOriginSecond := 1136185445
	resultTime := time.Unix(int64(baseOriginSecond+realDiffSecond), 0).Format("2006年01月02日")
	return resultTime
}
