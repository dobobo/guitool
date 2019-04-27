package main

import (
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"os/exec"
	"strings"
	"regexp"
	"github.com/rentiansheng/xlsx"
	"github.com/lxn/walk"
	. "github.com/lxn/walk/declarative"
)

type MyMainWindow struct {
	*walk.MainWindow
	edit   *walk.TextEdit
	sourcePath   string
	destPath string
	transPath string
	sourceSearchBox *walk.LineEdit
	destSearchBox *walk.LineEdit
	transSearchBox *walk.LineEdit
	selectExcelSheetName *walk.LineEdit
	dir string
}

func main() {
	mw := &MyMainWindow {}

	mw.dir, _ = os.Getwd()

	MW := MainWindow{
		AssignTo: &mw.MainWindow,
		Title: "自動翻訳",
		MinSize: Size {300, 200},
		Size   : Size {500, 400},
		Layout: VBox {},
		Children: []Widget {
			GroupBox {
                Layout: HBox {},
                Children: []Widget {
					LineEdit {
                        AssignTo: &mw.transSearchBox,
                    },
                    PushButton {
						Text: "翻訳Excelファイル",
						OnClicked: mw.transPbClicked,
                    },
				},
			},
			GroupBox {
                Layout: HBox {},
                Children: []Widget {
					Label{Font: Font{PointSize: 12}, Text: "Excelシート名指定"},
					LineEdit {
                        AssignTo: &mw.selectExcelSheetName,
                    },
				},
			},
			GroupBox {
                Layout: HBox {},
                Children: []Widget {
                    LineEdit {
                        AssignTo: &mw.sourceSearchBox,
                    },
                    PushButton {
						Text: "元データファイル",
						OnClicked: mw.sourcePbClicked,
					},
				},
			},
			GroupBox {
                Layout: HBox {},
                Children: []Widget {
					LineEdit {
                        AssignTo: &mw.destSearchBox,
                    },
                    PushButton {
						Text: "保存先フォルダ",
						OnClicked: mw.dest_pbClicked,
                    },
				},
			},
			PushButton {
				Text: "実行",
				OnClicked: mw.execute,
			},
		},
	}
	if _, err := MW.Run(); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
}

func (mw *MyMainWindow) sourcePbClicked() {

	dlg := new(walk.FileDialog)
	dlg.FilePath = mw.sourcePath
	dlg.Title    = "翻訳する元データファイルを選択してください"
	dlg.Filter   = "All files (*.*)|*.*"

	if ok, err := dlg.ShowOpen(mw); err != nil {
		walk.MsgBox(mw, "エラー", "ファイルオープンエラー", walk.MsgBoxOK)
		return
	} else if !ok {
		return
	}
	mw.sourcePath = dlg.FilePath
	s := fmt.Sprintf("%s\r\n", mw.sourcePath)
	mw.sourceSearchBox.SetText(s)
}

func (mw *MyMainWindow) dest_pbClicked() {

	dlg := new(walk.FileDialog)
	dlg.FilePath = mw.destPath
	dlg.Title    = "翻訳したファイルを保存するフォルダ先を選択してください"
	dlg.Filter   = "All files (*.*)|*.*"

	if ok, err := dlg.ShowBrowseFolder(mw); err != nil {
		walk.MsgBox(mw, "エラー", "ファイルオープンエラー", walk.MsgBoxOK)
		return
	} else if !ok {
		return
	}
	mw.destPath = dlg.FilePath
	s := fmt.Sprintf("%s\r\n", mw.destPath)
	mw.destSearchBox.SetText(s)
}

func (mw *MyMainWindow) transPbClicked() {

	dlg := new(walk.FileDialog)
	dlg.FilePath = mw.transPath
	dlg.Title    = "翻訳Excelファイルを選択してください"
	dlg.Filter   = "All files (*.*)|*.*"

	if ok, err := dlg.ShowOpen(mw); err != nil {
		walk.MsgBox(mw, "エラー", "ファイルオープンエラー", walk.MsgBoxOK)
		return
	} else if !ok {
		return
	}
	mw.transPath = dlg.FilePath
	s := fmt.Sprintf("%s\r\n", mw.transPath)
	mw.transSearchBox.SetText(s)
}

func (mw *MyMainWindow) execute() {
	transSearchBoxText := mw.transSearchBox.Text()
  sourceSearchBoxText := mw.sourceSearchBox.Text()
	destSearchBoxText := mw.destSearchBox.Text()
	
	dirNameSlice_hoge := strings.Split(sourceSearchBoxText, "\\")
	sourceName := dirNameSlice_hoge[len(dirNameSlice_hoge)-1]
	sourceName = strings.TrimRight(sourceName, "\r\n")

	if !strings.Contains(transSearchBoxText, ".xlsx") {
		walk.MsgBox(mw, "エラー", "翻訳ファイルはxlsx形式のものを選んでください", walk.MsgBoxOKCancel)
		return
	}
	if strings.Contains(sourceName, ".rvdata2") {
		sourceName = strings.Replace(sourceName, ".rvdata2", "", -1)
	}else {
		walk.MsgBox(mw, "エラー", "元データはrvdata2形式のものを選んでください", walk.MsgBoxOKCancel)
		return
	}

	dirNameSliceLtrim := dirNameSlice_hoge[:len(dirNameSlice_hoge)-1]

	// root_dir := strings.Join(dirNameSliceLtrim[:5], "/")

	dirNameLtrim := strings.Join(dirNameSliceLtrim, "/")
	dirNameLtrim = dirNameLtrim + "/" + sourceName

	if sourceName == "Scripts" {
		walk.MsgBox(mw, "エラー", "このゲームデータは翻訳しないでください。" + sourceName + ".rvdata2", walk.MsgBoxOKCancel)
		return
	}

	destFilePath := strings.TrimRight(destSearchBoxText, "\r\n")
	destFilePath += "./変換中"
	destFilePath = strings.TrimRight(destFilePath, "\r\n")
	destSearchBoxText = strings.TrimRight(destSearchBoxText, "\r\n")
	// destSearchBoxText --> 末尾に改行入ってない?
	if err := os.MkdirAll(mw.dir + "\\変換中\\", 0777); err != nil {
		walk.MsgBox(mw, "エラー", "変換するための一時フォルダを作成できません", walk.MsgBoxOKCancel)
		os.Exit(0)
	}

	sourceFilePath := strings.TrimRight(sourceSearchBoxText, "\r\n")

	err := exec.Command(mw.dir + "\\rv2da.exe", "-d", sourceFilePath, "-o", mw.dir + "\\変換中\\").Run()
	if err != nil {
		print(mw.dir)
		walk.MsgBox(mw, "エラー", "rvdata2をjsonに変換中にエラーが発生しました", walk.MsgBoxOKCancel)
		os.Exit(0)
	}
  
	sheetNum := 0

	dataFile := mw.dir + "\\変換中\\" + sourceName+".json"
 
	transMap := map[string]string{}
	untransMap := map[string]string{}
	transSearchBoxText = strings.TrimRight(transSearchBoxText, "\r\n")
	excel, err1 := xlsx.OpenFile(transSearchBoxText) 
	if err1 != nil {
		walk.MsgBox(mw, "エラー", "Excelファイルを開けませんでした", walk.MsgBoxOKCancel)
	  return
	}
  
	if mw.selectExcelSheetName.Text() != "" {
		sourceName =  mw.selectExcelSheetName.Text()
	}
	
	for index, sheet := range excel.Sheets {
	  if sheet.Name == sourceName {
			fmt.Println("Excelシートの存在確認: " + sheet.Name)
			sheetNum = index
			break;
	  }
	   if len(excel.Sheets)-1 <= index {
			walk.MsgBox(mw, "エラー", "error: 与えられた名前のExcelシートはありません。"+sourceName, walk.MsgBoxOKCancel)
			os.Exit(0)
	   }
	}
  
	for _, row := range excel.Sheets[sheetNum].Rows {
		for _, _ = range row.Cells {
			key := row.Cells[0].String()
			value := row.Cells[1].String()
			transMap[key] = value
			break
		}
	}

	untransMap = transMap
	jsonString, err := ioutil.ReadFile(dataFile)
	if err != nil {
		walk.MsgBox(mw, "エラー", "一時フォルダからjsonファイルを読み込ませんでした", walk.MsgBoxOKCancel)
		os.Exit(1)
	}

	threeOrMoreLineFeeds := ""
  
	sourceStr := string(jsonString)
	for key, value := range transMap {
	  replaceStr := "\\"+"\""
	  value = strings.Replace(value, "\"", replaceStr, -1) 
	  key = strings.Replace(key, "…", "…", -1)
	  keyTmp := strings.Replace(key, "*", "\\*", -1)
	  keyTmp = strings.Replace(keyTmp, "+", "\\+", -1)
	  keyTmp = strings.Replace(keyTmp, "{", "\\{", -1)
	  keyTmp = strings.Replace(keyTmp, "}", "\\}", -1)
	  keyTmp = strings.Replace(keyTmp, "^", "\\^", -1)
	  keyTmp = strings.Replace(keyTmp, "$", "\\$", -1)
	  keyTmp = strings.Replace(keyTmp, "-", "\\-", -1)
	  keyTmp = strings.Replace(keyTmp, "|", "\\|", -1)
	  keyTmp = strings.Replace(keyTmp, "(", "\\(", -1)
	  keyTmp = strings.Replace(keyTmp, ")", "\\)", -1)
	  keyTmp = strings.Replace(keyTmp, "+", "\\+", -1)
		keyTmp = strings.Replace(keyTmp, "?", "\\?", -1)
		
	  maxOneLine := 50
	  if strings.Contains(sourceName, "Map") || strings.Contains(sourceName, "Common") {
	  	  maxOneLine = 59
	  }

	  if len(value) > maxOneLine {
		if len(value) > 177 {
		  threeOrMoreLineFeeds = value + "\r\n"
		  fmt.Println("改行3回以上必要" + threeOrMoreLineFeeds)
		  continue
		}
		value = strings.Replace(value, "\r\n", " ", -1)
		sourceChar := ""
		sourceChars := strings.Split(value, "\r\n")
		value = ""
		for _, sourceCharsValue := range sourceChars {
		  char := strings.Split(sourceCharsValue, " ")
		  if len(sourceCharsValue) >= maxOneLine {
				charLength := 0
				for charIndex, charValue := range char {
					charLength += len(charValue) + 1
					if charLength > maxOneLine {
					char[charIndex] = "\\r\\n" + charValue
					charLength = len(charValue)
					}
					sourceChar = strings.Join(char, " ")    
				}
				value += sourceChar
			}
		}
	}

	  pettern := regexp.MustCompile("\"" + keyTmp + "\"")
	  fmt.Println(pettern)
	
	  tmp_sourceStr := sourceStr
	  sourceStr = pettern.ReplaceAllString(sourceStr, "\""+value+"\"")

	  if sourceStr != tmp_sourceStr {
		delete(untransMap, key)  
	  }
	}
	
	if err := os.MkdirAll(destSearchBoxText+"\\未翻訳", 0777); err != nil {
		walk.MsgBox(mw, "エラー", "未翻訳フォルダを作成できませんでした", walk.MsgBoxOKCancel)
		os.Exit(0)
	}

	if err := os.Remove(destSearchBoxText +"\\"+sourceName+".json"); err != nil {
	}
	trans_file, err := os.OpenFile(destSearchBoxText +"\\"+sourceName+".json", os.O_RDWR|os.O_CREATE|os.O_EXCL, 0666)
	if err != nil {
		walk.MsgBox(mw, "エラー", "既に保存するべきファイルが存在するため実行できません", walk.MsgBoxOKCancel)
	  log.Fatal(err)
	}
	fmt.Fprintln(trans_file, sourceStr)
  
	untrans_text := "日本語 : 英語 \r\n"
	for key, value := range untransMap {
	  untrans_text += key
	  untrans_text += ":" 
	  untrans_text += value
	  untrans_text += "\r\n"
	}

	if err := os.Remove(destSearchBoxText + "\\未翻訳\\"+sourceName+"未翻訳ver.txt"); err != nil {
	}
	untrans_file, err := os.OpenFile(destSearchBoxText + "\\未翻訳\\"+sourceName+"未翻訳ver.txt", os.O_RDWR|os.O_CREATE|os.O_EXCL, 0666)
	if err != nil {
	  walk.MsgBox(mw, "エラー", "既に保存するべきファイルが存在するため実行できません", walk.MsgBoxOKCancel)
	  log.Fatal(err)
	}
	fmt.Println("----------未翻訳-----------------")
	fmt.Println(untrans_text)
	fmt.Println("---------------------------------")
	fmt.Fprintln(untrans_file, untrans_text)
  
	
	erro := exec.Command(mw.dir + "\\rv2da.exe", "-c", destSearchBoxText +"\\"+sourceName+".json", "-o", destSearchBoxText).Run()
	if erro != nil {
		print(erro)
		print(destSearchBoxText +"\\"+sourceName+".json")
		walk.MsgBox(mw, "エラー", "翻訳したjsonファイルをrvdata2に変換中にエラーが発生しました", walk.MsgBoxOKCancel)
		os.Exit(0)
	}

	transJsonString, err := ioutil.ReadFile(destSearchBoxText +"\\"+sourceName+".json")
	if err != nil {
		walk.MsgBox(mw, "エラー", "翻訳済みのjsonファイルを読み込ませんでした", walk.MsgBoxOKCancel)
		os.Exit(1)
	}

	japaneseKey := "\"" +".*([ぁ-ん]*[ァ-ヶ]*[亜-熙])+.*" + "\"" 

	japanesePettern := regexp.MustCompile(japaneseKey)
	foobar := japanesePettern.FindAllStringSubmatch(string(transJsonString), -1)
	foobarText := ""
	for index := range foobar {
		foobarText += foobar[index][0] + " \r\n"
	}
	if e := os.MkdirAll(destSearchBoxText+"\\未翻訳ver2\\", 0777); e != nil {
		walk.MsgBox(mw, "エラー", "未翻訳フォルダを作成できませんでした", walk.MsgBoxOKCancel)
		os.Exit(0)
	}

	if err := os.Remove(destSearchBoxText + "\\未翻訳ver2\\"+sourceName+"未翻訳ver.txt"); err != nil {
	}
	foobar_file, err := os.OpenFile(destSearchBoxText + "\\未翻訳ver2\\"+sourceName+"未翻訳ver.txt", os.O_RDWR|os.O_CREATE|os.O_EXCL, 0666)
	if err != nil {
		walk.MsgBox(mw, "エラー", "既に保存するべきファイルが存在するため実行できません", walk.MsgBoxOKCancel)
		log.Fatal(err)
	}
	fmt.Fprintln(foobar_file, foobarText)

	transMap = make(map[string]string)
	untransMap = make(map[string]string)
	walk.MsgBox(mw, "完了", "処理が完了しました", walk.MsgBoxOKCancel)
	os.Exit(0)
}

