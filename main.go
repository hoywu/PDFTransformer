package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func convertToPDF(filePath []string, savePath string) {
	// 在当前线程上初始化 COM 库
	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	// 创建 COM 对象
	var unknownWord, unknownPPT *ole.IUnknown
	var err error
	unknownWord, err = oleutil.CreateObject("Word.Application")
	unknownPPT, err = oleutil.CreateObject("PowerPoint.Application")
	if err != nil {
		log.Fatalf("Error creating application object for %s: %v", filePath, err)
	}
	defer unknownWord.Release()
	defer unknownPPT.Release()

	// 获取 IDispatch 接口
	word, err := unknownWord.QueryInterface(ole.IID_IDispatch)
	ppt, err := unknownPPT.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		log.Fatalf("Error querying interface: %v", err)
	}
	defer word.Release()
	defer ppt.Release()

	// 设置应用程序不可见
	word.PutProperty("Visible", false)
	ppt.PutProperty("Visible", false)

	// 获取 Documents 对象
	documents := oleutil.MustGetProperty(word, "Documents").ToIDispatch()
	defer documents.Release()
	// 获取 Presentations 对象
	presentations := oleutil.MustGetProperty(ppt, "Presentations").ToIDispatch()
	defer presentations.Release()

	// 打开文件并转换为 PDF
	var document *ole.IDispatch
	for _, path := range filePath {
		log.Printf("Converting: %s", path)

		ext := strings.ToLower(filepath.Ext(path))
		filename := strings.TrimSuffix(filepath.Base(path), ext) + ".pdf"
		pdfPath := filepath.Join(savePath, filename)

		switch ext {
		case ".doc", ".docx":
			// 打开文件
			d, err := oleutil.CallMethod(documents, "Open", path)
			if err != nil {
				log.Printf("Error opening document: %v", err)
				continue
			}
			document = d.ToIDispatch()

			// 保存为 PDF
			const wdFormatPDF = 17
			_, err = oleutil.CallMethod(document, "SaveAs", pdfPath, wdFormatPDF)
			if err != nil {
				log.Printf("Error saving document: %v", err)
			}
			oleutil.MustCallMethod(document, "Close", false)
		case ".ppt", ".pptx":
			// 打开文件
			d, err := oleutil.CallMethod(presentations, "Open", path, true, true, false)
			if err != nil {
				log.Printf("Error opening presentation: %v", err)
				continue
			}
			document = d.ToIDispatch()

			// 保存为 PDF
			const ppSaveAsPDF = 32
			_, err = oleutil.CallMethod(document, "SaveAs", pdfPath, ppSaveAsPDF)
			if err != nil {
				log.Printf("Error saving presentation: %v", err)
			}
			oleutil.MustCallMethod(document, "Close")
		}
	}
	defer document.Release()

	oleutil.MustCallMethod(word, "Quit")
	oleutil.MustCallMethod(ppt, "Quit")
}

func main() {
	fmt.Print("Input path: ")
	var rootDir string
	fmt.Scanln(&rootDir)
	fmt.Print("Output path (leave blank to save to the \"PDF\" directory under the input path): ")
	var savePath string
	fmt.Scanln(&savePath)

	var filePath []string
	err := filepath.Walk(rootDir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		if info.IsDir() {
			return nil
		}

		ext := strings.ToLower(filepath.Ext(path))
		switch ext {
		case ".doc", ".docx", ".ppt", ".pptx":
			filePath = append(filePath, path)
		default:
			log.Printf("Unsupported file type: %s", path)
		}
		return nil
	})

	if err != nil {
		log.Fatalf("Error walking the directory: %v", err)
	}

	fmt.Printf("== %d files found ==\n", len(filePath))
	fmt.Println("Converting...")
	if savePath == "" {
		savePath = filepath.Join(rootDir, "PDF")
	}
	if err := os.MkdirAll(savePath, os.ModePerm); err != nil {
		log.Fatalf("Error creating directory: %v", err)
	}
	convertToPDF(filePath, savePath)

	fmt.Println("Press Enter to exit...")
	fmt.Scanln()
}
