package main

/*
#include <stdlib.h>
#include <stdbool.h>
*/
import "C"

import (
	"sync"

	"github.com/xuri/excelize/v2"
)

var (
	f    *excelize.File
	fSrc *excelize.File
	fDst *excelize.File
	mu   sync.Mutex
)

// ============================================================
// Funciones con un único libro
// ============================================================

//export OpenExcel
func OpenExcel(filename *C.char) C.int {
	mu.Lock()
	defer mu.Unlock()

	var err error
	f, err = excelize.OpenFile(C.GoString(filename))
	if err != nil {
		return -1
	}
	return 0
}

//export WriteCell
func WriteCell(sheet, cell, value *C.char) C.int {
	mu.Lock()
	defer mu.Unlock()

	if f == nil {
		return -1
	}

	sheetName := C.GoString(sheet)
	index, err := f.GetSheetIndex(sheetName)
	if index == -1 || err != nil {
		f.NewSheet(sheetName)
	}

	if err := f.SetCellValue(sheetName, C.GoString(cell), C.GoString(value)); err != nil {
		return -2
	}
	return 0
}

//export CopyRange
func CopyRange(srcSheet *C.char, dstSheet *C.char, startRow, endRow, startCol, endCol C.int) C.int {
	mu.Lock()
	defer mu.Unlock()

	if f == nil {
		return -1
	}

	src := C.GoString(srcSheet)
	dst := C.GoString(dstSheet)

	index, err := f.GetSheetIndex(dst)
	if index == -1 || err != nil {
		f.NewSheet(dst)
	}

	for i := int(startRow); i <= int(endRow); i++ {
		for j := int(startCol); j <= int(endCol); j++ {
			cell, _ := excelize.CoordinatesToCellName(j, i)
			val, _ := f.GetCellValue(src, cell)
			styleID, _ := f.GetCellStyle(src, cell)

			dstCell, _ := excelize.CoordinatesToCellName(j, i)
			f.SetCellValue(dst, dstCell, val)
			if styleID != 0 {
				f.SetCellStyle(dst, dstCell, dstCell, styleID)
			}
		}
	}
	return 0
}

//export SaveExcel
func SaveExcel(filename *C.char) C.int {
	mu.Lock()
	defer mu.Unlock()

	if f == nil {
		return -1
	}
	if err := f.SaveAs(C.GoString(filename)); err != nil {
		return -2
	}
	return 0
}

// ============================================================
// Funciones con dos libros (origen/destino)
// ============================================================

//export OpenExcelSrc
func OpenExcelSrc(filename *C.char) C.int {
	mu.Lock()
	defer mu.Unlock()

	var err error
	fSrc, err = excelize.OpenFile(C.GoString(filename))
	if err != nil {
		return -1
	}
	return 0
}

//export OpenExcelDst
func OpenExcelDst(filename *C.char) C.int {
	mu.Lock()
	defer mu.Unlock()

	var err error
	fDst, err = excelize.OpenFile(C.GoString(filename))
	if err != nil {
		fDst = excelize.NewFile()
	}
	return 0
}

// copiar merges compatible con versiones antiguas
func copyMerges(src, dst *excelize.File, srcSheet, dstSheet string) {
	merges, err := src.GetMergeCells(srcSheet)
	if err != nil {
		return
	}
	for _, m := range merges {
		start, end := m.GetStartAxis(), m.GetEndAxis()
		_ = dst.MergeCell(dstSheet, start, end)
	}
}

//export CopyRangeBetweenBooks
func CopyRangeBetweenBooks(srcSheet, dstSheet *C.char,
	startRow, endRow, startCol, endCol,
	dstStartRow, dstStartCol C.int, formulas C.bool) C.int {

	mu.Lock()
	defer mu.Unlock()

	if fSrc == nil || fDst == nil {
		return -1
	}

	src := C.GoString(srcSheet)
	dst := C.GoString(dstSheet)

	index, err := fDst.GetSheetIndex(dst)
	if index == -1 || err != nil {
		fDst.NewSheet(dst)
	}

	for i := int(startRow); i <= int(endRow); i++ {
		for j := int(startCol); j <= int(endCol); j++ {
			cell, _ := excelize.CoordinatesToCellName(j, i)
			styleID, _ := fSrc.GetCellStyle(src, cell)
			formula, _ := fSrc.GetCellFormula(src, cell)

			dstRow := int(dstStartRow) + (i - int(startRow))
			dstCol := int(dstStartCol) + (j - int(startCol))
			dstCell, _ := excelize.CoordinatesToCellName(dstCol, dstRow)

			if formulas && formula != "" {
				fDst.SetCellFormula(dst, dstCell, formula)
			} else {
				val, _ := fSrc.GetCellValue(src, cell)
				fDst.SetCellValue(dst, dstCell, val)
			}

			if styleID != 0 {
				style, err := fSrc.GetStyle(styleID)
				if err == nil && style != nil {
					newStyleID, _ := fDst.NewStyle(style)
					fDst.SetCellStyle(dst, dstCell, dstCell, newStyleID)
				}
			}
		}
	}

	copyMerges(fSrc, fDst, src, dst)
	return 0
}

//export CopySheetBetweenBooks
func CopySheetBetweenBooks(srcSheet, dstSheet *C.char, formulas C.bool) C.int {
	mu.Lock()
	defer mu.Unlock()

	if fSrc == nil || fDst == nil {
		return -1
	}

	src := C.GoString(srcSheet)
	dst := C.GoString(dstSheet)

	index, err := fDst.GetSheetIndex(dst)
	if index == -1 || err != nil {
		fDst.NewSheet(dst)
	}

	rows, err := fSrc.GetRows(src)
	if err != nil {
		return -2
	}

	for i, row := range rows {
		for j := range row {
			cell, _ := excelize.CoordinatesToCellName(j+1, i+1)
			styleID, _ := fSrc.GetCellStyle(src, cell)
			formula, _ := fSrc.GetCellFormula(src, cell)

			if formulas && formula != "" {
				fDst.SetCellFormula(dst, cell, formula)
			} else {
				val, _ := fSrc.GetCellValue(src, cell)
				fDst.SetCellValue(dst, cell, val)
			}

			if styleID != 0 {
				style, err := fSrc.GetStyle(styleID)
				if err == nil && style != nil {
					newStyleID, _ := fDst.NewStyle(style)
					fDst.SetCellStyle(dst, cell, cell, newStyleID)
				}
			}
		}
	}

	copyMerges(fSrc, fDst, src, dst)
	return 0
}

//export SaveExcelDst
func SaveExcelDst(filename *C.char) C.int {
	mu.Lock()
	defer mu.Unlock()

	if fDst == nil {
		return -1
	}
	if err := fDst.SaveAs(C.GoString(filename)); err != nil {
		return -2
	}
	return 0
}

//export CloseAllExcels
func CloseAllExcels() C.int {
	mu.Lock()
	defer mu.Unlock()

	if f != nil {
		f.Close()
		f = nil
	}
	if fSrc != nil {
		fSrc.Close()
		fSrc = nil
	}
	if fDst != nil {
		fDst.Close()
		fDst = nil
	}
	return 0
}

func main() {}
