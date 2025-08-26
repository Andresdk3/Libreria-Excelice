package main

/*
#include <stdlib.h>
*/
import "C"

import (
	"github.com/xuri/excelize/v2"
	"sync"
)

// ExcelHandler almacena el archivo en memoria
var (
	f      *excelize.File
	mu     sync.Mutex // control concurrente si se llama desde varios hilos
)

//export OpenExcel
func OpenExcel(filename *C.char) C.int {
	mu.Lock()
	defer mu.Unlock()

	var err error
	f, err = excelize.OpenFile(C.GoString(filename))
	if err != nil {
		// Si no existe, creamos uno nuevo
		f = excelize.NewFile()
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
	err := f.SetCellValue(C.GoString(sheet), C.GoString(cell), C.GoString(value))
	if err != nil {
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

	for i := int(startRow); i <= int(endRow); i++ {
		for j := int(startCol); j <= int(endCol); j++ {
			cell, _ := excelize.CoordinatesToCellName(j, i)
			val, _ := f.GetCellValue(src, cell)
			styleID, _ := f.GetCellStyle(src, cell)

			// Copiar a la misma posición pero en otra hoja
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

func main() {}
