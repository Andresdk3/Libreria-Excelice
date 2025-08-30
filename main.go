package main

/*
#include <stdlib.h>
#include <stdbool.h>
*/
import "C"

import (
	"encoding/json"
	"fmt"
	"sync"
	"unsafe"

	"github.com/xuri/excelize/v2"
)

var (
	mu     sync.Mutex
	libros = make(map[int]*excelize.File) // id → *excelize.File
	nextID = 1
)

// =====================
// Gestión de archivos
// =====================

//export Abrir_archivo
func Abrir_archivo(filename *C.char) C.int {
	mu.Lock()
	f, err := excelize.OpenFile(C.GoString(filename))
	if err != nil {
		mu.Unlock()
		return -1
	}
	id := nextID
	libros[id] = f
	nextID++
	mu.Unlock()
	return C.int(id)
}

//export Guardar_Excel
func Guardar_Excel(id C.int, filename *C.char) C.int {
	mu.Lock()
	file, ok := libros[int(id)]
	mu.Unlock()
	if !ok {
		return -1
	}
	if err := file.SaveAs(C.GoString(filename)); err != nil {
		return -2
	}
	return 0
}

//export Cerrar_archivo
func Cerrar_archivo(id C.int) C.int {
	mu.Lock()
	file, ok := libros[int(id)]
	if ok {
		delete(libros, int(id))
	}
	mu.Unlock()
	if !ok {
		return -1
	}
	file.Close()
	return 0
}

//export CloseAllExcels
func CloseAllExcels() C.int {
	mu.Lock()
	for id, file := range libros {
		file.Close()
		delete(libros, id)
	}
	mu.Unlock()
	return 0
}

// =====================
// Manejo de memoria C
// =====================

//export FreeString
func FreeString(str *C.char) {
	C.free(unsafe.Pointer(str))
}

// =====================
// Operaciones con hojas
// =====================

//export Leer_Hoja
func Leer_Hoja(id C.int, sheet *C.char) *C.char {
	mu.Lock()
	file, ok := libros[int(id)]
	mu.Unlock()
	if !ok {
		return C.CString(`{"error": "archivo no encontrado"}`)
	}

	rows, err := file.GetRows(C.GoString(sheet))
	if err != nil {
		return C.CString(fmt.Sprintf(`{"error": "%v"}`, err))
	}

	jsonData, err := json.Marshal(rows)
	if err != nil {
		return C.CString(fmt.Sprintf(`{"error": "%v"}`, err))
	}
	return C.CString(string(jsonData))
}



//export Escribir_Celda
func Escribir_Celda(id C.int, sheet *C.char, cell *C.char, value *C.char) C.int {
	mu.Lock()
	file, ok := libros[int(id)]
	mu.Unlock()
	if !ok {
		return -1
	}

	valStr := C.GoString(value)
	
	if len(valStr) > 0 && valStr[0] == '=' {
		if err := file.SetCellFormula(C.GoString(sheet), C.GoString(cell), C.GoString(value)); err != nil {
			return -2
		}
	} else {
		if err := file.SetCellValue(C.GoString(sheet), C.GoString(cell), C.GoString(value)); err != nil {
			return -2
		}
	}
	return 0
}

//export Descombinar_Rango
func Descombinar_Rango(id C.int, sheet *C.char, start *C.char, end *C.char) C.int {
	mu.Lock()
	file, ok := libros[int(id)]
	mu.Unlock()
	if !ok {
		return -1
	}
	if err := file.UnmergeCell(C.GoString(sheet), C.GoString(start), C.GoString(end)); err != nil {
		return -2
	}
	return 0
}

//export Listar_Hojas
func Listar_Hojas(id C.int) *C.char {
	mu.Lock()
	file, ok := libros[int(id)]
	mu.Unlock()
	if !ok {
		return C.CString(`{"error":"archivo no encontrado"}`)
	}
	sheets := file.GetSheetList()
	data, _ := json.Marshal(sheets)
	return C.CString(string(data))
}

// =====================
// Copiar datos
// =====================

// copiar merges compatible con versiones antiguas
func copyMerges(srcFile, dstFile *excelize.File, srcSheet, dstSheet string,
	startRow, startCol, dstStartRow, dstStartCol, endRow, endCol int) {

	merges, err := srcFile.GetMergeCells(srcSheet)
	if err != nil {
		return
	}

	for _, m := range merges {
		// Convertir rango de merge a coordenadas
		c1, r1, _ := excelize.CellNameToCoordinates(m.GetStartAxis())
		c2, r2, _ := excelize.CellNameToCoordinates(m.GetEndAxis())

		// Validar si el merge está dentro del rango a copiar
		if r1 < startRow || r2 > endRow || c1 < startCol || c2 > endCol {
			continue
		}

		// Calcular el offset
		rowOffset := dstStartRow - startRow
		colOffset := dstStartCol - startCol

		// Aplicar desplazamiento al rango destino
		newC1 := c1 + colOffset
		newR1 := r1 + rowOffset
		newC2 := c2 + colOffset
		newR2 := r2 + rowOffset

		newStart, _ := excelize.CoordinatesToCellName(newC1, newR1)
		newEnd, _ := excelize.CoordinatesToCellName(newC2, newR2)

		// Combinar en archivo destino
		_ = dstFile.MergeCell(dstSheet, newStart, newEnd)
	}
}

//export Copiar_rango
func Copiar_rango(
		srcID C.int, dstID C.int,
		srcSheet *C.char, dstSheet *C.char,
		startRow, endRow, startCol, endCol C.int,
		dstStartRow, dstStartCol C.int,
		formulas bool,
	) C.int {

	mu.Lock()
	srcFile, ok1 := libros[int(srcID)]
	dstFile, ok2 := libros[int(dstID)]
	mu.Unlock()
	if !ok1 {
		return -1
	}
	if !ok2 {
		return -2
	}

	src := C.GoString(srcSheet)
	dst := C.GoString(dstSheet)
	// Crear hoja destino si no existe
	index, err := dstFile.GetSheetIndex(dst)
	if index == -1 || err != nil {
		dstFile.NewSheet(dst)
	}

	// Copiar celdas (valores, fórmulas y estilos)
	for i := int(startRow); i <= int(endRow); i++ {
		for j := int(startCol); j <= int(endCol); j++ {
			cell, _ := excelize.CoordinatesToCellName(j, i)
			styleID, _ := srcFile.GetCellStyle(src, cell)
			formula, _ := srcFile.GetCellFormula(src, cell)

			// Ajustar destino
			dstRow := int(dstStartRow) + (i - int(startRow))
			dstCol := int(dstStartCol) + (j - int(startCol))
			dstCell, _ := excelize.CoordinatesToCellName(dstCol, dstRow)

			// Fórmulas o valores
			if formulas && formula != "" {
				dstFile.SetCellFormula(dst, dstCell, formula)
			} else {
				val, _ := srcFile.GetCellValue(src, cell)
				dstFile.SetCellValue(dst, dstCell, val)
			}

			// Estilos
			if styleID != 0 {
				style, err := srcFile.GetStyle(styleID)
				if err == nil && style != nil {
					newStyleID, _ := dstFile.NewStyle(style)
					dstFile.SetCellStyle(dst, dstCell, dstCell, newStyleID)
				}
			}
		}
	}

	// Copiar merges
	copyMerges(srcFile, dstFile, src, dst,
		int(startRow), int(startCol),
		int(dstStartRow), int(dstStartCol),
		int(endRow), int(endCol))

	// 📏 Copiar anchos de columna
	for j := int(startCol); j <= int(endCol); j++ {
		srcCol, _ := excelize.ColumnNumberToName(j)
		dstCol, _ := excelize.ColumnNumberToName(int(dstStartCol) + (j - int(startCol)))
		width, err := srcFile.GetColWidth(src, srcCol)
		if err == nil && width > 0 {
			_ = dstFile.SetColWidth(dst, dstCol, dstCol, width)
		}
	}

	// 📐 Copiar alturas de fila
	for i := int(startRow); i <= int(endRow); i++ {
		height, err := srcFile.GetRowHeight(src, i)
		if err == nil && height > 0 {
			dstRow := int(dstStartRow) + (i - int(startRow))
			_ = dstFile.SetRowHeight(dst, dstRow, height)
		}
	}
	return 0
}

//export Copiar_hoja
func Copiar_hoja(srcID C.int, dstID C.int, srcSheet *C.char, dstSheet *C.char, formulas bool) C.int {
	mu.Lock()
	srcFile, ok1 := libros[int(srcID)]
	dstFile, ok2 := libros[int(dstID)]
	mu.Unlock()
	if !ok1 {
		return -1
	}
	if !ok2 {
		return -2
	}

	dstFile.NewSheet(C.GoString(dstSheet))

	rows, err := srcFile.GetRows(C.GoString(srcSheet))
	if err != nil {
		return -3
	}

	for r, row := range rows {
		for c, val := range row {
			cell, _ := excelize.CoordinatesToCellName(c+1, r+1)
			if formulas {
				formula, _ := srcFile.GetCellFormula(C.GoString(srcSheet), cell)
				if formula != "" {
					dstFile.SetCellFormula(C.GoString(dstSheet), cell, formula)
				} else {
					dstFile.SetCellValue(C.GoString(dstSheet), cell, val)
				}
			} else {
				dstFile.SetCellValue(C.GoString(dstSheet), cell, val)
			}
		}
	}
	return 0
}

//export Copiar_hoja_completa
func Copiar_hoja_completa(srcID, dstID C.int, srcSheet *C.char, dstSheet *C.char) C.int {
	mu.Lock()
	srcFile, ok1 := libros[int(srcID)]
	dstFile, ok2 := libros[int(dstID)]
	mu.Unlock()
	if !ok1 {
		return -1
	}
	if !ok2 {
		return -2
	}

	srcIdx, err := srcFile.GetSheetIndex(C.GoString(srcSheet))
	if err != nil {
		return -3
	}
	dstIdx, _ := dstFile.NewSheet(C.GoString(dstSheet))
	if err := dstFile.CopySheet(srcIdx, dstIdx); err != nil {
		return -4
	}
	return 0
}

//export Eliminar_Fila
func Eliminar_Fila(id C.int, sheetName *C.char, fila C.int) C.int {
	mu.Lock()
	defer mu.Unlock()

	file, ok := libros[int(id)]
	if !ok {
		return -1
	}

	sheet := C.GoString(sheetName)

	// Excelize usa índice base 1 para filas
	err := file.RemoveRow(sheet, int(fila))
	if err != nil {
		return -2
	}
	return 0
}


func main() {}
