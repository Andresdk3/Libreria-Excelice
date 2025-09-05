package main

/*
#include <stdlib.h>
#include <stdbool.h>
*/
import "C"

import (
	"encoding/json"
	"fmt"
	"strconv"
	"sync"
	"unsafe"

	"github.com/xuri/excelize/v2"
)

var (
	mu     sync.RWMutex // Cambiado a RWMutex para lecturas concurrentes
	libros = make(map[int]*excelize.File)
	nextID = 1
	idPool = sync.Pool{ // Pool para reutilizar IDs
		New: func() interface{} {
			return new(int)
		},
	}
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
	mu.RLock()
	file, ok := libros[int(id)]
	mu.RUnlock()

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
	defer mu.Unlock()

	for id, file := range libros {
		file.Close()
		delete(libros, id)
	}
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
	mu.RLock()
	file, ok := libros[int(id)]
	mu.RUnlock()

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
	mu.RLock()
	file, ok := libros[int(id)]
	mu.RUnlock()

	if !ok {
		return -1
	}

	valStr := C.GoString(value)
	sheetStr := C.GoString(sheet)
	cellStr := C.GoString(cell)

	// Si comienza con "=" → escribir como fórmula
	if len(valStr) > 0 && valStr[0] == '=' {
		if err := file.SetCellFormula(sheetStr, cellStr, valStr); err != nil {
			return -4
		}
		return 0
	}

	// Intentar parsear como número
	if i, err := strconv.ParseInt(valStr, 10, 64); err == nil {
		if err := file.SetCellValue(sheetStr, cellStr, i); err != nil {
			return -3
		}
		return 0
	}

	// Intentar parsear como número decimal
	if f, err := strconv.ParseFloat(valStr, 64); err == nil {
		if err := file.SetCellValue(sheetStr, cellStr, f); err != nil {
			return -2
		}
		return 0
	}

	// Si no es número → escribir como texto
	if err := file.SetCellValue(sheetStr, cellStr, valStr); err != nil {
		return -1
	}

	return 0
}

//export Descombinar_Rango
func Descombinar_Rango(id C.int, sheet *C.char, start *C.char, end *C.char) C.int {
	mu.RLock()
	file, ok := libros[int(id)]
	mu.RUnlock()

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
	mu.RLock()
	file, ok := libros[int(id)]
	mu.RUnlock()

	if !ok {
		return C.CString(`{"error":"archivo no encontrado"}`)
	}

	sheets := file.GetSheetList()
	data, _ := json.Marshal(sheets)
	return C.CString(string(data))
}

// =====================
// Copiar datos (optimizado)
// =====================

// copiarMerges optimizado
func copiarMerges(srcFile, dstFile *excelize.File, srcSheet, dstSheet string,
	startRow, startCol, dstStartRow, dstStartCol, endRow, endCol int) {

	merges, err := srcFile.GetMergeCells(srcSheet)
	if err != nil {
		return
	}

	for _, m := range merges {
		c1, r1, _ := excelize.CellNameToCoordinates(m.GetStartAxis())
		c2, r2, _ := excelize.CellNameToCoordinates(m.GetEndAxis())

		// Verificar si el merge está dentro del rango
		if r1 < startRow || r2 > endRow || c1 < startCol || c2 > endCol {
			continue
		}

		rowOffset := dstStartRow - startRow
		colOffset := dstStartCol - startCol

		newStart, _ := excelize.CoordinatesToCellName(c1+colOffset, r1+rowOffset)
		newEnd, _ := excelize.CoordinatesToCellName(c2+colOffset, r2+rowOffset)

		dstFile.MergeCell(dstSheet, newStart, newEnd)
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

	mu.RLock()
	srcFile, ok1 := libros[int(srcID)]
	dstFile, ok2 := libros[int(dstID)]
	mu.RUnlock()

	if !ok1 || !ok2 {
		return -1
	}

	src := C.GoString(srcSheet)
	dst := C.GoString(dstSheet)

	// Verificar y crear hoja destino si no existe
	index, err := dstFile.GetSheetIndex(dst)
	if index == -1 || err != nil {
		dstFile.NewSheet(dst)
	}

	// Precalcular offsets
	rowOffset := int(dstStartRow) - int(startRow)
	colOffset := int(dstStartCol) - int(startCol)

	// Buffer para nombres de celdas
	cellCache := make(map[[2]int]string)

	// Copiar celdas
	for i := int(startRow); i <= int(endRow); i++ {
		for j := int(startCol); j <= int(endCol); j++ {
			// Usar cache para nombres de celdas
			srcKey := [2]int{j, i}
			if _, exists := cellCache[srcKey]; !exists {
				cellCache[srcKey], _ = excelize.CoordinatesToCellName(j, i)
			}
			srcCell := cellCache[srcKey]

			dstRow := i + rowOffset
			dstCol := j + colOffset
			dstKey := [2]int{dstCol, dstRow}
			if _, exists := cellCache[dstKey]; !exists {
				cellCache[dstKey], _ = excelize.CoordinatesToCellName(dstCol, dstRow)
			}
			dstCell := cellCache[dstKey]

			// Obtener estilo y fórmula
			styleID, _ := srcFile.GetCellStyle(src, srcCell)
			formula, _ := srcFile.GetCellFormula(src, srcCell)

			// Copiar fórmula o valor
			if formulas && formula != "" {
				dstFile.SetCellFormula(dst, dstCell, formula)
			} else {
				val, _ := srcFile.GetCellValue(src, srcCell)
				// Determinar tipo de dato eficientemente
				switch {
				case val == "":
					dstFile.SetCellValue(dst, dstCell, "")
				case isInteger(val):
					if i, err := strconv.ParseInt(val, 10, 64); err == nil {
						dstFile.SetCellValue(dst, dstCell, i)
					}
				case isFloat(val):
					if f, err := strconv.ParseFloat(val, 64); err == nil {
						dstFile.SetCellValue(dst, dstCell, f)
					}
				default:
					dstFile.SetCellValue(dst, dstCell, val)
				}
			}

			// Copiar estilo si existe
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
	copiarMerges(srcFile, dstFile, src, dst,
		int(startRow), int(startCol),
		int(dstStartRow), int(dstStartCol),
		int(endRow), int(endCol))

	// Copiar anchos de columna
	for j := int(startCol); j <= int(endCol); j++ {
		srcCol, _ := excelize.ColumnNumberToName(j)
		dstCol, _ := excelize.ColumnNumberToName(j + colOffset)
		if width, err := srcFile.GetColWidth(src, srcCol); err == nil && width > 0 {
			dstFile.SetColWidth(dst, dstCol, dstCol, width)
		}
	}

	// Copiar alturas de fila
	for i := int(startRow); i <= int(endRow); i++ {
		if height, err := srcFile.GetRowHeight(src, i); err == nil && height > 0 {
			dstFile.SetRowHeight(dst, i+rowOffset, height)
		}
	}

	return 0
}

// Funciones auxiliares para verificación rápida de tipos numéricos
func isInteger(s string) bool {
	for _, c := range s {
		if c < '0' || c > '9' {
			return false
		}
	}
	return s != "" && s != "0"
}

func isFloat(s string) bool {
	dotFound := false
	for i, c := range s {
		if c == '.' && !dotFound {
			dotFound = true
		} else if c < '0' || c > '9' {
			return false
		}
		if i == 0 && c == '-' {
			continue
		}
	}
	return dotFound
}

//export Copiar_hoja
func Copiar_hoja(srcID C.int, dstID C.int, srcSheet *C.char, dstSheet *C.char, formulas bool) C.int {
	mu.RLock()
	srcFile, ok1 := libros[int(srcID)]
	dstFile, ok2 := libros[int(dstID)]
	mu.RUnlock()

	if !ok1 || !ok2 {
		return -1
	}

	srcStr := C.GoString(srcSheet)
	dstStr := C.GoString(dstSheet)

	dstFile.NewSheet(dstStr)

	rows, err := srcFile.GetRows(srcStr)
	if err != nil {
		return -3
	}

	// Preallocar buffer para nombres de celdas
	cellNameCache := make(map[[2]int]string)

	for r, row := range rows {
		for c, val := range row {
			// Cachear nombres de celdas
			key := [2]int{c + 1, r + 1}
			if cell, exists := cellNameCache[key]; exists {
				if formulas {
					if formula, _ := srcFile.GetCellFormula(srcStr, cell); formula != "" {
						dstFile.SetCellFormula(dstStr, cell, formula)
						continue
					}
				}
				dstFile.SetCellValue(dstStr, cell, val)
			} else {
				cellName, _ := excelize.CoordinatesToCellName(c+1, r+1)
				cellNameCache[key] = cellName
				if formulas {
					if formula, _ := srcFile.GetCellFormula(srcStr, cellName); formula != "" {
						dstFile.SetCellFormula(dstStr, cellName, formula)
						continue
					}
				}
				dstFile.SetCellValue(dstStr, cellName, val)
			}
		}
	}
	return 0
}

//export Copiar_hoja_completa
func Copiar_hoja_completa(srcID, dstID C.int, srcSheet *C.char, dstSheet *C.char) C.int {
	mu.RLock()
	srcFile, ok1 := libros[int(srcID)]
	dstFile, ok2 := libros[int(dstID)]
	mu.RUnlock()

	if !ok1 || !ok2 {
		return -1
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
	mu.RLock()
	file, ok := libros[int(id)]
	mu.RUnlock()

	if !ok {
		return -1
	}

	err := file.RemoveRow(C.GoString(sheetName), int(fila))
	if err != nil {
		return -2
	}
	return 0
}

func main() {}
