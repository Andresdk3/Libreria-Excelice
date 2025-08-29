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
	Archivo_origen *excelize.File
	Archivo_segundario *excelize.File
	mu   sync.Mutex
)

//export Abrir_archivo
func Abrir_archivo(filename *C.char) C.int {
	mu.Lock()
	defer mu.Unlock()

	var err error
	if Archivo_origen == nil {
		Archivo_origen, err = excelize.OpenFile(C.GoString(filename))
		if err != nil {
			return -1
		}
		return 0
	}else{
		Archivo_segundario, err = excelize.OpenFile(C.GoString(filename))
		if err != nil {
			return -1
		}
		return 0
	}
}



//export Escribir_Celda
func Escribir_Celda(sheet, cell, value *C.char) C.int {
	mu.Lock()
	defer mu.Unlock()

	if Archivo_origen == nil {
		return -1
	}

	sheetName := C.GoString(sheet)
	index, err := Archivo_origen.GetSheetIndex(sheetName)
	if index == -1 || err != nil {
		Archivo_origen.NewSheet(sheetName)
	}

	if err := Archivo_origen.SetCellValue(sheetName, C.GoString(cell), C.GoString(value)); err != nil {
		return -2
	}
	return 0
}

//export Guardar_Excel
func Guardar_Excel(filename *C.char) C.int {
	mu.Lock()
	defer mu.Unlock()

	if Archivo_origen == nil {
		return -1
	}
	if err := Archivo_origen.SaveAs(C.GoString(filename)); err != nil {
		return -2
	}
	return 0
}


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
	srcSheet, dstSheet *C.char,
	startRow, endRow, startCol, endCol,
	dstStartRow, dstStartCol C.int,
	formulas C.bool,
	useSecondary C.bool, // true = copiar desde Archivo_segundario a Archivo_origen
	) C.int {

	mu.Lock()
	defer mu.Unlock()

	// Seleccionar libro de origen y destino
	var srcFile, dstFile *excelize.File
	if useSecondary {
		if Archivo_origen == nil || Archivo_segundario == nil {
			return -1
		}
		srcFile = Archivo_segundario
		dstFile = Archivo_origen
	} else {
		if Archivo_origen == nil {
			return -1
		}
		srcFile = Archivo_origen
		dstFile = Archivo_origen
	}

	src := C.GoString(srcSheet)
	dst := C.GoString(dstSheet)

	// Crear hoja destino si no existe
	index, err := dstFile.GetSheetIndex(dst)
	if index == -1 || err != nil {
		dstFile.NewSheet(dst)
	}

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

			// Estilos (recreación)
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

	return 0
}


//export Copiar_hoja
func Copiar_hoja(
	srcSheet, dstSheet *C.char,
	formulas C.bool,
	useSecondary C.bool, // true = desde Archivo_segundario → Archivo_origen
) C.int {
	mu.Lock()
	defer mu.Unlock()

	// Seleccionar origen y destino
	var srcFile, dstFile *excelize.File
	if useSecondary {
		if Archivo_origen == nil || Archivo_segundario == nil {
			return -1
		}
		srcFile = Archivo_segundario
		dstFile = Archivo_origen
	} else {
		if Archivo_origen == nil {
			return -1
		}
		srcFile = Archivo_origen
		dstFile = Archivo_origen
	}

	src := C.GoString(srcSheet)
	dst := C.GoString(dstSheet)

	// Crear hoja destino si no existe
	index, err := dstFile.GetSheetIndex(dst)
	if index == -1 || err != nil {
		dstFile.NewSheet(dst)
	}

	rows, err := srcFile.GetRows(src)
	if err != nil {
		return -2
	}

	for i, row := range rows {
		for j := range row {
			cell, _ := excelize.CoordinatesToCellName(j+1, i+1)
			styleID, _ := srcFile.GetCellStyle(src, cell)
			formula, _ := srcFile.GetCellFormula(src, cell)

			// Copiar fórmulas o valores
			if formulas && formula != "" {
				_ = dstFile.SetCellFormula(dst, cell, formula)
			} else {
				val, _ := srcFile.GetCellValue(src, cell)
				_ = dstFile.SetCellValue(dst, cell, val)
			}

			// Copiar estilos
			if styleID != 0 {
				style, err := srcFile.GetStyle(styleID)
				if err == nil && style != nil {
					newStyleID, _ := dstFile.NewStyle(style)
					_ = dstFile.SetCellStyle(dst, cell, cell, newStyleID)
				}
			}
		}
	}

	// Copiar merges de toda la hoja
	if len(rows) > 0 {
		copyMerges(
			srcFile, dstFile,
			src, dst,
			1, 1, 1, 1,
			len(rows), len(rows[0]),
		)
	}

	return 0
}


//export CloseAllExcels
func CloseAllExcels() C.int {
	mu.Lock()
	defer mu.Unlock()

	if Archivo_origen != nil {
		Archivo_origen.Close()
		Archivo_origen = nil
	}
	if Archivo_segundario != nil {
		Archivo_segundario.Close()
		Archivo_segundario = nil
	}
	return 0
}

func main() {}
