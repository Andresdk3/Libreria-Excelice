package main

/*
#include <stdlib.h>
#include <stdbool.h>
*/
import "C"

import (
	"encoding/json"
	"fmt"
	"sort"
	"strconv"
	"sync"
	"sync/atomic"
	"unsafe"

	"github.com/xuri/excelize/v2"
)

var (
	// Mapa concurrente para libros
	libros = &sync.Map{}
	// Contador atómico para IDs
	nextID int64 = 1
	// Pool de workers para procesamiento concurrente
	workerPool = NewWorkerPool(8) // 8 workers por defecto
)

// WorkerPool gestiona un pool de goroutines para procesamiento concurrente
type WorkerPool struct {
	workers   int
	taskQueue chan func()
	wg        sync.WaitGroup
}

// NewWorkerPool crea un nuevo pool de workers
func NewWorkerPool(workers int) *WorkerPool {
	pool := &WorkerPool{
		workers:   workers,
		taskQueue: make(chan func(), workers*2),
	}

	// Iniciar workers
	for i := 0; i < workers; i++ {
		pool.wg.Add(1)
		go pool.worker()
	}

	return pool
}

// worker procesa tareas del queue
func (p *WorkerPool) worker() {
	defer p.wg.Done()
	for task := range p.taskQueue {
		task()
	}
}

// Submit añade una tarea al pool
func (p *WorkerPool) Submit(task func()) {
	p.taskQueue <- task
}

// Close cierra el pool
func (p *WorkerPool) Close() {
	close(p.taskQueue)
	p.wg.Wait()
}

// SetWorkerPoolSize permite ajustar el tamaño del pool dinámicamente
//export SetWorkerPoolSize
func SetWorkerPoolSize(size C.int) {
	workerPool.Close()
	workerPool = NewWorkerPool(int(size))
}

// =====================
// Gestión de archivos
// =====================

//export Abrir_archivo
func Abrir_archivo(filename *C.char) C.int {
	f, err := excelize.OpenFile(C.GoString(filename))
	if err != nil {
		return -1
	}
	
	id := atomic.AddInt64(&nextID, 1)
	libros.Store(id, f)
	return C.int(id)
}

//export Guardar_Excel
func Guardar_Excel(id C.int, filename *C.char) C.int {
	file, ok := libros.Load(int64(id))
	if !ok {
		return -1
	}
	
	if err := file.(*excelize.File).SaveAs(C.GoString(filename)); err != nil {
		return -2
	}
	return 0
}

//export Cerrar_archivo
func Cerrar_archivo(id C.int) C.int {
	file, ok := libros.LoadAndDelete(int64(id))
	if !ok {
		return -1
	}
	
	file.(*excelize.File).Close()
	return 0
}

//export CloseAllExcels
func CloseAllExcels() C.int {
	libros.Range(func(key, value interface{}) bool {
		value.(*excelize.File).Close()
		libros.Delete(key)
		return true
	})
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
	file, ok := libros.Load(int64(id))
	if !ok {
		return C.CString(`{"error": "archivo no encontrado"}`)
	}

	rows, err := file.(*excelize.File).GetRows(C.GoString(sheet))
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
	file, ok := libros.Load(int64(id))
	if !ok {
		return -1
	}

	f := file.(*excelize.File)
	valStr := C.GoString(value)
	sheetStr := C.GoString(sheet)
	cellStr := C.GoString(cell)

	// Si comienza con "=" → escribir como fórmula
	if len(valStr) > 0 && valStr[0] == '=' {
		if err := f.SetCellFormula(sheetStr, cellStr, valStr); err != nil {
			return -4
		}
		return 0
	}

	// Intentar parsear como número
	if i, err := strconv.ParseInt(valStr, 10, 64); err == nil {
		if err := f.SetCellValue(sheetStr, cellStr, i); err != nil {
			return -3
		}
		return 0
	}

	// Intentar parsear como número decimal
	if fVal, err := strconv.ParseFloat(valStr, 64); err == nil {
		if err := f.SetCellValue(sheetStr, cellStr, fVal); err != nil {
			return -2
		}
		return 0
	}

	// Si no es número → escribir como texto
	if err := f.SetCellValue(sheetStr, cellStr, valStr); err != nil {
		return -1
	}

	return 0
}

//export Descombinar_Rango
func Descombinar_Rango(id C.int, sheet *C.char, start *C.char, end *C.char) C.int {
	file, ok := libros.Load(int64(id))
	if !ok {
		return -1
	}

	if err := file.(*excelize.File).UnmergeCell(C.GoString(sheet), C.GoString(start), C.GoString(end)); err != nil {
		return -2
	}
	return 0
}

//export Listar_Hojas
func Listar_Hojas(id C.int) *C.char {
	file, ok := libros.Load(int64(id))
	if !ok {
		return C.CString(`{"error":"archivo no encontrado"}`)
	}

	sheets := file.(*excelize.File).GetSheetList()
	data, _ := json.Marshal(sheets)
	return C.CString(string(data))
}

// =====================
// Copiar datos (optimizado con pool de workers)
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

	srcFile, ok1 := libros.Load(int64(srcID))
	dstFile, ok2 := libros.Load(int64(dstID))
	if !ok1 || !ok2 {
		return -1
	}

	src := C.GoString(srcSheet)
	dst := C.GoString(dstSheet)

	// Verificar y crear hoja destino si no existe
	index, err := dstFile.(*excelize.File).GetSheetIndex(dst)
	if index == -1 || err != nil {
		dstFile.(*excelize.File).NewSheet(dst)
	}

	// Precalcular offsets
	rowOffset := int(dstStartRow) - int(startRow)
	colOffset := int(dstStartCol) - int(startCol)

	// Usar waitgroup para procesamiento con pool
	var wg sync.WaitGroup
	errorChan := make(chan error, 1)
	var hasError atomic.Bool

	// Copiar celdas usando el pool de workers
	for i := int(startRow); i <= int(endRow); i++ {
		wg.Add(1)
		row := i // Capturar variable para la goroutine
		
		workerPool.Submit(func() {
			defer wg.Done()
			
			if hasError.Load() {
				return // Si hay error, salir
			}
			
			for j := int(startCol); j <= int(endCol); j++ {
				if hasError.Load() {
					return // Si hay error, salir
				}
				
				srcCell, _ := excelize.CoordinatesToCellName(j, row)
				dstRow := row + rowOffset
				dstCol := j + colOffset
				dstCell, _ := excelize.CoordinatesToCellName(dstCol, dstRow)

				// Obtener estilo y fórmula
				styleID, _ := srcFile.(*excelize.File).GetCellStyle(src, srcCell)
				formula, _ := srcFile.(*excelize.File).GetCellFormula(src, srcCell)

				// Copiar fórmula o valor
				if formulas && formula != "" {
					if err := dstFile.(*excelize.File).SetCellFormula(dst, dstCell, formula); err != nil {
						if !hasError.Swap(true) {
							errorChan <- err
						}
						return
					}
				} else {
					val, _ := srcFile.(*excelize.File).GetCellValue(src, srcCell)
					// Determinar tipo de dato eficientemente
					switch {
					case val == "":
						dstFile.(*excelize.File).SetCellValue(dst, dstCell, "")
					case isInteger(val):
						if iVal, err := strconv.ParseInt(val, 10, 64); err == nil {
							dstFile.(*excelize.File).SetCellValue(dst, dstCell, iVal)
						} else {
							dstFile.(*excelize.File).SetCellValue(dst, dstCell, val)
						}
					case isFloat(val):
						if fVal, err := strconv.ParseFloat(val, 64); err == nil {
							dstFile.(*excelize.File).SetCellValue(dst, dstCell, fVal)
						} else {
							dstFile.(*excelize.File).SetCellValue(dst, dstCell, val)
						}
					default:
						dstFile.(*excelize.File).SetCellValue(dst, dstCell, val)
					}
				}

				// Copiar estilo si existe
				if styleID != 0 {
					style, err := srcFile.(*excelize.File).GetStyle(styleID)
					if err == nil && style != nil {
						newStyleID, _ := dstFile.(*excelize.File).NewStyle(style)
						dstFile.(*excelize.File).SetCellStyle(dst, dstCell, dstCell, newStyleID)
					}
				}
			}
		})
	}

	wg.Wait()
	close(errorChan)
	
	// Verificar si hubo errores
	if hasError.Load() {
		return -2
	}

	// Copiar merges (no necesita concurrencia)
	copiarMerges(srcFile.(*excelize.File), dstFile.(*excelize.File), src, dst,
		int(startRow), int(startCol),
		int(dstStartRow), int(dstStartCol),
		int(endRow), int(endCol))

	// Copiar anchos de columna (no necesita concurrencia)
	for j := int(startCol); j <= int(endCol); j++ {
		srcCol, _ := excelize.ColumnNumberToName(j)
		dstCol, _ := excelize.ColumnNumberToName(j + colOffset)
		if width, err := srcFile.(*excelize.File).GetColWidth(src, srcCol); err == nil && width > 0 {
			dstFile.(*excelize.File).SetColWidth(dst, dstCol, dstCol, width)
		}
	}

	// Copiar alturas de fila (no necesita concurrencia)
	for i := int(startRow); i <= int(endRow); i++ {
		if height, err := srcFile.(*excelize.File).GetRowHeight(src, i); err == nil && height > 0 {
			dstFile.(*excelize.File).SetRowHeight(dst, i+rowOffset, height)
		}
	}

	return 0
}

// Funciones auxiliares para verificación rápida de tipos numéricos
func isInteger(s string) bool {
	if s == "" {
		return false
	}
	for i, c := range s {
		if i == 0 && c == '-' {
			continue
		}
		if c < '0' || c > '9' {
			return false
		}
	}
	return true
}

func isFloat(s string) bool {
	if s == "" {
		return false
	}
	dotFound := false
	for i, c := range s {
		if i == 0 && c == '-' {
			continue
		}
		if c == '.' && !dotFound {
			dotFound = true
		} else if c < '0' || c > '9' {
			return false
		}
	}
	return dotFound
}

//export Eliminar_Filas_Array
func Eliminar_Filas_Array(id C.int, sheetName *C.char, filas *C.int, count C.int) C.int {
	file, ok := libros.Load(int64(id))
	if !ok {
		return -1 // libro no encontrado
	}

	sheet := C.GoString(sheetName)
	length := int(count)

	if length == 0 {
		return 0 // nada que eliminar
	}

	// Convertir puntero C a slice de Go
	rows := (*[1 << 30]C.int)(unsafe.Pointer(filas))[:length:length]

	// Para evitar problemas con el corrimiento, ordenamos en orden descendente
	intRows := make([]int, length)
	for i, r := range rows {
		intRows[i] = int(r)
	}
	sort.Sort(sort.Reverse(sort.IntSlice(intRows)))

	// Eliminar cada fila
	for _, row := range intRows {
		if row <= 0 {
			continue
		}
		if err := file.(*excelize.File).RemoveRow(sheet, row); err != nil {
			return -2
		}
	}
	return 0
}


//export Copiar_hoja_completa
func Copiar_hoja_completa(srcID, dstID C.int, srcSheet *C.char, dstSheet *C.char) C.int {
	srcFile, ok1 := libros.Load(int64(srcID))
	dstFile, ok2 := libros.Load(int64(dstID))
	if !ok1 || !ok2 {
		return -1
	}

	srcIdx, err := srcFile.(*excelize.File).GetSheetIndex(C.GoString(srcSheet))
	if err != nil {
		return -3
	}

	dstIdx, _ := dstFile.(*excelize.File).NewSheet(C.GoString(dstSheet))
	if err := dstFile.(*excelize.File).CopySheet(srcIdx, dstIdx); err != nil {
		return -4
	}
	return 0
}

//export Eliminar_Fila
func Eliminar_Fila(id C.int, sheetName *C.char, fila C.int) C.int {
	file, ok := libros.Load(int64(id))
	if !ok {
		return -1
	}

	err := file.(*excelize.File).RemoveRow(C.GoString(sheetName), int(fila))
	if err != nil {
		return -2
	}
	return 0
}

func main() {
	// Aseguramos que el pool se cierre al terminar
	defer workerPool.Close()
}