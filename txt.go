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