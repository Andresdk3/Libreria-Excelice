import ctypes

# Cargamos la librería de Go
lib = ctypes.CDLL("../excel_wrapper.dll")

# Definimos argumentos
lib.OpenExcel.argtypes = [ctypes.c_char_p]
lib.WriteCell.argtypes = [ctypes.c_char_p, ctypes.c_char_p, ctypes.c_char_p]
lib.CopyRange.argtypes = [ctypes.c_char_p, ctypes.c_char_p,
                          ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int]
lib.SaveExcel.argtypes = [ctypes.c_char_p]


# Abrir libro ORIGEN
lib.OpenExcelSrc.argtypes = [ctypes.c_char_p]
lib.OpenExcelSrc.restype = ctypes.c_int

# Abrir libro DESTINO
lib.OpenExcelDst.argtypes = [ctypes.c_char_p]
lib.OpenExcelDst.restype = ctypes.c_int

# Copiar rango entre libros
lib.CopyRangeBetweenBooks.argtypes = [
    ctypes.c_char_p, ctypes.c_char_p,  # srcSheet, dstSheet
    ctypes.c_int, ctypes.c_int,        # startRow, endRow
    ctypes.c_int, ctypes.c_int,        # startCol, endCol
    ctypes.c_int, ctypes.c_int,        # dstStartRow, dstStartCol
    ctypes.c_bool                      # formulas
]
lib.CopyRangeBetweenBooks.restype = ctypes.c_int

lib.CopySheetBetweenBooks.argtypes = [ctypes.c_char_p, ctypes.c_char_p, ctypes.c_bool]
lib.CopySheetBetweenBooks.restype = ctypes.c_int

# Guardar destino
lib.SaveExcelDst.argtypes = [ctypes.c_char_p]
lib.SaveExcelDst.restype = ctypes.c_int

# 1. Abrir archivo
res = lib.OpenExcel(b"demo_salida.xlsx")

if res != 0:
    print("❌ Error: no se pudo abrir el archivo demo.xlsx")
else:    
    # 2. Escribir valores en una hoja
    lib.WriteCell(b"Sheet1", b"A1", b"Nombre")
    lib.WriteCell(b"Sheet1", b"B1", b"Edad")
    lib.WriteCell(b"Sheet1", b"A2", b"Ana")
    lib.WriteCell(b"Sheet1", b"B2", b"25")
    lib.WriteCell(b"Sheet1", b"A3", b"Luis")
    lib.WriteCell(b"Sheet1", b"B12", b"30")

    # 3. Copiar rango de datos con estilos de Sheet1 a Sheet2
    lib.CopyRange(b"Sheet1", b"Sheet2", 5, 7, 1, 2)  # A1:B3

    # 4. Guardar archivo en disco
    lib.SaveExcel(b"demo_salida_nueva.xlsx")

    print("Excel modificado y guardado como demo_salida.xlsx")


# 1. Abrir el libro origen
res = lib.OpenExcelSrc(b"demo_salida.xlsx")
if res != 0:
    print("❌ Error al abrir libro origen")
    exit()
else:
    # 2. Abrir el libro destino (si no existe, se crea nuevo)
    res = lib.OpenExcelDst(b"demo_salida_nueva.xlsx")
    if res != 0:
        print("❌ Error al abrir/crear libro destino")
        exit()
    else:
        # 3. Copiar rango A1:E10 de Sheet1 → Sheet2
        # Copiar A1:E10 de Sheet3 → pegar en Sheet1 comenzando en C5
        res = lib.CopyRangeBetweenBooks(
            b"Sheet3", b"Sheet1",
            1, 10,   # Rango origen filas 1-10
            1, 5,    # Rango origen columnas 1-5 (A-E)
            5, 3,    # Celda inicio destino = fila 5, columna 3 → C5
            True     # Incluir fórmulas
        )

        if res == 0:
            print("✅ Rango copiado correctamente")
        else:
            print("❌ Error al copiar rango:", res)

        # 4. Guardar el libro destino con los cambios
        res = lib.SaveExcelDst(b"demo_salida_nueva.xlsx")
        if res == 0:
            print("✅ Archivo destino guardado")
        else:
            print("❌ Error al guardar destino:", res)

res = lib.CopySheetBetweenBooks(b"Sheet3", b"Sheet_copy", True)
if res == 0:
    print("✅ Hoja copiada correctamente")
else:
    print("❌ Error al copiar hoja:", res)

# Guardar destino
lib.SaveExcelDst(b"demo_salida_nueva.xlsx")