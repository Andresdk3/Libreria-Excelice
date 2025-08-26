import ctypes

# Cargamos la librería de Go
lib = ctypes.CDLL("./excel_wrapper.dll")

# Definimos argumentos
lib.OpenExcel.argtypes = [ctypes.c_char_p]
lib.WriteCell.argtypes = [ctypes.c_char_p, ctypes.c_char_p, ctypes.c_char_p]
lib.CopyRange.argtypes = [ctypes.c_char_p, ctypes.c_char_p,
                          ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int]
lib.SaveExcel.argtypes = [ctypes.c_char_p]

# 1. Abrir archivo (si no existe, se crea en memoria)
res = lib.OpenExcel(b"demo_salida.xlsx")

if res != 0:
    print("❌ Error: no se pudo abrir el archivo demo.xlsx")
else:
    print("✅ Archivo abierto correctamente")

    # 2. Escribir valores en una hoja
    lib.WriteCell(b"Sheet1", b"A1", b"Nombre")
    lib.WriteCell(b"Sheet1", b"B1", b"Edad")
    lib.WriteCell(b"Sheet1", b"A2", b"Ana")
    lib.WriteCell(b"Sheet1", b"B2", b"25")
    lib.WriteCell(b"Sheet1", b"A3", b"Luis")
    lib.WriteCell(b"Sheet1", b"B10", b"30")

    # 3. Copiar rango de datos con estilos de Sheet1 a Sheet2
    lib.CopyRange(b"Sheet1", b"Sheet2", 5, 7, 1, 2)  # A1:B3

    # 4. Guardar archivo en disco
    lib.SaveExcel(b"demo_salida_nueva.xlsx")

    print("Excel modificado y guardado como demo_salida.xlsx")
