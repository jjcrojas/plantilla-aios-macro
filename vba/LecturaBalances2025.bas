Attribute VB_Name = "LecturaBalances2025"

Option Explicit

' ============================
'  LecturaBalances2025 (helpers)
' ============================

Private Function ArchivoExiste(ByVal ruta As String) As Boolean
    ArchivoExiste = (Dir(ruta) <> "")
End Function

Private Function AbrirLibroSeguro(ByVal ruta As String) As Workbook
    If Not ArchivoExiste(ruta) Then
        MsgBox "No existe el archivo: " & ruta, vbCritical
        Set AbrirLibroSeguro = Nothing
        Exit Function
    End If
    Set AbrirLibroSeguro = Workbooks.Open(Filename:=ruta, ReadOnly:=True)
End Function

Private Function HojaExiste(wb As Workbook, ByVal nombreHoja As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nombreHoja)
    HojaExiste = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function ObtenerHoja(wb As Workbook, ByVal nombreHoja As String) As Worksheet
    If wb Is Nothing Then
        Set ObtenerHoja = Nothing
        Exit Function
    End If
    If Not HojaExiste(wb, nombreHoja) Then
        MsgBox "La hoja '" & nombreHoja & "' no existe en " & wb.Name, vbCritical
        Set ObtenerHoja = Nothing
        Exit Function
    End If
    Set ObtenerHoja = wb.Worksheets(nombreHoja)
End Function

' Normaliza para comparar textos: minúsculas, trim y sin acentos básicos
Private Function Normaliza(ByVal s As String) As String
    Dim t As String
    t = LCase$(Trim$(s))
    t = Replace(t, "á", "a")
    t = Replace(t, "é", "e")
    t = Replace(t, "í", "i")
    t = Replace(t, "ó", "o")
    t = Replace(t, "ú", "u")
    t = Replace(t, "  ", " ")
    Normaliza = t
End Function

Private Function NzNum(v As Variant) As Double
    If IsNumeric(v) Then NzNum = CDbl(v) Else NzNum = 0
End Function

' Busca una celda por texto (case-insensitive). Si exacto falla, intenta por coincidencia parcial.
Private Function FindHeader(ws As Worksheet, ByVal texto As String) As Range
    Dim r As Range
    Set r = ws.Cells.Find(What:=texto, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If r Is Nothing Then
        Set r = ws.Cells.Find(What:=texto, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    End If
    Set FindHeader = r
End Function

' Devuelve la columna del encabezado buscado (por texto); 0 si no está.
Private Function ColumnaPorNombreFind(ws As Worksheet, ByVal texto As String) As Long
    Dim r As Range
    Set r = FindHeader(ws, texto)
    If r Is Nothing Then
        ColumnaPorNombreFind = 0
    Else
        ColumnaPorNombreFind = r.Column
    End If
End Function

' ============================
'  Lectura de la fila de datos principal (SIN depender de "ACTIVO")
'  Estrategia:
'   1) Detecta encabezados por nombres de administradoras y "SISTEMA" usando Find.
'   2) Ubica la fila del encabezado (tomando la fila del encabezado de "SISTEMA").
'   3) Elige la fila de datos debajo del encabezado con el mayor valor en "SISTEMA".
'   4) Lee Proteccion, Porvenir, Skandia, Colfondos y Alternativo (si existe).
' ============================
Private Function TryLeerFilaDatosPrincipal(ws As Worksheet, _
                                           ByRef vColf As Variant, ByRef vPorv As Variant, _
                                           ByRef vProt As Variant, ByRef vSkan As Variant, _
                                           Optional ByRef vAlter As Variant) As Boolean
    On Error GoTo ERR_HANDLER

    If ws Is Nothing Then Exit Function

    ' 1) Encabezados (Find)
    Dim hdrProt As Range, hdrPorv As Range, hdrSkan As Range, hdrAlter As Range, hdrColf As Range, hdrSis As Range
    Set hdrProt = FindHeader(ws, "PROTECCION")
    If hdrProt Is Nothing Then Set hdrProt = FindHeader(ws, "PROTECCIÓN")

    Set hdrPorv = FindHeader(ws, "PORVENIR")
    Set hdrSkan = FindHeader(ws, "SKANDIA")
    Set hdrAlter = FindHeader(ws, "SKANDIA_ALT") ' puede ser Nothing

    ' Colfondos puede venir como "CITI COLFONDOS" en 2025
    Set hdrColf = FindHeader(ws, "CITI COLFONDOS")
    If hdrColf Is Nothing Then Set hdrColf = FindHeader(ws, "COLFONDOS")

    Set hdrSis = FindHeader(ws, "SISTEMA")

    If hdrProt Is Nothing Or hdrPorv Is Nothing Or hdrSkan Is Nothing Or hdrColf Is Nothing Or hdrSis Is Nothing Then
        MsgBox "No se encontraron todos los encabezados (PROTECCION, PORVENIR, SKANDIA, COLFONDOS/SISTEMA) en " & ws.Parent.Name & "!" & ws.Name, vbCritical
        Exit Function
    End If

    ' 2) Fila del encabezado: tomamos la fila del encabezado de SISTEMA
    Dim filaHdr As Long
    filaHdr = hdrSis.Row

    ' 3) Columnas
    Dim cProt As Long, cPorv As Long, cSkan As Long, cAlter As Long, cColf As Long, cSis As Long
    cProt = hdrProt.Column
    cPorv = hdrPorv.Column
    cSkan = hdrSkan.Column
    cColf = hdrColf.Column
    cSis = hdrSis.Column
    cAlter = IIf(hdrAlter Is Nothing, 0, hdrAlter.Column)

    ' 4) Elegir fila de datos: mayor "SISTEMA" en las siguientes 60 filas
    Dim lastR As Long, r As Long, rStart As Long, rEnd As Long
    Dim maxSis As Double, rMax As Long, valSis As Variant
    lastR = ws.Cells(ws.Rows.Count, cSis).End(xlUp).Row
    rStart = filaHdr + 1
    rEnd = Application.WorksheetFunction.Min(filaHdr + 60, lastR)

    maxSis = -1: rMax = 0
    For r = rStart To rEnd
        valSis = ws.Cells(r, cSis).Value
        If IsNumeric(valSis) Then
            If CDbl(valSis) > maxSis Then
                maxSis = CDbl(valSis)
                rMax = r
            End If
        End If
    Next r

    If rMax = 0 Then
        MsgBox "No se encontró fila de datos válida en 'SISTEMA' debajo del encabezado en " & ws.Parent.Name & "!" & ws.Name, vbCritical
        Exit Function
    End If

    ' 5) Lectura con escala /1000
    vColf = NzNum(ws.Cells(rMax, cColf).Value) / 1000
    vPorv = NzNum(ws.Cells(rMax, cPorv).Value) / 1000
    vProt = NzNum(ws.Cells(rMax, cProt).Value) / 1000
    vSkan = NzNum(ws.Cells(rMax, cSkan).Value) / 1000
    If Not IsMissing(vAlter) And cAlter > 0 Then
        vAlter = NzNum(ws.Cells(rMax, cAlter).Value) / 1000
    End If

    TryLeerFilaDatosPrincipal = True
    Exit Function

ERR_HANDLER:
    MsgBox "Error leyendo fila principal en " & ws.Parent.Name & "!" & ws.Name & ": " & Err.Description, vbCritical
End Function

' ============================
'  Lectura del SISTEMA TOTAL (archivo "SISTEMA TOTAL ...")
' ============================
Public Sub LeerSistemaTotal_Find(ByVal rutaArchivo As String, Optional ByVal nombreHoja As String = "restot")
    On Error GoTo SALIDA
    Dim wb As Workbook, ws As Worksheet
    Set wb = AbrirLibroSeguro(rutaArchivo)
    Set ws = ObtenerHoja(wb, nombreHoja)
    If ws Is Nothing Then GoTo SALIDA

    ' Columnas por Find
    Dim cProt As Long, cPorv As Long, cSis As Long
    cProt = ColumnaPorNombreFind(ws, "PROTECCION")
    If cProt = 0 Then cProt = ColumnaPorNombreFind(ws, "PROTECCIÓN")
    cPorv = ColumnaPorNombreFind(ws, "PORVENIR")
    cSis = ColumnaPorNombreFind(ws, "SISTEMA")

    If cProt = 0 Or cPorv = 0 Or cSis = 0 Then
        MsgBox "No se hallaron columnas PROTECCION/PORVENIR/SISTEMA en " & wb.Name & "!" & ws.Name, vbCritical
        GoTo SALIDA
    End If

    ' Fila del encabezado: tomamos la fila del encabezado de SISTEMA
    Dim hdrSis As Range, filaHdr As Long
    Set hdrSis = FindHeader(ws, "SISTEMA")
    If hdrSis Is Nothing Then
        MsgBox "No se halló encabezado 'SISTEMA' en " & wb.Name & "!" & ws.Name, vbCritical
        GoTo SALIDA
    End If
    filaHdr = hdrSis.Row

    ' Elegir fila de datos: máximo SISTEMA en las siguientes 60 filas
    Dim lastR As Long, r As Long, rStart As Long, rEnd As Long
    Dim maxSis As Double, rMax As Long, valSis As Variant
    lastR = ws.Cells(ws.Rows.Count, cSis).End(xlUp).Row
    rStart = filaHdr + 1
    rEnd = Application.WorksheetFunction.Min(filaHdr + 60, lastR)

    maxSis = -1: rMax = 0
    For r = rStart To rEnd
        valSis = ws.Cells(r, cSis).Value
        If IsNumeric(valSis) Then
            If CDbl(valSis) > maxSis Then
                maxSis = CDbl(valSis)
                rMax = r
            End If
        End If
    Next r

    If rMax = 0 Then
        MsgBox "No se encontró fila de datos válida en 'SISTEMA' debajo del encabezado en " & wb.Name & "!" & ws.Name, vbCritical
        GoTo SALIDA
    End If

    ' Asignar variables globales (manteniendo tu escala y fórmula)
    vr_fondo = NzNum(ws.Cells(rMax, cSis).Value) / 1000
    porc_vrfondo = ((NzNum(ws.Cells(rMax, cProt).Value) + NzNum(ws.Cells(rMax, cPorv).Value)) / IIf(vr_fondo = 0, 1, vr_fondo)) / 10

SALIDA:
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

' ==================================
'  Lecturas por fondo (4 procedimientos)
'  Usan TryLeerFilaDatosPrincipal y muestran mensaje si falla.
' ==================================
Public Sub LeerVRFondo_MOD(ByVal rutaArchivo As String, Optional ByVal nombreHoja As String = "restot")
    Dim wb As Workbook, ws As Worksheet
    Set wb = AbrirLibroSeguro(rutaArchivo)
    Set ws = ObtenerHoja(wb, nombreHoja)
    If ws Is Nothing Then GoTo SALIDA

    If Not TryLeerFilaDatosPrincipal(ws, vr_fondo_colf_mod, vr_fondo_porv_mod, vr_fondo_prot_mod, vr_fondo_skan_mod, vr_fondo_alter_mod) Then
        MsgBox "No se pudo leer datos de MODERADO en " & rutaArchivo, vbCritical
    End If

SALIDA:
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

Public Sub LeerVRFondo_CON(ByVal rutaArchivo As String, Optional ByVal nombreHoja As String = "restot")
    Dim wb As Workbook, ws As Worksheet
    Set wb = AbrirLibroSeguro(rutaArchivo)
    Set ws = ObtenerHoja(wb, nombreHoja)
    If ws Is Nothing Then GoTo SALIDA

    If Not TryLeerFilaDatosPrincipal(ws, vr_fondo_colf_con, vr_fondo_porv_con, vr_fondo_prot_con, vr_fondo_skan_con) Then
        MsgBox "No se pudo leer datos de CONSERVADOR en " & rutaArchivo, vbCritical
    End If

SALIDA:
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

Public Sub LeerVRFondo_MR(ByVal rutaArchivo As String, Optional ByVal nombreHoja As String = "restot")
    Dim wb As Workbook, ws As Worksheet
    Set wb = AbrirLibroSeguro(rutaArchivo)
    Set ws = ObtenerHoja(wb, nombreHoja)
    If ws Is Nothing Then GoTo SALIDA

    If Not TryLeerFilaDatosPrincipal(ws, vr_fondo_colf_mr, vr_fondo_porv_mr, vr_fondo_prot_mr, vr_fondo_skan_mr) Then
        MsgBox "No se pudo leer datos de MAYOR RIESGO en " & rutaArchivo, vbCritical
    End If

SALIDA:
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

Public Sub LeerVRFondo_RP(ByVal rutaArchivo As String, Optional ByVal nombreHoja As String = "restot")
    Dim wb As Workbook, ws As Worksheet
    Set wb = AbrirLibroSeguro(rutaArchivo)
    Set ws = ObtenerHoja(wb, nombreHoja)
    If ws Is Nothing Then GoTo SALIDA

    If Not TryLeerFilaDatosPrincipal(ws, vr_fondo_colf_rp, vr_fondo_porv_rp, vr_fondo_prot_rp, vr_fondo_skan_rp) Then
        MsgBox "No se pudo leer datos de RETIRO PROGRAMADO en " & rutaArchivo, vbCritical
    End If

SALIDA:
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

' ==================================
'  Utilidad opcional para depuración
' ==================================
Public Sub DebugResumenVR(Optional ByVal titulo As String = "")
    Dim msg As String
    msg = "TRM=" & trm & vbCrLf & _
          "MOD: Colf=" & vr_fondo_colf_mod & ", Porv=" & vr_fondo_porv_mod & ", Prot=" & vr_fondo_prot_mod & _
          ", Skan=" & vr_fondo_skan_mod & ", Alter=" & vr_fondo_alter_mod & vbCrLf & _
          "CON: Colf=" & vr_fondo_colf_con & ", Porv=" & vr_fondo_porv_con & ", Prot=" & vr_fondo_prot_con & ", Skan=" & vr_fondo_skan_con & vbCrLf & _
          "MR : Colf=" & vr_fondo_colf_mr & ", Porv=" & vr_fondo_porv_mr & ", Prot=" & vr_fondo_prot_mr & ", Skan=" & vr_fondo_skan_mr & vbCrLf & _
          "RP : Colf=" & vr_fondo_colf_rp & ", Porv=" & vr_fondo_porv_rp & ", Prot=" & vr_fondo_prot_rp & ", Skan=" & vr_fondo_skan_rp
    If Len(titulo) > 0 Then msg = titulo & vbCrLf & msg
    MsgBox msg
End Sub


Public Sub LeerCotizantes491(ByVal ruta491 As String)
    Dim wb As Workbook, ws As Worksheet
    Set wb = AbrirLibroSeguro(ruta491)
    Set ws = wb.Worksheets("multifondos")

    ' Encuentra columnas por encabezado (ajusta el texto si en tu hoja se llama distinto)
    Dim colEntidad As Long, colCotizantes As Long
    colEntidad = ColumnaPorNombreFind(ws, "ENTIDAD")
    If colEntidad = 0 Then colEntidad = ColumnaPorNombreFind(ws, "Administrador") ' fallback si cambia

    colCotizantes = ColumnaPorNombreFind(ws, "COTIZANTES")
    If colCotizantes = 0 Then colCotizantes = ColumnaPorNombreFind(ws, "APORTANTES") ' fallback

    If colEntidad = 0 Or colCotizantes = 0 Then
        MsgBox "No se hallaron columnas 'ENTIDAD'/'COTIZANTES' en 491: multifondos", vbCritical
        wb.Close SaveChanges:=False
        Exit Sub
    End If

    ' Buscar filas por nombre de entidad
    Dim fPorv As Long, fProt As Long, fColf As Long, fSkan As Long
    fPorv = BuscarFilaPorTextoEnColumna(ws, colEntidad, "Porvenir")
    fProt = BuscarFilaPorTextoEnColumna(ws, colEntidad, "Protección")
    If fProt = 0 Then fProt = BuscarFilaPorTextoEnColumna(ws, colEntidad, "Proteccion") ' sin acento
    fColf = BuscarFilaPorTextoEnColumna(ws, colEntidad, "Colfondos")
    ' En 2025 algunas tablas muestran "CITI COLFONDOS"
    If fColf = 0 Then fColf = BuscarFilaPorTextoEnColumna(ws, colEntidad, "CITI COLFONDOS")
    fSkan = BuscarFilaPorTextoEnColumna(ws, colEntidad, "Skandia")

    If fPorv = 0 Or fProt = 0 Or fColf = 0 Or fSkan = 0 Then
        MsgBox "No se hallaron todas las entidades (Porvenir, Protección, Colfondos, Skandia) en 491: multifondos", vbCritical
        wb.Close SaveChanges:=False
        Exit Sub
    End If

    ' Asignar variables globales (las que tu macro ya usa)
    cot_porv = NzNum(ws.Cells(fPorv, colCotizantes).Value)
    cot_prot = NzNum(ws.Cells(fProt, colCotizantes).Value)
    cot_colf = NzNum(ws.Cells(fColf, colCotizantes).Value)
    cot_sk = NzNum(ws.Cells(fSkan, colCotizantes).Value)

    wb.Close SaveChanges:=False
End Sub

' Helper: buscar fila por texto (case-insensitive) en una columna dada
Private Function BuscarFilaPorTextoEnColumna(ws As Worksheet, ByVal col As Long, ByVal texto As String) As Long
    Dim lastR As Long, r As Long
    lastR = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    For r = 1 To lastR
        If Normaliza(ws.Cells(r, col).Text) = Normaliza(texto) Then
            BuscarFilaPorTextoEnColumna = r
            Exit Function
        End If
    Next r
    BuscarFilaPorTextoEnColumna = 0
End Function


 '--- Busca la fila en la hoja ws donde columna A contiene la etiqueta (ej: "jun-25") ---
Public Function BuscarFilaPorEtiqueta(ws As Worksheet, ByVal etiqueta As String) As Long
    Dim lastR As Long, r As Long, objetivo As String, celda As String
    If ws Is Nothing Then Exit Function
    objetivo = Normaliza(etiqueta)
    lastR = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For r = 1 To lastR
        celda = Normaliza(ws.Cells(r, 1).Text)
        ' Coincidencia parcial para tolerar espacios extra: ej. "jun-25   "
        If InStr(1, celda, objetivo, vbTextCompare) > 0 Then
            BuscarFilaPorEtiqueta = r
            Exit Function
        End If
    Next r
    BuscarFilaPorEtiqueta = 0
End Function

' --- Busca o crea la fila: si no existe, agrega una nueva y escribe la etiqueta en col A ---
Public Function GetOrCreateFila(ws As Worksheet, ByVal etiqueta As String) As Long
    Dim f As Long
    f = BuscarFilaPorEtiqueta(ws, etiqueta)
    If f = 0 Then
        f = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
        ws.Cells(f, 1).Value = etiqueta
    End If
    GetOrCreateFila = f
End Function


