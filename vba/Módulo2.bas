Attribute VB_Name = "Módulo2"

Option Explicit

' === AYUDA: asignación directa genérica (sin Copy/Paste) ===
Private Sub AsignarDirecto(ByVal wbOrigen As Workbook, _
                           ByVal nombreHojaOrigen As String, _
                           ByVal rangoOrigen As String, _
                           ByVal wbDestino As Workbook, _
                           ByVal nombreHojaDestino As String, _
                           ByVal celdaDestino As String)
    Dim wsO As Worksheet, wsD As Worksheet
    Set wsO = wbOrigen.Worksheets(nombreHojaOrigen)
    Set wsD = wbDestino.Worksheets(nombreHojaDestino)
    wsD.Range(celdaDestino).Resize(wsO.Range(rangoOrigen).Rows.Count, _
                                   wsO.Range(rangoOrigen).Columns.Count).Value = _
                                   wsO.Range(rangoOrigen).Value
End Sub

' === A) base.xlsx ? hoja "base mes" (A:E ? B:F) ===
Public Sub CopiarBalances_BaseMes()
    Dim rutaBase As String
    Dim wbBase As Workbook
    Dim wsBase As Worksheet
    Dim wbPlantilla As Workbook
    Dim lastRow As Long

    rutaBase = "C:\Users\jcrojas\OneDrive - Superfinanciera\Pensiones\InformesDelegatura\FORMATOS ACTUALIZADOS\ESTADOS FINANCIEROS\base.xlsx"
    Set wbBase = Workbooks.Open(Filename:=rutaBase, ReadOnly:=True)
    Set wsBase = wbBase.Worksheets("base")

    ' La plantilla es el libro que contiene la macro
    Set wbPlantilla = ThisWorkbook

    ' Detectar altura dinámica en columna A (ajusta si otra columna es más fiable)
    lastRow = wsBase.Cells(wsBase.Rows.Count, "A").End(xlUp).Row

    ' A1:E(lastRow) ? "base mes"!B1:F
    Call AsignarDirecto(wbBase, "base", "A1:E" & lastRow, _
                        wbPlantilla, "base mes", "B1")

    wbBase.Close SaveChanges:=False
    Set wsBase = Nothing: Set wbBase = Nothing: Set wbPlantilla = Nothing
End Sub

' === B) base_anu.xls ? hoja "base anual" (A:G ? B:H) ===
Public Sub CopiarBalances_BaseAnual()
    Dim rutaAnu As String
    Dim wbAnu As Workbook
    Dim wsAnu As Worksheet
    Dim wbPlantilla As Workbook
    Dim lastRow As Long

    rutaAnu = "C:\Users\jcrojas\OneDrive - Superfinanciera\Pensiones\InformesDelegatura\FORMATOS ACTUALIZADOS\ESTADOS FINANCIEROS\base_anu.xls"
    Set wbAnu = Workbooks.Open(Filename:=rutaAnu, ReadOnly:=True)
    Set wsAnu = wbAnu.Worksheets("base_anual")

    Set wbPlantilla = ThisWorkbook

    lastRow = wsAnu.Cells(wsAnu.Rows.Count, "A").End(xlUp).Row

    ' A1:G(lastRow) ? "base anual"!B1:H
    Call AsignarDirecto(wbAnu, "base_anual", "A1:G" & lastRow, _
                        wbPlantilla, "base anual", "B1")

    wbAnu.Close SaveChanges:=False
    Set wsAnu = Nothing: Set wbAnu = Nothing: Set wbPlantilla = Nothing
End Sub


