Attribute VB_Name = "Módulo1"
Public tmp30_44, tmp45_60, tmp60, aportantes, cons_fdos_admon, afiliados, mod_porv, con_porv, mr_porv, con_mod_porv, con_mr_porv
Public mod_mr_porv, mod_prot, cot_porv, cot_prot, cot_colf, cot_sk, sm_colombia
Public con_prot, mr_prot, con_mod_prot, con_mr_prot, mod_mr_prot, mod_colf, con_colf, mr_colf, con_mod_colf, con_mr_colf, mod_mr_colf, mod_sk, con_sk, mr_sk, con_mod_sk, con_mr_sk, mod_mr_sk, alt_sk
Public ruta, nombre, ULT_TRM, ULT_PEA, ULT_PIB, ULT_DEUDA, col_semestral_colombia, ULT_AFIL, hombres, mujeres, tmp29, nombre_obl, RUTA1, RUTA11, nombre_mod, nombre_con, nombre_mr, nombre_rp, tmesR
Public columna, columna_1, columna_2, columna_20, columna_21, columna_22, columna_19A, columna_19B, columna_16, columna_mes_1, columna_tri_1, columna_tri_2, columna_tri_5, columna_9, columna_10, columna_11, columna_11a, columna_23, columna_32, columna_33
Public trm, deuda_g, PIB, pea, cuenta1, cuenta2, cuenta3, cuenta5, cuenta6, cuenta4, cuenta7, cuenta8
Public activos, pasivos, comisiones, gastos, resul_neto, gastos_prot, gastos_porv, gastos_skan, gastos_colf, admon, comision511500, publicidad519015, otros517000
Public RUTA10, vr_fondo, porc_vrfondo, vr_fondo_prot_mod, vr_fondo_porv_mod, vr_fondo_skan_mod, vr_fondo_alter_mod, vr_fondo_colf_mod
Public vr_fondo_prot_con, vr_fondo_porv_con, vr_fondo_skan_con, vr_fondo_alter_con, vr_fondo_colf_con
Public vr_fondo_prot_mr, vr_fondo_porv_mr, vr_fondo_skan_mr, vr_fondo_alter_mr, vr_fondo_colf_mr
Public vr_fondo_prot_rp, vr_fondo_porv_rp, vr_fondo_skan_rp, vr_fondo_alter_rp, vr_fondo_colf_rp
Public total_pen, total_vej, total_inv, total_sob, traspasos_sistema, traspasos_prot, traspasos_porv, traspasos_skan, traspasos_colf
Public ULT_PEN, ULT_493, afiliados_fallecidos, ULT_COMISION_SONIA, ULT_COMISION_136, Aportes_recibidos
Public comision_ska_obl, comision_ska_seg, comision_por_obl, comision_por_seg, comision_pro_obl, comision_pro_seg, comision_col_obl, comision_col_seg
Public tmp_real_10, tmp_nominal_10, tmp_real_5, tmp_nominal_5, tmp_real_3, tmp_nominal_3, tmp_real_1, tmp_nominal_1
Public tmp_real_colf_12, tmp_nominal_colf_12, tmp_real_oldm_12, tmp_nominal_oldm_12, tmp_real_prot_12, tmp_nominal_prot_12, tmp_real_porv_12, tmp_nominal_porv_12
Public resultado, suma, Total1, dudaG, dudaEF, dudaNF, dudaAC, dudaF, dudaST, dudaGE, dudaEFE, dudaNFE, dudaACE, dudaFE, dudaSTE, otros
Public h17, Total11, tmes, tmpmes, dias_del_mes
Public fecha As Date
Public fecha_fin As Date



Option Explicit
' Devuelve la carpeta del año/mes, por ejemplo:
' C:\...\Historico_Rent_minima\2024\12 diciembre
Public Function CarpetaRentMinima(anio As Long, mes As Long) As String
    Dim base As String
    base = "C:\Users\jcrojas\OneDrive - Superfinanciera\Pensiones\InformesDelegatura\FORMATOS ACTUALIZADOS\PROCESOS MENSUALES\Rentabilidad Minima\Historico_Rent_minima\"
    
    ' "m mmmm" genera: 1 enero, 2 febrero, ..., 12 diciembre
    CarpetaRentMinima = base & anio & "\" & Format(DateSerial(anio, mes, 1), "m mmmm")
End Function

' Devuelve la ruta completa de un archivo de rentabilidad mínima
' Ejemplo:
'   RutaRentabilidad("Moderado", #12/31/2024#)
'   ? ...\2024\12 diciembre\Rent_Vr_Uni_Moderado.xlsm
Public Function RutaRentabilidad(nombreFondo As String, fechaReporte As Date) As String
    RutaRentabilidad = CarpetaRentMinima(Year(fechaReporte), Month(fechaReporte)) _
                       & "\Rent_Vr_Uni_" & nombreFondo & ".xlsm"
End Function

Function ObtenerTRMDesdeSeries(fechaReporte As Date) As Double
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rutaSeries As String
    Dim abriLibro As Boolean
    Dim lastRow As Long
    Dim i As Long
    Dim mejorFecha As Date
    Dim valorTRM As Double
    Dim fechaSerie As Date
    
    ' Ruta del archivo de series
    rutaSeries = "C:\Users\jcrojas\OneDrive - Superfinanciera\Pensiones\InformesDelegatura\FORMATOS ACTUALIZADOS\series PIB_PEA_TRM_DG.xlsm"
    
    ' Intentar usar el libro si ya está abierto
    On Error Resume Next
    Set wb = Workbooks("series PIB_PEA_TRM_DG.xlsm")
    On Error GoTo 0
    
    ' Si no estaba abierto, abrirlo en solo lectura
    If wb Is Nothing Then
        Set wb = Workbooks.Open(Filename:=rutaSeries, ReadOnly:=True)
        abriLibro = True
    End If
    
    ' Hoja con los datos
    Set ws = wb.Sheets("Hoja1")
    
    ' En Hoja1:
    '   B = FECHA (hasta)
    '   C = VALOR (TRM)
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    mejorFecha = 0
    valorTRM = 0
    
    For i = 2 To lastRow   ' si tus datos empiezan en fila 2
        If IsDate(ws.Cells(i, "B").Value) Then
            fechaSerie = CDate(ws.Cells(i, "B").Value)
            
            ' Última fecha <= fechaReporte
            If fechaSerie <= fechaReporte Then
                If fechaSerie >= mejorFecha Then
                    mejorFecha = fechaSerie
                    If IsNumeric(ws.Cells(i, "C").Value) Then
                        valorTRM = CDbl(ws.Cells(i, "C").Value)
                    End If
                End If
            End If
        End If
    Next i
    
    ObtenerTRMDesdeSeries = valorTRM
    
    ' Si abrimos el libro aquí, lo cerramos
    If abriLibro Then
        wb.Close SaveChanges:=False
    End If
End Function



Sub bajar()
'
' ***************** para balances de sociedades **********************
    Dim Pregunta As String
    Pregunta = MsgBox("Deseas Traer Balances", vbYesNo + vbQuestion, "EXCELeINFO")
    If Pregunta = vbNo Then
        'Macros
    Else
        Call CopiarBalances_BaseMes
    ' ***************** para balances de sociedades Anuales**********************
        Call CopiarBalances_BaseAnual
    End If


' ***************** para Actualizar los Indicadores de TRM, PIB, PEA Y DEUDA GUBERNAMENTAL **********************
'    Workbooks.Open Filename:="C:\Users\gjsarmiento\OneDrive - Superfinanciera\D\Janet Sarmiento\formatos\FORMATOS ACTUALIZADOS\series PIB_PEA_TRM_DG.xlsm"
    Workbooks.Open Filename:="C:\Users\jcrojas\OneDrive - Superfinanciera\Pensiones\InformesDelegatura\FORMATOS ACTUALIZADOS\series PIB_PEA_TRM_DG.xlsm"
    Sheets("Hoja1").Activate
    Range("A1").Select
    ULT_TRM = ActiveCell(3, 4)
    ULT_PEA = ActiveCell(3, 11)
    ULT_PIB = ActiveCell(3, 8)
    ULT_DEUDA = ActiveCell(3, 14)
' *****  PARA BAJAR TRM ********
Call ActualizarSeriesSinPortapapeles
    
' ***************** para Traer información de Afiliados F.491**********************
    Windows("Plantilla AIOS-probable.xlsm").Activate
    Sheets("formato").Activate
    Range("a1").Select
    fecha = ActiveCell(2, 4).Value
    fecha_fin = ActiveCell(2, 4).Value


' ***************** para abrir la base de salida del formato semestral **********************
    ruta = "C:\Users\jcrojas\OneDrive - Superfinanciera\Pensiones\AIOS\" + "Semestral_Colombia.xlsx"
    Workbooks.Open Filename:=ruta
    If Year(fecha) = 2017 And Month(fecha) = 12 Then col_semestral_colombia = 3
    If Year(fecha) = 2018 And Month(fecha) = 6 Then col_semestral_colombia = 4
    If Year(fecha) = 2018 And Month(fecha) = 12 Then col_semestral_colombia = 5
    If Year(fecha) = 2019 And Month(fecha) = 6 Then col_semestral_colombia = 6
    If Year(fecha) = 2019 And Month(fecha) = 12 Then col_semestral_colombia = 7
    If Year(fecha) = 2020 And Month(fecha) = 6 Then col_semestral_colombia = 8
    If Year(fecha) = 2020 And Month(fecha) = 12 Then col_semestral_colombia = 9
    If Year(fecha) = 2021 And Month(fecha) = 6 Then col_semestral_colombia = 10
    If Year(fecha) = 2021 And Month(fecha) = 12 Then col_semestral_colombia = 11
    If Year(fecha) = 2022 And Month(fecha) = 6 Then col_semestral_colombia = 12
    If Year(fecha) = 2022 And Month(fecha) = 12 Then col_semestral_colombia = 13
    If Year(fecha) = 2023 And Month(fecha) = 6 Then col_semestral_colombia = 14
    If Year(fecha) = 2023 And Month(fecha) = 12 Then col_semestral_colombia = 15
    If Year(fecha) = 2024 And Month(fecha) = 6 Then col_semestral_colombia = 16
    If Year(fecha) = 2024 And Month(fecha) = 12 Then col_semestral_colombia = 17
    If Year(fecha) = 2025 And Month(fecha) = 6 Then col_semestral_colombia = 18
    If Year(fecha) = 2025 And Month(fecha) = 12 Then col_semestral_colombia = 19
    If Year(fecha) = 2026 And Month(fecha) = 6 Then col_semestral_colombia = 20
    If Year(fecha) = 2026 And Month(fecha) = 12 Then col_semestral_colombia = 21

'    RUTA = "D:\Janet Sarmiento\formatos\FORMATOS ACTUALIZADOS\" + "Serie_Formato_ 491 AFILIADOS AFP.xlsm"
    ruta = "C:\Users\jcrojas\OneDrive - Superfinanciera\Pensiones\InformesDelegatura\FORMATOS ACTUALIZADOS\Formato 491\491 FORMATO TRANSMITIDO\" + "Serie_Formato_ 491 AFILIADOS AFP.xlsm"
    Workbooks.Open Filename:=ruta
    Sheets("TOTAL AFILIADOS").Activate
    Range("a1").Select
    ULT_AFIL = ActiveCell(2, 2)
    Sheets("informe de prensa").Activate
    Range("a1").Select
    ActiveCell(3, 3) = fecha
    hombres = ActiveCell(11, 3)
    mujeres = ActiveCell(11, 4)
    tmp29 = ActiveCell(81, 3) + ActiveCell(81, 4)
    tmp30_44 = ActiveCell(82, 3) + ActiveCell(82, 4)
    tmp45_60 = ActiveCell(83, 3) + ActiveCell(83, 4)
    tmp60 = ActiveCell(84, 3) + ActiveCell(84, 4)
    
    Sheets("multifondos").Activate
    Range("a1").Select
    ActiveCell(4, 3) = fecha
    aportantes = ActiveCell(25, 5)
    cons_fdos_admon = ((ActiveCell(8, 10) + ActiveCell(9, 10)) / ActiveCell(12, 10)) * 100
    afiliados = hombres + mujeres
        
    '*********************** NUMERO DE AFILIADOS POR FONDO ********************************
    mod_porv = ActiveCell(8, 3)
    con_porv = ActiveCell(8, 4)
    mr_porv = ActiveCell(8, 5)
    con_mod_porv = ActiveCell(8, 6)
    con_mr_porv = ActiveCell(8, 7)
    mod_mr_porv = ActiveCell(8, 8)
    
    mod_prot = ActiveCell(9, 3)
    con_prot = ActiveCell(9, 4)
    mr_prot = ActiveCell(9, 5)
    con_mod_prot = ActiveCell(9, 6)
    con_mr_prot = ActiveCell(9, 7)
    mod_mr_prot = ActiveCell(9, 8)
    
    mod_colf = ActiveCell(10, 3)
    con_colf = ActiveCell(10, 4)
    mr_colf = ActiveCell(10, 5)
    con_mod_colf = ActiveCell(10, 6)
    con_mr_colf = ActiveCell(10, 7)
    mod_mr_colf = ActiveCell(10, 8)
    
    mod_sk = ActiveCell(11, 3)
    con_sk = ActiveCell(11, 4)
    mr_sk = ActiveCell(11, 5)
    con_mod_sk = ActiveCell(11, 6)
    con_mr_sk = ActiveCell(11, 7)
    mod_mr_sk = ActiveCell(11, 8)
    alt_sk = ActiveCell(11, 9)
    
    
    '*********************** NUMERO DE COTIZANTES POR ENTIDAD ********************************
    cot_porv = ActiveCell(19, 10)
    cot_prot = ActiveCell(20, 10)
    cot_colf = ActiveCell(21, 10)
    cot_sk = ActiveCell(22, 10)
    
    '*********************** PROMEDIO DE COTIZACIÓN EN COLOMBIA ******************************
    Sheets("SM COLOMBIA").Activate
    Range("a1").Select
    ActiveCell(1, 1) = fecha
    sm_colombia = ActiveCell(8, 5)
  
    
    
    ' ************ para llenar la plantilla de AIOS *******************
    Windows("Plantilla AIOS-probable.xlsm").Activate
    Sheets("CARATULA").Activate
    Range("a1").Select
    ActiveCell(13, 5) = ULT_AFIL
    Sheets("formato").Activate
Range("A1").Select
columna = ActiveCell(3, 4)
columna_1 = ActiveCell(2, 22)
columna_2 = ActiveCell(3, 22)
columna_20 = ActiveCell(4, 22)
columna_21 = ActiveCell(5, 22)
columna_22 = ActiveCell(6, 22)
columna_19A = ActiveCell(7, 22)
columna_19B = ActiveCell(8, 22)
columna_16 = ActiveCell(9, 22)
columna_mes_1 = ActiveCell(20, 22)
columna_tri_1 = ActiveCell(21, 22)
columna_tri_2 = ActiveCell(22, 22)
columna_tri_5 = ActiveCell(24, 22)
columna_9 = ActiveCell(13, 22)
columna_10 = ActiveCell(17, 22)
columna_11 = ActiveCell(18, 22)
columna_11a = ActiveCell(19, 22)
columna_23 = ActiveCell(10, 22)
columna_32 = ActiveCell(25, 22)
columna_33 = ActiveCell(23, 22)


' Leer SIEMPRE la TRM desde series para la fecha actual
trm = ObtenerTRMDesdeSeries(CDate(fecha))
ActiveCell(15, 22).Value = trm   ' actualiza formato!V15 con la TRM del período


deuda_g = ActiveCell(16, 22)
PIB = ActiveCell(12, 22)
pea = ActiveCell(11, 22)


    ActiveCell(5, columna) = tmp29
    ActiveCell(6, columna) = tmp30_44
    ActiveCell(7, columna) = tmp45_60
    ActiveCell(8, columna) = tmp60
    ActiveCell(10, columna) = hombres
    ActiveCell(11, columna) = mujeres
    ActiveCell(13, columna) = aportantes
        
    Windows("Serie_Formato_ 491 AFILIADOS AFP.xlsm").Activate
    nombre = ActiveWindow.Caption
    Windows("Serie_Formato_ 491 AFILIADOS AFP.xlsm").Close SaveChanges:=False
    
    
    Range("a1").Select
    fecha = ActiveCell(2, 3).Value
    Sheets("cuentas").Activate
    Range("a1").Select
'    ActiveCell(2, 3) = fecha
    cuenta1 = ActiveCell(13, 7)
    cuenta2 = ActiveCell(15, 7)
    cuenta3 = ActiveCell(24, 8)
    cuenta5 = ActiveCell(21, 8)
    cuenta6 = ActiveCell(42, 6)
    'cuenta4 = cuenta5 + cuenta6
    cuenta4 = SumaSegura(cuenta5, cuenta6)
    cuenta7 = ActiveCell(29, 7)
    cuenta8 = ActiveCell(44, 5)
    
    activos = ActiveCell(6, 3)
    pasivos = ActiveCell(4, 3)
    comisiones = ActiveCell(13, 5)
    gastos = ActiveCell(15, 7)
    resul_neto = ActiveCell(44, 5)
    
    gastos_prot = ActiveCell(50, 3) - ActiveCell(57, 4)
    gastos_porv = ActiveCell(51, 3) - ActiveCell(69, 4)
    gastos_skan = ActiveCell(52, 3) - ActiveCell(81, 4)
    gastos_colf = ActiveCell(53, 3) - ActiveCell(93, 4)
    
    
    admon = ActiveCell(24, 8)
    comision511500 = ActiveCell(21, 8)
    publicidad519015 = ActiveCell(42, 5)
    otros517000 = ActiveCell(29, 8)
    
    Sheets("formato").Activate
    Range("a1").Select
    ActiveCell(50, columna) = cuenta1
    ActiveCell(51, columna) = cuenta2
    ActiveCell(52, columna) = cuenta3
    ActiveCell(53, columna) = cuenta4
    ActiveCell(54, columna) = cuenta5
    ActiveCell(55, columna) = cuenta6
    ActiveCell(56, columna) = cuenta7
    ActiveCell(58, columna) = cuenta8

    ' ************ para SACAR LA INFORMACIÓN del balance del sistema *******************

Call mes
RUTA10 = RUTA1 & nombre_obl
LeerSistemaTotal_Find RUTA10

        

' ===== Lecturas robustas por etiquetas (2025+) =====
RUTA10 = RUTA1 & nombre_mod: Call LeerVRFondo_MOD(RUTA10)
RUTA10 = RUTA1 & nombre_con: Call LeerVRFondo_CON(RUTA10)
RUTA10 = RUTA1 & nombre_mr: Call LeerVRFondo_MR(RUTA10)
RUTA10 = RUTA1 & nombre_rp: Call LeerVRFondo_RP(RUTA10)

' (opcional) ver resumen en pantalla:

'DebugResumenVR "Valores leídos de balances (ACTIVO)"
    
' ***************** para Traer información de Pensionados F.495 **********************
'    Windows("Plantilla AIOS-probable.xlsm").Activate
'    Sheets("formato").Activate
'    Range("a1").Select
'    fecha = ActiveCell(2, 4).Value
'    fecha_fin = ActiveCell(2, 4).Value

    ruta = "C:\Users\jcrojas\OneDrive - Superfinanciera\Pensiones\InformesDelegatura\FORMATOS ACTUALIZADOS\Formato 495\" + "Series_Formato-495 PENSIONADOS.xlsm"
    Workbooks.Open Filename:=ruta
    Sheets("Total pensionados").Activate
    Range("a1").Select
    ULT_PEN = ActiveCell(5, 2)
    Sheets("por entidad").Activate
    Range("a1").Select
    ActiveCell(6, 3) = fecha_fin
    total_pen = ActiveCell(67, 62)
    total_vej = ActiveCell(66, 60)
    total_inv = ActiveCell(66, 61)
    total_sob = ActiveCell(66, 62)
    
    Windows("Series_Formato-495 PENSIONADOS.xlsm").Activate
    nombre = ActiveWindow.Caption
    Windows("Series_Formato-495 PENSIONADOS.xlsm").Close SaveChanges:=False
    
' ***************** para Traer información de Novimientos de Afiliados F.493 **********************
    ruta = "C:\Users\jcrojas\OneDrive - Superfinanciera\Pensiones\InformesDelegatura\FORMATOS ACTUALIZADOS\" + "Serie_Formato_493 MOVIMIENTO AFILIADOS.xlsx"
    Workbooks.Open Filename:=ruta
    Sheets("Movimientos").Activate
    Range("a1").Select
    ULT_493 = ActiveCell(1, 15)
    Sheets("Traslados Entre AFP").Activate
    Range("a1").Select
    ActiveCell(11, 2) = fecha_fin
    ActiveCell(4, 4) = 99
    traspasos_sistema = ActiveCell(11, 69)
    ActiveCell(4, 4) = 2
    traspasos_prot = ActiveCell(11, 69)
    ActiveCell(4, 4) = 3
    traspasos_porv = ActiveCell(11, 69)
    ActiveCell(4, 4) = 9
    traspasos_skan = ActiveCell(11, 69)
    ActiveCell(4, 4) = 10
    traspasos_colf = ActiveCell(11, 69)
    
    '***********************
    Sheets("Fallecidos").Activate
    Range("a1").Select
    ActiveCell(11, 2) = fecha_fin
    ActiveCell(4, 4) = 99
    afiliados_fallecidos = ActiveCell(11, 13)
    
    Windows("Serie_Formato_493 MOVIMIENTO AFILIADOS.xlsx").Activate
    nombre = ActiveWindow.Caption
    Windows("Serie_Formato_493 MOVIMIENTO AFILIADOS.xlsx").Close SaveChanges:=False
   
 ' ***************** para Traer información de comisiones Carta Circular SONIA **********************
    ruta = "C:\Users\jcrojas\OneDrive - Superfinanciera\Pensiones\InformesDelegatura\FORMATOS ACTUALIZADOS\" + "Comisión FPO desde 2003.xlsx"
    Workbooks.Open Filename:=ruta
    Sheets("COTIZACION CORTE ANUAL").Activate
    Range("a1").Select
    ULT_COMISION_SONIA = ActiveCell(1, 48)
    ActiveCell(1, 1) = fecha_fin
    comision_ska_obl = ActiveCell(1, 2)
    comision_ska_seg = ActiveCell(1, 3)
'*
    comision_por_obl = ActiveCell(1, 6)
    comision_por_seg = ActiveCell(1, 7)
'*
    comision_pro_obl = ActiveCell(1, 14)
    comision_pro_seg = ActiveCell(1, 15)
'*
    comision_col_obl = ActiveCell(1, 18)
    comision_col_seg = ActiveCell(1, 19)
    
    Windows("Comisión FPO desde 2003.xlsx").Activate
    nombre = ActiveWindow.Caption
    Windows("Comisión FPO desde 2003.xlsx").Close SaveChanges:=False
  
 ' ***************** para Traer información del Rentabilidad Trimestral **********************
    If Month(fecha) = 12 Or Month(fecha) = 6 Then
'        RUTA = "D:\Janet Sarmiento\formatos\FORMATOS ACTUALIZADOS\" + "Serie 136 Moderado.xlsx"
        ruta = "C:\Users\jcrojas\OneDrive - Superfinanciera\Pensiones\InformesDelegatura\FORMATOS ACTUALIZADOS\" + "Formato_136_Meses.xlsm"
        Workbooks.Open Filename:=ruta
        Sheets("F_136").Activate
        Range("a1").Select
        ULT_COMISION_136 = ActiveCell(6, 12)
        Sheets("FORMATO OBL").Activate
        Range("a1").Select
        ActiveCell(7, 2) = fecha_fin
        ActiveCell(6, 2) = fecha_fin
        ActiveCell(7, 1) = ActiveCell(6, 1)
        Aportes_recibidos = ActiveCell(6, 5)
    
        Windows("Formato_136_Meses.xlsm").Activate
        nombre = ActiveWindow.Caption
        Windows("Formato_136_Meses.xlsm").Close SaveChanges:=False
    End If
' ***************** para Traer información de Rentabilidadeas SONIA **********************

        ruta = "C:\Users\jcrojas\OneDrive - Superfinanciera\Pensiones\InformesDelegatura\FORMATOS ACTUALIZADOS\PROCESOS MENSUALES\Rentabilidad Minima\Historico_Rent_minima\" + Trim(Str(Year(fecha))) + "\" + tmesR + "\Rent_Vr_Uni_Moderado.xlsm"

Dim rutaMod As String

' fecha ya se trae antes desde formato!D2
rutaMod = RutaRentabilidad("Moderado", fecha)

If Dir(rutaMod) = "" Then
    MsgBox "No se encontró el archivo de rentabilidad MODERADO para " & _
           Format(fecha, "mmmm yyyy") & vbCrLf & rutaMod, vbCritical
    Exit Sub    ' o Exit Function, según dónde estés
End If

Workbooks.Open Filename:=rutaMod


    Sheets("Consolidado").Activate
    Range("a1").Select
    ActiveCell(5, 4) = fecha_fin
    ActiveCell(4, 4) = DateAdd("yyyy", -10, fecha_fin)
    tmp_real_10 = ActiveCell(10, 4)
    tmp_nominal_10 = ActiveCell(11, 4)
'*
    ActiveCell(4, 4) = DateAdd("yyyy", -5, fecha_fin)
    tmp_real_5 = ActiveCell(10, 4)
    tmp_nominal_5 = ActiveCell(11, 4)
'*
    ActiveCell(4, 4) = DateAdd("yyyy", -3, fecha_fin)
    tmp_real_3 = ActiveCell(10, 4)
    tmp_nominal_3 = ActiveCell(11, 4)
'*
    ActiveCell(4, 4) = DateAdd("yyyy", -1, fecha_fin)
    tmp_real_1 = ActiveCell(10, 4)
    tmp_nominal_1 = ActiveCell(11, 4)
'*****
    Sheets("Colfondos").Activate
    Range("a1").Select
    ActiveCell(5, 4) = fecha_fin
    ActiveCell(4, 4) = DateAdd("yyyy", -1, fecha_fin)
    tmp_real_colf_12 = ActiveCell(10, 4)
    tmp_nominal_colf_12 = ActiveCell(11, 4)
'*
    Sheets("oldmutual").Activate
    Range("a1").Select
    ActiveCell(5, 4) = fecha_fin
    ActiveCell(4, 4) = DateAdd("yyyy", -1, fecha_fin)
    tmp_real_oldm_12 = ActiveCell(10, 4)
    tmp_nominal_oldm_12 = ActiveCell(11, 4)
'*
    Sheets("Protección").Activate
    Range("a1").Select
    ActiveCell(5, 4) = fecha_fin
    ActiveCell(4, 4) = DateAdd("yyyy", -1, fecha_fin)
    tmp_real_prot_12 = ActiveCell(10, 4)
    tmp_nominal_prot_12 = ActiveCell(11, 4)
'*
    Sheets("Porvenir").Activate
    Range("a1").Select
    ActiveCell(5, 4) = fecha_fin
    ActiveCell(4, 4) = DateAdd("yyyy", -1, fecha_fin)
    tmp_real_porv_12 = ActiveCell(10, 4)
    tmp_nominal_porv_12 = ActiveCell(11, 4)
   
    Windows("Rent_Vr_Uni_Moderado.xlsm").Activate
    nombre = ActiveWindow.Caption
    Windows("Rent_Vr_Uni_Moderado.xlsm").Close SaveChanges:=False
    
' ************ para llenar Boletin_AIOS TRIMESTRAL  *******************
If Month(fecha) = 3 Or Month(fecha) = 6 Or Month(fecha) = 9 Or Month(fecha) = 12 Then

    Dim wbTri As Workbook
    Dim wsAfiliados As Worksheet
    Dim wsAportantes As Worksheet
    Dim wsColombia As Worksheet
    Dim wsTraspasos As Worksheet
    Dim wsGastos As Worksheet
    Dim wsPromotores As Worksheet
    Dim wsRentabilidad As Worksheet
    Dim wsComisiones As Worksheet

    Dim filaTri1 As Long   ' para afiliados y comisiones
    Dim filaTri2 As Long   ' para aportantes, colombia, traspasos, gastos, promotores
    Dim filaTri5 As Long   ' para rentabilidad

    Dim etiquetaFecha As String
    Dim c As Range


' --- justo antes de abrir Boletin_AIOS TRIMESTRAL.xlsx ---
' MsgBox "TRM=" & trm & vbCrLf & _
       "vr_mod: Colf=" & vr_fondo_colf_mod & ", Porv=" & vr_fondo_porv_mod & ", Prot=" & vr_fondo_prot_mod & ", Skan=" & vr_fondo_skan_mod & ", Alter=" & vr_fondo_alter_mod & vbCrLf & _
       "vr_con: Colf=" & vr_fondo_colf_con & ", Porv=" & vr_fondo_porv_con & ", Prot=" & vr_fondo_prot_con & ", Skan=" & vr_fondo_skan_con & vbCrLf & _
       "vr_mr:  Colf=" & vr_fondo_colf_mr & ", Porv=" & vr_fondo_porv_mr & ", Prot=" & vr_fondo_prot_mr & ", Skan=" & vr_fondo_skan_mr & vbCrLf & _
       "vr_rp:  Colf=" & vr_fondo_colf_rp & ", Porv=" & vr_fondo_porv_rp & ", Prot=" & vr_fondo_prot_rp & ", Skan=" & vr_fondo_skan_rp

' MsgBox "Cotizantes 491 -> Colf=" & cot_colf & _
       ", Porv=" & cot_porv & ", Prot=" & cot_prot & ", Skan=" & cot_sk



' === Abrir Boletín trimestral y preparar hojas ===
ruta = "C:\Users\jcrojas\OneDrive - Superfinanciera\Pensiones\AIOS\Boletin_AIOS TRIMESTRAL.xlsx"
Set wbTri = Workbooks.Open(Filename:=ruta)
Set wsAfiliados = wbTri.Worksheets("afiliados")
Set wsAportantes = wbTri.Worksheets("aportantes")
Set wsColombia = wbTri.Worksheets("colombia")
Set wsTraspasos = wbTri.Worksheets("traspasos")
Set wsGastos = wbTri.Worksheets("gastos")
Set wsPromotores = wbTri.Worksheets("promotores")
Set wsRentabilidad = wbTri.Worksheets("rentabilidad")
Set wsComisiones = wbTri.Worksheets("comisiones")

' Etiqueta del período (ej: "jun-25")
etiquetaFecha = LCase(Format(fecha, "mmm-yy"))

' === Buscar o crear una fila por hoja (independientes) ===
Dim filaAf As Long, filaAport As Long, filaCol As Long, filaTrasp As Long
Dim filaGast As Long, filaProm As Long, filaRent As Long, filaCom As Long

filaAf = GetOrCreateFila(wsAfiliados, etiquetaFecha)
filaAport = GetOrCreateFila(wsAportantes, etiquetaFecha)
filaCol = GetOrCreateFila(wsColombia, etiquetaFecha)
filaTrasp = GetOrCreateFila(wsTraspasos, etiquetaFecha)
filaGast = GetOrCreateFila(wsGastos, etiquetaFecha)
filaProm = GetOrCreateFila(wsPromotores, etiquetaFecha)
filaRent = GetOrCreateFila(wsRentabilidad, etiquetaFecha)
filaCom = GetOrCreateFila(wsComisiones, etiquetaFecha)

' === LLENAR CADA HOJA ===

' --- AFILIADOS (usa filaAf) ---
With wsAfiliados
    ' Colfondos
    .Cells(filaAf, 2).Value = mod_colf
    .Cells(filaAf, 3).Value = con_colf
    .Cells(filaAf, 4).Value = mr_colf
    .Cells(filaAf, 5).Value = con_mod_colf
    .Cells(filaAf, 6).Value = con_mr_colf
    .Cells(filaAf, 7).Value = mod_mr_colf
    ' Porvenir
    .Cells(filaAf, 13).Value = mod_porv
    .Cells(filaAf, 14).Value = con_porv
    .Cells(filaAf, 15).Value = mr_porv
    .Cells(filaAf, 16).Value = con_mod_porv
    .Cells(filaAf, 17).Value = con_mr_porv
    .Cells(filaAf, 18).Value = mod_mr_porv
    ' Protección
    .Cells(filaAf, 19).Value = mod_prot
    .Cells(filaAf, 20).Value = con_prot
    .Cells(filaAf, 21).Value = mr_prot
    .Cells(filaAf, 22).Value = con_mod_prot
    .Cells(filaAf, 23).Value = con_mr_prot
    .Cells(filaAf, 24).Value = mod_mr_prot
    ' ING (sin datos)
    .Cells(filaAf, 25).Value = 0
    .Cells(filaAf, 26).Value = 0
    .Cells(filaAf, 27).Value = 0
    .Cells(filaAf, 28).Value = 0
    .Cells(filaAf, 29).Value = 0
    ' Skandia (mod + alt)
    .Cells(filaAf, 30).Value = mod_sk + alt_sk
    .Cells(filaAf, 31).Value = con_sk
    .Cells(filaAf, 32).Value = mr_sk
    .Cells(filaAf, 33).Value = con_mod_sk
    .Cells(filaAf, 34).Value = con_mr_sk
    .Cells(filaAf, 35).Value = mod_mr_sk
End With

' --- APORTANTES (usa filaAport) ---
' *Antes de este bloque asegúrate de haber ejecutado LeerCotizantes491 para que cot_* no sean cero*
With wsAportantes
    .Cells(filaAport, 2).Value = cot_colf
    .Cells(filaAport, 3).Value = 0
    .Cells(filaAport, 4).Value = cot_porv
    .Cells(filaAport, 5).Value = cot_prot
    .Cells(filaAport, 6).Value = 0
    .Cells(filaAport, 7).Value = cot_sk
End With

' --- COLOMBIA (usa filaCol) ---
With wsColombia
    ' MODERADO
    .Cells(filaCol, 2).Value = DivisionSegura(vr_fondo_colf_mod, trm)
    .Cells(filaCol, 3).Value = 0
    .Cells(filaCol, 4).Value = DivisionSegura(vr_fondo_porv_mod, trm)
    .Cells(filaCol, 5).Value = DivisionSegura(vr_fondo_prot_mod, trm)
    .Cells(filaCol, 6).Value = 0
    .Cells(filaCol, 7).Value = DivisionSegura(SumaSegura(vr_fondo_skan_mod, vr_fondo_alter_mod), trm)
    ' CONSERVADOR
    .Cells(filaCol, 9).Value = DivisionSegura(vr_fondo_colf_con, trm)
    .Cells(filaCol, 10).Value = 0
    .Cells(filaCol, 11).Value = DivisionSegura(vr_fondo_porv_con, trm)
    .Cells(filaCol, 12).Value = DivisionSegura(vr_fondo_prot_con, trm)
    .Cells(filaCol, 13).Value = 0
    .Cells(filaCol, 14).Value = DivisionSegura(vr_fondo_skan_con, trm)
    ' MAYOR RIESGO
    .Cells(filaCol, 16).Value = DivisionSegura(vr_fondo_colf_mr, trm)
    .Cells(filaCol, 17).Value = 0
    .Cells(filaCol, 18).Value = DivisionSegura(vr_fondo_porv_mr, trm)
    .Cells(filaCol, 19).Value = DivisionSegura(vr_fondo_prot_mr, trm)
    .Cells(filaCol, 20).Value = 0
    .Cells(filaCol, 21).Value = DivisionSegura(vr_fondo_skan_mr, trm)
    ' RETIRO PROGRAMADO
    .Cells(filaCol, 23).Value = DivisionSegura(vr_fondo_colf_rp, trm)
    .Cells(filaCol, 24).Value = 0
    .Cells(filaCol, 25).Value = DivisionSegura(vr_fondo_porv_rp, trm)
    .Cells(filaCol, 26).Value = DivisionSegura(vr_fondo_prot_rp, trm)
    .Cells(filaCol, 27).Value = 0
    .Cells(filaCol, 28).Value = DivisionSegura(vr_fondo_skan_rp, trm)
End With

' --- TRASPASOS (usa filaTrasp) ---
With wsTraspasos
    .Cells(filaTrasp, 2).Value = traspasos_colf
    .Cells(filaTrasp, 3).Value = 0
    .Cells(filaTrasp, 4).Value = traspasos_porv
    .Cells(filaTrasp, 5).Value = traspasos_prot
    .Cells(filaTrasp, 6).Value = 0
    .Cells(filaTrasp, 7).Value = traspasos_skan
End With

' --- GASTOS (usa filaGast) ---
With wsGastos
    .Cells(filaGast, 2).Value = SafeDivide(gastos_colf, trm)
    .Cells(filaGast, 3).Value = 0
    .Cells(filaGast, 4).Value = SafeDivide(gastos_porv, trm)
    .Cells(filaGast, 5).Value = SafeDivide(gastos_prot, trm)
    .Cells(filaGast, 6).Value = 0
    .Cells(filaGast, 7).Value = SafeDivide(gastos_skan, trm)
End With

' --- PROMOTORES (usa filaProm) ---
With wsPromotores
    .Cells(filaProm, 2).Value = "n.d."
    .Cells(filaProm, 3).Value = "n.d."
    .Cells(filaProm, 4).Value = "n.d."
    .Cells(filaProm, 5).Value = "n.d."
    .Cells(filaProm, 6).Value = "n.d."
    .Cells(filaProm, 7).Value = "n.d."
End With

' --- RENTABILIDAD (usa filaRent) ---
With wsRentabilidad
    ' Nominal 12m
    .Cells(filaRent, 2).Value = IIf(IsNumeric(tmp_nominal_colf_12), CDbl(tmp_nominal_colf_12) * 100, 0)
    .Cells(filaRent, 3).Value = 0
    .Cells(filaRent, 4).Value = IIf(IsNumeric(tmp_nominal_porv_12), CDbl(tmp_nominal_porv_12) * 100, 0)
    .Cells(filaRent, 5).Value = IIf(IsNumeric(tmp_nominal_prot_12), CDbl(tmp_nominal_prot_12) * 100, 0)
    .Cells(filaRent, 6).Value = 0
    .Cells(filaRent, 7).Value = IIf(IsNumeric(tmp_nominal_oldm_12), CDbl(tmp_nominal_oldm_12) * 100, 0)
    ' Real 12m
    .Cells(filaRent, 10).Value = IIf(IsNumeric(tmp_real_colf_12), CDbl(tmp_real_colf_12) * 100, 0)
    .Cells(filaRent, 11).Value = 0
    .Cells(filaRent, 12).Value = IIf(IsNumeric(tmp_real_porv_12), CDbl(tmp_real_porv_12) * 100, 0)
    .Cells(filaRent, 13).Value = IIf(IsNumeric(tmp_real_prot_12), CDbl(tmp_real_prot_12) * 100, 0)
    .Cells(filaRent, 14).Value = 0
    .Cells(filaRent, 15).Value = IIf(IsNumeric(tmp_real_oldm_12), CDbl(tmp_real_oldm_12) * 100, 0)
End With

' --- COMISIONES (usa filaCom) ---
With wsComisiones
    .Cells(filaCom, 2).Value = comision_col_obl * 100
    .Cells(filaCom, 3).Value = comision_col_seg * 100
    .Cells(filaCom, 4).Value = 0
    .Cells(filaCom, 5).Value = 0
    .Cells(filaCom, 6).Value = comision_por_obl * 100
    .Cells(filaCom, 7).Value = comision_por_seg * 100
    .Cells(filaCom, 8).Value = comision_pro_obl * 100
    .Cells(filaCom, 9).Value = comision_pro_seg * 100
    .Cells(filaCom, 10).Value = 0
    .Cells(filaCom, 11).Value = 0
    .Cells(filaCom, 12).Value = comision_ska_obl * 100
    .Cells(filaCom, 13).Value = comision_ska_seg * 100
End With

wbTri.Close SaveChanges:=True

End If
        
    ' ************ para llenar Boletin_AIOS MENSUAL  *******************
' ************ para llenar Boletin_AIOS MENSUAL  *******************

Dim wbMensual As Workbook
Dim wsMensual As Worksheet
Dim filaMensual As Long
Dim textoFecha As String
Dim ultFila As Long
Dim i As Long

' Texto que aparece en la columna A del boletín (ej: "dic-23")
textoFecha = Format(fecha, "mmm-yy")   ' ej: dic-23

ruta = "C:\Users\jcrojas\OneDrive - Superfinanciera\Pensiones\AIOS\" & _
       "Boletin_AIOS MENSUAL.xlsx"

Set wbMensual = Workbooks.Open(Filename:=ruta)
Set wsMensual = wbMensual.Worksheets("HOJA1")

' Buscar manualmente la fila donde está la fecha en la columna A
ultFila = wsMensual.Cells(wsMensual.Rows.Count, "A").End(xlUp).Row
filaMensual = 0

For i = 1 To ultFila
    If LCase$(Trim$(wsMensual.Cells(i, 1).Text)) = LCase$(Trim$(textoFecha)) Then
        filaMensual = i
        Exit For
    End If
Next i

If filaMensual = 0 Then
    MsgBox "No se encontró la fecha " & textoFecha & _
           " en la columna A de HOJA1 del Boletín mensual.", vbCritical
    wbMensual.Close SaveChanges:=False
    Set wsMensual = Nothing
    Set wbMensual = Nothing
    Exit Sub
End If

With wsMensual
    ' Afiliados, aportantes y traspasos
    .Cells(filaMensual, 2).Value = afiliados
    .Cells(filaMensual, 3).Value = aportantes
    .Cells(filaMensual, 4).Value = traspasos_sistema

    ' Fondos administrados en USD (vr_fondo / TRM)
    .Cells(filaMensual, 5).Value = SafeDivide(vr_fondo, trm)

    ' Rentabilidades 12 meses (nominal y real)
' Rentabilidades 12 meses (nominal y real)
If IsNumeric(tmp_nominal_1) Then
    .Cells(filaMensual, 14).Value = CDbl(tmp_nominal_1) * 100
Else
    .Cells(filaMensual, 14).Value = 0
End If

If IsNumeric(tmp_real_1) Then
    .Cells(filaMensual, 15).Value = CDbl(tmp_real_1) * 100
Else
    .Cells(filaMensual, 15).Value = 0
End If


    ' Número de fondos y concentración de fondos administrados
    .Cells(filaMensual, 16).Value = 4
    .Cells(filaMensual, 17).Value = cons_fdos_admon

    ' Porcentaje del valor del fondo sobre el total
    .Cells(filaMensual, 18).Value = porc_vrfondo
End With

wbMensual.Close SaveChanges:=True
Set wsMensual = Nothing
Set wbMensual = Nothing

        
'***************** trae la información de limites hojas 16 y 17
    ruta = RUTA11 + "LIMITES del nuevo.xlsm"
    
    Workbooks.Open Filename:=ruta
    Sheets("AIOS").Activate
    Range("a1").Select
    Total1 = ActiveCell(4, 28)
    dudaG = ActiveCell(4, 3)
    dudaEF = ActiveCell(4, 5)
    dudaNF = ActiveCell(4, 7)
    dudaAC = ActiveCell(4, 9)
    dudaF = ActiveCell(4, 11)
    dudaST = ActiveCell(4, 13)
    dudaGE = ActiveCell(4, 15)
    dudaEFE = ActiveCell(4, 17)
    dudaNFE = ActiveCell(4, 19)
    dudaACE = ActiveCell(4, 21)
    dudaFE = ActiveCell(4, 23)
    dudaSTE = ActiveCell(4, 25)
    otros = ActiveCell(4, 27)
    h17 = dudaGE + dudaEFE + dudaNFE + dudaACE + dudaFE + dudaSTE
    

If Month(fecha) = 12 Or Month(fecha) = 6 Then
    ' *** BLOQUE ANTIGUO PARA AÑOS < 2017 ***
    ' Antes se activaba el Boletín anual/semestral aquí, pero
    ' para años recientes no es necesario y causa error si el libro no está abierto.
    'If Month(fecha) = 12 Then
    '    Windows("Boletin_AIOS ANUAL.xls").Activate
    'Else
    '    Windows("Boletin_AIOS SEMESTRAL.xls").Activate
    'End If

    If Year(fecha) < 2017 Then
        Sheets("16").Activate
        Range("A1").Select

        'ActiveCell(columna_16, 2) = Total1 / trm
        ActiveCell(columna_tri_2, 2) = SafeDivide(Total11, trm)
        ActiveCell(columna_16, 3) = dudaG
        ActiveCell(columna_16, 4) = dudaEF
        ActiveCell(columna_16, 5) = dudaNF
        ActiveCell(columna_16, 6) = dudaAC
        ActiveCell(columna_16, 7) = dudaF
        ActiveCell(columna_16, 8) = dudaST
        ActiveCell(columna_16, 9) = dudaGE
        ActiveCell(columna_16, 10) = dudaEFE
        ActiveCell(columna_16, 11) = dudaNFE
        ActiveCell(columna_16, 12) = dudaACE
        ActiveCell(columna_16, 13) = dudaFE
        ActiveCell(columna_16, 14) = dudaSTE
        ActiveCell(columna_16, 15) = otros

        Sheets("17").Activate
        Range("A1").Select
        ActiveCell(10, columna_1) = h17 * 100

        Windows("limites.xlsm").Activate
               nombre = ActiveWindow.Caption
        Windows("limites.xlsm").Close SaveChanges:=False
    Else
        ' Para años >= 2017 usa el procedimiento robusto:
        EscribirSemestral_Integral
    End If


End If

If Month(fecha) = 12 Then
    ' Cerrar Boletin_AIOS ANUAL.xls solo si está abierto
    If LibroAbierto("Boletin_AIOS ANUAL.xls") Then
        Workbooks("Boletin_AIOS ANUAL.xls").Close SaveChanges:=True
    End If

ElseIf Month(fecha) = 6 Then
    ' Cerrar Boletin_AIOS SEMESTRAL.xls solo si está abierto
    If LibroAbierto("Boletin_AIOS SEMESTRAL.xls") Then
        Workbooks("Boletin_AIOS SEMESTRAL.xls").Close SaveChanges:=True
    End If
End If


    ' --- Escribir límites en Boletín_AIOS MENSUAL (HOJA1) ---
' --- Completar Boletin_AIOS MENSUAL con los límites (columnas 6–13 y 19) ---

Dim wbMensual2 As Workbook
Dim wsMensual2 As Worksheet
Dim filaMensual2 As Long
Dim textoFecha2 As String
Dim ultFila2 As Long
Dim j As Long

textoFecha2 = Format(fecha, "mmm-yy")

ruta = "C:\Users\jcrojas\OneDrive - Superfinanciera\Pensiones\AIOS\" & _
       "Boletin_AIOS MENSUAL.xlsx"

Set wbMensual2 = Workbooks.Open(Filename:=ruta)
Set wsMensual2 = wbMensual2.Worksheets("HOJA1")

ultFila2 = wsMensual2.Cells(wsMensual2.Rows.Count, "A").End(xlUp).Row
filaMensual2 = 0

For j = 1 To ultFila2
    If LCase$(Trim$(wsMensual2.Cells(j, 1).Text)) = LCase$(Trim$(textoFecha2)) Then
        filaMensual2 = j
        Exit For
    End If
Next j

If filaMensual2 = 0 Then
    MsgBox "No se encontró la fecha " & textoFecha2 & _
           " en la columna A de HOJA1 (límites).", vbCritical
    wbMensual2.Close SaveChanges:=False
    Set wsMensual2 = Nothing
    Set wbMensual2 = Nothing
    Exit Sub
End If

With wsMensual2
    ' Col. 6: Total1 / TRM (en USD)
    .Cells(filaMensual2, 6).Value = SafeDivide(Total1, trm)

    ' Cols. 7–13: porcentajes de límites
    .Cells(filaMensual2, 7).Value = dudaG * 100
    .Cells(filaMensual2, 8).Value = dudaEF * 100
    .Cells(filaMensual2, 9).Value = dudaNF * 100
    .Cells(filaMensual2, 10).Value = dudaAC * 100
    .Cells(filaMensual2, 11).Value = dudaF * 100
    .Cells(filaMensual2, 12).Value = h17 * 100
    .Cells(filaMensual2, 13).Value = otros * 100

    ' Col. 19: TRM usada
    .Cells(filaMensual2, 19).Value = trm
End With

wbMensual2.Close SaveChanges:=True
Set wsMensual2 = Nothing
Set wbMensual2 = Nothing


    
    Windows("Plantilla AIOS-probable.xlsm").Activate
    Sheets("CARATULA").Activate
    Range("a1").Select
    ActiveCell(20, 5) = ULT_PEN
    ActiveCell(22, 5) = ULT_493
    ActiveCell(24, 5) = ULT_COMISION_SONIA
    ActiveCell(26, 5) = ULT_COMISION_136
End Sub

Sub mes()
    Select Case Month(fecha)
    Case 1
        tmes = "enero"
        tmesR = "1 enero"
    Case 2
        tmes = "febrero"
        tmesR = "2 febrero"
    Case 3
        tmes = "marzo"
        tmesR = "3 marzo"
    Case 4
        tmes = "abril"
        tmesR = "4 abril"
    Case 5
        tmes = "mayo"
        tmesR = "5 mayo"
    Case 6
        tmes = "junio"
        tmesR = "6 junio"
    Case 7
        tmes = "julio"
        tmesR = "7 julio"
    Case 8
        tmes = "agosto"
        tmesR = "8 agosto"
    Case 9
        tmes = "septiembre"
        tmesR = "9 septiembre"
    Case 10
        tmes = "octubre"
        tmesR = "10 octubre"
    Case 11
        tmes = "noviembre"
        tmesR = "11 noviembre"
    Case 12
        tmes = "diciembre"
        tmesR = "12 diciembre"
    End Select
    RUTA1 = "C:\Users\jcrojas\OneDrive - Superfinanciera\Pensiones\InformesDelegatura\FORMATOS ACTUALIZADOS\Balances\" + Trim(Str(Year(fecha))) + "\" + tmesR + "\"
    RUTA11 = "C:\Users\jcrojas\OneDrive - Superfinanciera\Pensiones\InformesDelegatura\FORMATOS ACTUALIZADOS\LIMITES\" + Trim(Str(Year(fecha))) + "\" + tmesR + "\"
    
    tmpmes = Month(fecha)
    Select Case tmpmes
        Case 2
            If Year(fecha) Mod 4 = 0 Then
                dias_del_mes = 29
            Else
                dias_del_mes = 28
            End If
        Case 1, 3, 5, 7, 8, 10, 12: dias_del_mes = 31
        Case 4, 6, 9, 11:           dias_del_mes = 30
    End Select
    
    nombre_obl = "SISTEMA TOTAL " + tmes + (Str(Year(fecha))) + ".xls"
    nombre_mod = "MODERADO " + tmes + (Str(Year(fecha))) + ".xls"
    nombre_con = "CONSERVADOR " + tmes + (Str(Year(fecha))) + ".xls"
    nombre_mr = "MAYOR RIESGO " + tmes + (Str(Year(fecha))) + ".xls"
    nombre_rp = "RETIRO PROGRAMADO " + tmes + (Str(Year(fecha))) + ".xls"
   
End Sub



Function DivisionSegura(valor As Variant, divisor1 As Variant, Optional divisor2 As Variant = 1, Optional divisor3 As Variant = 1) As Variant
    ' Valida que todos los argumentos sean numéricos y distintos de cero
    If IsNumeric(valor) And IsNumeric(divisor1) And IsNumeric(divisor2) And IsNumeric(divisor3) Then
        If CDbl(divisor1) <> 0 And CDbl(divisor2) <> 0 And CDbl(divisor3) <> 0 Then
            DivisionSegura = CDbl(valor) / CDbl(divisor1) / CDbl(divisor2) / CDbl(divisor3)
        Else
            DivisionSegura = "no disponible"
        End If
    Else
        DivisionSegura = "no disponible"
    End If
End Function


Function SumaSegura(valor1 As Variant, valor2 As Variant, Optional valor3 As Variant = 0, Optional valor4 As Variant = 0) As Variant
    ' Valida que todos los argumentos sean numéricos antes de sumarlos
    Dim suma As Double
    suma = 0
    
    If IsNumeric(valor1) Then suma = suma + CDbl(valor1)
    If IsNumeric(valor2) Then suma = suma + CDbl(valor2)
    If IsNumeric(valor3) Then suma = suma + CDbl(valor3)
    If IsNumeric(valor4) Then suma = suma + CDbl(valor4)
    
    ' Si no se sumó nada, devolver "no disponible"
    If suma = 0 And (Not IsNumeric(valor1) Or Not IsNumeric(valor2)) Then
        SumaSegura = "no disponible"
    Else
        SumaSegura = suma
    End If
End Function



Private Function SafeDivide(ByVal num As Variant, ByVal den As Variant) As Double
    ' Manejo de errores y tipos
    If IsError(num) Or IsError(den) Then
        SafeDivide = 0
        Exit Function
    End If
    
    If Not IsNumeric(num) Or Not IsNumeric(den) Then
        SafeDivide = 0
        Exit Function
    End If
    
    ' Conversión y validación de cero
    If CDbl(den) = 0 Then
        SafeDivide = 0   ' <-- Cambia aquí si quieres devolver "" o Null
    Else
        SafeDivide = CDbl(num) / CDbl(den)
    End If
End Function

Private Function LibroAbierto(nombreLibro As String) As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.Name, nombreLibro, vbTextCompare) = 0 Then
            LibroAbierto = True
            Exit Function
        End If
    Next wb
End Function


' ===========================================
'  Procedimiento robusto para el SEMESTRAL
'  - Abre o toma Semestral_Colombia.xlsx
'  - Valida columna de destino
'  - Escribe filas 30–44 en Hoja1
'  - Guarda y cierra
'  Requiere que ya estén cargadas:
'  Total1, dudaG, dudaEF, dudaNF, dudaAC, dudaF, dudaST,
'  dudaGE, dudaEFE, dudaNFE, dudaACE, dudaFE, dudaSTE, otros, h17, trm
' ===========================================



Public Sub EscribirSemestral_Integral()
    ' ======= CONFIGURACIÓN =======
    Const RUTA_SEMESTRAL As String = "C:\Users\jcrojas\OneDrive - Superfinanciera\Pensiones\AIOS\Semestral_Colombia.xlsx"
    Const NOMBRE_HOJA1 As String = "Hoja1"   ' si tu pestaña se llama "Hoja", el bloque robusto la detecta
    Const LIMPIAR_RESIDUOS_LEGACY As Boolean = True
    Const LIMPIEZA_FILA29_DESDE As Long = 35     ' AI
    Const LIMPIEZA_FILA29_HASTA As Long = 60     ' BH (ajusta si lo necesitas)
    Const PIB_EN_COP As Boolean = True           ' <-- CAMBIA A False si tu PIB ya está en US$

    Dim wbSem As Workbook
    Dim wsSem As Worksheet
    Dim colDest As Long

    ' ======= VALIDACIONES BÁSICAS =======
    If Month(fecha) <> 6 And Month(fecha) <> 12 Then Exit Sub

    ' Columna destino
    colDest = col_semestral_colombia
    If colDest = 0 Then
        colDest = GetColSemestralColombia(fecha)
        col_semestral_colombia = colDest
    End If
    If colDest = 0 Then
        MsgBox "No se determinó columna del semestral para " & Format(fecha, "mmmm yyyy"), vbCritical
        Exit Sub
    End If

    ' TRM válida
    If Not IsNumeric(trm) Or trm <= 0 Then
        MsgBox "TRM inválida o no cargada para el semestral.", vbCritical
        Exit Sub
    End If

    ' ======= ABRIR LIBRO Y OBTENER HOJA =======
    On Error Resume Next
    Set wbSem = Workbooks("Semestral_Colombia.xlsx") ' si ya está abierto
    On Error GoTo 0
    If wbSem Is Nothing Then
        Set wbSem = Workbooks.Open(Filename:=RUTA_SEMESTRAL, ReadOnly:=False)
        If wbSem Is Nothing Then
            MsgBox "No se pudo abrir Semestral_Colombia.xlsx.", vbCritical
            Exit Sub
        End If
    End If

    ' Selección robusta de hoja
    Set wsSem = Nothing
    On Error Resume Next
    Set wsSem = wbSem.Worksheets(NOMBRE_HOJA1)
    If wsSem Is Nothing Then Set wsSem = wbSem.Worksheets("Hoja") ' alternativa
    If wsSem Is Nothing Then Set wsSem = wbSem.Worksheets(1)      ' primera hoja
    On Error GoTo 0
    If wsSem Is Nothing Then
        MsgBox "No se encontró hoja válida en Semestral_Colombia.xlsx.", vbCritical
        Exit Sub
    End If

    ' ======= LIMPIEZA OPCIONAL DE RESIDUOS LEGACY (fila 29 en columnas lejanas) =======
    If LIMPIAR_RESIDUOS_LEGACY And Year(fecha) >= 2017 Then
        wsSem.Range(wsSem.Cells(29, LIMPIEZA_FILA29_DESDE), wsSem.Cells(29, LIMPIEZA_FILA29_HASTA)).ClearContents
    End If

    ' ======= BLOQUE A — POBLACIÓN, APORTE Y COBERTURA (filas 3–19; + 25–29) =======
    With wsSem
        ' Afiliados y estructura por edades
        .Cells(3, colDest).Value = afiliados

.Cells(4, colDest).Value = SafeDivide(tmp29, afiliados) * 100
.Cells(5, colDest).Value = SafeDivide(tmp30_44, afiliados) * 100
.Cells(6, colDest).Value = SafeDivide(tmp45_60, afiliados) * 100
.Cells(7, colDest).Value = SafeDivide(tmp60, afiliados) * 100

.Cells(8, colDest).Value = _
    SafeDivide(tmp29, afiliados) * 100 + _
    SafeDivide(tmp30_44, afiliados) * 100 + _
    SafeDivide(tmp45_60, afiliados) * 100 + _
    SafeDivide(tmp60, afiliados)

.Cells(10, colDest).Value = SafeDivide(mujeres, afiliados)
.Cells(12, colDest).Value = SafeDivide(afiliados, pea)
.Cells(13, colDest).Value = SafeDivide(aportantes, pea)      ' ? ahora saldrá 0.2907, y el formato % lo mostrará 29..Cells(13, colDest).Value = SafeDivide(aportantes, pea)      ' ? ahora saldrá 0.2907, y el formato % lo mostrará 29.07

        .Cells(14, colDest).Value = SafeDivide(aportantes, afiliados) ' Aportantes / Afiliados

        ' Salario promedio (SM en US$)
        .Cells(15, colDest).Value = SafeDivide(sm_colombia, trm)

        ' Pensionados (F.495)
        .Cells(16, colDest).Value = total_pen
        .Cells(17, colDest).Value = SafeDivide(total_inv, total_pen)           ' % Inválidos
        .Cells(18, colDest).Value = SafeDivide(total_vej, total_pen)           ' % Vejez
        .Cells(19, colDest).Value = SafeDivide(total_sob, total_pen)           ' % Sobrevivencia

        ' (20) marcador "no disponible"
        .Cells(20, colDest).Value = "no disponible"

        ' Fallecidos / 1000
        .Cells(25, colDest).Value = SafeDivide(afiliados_fallecidos, 1000)

        ' Traspasos (F.493)
        .Cells(26, colDest).Value = traspasos_sistema
        .Cells(27, colDest).Value = SafeDivide(traspasos_sistema, afiliados)

        ' Fondos administrados (US$) y relación al PIB
        .Cells(28, colDest).Value = SafeDivide(vr_fondo, trm)

        If PIB_EN_COP Then
            ' PIB en COP: relación en %
            .Cells(29, colDest).Value = SafeDivide(vr_fondo, PIB)
        Else
            ' PIB en US$: relación en %
            .Cells(29, colDest).Value = SafeDivide(SafeDivide(vr_fondo, trm), PIB)
        End If
    End With

    ' ======= BLOQUE B — LÍMITES / COMPOSICIÓN (filas 30–44) =======
    With wsSem
        .Cells(30, colDest).Value = SafeDivide(Total1, trm) ' Total sistema en US$
        .Cells(31, colDest).Value = dudaG
        .Cells(32, colDest).Value = dudaEF
        .Cells(33, colDest).Value = dudaNF
        .Cells(34, colDest).Value = dudaAC
        .Cells(35, colDest).Value = dudaF
        .Cells(36, colDest).Value = dudaST
        .Cells(37, colDest).Value = dudaGE
        .Cells(38, colDest).Value = dudaEFE
        .Cells(39, colDest).Value = dudaNFE
        .Cells(40, colDest).Value = dudaACE
        .Cells(41, colDest).Value = dudaFE
        .Cells(42, colDest).Value = dudaSTE
        .Cells(43, colDest).Value = otros
        .Cells(44, colDest).Value = h17    ' inversiones en moneda extranjera (%)
    End With

    ' ======= BLOQUE C — TAMAÑO/ESTRUCTURA FINANCIERA Y DESEMPEÑO (filas 45–75) =======
    Dim p As Double, p1 As Double, comision_promedio As Double
    p = SafeDivide((activos - pasivos), trm)         ' Patrimonio en US$
    p1 = SafeDivide(vr_fondo, trm)                   ' Fondos en US$

    If IsNumeric(comision_col_obl) And IsNumeric(comision_por_obl) And _
       IsNumeric(comision_pro_obl) And IsNumeric(comision_ska_obl) Then
        comision_promedio = (comision_col_obl + comision_por_obl + comision_pro_obl + comision_ska_obl) / 4
    Else
        comision_promedio = 0
    End If

    With wsSem
        .Cells(45, colDest).Value = SafeDivide(SafeDivide(vr_fondo, trm), deuda_g)
        .Cells(46, colDest).Value = 4
        .Cells(47, colDest).Value = porc_vrfondo

        .Cells(48, colDest).Value = SafeDivide(activos, trm)
        .Cells(49, colDest).Value = SafeDivide(pasivos, trm)
        .Cells(50, colDest).Value = p

        .Cells(51, colDest).Value = comisiones
        .Cells(52, colDest).Value = gastos
        .Cells(53, colDest).Value = comisiones - gastos
        .Cells(54, colDest).Value = resul_neto

        .Cells(55, colDest).Value = admon
        .Cells(56, colDest).Value = comision511500
        .Cells(57, colDest).Value = publicidad519015
        .Cells(58, colDest).Value = comision511500 + publicidad519015
        .Cells(59, colDest).Value = otros517000
        .Cells(60, colDest).Value = admon + comision511500 + publicidad519015 + otros517000

        .Cells(61, colDest).Value = SafeDivide(SafeDivide(Aportes_recibidos, trm), SafeDivide(aportantes, 1000)) * 1000
        
        .Cells(62, colDest).Value = SafeDivide(gastos, SafeDivide(Aportes_recibidos, trm)) * 100

        .Cells(63, colDest).Value = SafeDivide(p, p1) * 100           ' Patrimonio / Fondos (%)
        
        .Cells(64, colDest).Value = SafeDivide(p, afiliados) * 1000000 ' Patrimonio por afiliado

        .Cells(65, colDest).Value = SafeDivide(resul_neto, comisiones) * 100
        .Cells(66, colDest).Value = SafeDivide(resul_neto, p) * 100

        .Cells(67, colDest).Value = SafeDivide(gastos, afiliados) * 1000000
        .Cells(68, colDest).Value = SafeDivide(comisiones, aportantes) * 1000000

        .Cells(69, colDest).Value = SafeDivide(admon, SafeDivide(SafeDivide(Aportes_recibidos, trm), SafeDivide(aportantes, 1000)) * 1000)

        .Cells(70, colDest).Value = 16
        .Cells(71, colDest).Value = comision_promedio * 100
        .Cells(72, colDest).Value = 0
        .Cells(73, colDest).Value = 0
        .Cells(74, colDest).Value = (3 - comision_promedio * 100) * 0.25
        .Cells(75, colDest).Value = (3 - comision_promedio * 100) * 0.75

        ' Serie histórica compacta
        .Cells(77, colDest).Value = comisiones
        .Cells(78, colDest).Value = p1
        .Cells(79, colDest).Value = SafeDivide(comisiones, p1)
        .Cells(80, colDest).Value = Year(fecha) - 1994
    End With

    ' ======= BLOQUE D — RENTABILIDADES (filas 82–89) =======
    With wsSem
        .Cells(82, colDest).Value = IIf(IsNumeric(tmp_nominal_10), CDbl(tmp_nominal_10) * 100, 0)
        .Cells(83, colDest).Value = IIf(IsNumeric(tmp_real_10), CDbl(tmp_real_10) * 100, 0)

        .Cells(84, colDest).Value = IIf(IsNumeric(tmp_nominal_5), CDbl(tmp_nominal_5) * 100, 0)
        .Cells(85, colDest).Value = IIf(IsNumeric(tmp_real_5), CDbl(tmp_real_5) * 100, 0)

        .Cells(86, colDest).Value = IIf(IsNumeric(tmp_nominal_3), CDbl(tmp_nominal_3) * 100, 0)
        .Cells(87, colDest).Value = IIf(IsNumeric(tmp_real_3), CDbl(tmp_real_3) * 100, 0)

        .Cells(88, colDest).Value = IIf(IsNumeric(tmp_nominal_1), CDbl(tmp_nominal_1) * 100, 0)
        .Cells(89, colDest).Value = IIf(IsNumeric(tmp_real_1), CDbl(tmp_real_1) * 100, 0)
    End With

    ' ======= FORMATO =======
    Call AplicarFormatoSemestral(wsSem, colDest)

    ' ======= GUARDAR Y CERRAR =======
    wbSem.Close SaveChanges:=True
End Sub


' ===========================================
'  Función auxiliar: columna destino en Semestral_Colombia.xlsx
'  Devuelve la columna según año/mes (mismo mapeo que tenías)
' ===========================================

' ===========================================
'  Función auxiliar: columna destino en Semestral_Colombia.xlsx
'  Devuelve la columna según año/mes (mismo mapeo que tenías)
' ===========================================
Private Function GetColSemestralColombia(ByVal f As Date) As Long
    Dim y As Long, m As Long
    y = Year(f): m = Month(f)
    GetColSemestralColombia = 0

    Select Case True
        Case y = 2017 And m = 12: GetColSemestralColombia = 3
        Case y = 2018 And m = 6:  GetColSemestralColombia = 4
        Case y = 2018 And m = 12: GetColSemestralColombia = 5
        Case y = 2019 And m = 6:  GetColSemestralColombia = 6
        Case y = 2019 And m = 12: GetColSemestralColombia = 7
        Case y = 2020 And m = 6:  GetColSemestralColombia = 8
        Case y = 2020 And m = 12: GetColSemestralColombia = 9
        Case y = 2021 And m = 6:  GetColSemestralColombia = 10
        Case y = 2021 And m = 12: GetColSemestralColombia = 11
        Case y = 2022 And m = 6:  GetColSemestralColombia = 12
        Case y = 2022 And m = 12: GetColSemestralColombia = 13
        Case y = 2023 And m = 6:  GetColSemestralColombia = 14
        Case y = 2023 And m = 12: GetColSemestralColombia = 15
        Case y = 2024 And m = 6:  GetColSemestralColombia = 16
        Case y = 2024 And m = 12: GetColSemestralColombia = 17
        Case y = 2025 And m = 6:  GetColSemestralColombia = 18
        Case y = 2025 And m = 12: GetColSemestralColombia = 19
        Case y = 2026 And m = 6:  GetColSemestralColombia = 20
        Case y = 2026 And m = 12: GetColSemestralColombia = 21
    End Select
End Function





Private Sub AplicarFormatoSemestral(ByVal wsSem As Worksheet, ByVal colDest As Long)
    ' ===== Conteos (enteros) =====
    wsSem.Cells(3, colDest).NumberFormat = "#,##0"   ' Afiliados
    wsSem.Cells(9, colDest).NumberFormat = "#,##0"   ' Afiliados miles
    wsSem.Cells(11, colDest).NumberFormat = "#,##0"  ' Aportantes
    wsSem.Cells(16, colDest).NumberFormat = "#,##0"  ' Total pensionados
    wsSem.Cells(25, colDest).NumberFormat = "#,##0.00"  ' Fallecidos / 1000"
    wsSem.Cells(26, colDest).NumberFormat = "#,##0"  ' Traspasos
    wsSem.Cells(70, colDest).NumberFormat = "#,##0"  ' marcador
    wsSem.Cells(72, colDest).NumberFormat = "#,##0"
    wsSem.Cells(73, colDest).NumberFormat = "#,##0"

    ' ===== Porcentajes (dos decimales) =====
    wsSem.Cells(4, colDest).NumberFormat = "#,##0.00"
    wsSem.Cells(5, colDest).NumberFormat = "#,##0.00"
    wsSem.Cells(6, colDest).NumberFormat = "#,##0.00"
    wsSem.Cells(7, colDest).NumberFormat = "#,##0.00"
    wsSem.Cells(8, colDest).NumberFormat = "#,##0.00"
    wsSem.Cells(10, colDest).NumberFormat = "#,##0.00"   ' Mujeres/Afiliados
    wsSem.Cells(12, colDest).NumberFormat = "#,##0.00"   ' Afiliados/PEA
    wsSem.Cells(13, colDest).NumberFormat = "#,##0.00"   ' Aportantes/PEA
    wsSem.Cells(14, colDest).NumberFormat = "#,##0.00"   ' Aportantes/Afiliados
    wsSem.Cells(17, colDest).NumberFormat = "#,##0.00"   ' % Inválidos
    wsSem.Cells(18, colDest).NumberFormat = "#,##0.00"   ' % Vejez
    wsSem.Cells(19, colDest).NumberFormat = "#,##0.00"   ' % Sobrevivencia
    wsSem.Cells(27, colDest).NumberFormat = "#,##0.00"   ' Traspasos/Afiliados
    wsSem.Cells(29, colDest).NumberFormat = "#,##0.00%"   ' Fondos/PIB
    wsSem.Cells(45, colDest).NumberFormat = "#,##0.00%"   ' Fondos US$/Deuda (%)
    wsSem.Cells(62, colDest).NumberFormat = "#,##0.00%"   ' Gastos/Aportes (%)
    wsSem.Cells(63, colDest).NumberFormat = "#,##0.00%"   ' Patrimonio/Fondos (%)
    wsSem.Cells(65, colDest).NumberFormat = "#,##0.00%"   ' Resultado/Comisiones (%)
    wsSem.Cells(66, colDest).NumberFormat = "#,##0.00%"   ' Resultado/Patrimonio (%)
    wsSem.Cells(71, colDest).NumberFormat = "#,##0.00%"   ' Comisión promedio (%)

    ' ===== US$ (dos decimales) =====
    wsSem.Cells(15, colDest).NumberFormat = "#,##0.00"   ' SM en US$
    wsSem.Cells(28, colDest).NumberFormat = "#,##0.00"   ' Fondos US$
    wsSem.Cells(30, colDest).NumberFormat = "#,##0.00"   ' Total sistema US$
    wsSem.Cells(48, colDest).NumberFormat = "#,##0.00"   ' Activos US$
    wsSem.Cells(49, colDest).NumberFormat = "#,##0.00"   ' Pasivos US$
    wsSem.Cells(50, colDest).NumberFormat = "#,##0.00"   ' Patrimonio US$
    wsSem.Cells(61, colDest).NumberFormat = "#,##0.00"   ' Aportes/aportante (US$)
    wsSem.Cells(78, colDest).NumberFormat = "#,##0.00"   ' Fondos US$

    ' ===== Ratios y costos unitarios (dos decimales) =====
    wsSem.Cells(47, colDest).NumberFormat = "#,##0.00"    ' porc_vrfondo
    wsSem.Cells(51, colDest).NumberFormat = "#,##0.00"    ' Comisiones
    wsSem.Cells(52, colDest).NumberFormat = "#,##0.00"    ' Gastos
    wsSem.Cells(53, colDest).NumberFormat = "#,##0.00"    ' Comisiones - Gastos
    wsSem.Cells(54, colDest).NumberFormat = "#,##0.00"    ' Resultado neto
    wsSem.Cells(58, colDest).NumberFormat = "#,##0.00"    ' 511500 + 519015
    wsSem.Cells(60, colDest).NumberFormat = "#,##0.00"    ' Total gastos estructura
    wsSem.Cells(64, colDest).NumberFormat = "#,##0.00"    ' Patrimonio por afiliado
    wsSem.Cells(67, colDest).NumberFormat = "#,##0.00"    ' Gastos por afiliado
    wsSem.Cells(68, colDest).NumberFormat = "#,##0.00"    ' Comisiones por aportante
    wsSem.Cells(69, colDest).NumberFormat = "#,##0.00"    ' Admon / (Aportes por aportante)
    wsSem.Cells(74, colDest).NumberFormat = "#,##0.00"    ' cálculo
    wsSem.Cells(75, colDest).NumberFormat = "#,##0.00"    ' cálculo
    wsSem.Cells(79, colDest).NumberFormat = "#,##0.0000"  ' Comisiones/Fondos (ratio con 4 decimales)
    wsSem.Cells(80, colDest).NumberFormat = "#,##0"       ' Años desde 1994

    ' ===== Rentabilidades (dos decimales) =====
    wsSem.Range(wsSem.Cells(82, colDest), wsSem.Cells(89, colDest)).NumberFormat = "#,##0.00"
End Sub

Sub ActualizarSeriesSinPortapapeles()
    Dim wbSeries As Workbook
    Dim wsO As Worksheet
    Dim wsC As Worksheet
    Dim wsCar As Worksheet
    Dim fTRM As Long, fPEA As Long, fPIB As Long, fDG As Long

    ' Abre el libro de series si no está abierto
    On Error Resume Next
    Set wbSeries = Workbooks("series PIB_PEA_TRM_DG.xlsm")
    On Error GoTo 0
    If wbSeries Is Nothing Then
        Set wbSeries = Workbooks.Open( _
            Filename:="C:\Users\jcrojas\OneDrive - Superfinanciera\Pensiones\InformesDelegatura\FORMATOS ACTUALIZADOS\series PIB_PEA_TRM_DG.xlsm", _
            ReadOnly:=True)
        If wbSeries Is Nothing Then
            MsgBox "No se pudo abrir 'series PIB_PEA_TRM_DG.xlsm'", vbCritical
            Exit Sub
        End If
    End If

    Set wsO = wbSeries.Worksheets("Hoja1")
    Set wsC = ThisWorkbook.Worksheets("CUENTAS")
    Set wsCar = ThisWorkbook.Worksheets("CARATULA")

    ' Últimos valores (fila 3)
    ULT_TRM = wsO.Cells(3, "D").Value
    ULT_PEA = wsO.Cells(3, "K").Value
    ULT_PIB = wsO.Cells(3, "H").Value
    ULT_DEUDA = wsO.Cells(3, "N").Value

    ' TRM (B:C -> K:L)
    fTRM = wsO.Cells(wsO.Rows.Count, "C").End(xlUp).Row
    wsC.Range("K2").Resize(fTRM - 2 + 1, 2).Value = wsO.Range("B3:C" & fTRM).Value

    ' PEA (I:J -> N:O)
    fPEA = wsO.Cells(wsO.Rows.Count, "J").End(xlUp).Row
    wsC.Range("N2").Resize(fPEA - 2 + 1, 2).Value = wsO.Range("I3:J" & fPEA).Value

    ' PIB (F:G -> Q:R)
    fPIB = wsO.Cells(wsO.Rows.Count, "G").End(xlUp).Row
    wsC.Range("Q2").Resize(fPIB - 2 + 1, 2).Value = wsO.Range("F3:G" & fPIB).Value

    ' Deuda (L:M -> T:U)
    fDG = wsO.Cells(wsO.Rows.Count, "M").End(xlUp).Row
    wsC.Range("T2").Resize(fDG - 2 + 1, 2).Value = wsO.Range("L3:M" & fDG).Value

    ' CARÁTULA (sin ActiveCell)
    With wsCar
        .Range("E8").Value = ULT_TRM
        .Range("E10").Value = ULT_PEA
        .Range("E9").Value = ULT_PIB
        .Range("E11").Value = ULT_DEUDA
    End With

    If wbSeries.ReadOnly Then wbSeries.Close SaveChanges:=False
End Sub


