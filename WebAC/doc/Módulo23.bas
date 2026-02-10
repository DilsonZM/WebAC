Attribute VB_Name = "Módulo23"
Sub Criticidad()
    Dim wsCrit As Worksheet
    Dim wsRes  As Worksheet
    Dim lastRow As Long
    Dim i       As Long, j As Integer
    Dim isEmptyRow As Boolean
    Dim resRow  As Long
    
    Set wsCrit = Worksheets("Calificación de Criterios")
    Set wsRes = Worksheets("Resultados de Criticidad")
    
    ' Borrar todo desde la segunda fila hacia abajo en la hoja de resultados
    wsRes.Rows("2:" & wsRes.Rows.Count).ClearContents
    
    lastRow = wsCrit.Cells(wsCrit.Rows.Count, "I").End(xlUp).Row
    resRow = 2  ' Iniciar en la segunda fila de la hoja de resultados
    
    For i = 2 To lastRow
        ' Detectar si la fila está vacía en columnas 1–8
        isEmptyRow = True
        For j = 1 To 8
            If Not IsEmpty(wsCrit.Cells(i, j).value) Then
                isEmptyRow = False
                Exit For
            End If
        Next j
        
        ' Sólo procesar si la fila no está vacía Y Aplica (col R) <> 0
        If Not isEmptyRow And wsCrit.Cells(i, 18).value <> 0 Then
            ' — Copiar primeras 8 columnas —
            For j = 1 To 8
                wsRes.Cells(resRow, j).value = wsCrit.Cells(i, j).value
            Next j
            
            ' — Calcular Criticidad en S (columna 19) —
            wsCrit.Cells(i, 19).FormulaR1C1 = _
              "=OFFSET(Matriz!R2C1,6-(6-IF((5-INT([@[FLEXIBILIDAD OPERACIONAL]]*MAX(MATRIX3[@[FINANCIERO]:[TI]])/5))=0,1,5-INT([@[FLEXIBILIDAD OPERACIONAL]]*MAX(MATRIX3[@[FINANCIERO]:[TI]])/5))),RC[-1])"
            Dim vCrit As Variant
            vCrit = wsCrit.Cells(i, 19).value
            wsCrit.Cells(i, 19).value = vCrit
            
            ' — Volcar criticidad en resultados —
            wsRes.Cells(resRow, 9).value = vCrit
            
            ' — Clasificación en columna 10 —
            wsRes.Cells(resRow, 10).FormulaR1C1 = _
              "=IF(RC[-1]<=30,""Baja"",IF(RC[-1]<=74,""Media"",""Alta""))&"" Criticidad"""
            wsRes.Cells(resRow, 10).value = wsRes.Cells(resRow, 10).value
            
            ' — Comentarios (columna 11) —
            wsRes.Cells(resRow, 11).value = wsCrit.Cells(i, 20).value
            
            resRow = resRow + 1
        Else
            ' No aplica: limpiar S
            wsCrit.Cells(i, 19).ClearContents
        End If
    Next i
    
    ' Luego ejecutamos la macro de orden y gráfico
    Order
    CrearTablaYGraficoDinamico
End Sub

