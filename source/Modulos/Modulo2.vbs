
'MantenimientoDB
Sub MantenimientoDB()
    busqSocio.Show (0)
End Sub

'Grupo
Sub MantenimientoGrupo()
    mGrupos.Show (0)
End Sub

'Accionista
Sub MantenimientoAccionista()
    mAccionista.Show (0)
End Sub

'ActualizarMain
Sub ActualizarMain()
    If Hoja1.Range("HISTORICO") = "SI" Then
        strSQL = "SELECT ID_FICHA, NOMBRE_GRUPO, DOI, CODIGO_SOCIO, NOMBRE_SOCIO, NOMBRE_ACCIONISTA, PARTICIPACION / 100, NOMBRE_TIPO_FICHA, FECHA_FICHA, SOLICITUD, NOMBRE_PRODUCTO, NOMBRE_MONEDA, MONTO FROM " & _
            " ((((((((FICHA_ACCIONISTA FA LEFT JOIN FICHA F ON FA.ID_FICHA_FK = F.ID_FICHA)" & _
            " LEFT JOIN TIPO_FICHA TF ON F.ID_TIPO_FICHA_FK = TF.ID_TIPO_FICHA)" & _
            " LEFT JOIN ACCIONISTA A ON FA.ID_ACCIONISTA_FK = A.ID_ACCIONISTA)" & _
            " LEFT JOIN PRESTAMO P ON F.ID_PRESTAMO_FK = P.ID_PRESTAMO)" & _
            " LEFT JOIN SOCIO S ON S.ID_SOCIO = P.ID_SOCIO_FK)" & _
            " LEFT JOIN PRODUCTO PROD ON P.ID_PRODUCTO_FK = PROD.ID_PRODUCTO)" & _
            " LEFT JOIN MONEDA M ON P.ID_MONEDA_FK = M.ID_MONEDA)" & _
            " LEFT JOIN GRUPO G ON S.ID_GRUPO_FK = G.ID_GRUPO)" & _
            " WHERE S.ANULADO = FALSE" & _
            "   AND P.ANULADO = FALSE" & _
            "   AND F.ANULADO = FALSE" & _
            "   AND FA.ANULADO = FALSE" & _
            "   AND A.ANULADO = FALSE" & _
            " ORDER BY NOMBRE_GRUPO, NOMBRE_SOCIO, FECHA_FICHA DESC"
    Else
        strSQL = "SELECT F.ID_FICHA, NOMBRE_GRUPO, DOI, CODIGO_SOCIO, NOMBRE_SOCIO, NOMBRE_ACCIONISTA, PARTICIPACION / 100, NOMBRE_TIPO_FICHA, FECHA_FICHA, SOLICITUD, NOMBRE_PRODUCTO, NOMBRE_MONEDA, MONTO FROM " & _
            " (((((((((FICHA_ACCIONISTA FA LEFT JOIN FICHA F ON FA.ID_FICHA_FK = F.ID_FICHA)" & _
            " LEFT JOIN (SELECT ID_PRESTAMO_FK, MAX(FECHA_FICHA) AS MAXFICHA FROM FICHA F WHERE F.ANULADO = FALSE GROUP BY ID_PRESTAMO_FK) MAXF ON MAXF.MAXFICHA = F.FECHA_FICHA AND MAXF.ID_PRESTAMO_FK = F.ID_PRESTAMO_FK)" & _
            " LEFT JOIN TIPO_FICHA TF ON F.ID_TIPO_FICHA_FK = TF.ID_TIPO_FICHA)" & _
            " LEFT JOIN ACCIONISTA A ON FA.ID_ACCIONISTA_FK = A.ID_ACCIONISTA)" & _
            " LEFT JOIN PRESTAMO P ON F.ID_PRESTAMO_FK = P.ID_PRESTAMO)" & _
            " LEFT JOIN SOCIO S ON S.ID_SOCIO = P.ID_SOCIO_FK)" & _
            " LEFT JOIN PRODUCTO PROD ON P.ID_PRODUCTO_FK = PROD.ID_PRODUCTO)" & _
            " LEFT JOIN MONEDA M ON P.ID_MONEDA_FK = M.ID_MONEDA)" & _
            " LEFT JOIN GRUPO G ON S.ID_GRUPO_FK = G.ID_GRUPO)" & _
            " WHERE S.ANULADO = FALSE" & _
            "   AND P.ANULADO = FALSE" & _
            "   AND F.ANULADO = FALSE" & _
            "   AND FA.ANULADO = FALSE" & _
            "   AND A.ANULADO = FALSE" & _
            "   AND MAXF.ID_PRESTAMO_FK IS NOT NULL" & _
            " ORDER BY NOMBRE_GRUPO, NOMBRE_SOCIO, FECHA_FICHA DESC"
    End If
    
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    If rs.RecordCount > 0 Then
        Hoja1.Range(Hoja1.Range("dataSet"), Hoja1.Range("dataSet").End(xlDown)).EntireRow.ClearContents
        Hoja1.Range("dataSet").CopyFromRecordset rs
        Hoja1.Range("ALTERNADO").AutoFill Destination:=Range(Hoja1.Range("ALTERNADO"), Hoja1.Range("dataSet").End(xlDown).Offset(0, -1))
        
        Hoja1.Range("PARTICIPACION").EntireColumn.NumberFormat = "0.00%"
        Hoja1.Range("MONTO").EntireColumn.NumberFormat = "#,##0.00"
    End If
Handle:
    If cnn.Errors.count > 0 Then
        'Log del Error
        Call Error_Handle(cnn.Errors.Item(0).Source, "Mulo2 - ActualizarMain", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Sub VerFichaCliente()

    On Error GoTo Handle:
    idFicha = ActiveSheet.Cells(Selection.Row, ActiveSheet.Range("ID_FICHA").Column).Value
    On Error GoTo 0
    
    If Selection.count = 1 Then
        If ActiveSheet.Name = "MAIN" Then
            If idFicha > 0 Then
                
                strSQL = "SELECT EXTENSION FROM FICHA WHERE ID_FICHA = " & idFicha
                                 
                OpenDB
                On Error GoTo Handle:
                rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
                On Error GoTo 0
                
                On Error Resume Next
                ActiveWorkbook.FollowHyperlink ActiveWorkbook.Path & Application.PathSeparator & "RECURSOS" & Application.PathSeparator & idFicha & "." & rs.Fields("EXTENSION").Value
                On Error GoTo 0
                
            End If
        End If
    End If
Handle:
    If cnn.Errors.count > 0 Then
        'Log del Error
        Call Error_Handle(cnn.Errors.Item(0).Source, "Mulo2 - VerFichaCliente", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub
