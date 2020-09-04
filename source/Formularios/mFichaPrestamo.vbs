Private Sub btCancelar_Click()
    Unload Me
End Sub

Private Sub btDesvincular_Click()
    If ListBox1.ListIndex <> -1 Then
        'Confirmar la Desvinculacion
        Dim resp As Integer
        resp = MsgBox("Esta seguro que desea desvicuncular esta ficha del prestamo?", vbYesNo + vbQuestion, ListBox1.List(ListBox1.ListIndex, 1))
        If resp = vbYes Then
        
            OpenDB
            On Error GoTo Handle:
            cnn.Execute ("UPDATE FICHA_PRESTAMO SET ANULADO = TRUE WHERE ID_FICHA_PRESTAMO = " & ListBox1.List(ListBox1.ListIndex, 0))
            On Error GoTo 0
            
            closeRS
            
            ActualizarHoja
            ActualizarLista
        End If
    Else
        MsgBox "Seleccione una entrada"
    End If
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btDesvincular_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Private Sub UserForm_Initialize()
    Dim cont As Integer
    
    'Query para obtener los Datos del Socio
    strSQL = "SELECT DOI, CODIGO_SOCIO, NOMBRE_SOCIO, SOLICITUD, MONTO, NOMBRE_MONEDA, NOMBRE_PRODUCTO" & _
        " FROM ((PRESTAMO P LEFT JOIN SOCIO S ON S.ID_SOCIO = P.ID_SOCIO_FK)" & _
        " LEFT JOIN PRODUCTO PROD ON PROD.ID_PRODUCTO = P.ID_PRODUCTO_FK)" & _
        " LEFT JOIN MONEDA M ON M.ID_MONEDA = P.ID_MONEDA_FK" & _
        " WHERE ID_PRESTAMO = " & idPrestamo
    
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    If rs.RecordCount > 0 Then
        lbNombre.Caption = lbNombre.Caption & rs.Fields("NOMBRE_SOCIO")
        lbCodigo.Caption = lbCodigo.Caption & rs.Fields("CODIGO_SOCIO")
        lbDOI.Caption = lbDOI.Caption & rs.Fields("DOI")
        lbSolicitud.Caption = lbSolicitud.Caption & rs.Fields("SOLICITUD")
        lbProducto.Caption = lbProducto.Caption & rs.Fields("NOMBRE_PRODUCTO")
        lbMoneda.Caption = lbMoneda.Caption & rs.Fields("NOMBRE_MONEDA")
        lbMonto.Caption = lbMonto.Caption & rs.Fields("MONTO")
    End If
    
    'Actualizar ListBox
    ActualizarHoja
    ActualizarLista
    
Handle:
    If cnn.Errors.count > 0 Then
        'Log del Error
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - UserForm_Initialize", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Private Sub ActualizarHoja()
    
    strSQL = "SELECT ID_FICHA_PRESTAMO, FECHA_FICHA_P, NOMBRE_TIPO_FICHA, USUARIO_INGRESA, FECHA_INGRESA FROM FICHA_PRESTAMO FP LEFT JOIN TIPO_FICHA TF ON TF.ID_TIPO_FICHA = FP.ID_TIPO_FICHA_FK" & _
            " WHERE FP.ANULADO = FALSE AND ID_PRESTAMO_FK = " & idPrestamo
    
    strSQL = strSQL & " ORDER BY FECHA_FICHA_P DESC"
    
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 1).CurrentRegion.ClearContents
    
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 1) = "ID_FICHA_PRESTAMO"
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 2) = "FECHA_FICHA"
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 3) = "NOMBRE_TIPO_FICHA"
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 4) = "USUARIO_INGRESA"
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 5) = "FECHA_INGRESA"
    
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    If rs.RecordCount > 0 Then
        ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Range("A2").CopyFromRecordset rs
    End If
    
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 1).EntireColumn.NumberFormat = "0"
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 2).EntireColumn.NumberFormat = "DD/MM/YYYY"
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 5).EntireColumn.NumberFormat = "DD/MM/YYYY HH:mm:SS"
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - ActualizarHoja", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub


'Agrega la Hoja Temporal a la ListBox
Private Sub ActualizarLista()
    With ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP)
        ListBox1.ColumnWidths = "0;70;90;150;100"
        ListBox1.ColumnCount = 5
        ListBox1.ColumnHeads = True
        
        'En caso halla mas de una fila
        If .Range("A3") <> "" Then
            ListBox1.RowSource = .Name & "!A2:E" & .Range("A2").End(xlDown).Row
        Else
            'En caso halla solamente una fila
            If .Range("A2") <> "" Then
                ListBox1.RowSource = .Name & "!A2:E2"
            'En caso no hallan datos
            Else
                ListBox1.RowSource = ""
                ListBox1.ColumnHeads = False
            End If
        End If
    End With
    
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    busqPrestamo.ActualizarHoja
    busqPrestamo.ActualizarLista
End Sub
