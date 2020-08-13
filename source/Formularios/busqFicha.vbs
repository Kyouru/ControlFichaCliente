Private Sub btAtras_Click()
    Unload Me
    busqPrestamo.Show (0)
End Sub

Private Sub btEliminar_Click()
    'Solo si se selecciono algun item de la lista
    If ListBox1.ListIndex <> -1 Then
        'Solo si se selecciono algun item de la lista y no es vacio
        If ListBox1.List(ListBox1.ListIndex) <> "" Then
            Dim resp As Integer
            resp = MsgBox("Esta seguro que desea eliminar esta ficha?", vbYesNo + vbQuestion, "Borrar Ficha")
            If resp = vbYes Then
                OpenDB
                strSQL = "UPDATE FICHA SET ANULADO = TRUE WHERE ID_FICHA = " & ListBox1.List(ListBox1.ListIndex)
                On Error GoTo Handle:
                cnn.Execute (strSQL)
                On Error GoTo 0
                
                Set rs = Nothing
                
                ActualizarHoja
                ActualizarLista
            End If
        Else
            MsgBox "Seleccione una entrada con datos"
        End If
    Else
        MsgBox "Seleccione una entrada"
    End If
Handle:
    If cnn.Errors.count > 0 Then
        'Log del Error
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btEliminar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub


Private Sub btModificar_Click()
    'Solo si se selecciono algun item de la lista
    If ListBox1.ListIndex <> -1 Then
        'Solo si se selecciono algun item de la lista y no es vacio
        If ListBox1.List(ListBox1.ListIndex) <> "" Then
            idFicha = ListBox1.List(ListBox1.ListIndex)
            'ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SEGUIMIENTO") = ""
            Unload Me
            modFicha.Show (0)
        End If
    Else
        MsgBox "Seleccione una entrada"
    End If
End Sub

Private Sub btNuevo_Click()
    Unload Me
    newFicha.Show (0)
End Sub

Private Sub btSeleccionar_Click()
    'Solo si se selecciono algun item de la lista
    If ListBox1.ListIndex <> -1 Then
        'Solo si se selecciono algun item de la lista y no es vacio
        If ListBox1.List(ListBox1.ListIndex) <> "" Then
            idFicha = ListBox1.List(ListBox1.ListIndex)
            'ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SEGUIMIENTO") = ""
            Unload Me
            modFicha.Show (0)
        End If
    Else
        MsgBox "Seleccione una entrada"
    End If
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btSeleccionar_Click
End Sub

Private Sub UserForm_Initialize()

    strSQL = "SELECT DOI, CODIGO_SOCIO, SOLICITUD, NOMBRE_PRODUCTO, NOMBRE_SOCIO, NOMBRE_MONEDA, MONTO " & _
    "FROM ((PRESTAMO LEFT JOIN SOCIO ON SOCIO.ID_SOCIO = PRESTAMO.ID_SOCIO_FK)" & _
    "LEFT JOIN MONEDA ON MONEDA.ID_MONEDA = PRESTAMO.ID_MONEDA_FK) " & _
    "LEFT JOIN PRODUCTO ON PRODUCTO.ID_PRODUCTO = PRESTAMO.ID_PRODUCTO_FK " & _
    "WHERE ID_PRESTAMO = " & idPrestamo
    
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        lbNombre.Caption = lbNombre.Caption & rs.Fields("NOMBRE_SOCIO")
        lbCodigo.Caption = lbCodigo.Caption & rs.Fields("CODIGO_SOCIO")
        lbDOI.Caption = lbDOI.Caption & rs.Fields("DOI")
        lbSolicitud.Caption = lbSolicitud.Caption & rs.Fields("SOLICITUD")
        lbProducto.Caption = lbProducto.Caption & rs.Fields("NOMBRE_PRODUCTO")
        lbMoneda.Caption = lbMoneda.Caption & rs.Fields("NOMBRE_MONEDA")
        lbMonto.Caption = lbMonto.Caption & Format(rs.Fields("MONTO"), "#,##0.00")
    End If
    
    Set rs = Nothing
    
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

Public Sub ActualizarHoja()
    strSQL = "SELECT ID_FICHA, NOMBRE_TIPO_FICHA, FECHA_FICHA, USUARIO, FECHA_INGRESO FROM ((" & _
    "FICHA LEFT JOIN PRESTAMO ON PRESTAMO.ID_PRESTAMO = FICHA.ID_PRESTAMO_FK)" & _
    " LEFT JOIN SOCIO ON SOCIO.ID_SOCIO = PRESTAMO.ID_SOCIO_FK)" & _
    " LEFT JOIN TIPO_FICHA ON TIPO_FICHA.ID_TIPO_FICHA = FICHA.ID_TIPO_FICHA_FK" & _
    " WHERE PRESTAMO.ID_PRESTAMO = " & idPrestamo & _
    " AND FICHA.ANULADO = FALSE ORDER BY FECHA_FICHA DESC"
    
    
    With ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP)
        'Limpiar Hoja Temporal
        .Cells(1, 1).CurrentRegion.ClearContents
        
        .Cells(1, 1) = "ID_FICHA"
        .Cells(1, 2) = "TIPO_FICHA"
        .Cells(1, 3) = "FECHA_FICHA"
        .Cells(1, 4) = "USUARIO"
        .Cells(1, 5) = "FECHA_INGRESO"
        
        OpenDB
        On Error GoTo Handle:
        rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
        On Error GoTo 0
        
        If rs.RecordCount > 0 Then
            .Range("A2").CopyFromRecordset rs
        End If
        closeRS
        
        .Cells(1, 1).EntireColumn.NumberFormat = "0"
        .Cells(1, 3).EntireColumn.NumberFormat = "DD/MM/YYYY"
        .Cells(1, 5).EntireColumn.NumberFormat = "DD/MM/YYYY"
        
    End With
    
Handle:
    If cnn.Errors.count > 0 Then
        'Log del Error
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - ActualizarHoja", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

'Agrega la Hoja Temporal a la ListBox
Public Sub ActualizarLista()
    With ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP)
        ListBox1.ColumnWidths = "0;80;70;220;60"
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
        
        If idFicha > 0 Then
            For i = 0 To (ListBox1.ListCount - 1)
                If ListBox1.List(i, 0) = idFicha Then
                    ListBox1.ListIndex = i
                    Exit For
                End If
            Next
        Else
            If ListBox1.ListCount > 0 Then
                ListBox1.ListIndex = 0
            End If
        End If
        
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    closeRS
End Sub
