

Private Sub btAtras_Click()
    Unload Me
    busqSocio.Show (0)
End Sub

'Actualiza la Lista
Private Sub btBuscar_Click()
    ActualizarHoja
    ActualizarLista
End Sub

Private Sub btEliminar_Click()
    'Solo si se selecciono algun item de la lista
    If ListBox1.ListIndex <> -1 Then
        'Solo si se selecciono algun item de la lista y no es vacio
        If ListBox1.List(ListBox1.ListIndex) <> "" Then
            'Confirmacion antes de anular el Prestamo
            Dim resp As Integer
            resp = MsgBox("Est・seguro que desea eliminar este pr駸tamo?", vbYesNo + vbQuestion, ListBox1.List(ListBox1.ListIndex, 3))
            If resp = vbYes Then
                OpenDB
                On Error GoTo Handle:
                strSQL = "UPDATE PRESTAMO SET ANULADO = TRUE WHERE ID_PRESTAMO = " & ListBox1.List(ListBox1.ListIndex)
                cnn.Execute (strSQL)
                On Error GoTo 0
                
                closeRS
                
                'Actualizar la ListBox
                ActualizarHoja
                ActualizarLista
            End If
        End If
    Else
        MsgBox "Seleccione una entrada"
    End If
Handle:
    If cnn.Errors.count > 0 Then
        'Log del Error
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - UserForm_Initialize", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

'Limpia todos los campos de busqueda
Private Sub btLimpiar_Click()
    cmbProducto.Text = ""
    cmbMoneda.Text = ""
    tbSolicitud.Text = ""
    tbMonto.Text = ""
End Sub

'Modifica el Prestamo Seleccionado
Private Sub btModificar_Click()
    'Solo si se selecciono algun item de la lista
    If ListBox1.ListIndex <> -1 Then
        idPrestamo = ListBox1.List(ListBox1.ListIndex)
        modPrestamo.Show (0)
    Else
        MsgBox "Seleccione una entrada"
    End If
End Sub

'Nuevo Prestamo
Private Sub btNuevo_Click()
    newPrestamo.Show (0)
End Sub

'Busca todas las Condiciones del Prestamo seleccionado
Private Sub btSeleccionar_Click()
    'Solo si se selecciono algun item de la lista
    If ListBox1.ListIndex <> -1 Then
        idPrestamo = ListBox1.List(ListBox1.ListIndex)
        
        Unload Me
        busqFicha.Show (0)
    Else
        MsgBox "Seleccione una entrada"
    End If
End Sub

Private Sub cmbMoneda_Change()
    btBuscar_Click
End Sub

Private Sub cmbProducto_Change()
    btBuscar_Click
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    'Solo si se selecciono algun item de la lista
    If ListBox1.ListIndex <> -1 Then
        btSeleccionar_Click
    End If
End Sub

Private Sub tbDesembolso_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btBuscar_Click
        KeyCode = 0
    End If
End Sub

Private Sub tbMonto_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btBuscar_Click
        KeyCode = 0
    End If
End Sub

Private Sub tbSolicitud_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btBuscar_Click
        KeyCode = 0
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim cont As Integer
    
    'Query para obtener los Datos del Socio
    strSQL = "SELECT DOI, CODIGO_SOCIO, NOMBRE_SOCIO FROM SOCIO" & _
    " WHERE ID_SOCIO = " & idSocio
    
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    If rs.RecordCount > 0 Then
        lbNombre.Caption = lbNombre.Caption & rs.Fields("NOMBRE_SOCIO")
        lbCodigo.Caption = lbCodigo.Caption & rs.Fields("CODIGO_SOCIO")
        lbDOI.Caption = lbDOI.Caption & rs.Fields("DOI")
    End If
    closeRS
    
    'Query para obtener los Todos de los Productos de los Prestamos
    strSQL = "SELECT * FROM PRODUCTO"
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    If rs.RecordCount > 0 Then
        cmbProducto.Clear
        cont = 0
        Do While Not rs.EOF
            cmbProducto.AddItem rs.Fields("NOMBRE_PRODUCTO")
            cmbProducto.List(cont, 1) = rs.Fields("ID_PRODUCTO")
            cont = cont + 1
            rs.MoveNext
        Loop
    End If
    closeRS
    
    'Query para obtener los Todas las Monedas de los Prestamos
    strSQL = "SELECT * FROM MONEDA"
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    If rs.RecordCount > 0 Then
        cmbMoneda.Clear
        cont = 0
        Do While Not rs.EOF
            cmbMoneda.AddItem rs.Fields("NOMBRE_MONEDA")
            cmbMoneda.List(cont, 1) = rs.Fields("ID_MONEDA")
            cont = cont + 1
            rs.MoveNext
        Loop
    End If
    closeRS
    
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

'Se Solicita todos los Prestamos que cumplan los filtros del Socio seleccionado previamente y se Copian a una hoja Temporal para luego poder agregarlos a la ListBox
Public Sub ActualizarHoja()

    strSQL = "SELECT ID_PRESTAMO, SOLICITUD, NOMBRE_PRODUCTO, NOMBRE_MONEDA, MONTO, " & _
    "ID_SOCIO_FK " & _
    "FROM ((PRESTAMO LEFT JOIN PRODUCTO ON PRODUCTO.ID_PRODUCTO = PRESTAMO.ID_PRODUCTO_FK)" & _
    " LEFT JOIN MONEDA ON MONEDA.ID_MONEDA = PRESTAMO.ID_MONEDA_FK)" & _
    " WHERE ID_SOCIO_FK=" & idSocio
    
    If tbSolicitud.Text <> "" Then
        strSQL = strSQL & " AND SOLICITUD LIKE '%" & tbSolicitud.Text & "%'"
    End If
    If cmbProducto.ListIndex <> -1 Then
        strSQL = strSQL & " AND ID_PRODUCTO = " & cmbProducto.List(cmbProducto.ListIndex, 1)
    End If
    If cmbMoneda.ListIndex <> -1 Then
        strSQL = strSQL & " AND ID_MONEDA = " & cmbMoneda.List(cmbMoneda.ListIndex, 1)
    End If
    If tbMonto.Text <> "" Then
        strSQL = strSQL & " AND MONTO LIKE '%" & tbMonto.Text & "%'"
    End If
    strSQL = strSQL & " AND PRESTAMO.ANULADO = FALSE"
        
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 1).CurrentRegion.ClearContents
    
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 1) = "ID_PRESTAMO"
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 2) = "SOLICITUD"
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 3) = "PRODUCTO"
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 4) = "MONEDA"
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 5) = "MONTO"
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 6) = "ID_SOCIO_FK"
    
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    If rs.RecordCount > 0 Then
        ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Range("A2").CopyFromRecordset rs
    End If
    
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 1).EntireColumn.NumberFormat = "0"
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 5).EntireColumn.NumberFormat = "#,##0.00"
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 6).EntireColumn.NumberFormat = "0"
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - ActualizarHoja", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub


'Agrega la Hoja Temporal a la ListBox
Public Sub ActualizarLista()
    With ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP)
        ListBox1.ColumnWidths = "0;80;60;60;80;0;"
        ListBox1.ColumnCount = 6
        ListBox1.ColumnHeads = True
        
        'En caso halla mas de una fila
        If .Range("A3") <> "" Then
            ListBox1.RowSource = .Name & "!A2:F" & .Range("A2").End(xlDown).Row
        Else
            'En caso halla solamente una fila
            If .Range("A2") <> "" Then
                ListBox1.RowSource = .Name & "!A2:F2"
            'En caso no hallan datos
            Else
                ListBox1.RowSource = ""
                ListBox1.ColumnHeads = False
            End If
        End If
        
        'En caso de que se provenga de un nivel superior (busqCondicion -> Atras) se selecciona el prestamo al que pertenecia la Condicion
        'Case contrario se selecciona el primer prestamo si lo hubiese
        If idPrestamo > 0 Then
            For i = 0 To (ListBox1.ListCount - 1)
                If ListBox1.List(i, 0) = idPrestamo Then
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
