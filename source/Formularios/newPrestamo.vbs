

Private Sub btCancelar_Click()
    Unload Me
End Sub

Private Sub btGuardar_Click()
    If IsNumeric(tbMonto.Text) Then
    If tbSolicitud.Text <> "" Then
    If cmbProducto.ListIndex <> -1 Then
    If cmbMoneda.ListIndex <> -1 Then
    
    strSQL = "INSERT INTO PRESTAMO (SOLICITUD, ID_PRODUCTO_FK, ID_MONEDA_FK, MONTO, ID_SOCIO_FK) VALUES ('" & tbSolicitud.Text & "'," & _
    cmbProducto.List(cmbProducto.ListIndex, 1) & "," & cmbMoneda.List(cmbMoneda.ListIndex, 1) & "," & _
    tbMonto.Text & "," & idSocio & ")"
    
    OpenDB
    
    On Error GoTo Handle:
    cnn.Execute strSQL
    strSQL = "SELECT @@IDENTITY"
    rs.Open strSQL, cnn
    On Error GoTo 0
    
    idPrestamo = rs.Fields(0)
    
    Set rs = Nothing
    
    busqPrestamo.ActualizarHoja
    busqPrestamo.ActualizarLista
    
    Unload Me
    
    Else
        MsgBox "Moneda Incorrecto"
    End If
    Else
        MsgBox "Producto Incorrecto"
    End If
    Else
        MsgBox "Solicitud Vacia"
    End If
    Else
        MsgBox "Monto Incorrecto"
    End If
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btGuardar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Private Sub UserForm_Initialize()

    strSQL = "SELECT * FROM SOCIO" & _
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
    
    Set rs = Nothing
    
    Dim cont As Integer
    strSQL = "SELECT * FROM PRODUCTO"
    
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
    
    Set rs = Nothing
    
    'cmbProducto.ListIndex = 0
    
    strSQL = "SELECT * FROM MONEDA"
    
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
    'cmbMoneda.ListIndex = 0
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - UserForm_Initialize", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    closeRS
End Sub
