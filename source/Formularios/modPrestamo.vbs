
Private Sub btCancelar_Click()
    Unload Me
End Sub

Private Sub btGuardar_Click()
    If tbMonto.Text = "" Or IsNumeric(tbMonto.Text) Then
    If cmbProducto.ListIndex <> -1 Then
    If cmbMoneda.ListIndex <> -1 Then
    
        strSQL = "UPDATE PRESTAMO SET SOLICITUD = '" & tbSolicitud.Text & "', ID_PRODUCTO_FK = " & _
        cmbProducto.List(cmbProducto.ListIndex, 1) & ", ID_MONEDA_FK = " & _
        cmbMoneda.List(cmbMoneda.ListIndex, 1) & ", MONTO = "
        If tbMonto.Text <> "" Then
            strSQL = strSQL & tbMonto.Text
        Else
            strSQL = strSQL & "NULL"
        End If
        strSQL = strSQL & " WHERE ID_PRESTAMO = " & idPrestamo
        
        OpenDB
        On Error GoTo Handle:
        cnn.Execute strSQL
        On Error GoTo 0
        
        Set rs = Nothing
        
        busqPrestamo.ActualizarHoja
        busqPrestamo.ActualizarLista
        
        Unload Me
    Else
        MsgBox "Error en Moneda"
    End If
    Else
        MsgBox "Error en Producto"
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
    
    Dim cont As Integer
    strSQL = "SELECT * FROM PRODUCTO WHERE ANULADO = FALSE"
    
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
    
    Set rs = Nothing
    
    strSQL = "SELECT * FROM MONEDA WHERE ANULADO = FALSE"
    
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
    
    Set rs = Nothing
    
    strSQL = "SELECT * FROM (((PRESTAMO LEFT JOIN SOCIO ON SOCIO.ID_SOCIO = PRESTAMO.ID_SOCIO_FK) " & _
    "LEFT JOIN MONEDA ON MONEDA.ID_MONEDA = PRESTAMO.ID_MONEDA_FK) " & _
    "LEFT JOIN PRODUCTO ON PRODUCTO.ID_PRODUCTO = PRESTAMO.ID_PRODUCTO_FK) " & _
    "WHERE ID_PRESTAMO = " & idPrestamo
    
    
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    If rs.RecordCount > 0 Then
        lbNombre.Caption = lbNombre.Caption & rs.Fields("NOMBRE_SOCIO")
        lbCodigo.Caption = lbCodigo.Caption & rs.Fields("CODIGO_SOCIO")
        lbDOI.Caption = lbDOI.Caption & rs.Fields("DOI")
        tbSolicitud.Text = rs.Fields("SOLICITUD")
        cmbProducto.Text = rs.Fields("NOMBRE_PRODUCTO")
        cmbMoneda.Text = rs.Fields("NOMBRE_MONEDA")
        tbMonto.Text = rs.Fields("MONTO")
    End If
    
    
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
