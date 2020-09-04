
Private Sub btCancelar_Click()
    Unload Me
End Sub

Private Sub btGuardar_Click()
    Dim valido As Boolean
    
    valido = True
    
    If tbNombre.Text = "" Then
        MsgBox "Nombre Vacio"
        valido = False
    End If
    
    If valido Then
        strSQL = "UPDATE REPRESENTANTE SET NOMBRE_REPRESENTANTE = '" & tbNombre.Text & "'"

        If tbCargo.Text <> "" Then
            strSQL = strSQL & ", CARGO_REPRESENTANTE = '" & tbCargo.Text & "'"
        Else
            strSQL = strSQL & ", CARGO_REPRESENTANTE = NULL"
        End If
        
        If tbPPoderes.Text <> "" Then
            strSQL = strSQL & ", PRINCIPALES_PODERES = '" & tbPPoderes.Text & "'"
        Else
            strSQL = strSQL & ", PRINCIPALES_PODERES = NULL"
        End If
        
        strSQL = strSQL & " WHERE ID_REPRESENTANTE = " & idRepresentante
        
        OpenDB
        
        On Error GoTo Handle:
        cnn.Execute (strSQL)
        On Error GoTo 0
        
        Set rs = Nothing
        
        mRepresentante.ActualizarHoja
        mRepresentante.ActualizarLista
        
        Unload Me
    End If
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btGuardar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Private Sub tbNombre_Change()
    Dim pos As Integer
    pos = tbNombre.SelStart
    tbNombre.Text = UCase(tbNombre.Text)
    tbNombre.SelStart = pos
End Sub

Private Sub UserForm_Initialize()
    
    OpenDB
    
    strSQL = "SELECT * FROM REPRESENTANTE R" & _
    " WHERE ID_REPRESENTANTE = " & idRepresentante
    
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    If rs.RecordCount > 0 Then
        tbNombre.Text = rs.Fields("NOMBRE_REPRESENTANTE")
        If Not IsNull(rs.Fields("CARGO_REPRESENTANTE")) Then
            tbCargo.Text = rs.Fields("CARGO_REPRESENTANTE")
        End If
        If Not IsNull(rs.Fields("PRINCIPALES_PODERES")) Then
            tbPPoderes.Text = rs.Fields("PRINCIPALES_PODERES")
        End If
        
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





