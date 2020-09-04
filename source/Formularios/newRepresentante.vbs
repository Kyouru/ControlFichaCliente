Private Sub btCancelar_Click()
    Unload Me
End Sub

Private Sub tbCargo_Change()
    Dim pos As Integer
    pos = tbCargo.SelStart
    tbCargo.Text = UCase(tbCargo.Text)
    tbCargo.SelStart = pos
End Sub

Private Sub tbNombre_Change()
    Dim pos As Integer
    pos = tbNombre.SelStart
    tbNombre.Text = UCase(tbNombre.Text)
    tbNombre.SelStart = pos
End Sub

Private Sub tbPPoderes_Change()
    Dim pos As Integer
    pos = tbPPoderes.SelStart
    tbPPoderes.Text = UCase(tbPPoderes.Text)
    tbPPoderes.SelStart = pos
End Sub

Private Sub btGuardar_Click()
    Dim Cargo As Boolean
    Dim PPoderes As Boolean
    
    Cargo = True
    PPoderes = True
    
    If tbNombre.Text = "" Then
        MsgBox "Nombre Vacio"
        Exit Sub
    End If
    
    If tbCargo.Text = "" Then
        Cargo = False
    End If
    
    If tbPPoderes.Text = "" Then
        PPoderes = False
    End If
    
    strSQL = "INSERT INTO REPRESENTANTE (NOMBRE_REPRESENTANTE, CARGO_REPRESENTANTE, PRINCIPALES_PODERES, ID_SOCIO_FK)" & _
            "VALUES ('" & tbNombre.Text & "', '" & tbCargo.Text & "', '" & tbPPoderes.Text & "', " & idSocio & ")"
    
    strSQL = "INSERT INTO REPRESENTANTE (NOMBRE_REPRESENTANTE, ID_SOCIO_FK"
    
    If Cargo Then
        strSQL = strSQL & ", CARGO_REPRESENTANTE"
    End If
    
    If PPoderes Then
        strSQL = strSQL & ", PRINCIPALES_PODERES"
    End If
    
    strSQL = strSQL & ") VALUES ('" & tbNombre.Text & "', " & idSocio
    
    If Cargo Then
        strSQL = strSQL & ", '" & tbCargo.Text & "'"
    End If
    
    If PPoderes Then
        strSQL = strSQL & ", '" & tbPPoderes.Text & "'"
    End If
    
    strSQL = strSQL & ")"
    
    OpenDB
    On Error GoTo Handle:
    cnn.Execute strSQL
    On Error GoTo 0
    
    On Error Resume Next
    newFicha.actualizarRepresentantes
    On Error Resume Next
    modFicha.actualizarRepresentantes
    
    On Error Resume Next
    mRepresentante.ActualizarHoja
    On Error Resume Next
    mRepresentante.ActualizarLista
    On Error GoTo 0
    
    Unload Me
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btGuardar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    closeRS
End Sub
