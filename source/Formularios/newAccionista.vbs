Private Sub btCancelar_Click()
    Unload Me
End Sub

Private Sub btGuardar_Click()
    If tbDOI.Text <> "" Then
        If tbNombre.Text <> "" Then
            Dim j As Integer
            strSQL = "INSERT INTO ACCIONISTA (DOI_ACCIONISTA, NOMBRE_ACCIONISTA) VALUES ('" & tbDOI.Text & "', '" & tbNombre.Text & "')"
            
            OpenDB
            On Error GoTo Handle:
            cnn.Execute strSQL
            On Error GoTo 0
            
            On Error Resume Next
            newFicha.actualizarAccionistas
            modFicha.actualizarAccionistas
            
            mAccionista.ActualizarHoja
            mAccionista.ActualizarLista
            On Error GoTo 0
            
            Unload Me
            
        Else
            MsgBox "Nombre Vacio"
        End If
    Else
        MsgBox "DOI Vacio"
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

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    closeRS
End Sub
