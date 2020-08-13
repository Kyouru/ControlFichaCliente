
Private Sub btCancelar_Click()
    Unload Me
End Sub

Private Sub btGuardar_Click()
    If tbDOI.Text <> "" Then
    If tbNombre.Text <> "" Then
        strSQL = "UPDATE ACCIONISTA SET DOI_ACCIONISTA = '" & tbDOI.Text & "', NOMBRE_ACCIONISTA = '" & _
        tbNombre.Text & "' WHERE ID_ACCIONISTA = " & idAccionista
        
        OpenDB
        On Error GoTo Handle:
        cnn.Execute (strSQL)
        On Error GoTo 0
        
        Set rs = Nothing
        
        mAccionista.ActualizarHoja
        mAccionista.ActualizarLista
        
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

Private Sub UserForm_Initialize()
    
    OpenDB
    
    strSQL = "SELECT * FROM ACCIONISTA" & _
    " WHERE ID_ACCIONISTA = " & idAccionista
    
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    If rs.RecordCount > 0 Then
        tbNombre.Text = rs.Fields("NOMBRE_ACCIONISTA")
        tbDOI.Text = rs.Fields("DOI_ACCIONISTA")
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

