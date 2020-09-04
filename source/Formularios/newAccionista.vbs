Private Sub btCancelar_Click()
    Unload Me
End Sub

Private Sub btGuardar_Click()
    If tbDOI.Text <> "" Then
        If tbNombre.Text <> "" Then
            
            strSQL = "INSERT INTO ACCIONISTA (DOI_ACCIONISTA, NOMBRE_ACCIONISTA, ID_NACIONALIDAD_FK) VALUES ('" & tbDOI.Text & "', '" & tbNombre.Text & "', " & Me.cmbNacionalidad.List(Me.cmbNacionalidad.ListIndex, 1) & ")"
            
            OpenDB
            On Error GoTo Handle:
            cnn.Execute strSQL
            On Error GoTo 0
            
            On Error Resume Next
            newFicha.actualizarAccionistas
            On Error Resume Next
            modFicha.actualizarAccionistas
            
            On Error Resume Next
            mAccionista.ActualizarHoja
            On Error Resume Next
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

Private Sub UserForm_Initialize()
    Dim i As Integer
    
    strSQL = "SELECT * FROM NACIONALIDAD WHERE ANULADO = FALSE"
    
    OpenDB
    
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    If rs.RecordCount > 0 Then
        i = 0
        Do While Not rs.EOF
            Me.cmbNacionalidad.AddItem rs.Fields("NOMBRE_NACIONALIDAD")
            Me.cmbNacionalidad.List(i, 1) = rs.Fields("ID_NACIONALIDAD")
            i = i + 1
            rs.MoveNext
        Loop
    End If
    
Handle:
    If cnn.Errors.count > 0 Then
        'Log del Error
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - UserForm_Initialize", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    closeRS
End Sub
