
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
    
    Set rs = Nothing
    
    strSQL = "SELECT * FROM ACCIONISTA A LEFT JOIN NACIONALIDAD N ON A.ID_NACIONALIDAD_FK = N.ID_NACIONALIDAD" & _
    " WHERE ID_ACCIONISTA = " & idAccionista
    
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    If rs.RecordCount > 0 Then
        tbNombre.Text = rs.Fields("NOMBRE_ACCIONISTA")
        tbDOI.Text = rs.Fields("DOI_ACCIONISTA")
        Me.cmbNacionalidad.Text = rs.Fields("NOMBRE_NACIONALIDAD")
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

