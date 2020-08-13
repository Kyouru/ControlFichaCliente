
Private Sub btCancelar_Click()
    Unload Me
End Sub

Private Sub btGuardar_Click()
    If tbCodigo.Text <> "" And tbDOI.Text <> "" And tbNombre.Text <> "" Then
        strSQL = "INSERT INTO SOCIO (ID_GRUPO_FK, DOI, CODIGO_SOCIO, NOMBRE_SOCIO) " & _
        "VALUES (" & cmbGrupo.List(cmbGrupo.ListIndex, 1) & ", '" & tbDOI.Text & "', '" & _
        tbCodigo.Text & "', '" & Replace(tbNombre.Text, "'", "''") & "')"
        
        OpenDB
        On Error GoTo Handle:
        cnn.Execute strSQL
        strSQL = "SELECT @@IDENTITY"
        rs.Open strSQL, cnn
        On Error GoTo 0
        
        idSocio = rs.Fields(0)
        idPrestamo = 0
        Set rs = Nothing
        
        Unload Me
    Else
        MsgBox "Informacion Incompleta"
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
    strSQL = "SELECT * FROM GRUPO WHERE GRUPO.ANULADO = FALSE ORDER BY NOMBRE_GRUPO"
    
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    If rs.RecordCount > 0 Then
        cmbGrupo.Clear
        i = 0
        Do While Not rs.EOF
            cmbGrupo.AddItem rs.Fields("NOMBRE_GRUPO")
            cmbGrupo.List(i, 1) = rs.Fields("ID_GRUPO")
            i = i + 1
            rs.MoveNext
        Loop
    End If
    
    cmbGrupo.Value = "NINGUNO"
    
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
