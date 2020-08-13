
Private Sub btCancelar_Click()
    Unload Me
End Sub

Private Sub btGuardar_Click()
    If cmbGrupo.ListIndex <> -1 Then
    If tbCodigo.Text <> "" Then
    If tbDOI.Text <> "" Then
    If tbNombre.Text <> "" Then
        strSQL = "UPDATE SOCIO SET ID_GRUPO_FK = " & cmbGrupo.List(cmbGrupo.ListIndex, 1) & _
        ", CODIGO_SOCIO = '" & tbCodigo.Text & "', DOI = '" & tbDOI.Text & "', NOMBRE_SOCIO = '" & _
        tbNombre.Text & "' WHERE ID_SOCIO = " & idSocio
        
        OpenDB
        On Error GoTo Handle:
        cnn.Execute (strSQL)
        On Error GoTo 0
        
        Set rs = Nothing
        
        busqSocio.ActualizarHoja
        busqSocio.ActualizarLista
        
        Unload Me
    Else
        MsgBox "Nombre Vacio"
    End If
    Else
        MsgBox "DOI Vacio"
    End If
    Else
        MsgBox "Codigo de Socio Vacio"
    End If
    Else
        MsgBox "Grupo no Valido"
    End If
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btGuardar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
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
    
    Set rs = Nothing
    
    strSQL = "SELECT * FROM SOCIO" & _
    " WHERE ID_SOCIO = " & idSocio
    
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    If rs.RecordCount > 0 Then
        i = 0
        tbNombre.Text = rs.Fields("NOMBRE_SOCIO")
        tbCodigo.Text = rs.Fields("CODIGO_SOCIO")
        tbDOI.Text = rs.Fields("DOI")
        While i < cmbGrupo.ListCount
            If cmbGrupo.List(i, 1) = CInt(rs.Fields("ID_GRUPO_FK")) Then
                cmbGrupo.ListIndex = i
            End If
            i = i + 1
        Wend
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
