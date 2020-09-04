
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
        strSQL = "UPDATE GERENCIA SET NOMBRE_GERENCIA = '" & tbNombre.Text & "'"

        If tbDOI.Text <> "" Then
            strSQL = strSQL & ", DOI_GERENCIA = '" & tbDOI.Text & "'"
        Else
            strSQL = strSQL & ", DOI_GERENCIA = NULL"
        End If
        
        If tbCargo.Text <> "" Then
            strSQL = strSQL & ", CARGO_GERENCIA = '" & tbCargo.Text & "'"
        Else
            strSQL = strSQL & ", CARGO_GERENCIA = NULL"
        End If
        
        If Not IsNumeric(tbAno.Text) Then
            strSQL = strSQL & ", ANTIGUEDAD = NULL"
        Else
            If CInt(tbAno.Text) > CInt(Format(Now(), "yyyy")) Then
                strSQL = strSQL & ", ANTIGUEDAD = NULL"
            Else
                If CInt(tbAno.Text) < 1850 Then
                    strSQL = strSQL & ", ANTIGUEDAD = NULL"
                Else
                    strSQL = strSQL & ", ANTIGUEDAD = " & tbAno.Text
                End If
            End If
        End If
        
        If cmbFormacion.ListIndex <> -1 Then
            strSQL = strSQL & ", ID_FORMACION_FK = " & cmbFormacion.List(cmbFormacion.ListIndex, 1)
        Else
            strSQL = strSQL & ", ID_FORMACION_FK = NULL"
        End If
        
        strSQL = strSQL & " WHERE ID_GERENCIA = " & idGerencia
        
        OpenDB
        
        On Error GoTo Handle:
        cnn.Execute (strSQL)
        On Error GoTo 0
        
        Set rs = Nothing
        
        mGerencia.ActualizarHoja
        mGerencia.ActualizarLista
        
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
    Dim i As Integer
    
    strSQL = "SELECT * FROM FORMACION WHERE ANULADO = FALSE"
    
    OpenDB
    
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    If rs.RecordCount > 0 Then
        i = 0
        Do While Not rs.EOF
            Me.cmbFormacion.AddItem rs.Fields("NOMBRE_FORMACION")
            Me.cmbFormacion.List(i, 1) = rs.Fields("ID_FORMACION")
            i = i + 1
            rs.MoveNext
        Loop
    End If
    
    Set rs = Nothing
    
    strSQL = "SELECT * FROM GERENCIA G LEFT JOIN FORMACION F ON F.ID_FORMACION = G.ID_FORMACION_FK" & _
    " WHERE ID_GERENCIA = " & idGerencia
    
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    If rs.RecordCount > 0 Then
        tbNombre.Text = rs.Fields("NOMBRE_GERENCIA")
        
        If Not IsNull(rs.Fields("DOI_GERENCIA")) Then
            tbDOI.Text = rs.Fields("DOI_GERENCIA")
        End If
        
        If Not IsNull(rs.Fields("ID_FORMACION_FK")) Then
            cmbFormacion.Text = rs.Fields("NOMBRE_FORMACION")
        End If
        
        If Not IsNull(rs.Fields("CARGO_GERENCIA")) Then
            tbCargo.Text = rs.Fields("CARGO_GERENCIA")
        End If
        
        If Not IsNull(rs.Fields("ANTIGUEDAD")) Then
            tbAno.Text = rs.Fields("ANTIGUEDAD")
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



