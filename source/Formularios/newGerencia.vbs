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


Private Sub btGuardar_Click()
    Dim Doi As Boolean
    Dim Cargo As Boolean
    Dim Ano As Boolean
    Dim Formacion As Boolean
    
    Doi = True
    Cargo = True
    Ano = True
    Formacion = True
    
    If tbNombre.Text = "" Then
        MsgBox "Nombre Vacio"
        Exit Sub
    End If
    
    If tbDOI.Text = "" Then
        Doi = False
    End If
    
    If tbCargo.Text = "" Then
        Cargo = False
    End If
    
    If cmbFormacion.ListIndex = -1 Then
        Formacion = False
    End If
    
    If tbAno.Text = "" Then
        Ano = False
    Else
        If Not IsNumeric(tbAno.Text) Then
            Ano = False
            MsgBox "Año Antiguedad no numerico"
            Exit Sub
        Else
            If CInt(tbAno.Text) > CInt(Format(Now(), "yyyy")) Then
                Ano = False
                MsgBox "Año Antiguedad no puede ser en el futuro"
                Exit Sub
            Else
                If CInt(tbAno.Text) < 1850 Then
                    Ano = False
                    MsgBox "Introdusca el Año con el formato correcto (ej. 1980)"
                    Exit Sub
                End If
            End If
        End If
    End If
    
    strSQL = "INSERT INTO GERENCIA (NOMBRE_GERENCIA, " & _
            "CARGO_GERENCIA, ANTIGUEDAD, ID_FORMACION_FK, DOI_GERENCIA, ID_SOCIO_FK)" & _
            "VALUES ('" & tbNombre.Text & "', '" & tbCargo.Text & "', " & tbAno.Text & ", " & cmbFormacion.List(cmbFormacion.ListIndex, 1) & ", '" & tbDOI.Text & "', " & idSocio & ")"
    
    strSQL = "INSERT INTO GERENCIA (NOMBRE_GERENCIA, ID_SOCIO_FK"
    
    If Doi Then
        strSQL = strSQL & ", DOI_GERENCIA"
    End If
    
    If Cargo Then
        strSQL = strSQL & ", CARGO_GERENCIA"
    End If
    
    If Formacion Then
        strSQL = strSQL & ", ID_FORMACION_FK"
    End If
    
    If Ano Then
        strSQL = strSQL & ", ANTIGUEDAD"
    End If
    
    strSQL = strSQL & ") VALUES ('" & tbNombre.Text & "', " & idSocio
    
    If Doi Then
        strSQL = strSQL & ", '" & tbDOI.Text & "'"
    End If
    
    If Cargo Then
        strSQL = strSQL & ", '" & tbCargo.Text & "'"
    End If
    
    If Formacion Then
        strSQL = strSQL & ", " & cmbFormacion.List(cmbFormacion.ListIndex, 1)
    End If
    
    If Ano Then
        strSQL = strSQL & ", " & tbAno.Text
    End If
    
    strSQL = strSQL & ")"
    
    OpenDB
    On Error GoTo Handle:
    cnn.Execute strSQL
    On Error GoTo 0
    
    On Error Resume Next
    newFicha.actualizarGerencias
    On Error Resume Next
    modFicha.actualizarGerencias
    
    On Error Resume Next
    mGerencia.ActualizarHoja
    On Error Resume Next
    mGerencia.ActualizarLista
    On Error GoTo 0
    
    Unload Me
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btGuardar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
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
