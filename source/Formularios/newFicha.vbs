Private Sub btBuscar_Click()
    Dim path_rep As String
    path_rep = openDialog
    
    If path_rep <> "FALSO" Then
        tbRuta.Value = path_rep
    End If
End Sub

Private Function openDialog() As String
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

   With fd
        .InitialFileName = "C:\"
      .AllowMultiSelect = False

      ' Set the title of the dialog box.
      .Title = "Por favor la Ficha de Cliente"

      ' Clear out the current filters, and add our own.
      .Filters.Clear
      .Filters.Add "Todos", "*.*"

      ' Show the dialog box. If the .Show method returns True, the
      ' user picked at least one file. If the .Show method returns
      ' False, the user clicked Cancel.
      If .Show = True Then
        openDialog = .SelectedItems(1)
    Else
        openDialog = "FALSO"
      End If
   End With
End Function

Private Sub btCalendario_Click()
    tmpFecha = 0
    frmCalendario.Show
    If tmpFecha > 0 Then
        tbFecha.Text = Format(tmpFecha, "DD/MM/YYYY")
    Else
        tbFecha.Text = ""
    End If
End Sub

Private Sub btCancelar_Click()
    Unload Me
    busqFicha.Show (0)
End Sub

Private Sub btGuardar_Click()
    Dim i As Integer
    Dim j As Integer
    Dim suma As Double
    Dim valido As Boolean
    Dim strAccionistas As String
    valido = True
    strAccionistas = ";"
    suma = 0
    For i = 1 To 16
        If Me.Controls("ComboBox" & i).Visible Then
            If Me.Controls("ComboBox" & i).ListIndex <> -1 And Me.Controls("TextBox" & i).Value <> "" Then
                If InStr(strAccionistas, ";" & Me.Controls("ComboBox" & i).List(Me.Controls("ComboBox" & i).ListIndex, 1) & ";") = 0 Then
                    strAccionistas = strAccionistas & Me.Controls("ComboBox" & i).List(Me.Controls("ComboBox" & i).ListIndex, 1) & ";"
                    If IsNumeric(Me.Controls("TextBox" & i).Value) Then
                        If CDbl(Me.Controls("TextBox" & i).Value) >= 0 And CDbl(Me.Controls("TextBox" & i).Value) <= 100 Then
                            suma = suma + CDbl(Me.Controls("TextBox" & i).Value)
                        Else
                            MsgBox "Error Valor fuera de rango Accionista " & i
                            valido = False
                            Exit For
                        End If
                    Else
                        MsgBox "Error Valor no Numerico Accionista " & i
                        valido = False
                        Exit For
                    End If
                Else
                    MsgBox "Accionista " & i & " Repetido"
                    valido = False
                    Exit For
                End If
            Else
                MsgBox "Error Accionista " & i
                valido = False
                Exit For
            End If
        Else
            Exit For
        End If
    Next i
    
    'Validaciones
    If Not valido Then
        Exit Sub
    End If
    
    If lbTotal.Caption <> "100.00" Then
        MsgBox "No Suma 100%"
        Exit Sub
    End If
        
    If tbFecha.Value = "" Then
        MsgBox "Falta Fecha de la Ficha"
        Exit Sub
    Else
        If Not IsDate(tbFecha.Value) Then
            MsgBox "Fecha Invalida"
            Exit Sub
        End If
    End If
    
    If cmbTipoFicha.ListIndex = -1 Then
        MsgBox "Tipo Ficha Invalida"
        Exit Sub
    End If
    
    If tbRuta.Text = "" Then
        MsgBox "Ruta Vacia"
        Exit Sub
    Else
        If Dir(tbRuta.Text) = "" Then
            MsgBox "No existe archivo " & tbRuta.Text
            Exit Sub
        End If
    End If
    
    strSQL = "INSERT INTO FICHA (ID_PRESTAMO_FK, ID_TIPO_FICHA_FK, FECHA_FICHA, FECHA_INGRESO, USUARIO, EXTENSION) VALUES " & _
                            "(" & idPrestamo & ", " & cmbTipoFicha.List(cmbTipoFicha.ListIndex, 1) & ", #" & fechaStrStr(tbFecha.Value) & "#, #" & Format(Now(), "yyyy-MM-dd HH:mm:ss") & "#, '" & Application.UserName & "', '" & Split(Split(tbRuta.Text, "\")(UBound(Split(tbRuta.Text, "\"))), ".")(UBound(Split(Split(tbRuta.Text, "\")(UBound(Split(tbRuta.Text, "\"))), "."))) & "')"
                        
    OpenDB
    On Error GoTo Handle:
    cnn.Execute strSQL
    
    strSQL = "SELECT @@IDENTITY"
    
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    idFicha = rs.Fields(0).Value
    
    
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    Call fso.CopyFile(tbRuta.Text, ActiveWorkbook.Path & "\RECURSOS\" & idFicha & "." & Split(Split(tbRuta.Text, "\")(UBound(Split(tbRuta.Text, "\"))), ".")(UBound(Split(Split(tbRuta.Text, "\")(UBound(Split(tbRuta.Text, "\"))), "."))), 1)
    
    Set rs = Nothing
    
    For j = 1 To i - 1
        strSQL = "INSERT INTO FICHA_ACCIONISTA (ID_FICHA_FK, ID_ACCIONISTA_FK, PARTICIPACION) VALUES " & _
                "(" & idFicha & ", " & Me.Controls("ComboBox" & j).List(Me.Controls("ComboBox" & j).ListIndex, 1) & ", " & Me.Controls("TextBox" & j) & ")"
        On Error GoTo Handle:
        cnn.Execute strSQL
        On Error GoTo 0
    Next j
    
    MsgBox "Registro Exitoso"
    
    Unload Me
    
    busqFicha.Show (0)
    
Handle:
    If cnn.Errors.count > 0 Then
        'Log del Error
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btGuardar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Private Sub btMas_Click()
    If CInt(nAccionistas.Caption) < 16 Then
        nAccionistas.Caption = CInt(nAccionistas.Caption) + 1
        Me.Controls("ComboBox" & nAccionistas.Caption).Visible = True
        Me.Controls("TextBox" & nAccionistas.Caption).Visible = True
        btCancelar.Top = btCancelar.Top + 30
        btGuardar.Top = btGuardar.Top + 30
        Me.Height = Me.Height + 30
    End If
End Sub

Private Sub btMenos_Click()
    If CInt(nAccionistas.Caption) > 1 Then
        Me.Controls("ComboBox" & nAccionistas.Caption).Visible = False
        Me.Controls("TextBox" & nAccionistas.Caption).Visible = False
        nAccionistas.Caption = CInt(nAccionistas.Caption) - 1
        btCancelar.Top = btCancelar.Top - 30
        btGuardar.Top = btGuardar.Top - 30
        Me.Height = Me.Height - 30
    End If
End Sub

Private Sub btAccionista_Click()
    newAccionista.Show (0)
End Sub

Private Sub verificarParticipacion(txtBox As Object)
    If IsNumeric(txtBox.Value) Then
        If CDbl(txtBox.Value) >= 0 And CDbl(txtBox.Value) <= 100 Then
            txtBox.Value = Application.WorksheetFunction.RoundDown(txtBox.Value, 2)
            actualizarTotal
        Else
            MsgBox "Fuera de Rango"
            txtBox.Value = ""
        End If
    Else
        MsgBox "Valor no numerico"
        txtBox.Value = ""
    End If
End Sub

Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    verificarParticipacion Me.TextBox1
End Sub

Private Sub TextBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    verificarParticipacion Me.TextBox2
End Sub

Private Sub TextBox3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    verificarParticipacion Me.TextBox3
End Sub

Private Sub TextBox4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    verificarParticipacion Me.TextBox4
End Sub

Private Sub TextBox5_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    verificarParticipacion Me.TextBox5
End Sub

Private Sub TextBox6_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    verificarParticipacion Me.TextBox6
End Sub

Private Sub TextBox7_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    verificarParticipacion Me.TextBox7
End Sub

Private Sub TextBox8_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    verificarParticipacion Me.TextBox8
End Sub

Private Sub TextBox9_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    verificarParticipacion Me.TextBox9
End Sub

Private Sub TextBox10_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    verificarParticipacion Me.TextBox10
End Sub

Private Sub TextBox11_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    verificarParticipacion Me.TextBox11
End Sub

Private Sub TextBox12_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    verificarParticipacion Me.TextBox12
End Sub

Private Sub TextBox13_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    verificarParticipacion Me.TextBox13
End Sub

Private Sub TextBox14_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    verificarParticipacion Me.TextBox14
End Sub

Private Sub TextBox15_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    verificarParticipacion Me.TextBox15
End Sub

Private Sub TextBox16_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    verificarParticipacion Me.TextBox16
End Sub

Private Sub UserForm_Initialize()
    Dim i As Integer
    strSQL = "SELECT DOI, CODIGO_SOCIO, SOLICITUD, NOMBRE_PRODUCTO, NOMBRE_SOCIO, NOMBRE_MONEDA, MONTO " & _
    "FROM ((PRESTAMO LEFT JOIN SOCIO ON SOCIO.ID_SOCIO = PRESTAMO.ID_SOCIO_FK)" & _
    "LEFT JOIN MONEDA ON MONEDA.ID_MONEDA = PRESTAMO.ID_MONEDA_FK) " & _
    "LEFT JOIN PRODUCTO ON PRODUCTO.ID_PRODUCTO = PRESTAMO.ID_PRODUCTO_FK " & _
    "WHERE ID_PRESTAMO = " & idPrestamo
    
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    If rs.RecordCount > 0 Then
        lbNombre.Caption = lbNombre.Caption & rs.Fields("NOMBRE_SOCIO")
        lbCodigo.Caption = lbCodigo.Caption & rs.Fields("CODIGO_SOCIO")
        lbDOI.Caption = lbDOI.Caption & rs.Fields("DOI")
        lbSolicitud.Caption = lbSolicitud.Caption & rs.Fields("SOLICITUD")
        lbProducto.Caption = lbProducto.Caption & rs.Fields("NOMBRE_PRODUCTO")
        lbMoneda.Caption = lbMoneda.Caption & rs.Fields("NOMBRE_MONEDA")
        lbMonto.Caption = lbMonto.Caption & Format(rs.Fields("MONTO"), "#,##0.00")
    End If
    
    Set rs = Nothing
    
    strSQL = "SELECT * FROM TIPO_FICHA WHERE ANULADO = FALSE"
    
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    If rs.RecordCount > 0 Then
        
        i = 0
        Do While Not rs.EOF
            cmbTipoFicha.AddItem rs.Fields("NOMBRE_TIPO_FICHA")
            cmbTipoFicha.List(i, 1) = rs.Fields("ID_TIPO_FICHA")
            i = i + 1
            rs.MoveNext
        Loop
    End If
    cmbTipoFicha.ListIndex = 0
    
    Set rs = Nothing
    
    actualizarAccionistas
    
Handle:
    If cnn.Errors.count > 0 Then
        'Log del Error
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - UserForm_Initialize", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Public Sub actualizarAccionistas()
    Dim i As Integer
    Dim arr(1 To 16) As String
    i = 0
    
    strSQL = "SELECT * FROM ACCIONISTA WHERE ANULADO = FALSE ORDER BY NOMBRE_ACCIONISTA"
    
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    If rs.RecordCount > 0 Then
        For i = 1 To 16
            If Me.Controls("ComboBox" & i).ListIndex <> -1 Then
                arr(i) = Me.Controls("ComboBox" & i).List(Me.Controls("ComboBox" & i).ListIndex, 0)
            End If
            Me.Controls("ComboBox" & i).Clear
        Next i
        i = 0
        Do While Not rs.EOF
            For j = 1 To 16
                Me.Controls("ComboBox" & j).AddItem rs.Fields("NOMBRE_ACCIONISTA")
                Me.Controls("ComboBox" & j).List(i, 1) = rs.Fields("ID_ACCIONISTA")
            Next j
            i = i + 1
            rs.MoveNext
        Loop
        
        For i = 1 To 16
            Me.Controls("ComboBox" & i).Value = arr(i)
        Next i
    End If
    
Handle:
    If cnn.Errors.count > 0 Then
        'Log del Error
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - actualizarAccionistas", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Private Sub actualizarTotal()
    Dim i As Integer
    Dim suma As Double
    suma = 0
    For i = 1 To 16
        If Me.Controls("TextBox" & i).Visible Then
            If IsNumeric(Me.Controls("TextBox" & i).Value) Then
                suma = suma + CDbl(Me.Controls("TextBox" & i).Value)
            End If
        Else
            Exit For
        End If
    Next i
    lbTotal.Caption = Format(suma, "##0.00")
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    closeRS
End Sub
