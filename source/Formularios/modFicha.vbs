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

Private Sub btGerencia_Click()
    newGerencia.Show (0)
End Sub

Private Sub btAccionista_Click()
    newAccionista.Show (0)
End Sub

Private Sub btRepresentante_Click()
    newRepresentante.Show (0)
End Sub

Private Sub btCancelar_Click()
    Unload Me
    If Not desdeHoja Then
        desdeHoja = False
        busqFicha.Show (0)
    Else
        desdeHoja = False
    End If
End Sub

Private Sub btFichaActual_Click()

    On Error Resume Next
    ActiveWorkbook.FollowHyperlink tbRuta
    On Error GoTo 0
Handle:
    If cnn.Errors.count > 0 Then
        'Log del Error
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btFichaActual_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub
Private Sub btGuardar_Click()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim m As Integer
    
    Dim suma As Double
    Dim strAccionistas As String
    Dim strRepresentantes As String
    Dim strGerencias As String
    Dim fechahorasys As Date
    
    strAccionistas = ";"
    suma = 0
    
    'Validar Accionistas
    For i = 1 To 16
        If Me.Controls("cmbAccionista" & i).Visible Then
            If Me.Controls("cmbAccionista" & i).ListIndex <> -1 And Me.Controls("TextBox" & i).Value <> "" Then
                If InStr(strAccionistas, ";" & Me.Controls("cmbAccionista" & i).List(Me.Controls("cmbAccionista" & i).ListIndex, 1) & ";") = 0 Then
                    strAccionistas = strAccionistas & Me.Controls("cmbAccionista" & i).List(Me.Controls("cmbAccionista" & i).ListIndex, 1) & ";"
                    If IsNumeric(Me.Controls("TextBox" & i).Value) Then
                        If CDbl(Me.Controls("TextBox" & i).Value) >= 0 And CDbl(Me.Controls("TextBox" & i).Value) <= 100 Then
                            suma = suma + CDbl(Me.Controls("TextBox" & i).Value)
                        Else
                            MsgBox "Error Valor fuera de rango Accionista " & i
                            Exit Sub
                        End If
                    Else
                        MsgBox "Error Valor no Numerico Accionista " & i
                        Exit Sub
                    End If
                Else
                    MsgBox "Accionista " & i & " Repetido"
                    Exit Sub
                End If
            Else
                MsgBox "Error Accionista " & i
                Exit Sub
            End If
        Else
            Exit For
        End If
    Next i
    
    'Validar Representantes Legales
    For j = 1 To 8
        If Me.Controls("cmbRepresentante" & j).Visible Then
            If Me.Controls("cmbRepresentante" & j).ListIndex <> -1 Then
                If InStr(strRepresentante, ";" & Me.Controls("cmbRepresentante" & j).List(Me.Controls("cmbRepresentante" & j).ListIndex, 1) & ";") = 0 Then
                    strRepresentantes = strRepresentantes & Me.Controls("cmbRepresentante" & j).List(Me.Controls("cmbRepresentante" & j).ListIndex, 1) & ";"
                Else
                    MsgBox "Representante Legal " & j & " Repetido"
                    Exit Sub
                End If
            Else
                MsgBox "Error Representante Legal " & j
                Exit Sub
            End If
        Else
            Exit For
        End If
    Next j
    
    'Validar Gerencia
    valido = True
    For k = 1 To 5
        If Me.Controls("cmbGerencia" & k).Visible Then
            If Me.Controls("cmbGerencia" & k).ListIndex <> -1 Then
                If InStr(strGerencias, ";" & Me.Controls("cmbGerencia" & k).List(Me.Controls("cmbGerencia" & k).ListIndex, 1) & ";") = 0 Then
                    strGerencias = strGerencias & Me.Controls("cmbGerencia" & k).List(Me.Controls("cmbGerencia" & k).ListIndex, 1) & ";"
                Else
                    MsgBox "Gerencia " & k & " Repetido"
                    Exit Sub
                End If
            Else
                MsgBox "Error Gerencia " & k
                Exit Sub
            End If
        Else
            Exit For
        End If
    Next k
    
    If lbTotal.Caption <> "100.00" Then
        MsgBox "No Suma 100%"
        closeRS
        Exit Sub
    End If
        
    If tbFecha.Value = "" Then
        MsgBox "Falta Fecha de la Ficha"
        closeRS
        Exit Sub
    Else
        If Not IsDate(tbFecha.Value) Then
            MsgBox "Fecha Invalida"
            closeRS
            Exit Sub
        End If
    End If
    
    If tbRuta.Text = "" Then
        MsgBox "Ruta Vacia"
        closeRS
        Exit Sub
    Else
        If Dir(tbRuta.Text) = "" Then
            MsgBox "No existe archivo " & tbRuta.Text
            closeRS
            Exit Sub
        End If
    End If
    
    fechahorasys = Now()
    
    OpenDB
    
    strSQL = "INSERT INTO FICHA_MOD (ID_FICHA_FK, FECHA_FICHA, FECHA_MODIFICA, USUARIO_MODIFICA, EXTENSION) VALUES " & _
                            "(" & idFicha & ", #" & fechaStrStr(tbFecha.Value) & "#, #" & Format(Now(), "yyyy-MM-dd HH:mm:ss") & "#, '" & Application.UserName & "', '" & Split(Split(tbRuta.Text, "\")(UBound(Split(tbRuta.Text, "\"))), ".")(UBound(Split(Split(tbRuta.Text, "\")(UBound(Split(tbRuta.Text, "\"))), "."))) & "')"
    
    On Error GoTo Handle:
    cnn.Execute strSQL
    
    strSQL = "SELECT @@IDENTITY"
    
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    idMod = rs.Fields(0).Value
    
    
    strSQL = "UPDATE FICHA_PRESTAMO SET FECHA_FICHA_P = #" & fechaStrStr(tbFecha.Value) & "# WHERE ID_FICHA_FK = " & idFicha
    
    On Error GoTo Handle:
    cnn.Execute strSQL
    On Error GoTo 0
    
    
    'Flag Integrantes Antiguos
        
    strSQL = "UPDATE FICHA_ACCIONISTA SET ID_FICHA_MOD_SIGUIENTE = " & idMod & " WHERE ID_FICHA_FK = " & idFicha & " AND ID_FICHA_MOD_SIGUIENTE = 0"
    On Error GoTo Handle:
    cnn.Execute strSQL
    On Error GoTo 0
    
    strSQL = "UPDATE FICHA_GERENCIA SET ID_FICHA_MOD_SIGUIENTE = " & idMod & " WHERE ID_FICHA_FK = " & idFicha & " AND ID_FICHA_MOD_SIGUIENTE = 0"
    On Error GoTo Handle:
    cnn.Execute strSQL
    On Error GoTo 0
    
    strSQL = "UPDATE FICHA_REPRESENTANTE SET ID_FICHA_MOD_SIGUIENTE = " & idMod & " WHERE ID_FICHA_FK = " & idFicha & " AND ID_FICHA_MOD_SIGUIENTE = 0"
    On Error GoTo Handle:
    cnn.Execute strSQL
    On Error GoTo 0
    
    
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    Call fso.CopyFile(tbRuta.Text, ActiveWorkbook.Path & "\RECURSOS\" & idMod & "." & Split(Split(tbRuta.Text, "\")(UBound(Split(tbRuta.Text, "\"))), ".")(UBound(Split(Split(tbRuta.Text, "\")(UBound(Split(tbRuta.Text, "\"))), "."))), 1)
    
    Set rs = Nothing
    
    For m = 1 To i - 1
        strSQL = "INSERT INTO FICHA_ACCIONISTA (ID_FICHA_FK, ID_ACCIONISTA_FK, PARTICIPACION, ID_FICHA_MOD_SIGUIENTE) VALUES " & _
                "(" & idFicha & ", " & Me.Controls("cmbAccionista" & m).List(Me.Controls("cmbAccionista" & m).ListIndex, 1) & ", " & Me.Controls("TextBox" & m) & ", 0)"
        On Error GoTo Handle:
        cnn.Execute strSQL
        On Error GoTo 0
    Next m
    
    For m = 1 To k - 1
        strSQL = "INSERT INTO FICHA_GERENCIA (ID_FICHA_FK, ID_GERENCIA_FK, ID_FICHA_MOD_SIGUIENTE) VALUES " & _
                "(" & idFicha & ", " & Me.Controls("cmbGerencia" & m).List(Me.Controls("cmbGerencia" & m).ListIndex, 1) & ", 0)"
        On Error GoTo Handle:
        cnn.Execute strSQL
        On Error GoTo 0
    Next m
    
    For m = 1 To j - 1
        strSQL = "INSERT INTO FICHA_REPRESENTANTE (ID_FICHA_FK, ID_REPRESENTANTE_FK, ID_FICHA_MOD_SIGUIENTE) VALUES " & _
                "(" & idFicha & ", " & Me.Controls("cmbRepresentante" & m).List(Me.Controls("cmbRepresentante" & m).ListIndex, 1) & ", 0)"
        On Error GoTo Handle:
        cnn.Execute strSQL
        On Error GoTo 0
    Next m
    
    ActualizarMain
    MsgBox "Registro Exitoso"
    
    Unload Me
    
    If Not desdeHoja Then
        desdeHoja = False
        busqFicha.Show (0)
    Else
        desdeHoja = False
    End If
    
Handle:
    If cnn.Errors.count > 0 Then
        'Log del Error
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btGuardar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Private Sub btMasA_Click()
    If CInt(nAccionistas.Caption) < 16 Then
        nAccionistas.Caption = CInt(nAccionistas.Caption) + 1
        Me.Controls("cmbAccionista" & nAccionistas.Caption).Visible = True
        Me.Controls("TextBox" & nAccionistas.Caption).Visible = True
        
        AjustarTopHeight
    End If
End Sub

Private Sub btMenosA_Click()
    If CInt(nAccionistas.Caption) > 1 Then
        Me.Controls("cmbAccionista" & nAccionistas.Caption).Visible = False
        Me.Controls("TextBox" & nAccionistas.Caption).Visible = False
        nAccionistas.Caption = CInt(nAccionistas.Caption) - 1
        
        AjustarTopHeight
    End If
End Sub

Private Sub btMasG_Click()
    If CInt(nGerencia.Caption) < 5 Then
        nGerencia.Caption = CInt(nGerencia.Caption) + 1
        Me.Controls("cmbGerencia" & nGerencia.Caption).Visible = True
        
        AjustarTopHeight
    End If
End Sub

Private Sub btMenosG_Click()
    If CInt(nGerencia.Caption) > 1 Then
        Me.Controls("cmbGerencia" & nGerencia.Caption).Visible = False
        nGerencia.Caption = CInt(nGerencia.Caption) - 1
        
        AjustarTopHeight
    End If
End Sub

Private Sub btMasRL_Click()
    If CInt(nRepresentanteLegal.Caption) < 8 Then
        nRepresentanteLegal.Caption = CInt(nRepresentanteLegal.Caption) + 1
        Me.Controls("cmbRepresentante" & nRepresentanteLegal.Caption).Visible = True
        
        AjustarTopHeight
    End If
End Sub

Private Sub btMenosRL_Click()
    If CInt(nRepresentanteLegal.Caption) > 1 Then
        Me.Controls("cmbRepresentante" & nRepresentanteLegal.Caption).Visible = False
        nRepresentanteLegal.Caption = CInt(nRepresentanteLegal.Caption) - 1
        
        AjustarTopHeight
    End If
End Sub

Private Sub AjustarTopHeight()
    Dim inicioHeightForm As Integer
    Dim inicioCancelarTop As Integer
    Dim inicioGuardarTop As Integer
    
    inicioHeightForm = 210
    inicioCancelarTop = 150
    inicioGuardarTop = 150
    
    Dim maxCmb As Integer
    maxCmb = CInt(nGerencia.Caption)
    If maxCmb < CInt(nRepresentanteLegal.Caption) Then
        maxCmb = CInt(nRepresentanteLegal.Caption)
    End If
    If maxCmb < CInt(nAccionistas.Caption) - 3 Then
        maxCmb = CInt(nAccionistas.Caption) - 3
    End If
    
    btCancelar.Top = inicioCancelarTop + 30 * (maxCmb - 1)
    btGuardar.Top = inicioGuardarTop + 30 * (maxCmb - 1)
    Me.Height = inicioHeightForm + 30 * (maxCmb - 1)
    
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
    
    actualizarAccionistas
    
    strSQL = "SELECT * " & _
    "FROM ((((FICHA F LEFT JOIN SOCIO S ON F.ID_SOCIO_FK = S.ID_SOCIO) " & _
    " LEFT JOIN (SELECT FM.ID_FICHA_FK, FM.ID_FICHA_MOD, FM.FECHA_FICHA, FM.FECHA_MODIFICA, FM.USUARIO_MODIFICA, EXTENSION FROM FICHA_MOD FM RIGHT JOIN (SELECT ID_FICHA_FK, MAX(FECHA_MODIFICA) AS ULTMOD FROM FICHA_MOD WHERE ANULADO = FALSE GROUP BY ID_FICHA_FK) AS FFMAX ON FFMAX.ULTMOD = FM.FECHA_MODIFICA AND FFMAX.ID_FICHA_FK = FM.ID_FICHA_FK) AS FFM ON FFM.ID_FICHA_FK = F.ID_FICHA)" & _
    " LEFT JOIN FICHA_ACCIONISTA FA ON FA.ID_FICHA_FK = F.ID_FICHA)" & _
    " LEFT JOIN ACCIONISTA A ON A.ID_ACCIONISTA = FA.ID_ACCIONISTA_FK)" & _
    " WHERE ID_FICHA = " & idFicha & _
    " AND FA.ID_FICHA_MOD_SIGUIENTE = 0"
    OpenDB
    
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    If rs.RecordCount > 0 Then
        idSocio = rs.Fields("ID_SOCIO_FK")
        idMod = rs.Fields("ID_FICHA_MOD")
        
        lbNombre.Caption = lbNombre.Caption & rs.Fields("NOMBRE_SOCIO")
        lbCodigo.Caption = lbCodigo.Caption & rs.Fields("CODIGO_SOCIO")
        lbDOI.Caption = lbDOI.Caption & rs.Fields("DOI")
        
        tbFecha.Value = Format(rs.Fields("FECHA_FICHA"), "DD/MM/YYYY")
        
        tbRuta.Value = ActiveWorkbook.Path & Application.PathSeparator & "RECURSOS" & Application.PathSeparator & idMod & "." & rs.Fields("EXTENSION").Value
                
        If rs.Fields("NOMBRE_ACCIONISTA") <> "" Then
            cmbAccionista1.Value = rs.Fields("NOMBRE_ACCIONISTA")
        End If
        
        If rs.Fields("PARTICIPACION") <> "" Then
            TextBox1.Value = rs.Fields("PARTICIPACION")
        End If
        
        rs.MoveNext
        i = 2
        Do While Not rs.EOF
            Me.Controls("cmbAccionista" & i).Visible = True
            Me.Controls("TextBox" & i).Visible = True
            Me.Controls("cmbAccionista" & i).Value = rs.Fields("NOMBRE_ACCIONISTA")
            Me.Controls("TextBox" & i).Value = rs.Fields("PARTICIPACION")
            nAccionistas.Caption = i
            i = i + 1
            rs.MoveNext
        Loop
    End If
    
    actualizarGerencias
    actualizarRepresentantes
    
    OpenDB
    
    'Gerencias
    strSQL = "SELECT * " & _
    "FROM ((FICHA F LEFT JOIN FICHA_GERENCIA FG ON FG.ID_FICHA_FK = F.ID_FICHA) " & _
    "LEFT JOIN GERENCIA G ON G.ID_GERENCIA = FG.ID_GERENCIA_FK) " & _
    "WHERE ID_FICHA = " & idFicha & _
    " AND ID_FICHA_GERENCIA IS NOT NULL" & _
    " AND ID_GERENCIA IS NOT NULL" & _
    "   AND FG.ID_FICHA_MOD_SIGUIENTE = 0"
    
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    If rs.RecordCount > 0 Then
        i = 1
        Do While Not rs.EOF
            Me.Controls("cmbGerencia" & i).Visible = True
            Me.Controls("cmbGerencia" & i).Value = rs.Fields("NOMBRE_GERENCIA")
            nGerencia.Caption = i
            i = i + 1
            rs.MoveNext
        Loop
    End If
    
    'Representantes
    strSQL = "SELECT * " & _
    "FROM ((FICHA F LEFT JOIN FICHA_REPRESENTANTE FR ON FR.ID_FICHA_FK = F.ID_FICHA) " & _
    "LEFT JOIN REPRESENTANTE R ON R.ID_REPRESENTANTE = FR.ID_REPRESENTANTE_FK) " & _
    "WHERE ID_FICHA = " & idFicha & _
    " AND ID_FICHA_REPRESENTANTE IS NOT NULL" & _
    " AND ID_REPRESENTANTE IS NOT NULL" & _
    "   AND FR.ID_FICHA_MOD_SIGUIENTE = 0"
    
    Set rs = Nothing
    
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    If rs.RecordCount > 0 Then
        i = 1
        Do While Not rs.EOF
            Me.Controls("cmbRepresentante" & i).Visible = True
            Me.Controls("cmbRepresentante" & i).Value = rs.Fields("NOMBRE_REPRESENTANTE")
            nRepresentanteLegal.Caption = i
            i = i + 1
            rs.MoveNext
        Loop
    End If
    
    AjustarTopHeight
    
Handle:
    If cnn.Errors.count > 0 Then
        'Log del Error
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - UserForm_Initialize", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
    
    actualizarTotal
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
            If Me.Controls("cmbAccionista" & i).ListIndex <> -1 Then
                arr(i) = Me.Controls("cmbAccionista" & i).List(Me.Controls("cmbAccionista" & i).ListIndex, 0)
            End If
            Me.Controls("cmbAccionista" & i).Clear
        Next i
        i = 0
        Do While Not rs.EOF
            For j = 1 To 16
                Me.Controls("cmbAccionista" & j).AddItem rs.Fields("NOMBRE_ACCIONISTA")
                Me.Controls("cmbAccionista" & j).List(i, 1) = rs.Fields("ID_ACCIONISTA")
            Next j
            i = i + 1
            rs.MoveNext
        Loop
        
        For i = 1 To 16
            Me.Controls("cmbAccionista" & i).Value = arr(i)
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

Public Sub actualizarGerencias()
    Dim i As Integer
    Dim arr(1 To 5) As String
    i = 0
    
    strSQL = "SELECT * FROM GERENCIA WHERE ANULADO = FALSE AND ID_SOCIO_FK = " & idSocio & " ORDER BY NOMBRE_GERENCIA"
    
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    If rs.RecordCount > 0 Then
        For i = 1 To 5
            If Me.Controls("cmbGerencia" & i).ListIndex <> -1 Then
                arr(i) = Me.Controls("cmbGerencia" & i).List(Me.Controls("cmbGerencia" & i).ListIndex, 0)
            End If
            Me.Controls("cmbGerencia" & i).Clear
        Next i
        i = 0
        Do While Not rs.EOF
            For j = 1 To 5
                Me.Controls("cmbGerencia" & j).AddItem rs.Fields("NOMBRE_GERENCIA")
                Me.Controls("cmbGerencia" & j).List(i, 1) = rs.Fields("ID_GERENCIA")
            Next j
            i = i + 1
            rs.MoveNext
        Loop
        
        For i = 1 To 5
            Me.Controls("cmbGerencia" & i).Value = arr(i)
        Next i
    End If
    
Handle:
    If cnn.Errors.count > 0 Then
        'Log del Error
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - actualizarGerencia", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Public Sub actualizarRepresentantes()
    Dim i As Integer
    Dim arr(1 To 8) As String
    i = 0
    
    strSQL = "SELECT * FROM REPRESENTANTE WHERE ANULADO = FALSE AND ID_SOCIO_FK = " & idSocio & " ORDER BY NOMBRE_REPRESENTANTE"
    
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    If rs.RecordCount > 0 Then
        For i = 1 To 8
            If Me.Controls("cmbRepresentante" & i).ListIndex <> -1 Then
                arr(i) = Me.Controls("cmbRepresentante" & i).List(Me.Controls("cmbRepresentante" & i).ListIndex, 0)
            End If
            Me.Controls("cmbRepresentante" & i).Clear
        Next i
        i = 0
        Do While Not rs.EOF
            For j = 1 To 8
                Me.Controls("cmbRepresentante" & j).AddItem rs.Fields("NOMBRE_REPRESENTANTE")
                Me.Controls("cmbRepresentante" & j).List(i, 1) = rs.Fields("ID_REPRESENTANTE")
            Next j
            i = i + 1
            rs.MoveNext
        Loop
        
        For i = 1 To 8
            Me.Controls("cmbRepresentante" & i).Value = arr(i)
        Next i
    End If
    
Handle:
    If cnn.Errors.count > 0 Then
        'Log del Error
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - actualizarRepresentante", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
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


