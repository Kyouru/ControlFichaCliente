Private Sub btAtras_Click()
    Unload Me
    busqSocio.Show (0)
End Sub

Private Sub btEliminar_Click()
    'Solo si se selecciono algun item de la lista
    If ListBox1.ListIndex <> -1 Then
        'Solo si se selecciono algun item de la lista y no es vacio
        If ListBox1.List(ListBox1.ListIndex) <> "" Then
            Dim resp As Integer
            Dim fechahorasys As Date
            resp = MsgBox("Esta seguro que desea eliminar esta ficha?", vbYesNo + vbQuestion, "Borrar Ficha")
            If resp = vbYes Then
                OpenDB
                fechahorasys = Now()
    
                strSQL = "UPDATE FICHA SET ANULADO = TRUE, FECHA_ANULADO = #" & Format(fechahorasys, "yyyy-MM-dd HH:mm:ss") & "#, USUARIO_ANULA = '" & Application.UserName & "' WHERE ID_FICHA = " & ListBox1.List(ListBox1.ListIndex)
                On Error GoTo Handle:
                cnn.Execute (strSQL)
                On Error GoTo 0
                
                Set rs = Nothing
                
                ActualizarHoja
                ActualizarLista
            End If
        Else
            MsgBox "Seleccione una entrada con datos"
        End If
    Else
        MsgBox "Seleccione una entrada"
    End If
Handle:
    If cnn.Errors.count > 0 Then
        'Log del Error
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btEliminar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Private Sub btModificar_Click()
    'Solo si se selecciono algun item de la lista
    If ListBox1.ListIndex <> -1 Then
        'Solo si se selecciono algun item de la lista y no es vacio
        If ListBox1.List(ListBox1.ListIndex) <> "" Then
            idFicha = ListBox1.List(ListBox1.ListIndex)
            Unload Me
            modFicha.Show (0)
        End If
    Else
        MsgBox "Seleccione una entrada"
    End If
End Sub

Private Sub btNuevo_Click()
    Unload Me
    newFicha.Show (0)
End Sub

Private Sub btSeleccionar_Click()
    'Solo si se selecciono algun item de la lista
    If ListBox1.ListIndex <> -1 Then
        'Solo si se selecciono algun item de la lista y no es vacio
        If ListBox1.List(ListBox1.ListIndex) <> "" Then
            idFicha = ListBox1.List(ListBox1.ListIndex)
            FechaFicha = Format(ListBox1.List(ListBox1.ListIndex, 1), "YYYY-MM-DD")
            Unload Me
            busqPrestamo.Show (0)
        End If
    Else
        MsgBox "Seleccione una entrada"
    End If
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btSeleccionar_Click
End Sub

Private Sub UserForm_Initialize()

    strSQL = "SELECT DOI, CODIGO_SOCIO, NOMBRE_SOCIO, NOMBRE_GRUPO " & _
    "FROM SOCIO LEFT JOIN GRUPO ON GRUPO.ID_GRUPO = SOCIO.ID_GRUPO_FK" & _
    " WHERE ID_SOCIO = " & idSocio
    
    OpenDB
    
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    If rs.RecordCount > 0 Then
        lbNombre.Caption = lbNombre.Caption & rs.Fields("NOMBRE_SOCIO")
        lbCodigo.Caption = lbCodigo.Caption & rs.Fields("CODIGO_SOCIO")
        lbDOI.Caption = lbDOI.Caption & rs.Fields("DOI")
        lbGrupo.Caption = lbGrupo.Caption & rs.Fields("NOMBRE_GRUPO")
    End If
    
    Set rs = Nothing
    ActualizarHoja
    ActualizarLista
    
Handle:
    If cnn.Errors.count > 0 Then
        'Log del Error
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - UserForm_Initialize", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Public Sub ActualizarHoja()
    strSQL = "SELECT F.ID_FICHA, FFM.FECHA_FICHA, FFM.USUARIO_MODIFICA, FFM.FECHA_MODIFICA FROM ((" & _
    " FICHA F LEFT JOIN SOCIO S ON S.ID_SOCIO = F.ID_SOCIO_FK)" & _
    " LEFT JOIN (SELECT FM.ID_FICHA_FK, FM.FECHA_FICHA, FM.FECHA_MODIFICA, FM.USUARIO_MODIFICA, EXTENSION FROM FICHA_MOD FM RIGHT JOIN (SELECT ID_FICHA_FK, MAX(ID_FICHA_MOD) AS MAXIDFM FROM FICHA_MOD WHERE ANULADO = FALSE GROUP BY ID_FICHA_FK) AS FFMAX ON FFMAX.MAXIDFM = FM.ID_FICHA_MOD AND FFMAX.ID_FICHA_FK = FM.ID_FICHA_FK) AS FFM ON FFM.ID_FICHA_FK = F.ID_FICHA)" & _
    " WHERE S.ID_SOCIO = " & idSocio & _
    " AND F.ANULADO = FALSE" & _
    " ORDER BY FFM.FECHA_FICHA DESC"
    
    With ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP)
        'Limpiar Hoja Temporal
        .Cells(1, 1).CurrentRegion.ClearContents
        
        .Cells(1, 1) = "ID_FICHA"
        .Cells(1, 2) = "FECHA_FICHA"
        .Cells(1, 3) = "USUARIO_MODIFICA"
        .Cells(1, 4) = "FECHA_MODIFICA"
        
        OpenDB
        On Error GoTo Handle:
        rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
        On Error GoTo 0
        
        If rs.RecordCount > 0 Then
            .Range("A2").CopyFromRecordset rs
        End If
        closeRS
        
        .Cells(1, 1).EntireColumn.NumberFormat = "0"
        .Cells(1, 2).EntireColumn.NumberFormat = "DD/MM/YYYY"
        .Cells(1, 4).EntireColumn.NumberFormat = "DD/MM/YYYY"
        
    End With
    
Handle:
    If cnn.Errors.count > 0 Then
        'Log del Error
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - ActualizarHoja", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

'Agrega la Hoja Temporal a la ListBox
Public Sub ActualizarLista()
    With ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP)
        ListBox1.ColumnWidths = "0;70;220;60"
        ListBox1.ColumnCount = 4
        ListBox1.ColumnHeads = True
        'En caso halla mas de una fila
        If .Range("A3") <> "" Then
            ListBox1.RowSource = .Name & "!A2:D" & .Range("A2").End(xlDown).Row
        Else
            'En caso halla solamente una fila
            If .Range("A2") <> "" Then
                ListBox1.RowSource = .Name & "!A2:D2"
            'En caso no hallan datos
            Else
                ListBox1.RowSource = ""
                ListBox1.ColumnHeads = False
            End If
        End If
        
        If idFicha > 0 Then
            For i = 0 To (ListBox1.ListCount - 1)
                If ListBox1.List(i, 0) = idFicha Then
                    ListBox1.ListIndex = i
                    Exit For
                End If
            Next
        Else
            If ListBox1.ListCount > 0 Then
                ListBox1.ListIndex = 0
            End If
        End If
        
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    closeRS
End Sub
