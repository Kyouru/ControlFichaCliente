Private Sub btLimpiar_Click()
    tbDOI.Value = ""
    tbNombre.Value = ""
    btBuscar_Click
End Sub

Private Sub btModificar_Click()
    'Solo si se selecciono algun item de la lista
    If ListBox1.ListIndex <> -1 Then
        'Solo si se selecciono algun item de la lista y no es vacio
        If ListBox1.List(ListBox1.ListIndex) <> "" Then
            idAccionista = ListBox1.List(ListBox1.ListIndex)
            modAccionista.Show (0)
        End If
    Else
        MsgBox "Seleccione una entrada"
    End If
End Sub

Private Sub btNuevo_Click()
    newAccionista.Show (0)
End Sub

Private Sub tbDOI_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btBuscar_Click
        KeyCode = 0
    End If
End Sub

Private Sub tbNombre_Change()
    Dim pos As Integer
    pos = tbNombre.SelStart
    tbNombre.Text = UCase(tbNombre.Text)
    tbNombre.SelStart = pos
End Sub

Private Sub tbNombre_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btBuscar_Click
        KeyCode = 0
    End If
End Sub

Private Sub btBuscar_Click()
    ActualizarHoja
    ActualizarLista
End Sub

Private Sub btEliminar_Click()
    'Solo si se selecciono algun item de la lista
    If ListBox1.ListIndex <> -1 Then
        'Solo si se selecciono algun item de la lista y no es vacio
        If ListBox1.List(ListBox1.ListIndex) <> "" Then
            'Confirmacion antes de anular el Prestamo
            Dim resp As Integer
            resp = MsgBox("Esta seguro que desea eliminar este accionista?", vbYesNo + vbQuestion, ListBox1.List(ListBox1.ListIndex, 3))
            If resp = vbYes Then
            
                OpenDB
                On Error GoTo Handle:
                strSQL = "UPDATE ACCIONISTA SET ANULADO = TRUE WHERE ID_ACCIONISTA = " & ListBox1.List(ListBox1.ListIndex)
                cnn.Execute (strSQL)
                On Error GoTo 0
                
                closeRS
            
                'Actualizar la ListBox
                ActualizarHoja
                ActualizarLista
            End If
        End If
    Else
        MsgBox "Seleccione una entrada"
    End If
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btEliminar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Private Sub UserForm_Initialize()
    
    'Actualizar la Lista de Socios
    ActualizarHoja
    ActualizarLista
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - UserForm_Initialize", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

'Se Solicita todos los Socios que cumplan los filtros y se Copian a una hoja Temporal para luego poder agregarlos a la ListBox
Public Sub ActualizarHoja()
    'Limpiar Hoja
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Range("A1").CurrentRegion.ClearContents
    
    strSQL = "SELECT ID_ACCIONISTA, DOI_ACCIONISTA, NOMBRE_ACCIONISTA, ANULADO FROM ACCIONISTA WHERE 1=1"
    
    If tbNombre.Text <> "" Then
        strSQL = strSQL & " AND NOMBRE_ACCIONISTA LIKE '%" & tbNombre.Text & "%'"
    End If
    If tbDOI.Text <> "" Then
        strSQL = strSQL & " AND DOI_ACCIONISTA LIKE '%" & tbDOI.Text & "%'"
    End If
    
    strSQL = strSQL & " ORDER BY NOMBRE_ACCIONISTA"
    
    
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 1).CurrentRegion.ClearContents
    
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 1) = "ID"
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 2) = "DOI"
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 3) = "NOMBRE"
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 4) = "ANULADO"
    
    'ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 3).EntireColumn.NumberFormat = "@"
    'ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 4).EntireColumn.NumberFormat = "@"
    
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    If rs.RecordCount > 0 Then
        ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Range("A2").CopyFromRecordset rs
    End If
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - ActualizarHoja", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

'Agrega la Hoja Temporal a la ListBox
Public Sub ActualizarLista()
    With ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP)
        ListBox1.ColumnWidths = "30;120;250;60"
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
        
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    closeRS
End Sub
