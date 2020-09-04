
Private Sub btAgregar_Click()
    Dim myValue As Variant
    myValue = InputBox("Nombre de la Nueva Formacion:", "Nueva Formacion")
    If myValue <> "" Then
    
        OpenDB
        
        strSQL = "SELECT * FROM FORMACION WHERE NOMBRE_FORMACION = '" & UCase(myValue) & "'"
        
        On Error GoTo Handle:
        rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
        On Error GoTo 0
        
        If rs.RecordCount = 0 Then
        
            strSQL = "INSERT INTO FORMACION (NOMBRE_FORMACION) VALUES ('" & UCase(myValue) & "');"
            
            On Error GoTo Handle:
            cnn.Execute strSQL
            On Error GoTo 0
            
            closeRS
            
            ActualizarHoja
            ActualizarLista
        Else
            MsgBox "Formacion ya existe"
        End If
    End If
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btAgregar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btModificar_Click
End Sub

Private Sub btCerrar_Click()
    Unload Me
End Sub

Public Sub ActualizarHoja()
    strSQL = "SELECT ID_FORMACION, NOMBRE_FORMACION FROM FORMACION F WHERE F.ANULADO = FALSE ORDER BY NOMBRE_FORMACION"
    
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Range("A2"), _
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Range("A2").End(xlDown)).ClearContents
    
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 1) = "ID_FORMACION"
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 2) = "NOMBRE_FORMACION"
    
    If rs.RecordCount > 0 Then
        ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Range("A2").CopyFromRecordset rs
    End If
    
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Cells(1, 1).EntireColumn.NumberFormat = "0"
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - ActualizarHoja", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Public Sub ActualizarLista()
    With ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP)
        ListBox1.ColumnWidths = "40;100;"
        ListBox1.ColumnCount = 2
        ListBox1.ColumnHeads = True
        
        If .Range("A3") <> "" Then
            ListBox1.RowSource = .Name & "!A2:B" & .Range("A2").End(xlDown).Row
        Else
            If .Range("A2") <> "" Then
                ListBox1.RowSource = .Name & "!A2:B2"
            Else
                ListBox1.RowSource = ""
                ListBox1.ColumnHeads = False
            End If
        End If
    End With
End Sub

Private Sub btEliminar_Click()
    If ListBox1.ListIndex <> -1 Then
        Dim resp As Integer
        resp = MsgBox("Esta seguro que desea eliminar esta formacion?", vbYesNo + vbQuestion, ListBox1.List(ListBox1.ListIndex, 1))
        If resp = vbYes Then
        
            OpenDB
            On Error GoTo Handle:
            cnn.Execute ("UPDATE FORMACION SET ANULADO = TRUE WHERE ID_FORMACION = " & ListBox1.List(ListBox1.ListIndex, 0))
            On Error GoTo 0
            
            closeRS
            
            ActualizarHoja
            ActualizarLista
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

Private Sub btModificar_Click()
    If ListBox1.ListIndex <> -1 Then
        Dim myValue As Variant
        myValue = InputBox("Nuevo Nombre de la Formacion:", "Modificar Formacion", ListBox1.List(ListBox1.ListIndex, 1))
        If myValue <> "" Then
    
            OpenDB
            
            strSQL = "SELECT * FROM FORMACION WHERE NOMBRE_FORMACION = '" & UCase(myValue) & "'"
            
            On Error GoTo Handle:
            rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
            On Error GoTo 0
            
            If rs.RecordCount = 0 Then
            
                strSQL = "UPDATE FORMACION SET NOMBRE_FORMACION = '" & UCase(myValue) & "' WHERE ID_FORMACION = " & ListBox1.List(ListBox1.ListIndex, 0)
                
                On Error GoTo Handle:
                cnn.Execute strSQL
                On Error GoTo 0
                
                closeRS
                
                ActualizarHoja
                ActualizarLista
            Else
                MsgBox "Formacion ya existe"
            End If
        End If
    Else
        MsgBox "Seleccione una entrada"
    End If
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btModificar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Private Sub UserForm_Initialize()
    ActualizarHoja
    ActualizarLista
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    closeRS
End Sub




