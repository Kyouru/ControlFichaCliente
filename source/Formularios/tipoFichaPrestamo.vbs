Private Sub btAceptar_Click()
    OpenDB
    
    If cmbTipoFicha.List(cmbTipoFicha.ListIndex, 1) = 1 Then
    
        'strSQL = "SELECT * FROM FICHA_PRESTAMO FP INNER JOIN (SELECT ID_PRESTAMO_FK, MAX(FECHA_INGRESA) AS MFECHAI FROM FICHA_PRESTAMO GROUP BY ID_PRESTAMO_FK) MFP ON MFP.ID_PRESTAMO_FK = FP.ID_PRESTAMO_FK AND MFP.MFECHAI = FP.FECHA_INGRESA" & _
        '        " WHERE ID_TIPO_FICHA_FK = 1 AND FP.ID_FICHA_FK = " & idFicha & " AND FP.ID_PRESTAMO_FK <> " & idPrestamo & " AND FP.ANULADO = FALSE"

        strSQL = "SELECT * FROM (SELECT FP.ID_PRESTAMO_FK, FP.ID_FICHA_FK, MFP.MAXFF, MAX(FP.FECHA_INGRESA) AS MAXFI FROM FICHA_PRESTAMO FP RIGHT JOIN (SELECT ID_PRESTAMO_FK AS MAXIDPREST, MAX(FECHA_FICHA_P) AS MAXFF FROM FICHA_PRESTAMO GROUP BY ID_PRESTAMO_FK) MFP ON MFP.MAXIDPREST = FP.ID_PRESTAMO_FK AND MFP.MAXFF = FP.FECHA_FICHA_P GROUP BY FP.ID_PRESTAMO_FK, FP.ID_FICHA_FK, MFP.MAXFF) FP LEFT JOIN FICHA_PRESTAMO FP2 ON FP2.ID_PRESTAMO_FK = FP.ID_PRESTAMO_FK AND FP2.ID_FICHA_FK = FP.ID_FICHA_FK AND FP2.FECHA_FICHA_P = FP.MAXFF AND FP2.FECHA_INGRESA = FP.MAXFI" & _
                " WHERE FP2.ID_TIPO_FICHA_FK = 1 AND FP2.ID_FICHA_FK = " & idFicha & " AND FP2.ANULADO = FALSE"
    
        On Error GoTo Handle:
        rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
        On Error GoTo 0
        
        If rs.RecordCount > 0 Then
            MsgBox "Ficha Original ya se encuentra tomada"
            closeRS
            Exit Sub
        End If
    End If
    
    Set rs = Nothing
    
    strSQL = "INSERT INTO FICHA_PRESTAMO (ID_FICHA_FK, ID_PRESTAMO_FK, FECHA_FICHA_P, ID_TIPO_FICHA_FK, USUARIO_INGRESA, FECHA_INGRESA) VALUES (" & idFicha & ", " & idPrestamo & ", #" & FechaFicha & "#, " & cmbTipoFicha.List(cmbTipoFicha.ListIndex, 1) & ", '" & Application.UserName & "', #" & Format(Now(), "yyyy-MM-dd HH:mm:ss") & "#)"
    
    On Error GoTo Handle:
    cnn.Execute strSQL
    On Error GoTo 0
    
    closeRS
    
    ActualizarMain
    Unload Me
    
    busqPrestamo.ActualizarHoja
    busqPrestamo.ActualizarLista
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btAceptar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Private Sub btCancelar_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    OpenDB
    
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
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - UserForm_Initialize", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub
