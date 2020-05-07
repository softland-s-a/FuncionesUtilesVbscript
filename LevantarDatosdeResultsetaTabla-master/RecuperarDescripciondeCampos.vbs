Function RecuperarDescripciondeCampo(oField, oTable)
    'Recibe como parametro un nombre de un objeto campo del sistema y su tabla y recupera la descripcion para el valor actual del campo
    Dim sSql, oRd, oRE, sSelectDescripcion, sCampoDescripcion, sCampoaReemplazar
    
    sSql =  SELECT SelectforValidate,Description FROM CWTMFIELDS WHERE FIELDNAME = ' & oField.Name & ' 
    Set oRd = TMInstance.Openresultset(CStr(sSql))
    Do While Not oRd.EOF
        sCampoDescripcion = oRd(Description).Value
        sSelectDescripcion = oRd(SelectforValidate).Value
        oRd.movenext
    Loop
    oRd.Close
    Set oRd = Nothing
    
    If sCampoDescripcion  X.NULLDESCRIPTION Then
        Set oRE = New RegExp
        'Patron para regular expression Encontrar coincidencias donde el texto este entre ! y dentro haya solo letras, numeros o un guion bajo
        oRE.Pattern = ![0-9a-zA-Z_]!
        oRE.Global = True
        Set oMatches = oRE.Execute(sSelectDescripcion)
        
        For Each oMatch In oMatches
            sCampoaReemplazar = replace(oMatch.Value, !, )
            sSelectDescripcion = replace(sSelectDescripcion, oMatch.Value, oTable.rows(oField.row.Number).fields(sCampoaReemplazar).SqlValue)
        Next
    
        Set oRE = Nothing
        
        Set oRd = TMInstance.Openresultset(CStr(sSelectDescripcion))
            
        Do While Not oRd.EOF
            sCampoDescripcion = replace(sCampoDescripcion, Left(sCampoDescripcion, InStr(sCampoDescripcion, .)), )
            sResultado = oRd(sCampoDescripcion).Value
            oRd.movenext
        Loop
        
        oRd.Close
        Set oRd = Nothing
    Else
        sResultado = 
    End If
    
    RecuperarDescripciondeCampo = sResultado
    
End Function
