Function LevantarDatosdeResultsetaTabla(sNombreTabla, oRd, bGrisarcampos)
    
    Dim oField, oColumn, oTable
    'Levanta datos del resultset recibido como parametro a la tabla.
    'sNombreTabla = Tabla grilla a la que se quiere levantar el dato
    'oRd = Resultset que contiene datos a levantar a la tabla. Los campos del resultset deben tener el mismo nombre que los campos de la tabla!
    'bGrisarcampos = Define si luego de levantar el dato a cada campo se debe grisar el mismo.
    
    Set oTable = TMInstance.Table.rows(1).Tables(sNombreTabla)
    
    Do While Not oRd.EOF
        With oTable.rows
            For Each oField In .Add.fields
                For Each oColumn In oRd.rdoColumns
                    If oColumn.Name = oField.Name Then
                        oField.Enabled = True
                        oField.Value = ResuelvoSegunType(oRd(oColumn.Name))
                        oField.Description = RecuperarDescripciondeCampo(oField, oTable)
                        If bGrisarcampos = True Then
                            oField.Enabled = False
                        End If
                        Exit For
                    End If
                Next
            Next
        End With
        oRd.movenext
    Loop

End Function