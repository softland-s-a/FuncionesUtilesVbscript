Function ValidoRegistry()
    '***Consulto cada valor para indicar error***
    If UCase(m_sShortDate) <> "DD/MM/YYYY" Then
        ValidoRegistry = "El Formato de la fecha del Sistema Operativo debe ser dd/mm/yyyy"
        oStruct.Result = "STOP_IT"
        Exit Function
    Else
        If m_sAutoFlushCache <> "1" Then
            ValidoRegistry = "Debe tener habilitado el tilde de No Utilizar el Cache de acceso a Base de Datos"
            oStruct.Result = "STOP_IT"
            Exit Function
        Else
            If m_sDecimal <> "." Then
                ValidoRegistry = "El Separador de Decimales debe ser " & Chr(34) + "." + Chr(34) + " (Punto)"
                oStruct.Result = "STOP_IT"
                Exit Function
            Else
                If m_strSystemDelimiter <> ";" Then
                    ValidoRegistry = "El Separador de listas debe ser " & Chr(34) + ";" + Chr(34) + " (Punto y Coma)"
                    oStruct.Result = "STOP_IT"
                    Exit Function
                End If
            End If
        End If
    End If
End Function
