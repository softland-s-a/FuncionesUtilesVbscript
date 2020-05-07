Function ResuelvoSegunType(RdField)
    'Recibe como parametro un campo de un resultset y lo convierte al formato requerido por softland para completarlo en pantalla
    Select Case RdField.Type
        Case 11
            ResuelvoSegunType = Year(RdField.Value) * 10000 + Month(RdField.Value) * 100 + Day(RdField.Value)
        Case 2
            ResuelvoSegunType = "0" & RdField.Value
        Case Else
            ResuelvoSegunType = RdField.Value
    End Select
End Function