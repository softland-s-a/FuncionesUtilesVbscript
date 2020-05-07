Function readFromRegistry(strRegistryKey, strDefault)
    Dim WshShell, value

    On Error Resume Next
    Set WshShell = CreateObject("WScript.Shell")
    value = WshShell.RegRead(strRegistryKey)

    If Err.Number <> 0 Then
        readFromRegistry = strDefault
    Else
        readFromRegistry = value
    End If

    Set WshShell = Nothing
End Function
