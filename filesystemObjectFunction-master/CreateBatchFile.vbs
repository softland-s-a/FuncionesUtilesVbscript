Public Function CreateBatchFile(ByVal path As String, ByVal filename As String, bcpcommand As String) As Boolean

On Error Resume Next
Dim FSO As FileSystemObject
Dim F1 As TextStream
Set FSO = New FileSystemObject
Set F1 = FSO.OpenTextFile(path & filename, ForWriting, True, TristateFalse)

'' what ever you want to write in it
F1.WriteLine "cd " & Chr(34) & path & Chr(34)
F1.WriteLine bcpcommand
F1.WriteBlankLines 1

F1.Close


Set F1 = Nothing
Set FSO = Nothing
CreateBatchFile = True
Exit Function

If Err.Number <> 0 Then
    CreateBatchFile = False
End If

On Error GoTo 0

End Function