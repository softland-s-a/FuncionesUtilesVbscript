Public Function DeleteBatchFile(ByVal path As String, ByVal filename As String)

On Error Resume Next
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    
    FSO.DeleteFile (path & filename)
    
    Set FSO = Nothing
    DeleteBatchFile = True
    Exit Function

    If Err.Number <> 0 Then
        DeleteBatchFile = False
    End If
On Error GoTo 0

End Function