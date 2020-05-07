Function Generofmt(sPath As String, sCodtra As String, sTablabase As String, sDatabase As String, sUsername As String, sPassword As String)

'Dim Who As Object
'Dim Who As cwTMImpl.ObjectInstance

Dim sSql As String, bOkcreate As Boolean, bOkDelete As Boolean


'bcp AdventureWorks2012.HumanResources.Department format nul -c -f Department-c.fmt -T

bOkcreate = CreateBatchFile(sPath, "Generafmt.bat", CStr("bcp " & sDatabase & ".dbo." & sTablabase & " format nul -c -f " & sTablabase & ".fmt -t ; -U " & sUsername & " -P " & sPassword))

Shell (Chr(34) & sPath & "Generafmt.bat" & Chr(34))

'bOkDelete = DeleteBatchFile(CStr(sPath), "Generafmt.bat")

'If bOkcreate = False Or bOkDelete = False Then
'    Generofmt = False
'Else
    Generofmt = True
'End If

End Function