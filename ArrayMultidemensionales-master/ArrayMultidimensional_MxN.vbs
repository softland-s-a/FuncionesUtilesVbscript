Dim oConn,ListaResultado, rd

ListaResultado =  Array()
cNombre_Empresa = "TEMP2"

OpenConnection

sQItem = "SELECT top 5 * FROM VTMCLH"
Set rd = oConn.Execute(CStr(sQItem))


ListaResultado = ArrayMultidimensionalValores

For each itemList in ListaResultado

		sNrocta = itemList(0)
		lImport =  itemList(1)
		sSigno =  itemList(2)

Next


Function  ArrayMultidimensionalValores()
	 Dim aItem, aList
	 Dim IContadorItem, IContadorList

	 aItem = Array()
	 aList = Array()
	 IContadorItem = 0
	 IContadorList = 0

 do While not rd.EOF



		For each id in rd.Fields

			Redim preserve aItem(IContadorItem)
			aItem(IContadorItem) = id.value
			IContadorItem = IContadorItem + 1

		Next

		IContadorItem = 0
		Redim preserve aList(IContadorList)
		aList(IContadorList) = aItem
		IContadorList = IContadorList + 1

		rd.MoveNext
	Loop

	ArrayMultidimensionalValores =  aList

End Function
Sub OpenConnection()
    Set oConn = CreateObject("ADODB.Connection")
    DBProperties.CompanyName = cNombre_Empresa
    oConn.Provider = "sqloledb"
    oConn.Properties("Data Source").Value = DBProperties.Server
    oConn.Properties("Initial Catalog").Value = DBProperties.Database
    oConn.Properties("User ID").Value = DBProperties.User
    oConn.Properties("Password").Value = DBProperties.Password
    oConn.Open
End Sub
Sub CloseConnection()
    oConn.Close
    Set oConn = Nothing
End Sub
