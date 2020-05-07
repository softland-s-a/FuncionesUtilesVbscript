'Decalarar variables publicas
Public oTranslate, oMessage, oInstance

'Agregar al principio del codigo, funcion para instanciar objeto de traducciones
InstanceTranslate

With oInstance.DataAccess
	If .PerformedOperations.Count = 0 Then 'Si oPerformedoperations da 0 quiere decir que no genero registracion
		If sErrorMessage & vbNullString = vbNullString Then 'Quiere decir que no se genero registracion y no pasa bien el error
			'Traduzco ultimo error detectado en pantalla
			sErrorMessage = TraduzcoError
			If sErrorMessage & vbNullString = vbNullString Then 'Si traduciendo el error, sigue vacio, doy mensaje al usuario
				sErrorMessage = "Se produjo un error al generar asiento de provisiones, intente generarlo manualmente para verificar la descripcion del error. "
			End If
			
			'sErrorMessage queda con el mejor error posible para loguearlo
		End If
	Else 'Si no hay error guardo datos del comprobante generado
		Dim oKeys, oKey, oPerform, lIndice
			For Each oPerform In .PerformedOperations
				Set oKeys = oPerform.Keys
				For lIndice = 1 To oKeys.Count
					Select Case Right(oKeys(lIndice).Name, 6)
						Case "MODFOR"
							m_sModForProv = CStr(oKeys(lIndice).Value)
						Case "CODFOR"
							m_sCodForProv = CStr(oKeys(lIndice).Value)
						Case "NROFOR"
							m_lNroForProv = CLng(oKeys(lIndice).Value)
					End Select
				Next
			Next
		
		'En las variables m_sModForProv, m_sCodForProv, m_sNroForProv quedan los datos del comprobante generado
			
	End If
End With
    
	
Function InstanceTranslate()
    Set WSHShell = CreateObject("WScript.Shell")
    Set oTranslate = CreateObject("GRWTranslate.GRWTraducciones")
    oTranslate.DatabasePath = WSHShell.CurrentDirectory & "\..\..\Language\Language.mdb"
    
    Set WSHShell = Nothing
End Function

Function TraduzcoError()
    Dim lContador
    For lContador = oInstance.Messagecount To 1 Step -1 'Recorro objeto de errores desde el ultimo mensaje al primero
        If oInstance.Messages(CLng(lContador)).Description <> "" Then
            sErrorMessage = oInstance.Messages(CLng(lContador)).Description
            sErrorMessage = oTranslate.Translate(sErrorMessage)
            sErrorMessage = CStr(Replace(sErrorMessage, "'", "-"))
            TraduzcoError = sErrorMessage
            Exit Function
        End If
    Next
End Function
