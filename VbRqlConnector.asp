<%
	Private Const DCOM = "DCOM"
	Private Const WebService11 = "WebService11"
	Private Const WebService11Url = "http://localhost/CMS/WebService/RqlWebService.svc"

	Class VbRqlConnector
		Private Function InitializeConnectionType()
			If(GetConnectionType() = "") Then
				If(TestConnection(WebService11Url) = "OK") Then
					SetConnectionType(WebService11)
				Else
					SetConnectionType(DCOM)
				End If
			End If
		End Function

		Private Function TestConnection(URL)
			Dim objHTTP
			Dim sHTML
			Set objHTTP = Server.CreateObject ("Microsoft.XMLHTTP")
			objHTTP.open "GET", URL, False
			objHTTP.send
			TestConnection = objHTTP.statusText
		End Function

		Private Function SetConnectionType(ConnectionType)
			Session("RqlConnectionType") = ConnectionType
		End Function

		Private Function GetConnectionType()
			GetConnectionType = Session("RqlConnectionType")
		End Function
		
		Public Function SendRql(Rql)
			Dim RqlResponse
			
			Select Case GetConnectionType()
				Case DCOM
					RqlResponse = SendRqlViaDCOM(Rql)
				Case WebService11
					RqlResponse = SendRqlViaWebService11(Rql)
			End Select
			
			SendRql = RqlResponse
		End Function

		Private Function SendRqlViaWebService11(Rql)
			Rql = "<![CDATA[" & Rql & "]]>"
			Set oXmlHTTP = CreateObject("Microsoft.XMLHTTP")
			oXmlHTTP.Open "POST", WebService11Url, False 
			oXmlHTTP.setRequestHeader "Content-Type", "text/xml; charset=utf-8" 
			oXmlHTTP.setRequestHeader "SOAPAction", "http://tempuri.org/RDCMSXMLServer/action/XmlServer.Execute"
			
			Dim SOAPMessage
			SOAPMessage = ""
			SOAPMessage = SOAPMessage & "<s:Envelope xmlns:s=""http://schemas.xmlsoap.org/soap/envelope/"">"
			SOAPMessage = SOAPMessage & "<s:Body><q1:Execute xmlns:q1=""http://tempuri.org/RDCMSXMLServer/message/""><sParamA>" & Rql  & "</sParamA><sErrorA></sErrorA><sResultInfoA></sResultInfoA></q1:Execute></s:Body>"
			SOAPMessage = SOAPMessage & "</s:Envelope>"
			
			oXmlHTTP.send SOAPMessage 
			SendRqlViaWebService11 = oXmlHTTP.responseXML.xml
		End Function

		Private Function SendRqlViaDCOM(Rql)
			Dim objIO	'Declare the objects
			Dim xmlData, sError, retXml
			Set objIO = Server.CreateObject("RDCMSASP.RdPageData")
			objIO.XmlServerClassName = "RDCMSServer.XmlServer"
			
			xmlData = objIO.ServerExecuteXml (Rql, sError) 

			Set objIO = Nothing
			
			If xmlData = "" Then
				retXml = "<ERRORTEXT>" & sError & "</ERRORTEXT>"
			Else
				retXml = xmlData
			End If
			
			SendRqlViaDCOM = retXml
		End Function
	End Class
%>