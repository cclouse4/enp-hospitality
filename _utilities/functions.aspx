<script runat="server">

    ' ********************
	' Name: ProcessSQL
	' Function: processes a string to build a SQL query.  Makes field Null if the string is an emptystring
	'   Also eliminates problem of single quotes.  It will replace all apostrophes with two apostrophes and wrap the entire statement with apostrophes.
	'   Do not use this with fields that are dates or numbers.
	' Author: N/A
	' Last Update:
	' ********************
   Function strProcessSQL(strFieldName)
	   	  If IsDBNull(strFieldName) Then
		  	 strProcessSQL = "NULL"
		  Else
	   	  	If strFieldName = "" Then
		  	 	strProcessSQL = "NULL"
		  	Else
				If Not IsDBNull(strFieldName) Then
		  	 		strProcessSQL = "'" & Replace(strFieldName, "'", "''") & "'"
				Else
					strProcessSQL = "NULL"
				End If
		  	End If
		End If
   End Function

Function isUnsubscribed(emailAddress)

	' ExactTarget API methods
	' Dim all variables
	
	Dim objSXH
		
	'Eat 'n Park
	Dim strUsername, strPassword
	Dim listID 
	
	
	strUsername = "eatnpark.api"
	strPassword = "enp@api1"
	
	Dim strUrl As String
	Dim strXML As String
	Dim ObjXML 
	Dim subStatus
	
	' Create the full URL
	strUrl = "http://www.exacttarget.com/api/integrate.asp"
	
	
	objSXH = nothing
	
	' Create the HTTP Object
	objSXH = server.CreateObject("msxml2.ServerXMLHttp.3.0")
	objSXH.open("POST", strUrl,false)
	
	' Refresh Groups
	strXML = "<?xml version=""1.0""?><exacttarget><authorization><username>" & strUsername & "</username><password>" & strPassword & "</password></authorization><system><system_name>subscriber</system_name><action>retrieve</action><search_type>listid</search_type><search_value>" & listID & "</search_value><search_value2>" & Replace(emailAddress,"'","") & "</search_value2><showChannelID></showChannelID></system></exacttarget>"
		
	' Post the XML body
	objSXH.setRequestHeader("Content-Type","application/x-www-form-urlencoded")
	objSXH.send ("qf=xml&xml=" & Server.URLEncode(strXML))

	' This step if there is is no error will begin to parse the response from ExactTarget
	If ObjSXH.status = 200 Then
	
		'Create the MSXML Document Object from the response stream
		ObjXML = ObjSXH.responseXML
		subStatus = getXMLValue( ObjXML, "Status", "" )
		'response.write(objSXH.responsexml.xml)
		'response.end
	Else
		'Response.Write "An error occurred." & objSXH.responsexml.xml
'		Response.write "Status " & objSXH.status & "<br>"
'		Response.write "Status text " & objSXH.statustext
	end if
		
	'Clean up objects and remove them from memory
	'objChildren = Nothing
	'objChildnode = Nothing
	' objNodeList = Nothing
	ObjXML = Nothing
	ObjSXH = Nothing

	If subStatus = "Unsubscribed" Then
		isUnsubscribed = true
	Else
		isUnsubscribed = false
	End If
	
End Function

Function getXMLValue( xml, parentNode, node )
	'if you want to save the file, you can use this to do so
	'objXML.Save(Server.MapPath("listadd.xml"))
	Dim objNodeList
	Dim intNodeCount As Integer
	
	' Get the first subscriber in the XML Doc and get the list info
	' Node names are case sensitive
	objNodeList = xml.GetElementsByTagName(parentNode)
	intNodeCount = objNodeList.length
	
	'if no nodes, then there was an error
	if intNodeCount = 0 then
		'response.write objSXH.responsexml.xml
	Else
		' this loop will work through each subscriber node
			getXMLValue = objNodeList.item(0).text
		'response.write "</p>"
	end if

End Function

</script>