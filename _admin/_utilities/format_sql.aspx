<script runat="server">

	' String Functions for Building SQL Statements

    ' ********************
	' Name: FormatBoolean
	' Function: processes a string (data variable) to build a SQL query.  Assumes a boolean value (-1 or 0)
	'	If the value is emptystring, then it assumes false or 0
	' Author: N/A
	' Last Update:
	' ********************
   Function strFormatBoolean(strFieldName)
   	  
	  
	  If IsDBNull(strFieldName) Then
	  	  strFormatBoolean = 0
      Else
	  	  strFieldName = CStr(strFieldName)
		  
		  If strFieldName = "" Then
			 strFormatBoolean = 0
		  Else
			 strFormatBoolean = CInt(strFieldName)
		  End If
	  End IF
	  
   End Function

  
    ' ********************
	' Name: FormatDate
	' Function: processes a string (data variable) to build a SQL query.  Makes field Null if the string is an emptystring
	'   Handles data values, will use the pound signs if you indicate that it is Access
	' Author: N/A
	' Last Update:
	' ********************
   Function strFormatDate(strFieldName, intDatabaseType)
   	  If strFieldName = "" Then
	  	 strFormatDate = "NULL"
	  Else
	  	 If intDatabaseType = 0 Then
	  	 	strFormatDate = "'" & strFieldName & "'"  ' SQL Server
		 Else
		 	strFormatDate = "#" & strFieldName & "#"  ' Access
		 End If
	  End If
   End Function

    ' ********************
	' Name: FormatNumber
	' Function: processes a string to build a SQL query.  Makes field Null if the string is an emptystring
	'   only use this function if you have a number field.  It eliminates dollar signs and commas.
	' Author: N/A
	' Last Update:
	' ********************
   Function strFormatNumber(strFieldName)
   
   	   Dim strTempString

	   If strFieldName = "" Then
	   	  strFormatNumber = "NULL"
		  Exit Function
	   End If

	   strTempString = strFieldName
 	   strTempString = Replace(strTempString,",", "")
	   strTempString = Replace(strTempString,"$", "")
	   
	   If Not IsNumeric(strTempString) Then
		  strTempString = "NULL"
	   End If
	   strFormatNumber = strTempString
	   
   End Function
   
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
			strFieldName = CStr(strFieldName)
			
	   	  	If strFieldName = "" Then
		  	 	strProcessSQL = "NULL"
		  	Else
		  	 	strProcessSQL = "'" & Replace(strFieldName, "'", "''") & "'"
		  	End If
		End If
		
   End Function

    
    ' ********************
	' Name: BooleanConversion
	' Function: Change a checkbox value to 0 or 1 for database insertion.
	' Author: Michael Yunn
	' Last Update: 1/17/2006
	' ********************
  Function BooleanConversion(strFieldName)
   	  If strFieldName = "" Then
	  	 BooleanConversion = 0
	  Else
	  	 BooleanConversion = 1
	  End If
   End Function
   
  Function BooleanToggleConversion(bValue)
   	  If bValue = "1" Then
	  	 BooleanToggleConversion = "0"
	  Else
	  	 BooleanToggleConversion = "1"
	  End If
   End Function
   
</script>