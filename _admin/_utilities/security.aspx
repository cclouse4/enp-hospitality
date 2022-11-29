<script runat="server">

	Const IMAGE_NAME = 0
	Const IMAGE_CODE = 1
	
	Dim arrSecurityImageCode(14, 1)
	
    function login_validate()
	
	
		If Session("LoggedIn") = 1 Then
			' Response.Write("Login=true")
		Else
			Response.Redirect("login.aspx?logged_out=t")
		End If
		
	
	end function
	
	Function bIsCaseEqual(strStringCompare1, strStringCompare2)
	
		Dim i
		
		bIsCaseEqual = True
		
		' Double Check Lengths
		If Not (Len(strStringCompare1) = Len(strStringCompare2)) Then
			bIsCaseEqual = False
			Exit Function
		End If
			
		For i = 1 To Len(strStringCompare1)
			If Asc(Mid(strStringCompare1, i, i)) <> Asc(Mid(strStringCompare2, i, i)) Then
				bIsCaseEqual = False
				Exit Function
			End If
		Next
	End Function
	
	function Pick_Images(intIndex)
	
		return arrSecurityImageCode(intIndex, IMAGE_NAME)
		
	end function
	
	function Pick_Code(intIndex)
	
		return arrSecurityImageCode(intIndex, IMAGE_CODE)
		
	end function
	
	function Load_Images()
	
		arrSecurityImageCode(0, IMAGE_NAME) = "image1.jpg"
		arrSecurityImageCode(0, IMAGE_CODE) = "JHyGV"
		arrSecurityImageCode(1, IMAGE_NAME) = "image2.jpg"
		arrSecurityImageCode(1, IMAGE_CODE) = "H3AGV"
		arrSecurityImageCode(2, IMAGE_NAME) = "image3.jpg"
		arrSecurityImageCode(2, IMAGE_CODE) = "MeyGL"
		arrSecurityImageCode(3, IMAGE_NAME) = "image4.jpg"
		arrSecurityImageCode(3, IMAGE_CODE) = "BKde8z"
		arrSecurityImageCode(4, IMAGE_NAME) = "image5.jpg"
		arrSecurityImageCode(4, IMAGE_CODE) = "VMJ2"
		arrSecurityImageCode(5, IMAGE_NAME) = "image6.jpg"
		arrSecurityImageCode(5, IMAGE_CODE) = "z7w5M"
		arrSecurityImageCode(6, IMAGE_NAME) = "image7.jpg"
		arrSecurityImageCode(6, IMAGE_CODE) = "A7Ln"
		arrSecurityImageCode(7, IMAGE_NAME) = "image8.jpg"
		arrSecurityImageCode(7, IMAGE_CODE) = "Hzs6K"
		arrSecurityImageCode(8, IMAGE_NAME) = "image9.jpg"
		arrSecurityImageCode(8, IMAGE_CODE) = "GFTV3L"
		arrSecurityImageCode(9, IMAGE_NAME) = "image10.jpg"
		arrSecurityImageCode(9, IMAGE_CODE) = "M7T4HA"
		arrSecurityImageCode(10, IMAGE_NAME) = "image11.jpg"
		arrSecurityImageCode(10, IMAGE_CODE) = "yc6cc"
		arrSecurityImageCode(11, IMAGE_NAME) = "image12.jpg"
		arrSecurityImageCode(11, IMAGE_CODE) = "KTVue"
		arrSecurityImageCode(12, IMAGE_NAME) = "image13.jpg"
		arrSecurityImageCode(12, IMAGE_CODE) = "yHM2B6"
		arrSecurityImageCode(13, IMAGE_NAME) = "image14.jpg"
		arrSecurityImageCode(13, IMAGE_CODE) = "zyXs54"
		arrSecurityImageCode(14, IMAGE_NAME) = "image15.jpg"
		arrSecurityImageCode(14, IMAGE_CODE) = "5eBKL"
		
	end function
	
	function strFixSQL(sql)
	
		dim strTempText
		dim i, badChars(6)
		
		strTempText = Trim(Replace(sql,"'","''"))
		
		badChars(0) = "select"
		badChars(1) = "drop"
		badChars(2) = ";"
		badChars(3) = "--"
		badChars(4) = "insert"
		badChars(5) = "delete"
		badChars(6) = "xp_"

		for i = 0 to uBound(badChars) 
			strTempText = replace(strTempText, badChars(i), "") 
		next 
		
		
		strFixSQL = strTempText
	
	End function
	
	function strFixSQLNum(sql)
	
		dim strTempText
		dim i, badChars(6)
		
		strTempText = Replace(sql,"'","''")
		
		badChars(0) = "select"
		badChars(1) = "drop"
		badChars(2) = ";"
		badChars(3) = "--"
		badChars(4) = "insert"
		badChars(5) = "delete"
		badChars(6) = "xp_"

		for i = 0 to uBound(badChars) 
			strTempText = replace(strTempText, badChars(i), "") 
		next 
		
		If Not (IsNumeric(strTempText) Or strTempText = "") Then
			
			Response.Write("Invalid Input")
			Response.End
		
		End If
		
		strFixSQLNum = strTempText
	
	End function
	
	
	Public Function strIsEmpty(ByVal argString)
		strIsEmpty = CBool((Trim(argString & "")) = "")
	End Function

	
	Public Function CleanSQL(argString, argType)
	
			If (strIsEmpty(argString)) Then
					CleanSQL = ""
					Exit Function
			End If
			
			dim i, badChars(6)
			dim strTempText
			
			strTempText = Trim(argString)
			
			badChars(0) = "select"
			badChars(1) = "drop"
			badChars(2) = ";"
			badChars(3) = "--"
			badChars(4) = "insert"
			badChars(5) = "delete"
			badChars(6) = "xp_"
	
			for i = 0 to uBound(badChars) 
				strTempText = replace(strTempText, badChars(i), "") 
			next 
						
			'Clean up SQL
			If (LCase(strTempText) = "null") Then                                                            ' Nulls
					CleanSQL = Trim(strTempText)
			Else
			
				Select Case LCase(argType)
					Case "i" 		

						If strTempText <> "" Then
							If IsNumeric(strTempText) Then
								CleanSQL = CLng(strTempText)      ' Int
							Else
								CleanSQL = ""
							End If
						Else
							CleanSQL = ""
						End If
						
					Case "d" 
					
					
						If strTempText <> "" Then
							If IsNumeric(strTempText) Then
								CleanSQL = CDbl(strTempText)      ' Double
							Else
								CleanSQL = ""
							End If
						Else
							CleanSQL = ""
						End If
						
					Case Else
						CleanSQL = strTempText  ' Alpha
						CleanSQL = Replace(CleanSQL, "--", " ")
						CleanSQL = Replace(CleanSQL, "==", " ")
						CleanSQL = Replace(CleanSQL, ";",  " ")
						CleanSQL = Replace(CleanSQL, "'",  "''")
				End Select
			End If
			
	End Function

</script>