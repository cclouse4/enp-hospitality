<script runat="server">
	
function linkName(strTempText)
	Dim strSQL AS String
	Dim strKeyword As String
	Dim intStart As Integer
	Dim intEnd As Integer
	Dim intContentID As Integer
	Dim intSectionID As Integer
	Dim dbo As New DBObject()
	
	' get the keyword
	if InStr(strTempText, "[[") then
		intStart = InStr(strTempText, "[[") + 2
		intEnd = InStr(strTempText, "]]")
		strKeyword = mid(strTempText, intStart, intEnd - intStart)
		
		strSQL = "select * from tblCMSContent where LinkName = '"& strKeyword &"'"
		dbo.RunQuery(strSQL)
		while dbo.Read()
			intSectionID = dbo.GetItem("SectionID")
			intContentID = dbo.GetItem("ContentID")
			
			' replace with the new link
			strTempText = Replace(strTempText, "[["& strKeyword &"]]", "<a href=""secondary.aspx?id="& intSectionID &"&contentID="& intContentID &""">"& dbo.GetItem("Title") &"</a>")
		
		end while
		dbo.EndQuery()
	end if
	
	dbo.CloseConnection()
	return strTempText

end function
	
	Function GetFileType(strFile)
		
		dim aryExtension(27,1)
		
		'Valid web displayable image extensions
		aryExtension(0,0) = "jpg"
		aryExtension(0,1) = "imageDisplay"
		aryExtension(1,0) = "gif"
		aryExtension(1,1) = "imageDisplay"	
		aryExtension(2,0) = "bmp"
		aryExtension(2,1) = "imageDisplay"
		
		'Other valid image extensions
		aryExtension(3,0) = "psd"
		aryExtension(3,1) = "image file"
		aryExtension(4,0) = "tif"
		aryExtension(4,1) = "image file"
		aryExtension(5,0) = "ico"
		aryExtension(5,1) = "image file"
		aryExtension(6,0) = "map"
		aryExtension(6,1) = "image file"
		
		'Valid music extensions
		aryExtension(7,0) = "mp3"
		aryExtension(7,1) = "music file"
		aryExtension(8,0) = "cda"
		aryExtension(8,1) = "music file"
		aryExtension(9,0) = "m3u"
		aryExtension(9,1) = "music file"				
		aryExtension(10,0) = "wma"
		aryExtension(10,1) = "music file"
		aryExtension(11,0) = "ra"
		aryExtension(11,1) = "music file"
		aryExtension(12,0) = "ram"
		aryExtension(12,1) = "music file"
		aryExtension(13,0) = "aif"
		aryExtension(13,1) = "music file"
		aryExtension(14,0) = "au"
		aryExtension(14,1) = "music file"
		aryExtension(15,0) = "mp2"
		aryExtension(15,1) = "music file"
		aryExtension(16,0) = "wav"
		aryExtension(16,1) = "music file"
		
		'Valid video extensions
		aryExtension(17,0) = "mov"
		aryExtension(17,1) = "video file"
		aryExtension(18,0) = "qt"
		aryExtension(18,1) = "video file"
		aryExtension(19,0) = "qtvr"
		aryExtension(19,1) = "video file"
		aryExtension(20,0) = "qtvr"
		aryExtension(20,1) = "video file"
		
		'Valid document extensions
		aryExtension(21,0) = "pdf"
		aryExtension(21,1) = "document"
		aryExtension(22,0) = "wpd"
		aryExtension(22,1) = "document"
		aryExtension(23,0) = "txt"
		aryExtension(23,1) = "document"
		aryExtension(24,0) = "dlc"
		aryExtension(24,1) = "document"
		aryExtension(25,0) = "xls"
		aryExtension(25,1) = "document"
		aryExtension(26,0) = "ppt"
		aryExtension(26,1) = "document"
		aryExtension(27,0) = "doc"
		aryExtension(27,1) = "document"
		
		Dim intFileLength
		Dim intDotLocation
		Dim intExtensionLength
		Dim fileExtension
		
		Dim strExtension
		Dim strExtensionType
		Dim intCounter
		Dim strFinalType
		
		If not strFile = "" Then
			intFileLength = Len(strFile)
			intDotLocation = InStr(strFile, ".")
			intExtensionLength = intFileLength - intDotLocation
			
			If NOT intDotLocation = 0 Then
			   fileExtension = Right(strFile, intExtensionLength) 
			   for intCounter = LBound(aryExtension) to UBound(aryExtension)
			   		strExtension   = aryExtension(intCounter,0)
		       		strExtensionType = aryExtension(intCounter,1)
					if UCase(strExtension) = UCase(fileExtension) then
						strFinalType = strExtensionType
					end if
				next
				If strFinalType = "" Then
				    strFinalType = "file"
				End If	
			Else
			   strFinalType = "file"
			End If
			
		End If
		GetFileType = strFinalType
		
	End Function			
		
		
	Function strRemoveFontsFromText(strText)

		Dim strTempString
		Dim intImgSpot
		Dim strLeftString
		Dim strLeftPop
		Dim strPop 
		
		strTempString = strText
	
		' remove all first parts
		Do Until Instr(UCase(strTempString), "<FONT") = 0 Or Instr(UCase(strTempString), "<FONT") = 0
			' Get to first part.
			intImgSpot = Instr(UCase(strTempString), "<FONT")
			strLeftPop = ""
			strPop = ""
			strLeftString = Right(strTempString, Len(strTempString) - (intImgSpot - 1))
			
			Do Until strLeftString = "" Or strPop = ">"
				strPop = Left(strLeftString, 1)
				strLeftPop = strLeftPop & strPop
				strLeftString = Right(strLeftString, Len(strLeftString) - 1)
			Loop
			strTempString = Replace(strTempString, strLeftPop, "")
		Loop

		' remove all last parts
		Do Until Instr(UCase(strTempString), "</FONT") = 0 Or Instr(UCase(strTempString), "</FONT") = 0
			' Get to first part.
			intImgSpot = Instr(UCase(strTempString), "</FONT")
			strLeftPop = ""
			strPop = ""
			strLeftString = Right(strTempString, Len(strTempString) - (intImgSpot - 1))
			
			Do Until strLeftString = "" Or strPop = ">"
				strPop = Left(strLeftString, 1)
				strLeftPop = strLeftPop & strPop
				strLeftString = Right(strLeftString, Len(strLeftString) - 1)
			Loop
			strTempString = Replace(strTempString, strLeftPop, "")
		Loop
		
		
		'body craziness
		Do Until Instr(UCase(strTempString), "<BODY") = 0 Or Instr(UCase(strTempString), "<BODY") = 0
			' Get to first part.
			intImgSpot = Instr(UCase(strTempString), "<BODY")
			strLeftPop = ""
			strPop = ""
			strLeftString = Right(strTempString, Len(strTempString) - (intImgSpot - 1))
			
			Do Until strLeftString = "" Or strPop = ">"
				strPop = Left(strLeftString, 1)
				strLeftPop = strLeftPop & strPop
				strLeftString = Right(strLeftString, Len(strLeftString) - 1)
			Loop
			strTempString = Replace(strTempString, strLeftPop, "")
		Loop		
		
		
		strRemoveFontsFromText = strTempString
	End Function
	
	Function strRemoveTagFromText(strTag, strText)

		Dim strTempString
		Dim intImgSpot
		Dim strLeftString
		Dim strLeftPop
		Dim strPop 
		
		strTempString = strText
	
		' remove all first parts
		Do Until Instr(UCase(strTempString), "<"&strTag) = 0 Or Instr(UCase(strTempString), "<"&strTag) = 0
			' Get to first part.
			intImgSpot = Instr(UCase(strTempString), "<"&strTag)
			strLeftPop = ""
			strPop = ""
			strLeftString = Right(strTempString, Len(strTempString) - (intImgSpot - 1))
			
			Do Until strLeftString = "" Or strPop = ">"
				strPop = Left(strLeftString, 1)
				strLeftPop = strLeftPop & strPop
				strLeftString = Right(strLeftString, Len(strLeftString) - 1)
			Loop
			strTempString = Replace(strTempString, strLeftPop, "")
		Loop
		strRemoveTagFromText = strTempString
	end function	


	function strip(p_Content)
	
		p_Content = Replace(p_Content, "<HTML>", " ")
		p_Content = Replace(p_Content, "<p>", "")		
		p_Content = Replace(p_Content, "<P>", "<br>")	
		p_Content = Replace(p_Content, "<P class=MsoNormal style=""MARGIN: 0in 0in 0pt"" align=left>", "")
		p_Content = Replace(p_Content, "<P align=left>", "")
		p_Content = Replace(p_Content, "</P>", "")
		p_Content = Replace(p_Content, "</p>", "")
		p_Content = Replace(p_Content, "<HEAD>", " ")

		p_Content = Replace(p_Content, "</HEAD>", " ")						
		p_Content = Replace(p_Content, "<BODY>", " ")								
		p_Content = Replace(p_Content, "</BODY>", " ")				
		p_Content = Replace(p_Content, "</HTML>", " ")			
		p_Content = Replace(p_Content, "<STRONG>", "<b>")	
		p_Content = Replace(p_Content, "</STRONG>", "</b>")	
		p_Content = Replace(p_Content, "Â", "")
		p_Content = Replace(p_Content, "â„¢", "&trade;")
		
		p_Content = strRemoveTagFromText("META", p_Content)
		p_Content = strRemoveTagFromText("HTML", p_Content)				
		p_Content = strRemoveTagFromText("BODY", p_Content)		
		p_Content = strRemoveTagFromText("LINK", p_Content)		
		
		strip = p_Content
	end function

</script>