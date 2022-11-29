<script runat="server">

Function buildPagination(ByVal qryPg As String, ByVal RECORDS_PER_PAGE As Integer, ByVal iRecCount As Integer, ByVal iFilterVal As Integer) As String
	Dim iPageNum As Integer
	Dim iStart As Integer
	Dim iEnd As Integer
	Dim iPageCount = System.Math.Ceiling(iRecCount / RECORDS_PER_PAGE)
	Dim iPageCounter As Integer
	Dim iPageIndex As Integer
	Dim sPageContent As String = ""
	
	If qryPg = "" Then
		iPageNum = 1
		iStart = 1
	Else
		iPageNum = qryPg
		iStart = (RECORDS_PER_PAGE * (iPageNum - 1)) + 1
	End If
	
	If RECORDS_PER_PAGE + iStart > iRecCount
		iEnd = iRecCount
	Else
		iEnd = iStart + RECORDS_PER_PAGE - 1
	End If
	
	if iPageCount > 1 then
		For iPageIndex = 1 To iPageCount    
			If Not iPageIndex = 1 Then
				sPageContent += " | "
			End If
			if iPageIndex = 1 then
				' for the first page, see if we are on a filter. Only filters can be active if higher than 1
				if iFilterVal = CInt(Request.QueryString("filter")) then
					' we are on a filter.. check the qrystring for a page value
					if iPageNum = 1 then
						sPageContent += CStr(iPageIndex)
					else
						sPageContent += "<a href='secondary.aspx?id=65&p=0&pg=" & iPageIndex & "&key="& Request.QueryString("key") &"&filter="& iFilterVal &"'>"& iPageIndex & "</a>"
					end if
				else
					' no filter.. page 1 is active
					sPageContent += CStr(iPageIndex)
				end if
			else
				' page num could be 2.. see if there is a filter for this obj
				if CInt(Request.QueryString("filter")) = iFilterVal then
					' there is a filter.. only make active the page that is in the query string
					if iPageNum <> iPageIndex then
						sPageContent += "<a href='secondary.aspx?id=65&p=0&pg=" & iPageIndex & "&key="& Request.QueryString("key") &"&filter="& iFilterVal &"'>"& iPageIndex & "</a>"
					else
						sPageContent += CStr(iPageIndex)
					end if
				else
					sPageContent += "<a href='secondary.aspx?id=65&p=0&pg=" & iPageIndex & "&key="& Request.QueryString("key") &"&filter="& iFilterVal &"'>"& iPageIndex & "</a>"
				end if
			end if
			
			iPageCounter = iPageCounter + 1
			
			If iPageCounter > 20 Then
				sPageContent += "<br/>"
				iPageCounter = 0
			End If
		Next
	end if
	return sPageContent
end function

</script>