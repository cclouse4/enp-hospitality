<%@ Import Namespace="System.Collections" %>
<script runat="server">

	function RenderMenu(lngNewSectionID)
		
		'Response.Write("<br><b style='font-size:14px;font-weight:bold;color:#000;'><u>Website Content</u></b><br><Br/>")
		'Response.Write(getCMSContent())
		'Response.Write("<br><a class=""menuitem"" href=""SectionDetail.aspx?Action=Add"">Add New Section</a>")
		Response.Write("<br/>")
'		Response.Write("<b style='font-size:14px;font-weight:bold;color:#000;'><!--<a class=""adminnav"" href=""AdminNavDisplay.aspx?AdminNavDisplay=0"">--><u>Admins</u></b><br/><br/>")
		Response.Write("<h2 style='margin:0px;padding:0px;'><u>Content</u></h2>")		
		Response.Write("<a class='menuitem' href='HomepageDetail.aspx'>Homepage</a><br/>")
		Response.Write("<a class='menuitem' href='ContentList.aspx'>Site Content</a><br/>")
		Response.Write("<a class='menuitem' href='TimelineList.aspx'>Ecosteps Timeline Content</a><br/>")
		'Response.Write("<a class='menuitem' href='ToutList.aspx'>Manage Landing Touts</a><br/>")
		
		'Response.Write("<br/><h2 style='margin:0px;padding:0px;'><u>Venues</u></h2>")		
		'Response.Write("<a class='menuitem' href='VenueList.aspx'>Venue List</a><br/>")
		'Response.Write("<a class='menuitem' href='LandmarkList.aspx'>Manage Landmarks</a><br/>")
		
		'Response.Write("<br/><h2 style='margin:0px;padding:0px;'><u>Flavors</u></h2>")		
		'Response.Write("<a class='menuitem' href='GlossaryList.aspx'>Event Glossary Terms</a><br/>")
		
		'Response.Write("<br/><h2 style='margin:0px;padding:0px;'><u>Media</u></h2>")
		'Response.Write("<a class='menuitem' href='GalleryList.aspx'>Galleries</a><br/>")
		'Response.Write("<a class='menuitem' href='PeopleList.aspx'>People</a><br/>")
				
		If Session("UserType") = 1 Then
			Response.Write("<br/><a class='menuitem' href='UserList.aspx'>Users</a><br/>")
		End If
		
		Response.Write("<br><b><a class=""menuitem"" href=""logout.aspx"">Logout</a></b>")
	end function
	
	public function getCMSContent() 
		Dim db As New DBObject(ConfigurationSettings.AppSettings("connString"))
		Dim items As New System.Collections.ArrayList()
		Dim sOutput As String
		
		db.OpenConnection()
		
		try
			' get the menu content and stuff it in an arraylist object
			db.RunQuery("select * from tblContent order by SectionID, SortOrder")
			while db.InResults()
				Dim s() As String = { db.GetItem("ContentID"), db.GetItem("ContentTitle"), db.GetItem("SectionID"), db.GetItem("SortOrder"), db.GetItem("isLocked") }
				items.Add(s)	
			end while
		catch e As Exception
			Response.Write("Error grabbing menu items")
			Response.End()
		finally
			db.CloseConnection()
		end try		
		
		' start outputting menu items
		' loop through the array list and pull out the root items
		Dim rootItems As New System.Collections.SortedList()
		Dim i As Integer
		Dim v() As String
		
		for i = 0 to items.Count - 1 
			v = items.Item(i)

			if v(2) = "" or v(2) = 0 then
				rootItems.Add(v(0),v)
			end if
		next i
		
		' loop through the new root array and display them along with their children
		Dim k As Integer
		for i = 0 to rootItems.Count - 1
			v = rootItems.GetByIndex(i)
			
			sOutput += "<a href='SectionDetail.aspx?Action=Edit&ContentID="& v(0) &"'>"& v(1) &"</a>"
			if v(4) = "False" then
				sOutput += "&nbsp;&nbsp;&nbsp;&nbsp;<a href='javascript:deleteSection("& v(0) &")'>[x]</a>"
			end if
			sOutput += "<br/>"
			' get the children
			sOutput += getChildItems(items,v(0),0,"")
		next 
		
		return sOutput
		
	end function
	
	public function getChildItems(ByVal list As System.Collections.ArrayList, ByVal iSectionID As Integer, ByVal iLevel As Integer, ByVal sColl As String)
		Dim found As Boolean = false
		Dim k As Integer
		Dim i As Integer
		Dim sOutput As String = sColl
			for k = 0 to list.Count - 1
				' Dim _v() = list(k)
				'if len(_v(2)) > 0 then
				'	if iSectionID = CInt(_v(2)) then
				'		for i = 0 to iLevel
				'			sOutput += "&nbsp;&nbsp;&nbsp;&nbsp;"
				'		next i
				'		if InStr(sOutput,"SectionID="& _v(0)) = 0 then
				'			sOutput += "- <a href='SectionDetail.aspx?Action=Edit&ContentID="& _v(0) &"'>"& _v(1) &"</a>"
				'			if _v(4) = "False" then
				'				sOutput += "&nbsp;&nbsp;&nbsp;&nbsp;<a href='javascript:deleteSection("& _v(0) &")'>[x]</a>"
				'			end if
				'			sOutput += "<br/>"
				'			sOutput = getChildItems(list,_v(0),iLevel + 1,sOutput)
				'		end if
				'	end if
				'end if
			next k
		
		return sOutput
	end function
	
	
	' for select boxes
	public function getCMSSelect(ByVal iSelected As Integer) 
		Dim db As New DBObject(ConfigurationSettings.AppSettings("connString"))
		Dim items As New System.Collections.ArrayList()
		Dim sOutput As String
		
		db.OpenConnection()
		
		try
			' get the menu content and stuff it in an arraylist object
			db.RunQuery("select * from tblContent order by SectionID, SortOrder")
			while db.InResults()
				Dim s() As String = { db.GetItem("ContentID"), db.GetItem("ContentTitle"), db.GetItem("SectionID"), db.GetItem("SortOrder"), db.GetItem("isLocked") }
				items.Add(s)	
			end while
		catch e As Exception
			Response.Write("Error grabbing menu items")
			Response.End()
		finally
			db.CloseConnection()
		end try		
		
		' start outputting menu items
		' loop through the array list and pull out the root items
		Dim rootItems As New System.Collections.SortedList()
		Dim i As Integer
		Dim v() As String
		
		for i = 0 to items.Count - 1 
			v = items.Item(i)

			if v(2) = "" or v(2) = 0 then
				rootItems.Add(v(0),v)
			end if
		next i
		
		' loop through the new root array and display them along with their children
		Dim k As Integer
		for i = 0 to rootItems.Count - 1
			v = rootItems.GetByIndex(i)
			
			sOutput += "<option value='"& v(0) &"'"
			if v(0) = iSelected then
				sOutput += " selected=selected"
			end if
			sOutput += ">"& v(1) &"</option>"
			' get the children
			sOutput += getChildSelect(items,v(0),0,"",iSelected)
		next 
		
		return sOutput
		
	end function
	
	public function getChildSelect(ByVal list As System.Collections.ArrayList, ByVal iSectionID As Integer, ByVal iLevel As Integer, ByVal sColl As String, ByVal iSelected As Integer)
		Dim found As Boolean = false
		Dim k As Integer
		Dim i As Integer
		Dim sOutput As String = sColl
			for k = 0 to list.Count - 1
				'Dim _v() = list(k)
				'if len(_v(2)) > 0 then
				'	if iSectionID = CInt(_v(2)) then
				'		if InStr(sOutput,"ContentID="& _v(0)) = 0 then
				'			sOutput += "<option value='"& _v(0) &"'"
				'			if _v(0) = iSelected then
				'				sOutput += " selected=selected"
				'			end if
				'			sOutput += ">"
				'			for i = 0 to iLevel
				'				sOutput += "&nbsp;&nbsp;&nbsp;&nbsp;"
				'			next i
				'			sOutput += _v(1) &"</option>"
				'			sOutput = getChildSelect(list,_v(0),iLevel + 1,sOutput,iSelected)
				'		end if
				'	end if
				'end if
			next k
		
		return sOutput
	end function
	
	
</script>
