<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="iso-8859-1" Debug="true" %>
<%@ Import Namespace="Chemistry" %>
<!--#include file="_utilities/security.aspx" -->
<!--#include file="includes/menu_functions.aspx" -->
<!--#include file="includes/sitewide.aspx" -->
<%
	' page variables
	Dim sPageData As String = ""
	Dim i As Integer
	' template variables -- These control the page fucntionality and display
	Dim sPageType As String = "List"					' Use Detail, List, Delete
	Dim sPageObject As String = "Image"					' The type of page, used for text display to identify what is going on
	Dim sFormActionURL As String = "GalleryAction.aspx"	' filename of the action page used for this module
	Dim sFormDetailURL As String = "GalleryDetail.aspx"	' filename of the detail page used for this module
	Dim sPrimaryKey As String = "id"				' primary key of the table associated with this module
	Dim sTableName As String = "tblGallery"				' database table name
	Dim bCanAdd As Boolean = true
	Dim iRecCount As Integer = 0
	Dim sColumnHeaders() As String = { "Image","Type","Sort","Action" }
	' data specific vars
	Dim iID As Integer
	Dim sTitle As String
	Dim sThumbnail As String
	Dim sType() As String = { "","Cake","Events","Food" }
	Dim t As Integer = CInt(Request.QueryString("t"))
	Dim img As String
	
	if t = 0 then
		t = 1
	end if
	
	sPageData += "<script>function change() { window.location = 'GalleryList.aspx?t=' + $('#ttt option:selected').val(); } $(document).ready(function() { $('#titleCell').append('<div style=""float:right;padding-right: 100px;padding-top:10px;clear:right;"">Show Gallery: <select name=t id=ttt onchange=""change();""><option value=1"
	if t = 1 then
		sPageData += " selected=selected"
	end if
	sPageData += ">Cake</option><option value=2"
	if t = 2 then
		sPageData += " selected=selected"
	end if
	sPageData += ">Events</option><option value=3"
	if t = 3 then
		sPageData += " selected=selected"
	end if
	sPageData += ">Food</option></select></div>'); });</script>"
	
	Dim db As New DBObject(ConfigurationSettings.AppSettings("connString"))
	db.OpenConnection()
	
	db.RunQuery("select * from "& sTableName &" where type = "& t &" order by SortOrder asc")
	
	while db.InResults()
		iID = db.GetItem(sPrimaryKey)
		sTitle = db.GetItem("image")
		sThumbnail = db.GetItem("thumbnail")
		
		sPageData += "<tr><th scope='row' class='specalt' width='500'>"
		
		if sThumbnail <> "" Then
			sPageData += "<img src='../upload/gallery/thumbs/"& sThumbnail &"' height='100'/>"
		Else
			sPageData += "<img src='../upload/gallery/"& sTitle &"' height='100'/>"
		End If
		
		sPageData += "</th><th class='specalt'>"& sType(db.GetItem("type")) &"</th>"
		if iRecCount = 0 then
			if db.NumRows() > 1 then
	        	sPageData += "<th class='specalt' width='100'><a href='"& sFormActionURL &"?Action=Sort&"& sPrimaryKey &"="& iID &"&SortDirection=1&t="& t &"'><img src='images/arrowDown.gif' border='0'></a> "
			else
				sPageData += "<th class='specalt' width='100'></th>"
			end if
        else if iRecCount = db.NumRows() - 1 then
           sPageData += "<th class='specalt' width='100'><a href='"& sFormActionURL &"?Action=Sort&"& sPrimaryKey &"="& iID &"&SortDirection=0&t="& t &"'><img src='images/arrowUp.gif' border='0'/></a></th>"
        else
        	sPageData += "<th class='specalt' width='100'><a href='"& sFormActionURL &"?Action=Sort&"& sPrimaryKey &"="& iID &"&SortDirection=1&t="& t &"'><img src='images/arrowDown.gif' border='0'></a> <a href='"& sFormActionURL &"?Action=Sort&"& sPrimaryKey &"="& iID &"&SortDirection=0&t="& t &"'><img src='images/arrowUp.gif' border='0'/></a></th>"
        end if
		sPageData += "<th class='specalt' width='250'><a href='"& sFormDetailURL &"?Action=Edit&"& sPrimaryKey &"="& iID &"'>edit</a> | <a href='javascript:DeleteListing("& iID &");'>delete</a>"

		sPageData += "</th>"
		iRecCount += 1
	end while	
	db.CloseConnection()
	
	if len(sPageData) = 0 then
		sPageData = "<tr><th scope=row' class='specalt' colspan='4'><i>There are no "& sPageObject &"s found in the database.</i></th></tr>"
	end if
%>
<!--#include file="templates/List.aspx" -->