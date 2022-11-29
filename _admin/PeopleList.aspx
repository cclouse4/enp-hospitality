<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="iso-8859-1" %>
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
	Dim sPageObject As String = "People"					' The type of page, used for text display to identify what is going on
	Dim sFormActionURL As String = "PeopleAction.aspx"	' filename of the action page used for this module
	Dim sFormDetailURL As String = "PeopleDetail.aspx"	' filename of the detail page used for this module
	Dim sPrimaryKey As String = "id"				' primary key of the table associated with this module
	Dim sTableName As String = "tblPeople"				' database table name
	Dim bCanAdd As Boolean = true
	Dim iRecCount As Integer = 0
	
	Dim sColumnHeaders() As String = { "Name","Sort","Action" }
	' data specific vars
	Dim iID As Integer
	Dim sTitle As String
	
	Dim db As New DBObject(ConfigurationSettings.AppSettings("connString"))
	db.OpenConnection()
	
	db.RunQuery("select * from "& sTableName &" order by SortOrder asc")
	
	while db.InResults()
		iID = db.GetItem(sPrimaryKey)
		sTitle = db.GetItem("fname") &" "& db.GetItem("lname")
		
		sPageData += "<tr><th scope='row' class='specalt' width='500'>"& sTitle &"</th>"
		if iRecCount = 0 then
			if db.NumRows() > 1 then
	        	sPageData += "<th class='specalt' width='100'><a href='"& sFormActionURL &"?Action=Sort&"& sPrimaryKey &"="& iID &"&SortDirection=1'><img src='images/arrowDown.gif' border='0'></a> "
			else
				sPageData += "<th class='specalt' width='100'></th>"
			end if
        else if iRecCount = db.NumRows() - 1 then
           sPageData += "<th class='specalt' width='100'><a href='"& sFormActionURL &"?Action=Sort&"& sPrimaryKey &"="& iID &"&SortDirection=0'><img src='images/arrowUp.gif' border='0'/></a></th>"
        else
        	sPageData += "<th class='specalt' width='100'><a href='"& sFormActionURL &"?Action=Sort&"& sPrimaryKey &"="& iID &"&SortDirection=1'><img src='images/arrowDown.gif' border='0'></a> <a href='"& sFormActionURL &"?Action=Sort&"& sPrimaryKey &"="& iID &"&SortDirection=0'><img src='images/arrowUp.gif' border='0'/></a></th>"
        end if
		sPageData += "<th class='specalt' width='250'><a href='"& sFormDetailURL &"?Action=Edit&"& sPrimaryKey &"="& iID &"'>edit</a> | <a href='javascript:DeleteListing("& iID &")'>delete</a>"

		sPageData += "</th>"
		
		iRecCount += 1
	end while	
	db.CloseConnection()
	
	if len(sPageData) = 0 then
		sPageData = "<tr><th scope=row' class='specalt' colspan='4'><i>There are no "& sPageObject &"s found in the database.</i></th></tr>"
	end if
%>
<!--#include file="templates/List.aspx" -->