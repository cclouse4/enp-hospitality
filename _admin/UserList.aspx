<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="iso-8859-1" %>
<%@ Import Namespace="Chemistry" %>
<!--#include file="_utilities/security.aspx" -->
<!--#include file="includes/menu_functions.aspx" -->
<!--#include file="includes/sitewide.aspx" -->
<%

	if Session("UserType") <> 1 then
		Response.Redirect("splash.aspx")
	end if
	
	' page variables
	Dim sPageData As String = ""
	Dim i As Integer
	' template variables -- These control the page fucntionality and display
	Dim sPageType As String = "List"					' Use Detail, List, Delete
	Dim sPageObject As String = "User"					' The type of page, used for text display to identify what is going on
	Dim sFormActionURL As String = "UserAction.aspx"	' filename of the action page used for this module
	Dim sFormDetailURL As String = "UserDetail.aspx"	' filename of the detail page used for this module
	Dim sPrimaryKey As String = "UserID"				' primary key of the table associated with this module
	Dim sTableName As String = "tblUser"				' database table name
	Dim bCanAdd As Boolean = true
	
	Dim sColumnHeaders() As String = { "Username","Type","Actions" }
	' data specific vars
	Dim iID As Integer
	Dim sName As String
	Dim sUserType() As String = { "","Administrator", "User" }
	Dim iUserType As Integer
	
	Dim iSortOrder As Integer
	Dim bActive As Boolean
	Dim iNumRecords As Integer
    Dim iRecCount As Integer = 0
	
	Dim db As New DBObject(ConfigurationSettings.AppSettings("connString"))
	db.OpenConnection()
	
	db.RunQuery("select count(*) as mx from "& sTableName)
    while db.InResults()
    	iNumRecords = db.GetItem("mx")
    end while

	
	db.RunQuery("select * from "& sTableName &" order by Username asc")
	while db.InResults()
		iID = db.GetItem(sPrimaryKey)
		sName = db.GetItem("Username")
		iUserType = db.GetItem("UserType")
		
		sPageData += "<tr><th scope='row' class='specalt' width='500'>"& sName &"</th>"
		sPageData += "<th class='specalt' width='200'>"& sUserType(iUserType) &"</th>"
		sPageData += "<th class='specalt' width='200'>&nbsp;&nbsp;<a href='"& sFormDetailURL &"?Action=Edit&"& sPrimaryKey &"="& iID &"'>Edit</a> / <a href='javascript:DeleteListing("& iID &");'>Remove</a>&nbsp;</th>"
		
		iRecCount += 1
	end while	

	db.CloseConnection()
	
	if len(sPageData) = 0 then
		sPageData = "<tr><th scope=row' class='specalt' colspan='3'><i>There are no users found in the database.</i></th></tr>"
	end if
%>
<!--#include file="templates/List.aspx" -->