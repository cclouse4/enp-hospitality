<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="iso-8859-1" Debug="True" %>
<%@ Import Namespace="Chemistry" %>
<!--#include file="includes/sitewide.aspx" -->
<!--#include file="includes/image.aspx" -->
<!--#include file="../_src/classes.aspx"-->
<!--#include file="../_utilities/functions.aspx"-->
<%
	' page variables - do not change
	Dim sAction As String = Request.Form("Action")
	Dim iSortOrder As Integer = CInt(Request.QueryString("SortID"))
	Dim iSortDirection As Integer = Request.QueryString("SortDirection")
	Dim iSortID As Integer
	Dim iOldSortOrder As Integer
	Dim i As Integer
	Dim db As New DBObject(ConfigurationSettings.AppSettings("connString"))
	db.OpenConnection()
	' data variables
	Dim sTitle As String = strProcessSQL(Request.Form("title"))
	Dim sPassword As String = strProcessSQL(Request.Form("password"))
	Dim iUserType As Integer = Cint(Request.Form("type"))
	Dim iActive As Integer = Cint(Request.Form("active"))
	
	'Dim sFiles() As String = UploadFiles("upload/directory",Request.Files)
	Dim sPostFieldName As String = Request.QueryString("Field")
	' used for template purposes. these control what goes into the database
	'###############################################################################################################################
	Dim sBridgeQueries As String = ""
	Const HAS_SORT_ORDER As Boolean = false			' set this to true to enable sort actions
	Const SORT_WHERE As String = ""
	Const HAS_FILES As Boolean = false				' set this to true to enable file deletion actions
	Const DEBUG As Boolean = false				' set this to true to output the query that bombed the script. false for live
	Dim sFileFieldName() As String = { "", "" } ' name of the file field name, used for file deletion
	Dim sTableName As String = "tblUser"				' name of the database table to perform actions on
	Dim sPrimaryKey As String = "UserID"				' name of the primary key field in the database
	Dim sRedirectURL As String = "UserList.aspx"		' url to redirect to once action is complete
	Dim sDetailURL As String = "UserDetail.aspx"		' url to redirect to detail page
	Dim sColumnNames() As String = { "Username","Password","UserType","Active" }
	Dim sValues() As String = { sTitle,sPassword,iUserType,iActive }
	'###############################################################################################################################
	' format the variables if there is a query string
	Dim iID As Integer = CInt(Request.Form(sPrimaryKey))
	if len(Request.QueryString(sPrimaryKey)) > 0 then
		iID = CInt(Request.QueryString(sPrimaryKey))
	end if
	
	if len(Request.QueryString("Action")) > 0 then
		sAction = Request.QueryString("Action")
	end if
	
	' bridge queries (seperate with | for splitting later into arrays and processing)
	'if iLocation > 0 then
	'	sBridgeQueries += "delete from brgLocation where DirectoryID = "& iID &" | "
	'	sBridgeQueries += "insert into brgLocation (DirectoryID,LocationID) values ("& iID &","& iLocation &") | "
		
	'end if
	
	' perform the actions
%>
<!--#include file="templates/Action.aspx" -->