<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="iso-8859-1" Debug="True" %>
<%@ Import Namespace="Chemistry" %>
<!--#include file="includes/sitewide.aspx" -->
<!--#include file="../_src/classes.aspx"-->
<%
	' page variables - do not change
	Dim sAction As String = Request.Form("Action")
	Dim iSortOrder As Integer = CInt(Request.QueryString("SortID"))
	Dim iSortDirection As Integer = Request.QueryString("SortDirection")
	Dim iSortID As Integer
	Dim iOldSortOrder As Integer
	Dim i As Integer
	Dim sPostFieldName As String = Request.QueryString("Field")
	Dim db As New DBObject(ConfigurationSettings.AppSettings("connString"))
	db.OpenConnection()
	

	' upload files
	try
		Dim bFile As Boolean = Utilities.Upload(Request.Files,Server.MapPath("../") & ConfigurationSettings.AppSettings("uploadPath") &"\venues")
	catch e As Exception
		'response.write(e.message())
		'response.end()
	end try
	
	Dim sFiles() As String = Utilities.GetUploadedFileNames()
	' data variables	
	Dim v As New Help
	v.name = db.Prepare(""& Request.Form("name"))
	v.url = db.Prepare(""& Request.Form("url"))
	
	' used for template purposes. these control what goes into the database
	'###############################################################################################################################
	Dim sBridgeQueries As String = ""
	Const SORT_WHERE As String = ""
	Const HAS_SORT_ORDER As Boolean = true			' set this to true to enable sort actions
	Const HAS_FILES As Boolean = false				' set this to true to enable file deletion actions
	Const DEBUG As Boolean = false					' set this to true to output the query that bombed the script. false for live
	Dim sFileFieldName() As String = { "" } ' name of the file field name, used for file deletion
	Dim sTableName As String = "tblHelpBar"				' name of the database table to perform actions on
	Dim sPrimaryKey As String = "id"				' name of the primary key field in the database
	Dim sRedirectURL As String = "HelpList.aspx"		' url to redirect to once action is complete
	Dim sDetailURL As String = "HelpDetail.aspx"		' url to redirect to detail page
	Dim sColumnNames() As String = { "name","url","SortOrder" }
	Dim sValues() As String = { v.name,v.url,1 }
	'###############################################################################################################################

	' format the variables if there is a query string
	Dim iID As Integer = CInt(Request.Form(sPrimaryKey))
	if len(Request.QueryString(sPrimaryKey)) > 0 then
		iID = CInt(Request.QueryString(sPrimaryKey))
	end if
	
	if len(Request.QueryString("Action")) > 0 then
		sAction = Request.QueryString("Action")
	end if
	
	Dim _s As String
	Dim _SortOrder As String
	
	' perform the actions

%>
<!--#include file="templates/Action.aspx" -->