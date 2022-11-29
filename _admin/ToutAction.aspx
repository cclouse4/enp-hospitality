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
		Dim bFile As Boolean = Utilities.Upload(Request.Files,Server.MapPath("../") & ConfigurationSettings.AppSettings("uploadPath") &"\landingTout")
	catch e As Exception
		'response.write(e.message())
		'response.end()
	end try
	
	Dim sFiles() As String = Utilities.GetUploadedFileNames()
	' data variables	
	Dim p As New LandingTout
	p.id = CInt(Request.Form("id"))
	p.url = db.Prepare(""& Request.Form("url"),false)
	p.name = db.Prepare(""& Request.Form("name"),false)
	p.image = db.Prepare(""& Utilities.GetFilenameByKey("image"),false)
	p.text = db.Prepare(""& Request.Form("text"),false)
	
	' used for template purposes. these control what goes into the database
	'###############################################################################################################################
	Dim sBridgeQueries As String = ""
	Dim SORT_WHERE As String = ""
	Const HAS_SORT_ORDER As Boolean = true			' set this to true to enable sort actions
	Const HAS_FILES As Boolean = true				' set this to true to enable file deletion actions
	Const DEBUG As Boolean = false					' set this to true to output the query that bombed the script. false for live
	Dim sFileFieldName() As String = { "image" } ' name of the file field name, used for file deletion
	Dim sTableName As String = "tblTout"				' name of the database table to perform actions on
	Dim sPrimaryKey As String = "id"				' name of the primary key field in the database
	Dim sRedirectURL As String = "ToutList.aspx"		' url to redirect to once action is complete
	Dim sDetailURL As String = "ToutDetail.aspx"		' url to redirect to detail page
	Dim sColumnNames() As String = { "name","image","url","text","SortOrder" }
	Dim sValues() As String = { p.name,p.image,p.url,p.text,1 }
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