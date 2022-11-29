<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="iso-8859-1" Debug="True" %>
<%@ Import Namespace="Chemistry" %>
<!--#include file="includes/sitewide.aspx" -->
<!--#include file="../_src/classes.aspx"-->
<%
	' page variables - do not change
	Dim sAction As String = "Edit"'Request.Form("Action")
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
		Dim bFile As Boolean = Utilities.Upload(Request.Files,Server.MapPath("../") & ConfigurationSettings.AppSettings("uploadPath") &"\page")
	catch e As Exception
		'response.write(e.message())
		'response.end()
	end try
	
	Dim sFiles() As String = Utilities.GetUploadedFileNames()
	' data variables	
	Dim h As New Homepage
	h.metaTitle = db.Prepare(""& Request.Form("metaTitle"))
	h.metaKeywords = db.Prepare(""& Request.Form("metaKeywords"))
	h.metaDescription = db.Prepare(""& Request.Form("metaDescription"))
	h.image = db.Prepare(""& Utilities.GetFilenameByKey("image"),false)
	h.header1 = db.Prepare(""& Request.Form("header1"))
	h.header2 = db.Prepare(""& Request.Form("header2"))
	h.caption = db.Prepare(""& Request.Form("caption"))
	h.url = db.Prepare(""& Request.Form("url"))
	h.urlText = db.Prepare(""& Request.Form("urltext"))
	h.tout(0) = new Tout( db.Prepare(""& Utilities.GetFilenameByKey("tout1img")), db.Prepare(""& Request.Form("tout1title")), db.Prepare(""& Request.Form("tout1text")), db.Prepare(""& Request.Form("tout1url")), db.Prepare(""& Request.Form("tout1urltext")) )
	h.tout(1) = new Tout( db.Prepare(""& Utilities.GetFilenameByKey("tout2img")), db.Prepare(""& Request.Form("tout2title")), db.Prepare(""& Request.Form("tout2text")), db.Prepare(""& Request.Form("tout2url")), db.Prepare(""& Request.Form("tout2urltext")) )
	h.tout(2) = new Tout( db.Prepare(""& Utilities.GetFilenameByKey("tout3img")), db.Prepare(""& Request.Form("tout3title")), db.Prepare(""& Request.Form("tout3text")), db.Prepare(""& Request.Form("tout3url")), db.Prepare(""& Request.Form("tout3urltext")) )	
	
	' used for template purposes. these control what goes into the database
	'###############################################################################################################################
	Dim sBridgeQueries As String = ""
	Const SORT_WHERE As String = ""
	Const HAS_SORT_ORDER As Boolean = false			' set this to true to enable sort actions
	Const HAS_FILES As Boolean = true				' set this to true to enable file deletion actions
	Const DEBUG As Boolean = false					' set this to true to output the query that bombed the script. false for live
	Dim sFileFieldName() As String = { "image", "tout1img","tout2img","tout3img" } ' name of the file field name, used for file deletion
	Dim sTableName As String = "tblHomepage"				' name of the database table to perform actions on
	Dim sPrimaryKey As String = "id"				' name of the primary key field in the database
	Dim sRedirectURL As String = "HomepageDetail.aspx"		' url to redirect to once action is complete
	Dim sDetailURL As String = "HomepageDetail.aspx"		' url to redirect to detail page
	Dim sColumnNames() As String = { "metaTitle","metaKeywords","metaDescription","image","header1","header2","caption","url","urltext","tout1img","tout1title","tout1text","tout1url","tout1urltext","tout2img","tout2title","tout2text","tout2url","tout2urltext","tout3img","tout3title","tout3text","tout3url","tout3urltext" }
	Dim sValues() As String = { h.metaTitle,h.metaKeywords,h.metaDescription,h.image,h.header1,h.header2,h.caption,h.url,h.urlText,h.tout(0).img,h.tout(0).title,h.tout(0).text,h.tout(0).url,h.tout(0).urlText,h.tout(1).img,h.tout(1).title,h.tout(1).text,h.tout(1).url,h.tout(1).urlText,h.tout(2).img,h.tout(2).title,h.tout(2).text,h.tout(2).url,h.tout(2).urlText }
	'###############################################################################################################################
	
	' format the variables if there is a query string
	Dim iID As Integer = 1'CInt(Request.Form(sPrimaryKey))
	if len(Request.QueryString(sPrimaryKey)) > 0 then
		iID = 1'CInt(Request.QueryString(sPrimaryKey))
	end if
	
	if len(Request.QueryString("Action")) > 0 then
		sAction = Request.QueryString("Action")
	end if
	
	Dim _s As String
	Dim _SortOrder As String
	
	' perform the actions

%>
<!--#include file="templates/Action.aspx" -->