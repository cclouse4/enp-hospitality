<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="iso-8859-1" Debug="true" %>
<%@ Import Namespace="Chemistry" %>
<!--#include file="includes/menu_functions.aspx" -->
<!--#include file="_utilities/formbuilder_class.aspx" -->
<!--#include file="includes/sitewide.aspx" -->
<!--#include file="../_src/classes.aspx"-->
<%
	' page variables
	Dim sAction As String = Request.QueryString("Action")
	Dim oForm As New FormBuilder()
	Dim db As New DBObject(ConfigurationSettings.AppSettings("connString"))
	db.OpenConnection()
	' template variables -- These control the page fucntionality and display
	Dim sPageType As String = "Detail"					' Use Detail, List, Delete
	Dim sPageObject As String = "Homepage"					' The type of page, used for text display to identify what is going on
	Dim sFormActionURL As String = "HomepageAction.aspx"		' filename of the action page used for this module
	Dim sPrimaryKey As String = "id"					' primary key of the table associated with this module
	Dim sTableName As String = "tblHomepage"					' database table name
	' data variables for data retrieval
	Dim iID As Integer = CInt(Request.QueryString(sPrimaryKey))
	Dim i As Integer
	Dim bActive As Boolean = true
	
	Dim h As New Homepage
	Dim sImgString(3) As String
	
	db.RunQuery("select * from "& sTableName &" where "& sPrimaryKey &" = 1")
	while db.InResults()
		h.metaKeywords = db.GetItem("metaKeywords")
		h.metaDescription = db.GetItem("metaDescription")
		h.image = db.GetItem("image")
		h.header1 = db.GetItem("header1")
		h.header2 = db.GetItem("header2")
		h.metaTitle = db.GetItem("metaTitle")
		h.caption = db.GetItem("caption")
		h.url = db.GetItem("url")
		h.urlText = db.GetItem("urltext")
		h.tout(0) = new Tout(db.GetItem("tout1img"), db.GetItem("tout1title"), db.GetItem("tout1text"), db.GetItem("tout1url"), db.GetItem("tout1urltext"))
		h.tout(1) = new Tout(db.GetItem("tout2img"), db.GetItem("tout2title"), db.GetItem("tout2text"), db.GetItem("tout2url"), db.GetItem("tout2urltext"))
		h.tout(2) = new Tout(db.GetItem("tout3img"), db.GetItem("tout3title"), db.GetItem("tout3text"), db.GetItem("tout3url"), db.GetItem("tout3urltext"))
	end while	
	db.CloseConnection()
	
	for i = 0 to 2
		if len(h.tout(i).img) > 0 then
			sImgString(i) = "<a href='../upload/touts/"& h.tout(i).img &"' target='_blank'>"& h.tout(i).img &"</a> <a href='"& sFormActionURL &"?Action=RemoveFile&Field=tout"& i+1 &"img&"& sPrimaryKey &"="& iID &"'>[x]</a>"
		end if	
	next i	
	
	if len(h.image) > 0 then
		sImgString(3) = "<a href='../upload/touts/"& h.image &"' target='_blank'>"& h.image &"</a> <a href='"& sFormActionURL &"?Action=RemoveFile&Field=image&"& sPrimaryKey &"="& iID &"'>[x]</a>"
	end if
	
	' construct the form modules
	'oForm.NewFormElement("Meta Title","text","metaTitle",h.metaTitle,50,false,Nothing)
	'oForm.NewFormElement("Meta Keywords","text","metaKeywords",h.metaKeywords,50,false,Nothing)
	'oForm.NewFormElement("Meta Description","text","metaDescription",h.metaDescription,50,false,Nothing)
	oForm.NewFormElement("Header Image (934px x 293px)","cell",Nothing,0,false,Nothing)
	oForm.NewFormElement("Image","file","image","",0,false,sImgString(3))
	oForm.NewFormElement("Header Line 1","text","header1",h.header1,175,false,Nothing)
	oForm.NewFormElement("Header Line 2","text","header2",h.header2,175,false,Nothing)
	oForm.NewFormElement("Middle Text","textarea","caption",h.caption,50,false,Nothing)
	'oForm.NewFormElement("URL","text","url",h.url,50,false,Nothing)
	'oForm.NewFormElement("URL Text","text","urltext",h.urltext,50,false,Nothing)
	
	for _i As Integer = 0 to 2
		oForm.NewFormElement("Tout "& _i+1 &" (311px width x 195px height)","cell","","",0,false,Nothing)
		oForm.NewFormElement("Image","file","tout"& _i+1 &"img",h.tout(_i).img,50,false,sImgString(_i))
		'oForm.NewFormElement("Title","text","tout"& _i+1 &"title",h.tout(_i).title,50,false,Nothing)
		oForm.NewFormElement("Caption","text","tout"& _i+1 &"text",h.tout(_i).text,50,false,Nothing)
		'oForm.NewFormElement("URL","text","tout"& _i+1 &"url",h.tout(_i).url,50,false,Nothing)
		'oForm.NewFormElement("URL Text","text","tout"& _i+1 &"urltext",h.tout(_i).urlText,50,false,Nothing)
	next _i		
%>
<!--#include file="templates/Detail.aspx" -->