<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="iso-8859-1" %>
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
	Dim sPageObject As String = "Content"					' The type of page, used for text display to identify what is going on
	Dim sFormActionURL As String = "TimelineAction.aspx"		' filename of the action page used for this module
	Dim sPrimaryKey As String = "id"					' primary key of the table associated with this module
	Dim sTableName As String = "tblTimeline"					' database table name
	' data variables for data retrieval
	Dim iID As Integer = CInt(Request.QueryString(sPrimaryKey))
	Dim i As Integer
	Dim bActive As Boolean = true
	Dim sImageString(5) As String
	
	Dim p as New TimelineEntry
	
	db.RunQuery("select * from "& sTableName &" where "& sPrimaryKey &" = "& iID)
	while db.InResults()
		p.id = db.GetItem("id")
		p.year = db.GetItem("year")
		p.name = db.GetItem("name")
		p.mainText = db.GetItem("mainText")
	end while
	
	'Dim sParentSectionList As String = "<option value='-1'>- Select One -</option><option value='0'>Not in Navigation</option>"
	'db.RunQuery("select * from tblContent where parentId = -1 order by id asc")
	'while db.InResults()
	'	sParentSectionList += "<option value='"& db.GetItem("id") &"'"
	'	if db.GetItem("id") = p.parentId then
	'		sParentSectionList += " selected=selected"
	'	end if
	'	sParentSectionList += ">"& db.GetItem("name") &"</option>"
	'end while
	
	'Dim selIds As String
	'db.RunQuery("select * from brgContentTout where contentId = "& p.id)
	'while db.InResults()
	'	selIds += db.GetItem("toutId") &"|"
	'end while
	' get the touts
	
	'Dim sTouts As String
	'db.RunQuery("select * from tblTout order by name asc")
	'while db.InResults()
	'	sTouts += "<input type='checkbox' name='touts[]' value='"& db.GetItem("id") &"'"
	'	for each id As String in split(selIds,"|")
	'		if id = db.getItem("id") then
	'			sTouts += " CHECKED"
	'			exit for
	'		end if
	'	next id
	'	sTouts += "> "& db.GetItem("name") &" <br/>"
	'end while
	
	db.CloseConnection()
	
	'if len(p.backgroundImage) > 0 then
	'	sImageString(0) = "<a href='../upload/page/"& p.backgroundImage &"' target='_blank'>"& p.backgroundImage &"</a> <a href='ContentAction.aspx?Action=RemoveFile&Field=backgroundImage&id="& iID &"'>[x]</a>"
	'end if
	
	'if len(p.image) > 0 then
	'	sImageString(1) = "<a href='../upload/page/"& p.image &"' target='_blank'>"& p.image &"</a> <a href='ContentAction.aspx?Action=RemoveFile&Field=image&id="& iID &"'>[x]</a>"
	'end if
	
	'if len(p.tout1img) > 0 then
	'	sImageString(2) = "<a href='../upload/page/"& p.tout1img &"' target='_blank'>"& p.tout1img &"</a> <a href='ContentAction.aspx?Action=RemoveFile&Field=tout1img&id="& iID &"'>[x]</a>"
	'end if
	
	'if len(p.tout2img) > 0 then
	'	sImageString(3) = "<a href='../upload/page/"& p.tout2img &"' target='_blank'>"& p.tout2img &"</a> <a href='ContentAction.aspx?Action=RemoveFile&Field=tout2img&id="& iID &"'>[x]</a>"
	'end if
	
	'if len(p.tout3img) > 0 then
		'sImageString(4) = "<a href='../upload/page/"& p.tout3img &"' target='_blank'>"& p.tout3img &"</a> <a href='ContentAction.aspx?Action=RemoveFile&Field=tout3img&id="& iID &"'>[x]</a>"
	'end if
	
	' construct the form modules
	'oForm.NewFormElement("Parent Section","select","parentId",sParentSectionList,0,false,Nothing)
	oForm.NewFormElement("Year","text","year",p.year,4,false,Nothing)
	oForm.NewFormElement("Name","text","name",p.name,150,false,Nothing)
	'oForm.NewFormElement("Page Title (shown on site)","text","pageTitle",p.pageTitle,50,false,Nothing)
	'oForm.NewFormElement("URL of page (static pages only)","text","url",p.url,50,false,Nothing)
	'oForm.NewFormElement("Meta Title","text","metaTitle",p.metaTitle,100,false,Nothing)
	'oForm.NewFormElement("Meta Keywords","text","metaKeywords",p.metaKeywords,100,false,Nothing)
	'oForm.NewFormElement("Meta Description","text","metaDescription",p.metaDescription,100,false,Nothing)
	
	'oForm.NewFormElement("Image Management","cell",Nothing,0,false,Nothing)
	'oForm.NewFormElement("Background Image: (2000px x 608px)","file","backgroundImage","",0,false,sImageString(0))
	'oForm.NewFormElement("Header Image: (934px x 159px)","file","image","",0,false,sImageString(1))
	
	oForm.NewFormElement("Modal Popup Content","cell",Nothing,0,false,Nothing)
	'oForm.NewFormElement("Intro Text","text","introText",p.introText,100,false,Nothing)
	oForm.NewFormElement("Popup Text","textarea","mainText",p.mainText,50,false,Nothing)
	
	'oForm.NewFormElement("Secondary Content","cell",Nothing,0,false,Nothing)
	'oForm.NewFormElement("Title","text","sidebarTitle",p.sidebarTitle,100,false,Nothing)
	'oForm.NewFormElement("Text","textarea","sidebarContent",p.sidebarContent,50,false,Nothing)
	
	
	'oForm.NewFormElement("Images for Our Brands subsections (298px width by 188px height)","cell",Nothing,0,false,Nothing)
	'oForm.NewFormElement("Image 1:","file","tout1img",p.tout1img,0,false,sImageString(2))
	'oForm.NewFormElement("Image 2:","file","tout2img",p.tout2img,0,false,sImageString(3))
	'oForm.NewFormElement("Image 3:","file","tout3img",p.tout3img,0,false,sImageString(4))
	
	'oForm.NewFormElement("Use Landing Template","checkbox","landingPage",1,0,p.isLandingPage,Nothing)
	'oForm.NewFormElement("Landing Touts<br/><a href='ToutList.aspx'>Manage</a>","multiCheckbox",Nothing,sTouts,false,Nothing)
	
	'oForm.NewFormElement("Content Header","text","header",sHeader,50,false,Nothing)
	'oForm.NewFormElement("Content Blurb","textarea","content",sContent,50,false,Nothing)
	'oForm.NewFormElement("Page Body","textarea","body",sBody,50,false,Nothing)
	'oForm.NewFormElement("Footer Text","textarea","footertext",sFooterText,50,false,Nothing)
	
	'oForm.NewFormElement("Header Image","file","img","",0,false,sImageString)
	
	
	'oForm.NewFormElement("Active","checkbox","active",1,0,bActive,Nothing)
%>
<!--#include file="templates/Detail.aspx" -->