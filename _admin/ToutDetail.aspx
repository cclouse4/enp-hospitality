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
	Dim sPageObject As String = "Tout"					' The type of page, used for text display to identify what is going on
	Dim sFormActionURL As String = "ToutAction.aspx"		' filename of the action page used for this module
	Dim sPrimaryKey As String = "id"					' primary key of the table associated with this module
	Dim sTableName As String = "tblTout"					' database table name
	' data variables for data retrieval
	Dim iID As Integer = CInt(Request.QueryString(sPrimaryKey))
	Dim i As Integer
	Dim bActive As Boolean = true
	Dim sImageString As String
	
	Dim p as New LandingTout
	
	db.RunQuery("select * from "& sTableName &" where "& sPrimaryKey &" = "& iID)
	while db.InResults()
		p.id = db.GetItem("id")
		p.name = db.GetItem("name")
		p.text = db.GetItem("text")
		p.url = db.GetItem("url")
		p.image = db.GetItem("image")
	end while
	
	db.CloseConnection()
	
	if len(p.image) > 0 then
		sImageString = "<a href='../upload/landingTout/"& p.image &"' target='_blank'>"& p.image &"</a> <a href='ToutAction.aspx?Action=RemoveFile&Field=image&id="& iID &"'>[x]</a>"
	end if
	
	' construct the form modules	
	oForm.NewFormElement("Name","text","name",p.name,50,false,Nothing)
	oForm.NewFormElement("URL","text","url",p.url,50,false,Nothing)
	oForm.NewFormElement("Image: (180px x 120px)","file","image","",0,false,sImageString)
	oForm.NewFormElement("Text","textarea","text",p.text,50,false,Nothing)
%>
<!--#include file="templates/Detail.aspx" -->