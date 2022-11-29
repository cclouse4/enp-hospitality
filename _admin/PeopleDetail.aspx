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
	Dim sPageObject As String = "People"					' The type of page, used for text display to identify what is going on
	Dim sFormActionURL As String = "PeopleAction.aspx"		' filename of the action page used for this module
	Dim sPrimaryKey As String = "id"					' primary key of the table associated with this module
	Dim sTableName As String = "tblPeople"					' database table name
	' data variables for data retrieval
	Dim iID As Integer = CInt(Request.QueryString(sPrimaryKey))
	Dim i As Integer
	Dim bActive As Boolean = true
	Dim sel() As String = { "Chefs","Staff","Managers" }
	
	Dim p As New People
	Dim _is As String
	
	db.RunQuery("select * from "& sTableName &" where "& sPrimaryKey &" = "& iID)
	while db.InResults()
		p.id = db.GetItem("id")
		p.image = db.GetItem("image")
		p.setName(db.GetItem("fname"),db.GetItem("lname"))
		p.title = db.GetItem("title")
		p.contentTitle = db.GetItem("contentTitle")
		p.contentText = db.GetItem("content")
		p.email = db.GetItem("email")
		p.phone = db.GetItem("phone")
		p.type = db.GetItem("type")
	end while	
	
	if len(p.image) > 0 then
		_is = "<a href='../upload/people/"& p.image &"' target='_blank'>"& p.image &"</a> <a href='PeopleAction.aspx?Action=RemoveFile&Field=image&id="& iID &"'>[x]</a>"
	end if
	
	db.CloseConnection()
		
	' construct the form modules
	oForm.NewFormElement("First Name","text","fname",p.name(0),50,false,Nothing)
	oForm.NewFormElement("Last Name","text","lname",p.name(1),50,false,Nothing)
	oForm.NewFormElement("Image: (271px wide, variable height)","file","image",p.image,50,false,_is)
	oForm.NewFormElement("Type","select","type",sel,p.type,false,Nothing)
	oForm.NewFormElement("Title","text","title",p.title,50,false,Nothing)
	oForm.NewFormElement("Email","text","email",p.email,50,false,Nothing)
	oForm.NewFormElement("Phone","text","phone",p.phone,50,false,Nothing)
	
	oForm.NewFormElement("Content Title","text","contentTitle",p.contentTitle,50,false,Nothing)
	oForm.NewFormElement("Content","textarea","content",p.contentText,50,false,Nothing)
	
	
%>
<!--#include file="templates/Detail.aspx" -->