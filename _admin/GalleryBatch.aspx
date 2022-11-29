<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="iso-8859-1" Debug="True" %>
<%@ Import Namespace="Chemistry" %>
<!--#include file="includes/sitewide.aspx" -->
<!--#include file="includes/image.aspx" -->

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
	
    Dim sSQL As String
    Dim sImage As String
    Dim sThumb As String
    
    Dim arrThumbs
    
    Dim iCounter As Integer = 0
    
	db.OpenConnection()
	
 	sSQL = "SELECT * FROM tblGallery"
    
	db.RunQuery(sSQL)
    while db.InResults()
    	iCounter += 1
    end while
    
    Redim arrThumbs(2, iCounter)
    
    iCounter = 0
	sSQL = "SELECT * FROM tblGallery"
    
	db.RunQuery(sSQL)
    
    while db.InResults()
    
    	sImage = db.GetItem("Image")
        sThumb = "TH_" & sImage
       
        
        ' Delete Thumb If Visible
		if (System.IO.File.Exists(Server.MapPath("../") & ConfigurationSettings.AppSettings("uploadPath") &"\gallery\thumbs\" & sThumb) ) then
			System.IO.File.Delete(Server.MapPath("../") & ConfigurationSettings.AppSettings("uploadPath") &"\gallery\thumbs\" & sThumb)
		end if
        
		if (System.IO.File.Exists(Server.MapPath("../") & ConfigurationSettings.AppSettings("uploadPath") &"\gallery\mobile\" & sThumb) ) then
			System.IO.File.Delete(Server.MapPath("../") & ConfigurationSettings.AppSettings("uploadPath") &"\gallery\mobile\" & sThumb)
		end if
        
        Call MakeThumb(ConfigurationSettings.AppSettings("uploadPath") &"\gallery\", ConfigurationSettings.AppSettings("uploadPath") &"\gallery\thumbs", sImage, sThumb, 250, 0, 81, 1)
		
        
        Call MakeThumb(ConfigurationSettings.AppSettings("uploadPath") &"\gallery\", ConfigurationSettings.AppSettings("uploadPath") &"\gallery\mobile", sImage, sThumb, 320, 0, 61, 1)
        
		arrThumbs(0, iCounter) = db.GetItem("id")
        arrThumbs(1, iCounter) = sThumb
        
    	iCounter += 1
        
        ' If iCounter > 0 Then Exit While
        
    end while
 
    For i = 0 To UBound(arrThumbs, 2)
    
    	If arrThumbs(1, i) <> "" And IsNumeric(arrThumbs(0, i)) Then
        	'Response.Write("X:" &  arrThumbs(1, i) & "<BR>")
            'Response.Write("X:" &  arrThumbs(0, i) & "<BR>")
            
            sSQL  = "UPDATE tblGallery SET Thumbnail = '" & arrThumbs(1, i) & "' WHERE id = " & arrThumbs(0, i)
            db.RunOther(sSQL)
            
            Response.Write(sSQL & "<BR>")
            
        End If
    
    Next
    
    
    db.CloseConnection()

	Response.Write("TEST")
    
%>

