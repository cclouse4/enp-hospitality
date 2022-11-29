<%
response.buffer = true
%>
<!-- #include file="freeaspupload.inc" -->
<%
Function ExtractDirName(byVal strFilename)
    ' Removes the directory from a string that contains path and filename
    If Len(strFilename) > 1 Then
		For I = 1 to 2
			Dim X 
			For X = Len(strFilename) To 1  Step -1 
				If Mid(strFilename, X, 1) = "/" Then Exit  For 
			Next 
			If Len(strFilename) > 1 Then
				strFilename = Left(strFilename, X - 1)
			Else
				strFilename = "/"
				X = 2
			End If
		Next 
	   	ExtractDirName = Left(strFilename, X - 1)
    End If
End Function

Function ShowImageForType(strName)
' For the 'explorer': shows file-icons
	Dim strTemp
	strTemp = strName
	If strTemp <> "dir" Then
		strTemp = LCase(Right(strTemp, Len(strTemp) - InStrRev(strTemp, ".", -1, 1)))
	End If
	Select Case strTemp
		Case "asp", "aspx"
			strTemp = "asp"
		Case "dir"
			strTemp = "dir"
		Case "htm", "html"
			strTemp = "htm"
		Case "gif", "jpg"
			strTemp = "img"
		Case "txt"
			strTemp = "txt"
		Case Else
			strTemp = "misc"
	End Select
	strTemp = "<IMG SRC=""img/dir_" & strTemp & ".gif"" WIDTH=16 HEIGHT=16 BORDER=0>"
	ShowImageForType = strTemp
End Function

Dim errorMessage
function SaveFiles(PathToSaveTo)
' Saves potentially uploaded files - any errors are in errorMessage
    Dim Upload, fileName, fileSize, ks, i, fileKey
    Set Upload = New FreeASPUpload
    Upload.Save(PathToSaveTo)
    SaveFiles = ""
    ks = Upload.Errors.keys
	' if errors are returned by the component
    if (UBound(ks) <> -1) then
        SaveFiles = ""
        for each fileKey in Upload.Errors.keys
            errorMessage = errorMessage & Upload.Errors(fileKey)&" " 
        next
    end if
end function
%>

<%' Now to the Runtime code:
Dim strPath : strPath = "/upload/"   'Path of directory to show, default is root "/upload/" 

Dim strDirPath 

strDirPath = Request.QueryString("dirpath")
If strDirPath = "upload" Then
	strDirPath = "/upload/"
End If

If Len(strDirPath) > 1 Then  
	if right(strDirPath,1)="/" then
		strPath = strDirPath
	else
		strPath = strDirPath & "/"
	end if
end If
'response.write(Server.MapPath(strPath))

Dim strGetPath

If Left(strPath, 1) = "/" Then
	strPath = Right(strPath, Len(strPath) - 1)
End If

strGetPath = Request.ServerVariables("APPL_PHYSICAL_PATH") & strPath

' on submit - save files
if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	Dim uploadmessage

	uploadmessage = SaveFiles(strGetPath)
End If


If Left(strPath, 1) <> "/" Then
	strPath = "/" & strPath
End If

' Response.Write("x:" & strPath)



%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<title>{#advimage_dlg.dialog_title}</title>
	<script type="text/javascript" src="../../tiny_mce_popup.js"></script>
	<script type="text/javascript" src="../../utils/mctabs.js"></script>
	<script type="text/javascript" src="../../utils/form_utils.js"></script>
	<script type="text/javascript" src="../../utils/validate.js"></script>
	<script type="text/javascript" src="js/image.js"></script>
	<script type="text/javascript">
	<!--
	function FileChosen(FileName)
	{
	// fill the path - textbox and show a preview of the image
	 document.forms[0].elements['src'].value='../<%= strPath %>' + FileName;
	 ImageDialog.showPreviewImage('../<%= strPath %>' + FileName);
	}
	//-->
	</script>	
	<link href="css/aspimage.css" rel="stylesheet" type="text/css" />
	<base target="_self" />
</head>

<body id="advimage" style="display: none">
<form name="filebrowser" method="POST" enctype="multipart/form-data" action="image.asp?dirpath=<%= strPath %>" onSubmit="">
<% If errorMessage <>"" Then%><div class="error"><%= errorMessage %></div><% End If %>
<div class="tabs">
	<ul>
		<li id="general_tab" class="current"><span><a href="javascript:mcTabs.displayTab('general_tab','general_panel');" onMouseDown="return false;">{#advimage_dlg.tab_general}</a></span></li>
		<li id="appearance_tab"><span><a href="javascript:mcTabs.displayTab('appearance_tab','appearance_panel');" onMouseDown="return false;">{#advimage_dlg.tab_appearance}</a></span></li>
	</ul>
</div>

<div class="panel_wrapper">

	<div id="general_panel" class="panel current">
		<fieldset class="leftColumn">
			<legend>{#advimage_dlg.directory_browser}</legend>
			<div class="explorer">
				<table>
					<tr>
						<th>{#advimage_dlg.filename}</th>
						<th>{#advimage_dlg.filesize}</th>
						<!--<th>{#advimage_dlg.filetype}</th>-->
						<th>{#advimage_dlg.filemodified}</th>
					</tr>
					<tr>
						<td>&nbsp;<a href="?dirpath=<%=ExtractDirName(strPath)%>".."><img src="img/dir_pdir.gif" width="16" height="16" border="0" alt="{#advimage_dlg.parent_directory}">..</a></td>
						<td>&nbsp;</td>
						<!--<td>&nbsp;</td>-->
						<td>&nbsp;</td>
					</tr>
<% 


Dim objFSO 
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Dim objFolder 


'Response.Write("X:" & Request.ServerVariables("APPL_PHYSICAL_PATH") & "<BR>")
' strPath = Request.ServerVariables("APPL_PHYSICAL_PATH") & "upload\"

' Response.Write("Y:" & strGetPath)

' Set objFolder = objFSO.GetFolder(Server.MapPath(strPath))
Set objFolder = objFSO.GetFolder(strGetPath)
			Dim RowCount : RowCount = 0
			Dim objItem
			For Each objItem In objFolder.SubFolders
				If InStr(1, objItem, "_vti", 1) = 0 Then%>
				<tr <% If RowCount MOD 2 = 0 Then%>class="darkRow"<% End If %>>
					<td><%= ShowImageForType("dir") %>&nbsp;<a href="?dirpath=<%= strPath & objItem.Name %>"><%= objItem.Name %></a></td>
					<td><%'= objItem.Size %></td>
					<!--<td><%'= objItem.Type %></td>-->
					<td><%= objItem.DateCreated %></td>
				</tr>
<%				RowCount = RowCount + 1
				End If
			Next %>
<%			For Each objItem In objFolder.Files	%>
				<tr <% If RowCount MOD 2 = 0 Then%>class="darkRow"<% End If %>>
					<td nowrap><%= ShowImageForType(objItem.Name) %>&nbsp;<a href="Javascript:FileChosen('<%= objItem.Name %>')"><%= objItem.Name %></a></td>
					<td><%= objItem.Size %></td>
					<!--<td><%= objItem.Type %></td>-->
					<td nowrap><%= objItem.DateCreated %></td>
				</tr>
<%				RowCount = RowCount + 1
			Next 
			Set objItem = Nothing
Set objFolder = Nothing
Set objFSO = Nothing
%>
				</table>
			</div><!-- end explorer -->
		<table width="100%" border="0">
			<tr>
				<td width="80"><label for="upload">{#advimage_dlg.upload}</label></td>
				<td width="274"><input type="file" name="upload" id="upload" size="32"><input type="hidden" name="chosendir" value="<%= strPath %>"></td>
				<td width="*"><input id="insert" type="submit" name="Submit" value="Upload"></td>
			</tr>
		</table>			
		</fieldset><!-- end leftcolumn -->
		
		<fieldset class="rightColumn">
			<legend>{#advimage_dlg.preview}</legend>
			<div id="prev"></div>
			<input name="src" type="text" id="src" value="" onChange="ImageDialog.showPreviewImage(this.value);" />							
		</fieldset>

	</div><!-- end general panel -->

	<div id="appearance_panel" class="panel">
		<fieldset>
			<legend>{#advimage_dlg.tab_appearance}</legend>

			<table border="0" cellpadding="4" cellspacing="0">
				<tr> 
					<td class="column1"><label id="altlabel" for="alt">{#advimage_dlg.alt}</label></td> 
					<td><input id="alt" name="alt" type="text" value="" /></td> 
				</tr> 
				<tr> 
					<td class="column1"><label id="titlelabel" for="title">{#advimage_dlg.title}</label></td> 
					<td><input id="title" name="title" type="text" value="" /></td> 
				</tr>
				<tr> 
					<td class="column1"><label id="alignlabel" for="align">{#advimage_dlg.align}</label></td> 
					<td><select id="align" name="align" onChange="ImageDialog.updateStyle();ImageDialog.changeAppearance();"> 
							<option value="">{#not_set}</option> 
							<option value="baseline">{#advimage_dlg.align_baseline}</option>
							<option value="top">{#advimage_dlg.align_top}</option>
							<option value="middle">{#advimage_dlg.align_middle}</option>
							<option value="bottom">{#advimage_dlg.align_bottom}</option>
							<option value="text-top">{#advimage_dlg.align_texttop}</option>
							<option value="text-bottom">{#advimage_dlg.align_textbottom}</option>
							<option value="left">{#advimage_dlg.align_left}</option>
							<option value="right">{#advimage_dlg.align_right}</option>
						</select> 
					</td>
					<td rowspan="6" valign="top">
						<div class="alignPreview">
							<img id="alignSampleImg" src="img/sample.gif" alt="{#advimage_dlg.example_img}" />
							Lorem ipsum, Dolor sit amet, consectetuer adipiscing loreum ipsum edipiscing elit, sed diam
							nonummy nibh euismod tincidunt ut laoreet dolore magna aliquam erat volutpat.Loreum ipsum
							edipiscing elit, sed diam nonummy nibh euismod tincidunt ut laoreet dolore magna aliquam
							erat volutpat.
						</div>
					</td>
				</tr>

				<tr>
					<td class="column1"><label id="widthlabel" for="width">{#advimage_dlg.dimensions}</label></td>
					<td nowrap="nowrap">
						<input name="width" type="text" id="width" value="" size="5" maxlength="5" class="size" onChange="ImageDialog.changeHeight();" /> x 
						<input name="height" type="text" id="height" value="" size="5" maxlength="5" class="size" onChange="ImageDialog.changeWidth();" /> px
					</td>
				</tr>

				<tr>
					<td>&nbsp;</td>
					<td><table border="0" cellpadding="0" cellspacing="0">
							<tr>
								<td><input id="constrain" type="checkbox" name="constrain" class="checkbox" /></td>
								<td><label id="constrainlabel" for="constrain">{#advimage_dlg.constrain_proportions}</label></td>
							</tr>
						</table></td>
				</tr>

				<tr>
					<td class="column1"><label id="vspacelabel" for="vspace">{#advimage_dlg.vspace}</label></td> 
					<td><input name="vspace" type="text" id="vspace" value="" size="3" maxlength="3" class="number" onChange="ImageDialog.updateStyle();ImageDialog.changeAppearance();" />
					</td>
				</tr>

				<tr> 
					<td class="column1"><label id="hspacelabel" for="hspace">{#advimage_dlg.hspace}</label></td> 
					<td><input name="hspace" type="text" id="hspace" value="" size="3" maxlength="3" class="number" onChange="ImageDialog.updateStyle();ImageDialog.changeAppearance();" /></td> 
				</tr>

				<tr>
					<td class="column1"><label id="borderlabel" for="border">{#advimage_dlg.border}</label></td> 
					<td><input id="border" name="border" type="text" value="" size="3" maxlength="3" class="number" onChange="ImageDialog.updateStyle();ImageDialog.changeAppearance();" /></td> 
				</tr>

				<tr>
					<td><label for="class_list">{#class_name}</label></td>
					<td><select id="class_list" name="class_list"></select></td>
				</tr>

				<tr>
					<td class="column1"><label id="stylelabel" for="style">{#advimage_dlg.style}</label></td> 
					<td colspan="2"><input id="style" name="style" type="text" value="" onChange="ImageDialog.changeAppearance();" /></td> 
				</tr>

				<!-- <tr>
					<td class="column1"><label id="classeslabel" for="classes">{#advimage_dlg.classes}</label></td> 
					<td colspan="2"><input id="classes" name="classes" type="text" value="" onchange="selectByValue(this.form,'classlist',this.value,true);" /></td> 
				</tr> -->
			</table>
		</fieldset>
	</div><!-- end appearance panel -->
	
</div> <!-- end panel_wrapper -->

<div class="mceActionPanel">
	<div style="float: left">
		<input type="button" id="insert" name="insert" value="{#insert}" onClick="ImageDialog.insert();" />
	</div>
	<div style="float: right">
		<input type="button" id="cancel" name="cancel" value="{#cancel}" onClick="tinyMCEPopup.close();" />
	</div>
</div>
</form>
</body> 
</html> 
