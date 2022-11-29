<%
if len(Request.Form("Action")) > 0 then
	sAction = Request.Form("Action")
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<title><%= ConfigurationSettings.AppSettings("applicationName") &" - "& sPageObject &" "& sPageType %></title>

	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

	<link href="includes/screen.css" rel="stylesheet" type="text/css"/>
    <link href="niceforms.css" rel="stylesheet" type="text/css"/>
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.4.2/jquery.min.js"></script>
    <script language="javascript" type="text/javascript" src="includes/validation.js"></script>
    <script type="text/javascript" src="includes/scw.js"></script>
    <script type="text/javascript" src="includes/multiselectpanel.js"></script>
    <% if oForm.useEditor %>
    	<script type="text/javascript" src="editor/tiny_mce.js"></script>
        <script type="text/javascript">
		tinyMCE.init({
			mode : "textareas",
			theme : "advanced",
			plugins : "safari,spellchecker,pagebreak,style,layer,table,save,advhr,aspimage,advlink,emotions,iespell,inlinepopups,insertdatetime,preview,media,searchreplace,print,contextmenu,paste,directionality,fullscreen,noneditable,visualchars,nonbreaking,xhtmlxtras,template",
			// Theme options
			theme_advanced_buttons1 : "link,unlink,bold,italic,underline,strikethrough,bullist,numlist|,justifyleft,justifycenter,justifyright,justifyfull,fontselect,formatselect,fontsizeselect,sub,sup,tablecontrols,image,code",
			theme_advanced_buttons2 : "",
			theme_advanced_buttons3 : "",
			theme_advanced_buttons4 : "",
			theme_advanced_toolbar_location : "top",
			theme_advanced_toolbar_align : "left",
			theme_advanced_statusbar_location : "bottom",
			theme_advanced_resizing : true
			
			//content_css : "../styles/editor.css"

		});
		function formatSelectBox(boxName) {
			var i;
			var box = document.getElementById(boxName);
			if (box != null) {
				for (i = 0; i < box.options.length; i++) {
					box.options[i].selected = "selected";
				}
			}
			
			$('#title').val($('#title').val().replace("è","&egrave;"));
			$('#title').val($('#title').val().replace("é","&eacute;"));
			$('#title').val($('#title').val().replace("œ","&#156;"));
			$('#title').val($('#title').val().replace("ï","&#239;"));
			$('#title').val($('#title').val().replace("û","&#251;"));
			
			
			
			return true;
		}
		function deleteSection(id) {
			if (confirm("Are you sure you want to delete this section?")) {
				window.location = "SectionAction.aspx?Action=Delete&ContentID="+ id;
			}			
		}
		</script>
        <script language="javascript" type="text/javascript" src="includes/niceforms.js"></script>
    <% end if %>
</head>
<body>
    <table border=0 width="1200" cellspacing="0" cellpadding="0" height="100%">
    <tr><td width="200" valign="top" style="border-right: 1px solid #cccccc;">
        <center><img align="left" src="images/logo.gif"></center>
        <br clear=all>
        <div style="width:200px; height:95%; overflow: auto; padding-left: 10px;">
        <%
            RenderMenu(0)
        %>
		<!--#include file="../includes/menu.aspx" -->		
        </div>
        </td>
    <td width="1000" valign="top" align="center">
        <!-- Start Body -->
        <br/><br/>
        <div style="margin-left: 10px;">
            <table border="0" cellpadding="0" cellspacing="0" width="1000">
                <tr><td width="1000" align="left"><b class="title"><%= sPageObject &" "& sPageType %> - Administration</b></td></tr>
                <tr><td class="subtitle" align="left">
                <% select case sAction
                    case "Add"
                        Response.Write("Add New "& sPageObject)
                    case "Edit"
                        Response.Write("Modify "& sPageObject)
                    case "Delete"
                        Response.Write("Remove "& sPageObject)
                end select %>
                
            </table>
            <div id="container">
            <script type="text/javascript" language="javascript">
                //function validateOnSubmit() {
                    //var elem;
                    //var errs = 0;
                    //if (!validatePresent(document.forms.form.IndustryName, 'inf_from')) errs++;
                    //if (errs > 1)  alert('There are fields which need correction before sending');
                    //if (errs == 1) alert('There is a field which needs correction before sending');
                    //if (errs == 0) ? return true : return false;
                //};
                function deleteSubmit() { document.form.submit(); };
				$(document).ready(function() {
					$('#regionSelect').change(function() {
						var newID = $('#regionSelect option:selected').val();
						window.location = "ManageAreas.aspx?ContactID=<%= iID %>&RegionID=" + newID;
					});
				});
            </script>
            
            <form accept-charset="ISO-8859-15" action="<%= sFormActionURL %>" method="post" aclass="niceform" name="form" onSubmit="return formatSelectBox('PracticeContact_ValueSelected');" enctype="multipart/form-data">
                <input type="hidden" name="Action" value="<%= sAction %>"/>
                <input type="hidden" name="<%= sPrimaryKey %>" value="<%= iID %>"/>
                
                <table width="1000" cellpadding="0" cellspacing="0" border="0" id="displayTable">
                    <%= oForm.PrintFormItems() %>
                    <%= oForm.PrintSubmitButton("Save Changes") %>
                    
                    <tr>
                        <th class="topCap" colspan="2">&nbsp;</th> 
                    </tr>   
                </table>
            </form>     
            <!-- End Body -->
            </td></tr>
            </table>
        </div>
    </td></tr>
    </table>
</body>
</html>
