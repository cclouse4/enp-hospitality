<%@ Page Language="VB" ContentType="text/html" debug="true" ResponseEncoding="utf-8" %>
<%@ Import Namespace="Chemistry" %>
<!--#include file="_utilities/security.aspx" -->
<!--#include file="includes/sitewide.aspx" -->
<!--#include file="includes/menu_functions.aspx" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<title><%= ConfigurationSettings.AppSettings("applicationName")	%></title>

	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

	<link href="includes/screen.css" rel="stylesheet" type="text/css"/>
</head>
<body>
	<script language="javascript">
	function deleteSection(id) {
		if (confirm("Are you sure you want to delete this section?")) {
			window.location = "SectionAction.aspx?Action=Delete&ContentID="+ id;
		}			
	}
	</script>
    <table border=0 width="900" cellspacing="0" cellpadding="0" height="100%">
    <tr><td width="200" valign="top" style="border-right: 1px solid #cccccc;">
        <center><img align="left" src="images/logo.gif"></center>
        <br clear=all>
        <div style="width:200px; height:95%; overflow: show; padding-left: 10px;">
        	<% RenderMenu(0) %>
        </div>
        </td>
    <td width="700" valign="top">
        <!-- Start Body -->
		<% if request.QueryString("HTML") = "1" then %>
        	<p style="padding-left:20px;">HTML files generated <a href="../index.html" target="_blank">click here to view</a></p>
        <% end if %>
        <!-- End Body -->
        </td></tr>
        </table>
    </td></tr>
    </table>
</body>
</html>