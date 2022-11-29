<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<title><%= ConfigurationSettings.AppSettings("applicationName") &" - "& sPageObject &" "& sPageType %></title>

	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

	<link href="includes/screen.css" rel="stylesheet" type="text/css"/>
    <link href="niceforms.css" rel="stylesheet" type="text/css"/>
    
    <script language="javascript" type="text/javascript" src="includes/niceforms.js"></script>
    <script language="javascript" type="text/javascript" src="includes/validation.js"></script>
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.4.2/jquery.min.js"></script>
    <script language="javascript">
	function DeleteListing(ID) {
		if (confirm("Are you sure you want to delete this <%= sPageObject %>?")) {
			window.location = "<%= sFormActionURL %>?Action=Delete&<%= sPrimaryKey %>=" + ID +"&t=<%= CInt(Request.QueryString("t")) %>";
		}
	}
	</script>
</head>
<body>
    <table border=0 width="900" cellspacing="0" cellpadding="0" height="100%">
    <tr><td width="200" valign="top" style="border-right: 1px solid #cccccc;">
        <center><img align="left" src="images/logo.gif"></center>
        <br clear=all>
        <div style="width:200px; height:95%; overflow: auto; padding-left: 10px;">
        <%
            RenderMenu(0)
        %>			
        </div>
        </td>
    <td width="700" valign="top" align="center">
        <!-- Start Body -->
        <br/><br/>
        <table border="0" cellpadding="0" cellspacing="0" width="670">
        	<tr><td width="670" align="left" colspan="2"><b class="title"><%= sPageObject %> - Administration</b></td></tr>
            <tr>
            	<td align="left" id="titleCell"><b class="subtitle"><%= sPageObject &" "& sPageType %></b></td>
                <td align="right">
                <% if bCanAdd then %>
                	<form method="post" class="niceform" action="<%= sFormDetailURL %>">
                    	<input type="hidden" name="Action" value="Add"/>
                        <input type="hidden" name="t" value="<%= CInt(Request.QueryString("t")) %>"/>
                        <input type="submit" name="submit" id="submit" value="Add New <%= sPageObject %>"/>
                    </form>
                <% end if %>
                </td>                        
            </tr>
        </table>
        
        <table id="displayTable" border="0" cellpadding="0" cellspacing="0" width="670" summary="Administration for <%= sPageObject %>">
        	<tr>
            	<%
				For i = 0 to UBound(sColumnHeaders)
					Response.Write("<th scope='col' ")
					if i = 0 then
						Response.Write("class='endCap'")
					end if
					Response.Write(">"& sColumnHeaders(i) &"</th>")
				next i
				%>
            </tr>
            <%= sPageData %>
        </table>        
        <!-- End Body -->
        </td></tr>
        </table>
    </td></tr>
    </table>
</body>
</html>