<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" %>
<%@ Import Namespace="Chemistry" %>
<%
' variables
Dim sUser As String
Dim sPass As String
Dim sError As String = Request.QueryString("error")
Dim sFormAction As String = Request.Form("action")

' redirect the user if they are logged in
If Session("LoggedIn") = 1 then
	Response.Redirect("splash.aspx")
end if

' get the session data
If len(sError) = 0 then
	sUser = Session("Username")
	sPass = Session("Password")
end if

' process the login
If sFormAction = "login" then
	Dim db As New DBObject(ConfigurationSettings.AppSettings("connString"))
	
	sUser = db.Prepare(Request.Form("Login"),true)
	sPass = db.Prepare(Request.Form("Password"),true)
	Session("AdminNavDisplay") = 0
	
	Try
		db.OpenConnection()
		db.RunQuery("select * from tblUser where Username = "& sUser &" and Password = "& sPass)
		while db.InResults()
			Session("Username") = db.GetItem("Username")
			Session("Password") = db.GetItem("Password")
			Session("UserID") = db.GetItem("UserID")
            Session("UserType") = db.GetItem("UserType")
			Session("LoggedIn") = 1
			
		end while	
	Catch e as Exception
		sError = "<span style='color:#ff0000;'>Error: "& e.Message &"</span>"
		Response.Write(sError)
	Finally
		db.CloseConnection()
		
	End Try
	
	' redirect the user based on their login attempt success
	if Session("LoggedIn") = 1 then
		Response.Redirect("splash.aspx")
	Else
		'Response.Redirect("index.aspx?error=1")
	end if
end if

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
		
        <title><% Response.Write(ConfigurationSettings.AppSettings("applicationName")) %></title>
        
        <link href="includes/screen.css" rel="stylesheet" type="text/css"/>
        <link href="niceforms.css" rel="stylesheet" type="text/css"/>
        
		<script language="javascript" type="text/javascript" src="includes/niceforms.js"></script>
        <script language="javascript" type="text/javascript" src="includes/validation.js"></script>	
	</head>
	<body>
   		<table width="100%" height="300" cellspacing="0" cellpadding="0" border="0">
	        <tr align="center" valign="middle">
            	<td align="center" valign="middle">
                    <form method="post" action="index.aspx" name="login" id="login" class="niceform">
                    	<input type="hidden" name="action" value="login">
                        <table width="100%" height="100%" cellspacing="0" cellpadding="0" border="0">
	                        <tr align="center" valign="middle">
                            	<td align="center" valign="middle">
                            		<center>
                            			<table width="600" cellspacing="4" cellpadding="2" border="0">
                            				<tr><td colspan="3" align="center"><img src="images/logo.gif" border="0"></td></tr>
                            				<tr><td colspan="3" align="center">
                                				<table border="0">
                                					<tr><td colspan=3 align="center">
														<% If sError = "1" Then %>
                                                        	<font color="red">Incorrect login/password</font>
                                						<% End If %>
                                					</td></tr>
                                					<tr><td align="right">Login:&nbsp;&nbsp;&nbsp;</td>						
                                                        <td align="left"><input type="text" name="Login" tabindex="1" value="<%= sUser %>"></td>		
                                                        <td align="left" rowspan="3">&nbsp;&nbsp;&nbsp;&nbsp;
                                    						<input type="submit" name="submit" id="submit" value="Login" />
                               						<tr><td colspan="2" height="3" align="right"></td></tr> 
                                					<tr><td align="right">Password:&nbsp;&nbsp;&nbsp;</td>						
                                    					<td align="left"><input type="password" name="Password"  value="<%= sPass %>" tabindex="2"></td></tr>
                                				</table>
                            				</td></tr>
                            			</table>
                        			</center>
                    			</td>
                            </tr>
                    	</table>
                    </form>
                </td>
            </tr>
        </table>
	</body>
</html>