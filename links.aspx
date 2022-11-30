<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" Debug="true"%>
<%@ Import Namespace="Chemistry" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Collections" %>

<!--#include file="_src/classes.aspx"-->
<%
    Dim bMobileTemplate As Boolean = false
	Dim db As New DBObject(ConfigurationSettings.AppSettings("connString"))
	db.OpenConnection()
	
	Dim pid As Integer = 48
	Dim ppid As Integer = CInt(Request.QueryString("ppid"))
	Dim p As New Page
	
	if pid > 0 then
		p = Page.getById(db,pid,ppid)
	end if
	
	Dim timelineCollapsed as String
	Dim timelineExpanded as String
	
	' bMobileTemplate
	if (Request.QueryString("mobile") = "Y") Then
		bMobileTemplate = true
	End If
	
	Dim a As String = LCase(Request.ServerVariables("http_user_agent"))
	Dim sMobile() As String = { "ipod", "iphone", "android", "blackberry"}
	Dim is_iPad As Boolean
	
	for i As Integer = 0 to UBound(sMobile)
		if InStr(a,"ipad") > 0 then
			is_iPad = true
		end if 
		if InStr(a,sMobile(i)) > 0 and sMobile(i) <> "blackberry" then
			if Session("mobile") = "" then
				bMobileTemplate = true
			end if
		else if InStr(a,"blackberry") > 0 then
			if Session("mobile") = "" then
				bMobileTemplate = true
			end if
		end if
	next i
	If pid <> 4 Then
		bMobileTemplate = false
	End If
	
	if pid = 3
		Dim t As ArrayList = TimelineEntry.fillAll(db)

		Dim previousYear as Integer
		previousYear = 1000
		
		timelineCollapsed += "<div id='timeline'><table width='900' height='459'>"
		
		for each _t As TimelineEntry in t
			if _t.year <> previousYear then
				if previousYear <> 1000 then
					timelineCollapsed += "</ul></div></div></td></tr>"
				end if
				timelineCollapsed += "<tr><td valign='top' style='width:48px;background-image:url(images/ecostepstimeline/vertical.png);background-repeat:repeat-y;'><div style='height:48px;width:48px;background-image:url(images/ecostepstimeline/year_globe.png);background-repeat:no-repeat;line-height:48px;'><div style='display:block;text-align:center;'><span style='color:white;font-family: Arial, Helvetica, sans-serif;font-size:14px;font-weight:100;letter-spacing:1px;'>" & _t.year
				timelineCollapsed += "</span></div></div></td><td valign='top'><div style='padding-top: 22px; margin-left:1px;min-height:20px;'><img src='images/ecostepstimeline/horizontal.png' /><div style='width:596px; margin-top: 5px;margin-left: 40px;'><ul style='font-family: Arial, Helvetica, sans-serif;'><li><a href=""javascript:launchTimelineModal('timeline"& _t.id & "')"">"
				timelineCollapsed += _t.name
				timelineCollapsed += "</a></li>"
				previousYear = _t.year
			
			else
				timelineCollapsed += "<li><a href=""javascript:launchTimelineModal('timeline"& _t.id & "')"">"
				timelineCollapsed += _t.name
				timelineCollapsed += "</a></li>"
			
			end if	
			
			timelineExpanded += "<div id='timeline" & _t.id & "' title='" & _t.name & "' style='display:none; overflow: hidden;' class='timeline_modals'>"
			
			timelineExpanded += _t.mainText
			
			timelineExpanded += "<div style='height:10px;'></div><a style='font-size:12px;font-family: Arial, Helvetica, sans-serif;text-decoration: underline;color: #5c7db4;outline:none;' href=""javascript:CloseTimelineModal('timeline"& _t.id & "')"">Close</a></div>"
			
			
		next _t
		
		timelineCollapsed += "</ul></div></div></td></tr>"
		timelineCollapsed += "</table><img src='images/ecostepstimeline/horizontal_bottom.png' style='padding-top: 50px;margin-left:-15px;' /></div>"
		
	end if
    
    db.CloseConnection()
	
	If pid = 2 Then
		Response.Redirect("content.aspx?pid=39&ppid=2")
	End If
	
	
	If Not bMobileTemplate Then
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Eat'n Park Hospitality Group <% if len(p.pageTitle) > 0 then %>- <%= p.pageTitle %><% end if %></title>

<script src="//ajax.googleapis.com/ajax/libs/jquery/1.6.4/jquery.min.js" type="text/javascript"></script>
<script>

	//$(document).ready(function(){
	//	setTimeout(resetSize(), 2000);
	//});
	//function resetSize(){
	// 	$("#lGradient").css('height', $('#wrapper').height() + 2 +'px');
	//	$("#rGradient").css('height', $('#wrapper').height() + 2 +'px');
	//}

</script>
<script type="text/javascript" src="js/jquery-ui-1.8.12.custom.min.js"></script>
<link type="text/css" href="css/smoothness/jquery-ui-1.8.13.custom.css" rel="stylesheet" />	
<script src="js/modalPopup.js" type="text/javascript"></script>

<link href="css/style.css" rel="stylesheet" type="text/css" />

<!--
/**
 * @license
 * MyFonts Webfont Build ID 3867246, 2020-12-16T11:57:38-0500
 * 
 * The fonts listed in this notice are subject to the End User License
 * Agreement(s) entered into by the website owner. All other parties are 
 * explicitly restricted from using the Licensed Webfonts(s).
 * 
 * You may obtain a valid license at the URLs below.
 * 
 * Webfont: HelveticaLTWXX-Roman by Linotype
 * URL: https://www.myfonts.com/fonts/linotype/helvetica/pro-regular/
 * Copyright: Copyright © 2014 Monotype Imaging Inc. All rights reserved.
 * 
 * 
 * 
 * © 2020 MyFonts Incn*/

-->

<link rel="stylesheet" type="text/css" href="css/fonts.css"/>
</head>

<body>

<!--<div id="lGradient"></div>
<div id="rGradient"></div>-->
<div id="bodyContainer"></div>

	<div id="wrapper">
    
    	<div id="header">
        	<div id="header_inner">
            
                <a href="index.aspx"><img src="images/header_logo.png" alt="Eat'n Park Logo" /></a>
                <div id="main_nav">
                        
                        
                        <ul>
                            <li><a href="index.aspx">Home</a>&nbsp;&nbsp;|&nbsp;&nbsp;</li>
                            <li><a href="content.aspx?pid=1&ppid=0"<% if pid = 1 OR ppid = 1 then %> class="current_page"<% end if %>>Our Brands</a>&nbsp;&nbsp;|&nbsp;&nbsp;</li>
                            <li><a href="content.aspx?pid=2&ppid=0"<% if pid = 2 OR ppid = 2 then %> class="current_page"<% end if %>>Community Commitment</a>&nbsp;&nbsp;|&nbsp;&nbsp;</li>
                            <li><a href="content.aspx?pid=4&ppid=0"<% if pid = 4 OR ppid = 4 then %> class="current_page"<% end if %>>Careers</a>&nbsp;&nbsp;|&nbsp;&nbsp;</li>
                            <li><a href="content.aspx?pid=5&ppid=0"<% if pid = 5 OR ppid = 5 then %> class="current_page"<% end if %>>Awards</a></li>
                        </ul>
        
                </div><!--End Main Nav-->
            </div><!--End Header Inner-->
        </div> <!--End Header-->
    
    	<!--<div id="content" style="min-height:800px; width:945px;margin-left: auto;margin-right:auto;">-->
        
        <div id="content">

            
            
            <div id="middlebox_top">
            	<div id="middlebox_content">
                	<div id="middlebox_content_top">
                        
                        
                        <div id="middlebox_content_right" style="width: 800px; margin-left: 65px;">
                            
                            <%p.mainText = p.mainText.Replace("../", "")%>
                            
                            <%= p.mainText %>
                            
                            <!--<br/>-->
                            <% if pid = 1 or ppid = 1 or pid = 3 or ppid = 3 then %>
                                
                            <% else %>
                                <img style="padding-top: 20px;" src="images/content_bar_bottom.png" />
                            <% end if %>
                            
                            <!--<img src="images/lifesmiles.png" style="margin-top: 300px; margin-left:-325px;"/>-->
                            
                        </div><!--End Middlebox Content Right-->
                    </div><!--End Middlebox Content Top-->
                    
                    <div id="middlebox_content_bottom">
                    
                    	<%if ppid=1 then%>
                        	<div id="brands_images">
                        		<div id="brands_image_01">
                                	<% if len(p.tout1img)>0%>
                                    	<img src="upload/page/<%=p.tout1img%>"/>
                                    <%else%>
                            			<img src="images/imagebox/brand_image_enp_01.jpg"/>
                            		<%end if%>
                                    
                                </div>
                            
                            	<div id="brands_image_03">
                                	<% if len(p.tout2img)>0%>
                                    	<img src="upload/page/<%=p.tout2img%>"/>
                                    <%else%>
                            			<img src="images/imagebox/brand_image_enp_02.jpg"/>
                                    <%end if%>
                            	</div>
                            
                            	<div id="brands_image_03">
                                	<% if len(p.tout3img)>0%>
                                    	<img src="upload/page/<%=p.tout3img%>"/>
                                    <%else%>
                            			<img src="images/imagebox/brand_image_enp_03.jpg"/>
                                    <%end if%>
                            	</div>
                            </div>
                        	<!--<img src="images/brand_image_background.png" style="margin-left:-4px;"/>
                        	<img src="images/brand_image_background.png" style="margin-left:-2px;"/>
                            <img src="images/brand_image_background.png" style="margin-left:-2px;"/>-->
                            
                        <%else if pid=3'len(p.sidebarContent)>1 then%>
                        		<%=timelineCollapsed%>
                        		<%'p.sidebarContent = p.sidebarContent.Replace("../", "")%>
								<%'=p.sidebarContent%>

                        <%end if%>
                    
                     </div><!--End Middlebox Content Bottom-->
            	</div><!--End Middlebox Content-->
				
            </div><!--End Middlebox Top-->
            
           
            <div id="middlebox_bottom">
            
            </div><!--End Middlebox Bottom-->

         
    	</div><!--End Content-->
        
        <div id="footer">
        
        	<div id="contact">
            	<div id="contact_content">
                	
                    <div style="float: left;">
                    	<img src="images/contact_divider_1.png" alt="Divider" style="vertical-align:middle;"/><img src="images/contact_information.png" alt="Contact Information" style="margin-left: 25px; vertical-align: middle;"/>
                    </div>
                    	
                    <div style="float: left; margin-left: 75px;">
                    		<div style="float:left;"><img src="images/contact_divider_1.png" alt="Divider" style="vertical-align:middle;"/></div>
                        <!--<div style="margin-left: 20px;">-->
                        	
                        	<div style="float:left;margin-left: 30px;margin-top:5px;">
                            
                                <span class="top">ENPHG PITTSBURGH <br />HEADQUARTERS</span>
                                
                                <div style="height:15px;"></div>
                                
                                <span class="bottom" style="font-weight:bold;">412-461-2000</span>
                            
                        	</div>
                    </div>
                    
                    <div style="float: left; margin-left: 75px;">
                    		<div style="float:left;"><img src="images/contact_divider_1.png" alt="Divider" style="vertical-align:middle;"/></div>
                        <!--<div style="margin-left: 20px;">-->
                        	
                        	<div style="float:left;margin-left: 30px;margin-top:5px;">
                                
                                <span class="middle">Eat’n Park Hospitality Group<br />285 E. Waterfront Drive<br />Homestead, PA 15120</span>
                                
                                <div style="height:15px;"></div>
                                
                                <span class="bottom" style="font-weight:bold;"></span>
                            
                        	</div>
                    </div>
                
                
                </div><!--End Contact Content-->
            
                <div id="contact_bottom">
                
            	</div><!--End Contact Bottom-->
                
            </div><!--End Contact-->
            
            <div id="footer_bottom">
            	
                <div id="footer_bottom_content">
                	<div id="footer_bottom_content_left">
                    	<div class="top">Copyright &copy; 2015 Eat’n Park Hospitality Group, Inc. All rights reserved</div>
                        <div style="height:15px;"></div>
                        <div class="bottom"><a href="content.aspx?pid=41&ppid=0">Legal Notices</a>&nbsp;|&nbsp;<a href="content.aspx?pid=30&ppid=0">Privacy Policy</a></div>
                    </div>
                    
                    <div id="footer_bottom_content_right">
                    	<img style="padding-top: 3px;padding-left:10px;" src="images/footer_logo.png" alt="Eat'n Park" />
                    </div>
                
                </div><!--End Footer Bottom Content-->
            
            </div><!--End Footer Bottom-->
        	
        </div><!--End Footer-->
           
	</div> <!--End Wrapper-->
</div>
<div style="clear:both; height:10px;">&nbsp;</div>

<!--#include file="_includes/footer.aspx"-->


<%

	If pid = 4 Then   %>
    
<!--#include file="_includes/facebook-tag.aspx"-->
<%
	
	
	End If


%>


</body>

</html>

<% Else %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1"/>

<link rel="stylesheet" type="text/css" href="css/mobile.css" />
<link rel="stylesheet" media="screen" type="text/css" href="css/jquery-ui.min.css" />
<link rel="stylesheet" href="css/jquery.mobile-1.2.0.css" />
<script src="http://code.jquery.com/jquery-1.8.2.min.js"></script>


<meta name="keywords" content="" />
<meta name="description" content="" />
<title>ENP Hospitality.</title>


<style type="text/css">

	#contentBody { display:flex; flex-direction: column; }
	
	#lifesmilesimage { order: 2; margin-top: 20px; width: auto; text-align: center;}
	
	#lifesmilesimage img { margin: 0 auto; }
	
	#careersContent { order: 1; }
</style>

</head>

<body>
<center>

<div id="container">

<!--Header Begin-->
<div id="header">
    <a href="index.aspx"><img src="images/header_logo.png" width="194" height="41" alt="" style="margin: 20px 10px;"/></a>
</div>
<!--Header End-->

<!--Alert Begin-->

<div id="content">
    <div id="contentBody">

<h1><%= p.pageTitle %></h1>

                            <%
							p.mainText = p.mainText.Replace("../", "")
							p.mainText = p.mainText.Replace("<img src=""upload/LearnMoreCareers2.gif"" alt="""" width=""200"" height=""58"" />", "<br><b>Learn More About Careers At:</b><br>")
							%>
                            

                            <%= p.mainText %>
                    
    </div>
</div>

<!--Alert End-->



<!--Footer Container Begin-->
<div id="footer-container">
    <div id="footer" style="padding-right:5px" >
        <a href="index.aspx?do=full">Full Site</a>
        
      <br />
ENPHG PITTSBURGH HEADQUARTERS<br />
412-461-2000 <br />
<br />
Eat’n Park Hospitality Group<br />
285 E. Waterfront Drive<br />
Homestead, PA 15120 <br />

    </div>
    
    <div style="margin-top:10px" id="copyright">
        &copy; 2018 Eat'n Park Hospitality Group, Inc.<br/>
        All Rights Reserved.&nbsp; 
        <br /><br />
    </div>
</div>
<!--Footer Container End-->


<!-- GOOGLE TAG -->
 
 

</center>
</div>

</body>
</html>

<% End If %>

<% if pid = 3 %>
	<%=timelineExpanded%>
<% end if %>


