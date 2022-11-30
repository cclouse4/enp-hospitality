<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" %>
<%@ Import Namespace="Chemistry" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Collections" %>
<!--#include file="_src/classes.aspx"-->
<%
	Dim db As New DBObject(ConfigurationSettings.AppSettings("connString"))
	Dim p As New Homepage
	
	db.OpenConnection()


	if request.QueryString("do") = "full" then 
		Session("mobile") = "ignore"
	end if
	
	try
		db.RunQuery("select * from tblHomepage where id = 1")
		while db.InResults()
			'p.metaTitle = db.GetItem("metaTitle")
			'p.metaKeywords = db.GetItem("metaKeywords")
			'p.metaDescription = db.GetItem("metaDescription")
			p.pageTitle = ""
			p.image = db.GetItem("image")
			p.header1 = db.GetItem("header1")
			p.header2 = db.GetItem("header2")
			p.caption = db.GetItem("caption")
			'p.url = db.GetItem("url")
			p.tout(0) = new Tout(db.GetItem("tout1Img"), db.GetItem("tout1title"), db.GetItem("tout1text"), db.GetItem("tout1url"), db.GetItem("tout1urltext"))
			p.tout(1) = new Tout(db.GetItem("tout2Img"), db.GetItem("tout2title"), db.GetItem("tout2text"), db.GetItem("tout2url"), db.GetItem("tout2urltext"))
			p.tout(2) = new Tout(db.GetItem("tout3Img"), db.GetItem("tout3title"), db.GetItem("tout3text"), db.GetItem("tout3url"), db.GetItem("tout3urltext"))
		end while
	catch e As Exception
		Response.Write(e.Message)
	finally 

   
		
	end try

	db.CloseConnection()
%>



<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Eat'n Park Hospitality Group</title>

<script src="//ajax.googleapis.com/ajax/libs/jquery/1.6.4/jquery.min.js" type="text/javascript"></script>
<script>

	//$(document).ready(function(){
	//	$("#lGradient").css('height', $('body').height()+'px');
	//	$("#rGradient").css('height', $('body').height()+'px');
	//});

</script>

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
                <div id="main_nav_index">
                    
                        <ul>
                            <li><a href="index.aspx" class="current_page">Home</a>&nbsp;&nbsp;|&nbsp;&nbsp;</li>
                            <li><a href="content.aspx?pid=1&ppid=0">Our Brands</a>&nbsp;&nbsp;|&nbsp;&nbsp;</li>
                            <li><a href="content.aspx?pid=2&ppid=0">Community Commitment</a>&nbsp;&nbsp;|&nbsp;&nbsp;</li>
                            <li><a href="content.aspx?pid=4&ppid=0">Careers</a>&nbsp;&nbsp;|&nbsp;&nbsp;</li>
                            <li><a href="content.aspx?pid=5&ppid=0">Awards</a></li>
                        </ul>
        
                </div><!--End Main Nav-->
            </div><!--End Header Inner-->
        </div> <!--End Header-->
    
    	<div id="content_homepage">
            
            <div id="topbox_homepage">
            	<div id="topbox_homepage_header">
                	<!-- <%if len(p.image)>0 then%>
                		<img src="upload/page/<%=p.image%>" style="height:294px;width:934px;"/>
                    <%end if%> -->
                    <img src="upload/page/homepage-salad-revised-dec22.jpg" style="height:292px;width:936px;margin-top:3px;	border-width:1px;border-style: solid;"/>
                </div>
            
                <div id="topbox_homepage_content">
                
                    <!-- <img src="images/topbox_header.png" alt="We are a company of big dreams and humble beginnings" /> -->
										<h1 class="white-text">We are a company of big dreams and humble beginnings</h1>
						
                        <p>
							<%=p.header1%>
                        	<br/>
                            <%=p.header2%>
                        </p>
                    <!--<p>Started as a single car-hop style restaurant in Pittsburgh in 1949, Eat’n Park Hospitality Group has grown into a portfolio of regional foodservice concepts focused on personalized dining. <br/>We now serve more than xxx,xxx guests everyday in our restaurants, on college and corporate campuses, in retirement communities and hospitals, and in every state through our online store.</p>-->
                
                </div><!--End Topbox Content-->
			</div><!--End Topbox--> 
            
            
            <div id="middlebox_homepage">
                	<div id="middlebox_spacer">
                    
                    </div>
                	<div id="middlebox_homepage_left">
                    	<div id="middlebox_homepage_left_inner">
                            <div style="margin: 0 auto; max-width: 300px;">
                                <a href="content.aspx?pid=33&ppid=1"><img src="images/logos/home_logos_01.png?1=1" onMouseOver="this.src='images/logos/home_logos_rollover_01.png?1=1';" 
        onmouseout="this.src='images/logos/home_logos_01.png?1=1';" alt="Parkhurst Dining Services" /></a>
                            	<a href="content.aspx?pid=35&ppid=1"><img src="images/logos/smiley_cookie_logo_dec22.png?1=1" onMouseOver="this.src='images/logos/smiley_cookie_logo_rollover_dec22.png?1=1';" 
        onmouseout="this.src='images/logos/smiley_cookie_logo_dec22.png?1=1';" alt="Smiley Cookie: Share A Smile!" style='margin-left: 10px'/></a><br />
                            </div>
        
                            
                            <a href="content.aspx?pid=36&ppid=1"><img src="images/logos/home_logos_05.png" onMouseOver="this.src='images/logos/home_logos_rollover_05_rev.png';" 
        onmouseout="this.src='images/logos/home_logos_05.png';" alt="The Porch" style='margin-left: 5px'/></a>
        
        					<a href="http://www.enphospitality.com/content.aspx?pid=44&ppid=1"><img src="images/logos/home_logos_06.png?1=3" onMouseOver="this.src='images/logos/home_logos_rollover_06.png?1=3';" 
        onmouseout="this.src='images/logos/home_logos_06.png?1=3';" alt="Hello Bistro" style='margin-left: 25px'/></a>
        					<a href="content.aspx?pid=32&ppid=1"><img src="images/logos/eatnpark-logo-dec22.png?1=1" onMouseOver="this.src='images/logos/eatnpark-rollover-logo-dec22.png?1=1';" 
       onmouseout="this.src='images/logos/eatnpark-logo-dec22.png?1=1';" alt="Eat'n Park" style='margin-left: 25px'/></a> 
        					
                            
                       	</div>
                    </div><!--End Middlebox Left-->
                   	
                    <div id="middlebox_homepage_right">
                    
                    	<%p.caption = p.caption.Replace("../", "")%>
                        
                    	<%=p.caption%>
                        
                        <img src="images/middlebox_right_bar.gif" style="padding-top: 10px;" />
                        
                    </div><!--End Middlebox Right-->
                
			</div><!--End Middlebox-->
            
            <!-- <div id="imagebox_homepage">
            		<div style="height:2px;width:945px;">
                    </div>
					<table width="937">
                    	<tr height="197">
	                    	<td>
                            	<div id="imagebox_homepage_left">
                                	<%if len(p.tout(0).img)>0 then%>
                                    	<img src="upload/page/<%=p.tout(0).img%>" width="311" height="195" />
                                    <%else%>
                                		<img src="images/imagebox/cura.jpg" width="311" height="195"/>
                                    <%end if%>
                                    
                    				<div id="imagebox_homepage_left_inner">
                                        <p><%=p.tout(0).text%></p>
                        			</div>
                    			</div>
                            </td>
                            <td>
                            	<div id="imagebox_homepage_middle">
                                	<%if len(p.tout(1).img)>0 then%>
                                    	<img src="upload/page/<%=p.tout(1).img%>" width="311" height="195" />
                                    <%else%>
                                		<img src="images/imagebox/smiley.jpg" width="311" height="195" />
                                    <%end if%>
                                      
                    				<div id="imagebox_homepage_middle_inner">
                                    	<p><%=p.tout(1).text%></p>
                        			</div>
                    			</div>
                            </td>
                            
                            <td>
                            	<div id="imagebox_homepage_right">
                                
                                	<%if len(p.tout(2).img)>0 then%>
                                    	<img src="upload/page/<%=p.tout(2).img%>" width="311" height="195" />
                                    <%else%>
                                		<img src="images/imagebox/6penn.jpg" width="311" height="195" />
                    				<%end if%>
                                    
                                    <div id="imagebox_homepage_right_inner">
                                    	<p><%=p.tout(2).text%></p>
                        			</div>
                    			</div>
                            </td>
                        </tr>
                    </table>

            </div>End Imagebox -->
            
         
    	</div><!--End Content-->
        
        <div id="footer">
        
        	<div id="contact">
            	<!-- <div id="contact_content">
                	
                    <div style="float: left;">
                    	<img src="images/contact_divider_1.png" alt="Divider" style="vertical-align:middle;"/><img src="images/contact_information.png" alt="Contact Information" style="margin-left: 25px; vertical-align: middle;"/>
                    </div>
                    	
                    <div style="float: left; margin-left: 75px;">
                    		<div style="float:left;"><img src="images/contact_divider_1.png" alt="Divider" style="vertical-align:middle;"/></div>
                        	
                        	<div style="float:left;margin-left: 30px;margin-top:5px;">
                            
                                <span class="top">ENPHG PITTSBURGH <br />HEADQUARTERS</span>
                                
                                <div style="height:15px;"></div>
                                
                                <span class="bottom" style="font-weight:bold;">412-461-2000</span>
                            
                        	</div>
                    </div>
                    
                    <div style="float: left; margin-left: 75px;">
                    		<div style="float:left;"><img src="images/contact_divider_1.png" alt="Divider" style="vertical-align:middle;"/></div>
                        	
                        	<div style="float:left;margin-left: 30px;margin-top:5px;">
                                
                                <span class="middle">Eat’n Park Hospitality Group<br />285 E. Waterfront Drive<br />Homestead, PA 15120</span>
                                
                                <div style="height:15px;"></div>
                                
                                <span class="bottom" style="font-weight:bold;"></span>
                            
                        	</div>
                    </div>
                
                </div> -->
								
								<!--End Contact Content-->
            
                <div id="contact_bottom">
                
            	</div><!--End Contact Bottom-->
                
            </div><!--End Contact-->
            
            <div id="footer_bottom">
            	
							<div id="footer_bottom_content">
								<div id="footer_bottom_content_left">
										<div class="top">Copyright &copy; 2022 Eat’n Park Hospitality Group, Inc. All rights reserved</div>
											<div style="height:15px;"></div>
											<div class="bottom"><a href="content.aspx?pid=41&ppid=0">Legal Notices</a>&nbsp;|&nbsp;<a href="content.aspx?pid=30&ppid=0">Privacy Policy</a></div>
											<img style="padding-top: 30px;" src="images/header_logo_dec22.png" alt="Eat'n Park" />
									</div>
									
							
							</div><!--End Footer Bottom Content-->
					
					</div><!--End Footer Bottom-->
        	
        </div><!--End Footer-->
           
	</div> <!--End Wrapper-->
</div>

<!--#include file="_includes/footer.aspx"-->

</body>

</html>


