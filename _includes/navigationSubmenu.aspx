<%

db.OpenConnection()

Dim aSubMenu As ArrayList = Page.getSubMenu(db)
Dim idOfParent As Integer
Dim intCount As Integer

intCount = 0

db.CloseConnection()

Dim boolMenu As Boolean
boolMenu = false

if ppid = 1 OR pid = 1 then
    boolMenu = true
end if
if ppid = 2 OR pid = 2 then
    boolMenu = true
end if
if ppid = 3 OR pid = 3 then
    boolMenu = true
end if
if ppid = 4 OR pid = 4 then
    boolMenu = true
end if
if ppid = 5 OR pid = 5 then
boolMenu = true
end if


%>

		<% if pid = 30 OR pid = 41 then %>
        
        <% else %>

            <table border="0" style="border-collapse: collapse;">

            <% if boolMenu then %>
            <tr><td><img src="images/subnav_bar_top.png" /></td></tr>
            <% end if %>
            <% if ppid = 1 OR pid = 1 then %>
            
                <tr><td width="20"><img src="images/subnav_bar_item.png" /></td><td colspan="2"><a href="content.aspx?pid=1&ppid=0"><img src="images/subnav_ourbrands.png" alt="Our Brands" style="padding-left: 5px;"/></a></td></tr>
                
                <% idOfParent = 1 %>
                
            <% end if %>
            
            <% if ppid = 2 OR pid = 2 then %>
            
                <tr><td width="20"><img src="images/subnav_bar_item.png" /></td><td colspan="2"><a href="content.aspx?pid=37&ppid=2"><img src="images/subnav_community.png" alt="Community Commitment" style="padding-left: 5px;"/></a></td></tr>
                
                <% idOfParent = 2 %>
                
            <% end if %>
            
            <% if ppid = 3 OR pid = 3 then %>
            
                <tr><td width="20"><img src="images/subnav_bar_item.png" /></td><td colspan="2"><a href="content.aspx?pid=3&ppid=0"><img src="images/subnav_sustainability.png" alt="Sustainability" style="padding-left: 5px;"/></a></td></tr>
                
                <% idOfParent = 3 %>
                
            <% end if %>
            
            <% if ppid = 4 OR pid = 4 then %>
            
                <tr><td width="20"><img src="images/subnav_bar_item.png" /></td><td colspan="2"><a href="content.aspx?pid=4&ppid=0"><img src="images/subnav_careers.png" alt="Careers" style="padding-left: 5px;"/></a></td></tr>
                
                <% idOfParent = 4 %>
                
            <% end if %>
            
            <% if ppid = 5 OR pid = 5 then %>
            
                <tr><td width="20"><img src="images/subnav_bar_item.png" /></td><td colspan="2"><td colspan="3"><a href="content.aspx?pid=5&ppid=0"><img src="images/subnav_awards.png" alt="Awards" style="padding-left: 5px;"/></a></td></tr>
                
                <% idOfParent = 5 %>
                
            <% end if %>
            

            <% if idOfParent <> 0 then %>

            

            <% for each _p As Page in aSubMenu
                if _p.parentId = idOfParent then
                    if len(_p.url) > 0 then
                        %><tr><td width="20"><img src="images/subnav_bar_item.png" /></td><td><a href="<%= _p.url %><% if InStr(_p.url,"?") > 0 then %>&<% else %>?<% end if %>pid=<%= _p.id %>&ppid=<%= _p.parentId %>"<% if _p.id = pid then%> class="current_page"<% end if %>><span class="subnav_item"><%= _p.name %></span></a></td></tr><%
                    else
                        %><tr><td width="20"><img src="images/subnav_bar_item.png" /></td><td><a href="content.aspx?pid=<%= _p.id %>&ppid=<%= _p.parentId %>"<% if _p.id = pid then%> class="current_page"<% end if %>><span class="subnav_item"><%= _p.name %></span></a></td></tr><%
                    end if
                end if
             next %>
             <% end if %>

             <% if boolMenu then %>
             <tr><td><img src="images/subnav_bar_bottom.png" /></td></tr>

             <% end if %>
    
            
            
            </table>
            
       <% end if %>
        
        
        
        
