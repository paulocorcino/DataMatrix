<!--#include virtual="/bymidia/matrix/config/config.asp"-->
<%
	dim rs,sql,wban,hban,nban, xban,cod_area,posban, chaveban, urlban
	chaveban = request("key")
	db "leitura",rs,sql,"select * from tb_area_sysban where chave_area_sysban='"&chaveban&"'",conn
	if not rs.eof and not rs.bof then
%>
<table width="1%" border="0" cellspacing="3" cellpadding="0">
<%

		wban = rs("w_area_sysban")
		hban = rs("h_area_sysban")
		nban = rs("n_area_sysban")
		posban = rs("pos_area_sysban")
		cod_area = rs("cod_area_sysban")
		
		
		if posban = 1 then
			response.write "<tr>"
		end if
		
		for xban = 1 to nban
			
			if posban = 1 then
				response.write "<td>"
			else
				response.write "<tr><td>"
			end if
			
			db "leitura",rs,sql,"select top 1 * from BannerShow where cod_area_sysban="&cod_area,conn	
			
			if not rs.eof and not rs.bof then
				
				if trim(rs("url_ban_sysban")) <> "" and trim(rs("url_ban_sysban")) <> "#" then
					'//tem URL
					urlban = site&"/matrix/app_banner/redirect.asp?cod="&rs("cod_ban_sysban")
				else
					'//não tem URL
					urlban = "#"
				end if
				
						if rs("tpfile_ban_sysban") = "swf" then
			%>
				<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0" width="<%=wban%>" height="<%=hban%>">
				  <param name="movie" value="<%=ban_g&trim(rs("file_ban_sysban"))%><% if urlban <> "#" then %>?url=<%=urlban%><% end if %>">
				  <param name="quality" value="high">
				  <param name="wmode" value="transparent">
				  <param name="menu" value="false">
				  <embed src="<%=ban_g&trim(rs("file_ban_sysban"))%><% if urlban <> "#" then %>?url=<%=urlban%><% end if %>" width="<%=wban%>" height="<%=hban%>" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" wmode="transparent" menu="false"></embed>
				</object>
			<%	
						else
			%>
				<% if urlban <> "#" then %><a href="<%=urlban%>" target="_blank"><% end if %><img src="<%=ban_g&trim(rs("file_ban_sysban"))%>" width="<%=wban%>" height="<%=hban%>" border="0" ><% if urlban <> "#" then %></a><% end if %>
			<%
						end if
					
			
					if posban = 1 then
						response.write "</td>"
					else
						response.write "</tr></td>"
					end if
					
					db "normal",,sql,"update tb_ban_sysban set show="&rs("show")+1&" where cod_ban_sysban="&rs("cod_ban_sysban"),conn
			end if
		next
		
		if posban = 1 then
			response.write "</tr>"
		end if
	
%>
</table>
<%
	else
		response.write "Erro na area "&chaveban
	end if
%>