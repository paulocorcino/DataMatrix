<!--#include virtual="/bymidia/matrix/config/config.asp"-->
<%
	dim rs, sql
	db "leitura",rs,sql,"select * from tb_ban_sysban where cod_ban_sysban="&formus("cod",3),conn
	db "normal",,sql,"update tb_ban_sysban set click_ban_sysban="&rs("click_ban_sysban")+1&" where cod_ban_sysban="&formus("cod",3),conn
	if not rs.eof and not rs.bof then
		response.Redirect trim(rs("url_ban_sysban"))
	end if
%>