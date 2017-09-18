<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp"-->
<%
	dim rs,sql
	db "leitura",rs,sql,"select top 1 * from BannerShow where chave_area_sysban='"&request("key")&"'",conn	
			
		if not rs.eof and not rs.bof then
			db "normal",,sql,"update tb_ban_sysban set show="&rs("show")+1&" where cod_ban_sysban="&rs("cod_ban_sysban"),conn
			response.Redirect sites&ban_g&trim(rs("file_ban_sysban"))
		else
			if request("key")="albumpub" then
				response.redirect sites&"/matrix/album_book/inc_album_foto.asp?id="&request("id")
			end if
		end if
%>