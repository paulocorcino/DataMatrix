<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<%
	cache 1,true
	
	dim rs, sql
	
	db "leitura",rs,sql,"select lido_msg, cod_msg from mala_msg where cod_msg="&formus("id",3),conn
	
	if not rs.eof and not rs.bof then
		db "normal",,sql,"update mala_msg set lido_msg="& cint(rs("lido_msg")) + 1 &" where cod_msg="&rs("cod_msg"),conn
	end if
	
	response.Redirect "blank.gif"
	
%>