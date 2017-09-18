<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<%
	'protec(2)
	cache 1,false
	
	dim rs, sql
	db "leitura",rs,sql,"SELECT id, entrada, saida FROM dbo.enquetefoto_perguntas WHERE (criado = 1) AND (id_a = "&formus("id_a",0)&") and id<>"&formus("id",0)&" and (entrada >= CONVERT(DATETIME, '"&fdate(formus("incio",0),"yyyy-mm-dd")&" 00:00:00', 102)) AND (saida <= CONVERT(DATETIME, '"&fdate(formus("fim",0),"yyyy-mm-dd")&" 00:00:00', 102)) OR (criado = 1) AND (id_a = "&formus("id_a",0)&") and id<>"&formus("id",0)&" and  (entrada <= CONVERT(DATETIME, '"&fdate(formus("fim",0),"yyyy-mm-dd")&" 00:00:00', 102)) AND (saida >= CONVERT(DATETIME, '"&fdate(formus("incio",0),"yyyy-mm-dd")&" 00:00:00', 102))",conn
	'response.write sql
	if rs.eof and rs.bof then
		response.write 1
	else
		response.write 0
	end if
%>

