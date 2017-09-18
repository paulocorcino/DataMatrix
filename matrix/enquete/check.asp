<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<%
	'protec(2)
	cache 1,false
	
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Testando Peri&oacute;dos</title>
<link href="../css/estilo.css" rel="stylesheet" type="text/css" />
</head>

<body><p class="formtextarea">
<%
	dim rs, sql
	db "leitura",rs,sql,"SELECT id, entrada, saida FROM dbo.enquete_perguntas WHERE (criado = 1) AND (entrada >= CONVERT(DATETIME, '"&fdate(formus("incio",3),"yyyy-mm-dd")&" 00:00:00', 102)) AND (saida <= CONVERT(DATETIME, '"&fdate(formus("fim",3),"yyyy-mm-dd")&" 00:00:00', 102)) OR (entrada <= CONVERT(DATETIME, '"&fdate(formus("fim",3),"yyyy-mm-dd")&" 00:00:00', 102)) AND (saida >= CONVERT(DATETIME, '"&fdate(formus("incio",3),"yyyy-mm-dd")&" 00:00:00', 102)) and (id_a = "&formus("id_a",3)&")",conn
	'response.write sql
	if rs.eof and rs.bof then
%>
	Peri&oacute;do V&aacute;lido
	<% else %>
	Peri&oacute;do Inv&aacute;lido! Tente outro peri&oacute;do. 
	<% end if %>
</p>
</body>
</html>
