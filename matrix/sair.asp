<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<%
	'Cache
	cache 1,"true"
%>
<html>
<head>
<title>Saindo</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
.style1 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 12px;
}
-->
</style>
</head>
<body>
<span class="style1">Aguarde .... Saindo do sistema....</span>
<%

	dim sql
	db "normal",,sql,"update auditoria set logado=0 where sessionid='"&aspid()&"'",conn
	
	fdb(conn)
	
	session.Abandon()
%>
<script>
window.close();
</script>
</body>
</html>
