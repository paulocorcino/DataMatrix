<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<%
	cache 1,true
	
	dim rs, sql, email
	
	db "leitura",rs,sql,"select cod_email, email_email from mala_email where email_email='"&formus("email",3)&"'",conn
	
	if not rs.eof and not rs.bof then
		email = trim(rs("email_email"))
		db "normal",,sql,"update mala_email set atv_email=0 where cod_email="&rs("cod_email"),conn
	end if

	
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title><%=email%> Removido da lista</title>
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
<p class="style1">Seu e-mail foi removido com sucesso da nossa lista de e-mails.</p>
<p class="style1"><strong>E-Mail removido:</strong> <%=email%> </p>
<p class="style1">Pode fechar esta janela. </p>
</body>
</html>
