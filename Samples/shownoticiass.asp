<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp"-->
<%
	dim rs, sql
	cache 1, true
	db "leitura",rs,sql,"select * from noticiaopen where cod="&request("id"),conn
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Noticia - <%=rs("titulo")%></title>
<style type="text/css">
<!--
.style1 {font-family: Verdana, Arial, Helvetica, sans-serif}
.style2 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 10px;
	font-weight: bold;
}
.style3 {font-size: 10px; font-family: Verdana, Arial, Helvetica, sans-serif;}
.style4 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 12px;
	font-style: italic;
}
.style5 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; }
#nComentItemTextoDe {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 10px;
}
#nComentItemTituloDe {
	font-size: 10px;
	font-weight: bold;
}
-->
</style>
</head>

<body>
<p class="style1"><strong><%=rs("titulo")%></strong> - <span class="style3"><%=rs("data")%></span><br />
  <span class="style2"><%=rs("jornalista")%></span></p>
<p class="style4"><%=rs("comentario")%> </p>
<p class="style5"><%=rs("texto")%></p>
<p class="style2"><%=rs("fonte")%></p>
<br />
<%
	NoticiaImgComent = "/matrix/img/edit.gif"
	NoticiaComent(request("id"))
	NoticiaCount(request("id"))
%>
</body>
</html>
