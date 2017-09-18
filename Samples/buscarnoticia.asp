<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp"-->
<%
	cache 1, true
	
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<style type="text/css">
<!--
#imgnot {
	float: left;
	margin-right: 3px;
}
.style1 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 11px;
	font-weight: bold;
}
.style2 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 11px;
}
-->
</style>
</head>

<body>
<form id="formsn" name="formsn" method="get" action="buscarnoticia.asp">
  <table width="100%" border="0" cellspacing="2" cellpadding="0">
    <tr>
      <td width="38%"><input type="text" name="query" /></td>
      <td width="62%"><select name="area" id="area">
      </select>       <input type="submit" name="Submit" value="buscar" /></td>
    </tr>
  </table>
</form>
<%AddSessionsNoticia("formsn.area")%>
<table width="100%" border="0" cellspacing="3" cellpadding="0">

<%
	dim rs, sql
	
	'//-- Busca de noticia
	dim query, codarea, buscasql
	
	query = formus("query",3) 'Campos a serem capturados
	codarea = formus("area",3)
	
	if codarea <> "" and codarea <> "0" then
		buscasql = buscasql & "codarea="&codarea&""
	end if 
	
	if query <> "" then
		buscasql = buscasql & " and texto like '%"&query&"%' or comentario like '%"&query&"%'"
	else
		buscasql = buscasql & " and texto like '$%$%'"
	end if	
	
	if buscasql <> "" then
	
		if codarea = "0" then
			buscasql = replace(buscasql," and texto"," texto")
		end if
		
		buscasql = " where " & buscasql
	end if
	'//-busca de noticia	
	
	db "leitura",rs,sql,"select top 1000 * from noticiabusca" & buscasql,conn
	while not rs.eof
%>

  <tr><td>
<div>
	<%
		if rs("foto") <> "" then
	%>
<img src="fotos/<%=rs("foto")%>" alt="" name="imgnot" width="65" id="imgnot" />
	<%
		end if
	%>
<span class="style1"><a href="shownoticia.asp?id=<%=rs("cod")%>"><%=rs("titulo")%></a></span><br />
<span class="style2"><%=rs("comentario")%></span></div>
</td> </tr>
 <%
 		rs.movenext
	wend
 %> 
 
</table>


</body>
</html>
