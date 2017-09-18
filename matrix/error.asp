<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
				Response.Buffer = true
				Response.Expires = 0	
				Response.Expiresabsolute = Now() - 1
				Response.AddHeader "pragma","no-cache"
				Response.AddHeader "cache-control","private"
				Response.CacheControl = "no-cache"
				
	dim javaalert
	function alert(msg)
		response.write "<script>"&vbcrlf
		response.write "//ASP Alerta - Paulo Corcino pecjs@msn.com"&vbcrlf
		msg = replace(msg,"\","\\")
		
		msg = replace(msg,"&#39;","\'")
		msg = replace(msg,"""","\""")
		msg = replace(msg,"[n]","\n")
		response.write "alert('"&msg&"')"&vbcrlf
		response.Write "</script>"&vbcrlf
	end function
	
	'Usa-se: emails "de@site.com.br","para@site.com.br","assunto","mensagem","html ou texto"
	Dim email_r
	function emails(de,para,assunto,mensagem,op)
	Set email_r = Server.CreateObject("CDONTS.NewMail")
	email_r.From = de
	email_r.Subject = assunto
	email_r.To = para
	email_r.body = mensagem
	if op = "html" then
	email_r.BodyFormat = 0
	email_r.MailFormat = 0
	else
	email_r.BodyFormat = 1
	email_r.MailFormat = 1 
	end if
	email_r.send
	Set email_r = nothing
	end function
%>
<html>
<head>
<title>Error</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
body {
	background-color: #E3F0DB;
}
.borda {
	border: 1px solid #666666;
}
.titulo {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 36px;
	font-weight: normal;
	color: #006699;
}
.texto {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 14px;
}
.style1 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 10px;
}
.style2 {font-size: 12px}
.style5 {font-size: 11px}
-->
</style></head>

<body oncontextmenu="return false" onselectstart="return false" ondragstart="return false">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td><div align="center">
      <table width="680" border="0" align="center" cellpadding="8" cellspacing="3" bgcolor="#FFFFFF" class="borda">
        <tr>
          <td height="58">&nbsp;&nbsp;&nbsp;<img src="img/error.gif" width="31" height="31" align="absmiddle">&nbsp;&nbsp;<span class="titulo">Erro 500.100</span></td>
        </tr>
        <tr>
          <td height="397" valign="top" class="texto"><p>Ocorreu um erro na execu&ccedil;&atilde;o desta p&aacute;gina. O webmaster deste site deve ser contatado para corre&ccedil;&atilde;o deste erro. <br>
            </p>
              <p> <br>
        Clique no bot&atilde;o abaixo para... <br>
        <br>
        <a href="javascript:history.go(-2)" >Continuar a navega&ccedil;&atilde;o no site</a></p>
              <p><a  href="<%=request.ServerVariables("SCRIPT_NAME")%>?av=1&d=<%=server.URLEncode(request.QueryString("d"))%>&pg=<%= request.QueryString("pg")%>&f=<%=server.URLEncode(request.QueryString("f"))%>&n=<%= request.QueryString("n")%>&lin=<%=server.URLEncode(request.QueryString("lin"))%>">Avise o erro ao Webmaster</a><br>
                <br>
                <br>
                            </p>
              <table width="90%" border="0" align="center" cellpadding="0" cellspacing="3">
                <tr>
                  <td><span class="style2">Fonte do Erro </span></td>
                </tr>
                <tr>
                  <td height="24" valign="top" class="style5"><%=request.QueryString("f")%> (<%= request.QueryString("n")%>)</td>
                </tr>
                <tr>
                  <td><span class="style2">Descri&ccedil;&atilde;o do Erro </span></td>
                </tr>
                <tr>
                  <td height="22" valign="top" class="style5"><%= request.QueryString("d")%></td>
                </tr>
                <tr>
                  <td height="14" valign="top" class="style5"><span class="style2">P&aacute;gina de Origem do erro </span></td>
                </tr>
                <tr>
                  <td height="22" valign="top" class="style5"><%= request.QueryString("pg")%>, Linha <%=request.QueryString("lin")%></td>
                </tr>
              </table>
              <p>&nbsp;</p></td>
        </tr>
        <tr>
          <td height="21" bgcolor="#66FFCC"><span class="style1">DataMaxtrixs Vr.2.5</span></td>
        </tr>
      </table>
    </div></td>
  </tr>
</table>
<%
	if request.QueryString("av")<>"" then
	dim mg
	
		mg = "<b>Erro no site [ "&request.ServerVariables("HTTP_HOST")&" ] "&now()&"</b><br><br>"
		mg = mg + "Fonte: "&request.QueryString("f") &"  "&request.QueryString("n")&"<br>"
		mg = mg + "Descrição: "&request.QueryString("d")&"<br>"
		mg = mg + "Página: "&request.QueryString("pg")&" , Linha "&request.QueryString("lin")&"<br><br>"
		mg = mg + "DataMatrixs 2004 Vr. 2.5"
		
		 emails "erro@erro.com.br","pecjs@msn.com","Error no DataMatrixs",mg,"html"
		 
		 alert("Obrigado pela ajuda.[n]WebMaster contatado com exito!")
%>
	<script>
		history.go(-3)
	</script>
<%
		 
	end if
%>
</body>
</html>

