<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<%
	dim rs, sql
	db "leitura",rs,sql,"select * from albumlista",conn
	while not rs.eof
%>
<script type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="11%" rowspan="2"><img src="<%=sites&foto_root&rs("foto")%>" width="86" height="63" hspace="5" vspace="5"></td>
        <td width="89%" id="tituloalbum"><%=rs("album")%></td>
      </tr>
      <tr>
        <td valign="top" id="comentalbum"><%=rs("comentario")%><br>
          <a href="javascript:void(0)" id="albumlink" onClick="MM_openBrWindow('album.asp?cod=<%=rs("cod")%>','<%=rs("cod")%>','width=450,height=500')">[Abrir &Aacute;lbum] </a></td>
      </tr>
    </table>
<%
		rs.movenext
	wend
%>
