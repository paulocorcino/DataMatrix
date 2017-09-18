<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<%
	cache 2, true
	dim rs, sql
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>DataMatrix Demo</title>
<link href="css/estilo.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.style1 {font-size: 10px}
-->
</style>
</head>

<body>
<table width="750" height="604" border="0" cellpadding="0" cellspacing="4" class="table_master">
  <tr class="Titulo">
    <td colspan="3"><img src="matrix/img/empresa.gif" width="48" height="48" hspace="6" vspace="6" align="absmiddle" />DataMatrix Demonstra&ccedil;&atilde;o </td>
  </tr>
  <tr>
    <td width="149" valign="top">
	<!-- Tempo -->	
		<table width="80%" border="0" align="center" cellpadding="0" cellspacing="3" class="table_submaster">
          <tr class="titulos_sub">
            <td colspan="3">Tempo Hoje </td>
          </tr>
          <tr>
            <td colspan="3"><div id="cidade"></div></td>
          </tr>
          <tr>
            <td rowspan="2" align="center" valign="middle"><div id="imgtmp"></div></td>
            <td><span class="style1">Min</span></td>
            <td><div id="temMin"></div></td>
          </tr>
          <tr>
            <td><span class="style1">Max</span></td>
            <td><div id="temMax"></div></td>
          </tr>
          <tr>
            <td colspan="3"><div id="estado"></div><div id="comenttmp"></div></td>
          </tr>
        </table>
	    <%
		tempoEstado = "br"
		imgTempo = ""
		timerTempo = 7
		AspTempo(0)
	%>
	<!-- Tempo -->	</td>
    <td width="587" colspan="2" rowspan="2" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="3">
      <tr>
        <td height="18" colspan="2" align="right"><%=fdate(now(),"dd de mmmm de yyyy - hh:nn")%></td>
        </tr>
      <tr>
        <td height="500" align="left" valign="top"><p class="titulo_areas">&nbsp;</p>
		<!-- Noticias -->
        <!-- Noticias -->
        </td>
        <td width="150" valign="top">&nbsp;</td>
      </tr>
    </table>
    </td>
  </tr>
  <tr>
    <td height="395" valign="top"><table width="100%" border="0" cellspacing="2" cellpadding="3">
      <tr>
        <td><a href="default.asp">Inicio</a></td>
      </tr>
      <tr>
        <td><a href="#">Noticias</a></td>
      </tr>
      <tr>
        <td><a href="#">Galeria de Fotos </a></td>
      </tr>
      <tr>
        <td><a href="#">Sobre o DataMatrix</a> </td>
      </tr>
      <tr>
        <td><a href="#">Fotos do DataMatrix</a> </td>
      </tr>
      <tr>
        <td><a href="#">Contato</a></td>
      </tr>

    </table></td>
  </tr>
  <tr>
    <td colspan="3" class="base">DataMatrix 2004 v. 2.8 - Paulo Corcino </td>
  </tr>
</table>
</body>
</html>
