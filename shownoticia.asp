<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<%
	cache 2, true
	dim rs, sql
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><!-- InstanceBegin template="/Templates/template.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<!-- InstanceBeginEditable name="doctitle" -->
<title>DataMatrix Demo</title>
<!-- InstanceEndEditable -->
<link href="css/estilo.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.style1 {font-size: 10px}
-->
</style>
<!-- InstanceBeginEditable name="head" --><!-- InstanceEndEditable -->
</head>

<body>
<table width="750" height="604" border="0" cellpadding="0" cellspacing="4" class="table_master">
  <tr class="Titulo">
    <td colspan="3"><img src="matrix/img/empresa.gif" width="48" height="48" hspace="6" vspace="6" align="absmiddle" />DataMatrix Demonstra&ccedil;&atilde;o </td>
  </tr>
  <tr>
    <td width="149" valign="top">
	<table width="100%" border="0" cellspacing="2" cellpadding="3">
	  <tr>
	    <td><a href="default.asp">Inicio</a></td>
        </tr>
	  <tr>
	    <td><a href="noticias.asp">Noticias</a></td>
        </tr>
	  <tr>
	    <td><a href="galeria.asp">Galeria de Fotos </a></td>
        </tr>
		<%
			db "leitura",rs,sql,"select * from matrix_cont where sessao like '@%'",conn
			while not rs.eof 
		%>
	  <tr>
	    <td><a href="conteudo.asp?id=<%=rs("id")%>"><%=replace(rs("sessao"),"@","")%></a> </td>
        </tr>
		<%
				rs.movenext
			wend
		%>

	  <tr>
	    <td><a href="#">Contato</a></td>
        </tr>
	  
    </table><br /><!-- Tempo -->	
		<table width="90%" border="0" align="center" cellpadding="0" cellspacing="3" class="table_submaster">
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
	%><br />
	    <table width="90%" border="0" align="center" cellpadding="3" cellspacing="3" class="table_submaster">
          <tr>
            <td class="titulos_sub">Informativo</td>
          </tr>
          <tr>
            <td>
<%
	OptNews = "add"
	NewsGrupo = "Matrix"
	AddNews()
%></td>
          </tr>
      </table>
	<!-- Tempo -->	
	</td>
    <td width="587" colspan="2" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="3">
      <tr>
        <td height="18" colspan="2" align="right"><%=fdate(now(),"dd de mmmm de yyyy - hh:nn")%></td>
        </tr>
      <tr>
        <td height="500" align="left" valign="top">
		<!-- InstanceBeginEditable name="Conteudo" -->
		<%
				db "leitura",rs,sql,"select * from noticiaopen where cod="&request("id"),conn
		%>
		<p class="texto_1"><strong><%=rs("titulo")%></strong> - <span class="style3"><%=rs("data")%></span><br />
		  <span class="style1"><%=rs("jornalista")%></span></p>
		<p class="texto_2"><%=rs("comentario")%> </p>
		<p class="texto_2"><%=rs("texto")%></p>
		<p class="texto_2"><%=rs("fonte")%></p>
		<br />
		<table width="90%" border="0" align="center" cellpadding="2" cellspacing="3" class="table_submaster">
          <tr>
            <td class="titulos_sub">Coment&aacute;rios</td>
          </tr>
          <tr>
            <td><%
			NoticiaImgComent = "/matrix/img/edit.gif"
			NoticiaComent(request("id"))
			NoticiaCount(request("id"))
		%></td>
          </tr>
        </table>
		
		<!-- InstanceEndEditable -->        </td>
        <td width="150" valign="top">
		<table width="90%" border="0" align="center" cellpadding="3" cellspacing="3" class="table_submaster">
  <tr>
    <td class="titulos_sub">Super fotos </td>
  </tr>
 

				<%
				db "leitura",rs,sql,"select top 4 * from AlbumListClick",conn
				while not rs.eof
			%>
               <tr><td align="center" valign="middle"><div align="center"><a href="album.asp?id=<%=rs("cod")%>"><img src="fotos/<%=trim(rs("foto"))%>" width="86" height="63" hspace="2" vspace="2" border="0" class="table_submaster" /></a></div></td> 
               </tr>
			<%
					rs.movenext
				wend
			%>

 
</table><br />

		<table width="90%" border="0" align="center" cellpadding="3" cellspacing="3" class="table_submaster">
  <tr>
    <td class="titulos_sub">Emquete</td>
  </tr>
  <tr>
    <td>
	<%
	imgEnqV = ""
	imgEnqR = ""
	eEnquete(1) 'codigo da area
%>
	</td>
  </tr>
</table><br />
<%SysBanner("parceiros")%>
<br />
<br />


		</td>
      </tr>
    </table>    </td>
  </tr>
  
  <tr>
    <td colspan="3" class="base">DataMatrix 2004 v. 2.8 - Paulo Corcino </td>
  </tr>
</table>
</body>
<!-- InstanceEnd --></html>
