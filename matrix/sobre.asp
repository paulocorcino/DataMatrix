<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<%
	cache 1,"true"
%>

<html>
<head>
<title>contas</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="css/estilo.css" rel="stylesheet" type="text/css">
<META NAME='ROBOTS' CONTENT='INDEX,NOFOLLOW'>
<meta http-equiv="imagetoolbar" content="no">
<script language="JavaScript">

var msg="Direito Não";
function disableIE() {if (document.all) {return false;}
}
function disableNS(e) {
  if (document.layers||(document.getElementById&&!document.all)) {
    if (e.which==2||e.which==3) {return false;}
  }
}
if (document.layers) {
  document.captureEvents(Event.MOUSEDOWN);document.onmousedown=disableNS;
} else {
  document.onmouseup=disableNS;document.oncontextmenu=disableIE;
}
document.oncontextmenu=new Function("return false")
</script>
<style type="text/css">
<!--
.tabele_invi {
	filter: Alpha(Opacity=70);
	border: 1px solid #006633;
}
body {
	background-color: #FFFFFF;
}
.style2 {color: #FF0000}
-->
</style>
</head>

<body oncontextmenu="return false" onselectstart="return false" ondragstart="return false">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="middle"><div align="center">
      <table width="457" height="326"  border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td height="326"><div align="center">
            <form name="form1" method="post" action="check.asp">
              <table width="349"  border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="tabele_invi">
			
                <tr bgcolor="#F3F3F3">
                  <td height="149" align="center" valign="middle" class="fonte style2"><div align="center">
                    <p align="center"><strong>Programador<br>
                    </strong>Paulo Corcino - pecjs@msn.com<br>
                    </p>
                    <p align="center"><strong>Vers&atilde;o 2.5</strong><br>
                      Novo Banco de dados - SQLServer <br>
                      Novos sistema de Noticia<br>
                      Novo sistema de enquetes<br>
                      Sistema de Banner<br>
                      Sistema de materias<br>
                      Revis&atilde;o de C&oacute;digo Fonte </p>
                    <p align="center"><strong>Vers&atilde;o 2.0</strong><br>
                        Novo Design<br>
                        Revis&atilde;o do C&oacute;digo fonte<br>
                        Revis&atilde;o do Banco de dados<br>
                        Novo sistema de noticias<br>
                        Sistema de Album<br>
                        Sistema de Estatisticas Completa<br>
                        Novo sistema de Conte&uacute;do<br>
                        <br>
                        <br>
                        <strong>Vers&atilde;o 1.5</strong><br>
                        Novo sistema de noticias<br>
                        Novo sistema de conte&uacute;do<br>
                        Revis&atilde;o do c&oacute;digo fonte do sistema <br>
                        <br>
                        <strong>Vers&atilde;o 1.0</strong><br>
                        Primeira vers&atilde;o<br>                      
                        <br>
                        Paulo Corcino  - 2000 &nbsp;2004 - Todos os Direitos Reservados.</p>
                  </div></td>
                  </tr>
              </table>
            </form>
          </div></td>
        </tr>
      </table>
    </div></td>
  </tr>
</table>
</body>
</html>
