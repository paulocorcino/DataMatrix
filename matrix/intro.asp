<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>


<!--#include virtual="/bymidia/matrix/contador/dados.asp" -->
<%
	'segurança
	seguranca("2")
%>
<html>
<head>
<title>Intro</title>
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
body {
	background-color: #FFFFFF;
}
.style1 {
	font-size: 13px;
	font-weight: bold;
}
.style3 {font-size: 10px}
.style4 {color: #999999}
-->
</style></head>

<body class="fonte1"  oncontextmenu="return false" onselectstart="return false" ondragstart="return false">
<span class="style1"> Informa&ccedil;&otilde;es do Site - <%=now()%></span>
<hr size="1" noshade>
<br>
<table width="448"  border="0" cellpadding="3" cellspacing="2">
  <tr bgcolor="#F5F5F5" class="fonte2">
    <td width="378" bgcolor="#F5F5F5"><span class="style3">Visitantes On-Line </span></td>
    <td width="64"><div align="center"><span class="style3"><%=Application( "usersOnline" )%></span></div></td>
  </tr>
  <tr class="style3">
    <td>Visitas Hoje </td>
    <td><div align="center" class="style4"><%=sVisitorsToday%></div></td>
  </tr>
  <tr bgcolor="#F5F5F5" class="style3">
    <td><span class="style4">Total de Visitas </span></td>
    <td><div align="center" class="style4"><%=sVisitorsTotal%></div></td>
  </tr>
  <tr bgcolor="#FFFFFF" class="style3">
    <td>&nbsp;</td>
    <td><form name="form1" method="post" action="">
      <div align="center">
        <input name="Submit" type="submit" class="botao" value="Atualizar" onClick="window.open('<%=request.ServerVariables("SCRIPT_NAME")%>','_self')">
      </div>
    </form></td>
  </tr>
</table>
</body>
</html>
