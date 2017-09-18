<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<%
	'Cache
	cache 1,"true"
%>
<html>
<head>
<title>contas</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/estilo.css" rel="stylesheet" type="text/css">
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
	filter: Alpha(Opacity=90);
	border: 1px solid #006633;
}
.style1 {color: #666666}
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
      <table width="457" height="252"  border="0" cellpadding="0" cellspacing="0" background="../img/logo.gif">
        <tr>
          <td><div align="center">
            <form name="form1" method="post" action="check.asp">
              <table width="213"  border="0" cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="tabele_invi">
			  <% if request.QueryString("erro")<>"" then %>
                <tr bgcolor="#F3F3F3">
                  <td colspan="2" class="fonte style2"><div align="center">&nbsp;<%=request.QueryString("erro")%></div></td>
                  </tr>
				  <% end if %>
                <tr>
                  <td colspan="2" class="fonte style1">Nome do Usu&aacute;rio </td>
                </tr>
                <tr>
                  <td colspan="2"><input name="user" type="text" class="formtextarea" id="user" size="40" maxlength="255"></td>
                </tr>
                <tr>
                  <td colspan="2" class="fonte style1">Senha</td>
                </tr>
                <tr>
                  <td colspan="2"><input name="pass" type="password" class="formtextarea" id="pass" size="40" maxlength="255"></td>
                </tr>
                <tr>
                  <td width="67"><input type="hidden" name="pag" value="<%=request.QueryString("pag")%>">
                    <div align="left">
                      <input name="Submit" type="submit" class="botao" value="Acessar">
                    </div></td>
                  <td width="136"><div align="center" class="linksimples"></div></td>
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
