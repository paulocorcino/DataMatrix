<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Emoticom</title>
<style type="text/css">
<!--
.border {
	border: 1px solid #00CCCC;
}
body {
	margin: 0px;
}
-->
</style>
</head>

<body onLoad="focus()">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="middle"><table width="1%" border="0" cellpadding="4" cellspacing="1">
<%
		'Subistitui os codigos pelos icones	
	dim icos, ix, nCol, nXn, imgs
	nCol = 10
	icos = Split(tagIco, ",")
	imgs = Split(tagImg, ",")
	nXn  = 0
	
	For ix=0 to Ubound(icos)
	
		if nXn = 0 then
			response.write "<tr>"
		end if
		
			 response.write "<td class='border'><a href='javascript: void(0);' onclick='javascript:window.newmsg.document.getElementById(""ta"").value==""oi""'><img src='../../img/icones_mural/"&imgs(ix)&".gif' alt='' width='19' height='19' border='0' /></a></td>"
			 nXn = nXn + 1
			 
		if nXn = nCol then 
			response.write "</tr>"
			nXn = 0
		end if
		
	next
%>
</table></td>
  </tr>
</table>

</body>
</html>
