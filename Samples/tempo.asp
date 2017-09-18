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
</head>

<body>

		<div id="imgtmp"></div>							
		<div id="cidade"></div>
		<div id="estado"></div>
		<div id="temMax"></div>
		<div id="temMin"></div>
		<div id="comenttmp"></div>
		<pre>
	<%
		tempoEstado = "sp"
		imgTempo = ""
		timerTempo = 7
		AspTempo(0)
	%>
	</pre>
</body>
</html>
