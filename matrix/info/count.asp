<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp"-->
<%
	cache 1,"true"
	dim sql
	db "normal",,sql,"update matrix_info_stat set ok=1 where uid='"&formus("uid",3)&"'",conn
%>
