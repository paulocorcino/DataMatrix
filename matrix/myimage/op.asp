<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp"-->
<%
	cache 1,"true"
%>
<%
dim sql,rs
select case request.QueryString("op")
case "dst"
	db "normal",,sql,"update myimage set dst=0 where id_a="&request.QueryString("id"),conn
	er(7)
	db "normal",,sql,"update myimage set dst=1 where id="&request.QueryString("id_i"),conn
	
	er(9)
	voltar site&"/matrix/myimage/myimage.asp?id="&formus("id",3)&"&area="&formus("area",3),"","_self"
case "del"
	db "leitura",rs,sql,"select foto  from myimage where id="&request.QueryString("id_i"),conn
	er(13)
	delfile(foto_root&rs("foto"))
	delfile(fotog_root&rs("foto"))
	db "normal",,sql,"delete from myimage where id="&request.QueryString("id_i"),conn
	er(17)
	voltar site&"/matrix/myimage/myimage.asp?id="&formus("id",3)&"&area="&formus("area",3),"","_self"
end select
fdb(conn)
%>