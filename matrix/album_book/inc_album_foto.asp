<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<%
	
	'Cache
	cache 1, true
	
	dim rs, sql, img, usrClick
	
	'Cadastra o click desta foto
	db "leitura",rs,sql,"select id, foto, click from album_fotos where id="&formus("id",3),conn
	if not rs.eof and not rs.bof then
	
		'response.write request.Cookies
		On Error Resume Next
		usrClick = cint(request.Cookies("eFoto"&rs("id"))) + 1
		
		
		If Err.Number<>0 then
			usrClick = 1
			'response.Cookies("eFoto"&rs("id")) = usrClick + 1
		else	
			usrClick = cint(usrClick)
			'response.Cookies("eFoto"&rs("id")) = usrClick + 1
		end if
			
		
		if usrClick = 1 then
			db "normal",,sql,"update album_fotos set click="& cint(rs("click")) + 1 &" where id="&rs("id"),conn
		end if
		
	
		
		'redireciona pra foto
		response.Redirect sites&fotog_root&trim(rs("foto"))
		'
	
		
	end if 
	
%>