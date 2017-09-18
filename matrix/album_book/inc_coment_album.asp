<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<%
	cache 1,true
	
	'response.write request.QueryString
	'response.write request.form
	
	dim rs, sql, nTitulo, cod_album, nTituloTexto
	
	nTitulo = split(formus("itens",3),",")
	cod_album = formus("id",3)
	
	for i=lbound(nTitulo) to ubound(nTitulo)
		
		if nTitulo(i) = "start" then
			response.write "<div id='aAlbumComentStart'><a href='javascript:void(0)' onclick='AjaxLoadAlbum(AlbumLinkDefault,null,""aAlbumThumb"",0);'>"&server.HTMLEncode(request.form("imgstart"))&"</a></div>"
		else
		
			On Error Resume Next
				db "leitura",rs,sql,"select "&nTitulo(i)&" from albumlista where cod="&cod_album,conn
				If Err.Number<>0 then
					response.write "O Campo "&nTitulo(i)&" não existe."
				else
					if not rs.eof and not rs.bof then
						nTituloTexto = ucase(left(nTitulo(i),1))&right(lcase(nTitulo(i)),len(nTitulo(i))-1)
	
						if rs(nTitulo(i)) <> "" then
							response.write "<div id='aAlbumComent"&nTitulo(i)&"'>"&nTituloTexto&"</div>"
							response.write "<div id='aAlbumComent"&nTitulo(i)&"Texto'>"&chtml(rs(nTitulo(i)))&"</div>"
						end if
						
					end if
					
				end if
			end if 
		
	next
%>
