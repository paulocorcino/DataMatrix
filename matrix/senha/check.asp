<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<%
	'Cache
	cache 1,"false"

'======================================================
'Sistema de seguraça DataMatrixs2004 
'Sistema cria uma identidade chamada aspid
'======================================================
'Variaveis 
	Dim sql, rs, x, sys,id
	
	db "leitura",rs,sql,"SELECT top 1 id, nome,usuario, senha, nivel FROM dbo.senha WHERE (senha LIKE '"&formus("pass",0)&"') AND (usuario LIKE '"&formus("user",0)&"')",conn
	
		if rs.eof and rs.bof then
			'verifica se a senha mestre esta cadastrada %3FS%93%3CV
	
			fdb(rs)
			voltar "login.asp?o=0","Usuário inexistente.","_self"
			
		else
		id = cint(rs("id"))
			db "normal",,sql,"insert into auditoria(id_u, nome, ip, data, logado,nivel) values("&id&",'"&rs("nome")&"','"&request.ServerVariables("REMOTE_ADDR")&"','"&fDate(date(),"yyyymmdd")&"',1,"&rs("nivel")&")",conn
			'fdb(rs)
			db "leitura",rs,sql,"SELECT sessionid,nome,nivel FROM dbo.auditoria WHERE (nome LIKE N'"&rs("nome")&"') AND (ip LIKE '"&request.ServerVariables("REMOTE_ADDR")&"') AND (logado = 1)",conn
			
			'envia o usuário para página liberada
			'
			'response.write rs("sessionid")
			voltar site&"/matrix/index_online.asp?aspid="&rs("sessionid")&"&nome="&rs("nome")&"&nivel="&rs("nivel")&"&pag=senha/logad.asp","","datamatrixs"
			fdb(rs)
		end if
'	
	fdb(conn)
%>
