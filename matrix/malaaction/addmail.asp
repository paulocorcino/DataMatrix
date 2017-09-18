<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<%
	cache 1,true
	
	dim sql, rs, grupo, email, cod_grupo, nome_email, nome, nome_email_t
	
	grupo = formus("grupo",3)
	email = split(formus("email",3),",")

	if formus("nome",3) <> "" then
		nome = split(formus("nome",3),",")
	end if
	
	db "leitura",rs,sql,"select cod_grupo, nome_grupo from mala_grupos where nome_grupo='"&grupo&"'",conn
	if rs.eof and rs.bof then
		
		db "normal",,sql,"insert into mala_grupos(nome_grupo,criado) values ('"&grupo&"',1)",conn
		db "leitura",rs,sql,"select cod_grupo, nome_grupo from mala_grupos where nome_grupo='"&grupo&"'",conn
		
	end if 
	
	'Seleciona o grupo
	cod_grupo = rs("cod_grupo")
	
	for i = lbound(email) to ubound(email)
		
		On Error Resume Next
		'vewrifica se existe varios nomes
		nome_email_t = nome(i)
		
		If Err.Number<>0 then
			if nome(i) <> "" then
				nome_email = nome(i)
				nome_email_t = nome_email 
			else
				nome_email = split(email(i),"@")
				nome_email_t = nome_email(0)
			end if
		else
			nome_email = split(email(i),"@")
			nome_email_t = nome_email(0)
		end if
		
		
		db "leitura",rs,sql,"select email_email from mala_email where email_email='"&email(i)&"'",conn
		
		if rs.eof and rs.bof then
			db "normal",,sql,"insert into mala_email(nome_email, email_email, cod_grupo, criado) values ('"&nome_email_t&"','"&email(i)&"',"&cod_grupo&",1)",conn
		end if
		
	next
	
	response.Redirect "blank.gif"	
%>