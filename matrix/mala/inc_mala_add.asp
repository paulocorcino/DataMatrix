<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<%
	cache 1,true
	dim rs, sql, nome_email, email_email, nome_grupo, cod_grupo
	
	select case request("op")
	case "add"
	
		if request("post")<>"" then
		
			nome_email = formus("nomeMala",0)
			email_email = formus("emailMala",0)
			nome_grupo = formus("nome_grupo",0)
			
			db "leitura",rs,sql,"select cod_grupo from mala_grupos where nome_grupo='"&nome_grupo&"'",conn
			if rs.eof and rs.bof then				
				db "normal",,sql,"insert into mala_grupos(nome_grupo,criado) values ('"&nome_grupo&"',1)",conn
				db "leitura",rs,sql,"select cod_grupo from mala_grupos where nome_grupo='"&nome_grupo&"'",conn
			end if
			
			cod_grupo = rs("cod_grupo")	
			
			db "leitura", rs, sql, "select email_email, atv_email, cod_email from mala_email where email_email='"&email_email&"'",conn
			if rs.eof and rs.bof then
				db "normal",,sql, "insert into mala_email(nome_email, email_email, cod_grupo, criado) values ('"&nome_email&"','"&email_email&"',"&cod_grupo&",1)",conn
			else
				if cint(rs("atv_email")) = 0 then
					db "normal",,sql,"update mala_email set atv_email=1 where email_email='"&email_email&"'",conn
				end if
			end if 
			
			'alert("Seu E-Mail foi adiconado em nossa lista com sucesso!")
		end if
	
%>

	<div id="nNewsMsg">Receba nossos Informativos</div>
	<div id="nNewsMsgNome">Seu Nome</div>
	<input name="nomeMala" type="text" id="nomeMala" /><br />
	<div id="nNewsMsgEmail">Seu E-Mail</div>
	<input name="emailMala" type="text" id="emailMala" onBlur="if(!isEmail(this.value)){this.value=''}" /> 
	<br />
	<input type="button" name="NewsButton" value="Cadastrar" onclick="nFormNewAdd()" />
	<input name="post" type="hidden" value="1">
	<input name="nome_grupo" type="hidden" value="<%=request("nome_grupo")%>">
	<input name="op" type="hidden" value="add">

<%
	case "rem"
	
		email_email = formus("emailMala",0)
		
		if request("post")<>"" then
			db "normal",,sql,"update mala_email set atv_email=0 where email_email='"&email_email&"'",conn
			'alert("Seu E-Mail foi removido da nossa lista com sucesso!")
		end if

%>
	<div id="nNewsMsg">Remova seu e-mail de nossa lista</div>
	<div id="nNewsMsgRem">Digite o e-mail a ser removido</div>
	<input name="emailMalaRemove" type="text" id="emailMalaRemove" onBlur="if(!isEmail(this.value)){this.value=''}" />
	<br />
	<input type="button" name="NewsButtonRem" value="Remover" onclick="nFormNewRem()" />
	<input name="post" type="hidden" value="1">
	<input name="op" type="hidden" value="rem">
<%
	case else
		response.write "Parametro OptNews inválido"
	end select
%>