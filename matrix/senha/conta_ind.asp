<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<%
	'segurança
	protec("2")

	'Cache
	cache 1,"false"
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/estilo.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
body,td,th {
	font-family: MS Sans Serif, sans-serif, Arial, Helvetica;
}
body {
	background-color: #FFFFFF;
}
-->
</style>
</head>

<body>
<%
dim rs,sql,x
select case request.QueryString("op")
	'Contas ===================================================================================================================
		case "alt"
		'Contas ===================================================================================================================
				
				'Verifica se a o formulário foi enviado
					if isempty(request.form("submit")) then	
					
									
						'Se não existir querystring
						'if request.QueryString("id")="" then
'
							'invalido "?op=listar&aspid="&aspid(),0,0
							
						'end if						
						
						'Lê banco de dados
						'response.Write aspid()
						'db "leitura",rs,sql,"SELECT TOP 100 PERCENT id_u FROM dbo.auditoria WHERE (sessionid = '"&aspid()&"')",conn
						
						'response.write 1
						db "leitura",rs,sql,"select nome,usuario,id,senha,inf,inf_data,email from senha where id="&idaspid, conn
							'response.end()
							'Se o usuário não existir
							if rs.bof and rs.eof then
															
								invalido "?op=listar&aspid="&aspid(),1,rs
								
							end if
%>
 <form name="form1" method="post" action="<%=request.ServerVariables("SCRIPT_NAME")%>?op=alt&aspid=<%=aspid()%>&id=<%=request.QueryString("id")%>">
 
<table width="95%" border="0" align="center" cellpadding="4" cellspacing="0">
                
                <tr> 
                  <td height="20" valign="top" bgcolor="#EBEBEB"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Editar minha conta </strong></font></td>
                </tr>
                <tr>
                  <td height="20" valign="top"><%=erros%></td>
                </tr>
                <tr> 
                  <td height="20" valign="top" class="fonte"><font color="#666666"><strong>Nome</strong></font></td>
                </tr>
                <tr> 
                  <td height="20" valign="top" class="fonte"><font color="#666666">&nbsp; 
                    <input name="nome_comp" type="text" class="formtextarea" id="nome_comp" value="<%=rs("nome")%>" size="60" maxlength="255">
                    </font></td>
                </tr>
                <tr> 
                  <td height="20" valign="top" class="fonte"><font color="#666666"><strong>Seu 
                    e-mail</strong></font></td>
                </tr>
                <tr> 
                  <td height="20" valign="top" class="fonte"><font color="#666666">&nbsp; 
                    <input name="email" type="text" class="formtextarea" id="email" value="<%=rs("email")%>" size="40" maxlength="255">
                    </font></td>
                </tr>
                <tr> 
                  <td height="20" valign="top" class="fonte"><p><font color="#666666"><strong> 
                    &nbsp; 
                  <input name="inf" type="checkbox" class="formtextarea" id="inf" value="1" <% if rs("inf")="1" then %>checked<% end if %>>
Enviar informa&ccedil;&otilde;es do site  por e-mail <br>
                  </strong>(Voc&ecirc; recebe diariamente informa&ccedil;&otilde;es sobre o site enviadas pelo Matrix2004)</font></p>                    </td>
                </tr>
                <tr> 
                  <td height="20" valign="top" class="fonte"><font color="#666666"><strong>Nome 
                    do Usu&aacute;rio</strong></font></td>
                </tr>
                <tr> 
                  <td height="20" valign="top" class="fonte"><font color="#666666">&nbsp; 
                    <input name="nome" type="text" class="formtextarea" id="nome" value="<%=rs("usuario")%>" size="20" maxlength="50">
                    </font></td>
                </tr>
                <tr> 
                  <td height="20" valign="top" class="fonte"><font color="#666666"><strong>Senha 
                    do Usu&aacute;rio</strong></font></td>
                </tr>
                <tr> 
                  <td height="20" valign="top" class="fonte"><font color="#666666"> 
                    &nbsp; 
                    <input name="senha" type="password" class="formtextarea" id="senha" size="20" maxlength="50" value="<%=rs("senha")%>">
                    </font></td>
                </tr>
                <tr>
                  <td height="20" valign="top" class="fonte"><font color="#666666"><strong>Confimando 
                    Senha </strong></font></td>
                </tr>
                <tr>
                  <td height="20" valign="top" class="fonte"><font color="#666666"> 
                    &nbsp; 
                    <input name="senha_con" type="password" class="formtextarea" id="senha_con" size="20" maxlength="50" value="<%=rs("senha")%>">
                    </font></td>
                </tr>
                <tr> 
                  <td height="20" valign="top">&nbsp; <input type="hidden" value="<%=rs("id")%>" name="id">
<input name="Submit" type="submit" class="botao" value="Alterar" >
                    &nbsp;
                    <input name="button" type="button" class="botao" id="button" value="Cancelar" onClick="window.open('<%=request.ServerVariables("SCRIPT_NAME")%>?op=listar&aspid=<%=aspid()%>','_self')"></td>
                </tr>
              </table>
 </form>

<%
					elseif not isempty(request.form("submit")) and trim(formus("nome_comp",1))<>"" and trim(formus("email",0))<>"" and trim(formus("senha",0))<>"" then
						
						'verifique se inf esta selecionado
						if request.form("inf") = "1" then
							x="1"
						else
							x="0"
						end if
						
						'verificar se a senha foi redigitada corretamente
						if trim(formus("senha",0)) = trim(request.form("senha_con")) then
										
							'verificar se o usuário já esta cadastrado
							 db "leitura",rs,sql,"SELECT TOP 100 PERCENT id, usuario FROM dbo.senha WHERE (usuario LIKE N'"&ac("senha", request.form("nome"))&"') AND (id <> "&formus("id",0)&")",conn
							
							 	if rs.bof and rs.eof then
									

										'ok, cadastre
										db "normal",,sql,"UPDATE senha SET nome='"&ac("html", request.form("nome_comp"))&"', usuario='"&ac("senha", request.form("nome"))&"', senha='"&ac("senha", request.form("senha"))&"', email='"&ac("completo", request.form("email"))&"', inf="&x&" WHERE id="&request.form("id"),conn
										'redirecione o usuário
										voltar request.ServerVariables("SCRIPT_NAME")&"?op=listar&aspid="&aspid()&""," ","_self"
	
										
								else
									'erro
									voltar request.ServerVariables("SCRIPT_NAME")&"?op=alt&aspid="&aspid()&"&id="&request.form("id"),"Este Nome do Usuário já esta cadastrado.","_self"
								end if
								
								'Fecha banco de dados
								fdb(rs)
								
						else
							 'erro
							 voltar request.ServerVariables("SCRIPT_NAME")&"?op=alt&aspid="&aspid()&"&id="&request.form("id"),"Você confimou um senha diferente.","_self"
							 
						end if
						
					else
						'erro
						voltar request.ServerVariables("SCRIPT_NAME")&"?op=alt&aspid="&aspid()&"&id="&request.form("id"),"Dados incompletos.","_self"
					end if
			case "listar"
				
				voltar "../blank.asp?op=listar&aspid="&aspid(),"","_self"
			end select 
			fdb(conn)
					%>
</body>
</html>
