<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<%
	'segurança
	protec("2")

	'Cache
	cache 1,"true"
%>
<html>
<head>
<title>contas</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/estilo.css" rel="stylesheet" type="text/css">
<META NAME='ROBOTS' CONTENT='INDEX,NOFOLLOW'>
<meta http-equiv="imagetoolbar" content="no">
<script language="JavaScript">

var msg="Direito Não";
function disableIE() {if (document.all) {return false;}
}
function disableNS(e) {
  if (document.layers||(document.getElementById&&!document.all)) {
    if (e.which==2||e.which==3) {return false;}
  }
}
if (document.layers) {
  document.captureEvents(Event.MOUSEDOWN);document.onmousedown=disableNS;
} else {
  document.onmouseup=disableNS;document.oncontextmenu=disableIE;
}
document.oncontextmenu=new Function("return false")
</script>
<style type="text/css">
<!--
body,td,th {
	font-family: MS Sans Serif, sans-serif, Arial, Helvetica;
}
body {
	background-color: #FFFFFF;
}
.style1 {color: #FF0000}
.style2 {color: #636563}
.style4 {font-size: 10px}
-->
</style></head>

<body>
<%
	Dim x, rs, sql, cor, nivel, id, nome
	
		Select case request.QueryString("op")
	'Contas ===================================================================================================================
			case "listar"
	'Contas ===================================================================================================================
					
					'Conectando ao BD, em modo somente leitura
					if nivelaspid < 2 then
						db "leitura",rs,sql,"SELECT id, nome, nivel FROM senha",conn
					else
						db "leitura",rs,sql,"SELECT id, nome, nivel FROM senha where nivel>= 2",conn
					end if
					
					'Inserir javascript de deletar contas
					prep_deletar()
						
%>
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="20" valign="top" bgcolor="#EBEBEB"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>&nbsp;Contas</strong></font></td>
  </tr>
              <tr>
                  <td height="20" valign="top"><%=erros%></td>
                </tr>
  <tr>
    <td><div align="center"> <br>
            <table width="60%" border="0" cellspacing="2" cellpadding="0">
              <tr class="fundo2">
                <td>
                  <div align="center" class="fonte"><strong><font size="1">Conta</font></strong></div></td>
                <td width="25%"><div align="center" class="fonte"><strong><font size="1">N&iacute;vel</font></strong></div></td>
                <td width="15%">
                  <div align="center" class="fonte"><strong><font size="1">Alterar</font></strong></div></td>
                <td width="15%">
                  <div align="center" class="fonte"><strong><font size="1">Deletar</font></strong></div></td>
              </tr>
<%
					'Variavel para selecionar a cor da tabela
					x = 1
					while not rs.EOF
							if x=1 then
								'Se x igual a 1 exibir cor
								cor="#FFFFFF"
								'x agora é 0
								x=0
							elseif x=0 then
								'Se x igual a 0
								cor="#FBFBEA"
								'x agora é 1
								x=1
							end if
							
							'Verificar nivel--------------------\
							if rs("nivel")="1" then
								nivel = "Administrador"
							else
								nivel = "Operador"
							end if
							
							id = cint(rs("id"))
							nome = rs("nome")
%>
              <tr bgcolor="<%=cor%>">
                <td height="20">&nbsp;<span class="fonte"><font size="1"><%=nome%></font></span></td>
                <td><div align="center"><span class="fonte"><font size="1">&nbsp;<%=nivel%></font></span></div></td>
                <td>
                  <div align="center"><font size="1"><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=alt&aspid=<%=aspid()%>&id=<%=id%>" class="linksimples">Alterar</a></font></div></td>
                <td>
                  <div align="center"><font size="1"><a href="<%=deletar (request.ServerVariables("SCRIPT_NAME")&"?op=del&aspid="&aspid()&"&id="&id,nome)%>" class="linksimples">Deletar</a></font></div></td>
              </tr>
<%
						rs.MoveNext
					Wend
					
					'Se não encontrar valor, exibir nenhuma conta
					if rs.bof and rs.eof then
					
						'Por segurança o sistemanão permite o sistema não existir conta.
						db "normal",,sql,"INSERT INTO senha(nome,nome,senha,nivel,inf) VALUES ('Administrador Matrix','matrix','matrix@2004',1,0)",conn
%>
              <tr>
                <td colspan="4"><br><div align="center" class="fonte"><font size="1">Nenhuma conta cadastrada.<br>
                      <span class="style1">A conta padr&atilde;o foi criada. </span></font>
                    <form name="form1" method="post" action="">
                      <input name="Button" type="button" class="botao" value="Continuar" onclick="window.open('<%=request.ServerVariables("SCRIPT_NAME")&"?op=listar&aspid"&aspid()%>','_self')">
                    </form>
                </div></td>
              </tr>
<%
					end if
					
					'Fechando banco de dados 
					fdb(rs)
%>			  
            </table>
            <br>
    </div></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<%
			
			'Contas ===================================================================================================================
			case "cad"
			'Contas ===================================================================================================================
				
				'Verifica se a o formulário foi enviado
					if isempty(request.form("submit")) then				
%>
 <form name="form1" method="post" action="<%=request.ServerVariables("SCRIPT_NAME")%>?op=cad&aspid=<%=aspid()%>">
 
<table width="95%" border="0" align="center" cellpadding="4" cellspacing="0">
                
                <tr> 
                  <td height="20" valign="top" bgcolor="#EBEBEB"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Nova 
                    Conta
                    </strong></font></td>
                </tr>
                <tr>
                  <td height="20" valign="top"><%=erros%></td>
                </tr>
                <tr> 
                  <td height="20" valign="top" class="fonte"><font color="#666666"><strong>Nome</strong></font></td>
                </tr>
                <tr> 
                  <td height="20" valign="top" class="fonte"><font color="#666666">&nbsp; 
                    <input name="nome_comp" type="text" class="formtextarea" id="nome_comp" value="<%=request.QueryString("nome_comp")%>" size="60" maxlength="255">
                    </font></td>
                </tr>
                <tr> 
                  <td height="20" valign="top" class="fonte"><font color="#666666"><strong>Seu 
                    e-mail</strong></font></td>
                </tr>
                <tr> 
                  <td height="20" valign="top" class="fonte"><font color="#666666">&nbsp; 
                    <input name="email" type="text" class="formtextarea" id="email" value="<%=request.QueryString("email")%>" size="40" maxlength="255">
                    </font></td>
                </tr>
                <tr> 
                  <td height="20" valign="top" class="fonte"><p><font color="#666666"><strong> 
                    &nbsp; 
                  <input name="inf" type="checkbox" class="formtextarea" id="inf" value="1" <% if request.QueryString("inf")="1" then %>checked<% end if %>>
Enviar informa&ccedil;&otilde;es do site  por e-mail <br>
                  </strong>(Voc&ecirc; recebe diariamente informa&ccedil;&otilde;es sobre o site enviadas pelo Matrix2004)</font></p>                    </td>
                </tr>
                <tr>
                  <td height="20" valign="top" class="fonte"><p><font color="#666666"><strong>N&iacute;vel<br>
                  </strong>(Escolha o n&iacute;vel do usu&aacute;rio da conta)</font></p>                  </td>
                </tr>
                <tr>
                  <td height="20" valign="top" class="fonte">
				  <% if nivelaspid < 2 then %>
				  <table width="200" border="0" cellpadding="0" cellspacing="2">
                    <tr>
                      <td><span class="style2">
                        <label>
                          <input name="nivel" type="radio" value="1">
                          <span class="style4">Administrador</span></label>
                      </span></td>
                    </tr>
                    <tr>
                      <td><span class="style2">
                        <label>
                          <input name="nivel" type="radio" value="2" checked>
                          <span class="style4">Operador</span></label>
                      </span></td>
                    </tr>
                  </table>
				  <% else %>
				  <span class="style4">Operador</span>
				  <% end if %></td>
                </tr>
                <tr> 
                  <td height="20" valign="top" class="fonte"><font color="#666666"><strong>Nome 
                    do Usu&aacute;rio</strong></font></td>
                </tr>
                <tr> 
                  <td height="20" valign="top" class="fonte"><font color="#666666">&nbsp; 
                    <input name="nome" type="text" class="formtextarea" id="nome" value="<%=request.QueryString("nome")%>" size="20" maxlength="50">
                    </font></td>
                </tr>
                <tr> 
                  <td height="20" valign="top" class="fonte"><font color="#666666"><strong>Senha 
                    do Usu&aacute;rio</strong></font></td>
                </tr>
                <tr> 
                  <td height="20" valign="top" class="fonte"><font color="#666666"> 
                    &nbsp; 
                    <input name="senha" type="password" class="formtextarea" id="senha" size="20" maxlength="50">
                    </font></td>
                </tr>
                <tr>
                  <td height="20" valign="top" class="fonte"><font color="#666666"><strong>Confimando 
                    Senha </strong></font></td>
                </tr>
                <tr>
                  <td height="20" valign="top" class="fonte"><font color="#666666"> 
                    &nbsp; 
                    <input name="senha_con" type="password" class="formtextarea" id="senha_con" size="20" maxlength="50">
                    </font></td>
                </tr>
                <tr> 
                  <td height="20" valign="top">&nbsp; 
<input name="Submit" type="submit" class="botao" value="Cadastrar" >
                    &nbsp;
                    <input name="button" type="button" class="botao" id="button" value="Cancelar" onClick="window.open('<%=request.ServerVariables("SCRIPT_NAME")%>?op=listar&aspid=<%=aspid()%>','_self')"></td>
                </tr>
              </table>
 </form>

<%
					elseif not isempty(request.form("submit")) and trim(formus("nome_comp",1))<>"" and trim(formus("email",1))<>"" and trim(formus("senha",0))<>"" then
						
						'verifique se inf esta selecionado
						if request.form("inf") = "1" then
							x="1"
						else
							x="0"
						end if
						
						'verificar se a senha foi redigitada corretamente
						if trim(formus("senha",0)) = trim(formus("senha_con",0)) then
										
							'verificar se o usuário já esta cadastrado
							
							'response.end
							 db "leitura",rs,sql, "SELECT TOP 100 PERCENT id, usuario FROM dbo.senha WHERE (usuario LIKE N'"&formus("nome",1)&"')",conn
							
							 	if rs.bof and rs.eof then
									'ok, cadastre
									if nivelaspid < 2 then
										db "normal",,sql,"INSERT INTO senha(nome,usuario,senha,email,inf,nivel) VALUES ('"&formus("nome_comp",1)&"','"&formus("nome",1)&"','"&formus("senha",0)&"','"&formus("email",0)&"',"&x&","&formus("nivel",0)&")",conn
									else
										db "normal",,sql,"INSERT INTO senha(nome,usuario,senha,email,inf,nivel) VALUES ('"&formus("nome_comp",1)&"','"&formus("nome",1)&"','"&formus("senha",0)&"','"&formus("email",0)&"',"&x&",2)",conn
									end if
									'redirecione o usuário
									voltar request.ServerVariables("SCRIPT_NAME")&"?op=listar&aspid="&aspid()," ","_self"
								else
									'erro
									voltar request.ServerVariables("SCRIPT_NAME")&"?op=cad&aspid="&aspid()&"&nome_comp="&server.URLEncode(request.Form("nome_comp"))&"&email="&server.URLEncode(request.Form("email"))&"&inf="&x&"&nome="&server.URLEncode(request.Form("nome")),"Este Nome do Usuário já esta cadastrado.","_self"
								end if
								voltar request.ServerVariables("SCRIPT_NAME")&"?op=listar&aspid="&aspid()&""," ","_self"
								'Fecha banco de dados
								fdb(rs)
								
						else
							 'erro
							 voltar request.ServerVariables("SCRIPT_NAME")&"?op=cad&aspid="&aspid()&"&nome_comp="&server.URLEncode(request.Form("nome_comp"))&"&email="&server.URLEncode(request.Form("email"))&"&inf="&x&"&nome="&server.URLEncode(request.Form("nome")),"Você confimou um senha diferente.","_self"
							 
						end if
						
					else
						'erro
						voltar request.ServerVariables("SCRIPT_NAME")&"?op=cad&aspid="&aspid()&"&nome_comp="&server.URLEncode(request.Form("nome_comp"))&"&email="&server.URLEncode(request.Form("email"))&"&inf="&x&"&nome="&server.URLEncode(request.Form("nome")),"Dados incompletos.","_self"
					end if
					
		'Contas ===================================================================================================================
		case "alt"
		'Contas ===================================================================================================================
				
				'Verifica se a o formulário foi enviado
					if isempty(request.form("submit")) then	
					
									
						'Se não existir querystring
						if request.QueryString("id")="" then

							invalido "?op=listar&aspid="&aspid(),0,0
							
						end if						
						
						'Lê banco de dados
						db "leitura",rs,sql,"select nome,usuario,id,senha,nivel,inf,inf_data,email from senha where id="&formus("id",3), conn
							
							'Se o usuário não existir
							if rs.bof and rs.eof then
															
								invalido "?op=listar&aspid="&aspid(),1,rs
								
							end if
%>
 <form name="form1" method="post" action="<%=request.ServerVariables("SCRIPT_NAME")%>?op=alt&aspid=<%=aspid()%>&id=<%=request.QueryString("id")%>">
 
<table width="95%" border="0" align="center" cellpadding="4" cellspacing="0">
                
                <tr> 
                  <td height="20" valign="top" bgcolor="#EBEBEB"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Alterar 
                    Conta
                    </strong></font></td>
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
                  <td height="20" valign="top" class="fonte"><p><font color="#666666"><strong>N&iacute;vel<br>
                  </strong>(Escolha o n&iacute;vel do usu&aacute;rio da conta)</font></p>                  </td>
                </tr>
                <tr>
                  <td height="20" valign="top" class="fonte">
				  <% if nivelaspid < 2 then%>
				  <table width="200" border="0" cellpadding="0" cellspacing="2">
                    <tr>
                      <td><span class="style2">
                        <label>
                          <input name="nivel" type="radio" value="1"  <% if rs("nivel")="1" then %>checked<%end if%>>
                          <span class="style4">Administrador</span></label>
                      </span></td>
                    </tr>
                    <tr>
                      <td><span class="style2">
                        <label>
                          <input name="nivel" type="radio" value="2" <% if rs("nivel")="2" then %>checked<%end if%>>
                          <span class="style4">Operador</span></label>
                      </span></td>
                    </tr>
                  </table>
				  <% else %>
				  	<span class="style4">Operador</span> <input type="hidden" name="nivel" value="2">
				  <% end if %></td>
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
									
									'verifica se é o unico administrador
									if trim(request.form("nivel"))="2" then
										db "leitura", rs,sql, "SELECT id, nivel FROM senha WHERE nivel=1 and id<>"&request.Form("id"),conn
											if rs.bof and rs.eof then
												'erro
												voltar request.ServerVariables("SCRIPT_NAME")&"?op=alt&aspid="&aspid()&"&id="&request.form("id"),"Este Usuário é o único Administrador, por isto não é permitido alterar o seu nível.","_self"
											else
											
												'ok, cadastre
												db "normal",,sql,"UPDATE senha SET nome='"&ac("html", request.form("nome_comp"))&"', usuario='"&ac("senha", request.form("nome"))&"', senha='"&ac("senha", request.form("senha"))&"', email='"&ac("completo", request.form("email"))&"', inf="&x&", nivel="&request.form("nivel")&" WHERE id="&request.form("id"),conn
												'redirecione o usuário
												voltar request.ServerVariables("SCRIPT_NAME")&"?op=listar&aspid="&aspid()&""," ","_self"
											end if
									else
										'ok, cadastre
										db "normal",,sql,"UPDATE senha SET nome='"&ac("html", request.form("nome_comp"))&"', usuario='"&ac("senha", request.form("nome"))&"', senha='"&ac("senha", request.form("senha"))&"', email='"&ac("completo", request.form("email"))&"', inf="&x&", nivel="&request.form("nivel")&" WHERE id="&request.form("id"),conn
										'redirecione o usuário
										voltar request.ServerVariables("SCRIPT_NAME")&"?op=listar&aspid="&aspid()&""," ","_self"
									end if				
										
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
					
		'Contas ===================================================================================================================
		case "del"
		'Contas ===================================================================================================================
		
		'Verifica se existe querystring
		if trim(request.QueryString("id"))<>"" then
		
		'verifica se este id existe
			db "leitura",rs,sql,"SELECT id FROM senha WHERE id="&request.QueryString("id"),conn
				
				if rs.bof and rs.eof then
				'se vazio volta a listar
					invalido "?op=listar&aspid="&aspid(),1,rs
				else
					'verifica se é o unico administrador cadastrado
					db "leitura",rs,sql,"SELECT id FROM senha WHERE nivel=1 and id<>"&request.QueryString("id"),conn
						if rs.bof and rs.eof then
							'se verdadeiro volta listagem
							voltar request.ServerVariables("SCRIPT_NAME")&"?op=listar&aspid="&aspid(),"Impossível deletar este usuário no momento, pois é o unico Administrador.","_self"
						else
							'Deleta o usuário selecionado
							db "normal",,sql,"DELETE FROM senha WHERE id="&request.QueryString("id"),conn
							'volta a página de listagem
							voltar request.ServerVariables("SCRIPT_NAME")&"?op=listar&aspid="&aspid()," ","_self"
						end if
				end if
		else
			'se a querystring não existe volta a tela de listagem
			invalido "?op=listar&aspid="&aspid(),0,0,0
		end if
					
					
		end select
		
			'Encerrando  o Banco de dados
			fdb(conn)
%>
</body>
</html>