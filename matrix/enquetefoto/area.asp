<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<html>
<head>
<title>Enquetes</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/estilo.css" rel="stylesheet" type="text/css">
</head>

<body>
<%

	'Declara as variaveis 
	dim sql, rs
	
	'Seleciona a tela apartir da querystring [op]
	select case request.QueryString("op")
	
	'Tela de Cadastro
	case "cad"
	
		if isempty(request.form("submit")) then
%>
<table width="100%"  border="0" cellspacing="2" cellpadding="0">
  <tr>
    <td height="22" class="fundo3"><strong class="fonte">&nbsp;Cadastro de Areas </strong></td>
  </tr>
  <tr>
    <td><%=erros%></td>
  </tr>
  <tr>
    <td><form name="form1" method="post" action="<%=request.ServerVariables("SCRIPT_NAME")%>?op=cad">
      <table width="95%"  border="0" align="center" cellpadding="0" cellspacing="2">
        <tr>
          <td class="fonte">Nome da Area <br></td>
        </tr>
        <tr>
          <td><input name="area" type="text" class="formtextarea" id="area" size="50" maxlength="255"></td>
        </tr>

          <tr>
            <td><input name="Submit" type="submit" class="botao" value="Cadastrar">
            <input name="Submit2" type="button" class="botao" value="Cancelar" onClick="window.open('<%=request.ServerVariables("SCRIPT_NAME")%>?op=listar','_self')"></td>
          </tr>
        </table>
    </form></td>
  </tr>
</table>
<%
		elseif not isempty(request.form("submit")) and trim(formus("area",1))<>"" then
		
			db "leitura",rs,sql,"select id from enquetefoto_areas where areas='"&formus("area",1)&"'",conn
				if rs.eof and rs.bof then
					db "normal",,sql,"insert into enquetefoto_areas(areas) values ('"&formus("area",1)&"')",conn
					voltar request.ServerVariables("SCRIPT_NAME")&"?op=cad","Criado com sucesso","_self"
				else
					voltar request.ServerVariables("SCRIPT_NAME")&"?op=cad","Esta Área já foi criada","_self"
				end if
	
		else
		
			voltar request.ServerVariables("SCRIPT_NAME")&"?op=cad","É necessário preencher todos os Campos","_self"
	
		end  if
		
	'Tela de Alteração
	case "alt"
	
		if isempty(request.form("submit")) then
		
		db "leitura",rs,sql,"select id, areas from enquetefoto_areas where id="&request.QueryString("id"),conn
		if rs.eof and rs.bof then
			voltar request.ServerVariables("SCRIPT_NAME")&"?op=listar","","_self"
			fdb(rs)
			fdb(conn)
			response.end()	
		end if
%>
<table width="100%"  border="0" cellspacing="2" cellpadding="0">
  <tr>
    <td height="22" class="fundo3"><strong class="fonte">&nbsp;Alterar Area </strong></td>
  </tr>
  <tr>
    <td><%=erros%></td>
  </tr>
  <tr>
    <td><form name="form1" method="post" action="<%request.ServerVariables("SCRIPT_NAME")%>?op=alt">
      <table width="95%"  border="0" align="center" cellpadding="0" cellspacing="2">
        <tr>
          <td class="fonte">Nome da Area <br></td>
        </tr>
        <tr>
          <td><input name="area" type="text" class="formtextarea" id="area" size="50" maxlength="255" value="<%=trim(rs("areas"))%>"><input name="id" type="hidden" value="<%=rs("id")%>" ></td>
        </tr>

          <tr>
            <td><input name="Submit" type="submit" class="botao" value="Cadastrar">
            <input name="Submit22" type="button" class="botao" value="Cancelar" onClick="window.open('<%=request.ServerVariables("SCRIPT_NAME")%>?op=listar','_self')"></td>
          </tr>
            </table>
    </form></td>
  </tr>
</table>
<%
	elseif not isempty(request.Form("submit")) and trim(formus("area",1))<>"" then
		
		db "leitura",rs,sql,"select id from enquetefoto_areas where areas='"&formus("area",1)&"' and id<>"&formus("id",0),conn
		if rs.bof and rs.eof then
			db "normal",,sql,"update enquetefoto_areas set areas='"&formus("areas",1)&"' where id="&formus("id",0),conn
			voltar request.ServerVariables("SCRIPT_NAME")&"?op=listar","","_self"
		else
			voltar request.ServerVariables("SCRIPT_NAME")&"?op=alt&id="&formus("id",0),"Esta area já esta cadastrada","_self"
		end if
		
	else
		voltar request.ServerVariables("SCRIPT_NAME")&"?op=alt&id="&formus("id",0),"É necessário preencher alguma coisa","_self"
	end if
	
	'Tela de Deletar
	case "del"
		db "leitura",rs,sql,"select id from enquetefoto_areas where id="&request.QueryString("id"),conn
		if rs.eof and rs.bof then
			voltar request.ServerVariables("SCRIPT_NAME")&"?op=listar","","_self"
		else
			db "normal",,sql,"delete from enquetefoto_areas where id="&request.QueryString("id"),conn
			voltar request.ServerVariables("SCRIPT_NAME")&"?op=listar","","_self"
		end if

	'Tela de listagem
	case else
		 db "leitura",rs,sql,"select * from enquetefoto_areas",conn
		 prep_deletar()
%>
<table width="100%"  border="0" cellspacing="2" cellpadding="0">
  <tr>
    <td height="22" class="fundo3"><strong class="fonte">&nbsp;Listar Areas </strong></td>
  </tr>
  <tr>
    <td><%=erros%></td>
  </tr>
  <tr>
    <td><div align="center">
      <table width="500"  border="0" cellspacing="2" cellpadding="0">
        <tr class="fundo1">
          <td width="39"><div align="center" class="fonte">C&oacute;d.</div></td>
          <td width="260"><div align="center" class="fonte">&Aacute;rea</div></td>
          <td width="91" class="fonte"><div align="center">Alterar</div></td>
          <td width="100"><div align="center" class="fonte">Deletar</div></td>
        </tr>
		<%
		while not rs.eof
		%>
        <tr class="fonte">
          <td><div align="center"><%=rs("id")%></div></td>
          <td><div align="left">&nbsp;&nbsp;<%=trim(rs("areas"))%></div></td>
          <td><div align="center"><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=alt&id=<%=rs("id")%>">Alterar</a></div></td>
          <td><div align="center"><a href="<%=deletar(request.ServerVariables("SCRIPT_NAME")&"?op=del&id="&rs("id"),rs("areas"))%>" >Deletar</a></div></td>
        </tr>
		<%
			rs.movenext
		wend
		if rs.eof and rs.bof then
		%>
        <tr>
          <td colspan="4"> <div align="center" class="fonte">Nenhuma &aacute;rea cadastrada </div></td>
		  <%
		  	end if
		  %>
          </tr>
      </table>
    </div></td>
  </tr>
</table>
<%
	end select
%>
</body>
</html>
