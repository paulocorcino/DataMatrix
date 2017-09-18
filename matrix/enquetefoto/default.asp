<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<%
	cache 1,true
	protec(2)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Foto Enquete</title>
<link href="../css/estilos.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" type="text/javascript" src="../js/script.js"></script>
<script language="JavaScript" type="text/javascript" src="../js/form_fotoenquete.js"></script>
<script language="JavaScript" type="text/javascript" src="../js/form.js"></script>
<script>AspId = '<%=aspid%>'</script>
</head>
<!--[Info File]-->
<!--[!]
	Enquete Foto
	Criado: Paulo Corcino		
	@Ano: 2006 - @Version: 1.5
[!]-->
<!--[Info File]-->
<body>
<form method="post"  name="form" id="form">
  <table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td id="titulo_pagina"><div id="matrix_titulo" style="float:left">Foto Enquete</div></td>
    </tr>
	<%
		dim rs,sql,id, criado, x
		select case request.QueryString("op")
		case "cad_area"
		' // Listar Area ------------------------------------------------ //
		
		dim area
		
			if lcase(request.ServerVariables("REQUEST_METHOD")) <> "post" then
			'// Lista campos 
			
			id = formus("id",3)
			
			if id<>"" then
				subTitle("> Alterar Área")
				'//Alterar Dados
				db "leitura",rs,sql,"select top 1 * from enquetefoto_areas where id="&id,conn
				area = trim(rs("areas"))
				
			else
				subTitle("> Nova Área")
				'//nova area
				db "leitura",rs,sql,"select top 1 * from enquetefoto_areas where criado=0",conn
				if rs.eof and rs.bof then
					'//não existe registro
					db "normal",,sql,"insert into enquetefoto_areas(criado) values(0)",conn
					db "leitura",rs,sql,"select top 1 * from enquetefoto_areas where criado=0",conn
				end if						
				
			end if
			
			id = rs("id")
			criado = rs("criado")
	%>	
	
    <tr>
      <td id="menutab">
	  <!--[menu]-->
	   <%
	  	'// Gera Menus ----------------------------- >
		'// 0-menu;1-voltar;2-imprimir;3-confirmar;4-incluir;
		'// 5-apagar;6-editar;7-reportar;8-mostrarall;
		'// 9-ocultarall;10-item.
		'//----------------------------------------- >
		dim AspMenuOp(2) 'Define o numero de itens
		'//campos ---------------------------------
		AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?aspid=' + AspId)"
		AspMenuOp(1) = "Salvar|3|miniMenu_area()"
		'//----------------------------------------
		AspMenu() 'gera menu
		'// Gera Menus ---------------------------- >
		%>
       <!--[menu]--></td>
    </tr>
    <tr>
      <td id="conteudo_page">
	  <!--[conteudo]-->
	  <span class="texto_size_10">Cod</span><br />
      <input name="id" type="text" id="id" value="<%=id%>" size="3" readonly="true" />
      <br />
      <span class="texto_size_10">Nome da &Aacute;rea</span><br />
<input name="area" type="text" id="area" value="<%=area%>" size="30" maxlength="255" /><input name="criado" type="hidden" value="<%=criado%>" />
	  <!--[conteudo]--></td>
    </tr>
	
	<%
			else
			
				id = formus("id",0)
				area = formus("area",0)
				criado = cint(formus("criado",0))
				'///--------- [salva dados] ---->
				db "leitura",rs,sql,"select * from enquetefoto_areas where id<>"&id&" and areas='"&area&"'",conn
				'verifica se existe algo cadastrado com este Nome
					if rs.eof and rs.bof then 
					
						db "normal",,sql,"update enquetefoto_areas set areas='"&area&"', criado=1 where id="&id,conn
						'//Informação salva
						alert("Área salva com sucesso!")
						
						if criado = 0 then
							voltar request.ServerVariables("SCRIPT_NAME")&"?aspid="&aspid&"&op=cad_area","","_self"
						else
							voltar request.ServerVariables("SCRIPT_NAME")&"?aspid="&aspid&"&op=cad_area&id="&id,"","_self"
						end if
						
					else
						'// este registro já existe
						alert("Está área já foi cadastrada!")
						voltar request.ServerVariables("SCRIPT_NAME")&"?aspid="&aspid&"&op=cad_area&id="&id,"","_self"
											
					end if		
			
			end if
			
		case "list_area"
		' // Listar Área --------------------------------------------------- //
		
	%>
    <tr>
      <td id="menutab"><!--[menu]-->
          <%
	  	'// Gera Menus ----------------------------- >
		'// 0-menu;1-voltar;2-imprimir;3-confirmar;4-incluir;
		'// 5-apagar;6-editar;7-reportar;8-mostrarall;
		'// 9-ocultarall;10-item.
		'//----------------------------------------- >
		redim AspMenuOp(2) 'Define o numero de itens
		'//campos ---------------------------------
		AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?aspid=' + AspId)"
		AspMenuOp(1) = "Novo|4|WindowOpen('default.asp?op=cad_area&aspid=' + AspId)"
		'AspMenuOp(1) = "Salvar|3|miniMenu_area()"
		'//----------------------------------------
		AspMenu() 'gera menu
		'// Gera Menus ---------------------------- >
		%>
          <!--[menu]-->      </td>
    </tr>
    <tr>
      <td id="conteudo_page"><!--[conteudo]-->
          <div align="center">
            <%
		  subTitle("> Listar Área")
		  Dim GridData(3), GridName(3), GridSQL, GridDate, GridPoss
				
				
					GridSQL = "select * from enquetefoto_areas where criado=1;"
					
					GridName(0) = "Cod;35"
					GridName(1) = "Área;370"
					GridName(2) = "Alterar;60"
					GridName(3) = "Deletar;60"
									
					GridData(0) = "id;1;0"
					GridData(1) = "areas;0;0"
					GridData(2) = "*Alterar;1;default.asp?op=cad_area&id=[col0]&aspid=[aspid]&m=[mod]"
					GridData(3) = "*Deletar;1;javascript: deletar('"&request.ServerVariables("SCRIPT_NAME")&"?op=del&tp=area&aspid=[aspid]&id=[col0]&m=[mod]','[col1]')"
							
					GridDate = "dd/mm/yyyy"
					GridPoss = "center"
					'GridURLS = 1
					'// Chama a Função
							
					GridForm Gkeys,15
					
	%>
          </div>
        <!--[conteudo]-->      </td>
    </tr>
    <%
		case "cad_enquete"
		' // Cadastrar Enquete -------------------------------------------- //
		
		dim id_a, pergunta, entrada, alt, totalvotos, repeteuser, ativo, saida, novaenq
		novaenq = true
		id = formus("id",3)
			
			if id<>"" then
				subTitle("> Alterar Enquete")
				response.write "<script>var enquete_tp = 'upd';</script>"
				novaenq = false
				'//Alterar Dados
				db "leitura",rs,sql,"select top 1 * from enquetefoto_perguntas where id="&id,conn
				
				id_a = cint(trim(rs("id_a")))
				pergunta = trim(rs("pergunta"))
				entrada = fdate(rs("entrada"),"dd/mm/yyyy")
				alt = trim(rs("alt"))
				totalvotos = trim(rs("totalvotos"))
				repeteuser = trim(rs("repeteuser"))
				saida = fdate(rs("saida"),"dd/mm/yyyy")
											
			else
				subTitle("> Nova Enquete")
				response.write "<script>var enquete_tp = 'new';</script>"
				
				'//nova area
				
				db "leitura",rs,sql,"select top 1 * from enquetefoto_perguntas where criado=0",conn
				if rs.eof and rs.bof then
					'//não existe registro
					db "normal",,sql,"insert into enquetefoto_perguntas(criado) values(0)",conn
					db "leitura",rs,sql,"select top 1 * from enquetefoto_perguntas where criado=0",conn
				end if		
				
				entrada = fdate(now(),"dd/mm/yyyy")
						
				
			end if
			
			id = rs("id")
			criado = rs("criado")
		
	%>
	    <tr>
      <td id="menutab"><!--[menu]-->
        <%
	  	'// Gera Menus ----------------------------- >
		'// 0-menu;1-voltar;2-imprimir;3-confirmar;4-incluir;
		'// 5-apagar;6-editar;7-reportar;8-mostrarall;
		'// 9-ocultarall;10-item.
		'//----------------------------------------- >
		redim AspMenuOp(2) 'Define o numero de itens
		'//campos ---------------------------------
		AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?aspid=' + AspId)"
		AspMenuOp(1) = "Salvar|3|miniMenu_fotoenq()"
		'//----------------------------------------
		AspMenu() 'gera menu
		'// Gera Menus ---------------------------- >
		%>
        <!--[menu]--></td>
    </tr>
    <tr>
      <td id="conteudo_page">
	  <!--[conteudo]-->
	  <table width="50%" border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td width="39"><span class="texto_size_10">Cod</span><br />
            <input name="id" type="text" id="id" value="<%=id%>" size="3" maxlength="3" readonly="true" /></td>
          <td width="457"><span class="texto_size_10">&Aacute;rea</span><br />
	            <select name="cod_area" id="cod_area" onchange="validperiodo(document.form.entrada, document.form.saida, this, document.form.id)">
			<%
				db "leitura",rs,sql,"select * from enquetefoto_areas where criado=1", conn
				while not rs.eof
				
					response.write "<option value='"&rs("id")&"'"
					if id_a = cint(rs("id")) then
						response.write " selected "
					end if
					response.write " >"&trim(rs("areas"))&"</option>"
				
					rs.movenext
				wend
			%>
            </select>
				  </td>
        </tr>
      </table>
	  <span class="texto_size_10">Pergunta</span><br />
	  <!--[conteudo]-->
      <input name="pergunta" type="text" id="pergunta" size="50" maxlength="255" value="<%=pergunta%>" /><input name="dtainvalid" type="hidden" value="true" />
      <br />
      <table width="50%" border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td width="123"><span class="texto_size_10">Entrada</span><br />
            <input name="entrada" type="text" id="entrada" value="<%=entrada%>" onBlur="if(this.value==''){alert('Você deve digitar uma data');}else{dataVal(this);};validperiodo(this, document.form.saida, document.form.cod_area, document.form.id)" onFocus="this.value=''" size="13" maxlength="10" /></td>
          <td><span class="texto_size_10">Sa&iacute;da</span><br />
            <input name="saida" onBlur="if(this.value==''){alert('Você deve digitar uma data');}else{dataVal(this); };validperiodo(document.form.entrada, this, document.form.cod_area,document.form.id)" onFocus="this.value=''" type="text" id="saida" value="<%=saida%>" size="13" maxlength="10" /></td>
        </tr>
        <tr>
          <td class="texto_size_10">Status </td>
          <td><input name="priodo" type="text" class="red" id="priodo" value="Peri&oacute;do Inv&aacute;lido" size="20" readonly="true" /></td>
        </tr>
      </table>
      <span class="texto_size_10">Numero Alternativas</span><br />
	  <% if alt = "" then %>
      <select name="alt" onchange="listaUpload(this)">
        <option value="2" <% if alt = "2" then %>selected="selected"<% end if %>>2</option>
        <option value="3" <% if alt = "3" then %>selected="selected"<% end if %>>3</option>
        <option value="4" <% if alt = "4" then %>selected="selected"<% end if %>>4</option>
        <option value="5" <% if alt = "5" then %>selected="selected"<% end if %>>5</option>
        <option value="6" <% if alt = "6" then %>selected="selected"<% end if %>>6</option>
        </select>
		<% else %>
			<span class="texto_size_10"><%=alt%><input name="alt" type="hidden" value="<%=alt%>" /></span>
		  <% end if %>	
      <br />
		<div id="divalt"></div>
		<%
			if novaenq <> true then
				'//dados gravados no banco ------
				response.write "<script>var resp = '"
				db "leitura",rs,sql,"select * from enquetefoto_respostas where id_p="&id,conn
				while not rs.eof
				
						response.write trim(rs("id"))&","&sites&foto_root&trim(rs("fotos"))&","&trim(rs("votos"))&","&trim(rs("titulo"))&";"
					
					rs.movenext
				wend
				response.write "';</script>"
				'// -----------------------------
			end if
			
			response.write "<script>listaUpload(document.form.alt)</script>"
			response.write "<script>validperiodo(document.form.entrada, document.form.saida, document.form.cod_area, document.form.id)</script>"
		%>
      <br />
	  <% if alt <> "" then %>
	  <input name="isupdate" type="hidden" value="1" />
	  <% end if %>
      <input name="totalvotos" type="checkbox" id="totalvotos" value="1" <%if totalvotos = "1" then%> checked="checked" <%end if%> />
      <span class="texto_size_10">Exibir o total de votos.</span><br />
      <input name="repeteuser" type="checkbox" id="repeteuser" value="1" <%if repeteuser = "1" then%> checked="checked" <%end if%>/>
      <span class="texto_size_10">Permitir que o usu&aacute;rio vote mais de uma vez.</span>
	  		<script>
	if(isIE()){
		oDateMasks = new Mask("dd/mm/yyyy", "date");
		oDateMasks.attach(document.form.entrada);
		oDateMasks.attach(document.form.saida);
		
	}
	
	</script><br />
	</td>
    </tr>
    <%

			
		case "list_enquete"
		' // Listar Enquete ----------------------------------------------- //
		prep_deletar()
	%>
<tr>
  <td id="menutab">
  	  <!--[menu]-->
	   <%
	  	'// Gera Menus ----------------------------- >
		'// 0-menu;1-voltar;2-imprimir;3-confirmar;4-incluir;
		'// 5-apagar;6-editar;7-reportar;8-mostrarall;
		'// 9-ocultarall;10-item.
		'//----------------------------------------- >
		redim AspMenuOp(2) 'Define o numero de itens
		'//campos ---------------------------------
		AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?aspid=' + AspId)"
		AspMenuOp(1) = "Novo|4|WindowOpen('default.asp?op=cad_enquete&aspid=' + AspId)"
		'AspMenuOp(1) = "Salvar|3|miniMenu_area()"
		'//----------------------------------------
		AspMenu() 'gera menu
		'// Gera Menus ---------------------------- >
		%>
      <!--[menu]-->
	  </td>
    </tr>
    <tr>
      <td id="conteudo_page"><span class="texto_size_10">
        <!--[conteudo]-->
        Selecione a &Aacute;rea</span><br>
	<form name="form1" method="post">
	  <select name="id_a" class="formtextarea" onChange="window.open('<%=request.ServerVariables("SCRIPT_NAME")%>?aspid=<%=aspid%>&op=list_enquete&id_a='+this.value,'_self');">
	  <option value="" selected>Selecione a Area</option>
	  <%
	  	db "leitura",rs,sql,"select * from enquetefoto_areas where criado=1",conn
		while not rs.eof
	  %>
	  	<option value="<%=rs("id")%>"><%=rs("areas")%></option>
	<%
			rs.movenext
		wend
	%>
	    </select>
	  </form>
	  <div align="center">
	    <%
				  reDim GridData(6), GridName(6)
				
				
					GridSQL = "select * from show_enquetefoto where id_a = "&formus("id_a",3)&";"
					
					GridName(0) = "Pergunta;250"
					GridName(1) = "Entrada;80"
					GridName(2) = "Saída;80"
					GridName(3) = "Votos;80"
					GridName(4) = "Ativado;60"
					GridName(5) = "Alterar;60"
					GridName(6) = "Deletar;60"
									
					GridData(0) = "pergunta;0;"&request.ServerVariables("SCRIPT_NAME")&"?op=result&aspid=[aspid]&id_a={id_a}&id={id}"
					GridData(1) = "entrada;0;0"
					GridData(2) = "saida;0;0"
					GridData(3) = "totalvotos;1;0"
					GridData(4) = "$ativo;1;"&request.ServerVariables("SCRIPT_NAME")&"?op=atv&aspid=[aspid]&id_a={id_a}&atv={ativo}&id={id};$case({ativo}=1):On-Line:Off-Line"
					GridData(5) = "*Alterar;1;"&request.ServerVariables("SCRIPT_NAME")&"?op=cad_enquete&aspid=[aspid]&id={id}"
					GridData(6) = "*Deletar;1;javascript: deletar('"&request.ServerVariables("SCRIPT_NAME")&"?op=del&aspid=[aspid]&tp=enq&id={id}&id_a={id_a}','[col0]')"
							
					GridDate = "dd/mm/yyyy"
					GridPoss = "center"
					'GridURLS = 1
					'// Chama a Função
					if request.QueryString("id_a")<>"" then
						GridForm Gkeys,15
					end if
	%>	
      </div>
          <!--[conteudo]-->
      </td>
    </tr>
    <%
		case "del"
			
			function fotoenquet(cod)
			
				db "leitura",rs,sql,"select * from enquetefoto_respostas where id_p="&cod,conn
				while not rs.eof 
					delfile(foto_root&trim(rs("fotos")))
					delfile(fotog_root&trim(rs("fotos")))
					rs.movenext
				wend
				db "normal",,sql,"delete from enquetefoto_respostas where id_p="&cod,conn
				
				db "normal",,sql,"delete from enquetefoto_perguntas where id="&cod,conn
			
				'fotoenquet = true
			
			end function
			
			select case request.QueryString("tp")
			case "enq"
				fotoenquet(formus("id",3))
				response.Write "<script>WindowOpen('"&request.ServerVariables("SCRIPT_NAME")&"?op=list_enquete&id_a="&formus("id_a",3)&"&aspid="&aspid&"')</script>"
			case "area"
			
				
				db "leitura",rs,sql,"select id from enquetefoto_perguntas where id_a="&formus("id",3),conn
				while not rs.eof
					fotoenquet(rs("id"))
					rs.movenext
				wend
				
				db "normal",,sql,"delete from enquetefoto_areas where id="&formus("id",3),conn
				response.Write "<script>WindowOpen('"&request.ServerVariables("SCRIPT_NAME")&"?op=list_area&aspid="&aspid&"')</script>"
			case else
				response.Write "<script>WindowOpen('"&request.ServerVariables("SCRIPT_NAME")&"?aspid="&aspid&"')</script>"
			end select
		
		case else
		'// Menu Principal ------------------------------------------------ //
		
	%>
	
	    <tr>
      <td id="menutab">
	  <!--[menu]-->
	  
	  <!--[menu]--></td>
    </tr>
    <tr>
      <td id="conteudo_page">
	  <!--[conteudo]-->
	  <p class="texto_size_10">Selecione qual o tipo de informa&ccedil;&atilde;o deseja visualizar:</p>
        <ul class="texto_size_10">
          <li><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=cad_area&aspid=<%=aspid%>" target="_self">Criar nova &Aacute;rea de Fotos</a></li>
          <li><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=list_area&aspid=<%=aspid%>" target="_self">Listar &Aacute;reas Criadas</a></li>
          <li><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=cad_enquete&aspid=<%=aspid%>" target="_self">Criar nova Enquete </a></li>
          <li><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=list_enquete&aspid=<%=aspid%>" target="_self">Listar Enquetes </a></li>
        </ul>
      <p>&nbsp;</p>
	  <!--[conteudo]--></td>
    </tr>
	
	<%
		end select
	%>
  </table>
  <br />
</form>
</body>
</html>
