<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<%
	cache 1,false
	protec(2)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Mala Direta</title>
<link href="../css/estilos.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" type="text/javascript" src="../js/script.js"></script>
<script language="JavaScript" type="text/javascript" src="../js/form_mala.js"></script>
<script language="JavaScript" type="text/javascript" src="../js/form.js"></script>
<script>AspId = '<%=aspid%>'</script>

<%
	
		if request("op")="cad_msg" and lcase(request.ServerVariables("REQUEST_METHOD")) <> "post" then
		
	if formus("id",3)<>"" then
		id = formus("id",3)
	else
		db "leitura",rs,sql,"select cod_msg from mala_msg where criado=0",conn
		if rs.eof and rs.bof then
			db "normal",,sql,"insert into mala_msg(criado) values (0)",conn
			db "leitura",rs,sql,"select cod_msg from mala_msg where criado=0",conn
		end if
		id = rs("cod_msg")
	end if
	
%>
		<script type="text/javascript">
		  _editor_url = "../editor/";
		  _editor_lang = "en";
		</script>
		<script type="text/javascript" src="../editor/htmlarea.asp?tela=MalaDireta&id=<%=id%>"></script>
		
		<style type="text/css">
			textarea { background-color: #fff; border: 1px solid 00f; }
		</style>
		
		<script type="text/javascript">
		<!--
		var editor = null;
		function initEditor() {
		  editor = new HTMLArea("ta");
		
		  // comment the following two lines to see how customization works
		  editor.generate();
		  return false;
		
		  var cfg = editor.config; // this is the default configuration
		  cfg.registerButton({
			id        : "my-hilite",
			tooltip   : "Highlight text",
			image     : "ed_custom.gif",
			textMode  : false,
			action    : function(editor) {
						  editor.surroundHTML("<span class=\"hilite\">", "</span>");
						},
			context   : 'table'
		  });
		
		  cfg.toolbar.push(["linebreak", "my-hilite"]); // add the new button to the toolbar
		
		  // BEGIN: code that adds a custom button
		  // uncomment it to test
		  var cfg = editor.config; // this is the default configuration
		
		
		function clickHandler(editor, buttonId) {
		  switch (buttonId) {
			case "my-toc":
			  editor.insertHTML("<h1>Table Of Contents</h1>");
			  break;
			case "my-date":
			  editor.insertHTML((new Date()).toString());
			  break;
			case "my-bold":
			  editor.execCommand("bold");
			  editor.execCommand("italic");
			  break;
			case "my-hilite":
			  editor.surroundHTML("<span class=\"hilite\">", "</span>");
			  break;
		  }
		};
		cfg.registerButton("my-toc",  "Insert TOC", "ed_custom.gif", false, clickHandler);
		cfg.registerButton("my-date", "Insert date/time", "ed_custom.gif", false, clickHandler);
		cfg.registerButton("my-bold", "Toggle bold/italic", "ed_custom.gif", false, clickHandler);
		cfg.registerButton("my-hilite", "Hilite selection", "ed_custom.gif", false, clickHandler);
		
		cfg.registerButton("my-sample", "Class: sample", "ed_custom.gif", false,
		  function(editor) {
			if (HTMLArea.is_ie) {
			  editor.insertHTML("<span class=\"sample\">&nbsp;&nbsp;</span>");
			  var r = editor._doc.selection.createRange();
			  r.move("character", -2);
			  r.moveEnd("character", 2);
			  r.select();
			} else { // Gecko/W3C compliant
			  var n = editor._doc.createElement("span");
			  n.className = "sample";
			  editor.insertNodeAtSelection(n);
			  var sel = editor._iframe.contentWindow.getSelection();
			  sel.removeAllRanges();
			  var r = editor._doc.createRange();
			  r.setStart(n, 0);
			  r.setEnd(n, 0);
			  sel.addRange(r);
			}
		  }
		);
		
		
		  /*
		  cfg.registerButton("my-hilite", "Highlight text", "ed_custom.gif", false,
			function(editor) {
			  editor.surroundHTML('<span class="hilite">', '</span>');
			}
		  );
		  */
		  cfg.pageStyle = "body { background-color: #efd; } .hilite { background-color: yellow; } "+
						  ".sample { color: green; font-family: monospace; }";
		  cfg.toolbar.push(["linebreak", "my-toc", "my-date", "my-bold", "my-hilite", "my-sample"]); // add the new button to the toolbar
		  // END: code that adds a custom button
		
		  editor.generate();
		}
		function insertHTML() {
		  var html = prompt("Enter some HTML code here");
		  if (html) {
			editor.insertHTML(html);
		  }
		}
		function highlight() {
		  editor.surroundHTML('<span style="background-color: yellow">', '</span>');
		}
		
		
		//-->
		</script>
<% end if %>
</head>

<body <% if request("op")="cad_msg" and lcase(request.ServerVariables("REQUEST_METHOD")) <> "post" then 	 %> onload="initEditor();" <% end if %>>

<form method="post" name="form" id="form">
  <table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td id="titulo_pagina"><div id="matrix_titulo" style="float:left">Mala Direta</div></td>
    </tr>
	<%
		dim rs,sql,id, criado, x
		select case request("op")
		case "cad_area"
		
		dim cod_area, nome_area
		'//Cadastrar Áreas -----------------------------------------------------
		if lcase(request.ServerVariables("REQUEST_METHOD")) <> "post" then
		
			cod_area = formus("id",3)
			
			if cod_area<>"" then
				subTitle("> Alterar Área")
				'//Alterar Dados
				db "leitura",rs,sql,"select top 1 * from mala_area where cod_area="&cod_area,conn
				nome_area = trim(rs("nome_area"))
				
			else
				subTitle("> Nova Área")
				'//nova area
				db "leitura",rs,sql,"select top 1 * from mala_area where criado=0",conn
				if rs.eof and rs.bof then
					'//não existe registro
					db "normal",,sql,"insert into mala_area(criado) values(0)",conn
					db "leitura",rs,sql,"select top 1 * from mala_area where criado=0",conn
				end if						
				
			end if
			
			cod_area = rs("cod_area")
			criado = rs("criado")
		
	%>
	<tr>
      <td id="menutab">
	  	<%
	  	'// Gera Menus ----------------------------- >
		'// 0-menu;1-voltar;2-imprimir;3-confirmar;4-incluir;
		'// 5-apagar;6-editar;7-reportar;8-mostrarall;
		'// 9-ocultarall;10-item.
		'//----------------------------------------- >
		dim AspMenuOp(2) 'Define o numero de itens
		'//campos ---------------------------------
		AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?op=list_area&aspid=' + AspId)"
		AspMenuOp(1) = "Salvar|3|miniMenu_area()"
		'//----------------------------------------
		AspMenu() 'gera menu
		'// Gera Menus ---------------------------- >
		%>
	  </td>
	</tr>
	 <tr>
      <td id="conteudo_page"><span class="texto_size_10">Cod</span><br />
        <label>
        <input name="cod_area" type="text" id="cod_area" value="<%=cod_area%>" size="3" readonly="true" />
        <br />
        <span class="texto_size_10">&Aacute;rea</span><br />
<input name="nome_area" type="text" id="nome_area" value="<%=nome_area%>" size="40" maxlength="255" />
       </label></td>
	 </tr>
	 <%
	 	else
			'//salvar l area
			db "leitura", rs, sql, "select nome_area from mala_area where cod_area<>"&formus("cod_area",0)&" and nome_area='"&formus("nome_area",0)&"'",conn
			if rs.eof and rs.bof then
				db "normal",,sql,"update mala_area set nome_area='"&formus("nome_area",0)&"', criado=1 where cod_area="&formus("cod_area",0),conn
				alert("Salvo com sucesso!")
				voltar request.ServerVariables("SCRIPT_NAME")&"?op=cad_area&aspid="&aspid,"","_self"
			else
				alert("Está area já está cadastrada!")
				voltar request.ServerVariables("SCRIPT_NAME")&"?op=cad_area&aspid="&aspid&"&id="&formus("cod_area",0),"","_self"
			end if
			
		end if
	 	case "list_area"
	 %>
	 <tr>
	   <td id="menutab">
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
		'//----------------------------------------
		AspMenu() 'gera menu
		'// Gera Menus ---------------------------- >
		prep_deletar()
		%>
	   </td>
    </tr>
	 <tr>
	   <td id="conteudo_page">
	   <div align="center">
	    <%
		  subTitle("> Listar Área")
		  Dim GridData(3), GridName(3), GridSQL, GridDate, GridPoss
				
				
					GridSQL = "select * from mala_area where criado=1;"
					
					GridName(0) = "Cod;35"
					GridName(1) = "Área;370"
					GridName(2) = "Alterar;60"
					GridName(3) = "Deletar;60"
									
					GridData(0) = "cod_area;1;0"
					GridData(1) = "nome_area;0;0"
					GridData(2) = "*Alterar;1;default.asp?op=cad_area&id=[col0]&aspid=[aspid]&m=[mod]"
					GridData(3) = "*Deletar;1;javascript: deletar('"&request.ServerVariables("SCRIPT_NAME")&"?op=del&tp=area&aspid=[aspid]&id=[col0]&m=[mod]','[col1]')"
							
					GridDate = "dd/mm/yyyy"
					GridPoss = "center"
					'GridURLS = 1
					'// Chama a Função
							
					GridForm Gkeys,15
					
	%>
	   </div>
	   </td>
    </tr>

	
	<%
		case "cad_grupo"
		
		dim cod_grupo, nome_grupo
		'//Cadastrar Grupos -----------------------------------------------------
		if lcase(request.ServerVariables("REQUEST_METHOD")) <> "post" then
		
			cod_grupo = formus("id",3)
			
			if cod_grupo<>"" then
				subTitle("> Alterar Grupos")
				'//Alterar Dados
				db "leitura",rs,sql,"select top 1 * from mala_grupos where cod_grupo="&cod_grupo,conn
				nome_grupo = trim(rs("nome_grupo"))
				
			else
				subTitle("> Novo Grupo")
				'//nova area
				db "leitura",rs,sql,"select top 1 * from mala_grupos where criado=0",conn
				if rs.eof and rs.bof then
					'//não existe registro
					db "normal",,sql,"insert into mala_grupos(criado) values(0)",conn
					db "leitura",rs,sql,"select top 1 * from mala_grupos where criado=0",conn
				end if						
				
			end if
			
			cod_grupo = rs("cod_grupo")
			criado = rs("criado")
			
	%>
	 <tr>
	   <td id="menutab">
	   	<%
	  	'// Gera Menus ----------------------------- >
		'// 0-menu;1-voltar;2-imprimir;3-confirmar;4-incluir;
		'// 5-apagar;6-editar;7-reportar;8-mostrarall;
		'// 9-ocultarall;10-item.
		'//----------------------------------------- >
		redim AspMenuOp(2) 'Define o numero de itens
		'//campos ---------------------------------
		AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?op=list_grupo&aspid=' + AspId)"
		AspMenuOp(1) = "Salvar|3|miniMenu_grupo()"
		'//----------------------------------------
		AspMenu() 'gera menu
		'// Gera Menus ---------------------------- >
		%>
	   </td>
    </tr>
	 <tr>
	   <td id="conteudo_page"><span class="texto_size_10">Cod</span><br />
        <label>
        <input name="cod_grupo" type="text" id="cod_grupo" value="<%=cod_grupo%>" size="3" readonly="true" />
        <br />
        <span class="texto_size_10">Grupos</span><br />
<input name="nome_grupo" type="text" id="nome_grupo" value="<%=nome_grupo%>" size="40" maxlength="255" />
       </label></td>
    </tr>
	<%
	
		else
			'//salvar l area
			db "leitura", rs, sql, "select nome_grupo from mala_grupos where cod_grupo<>"&formus("cod_grupo",0)&" and nome_grupo='"&formus("nome_grupo",0)&"'",conn
			if rs.eof and rs.bof then
				db "normal",,sql,"update mala_grupos set nome_grupo='"&formus("nome_grupo",0)&"', criado=1 where cod_grupo="&formus("cod_grupo",0),conn
				alert("Salvo com sucesso!")
				voltar request.ServerVariables("SCRIPT_NAME")&"?op=cad_grupo&aspid="&aspid,"","_self"
			else
				alert("Este Grupo a já está cadastrado!")
				voltar request.ServerVariables("SCRIPT_NAME")&"?op=cad_grupo&aspid="&aspid&"&id="&formus("cod_grupo",0),"","_self"
			end if
			
		end if
		
		case "list_grupo"
	%>
	 <tr>
	   <td id="menutab">
	   	<%
	  	'// Gera Menus ----------------------------- >
		'// 0-menu;1-voltar;2-imprimir;3-confirmar;4-incluir;
		'// 5-apagar;6-editar;7-reportar;8-mostrarall;
		'// 9-ocultarall;10-item.
		'//----------------------------------------- >
		redim AspMenuOp(2) 'Define o numero de itens
		'//campos ---------------------------------
		AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?aspid=' + AspId)"
		AspMenuOp(1) = "Novo|4|WindowOpen('default.asp?op=cad_grupo&aspid=' + AspId)"
		'//----------------------------------------
		AspMenu() 'gera menu
		'// Gera Menus ---------------------------- >
		prep_deletar()
		%>
	   </td>
    </tr>
	 <tr>
	   <td id="conteudo_page">

	     <div align="center">
	    <%
		  subTitle("> Listar Grupos")
		  ReDim GridData(3), GridName(3)
				
				
					GridSQL = "select * from mala_grupos where criado=1;"
					
					GridName(0) = "Cod;35"
					GridName(1) = "Grupo;370"
					GridName(2) = "Alterar;60"
					GridName(3) = "Deletar;60"
									
					GridData(0) = "cod_grupo;1;0"
					GridData(1) = "nome_grupo;0;0"
					GridData(2) = "*Alterar;1;default.asp?op=cad_grupo&id=[col0]&aspid=[aspid]&m=[mod]"
					GridData(3) = "*Deletar;1;javascript: deletar('"&request.ServerVariables("SCRIPT_NAME")&"?op=del&tp=grupo&aspid=[aspid]&id=[col0]&m=[mod]','[col1]')"
							
					GridDate = "dd/mm/yyyy"
					GridPoss = "center"
					'GridURLS = 1
					'// Chama a Função
							
					GridForm Gkeys,15
					
	%>
	   </div>	   </td>
    </tr>
	<%
		case "cad_email"
		
		dim cod_email, nome_email, sexo_email, email_email
		
		'//Cadastrar E-Mails -----------------------------------------------------
		if lcase(request.ServerVariables("REQUEST_METHOD")) <> "post" then
		
			cod_email = formus("id",3)
			
			if cod_email<>"" then
				
				'//Alterar Dados
				db "leitura",rs,sql,"select top 1 * from mala_email where cod_email="&cod_email,conn

				nome_email = trim(rs("nome_email"))
				sexo_email = rs("sexo_email")
				email_email = trim(rs("email_email"))

				if email_email = "root@matrix" then
					email_email = "" 
					subTitle("> Novo E-Mail")
				else
					subTitle("> Alterar E-Mail")
				end if
				
				cod_grupo = rs("cod_grupo")
				
			else
				subTitle("> Novo E-Mail")
				'//nova area
				db "leitura",rs,sql,"select top 1 * from mala_email where criado=0",conn
				if rs.eof and rs.bof then
					'//não existe registro
					db "normal",,sql,"insert into mala_email(criado) values(0)",conn
					db "leitura",rs,sql,"select top 1 * from mala_email where criado=0",conn
				end if
				
				'// nova logica reservando id (beta)
				'-------------------------------------------					
				if not rs.eof and not rs.bof then
					db "leitura",rs,sql,"select top 1 * from mala_email where criado=0",conn
					db "normal",,sql,"update mala_email set criado=2 where cod_email="&rs("cod_email"),conn
				end if
				'//status 2 - reservado
				
			end if
			
			cod_email = rs("cod_email")
			criado = rs("criado")
			
	%>
	 <tr>
	   <td id="menutab">
	   	<%
	  	'// Gera Menus ----------------------------- >
		'// 0-menu;1-voltar;2-imprimir;3-confirmar;4-incluir;
		'// 5-apagar;6-editar;7-reportar;8-mostrarall;
		'// 9-ocultarall;10-item.
		'//----------------------------------------- >
		redim AspMenuOp(2) 'Define o numero de itens
		'//campos ---------------------------------
		AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?op=list_email&codemail="&cod_email&"&aspid=' + AspId)"
		AspMenuOp(1) = "Salvar|3|miniMenu_email()"
		'//----------------------------------------
		AspMenu() 'gera menu
		'// Gera Menus ---------------------------- >
		%>
	   </td>
    </tr>
	 <tr>
	   <td id="conteudo_page"><table width="100%" border="0" cellspacing="0" cellpadding="0">
         <tr>
           <td width="8%"><span class="texto_size_10">Cod</span><br />
            <input name="cod_email" type="text" id="cod_email" value="<%=cod_email%>" size="3" readonly /></td>
           <td width="92%"><span class="texto_size_10">Grupo</span><br />
             <select name="cod_grupo" id="cod_grupo">
			 <%
			 	db "leitura",rs,sql,"select * from mala_grupos where criado=1 order by nome_grupo", conn
				while not rs.eof
			 %>
               <option value="<%=rs("cod_grupo")%>" <% if cstr(rs("cod_grupo")) = trim(cod_grupo) then %>selected<% end if %>><%=trim(rs("nome_grupo"))%></option>
			  <%
			  		rs.movenext
				wend
			  %>
             </select>
</td>
         </tr>
       </table>
         <span class="texto_size_10">Nome</span><br />
       <input name="nome_email" type="text" id="nome_email" value="<%=nome_email%>" size="50" maxlength="255" />
       <br />
       <span class="texto_size_10">E-Mail</span><br />
       <input name="email_email" type="text" id="email_email" onBlur="if(!isEmail(this.value)){this.value=''}" value="<%=email_email%>" size="40" maxlength="255" />
       <br />
       <span class="texto_size_10">Sexo</span><br />
       <select name="sexo_email" id="sexo_email">
         <option value="1" <% if trim(sexo_email) = "1" then %>selected<% end if %>>Masculino</option>
         <option value="0" <% if trim(sexo_email) = "0" then %>selected<% end if %>>Feminino</option>
       </select>
</td>
    </tr>
	<%
		else
			'//salvar email
			db "leitura",rs,sql,"select email_email from mala_email where email_email='"&formus("email_email",0)&"' and cod_email<>"&formus("cod_email",0),conn
			if rs.eof and rs.bof then
				db "normal",,sql,"update mala_email set cod_grupo="&formus("cod_grupo",0)&",nome_email='"&formus("nome_email",0)&"',email_email='"&formus("email_email",0)&"',sexo_email="&formus("sexo_email",0)&", criado=1 where cod_email="&formus("cod_email",0),conn
				alert("Salvo com sucesso!")
				voltar request.ServerVariables("SCRIPT_NAME")&"?op=cad_email&aspid="&aspid,"","_self"
			else
				alert("Este E-Mail a já está cadastrado!")
				
				
				'//retorna ao valor incial
				db "leitura",rs,sql,"select criado from mala_email where cod_email="&formus("cod_email",0),conn
				if cint(rs("criado")) = 2 then
					db "normal",,sql,"update mala_email set criado=0 where cod_email="&formus("cod_email",0),conn
				end if
				
				voltar request.ServerVariables("SCRIPT_NAME")&"?op=cad_email&aspid="&aspid&"&id="&formus("cod_email",0),"","_self"
			end if
			
		end if
		case "list_email"
		
		'//retorna valores iniciais
		'//retorna ao valor incial
		if request("codemail") <> "" then
		'lista de e-mail
			db "leitura",rs,sql,"select criado from mala_email where cod_email="&formus("codemail",3),conn
			if cint(rs("criado")) = 2 then
				db "normal",,sql,"update mala_email set criado=0 where cod_email="&formus("codemail",3),conn
			end if
		end if

	%>
	 <tr>
	   <td id="menutab">
	   	<%
	  	'// Gera Menus ----------------------------- >
		'// 0-menu;1-voltar;2-imprimir;3-confirmar;4-incluir;
		'// 5-apagar;6-editar;7-reportar;8-mostrarall;
		'// 9-ocultarall;10-item.
		'//----------------------------------------- >
		redim AspMenuOp(4) 'Define o numero de itens
		'//campos ---------------------------------
		AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?aspid=' + AspId)"
		AspMenuOp(1) = "Novo|4|WindowOpen('default.asp?op=cad_email&aspid=' + AspId)"
		AspMenuOp(2) = "Alt.Stat|6|alt_statmala('"&request("idgrupo")&"')"
		AspMenuOp(3) = "Excluir|5|excluir_emailmala('"&request("idgrupo")&"')"
		'//----------------------------------------
		AspMenu() 'gera menu
		'// Gera Menus ---------------------------- >
		prep_deletar()
		%>
	   </td>
    </tr>
	 <tr>
	   <td id="conteudo_page">
	   <input name="sCod_vale" type="hidden" value="" />
	   <table width="620" border="0" cellspacing="0" cellpadding="0">
         <tr>
           <td><span class="texto_size_10">Selecione o Grupo</span></td>
           <td><span class="texto_size_10">Filtros</span></td>
           <td>&nbsp;</td>
           <td><span class="texto_size_10">Buscar</span></td>
           <td>&nbsp;</td>
           <td>&nbsp;</td>
         </tr>
         <tr>
           <td><select name="Grupos" id="Grupos">
             <option selected="selected">Todos os grupos</option>
             <%
			 	db "leitura",rs,sql,"select * from mala_grupos where criado=1 order by nome_grupo", conn
				while not rs.eof
			 %>
             <option value="<%=rs("cod_grupo")%>" <% if cstr(rs("cod_grupo")) = trim(request("idgrupo")) then %>selected<% end if %>><%=trim(rs("nome_grupo"))%></option>
             <%
			  		rs.movenext
				wend
			  %>
           </select></td>
           <td><select name="Sexo" id="Sexo">
             <option selected="selected">Ambos os Sexos</option>
             <option value="1">Masculino</option>
             <option value="0">Feminino</option>
           </select></td>
           <td><select name="Status" id="Status">
             <option>Ambos Status</option>
             <option value="1">Ativos</option>
             <option value="0">Inativos</option>
           </select></td>
           <td><input name="Querys" type="text" id="Querys" /></td>
           <td><select name="Por" id="Por">
             <option selected="selected">E-Mail e Nome</option>
             <option value="1">Nome</option>
             <option value="0">E-Mail</option>
           </select></td>
           <td><input type="button" name="Button2" value="Filtrar" onclick="QueryEmail()" /></td>
         </tr>
       </table>
	   <br />
	   <br />
	   	     <div align="center">
	    <%
		  subTitle("> Listar Email")
		  ReDim GridData(5), GridName(5)
				
					sql = "select * from mala_email where criado=1"
					
					if request("idgrupo") <> "" then
						sql = sql & " and cod_grupo = "&request("idgrupo")
					end if
					
					if request("idsexo") <> "" then
						sql = sql & " and sexo_email = "&request("idsexo")
					end if
					
					if request("idstatus") <> "" then
						sql = sql & " and atv_email = "&request("idstatus")
					end if
					
					if request("query") <> "" then
						if request("tpquery") = "1" then
							sql = sql & " and nome_email like '"&request("query")&"%'"
						elseif request("tpquery") = "0" then
							sql = sql & " and email_email like '"&request("query")&"%'"
						else
							sql = sql & " and nome_email like '"&request("query")&"%' or email_email like '"&request("query")&"%'"
						end if
					end if


					
					sql = sql & " order by nome_email"
					
					GridSQL = sql
					
					GridName(0) = "-;35"
					GridName(1) = "Cod;35"
					GridName(2) = "Nome;200"
					GridName(3) = "Email;150"
					GridName(4) = "Status;80"
					GridName(5) = "Editar;60"
									
					GridData(0) = "$cod_email;1;0;$case({cod_email}!0):<input type='checkbox' name='codEmail' value='{cod_email}' onclick='nList(this.value,document.form.sCod_vale)'>"
					GridData(1) = "cod_email;1;0"
					GridData(2) = "nome_email;0;0"
					GridData(3) = "email_email;0;0"
					GridData(4) = "$atv_email;1;0;$case({atv_email}!0):Ativo:Inativo"
					GridData(5) = "*Editar;1;default.asp?op=cad_email&id={cod_email}&aspid=[aspid]"
							
					GridDate = "dd/mm/yyyy"
					GridPoss = "center"
					'GridURLS = 1
					'// Chama a Função
					
					GridForm Gkeys,15
					
	%>
	   </div>
	   </td>
    </tr>
	<%
		case "cad_list_email"
		
		'//Cadastrar E-Mails lista -----------------------------------------------------
		if lcase(request.ServerVariables("REQUEST_METHOD")) <> "post" then
	%>
	<tr>
      <td id="menutab">
	  	<%
	  	'// Gera Menus ----------------------------- >
		'// 0-menu;1-voltar;2-imprimir;3-confirmar;4-incluir;
		'// 5-apagar;6-editar;7-reportar;8-mostrarall;
		'// 9-ocultarall;10-item.
		'//----------------------------------------- >
		redim AspMenuOp(2) 'Define o numero de itens
		'//campos ---------------------------------
		AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?aspid=' + AspId)"
		AspMenuOp(1) = "Salvar|3|miniMenu_listemail()"
		'//----------------------------------------
		AspMenu() 'gera menu
		'// Gera Menus ---------------------------- >
		%>
	  </td>
	</tr>
	 <tr>
      <td id="conteudo_page"><p><span class="texto_size_10">Grupo</span><br />
        <select name="cod_grupo" id="cod_grupo">
		<%
			 	db "leitura",rs,sql,"select * from mala_grupos where criado=1 order by nome_grupo", conn
				while not rs.eof
		%>
               <option value="<%=rs("cod_grupo")%>" <% if cstr(rs("cod_grupo")) = trim(request("idgrupo")) then %>selected<% end if %>><%=trim(rs("nome_grupo"))%></option>
		 <%
			  		rs.movenext
				wend
		 %>
        </select>
        <br />
          <span class="texto_size_10">Lista de E-Mails</span><br />
          <span class="texto_size_10">Insira os e-mails separados por ponto e virgula( ; )</span><br />
          <textarea name="list_mail" cols="60" rows="9" id="list_mail"></textarea>
          <input type="button" name="Button" value="Filtrar E-Mails" onclick="filtrarmail(document.form.list_mail)" />
      </p>
       </td>
	 </tr>
	<%
		else
			'//salvar
			'Verifica se existe algum criado = 0
			db "leitura",rs,sql,"select criado from mala_email where criado=0",conn
			if rs.eof and rs.bof then
				db "normal",,sql,"insert into mala_email(criado) values(0)",conn
			end if
			
			'Cadastra os e-mails 
			dim listemail, nomemail
			
			listemail = split(formus("list_mail",0),";")
			
			for i = lbound(listemail) to ubound(listemail)
				nomemail = empty
				nomemail = split(listemail(i),"@")
				'se o email existir ele ira ignorar o erro

				db "leitura",rs,sql,"select email_email from mala_email where email_email='"&trim(listemail(i))&"'",conn
				if rs.eof and rs.bof then
					db "normal",,sql,"update mala_email set nome_email='"&trim(nomemail(0))&"',cod_grupo="&formus("cod_grupo",0)&",email_email='"&trim(listemail(i))&"', criado=1 where criado=0",conn
				end if
				
			next
			alert("E-Mails Cadastrados")
			voltar request.ServerVariables("SCRIPT_NAME")&"?op=cad_list_email&aspid="&aspid,"","_self"
			
		end if
		case "cad_msg"
		
		dim cod_msg, nome_msg, titulo_msg, html_msg, modelo_msg, de_msg
		'//Cadastrar Modelo -----------------------------------------------------
		if lcase(request.ServerVariables("REQUEST_METHOD")) <> "post" then
		
			cod_msg = formus("id",3)
			modelo_msg = formus("modelo",3)
			
			
			if cod_msg<>"" then
			
				
				'//Alterar Dados
				db "leitura",rs,sql,"select top 1 * from mala_msg where cod_msg="&cod_msg,conn

				nome_msg = trim(rs("nome_msg"))
				html_msg = trim(rs("html_msg"))
				titulo_msg = trim(rs("titulo_msg"))
				de_msg = trim(rs("de_msg"))
				modelo_msg = rs("modelo_msg")
				
				if cstr(modelo_msg) = "1" then
					subTitle("> Editar Modelo")
				else
					subTitle("> Editar E-Mail")
				end if
				
			else
			
				'//nova area
				db "leitura",rs,sql,"select top 1 * from mala_msg where criado=0",conn
				if rs.eof and rs.bof then
					'//não existe registro
					db "normal",,sql,"insert into mala_msg(criado) values(0)",conn
					db "leitura",rs,sql,"select top 1 * from mala_msg where criado=0",conn
				end if						
				
				if cstr(modelo_msg) = "1" then
					subTitle("> Novo Modelo")
				else
					subTitle("> Novo E-Mail")
				end if
				
			end if
			
			cod_msg = rs("cod_msg")
			criado = rs("criado")
			
			'// Abre o modelo
			if request("modelocod")<>""  then
				db "leitura",rs,sql,"select html_msg, titulo_msg from mala_msg where cod_msg="&request("modelocod"),conn
				if not rs.eof and not rs.bof then
					titulo_msg = trim(rs("titulo_msg"))
					html_msg = trim(rs("html_msg"))
				end if
			end if
	%>
	 <tr>
	   <td id="menutab">
	   	  	<%
	  	'// Gera Menus ----------------------------- >
		'// 0-menu;1-voltar;2-imprimir;3-confirmar;4-incluir;
		'// 5-apagar;6-editar;7-reportar;8-mostrarall;
		'// 9-ocultarall;10-item.
		'//----------------------------------------- >
		redim AspMenuOp(2) 'Define o numero de itens
		'//campos ---------------------------------
		AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?aspid=' + AspId)"
		AspMenuOp(1) = "Salvar|3|miniMenu_msg()"
		'//----------------------------------------
		AspMenu() 'gera menu
		'// Gera Menus ---------------------------- >
		%>
	   </td>
    </tr>
	 <tr>
	   <td id="conteudo_page"><table width="100%" border="0" cellspacing="0" cellpadding="0">
           <tr>
             <td width="17%"><span class="texto_size_10">Cod</span><br />
             <input name="cod_msg" type="text" id="cod_msg" value="<%=cod_msg%>" size="3" readonly="true" /></td>
             <td width="83%">
			 <% if cstr(modelo_msg) <> "1" then %>
			 <span class="texto_size_10">&Aacute;rea</span><br />
               <select name="cod_area" id="cod_area">
			   <%
			   		db "leitura",rs,sql,"select * from mala_area where criado=1 order by nome_area",conn
					while not rs.eof
			   %>
			   	<option value="<%=rs("cod_area")%>" <% if cod_area = trim(rs("cod_area")) then %>selected="selected"<% end if %>><%=trim(rs("nome_area"))%></option>
			   <%
						rs.movenext
					wend
			   %>
               </select>
			   <% else %>
			   	&nbsp;
			   <% end if %></td>
           </tr>
         </table>
       <span class="texto_size_10">Nome Mensagem</span><br />
       <input name="nome_msg" type="text" id="nome_msg" value="<%=nome_msg%>" size="40" maxlength="255" />
       <br />
       <span class="texto_size_10">Assunto da Mensagem</span><br />
       <input name="titulo_msg" type="text" id="titulo_msg" value="<%=titulo_msg%>" size="40" maxlength="255" />
       <br />
	   <% 
	   	if cstr(modelo_msg) <> "1" then
		%>
		<span class="texto_size_10">De</span><br />
		<input name="de_msg" type="text" id="de_msg" onBlur="if(!isEmail(this.value)){this.value=''}" value="<%=de_msg%>" size="30" maxlength="255" />
<br />
       <span class="texto_size_10">Modelo</span><br />
	   
       <select name="modelocod" id="modelocod">
         <option value="0" selected="selected">N&atilde;o usar modelo</option>
		 <%
		 	db "leitura",rs,sql,"select nome_msg, cod_msg from mala_msg where modelo_msg=1",conn
			while not rs.eof
		 %>
	         <option value="<%=rs("cod_msg")%>" <% if request("modelocod") = trim(rs("cod_msg")) then %>selected="selected"<% end if %>><%=trim(rs("nome_msg"))%></option>
		<%
				rs.movenext
			wend
		%>
       </select>
       &nbsp;<input name="modelo_msg" type="hidden" value="0" />
       <input type="button" name="Button" value="Abrir Modelo" onclick="window.open('<%=request.ServerVariables("SCRIPT_NAME")%>?<%=request.QueryString%>&modelocod='+document.getElementById('modelocod').value,'_self')" />
       <br />
	   <%
	   		else
		%>
			<input name="modelo_msg" type="hidden" value="1" />
		<% end if %>
       <span class="texto_size_10">Mensagem</span><br />
	   <textarea id="ta" name="ta" style="width:99%" rows="25" cols="80"><%=html_msg%></textarea></td>
    </tr>
	<%
		else
			'//salvar modelo
			
			if formus("modelo_msg",0) = 1 then
			
				db "normal",,sql,"update mala_msg set nome_msg='"&formus("nome_msg",0)&"',enviado_msg=0, titulo_msg='"&formus("titulo_msg",0)&"', html_msg='"&formus("ta",0)&"', criado=1, modelo_msg="&formus("modelo_msg",0)&" where cod_msg="&formus("cod_msg",0),conn
			
			
				'modelo
				alert("Modelo Salvo com sucesso!")
				voltar request.ServerVariables("SCRIPT_NAME")&"?op=cad_msg&modelo=1&aspid="&aspid,"","_self"

			else
				db "normal",,sql,"update mala_msg set de_msg='"&formus("de_msg",0)&"', enviado_msg=0, cod_area="&formus("cod_area",0)&", nome_msg='"&formus("nome_msg",0)&"',titulo_msg='"&formus("titulo_msg",0)&"', html_msg='"&formus("ta",0)&"', criado=1, modelo_msg="&formus("modelo_msg",0)&" where cod_msg="&formus("cod_msg",0),conn
				'email
				alert("Mensagem Salva com sucesso!")
				voltar request.ServerVariables("SCRIPT_NAME")&"?op=cad_msg&aspid="&aspid,"","_self"

			end if
			
		end if
		case "list_msg"
	%>
	 <tr>
	   <td id="menutab">
	   	<%
		if request("modelo")="1" then
			 subTitle("> Listar Modelos")
			'// Gera Menus ----------------------------- >
			'// 0-menu;1-voltar;2-imprimir;3-confirmar;4-incluir;
			'// 5-apagar;6-editar;7-reportar;8-mostrarall;
			'// 9-ocultarall;10-item.
			'//----------------------------------------- >
			redim AspMenuOp(3) 'Define o numero de itens
			'//campos ---------------------------------
			AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?aspid=' + AspId)"
			AspMenuOp(1) = "Novo|4|WindowOpen('default.asp?op=cad_msg&modelo=1&aspid=' + AspId)"
			AspMenuOp(2) = "Excluir|5|excluir_msg('"&request("idarea")&"','"&request("enviado")&"','"&request("modelo")&"')"
			'//----------------------------------------
		
		else
			if request("enviado") <> "1" then
				 subTitle("> Caixa de Saida")
			else
				 subTitle("> E-Mails Enviados")
			end if 
			
			'// Gera Menus ----------------------------- >
			'// 0-menu;1-voltar;2-imprimir;3-confirmar;4-incluir;
			'// 5-apagar;6-editar;7-reportar;8-mostrarall;
			'// 9-ocultarall;10-item.
			'//----------------------------------------- >
			redim AspMenuOp(4) 'Define o numero de itens
			'//campos ---------------------------------
			AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?aspid=' + AspId)"
			AspMenuOp(1) = "Novo|4|WindowOpen('default.asp?op=cad_msg&aspid=' + AspId)"
			AspMenuOp(2) = "Enviar|3|enviar_msg()"
			AspMenuOp(3) = "Excluir|5|excluir_msg('"&request("idarea")&"','"&request("enviado")&"','"&request("modelo")&"')"
			'//----------------------------------------
		end if
		
		AspMenu() 'gera menu
		'// Gera Menus ---------------------------- >
		
		prep_deletar()
		
		%>
	   </td>
    </tr>
	 <tr>
	   <td id="conteudo_page">
	      <input name="sCod_vale" type="hidden" value="" />
		  <%
		  	if request("modelo")<>"1" then
		  %>
	   	     <span class="texto_size_10">Selecione a Área</span><br />
	     <select name="cod_area" id="cod_grupo">
	       <option selected="selected">Todas as áreas</option>
			 <%
			 	db "leitura",rs,sql,"select * from mala_area where criado=1 order by nome_area", conn
				while not rs.eof
			 %>
               <option value="<%=rs("cod_area")%>" <% if cstr(rs("cod_area")) = trim(request("idarea")) then %>selected<% end if %>><%=trim(rs("nome_area"))%></option>
			  <%
			  		rs.movenext
				wend
			  %>
         </select> 
	     <input type="button" name="Button" value="Abrir" onclick="window.open('default.asp?op=list_msg&enviado=<%=request("enviado")%>&aspid='+ AspId +'&idarea=' + document.getElementById('cod_area').value,'_self')" />
	     <br />
		 <% end if %>
	   	     <div align="center">
	    <%
		  
		 
		  dim enviadDataMsg
		  			
					if request("enviado") = "1" then
						 ReDim GridData(6), GridName(6)
						enviadDataMsg = "Enviado"
						sql = "select *, datasend_msg as data from mala_msg where criado=1"
						sql = sql & " and enviado_msg = 1"
					else
						 ReDim GridData(4), GridName(4)
						enviadDataMsg = "Criado"
						sql = "select *, data_msg as data from mala_msg where criado=1 and enviado_msg <> 1"
					end if 				
					
					if request("modelo") = "1" then
						sql = sql & " and modelo_msg = 1"
					else
						sql = sql & " and modelo_msg <> 1"
					end if 
					
					if request("idarea") <> "" then
						sql = sql & " and cod_area = "&request("idarea")
					end if
					
					sql = sql & " order by nome_msg"
					'response.write sql
					GridSQL = sql
					
					GridName(0) = "-;35"
					GridName(1) = "Cod;35"
					GridName(2) = "Nome;200"
					GridName(3) = enviadDataMsg & ";70"
					GridName(4) = "Editar;60"
					
					if request("enviado") = "1" then
						GridName(5) = "Enviado;70"
						GridName(6) = "Abertos;60"
					end if
									
					GridData(0) = "$cod_msg;1;0;$case({cod_msg}!0):<input type='checkbox' name='codMsg' value='{cod_msg}' onclick='nList(this.value,document.form.sCod_vale)'>"
					GridData(1) = "cod_msg;1;0"
					GridData(2) = "nome_msg;0;0"
					GridData(3) = "data;0;0"
					GridData(4) = "*Editar;1;default.asp?op=cad_msg&modelo={modelo_msg}&id={cod_msg}&aspid=[aspid]"
					
					if request("enviado") = "1" then
						GridData(5) = "send_msg;0;0"
						GridData(6) = "lido_msg;0;0"
					end if
						
					GridDate = "dd/mm/yyyy"
					GridPoss = "center"
					'GridURLS = 1
					'// Chama a Função
					
					GridForm Gkeys,15
					
	%>
	   </div>
	   </td>
    </tr>
	<%
		case "send"
	%>
	 <tr>
	   <td id="menutab">
	   	<%
	  	'// Gera Menus ----------------------------- >
		'// 0-menu;1-voltar;2-imprimir;3-confirmar;4-incluir;
		'// 5-apagar;6-editar;7-reportar;8-mostrarall;
		'// 9-ocultarall;10-item.
		'//----------------------------------------- >
		redim AspMenuOp(2) 'Define o numero de itens
		'//campos ---------------------------------
		AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?aspid=' + AspId)"
		AspMenuOp(1) = "Enviar|3|send_msg()"
		'//----------------------------------------
		AspMenu() 'gera menu
		'// Gera Menus ---------------------------- >
		%>
		
	   </td>
    </tr>
	 <tr>
	   <td id="conteudo_page"><span class="texto_size_10">Grupo de E-Mail </span><br />
		
		 <br />
		 <table width="70%" border="0" cellspacing="0" cellpadding="0">
           <tr>
             <td> <%
		dim chkCount, chkCountDiv, ChkText, ChkValue, ChkName
		chkCount = 0
		chkCountDiv = 0
		db "leitura",rs,sql,"select * from mala_grupos where criado=1 order by nome_grupo",conn
		while not rs.eof
			if chkCount = 0 then
				chkCountDiv = 1 + chkCountDiv
				ChkText = ChkText &  "<td width='[largCols]%'  valign='top'>"
			end if
				ChkValue = trim(rs("cod_grupo"))
				ChkName  = trim(rs("nome_grupo"))
				ChkText = ChkText & "<table width='100%' border='0' cellspacing='2' cellpadding='2'"&chr(asc(">"))&br
				ChkText = ChkText & "<tr"&chr(asc(">"))&br
				ChkText = ChkText & " <td width='1'"&chr(asc(">"))&br
				ChkText = ChkText & "<input type='checkbox' name='cod_grupo' id='cod_grupo' value='"&ChkValue&"'"
				ChkText = ChkText & " /"&chr(asc(">"))
				ChkText = ChkText & "</td"&chr(asc(">"))&br
				ChkText = ChkText & "<td class='texto_size_10'"&chr(asc(">"))&br
				ChkText = ChkText & "<span class='texto_size_10'"&chr(asc(">"))&ChkName&"</span"& chr(asc(">")) & br
				ChkText = ChkText & "</td"&chr(asc(">"))&br
				ChkText = ChkText & "</tr"&chr(asc(">"))&br
				ChkText = ChkText & "</table"&chr(asc(">"))&br
			if chkCount = 4 then
				ChkText = ChkText &  " </td>"
				chkCount = 0
			else
				chkCount = 1 + chkCount
			end if
			rs.movenext
		wend
		if chkCountDiv <> 0 then
			ChkText = replace(ChkText,"[largCols]",trim(100/chkCountDiv))
		end if
			response.write ChkText

		 %></td>
           </tr>
         </table>
		 <br />
         <span class="texto_size_10">Sexo</span><br /><input name="op" type="hidden" value="enviando" />
<input name="aspid" type="hidden" value="<%=aspid%>" />         <select name="sexo_email" id="sexo_email">
           <option value="1">Masculino</option>
           <option value="0">Feminino</option>
           <option value="2" selected="selected">Ambos</option>
         </select>
         <br />
         <span class="texto_size_10">Tipo de e-mail </span><br />
         <select name="atv_email" id="atv_email">
           <option value="1" selected="selected">Ativos</option>
           <option value="0">Inativos</option>
           <option value="2">Ambos</option><input name="Gpage" type="hidden" value="0" />
         </select><input name="cod_msg" type="hidden" id="cod_msg" value="<%=formus("sCod_vale",0)%>" />
         <br />
         <span class="texto_size_10">Mensagem Selecionadas</span><br />
         <textarea name="msg" cols="60" rows="6" id="msg" readonly="readonly"><%
		 	listemail = split(formus("sCod_vale",0),",")
			for i = lbound(listemail) to ubound(listemail)
			 	db "leitura",rs,sql,"select * from MalaListMsgArea where cod_msg="&listemail(i),conn
				if not rs.eof and not rs.bof then
		 %><%=rs("cod_msg")%> - <%=trim(rs("nome_msg"))%> - Area: <%=trim(rs("nome_area")) & chr(10)%><%
				end if
			next
		%>
		</textarea>
		
<br /></td>
    </tr>
	<%
		case "alt_stats"
			
				listemail = split(formus("sCod_vale",0),",")
				for i = lbound(listemail) to ubound(listemail)
					db "leitura",rs,sql,"select atv_email, cod_email from mala_email where cod_email="&listemail(i),conn
					if rs("atv_email") = 1 then
						db "normal",,sql,"update mala_email set atv_email=0 where cod_email="&rs("cod_email"),conn
					else
						db "normal",,sql,"update mala_email set atv_email=1 where cod_email="&rs("cod_email"),conn
					end if
				next	
				
				voltar request.ServerVariables("SCRIPT_NAME")&"?op=list_email&idgrupo="&formus("idgrupo",3)&"&aspid="&aspid,"","_self"
			
		case "del"
		
			select case request("tp")
			case "area"
			
				db "normal",,sql,"delete from mala_area where cod_area="&request("id"),conn
				voltar request.ServerVariables("SCRIPT_NAME")&"?op=list_area&aspid="&aspid,"","_self"
				
			case "grupo"
				
				db "normal",,sql,"delete from mala_email where cod_grupo="&	request("id"),conn	
				db "normal",,sql,"delete from mala_grupos where cod_grupo="&request("id"),conn
				voltar request.ServerVariables("SCRIPT_NAME")&"?op=list_grupo&aspid="&aspid,"","_self"
				
			case "email"
			
				listemail = split(formus("sCod_vale",0),",")
				for i = lbound(listemail) to ubound(listemail)
					db "normal",,sql,"delete from mala_email where cod_email="&	listemail(i),conn	
				next	
				
				voltar request.ServerVariables("SCRIPT_NAME")&"?op=list_email&idgrupo="&formus("idgrupo",3)&"&aspid="&aspid,"","_self"
			
			case "msg"
				
				listemail = split(formus("sCod_vale",0),",")
				for i = lbound(listemail) to ubound(listemail)
					db "normal",,sql,"delete from mala_msg where cod_msg="&	listemail(i),conn	
				next	
				
				voltar request.ServerVariables("SCRIPT_NAME")&"?op=list_msg&idarea="&formus("idarea",3)&"&modelo="&formus("modelo",3)&"&enviado="&formus("enviado",3)&"&aspid="&aspid,"","_self"
			
			case else
				voltar request.ServerVariables("SCRIPT_NAME")&"?aspid="&aspid,"","_self"
			end select	
			
			
		case "enviando"
	%>
	<tr>
	   <td id="menutab">	   	<%
	  	'// Gera Menus ----------------------------- >
		'// 0-menu;1-voltar;2-imprimir;3-confirmar;4-incluir;
		'// 5-apagar;6-editar;7-reportar;8-mostrarall;
		'// 9-ocultarall;10-item.
		'//----------------------------------------- >
		redim AspMenuOp(1) 'Define o numero de itens
		'//campos ---------------------------------
		AspMenuOp(0) = "Abortar|5|WindowOpen('default.asp?aspid=' + AspId)"
		
		'//----------------------------------------
		AspMenu() 'gera menu
		'// Gera Menus ---------------------------- >
		%></td>
    </tr>
	<tr>
	   <td id="conteudo_page">
	   <div id="msg"></div>
	   	   <script>
		   
		   var tempo = 0.5
			var militempo = 1000
			var x = tempo * militempo
			var textsl
			var textsplit
		   function AjaxLoads(pg,query,camada){
			//--------------------------------
			// Ajax - Paulo Corcino - 1.5
			//--------------------------------
			
			var http = null; //variavel
			
			 //--------------------------------------------------------------
			 // Detecta o componentes XML
			 //---------------------------------------------------------------
			
			if (typeof XMLHttpRequest != "undefined") {
				http = new XMLHttpRequest();
			}else if (window.ActiveXObject){
			  var aVersions = ["MSXML2.XMLHttp.5.0",
				"MSXML2.XMLHttp.4.0","MSXML2.XMLHttp.3.0",
				"MSXML2.XMLHttp","Microsoft.XMLHttp"];
		
			  for (var i = 0; i < aVersions.length; i++) {
				try{
					var http = new ActiveXObject(aVersions[i]);
				}
				catch(oError){
					//Do nothing
				}
			  }
			}
			
			  //-----------------------------------------------------------------------
			  // Página do Post
			  //-----------------------------------------------------------------------
			  web = pg;
				  
			  //------------------------------------------------------------------------
			  // Envio dos Dados
			  //------------------------------------------------------------------------
			  http.open('post', web, true);
			  //document.getElementById(camada).innerHTML = '<div aling=left class=fonte_10_b>Carregando....</div>';
			  http.setRequestHeader('Content-Type','application/x-www-form-urlencoded');
			
			  //-----------------------------------------------------------------------
			  // Leitura de Dados
			  //-----------------------------------------------------------------------
			  http.onreadystatechange = function()
			  {
					if(http.readyState == 4){
					//	if (http.status == 200) {	
							//alert(unescape(http.responseText.replace(/\+/g,' ')));
							textsl = unescape(http.responseText.replace(/\+/g,' '));
							textsplit = textsl.split("::");
									
							document.getElementById(camada).innerHTML = textsplit[1];
							window.setTimeout("pular(" + textsplit[0] + ")",x)
						//}else{
						//	document.getElementById(camada).innerHTML = "<div aling=left class=fonte_10_b>Erro na abertura da página: " + http.status + " - " + http.statusText + "</div>";
						//}
					}		 
			  };
			  
			//------------------------------------------------------------------------
			// Captura a Pagina
			//------------------------------------------------------------------------
		
			
			if(query!=''){  
				iTimeoutId = setTimeout(function(){http.send(query);},800);
			}else{
				//alert("aqui")
				if(document.getElementById('fotosel').value != "")
				{
				//ocasião especial
					iTimeoutId = setTimeout(function(){http.send('sel=' + document.getElementById('fotosel').value);},800);
					
				}else{
					
					iTimeoutId = setTimeout(function() {http.send(null);},800);
					
				}
			}
			  
		}
		   
		   

<%
	'If cint(iPageCurrent) < cint(iPageCount) Then
%>
AjaxLoads("sends.asp?aspid=<%=aspid%>&op=sends&Gpage=1","<%=request.querystring%>","msg")
function pular(p){
	if (p == "0"){
		//parado
	}else{
		AjaxLoads("sends.asp?aspid=<%=aspid%>&op=sends&Gpage="+p,"<%=request.querystring%>","msg")
	}//window.open("","_self")
}
<% 'end if %>
</script>
	   </td>
    </tr>
	<%
		case "sends"
	%>
   
	<%	
		case else
	%>
	 <tr>
	   <td id="menutab">&nbsp;</td>
    </tr>
	 <tr>
	   <td id="conteudo_page"><table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="307"><p class="texto_size_10">Selecione qual o tipo de informa&ccedil;&atilde;o deseja visualizar:</p>
	     <ul class="texto_size_10">
           <li><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=cad_msg&amp;aspid=<%=aspid%>" target="_self">Criar novo E-Mail </a></li>
	       <li><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=list_msg&amp;aspid=<%=aspid%>" target="_self"> Caixa de Saida</a></li>
	       <li><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=list_msg&amp;enviado=1&amp;aspid=<%=aspid%>">Itens Enviados</a> </li>
	       <li><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=cad_msg&amp;modelo=1&amp;aspid=<%=aspid%>" target="_self">Criar Modelo de E-Mail</a></li>
           <li><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=list_msg&amp;modelo=1&amp;aspid=<%=aspid%>">Listar Modelos</a></li>
	       <li><a href="default.asp?op=cad_email&amp;aspid=<%=aspid%>">Cadastrar E-Mails</a></li>
	       <li><a href="default.asp?op=cad_list_email&amp;aspid=<%=aspid%>">Cadastrar Lista de E-Mails </a></li>
	       <li><a href="default.asp?op=list_email&amp;aspid=<%=aspid%>">Listar E-Mails</a></li>
	       <li><a href="default.asp?op=cad_grupo&amp;aspid=<%=aspid%>">Criar Grupos</a></li>
	       <li><a href="default.asp?op=list_grupo&amp;aspid=<%=aspid%>">Listar Grupos</a></li>
	       <li><a href="default.asp?op=cad_area&amp;aspid=<%=aspid%>">Criar &Aacute;reas</a></li>
	       <li><a href="default.asp?op=list_area&amp;aspid=<%=aspid%>">Listar &Aacute;reas </a></li>
	     </ul>
       <p></p></td>
	   <%
	   		db "leitura",rs,sql,"select * from MalaResumo",conn
	   %>
    <td width="286"><table width="95%" border="0" align="center" cellpadding="2" cellspacing="1">
      <tr>
        <td colspan="2" bgcolor="#00FFCC" class="texto_size_10"><strong>Estat&iacute;sticas</strong></td>
        </tr>
      <tr>
        <td width="82%" class="texto_size_10">E-Mails Ativos </td>
        <td width="18%" align="right" class="texto_size_10"><%=rs("email_atv")%></td>
      </tr>
      <tr>
        <td class="texto_size_10">E-Mails Inativos </td>
        <td align="right" class="texto_size_10"><%=rs("email_inativo")%></td>
      </tr>
      <tr>
        <td class="texto_size_10"><strong>Total</strong></td>
        <td align="right" class="texto_size_10"><%=rs("email_cad")%></td>
      </tr>
      <tr>
        <td colspan="2"><img src="../img/blank.gif" width="1" height="3" /></td>
        </tr>
      <tr>
        <td class="texto_size_10">Caixa de Saida </td>
        <td align="right" class="texto_size_10"><%=rs("msg_saida")%></td>
      </tr>
      <tr>
        <td class="texto_size_10">Enviados</td>
        <td align="right" class="texto_size_10"><%=rs("msg_enviada")%></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td colspan="2" class="texto_size_10">Estimativa de audi&ecirc;ncia*</td>
        </tr>
      <tr>
        <td colspan="2">
		<%
			'cor do nivel
			dim nCor
			if rs("aprov") <> "" then			
				if cint(rs("aprov")) > 65 then
					nCor = "#00CC00"
				elseif cint(rs("aprov")) < 15 then
					nCor = "#FF0000"
				else
					nCor = "#99FF00"
				end if
			end if
		%>
		<div id="porcentaud" style="width:<%=rs("aprov")%>%; background-color:<%=nCor%>"><img src="../img/blank.gif" width="1" height="8" /></div><span class="texto_size_10"><%=rs("aprov")%>%</span></td>
        </tr>
      <tr>
        <td colspan="2" class="texto_size_10">*M&eacute;dia de e-mails lidos por mensagens enviadas. </td>
      </tr>
    </table></td>
  </tr>
</table>
</td>
    </tr>
	<% end select %>
  </table>
</form>
</body>
</html>
