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
<title>Noticias</title>
<link href="../css/estilos.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" type="text/javascript" src="../js/script.js"></script>
<script language="JavaScript" type="text/javascript" src="../js/form_noticia.js"></script>
<script language="JavaScript" type="text/javascript" src="../js/form.js"></script>
<script>AspId = '<%=aspid%>'</script>

<%
	if formus("op",3) = "cad_noticia" and lcase(request.ServerVariables("REQUEST_METHOD")) <> "post" then
		
	if formus("id",3)<>"" then
		id = formus("id",3)
	else
		db "leitura",rs,sql,"select id from not_noticia where criado=0",conn
		if rs.eof and rs.bof then
			db "normal",,sql,"insert into not_noticia(criado) values (0)",conn
			db "leitura",rs,sql,"select id from not_noticia where criado=0",conn
		end if
		id = rs("id")
	end if
%>
	<script type="text/javascript">
		  _editor_url = "../editor/";
		  _editor_lang = "en";
		</script>
		<script type="text/javascript" src="../editor/htmlarea.asp?tela=Noticia&id=<%=id%>"></script>
		
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
<!--[Info File]-->
<!--[!]
	Album Book
	Criado: Paulo Corcino		
	@Ano: 2006 - @Version: 1.5
[!]-->
<!--[Info File]-->
<body <% if formus("op",3) = "cad_noticia" and lcase(request.ServerVariables("REQUEST_METHOD")) <> "post" then %> onload="initEditor();" <% end if %>>
<form method="post" name="form" id="form">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td id="titulo_pagina"><div id="matrix_titulo" style="float:left">Noticias</div></td>
  </tr>
  <%
		dim rs,sql,id, criado, x, ids
		select case request.QueryString("op")
		case "cad_area"
		dim area
		'//Nova Area ----------------------------------
		if lcase(request.ServerVariables("REQUEST_METHOD")) <> "post" then
			
			id = formus("id",3)
			
			if id<>"" then
				subTitle("> Alterar Área")
				'//Alterar Dados
				db "leitura",rs,sql,"select top 1 * from not_area where id="&id,conn
				area = trim(rs("area"))
				
			else
				subTitle("> Nova Área")
				'//nova area
				db "leitura",rs,sql,"select top 1 * from not_area where criado=0",conn
				if rs.eof and rs.bof then
					'//não existe registro
					db "normal",,sql,"insert into not_area(criado) values(0)",conn
					db "leitura",rs,sql,"select top 1 * from not_area where criado=0",conn
				end if						
				
			end if
			
			id = rs("id")
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
      <input name="id" type="text" id="id" value="<%=id%>" size="3" readonly="true" />
      <br />
      <span class="texto_size_10">Nome da &Aacute;rea</span><br />
<input name="area" type="text" id="area" value="<%=area%>" size="30" maxlength="255" /><input name="criado" type="hidden" value="<%=criado%>" /></td>
    </tr>
    <%
		else
		
			id = formus("id",0)
			area = formus("area",0)
			criado = cint(formus("criado",0))
			'///--------- [salva dados] ---->
			db "leitura",rs,sql,"select * from not_area where id<>"&id&" and area='"&area&"'",conn
			'verifica se existe algo cadastrado com este Nome
				if rs.eof and rs.bof then 
				
					db "normal",,sql,"update not_area set area='"&area&"', criado=1 where id="&id,conn
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
		'//Listar Area ----------------------------------
		prep_deletar()
	%>
    <tr>
      <td id="menutab">
	  	  <%
	  	'// Gera Menus ----------------------------- >
		'// 0-menu;1-voltar;2-imprimir;3-confirmar;4-incluir;
		'// 5-apagar;6-editar;7-reportar;8-mostrarall;
		'// 9-ocultarall;10-item.
		'//----------------------------------------- >
		redim AspMenuOp(1) 'Define o numero de itens
		'//campos ---------------------------------
		AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?aspid=' + AspId)"
		'AspMenuOp(1) = "Salvar|3|miniMenu_area()"
		'//----------------------------------------
		AspMenu() 'gera menu
		'// Gera Menus ---------------------------- >
		%>
	  </td>
    </tr>
    <tr>
      <td id="conteudo_page">
	   <div align="center">
	    <%
		  subTitle("> Listar Área")
		  Dim GridData(3), GridName(3), GridSQL, GridDate, GridPoss
				
				
					GridSQL = "select * from not_area where criado=1;"
					
					GridName(0) = "Cod;35"
					GridName(1) = "Área;370"
					GridName(2) = "Alterar;60"
					GridName(3) = "Deletar;60"
									
					GridData(0) = "id;1;0"
					GridData(1) = "area;0;0"
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
		case "cad_noticia"
		
		'//Nova Noticia ----------------------------------
		if lcase(request.ServerVariables("REQUEST_METHOD")) <> "post" then
			
			id = formus("id",3)
			dim ntitulo, ntexto, nfonte, ndata, ncoment, njornalista
			ndata = fdate(now(),"dd/mm/yyyy")
			
			if id<>"" then
				subTitle("> Alterar Noticia")
				'//Alterar Dados
				db "leitura",rs,sql,"select top 1 * from not_noticia where id="&id,conn
				id_s = cstr(rs("id_s"))
				ntitulo = trim(rs("titulo"))
				ntexto = trim(rs("texto"))
				nfonte = trim(rs("fonte"))
				ndata = fdate(rs("data"),"dd/mm/yyyy")
				ncoment = trim(rs("coment"))
				njornalista = trim(rs("jornalista"))
				
			else
				subTitle("> Nova Noticia")
				'//nova area
				db "leitura",rs,sql,"select top 1 * from not_noticia where criado=0",conn
				if rs.eof and rs.bof then
					'//não existe registro
					db "normal",,sql,"insert into not_noticia(criado) values(0)",conn
					db "leitura",rs,sql,"select top 1 * from not_noticia where criado=0",conn
				end if						
				
			end if
			
			id = rs("id")
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
		AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?aspid=' + AspId)"
		AspMenuOp(1) = "Salvar|3|miniMenu_noticia()"
		'//----------------------------------------
		AspMenu() 'gera menu
		'// Gera Menus ---------------------------- >
		%>
	
	</td>
  </tr>
  <tr>
    <td id="conteudo_page">
	<!--[start webword]-->
			
		
	<!--[start webword]-->
	<table width="364" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="56" height="41"><span class="texto_size_10">Cod</span><br />
          <input name="id" type="text" id="id" value="<%=id%>" size="3" maxlength="3" /></td>
        <td width="308" class="texto_size_10">&Aacute;rea<br />
          <select name="id_s" id="id_s">
		  <%
				db "leitura",rs,sql,"select * from not_area where criado=1",conn
				while not rs.eof
			%>
            <option value="<%=rs("id")%>" <% if cstr(rs("id")) = id_s then%> selected <% end if %>><%=trim(rs("area"))%></option>
			<%
				rs.movenext
			wend
			%>
          </select> 
          <a href="javascript:alert(document.form.ta.value)"></a>          </td>
      </tr>
    </table>
	<span class="texto_size_10">Titulo</span><br />
	<input name="ntitulo" type="text" id="ntitulo" value="<%=ntitulo%>" size="50" maxlength="255" />
	<br />
	<span class="texto_size_10">Texto da Noticia</span><br />
	<table width="90%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><textarea id="ta" name="ta" style="width:99%" rows="25" cols="80"><%=ntexto%></textarea></td>
  </tr>
</table>

	
	<br />
    <span class="texto_size_10">Data da Publica&ccedil;&atilde;o</span><br />
	<input name="ndata" type="text" id="ndata" value="<%=ndata%>" onBlur="if(this.value==''){this.value='<%=fdate(now(),"dd/mm/yyyy")%>'}else{dataVal(this)}" onFocus="this.value=''" size="12" maxlength="10" />
	<br />
	<span class="texto_size_10">Resumo</span><br />
	<textarea name="ncoment" cols="50" rows="5" onkeypress="CharMax(this,255)" id="ncoment"><%=ncoment%></textarea>
	<br />
	<span class="texto_size_10">Jornalista</span><br />
	<input name="njornalista" type="text" id="njornalista" value="<%=njornalista%>" size="40" maxlength="255" />
	<br />
	<span class="texto_size_10">Fonte</span><br />
	<input name="nfonte" type="text" id="nfonte" value="<%=nfonte%>" size="40" maxlength="255" />
		<script>
	if(isIE()){
		oDateMasks = new Mask("dd/mm/yyyy", "date");
		oDateMasks.attach(document.form.ndata);
		
	}
	</script>
	</td>
  </tr>

  <%
		else
			'//Salva Noticia
			
			db "normal",,sql,"update not_noticia set id_s="&formus("id_s",1)&",titulo='"&formus("ntitulo",1)&"',texto='"&formus("ta",0)&"',fonte='"&formus("nfonte",1)&"',data='"&fdate(formus("ndata",0),"yyyymmdd")&"',coment='"&formus("ncoment",0)&"',jornalista='"&formus("njornalista",1)&"',criado=1 where id="&formus("id",0),conn
			alert("Noticia gravada com sucesso!!!")
			'for each i in request.form
				'response.write i & "<br>"
				'response.write request.form(i) & "<br>"
			'next
			'response.write "terste"
			voltar request.ServerVariables("SCRIPT_NAME")&"?aspid="&aspid&"&op=list_noticia&periodo="&right(formus("periodo",1),7)&"&id_s="&formus("id_s",1),"","_self"
			
			
		end if
	
  	case "list_noticia"
			'//Listar Noticias ------------------------------------
		prep_deletar()
		
		dim id_s, periodo, infoperiodo, DiaIni, DiaFim, splitData
		  
		id_s = formus("id_s",3)
		periodo = formus("periodo",3)
  %>
  <tr>
    <td id="menutab">
		<%
	  	'// Gera Menus ----------------------------- >
		'// 0-menu;1-voltar;2-imprimir;3-confirmar;4-incluir;
		'// 5-apagar;6-editar;7-reportar;8-mostrarall;
		'// 9-ocultarall;10-item.
		'//----------------------------------------- >
		redim AspMenuOp(6) 'Define o numero de itens
		'//campos ---------------------------------
		AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?aspid=' + AspId)"
		AspMenuOp(1) = "Novo|3|WindowOpen('default.asp?op=cad_noticia&aspid=' + AspId)"
		AspMenuOp(2) = "Alt Stat.|6|stat_noticia('"&id_s&"','"&periodo&"')"
		AspMenuOp(3) = "Destaq.|10|destaq_noticia('"&id_s&"','"&periodo&"')"
		AspMenuOp(4) = "Excluir|5|excluir_noticia('"&id_s&"','"&periodo&"')" 
		AspMenuOp(5) = "Coments|7|coment_noticia('"&id_s&"','"&periodo&"')"
		'//----------------------------------------
		AspMenu() 'gera menu
		'// Gera Menus ---------------------------- >
		%>
	</td>
  </tr>
  <tr>
    <td id="conteudo_page"><table width="373" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><span class="texto_size_10">Selecione a &Aacute;rea</span><br />
          <select name="id_s" id="id_s">
		  	<%
				db "leitura",rs,sql,"select * from not_area where criado=1",conn
				while not rs.eof
			%>
            <option value="<%=rs("id")%>" <% if cstr(rs("id")) = id_s then%> selected <% end if %>><%=trim(rs("area"))%></option>
			<%
				rs.movenext
			wend
			%>
          </select></td>
        <td>
		<%
			if id_s <> "" then
		%>
		<span class="texto_size_10">Filtrar por Peri&oacute;dos </span><br />
          <select name="periodo" id="periodo">
	              <option value="0" selected="selected">Todos os Peri&oacute;dos</option>
		  	<%
				db "leitura",rs,sql,"select * from NoticiaPeriodo where id_s="& id_s,conn
				while not rs.eof
			%>
				<option value="<%=rs("mes")%>/<%=rs("ano")%>" <% if periodo = rs("mes")&"/"&rs("ano") then%> selected <% end if %>><%=mesexts(rs("mes"))%>/<%=rs("ano")%></option>
			<%
					rs.movenext
				wend
			%></select>
			<%
				else
			%>
			<input type="hidden" name="periodo" value="" />
			<% end if %>
			</td>
        <td align="center" valign="bottom"><input type="button" name="Button" value="Listar Noticias" onclick="window.open('default.asp?op=list_noticia&id_s='+document.form.id_s.value+'&aspid=' + AspId + '&periodo=' + document.form.periodo.value,'_self');" /></td>
      </tr>
    </table>
	<input name="sCod_vale" type="hidden" value="" />
		   <div align="center">
	    <%

		  infoperiodo = ""
		  
		  '//Caso o periodo exista
		  if periodo <> "" and  periodo <> "0" then
		  
		  	splitData = split(periodo,"/")
			DiaIni = day(dateserial(splitData(1),splitData(0),01))
		  	DiaFim = day(dateserial(splitData(1),cint(splitData(0))+1,1-1))
			infoperiodo = " and  (data > CONVERT(DATETIME, '"&splitData(1)&"-"&splitData(0)&"-"&DiaIni&" 00:00:00', 102) AND data < CONVERT(DATETIME, '"&splitData(1)&"-"&splitData(0)&"-"&DiaFim&" 00:00:00', 102))"
			
  		  end if
		  
		  'response.write infoperiodo 
		  subTitle("> Listar Notícias")
		  reDim GridData(10), GridName(10)
				
				
					GridSQL = "select * from NoticiaList where id_s=" & id_s & " " & infoperiodo
					
					GridName(0) = "-;35"
					GridName(1) = "Cod;35"
					GridName(2) = "Titulo;150"
					GridName(3) = "Coment;250"
					GridName(4) = "Data;75"
					GridName(5) = "Visitas;60"
					GridName(6) = "Coments;60"
					GridName(7) = "Com.New;60"
					GridName(8) = "Status;60"
					GridName(9) = "Destaq;55"
					GridName(10) = "Editar;60"
														
					GridData(0) = "$id;1;0;$case({id}!0):<input type=checkbox name=codnoticia value={id} onclick='nList(this.value,document.form.sCod_vale)'>"
					GridData(1) = "id;1;0"
					GridData(2) = "titulo;0;0"
					GridData(3) = "coment;0;0"
					GridData(4) = "data;1;0"
					GridData(5) = "counter;1;0"
					GridData(6) = "coments;1;0"
					GridData(7) = "comentsnew;1;0"
					GridData(8) = "$atv;1;0;$case({atv}=0):Off-Line:On-Line"
					GridData(9) = "$dst;1;0;$case({dst}!1):Não:Sim"
					GridData(10) = "*Editar;1;default.asp?op=cad_noticia&id=[col0]&aspid=[aspid]"
					
							
					GridDate = "dd/mm/yyyy"
					GridPoss = "center"
					'GridURLS = 1
					'// Chama a Função
					
					if id_s <> "" then
						GridForm Gkeys,15
					end if
					
	%>
	    </div>
	
	</td>
  </tr>
   <%
  	case "del"
	
		select case request.QueryString("tp")
		case "nt"
		
			ids = split(formus("sCod_vale",0),",")
			
			for x = lbound(ids) to ubound(ids)
			
				db "normal",,sql,"delete from not_noticia where id="&ids(x),conn

			
			next
			voltar request.ServerVariables("SCRIPT_NAME")&"?aspid="&aspid&"&op=list_noticia&periodo="&formus("periodo",3)&"&id_s="&formus("id_s",3),"","_self"
		case "area"
			db "normal",,sql,"delete from not_area where id="&formus("id",3),conn
			voltar request.ServerVariables("SCRIPT_NAME")&"?op=list_area&aspid="&aspid,"","_self"
		case "coment"
			db "normal",,sql,"delete from not_coment where id="&formus("sCod_vales",0),conn
			voltar request.ServerVariables("SCRIPT_NAME")&"?"&replace(request.QueryString,"del","coment"),"","_self"
		case else
			voltar request.ServerVariables("SCRIPT_NAME")&"?aspid="&aspid,"","_self"
		end select
	
	case "stat"
	
			ids = split(formus("sCod_vale",0),",")
			
			for x = lbound(ids) to ubound(ids)
				
				db "leitura",rs,sql,"select atv from not_noticia where id="&ids(x),conn

				if cint(rs("atv")) = 1 then
					db "normal",,sql,"update not_noticia set atv=0 where id="&ids(x),conn
				else
					db "normal",,sql,"update not_noticia set atv=1 where id="&ids(x),conn
				end if
			
			next
			
			voltar request.ServerVariables("SCRIPT_NAME")&"?aspid="&aspid&"&op=list_noticia&periodo="&formus("periodo",3)&"&id_s="&formus("id_s",3),"","_self"
	
	case "dest"
  
  			ids = split(formus("sCod_vale",0),",")
			
			for x = lbound(ids) to ubound(ids)
				
				db "leitura",rs,sql,"select dst from not_noticia where id="&ids(x),conn
				'response.write ids(x) & " " & rs("dst")
				if cint(rs("dst")) = 1 then
					db "normal",,sql,"update not_noticia set dst=0 where id="&ids(x),conn
				else
					db "normal",,sql,"update not_noticia set dst=1 where id="&ids(x),conn
				end if
			
			next
  		voltar request.ServerVariables("SCRIPT_NAME")&"?aspid="&aspid&"&op=list_noticia&periodo="&formus("periodo",3)&"&id_s="&formus("id_s",3),"","_self"
	case "comentstat"
	
	  		ids = split(formus("sCod_vales",0),",")
			
			for x = lbound(ids) to ubound(ids)
			
				db "leitura",rs,sql,"select atv from not_coment where id="&ids(x),conn

				if cint(rs("atv")) = 1 then
					db "normal",,sql,"update not_coment set atv=0 where id="&ids(x),conn
				else
					db "normal",,sql,"update not_coment set atv=1 where id="&ids(x),conn
				end if
				
			next
				
				voltar request.ServerVariables("SCRIPT_NAME")&"?"&replace(request.QueryString,"comentstat","coment"),"","_self"
				
	
  	case "coment"
	'//Coment ------------------------------------
	%>
		<tr>
      <td id="menutab">
	  		<%
	  	'// Gera Menus ----------------------------- >
		'// 0-menu;1-voltar;2-imprimir;3-confirmar;4-incluir;
		'// 5-apagar;6-editar;7-reportar;8-mostrarall;
		'// 9-ocultarall;10-item.
		'//----------------------------------------- >
		redim AspMenuOp(3) 'Define o numero de itens
		'//campos ---------------------------------
		AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?"&replace(request.QueryString,"coment","list_noticia")&"')"
		AspMenuOp(1) = "Excluir|5|excluir_noticiaC('"&id_s&"','"&periodo&"','"&request("sCod_vale")&"')" 
		AspMenuOp(2) = "Alt Stat.|6|stat_noticiaC('"&id_s&"','"&periodo&"','"&request("sCod_vale")&"')"
			'//----------------------------------------
		AspMenu() 'gera menu
		'// Gera Menus ---------------------------- >
		%>
	  </td>
    </tr>
    <tr>
      <td id="conteudo_page">
	   <div align="center">

	   <input name="sCod_vales" type="hidden" value="" />
	  <%
		  subTitle("> Comentários")
		  redim GridData(6), GridName(6)
				
				
					GridSQL = "select * from NoticiaComentList where id_n="&request("sCod_vale")
					
					GridName(0) = "-;35"
					GridName(1) = "Cod;35"
					GridName(2) = "De;150"
					GridName(3) = "Email;120"
					GridName(4) = "Msg;360"
					GridName(5) = "Data;60"
					GridName(6) = "Status;60"

					GridData(0) = "$id;1;0;$case({id}!0):<input type=checkbox name=codnoticia value={id} onclick='nList(this.value,document.form.sCod_vales)'>"				
					GridData(1) = "id;1;0"
					GridData(2) = "de;0;0"
					GridData(3) = "email;0;0"
					GridData(4) = "msg;0;0"
					GridData(5) = "data;0;0"
					GridData(6) = "$atv;1;0;$case({atv}!0):on-line:off-line"

							
					GridDate = "dd/mm/yyyy"
					GridPoss = "center"
					'GridURLS = 1
					'// Chama a Função
							
					GridForm Gkeys,15
					
	%></div>
	  </td>
    </tr>
	<%
  	case else
		'//Menu Principal ------------------------------------
	%>
	<tr>
      <td id="menutab">&nbsp;</td>
    </tr>
    <tr>
      <td id="conteudo_page"><p class="texto_size_10">Selecione qual o tipo de informa&ccedil;&atilde;o deseja visualizar:</p>
        <ul class="texto_size_10">
          <li><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=cad_noticia&aspid=<%=aspid%>" target="_self">Nova Noticia </a></li>
          <li><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=list_noticia&amp;aspid=<%=aspid%>" target="_self">Listar Noticias </a></li>
          <li><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=cad_area&aspid=<%=aspid%>" target="_self">Nova &Aacute;rea </a></li>
          <li><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=list_area&aspid=<%=aspid%>" target="_self">Listar &Aacute;reas </a></li>
        </ul>
      <p>&nbsp;</p></td>
    </tr>
	<%
		end select
	%>
</table>
</form>

</body>
</html>


