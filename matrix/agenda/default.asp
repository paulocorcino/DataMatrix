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
<title>Agenda</title>
<link href="../css/estilos.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" type="text/javascript" src="../js/script.js"></script>
<script language="JavaScript" type="text/javascript" src="../js/form_agenda.js"></script>
<script language="JavaScript" type="text/javascript" src="../js/form.js"></script>
<script>AspId = '<%=aspid%>'</script>

<%
	
	if formus("op",3) = "cad_agenda" and lcase(request.ServerVariables("REQUEST_METHOD")) <> "post" then
		
	if formus("id",3)<>"" then
		id = formus("id",3)
	else
		db "leitura",rs,sql,"select cod_agenda from tb_agenda where criado=0",conn
		if rs.eof and rs.bof then
			db "normal",,sql,"insert into tb_agenda(criado) values (0)",conn
			db "leitura",rs,sql,"select cod_agenda from tb_agenda where criado=0",conn
		end if
		id = rs("cod_agenda")
	end if
%>
	<script type="text/javascript">
		  _editor_url = "../editor/";
		  _editor_lang = "en";
		</script>
		<script type="text/javascript" src="../editor/htmlarea.asp?tela=Agenda&id=<%=id%>"></script>
		
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

function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_jumpMenuGo(selName,targ,restore){ //v3.0
  var selObj = MM_findObj(selName); if (selObj) MM_jumpMenu(targ,selObj,restore);
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
<body  <% if formus("op",3) = "cad_agenda" and lcase(request.ServerVariables("REQUEST_METHOD")) <> "post" then %> onload="initEditor();" <% end if %>>
<form method="post" name="form" id="form">
  <table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td id="titulo_pagina"><div id="matrix_titulo" style="float:left">Agenda</div></td>
    </tr>
	<%
		dim rs,sql,id, criado, x
		select case request.QueryString("op")
		case "cad_agenda"
		dim titulo_agenda, texto_agenda, data_agenda, dest_agenda
		'//Nova Area ----------------------------------
		if lcase(request.ServerVariables("REQUEST_METHOD")) <> "post" then
			
			id = formus("id",3)
			data_agenda =  fdate(now(),"dd/mm/yyyy")
			dest_agenda = 0
			
			if id<>"" then
				subTitle("> Alterar Agendamento")
				'//Alterar Dados
				db "leitura",rs,sql,"select top 1 * from tb_agenda where cod_agenda="&id,conn
				titulo_agenda = trim(rs("titulo_agenda"))
				texto_agenda = trim(rs("texto_agenda"))
				data_agenda = fdate(rs("data_agenda"),"dd/mm/yyyy")
				dest_agenda = trim(rs("dest_agenda"))
				
			else
				subTitle("> Novo Agendamento")
				'//nova area
				db "leitura",rs,sql,"select top 1 * from tb_agenda where criado=0",conn
				if rs.eof and rs.bof then
					'//não existe registro
					db "normal",,sql,"insert into tb_agenda(criado) values(0)",conn
					db "leitura",rs,sql,"select top 1 * from tb_agenda where criado=0",conn
				end if						
				
			end if
			
			id = rs("cod_agenda")
			
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
		AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?aspid=' + AspId)"
		AspMenuOp(1) = "Salvar|3|miniMenu_agenda()"
		'//----------------------------------------
		AspMenu() 'gera menu
		'// Gera Menus ---------------------------- >
		%>
</td>
    </tr>
    <tr>
      <td id="conteudo_page"><span class="texto_size_10">Cod</span><br />
      <input name="cod_agenda" type="text" id="cod_agenda" value="<%=id%>" size="3" readonly="true" />
      <br />
      <span class="texto_size_10">Titulo</span><br />
<input name="titulo_agenda" type="text" id="titulo_agenda" value="<%=titulo_agenda%>" size="30" maxlength="255" />
<br />
<span class="texto_size_10">Data</span><br />
<input name="data_agenda" onBlur="if(this.value==''){this.value='<%=data_agenda%>'}else{dataVal(this)}" onFocus="this.value=''" type="text" id="data_agenda" value="<%=data_agenda%>" size="10" maxlength="10" />
<br />
<span class="texto_size_10">Texto</span><br />
<textarea id="ta" name="ta" style="width:99%" rows="25" cols="80"><%=texto_agenda%></textarea>
<br />
<span class="texto_size_10">Destacar</span><br />
<input name="dest_agenda"  type="radio" value="1" <%if dest_agenda = 1 then%>checked="checked"<% end if %> />
<span class="texto_size_10">
Sim 
<input name="dest_agenda" type="radio" value="0" <%if dest_agenda = 0 then%>checked="checked"<% end if %> />
N&atilde;o</span>
		<script>
	if(isIE()){
		oDateMasks = new Mask("dd/mm/yyyy", "date");
		oDateMasks.attach(document.form.data_agenda);
		
	}
	</script>
</td>
    </tr>
    <%
		else
		
			'///--------- [salva dados] ---->
			'db "leitura",rs,sql,"select * from album_area where id<>"&id&" and area='"&area&"'",conn
			'verifica se existe algo cadastrado com este Nome
				'if rs.eof and rs.bof then 
				
					db "normal",,sql,"update tb_agenda set titulo_agenda='"&formus("titulo_agenda",0)&"',dest_agenda="&formus("dest_agenda",0)&",data_agenda='"&fdate(formus("data_agenda",0),"yyyymmdd")&"',texto_agenda='"&formus("ta",0)&"', criado=1 where cod_agenda="&formus("cod_agenda",0),conn
					'//Informação salva
					alert("Área salva com sucesso!")
					
					'if criado = 0 then
						voltar request.ServerVariables("SCRIPT_NAME")&"?aspid="&aspid&"&op=cad_agenda","","_self"
					'else
						'voltar request.ServerVariables("SCRIPT_NAME")&"?aspid="&aspid&"&op=cad_area&id="&id,"","_self"
					'end if
					
				'else
					'// este registro já existe
					'alert("Está área já foi cadastrada!")
					'voltar request.ServerVariables("SCRIPT_NAME")&"?aspid="&aspid&"&op=cad_area&id="&id,"","_self"
										
				'end if		
			
		end if
		
		case "list_agenda"
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
		redim AspMenuOp(2) 'Define o numero de itens
		'//campos ---------------------------------
		AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?aspid=' + AspId)"
		AspMenuOp(1) = "Novo|4|WindowOpen('default.asp?op=cad_agenda&aspid=' + AspId)"
		'//----------------------------------------
		AspMenu() 'gera menu
		'// Gera Menus ---------------------------- >
		%>
	  </td>
    </tr>
    <tr>
      <td id="conteudo_page"><span class="texto_size_10">Selecione o Peri&oacute;do</span><br />
        <select name="periodo" id="periodo">
		<%
			db "leitura",rs,sql,"select * from AgendaPeriodo",conn
			while not rs.eof
		%>
          <option value="<%="01/"&rs("mes")&"/"&rs("ano")%>"><%=mesext(rs("mes"))%>/<%=rs("ano")%></option>
		 <%
		 		rs.movenext
			wend
		 %>
        </select> 
	    <input type="button" name="Button1" value="Abrir" onclick="WindowOpen('<%=request.ServerVariables("SCRIPT_NAME")%>?aspid=<%=aspid%>&op=list_agenda&periodo='+document.form.periodo.value)" />
	    <div align="center">
	    <%
		  subTitle("> Listar Agenda")
	
		  Dim GridData(3), GridName(3), GridSQL, GridDate, GridPoss
				
				  if formus("periodo",3) <> "" then	
					GridSQL = "select * from AgendaMemoResum where mes="&fdate(formus("periodo",3),"mm")&" and ano="&fdate(formus("periodo",3),"yyyy")&";"
					
					GridName(0) = "Data;80"
					GridName(1) = "Titulo;370"
					GridName(2) = "Alterar;60"
					GridName(3) = "Deletar;60"
									
					GridData(0) = "data_agenda;1;0"
					GridData(1) = "titulo_agenda;0;0"
					GridData(2) = "*Alterar;1;default.asp?op=cad_agenda&id={cod_agenda}&aspid=[aspid]&m=[mod]"
					GridData(3) = "*Deletar;1;javascript: deletar('"&request.ServerVariables("SCRIPT_NAME")&"?op=del&aspid=[aspid]&id={cod_agenda}&m=[mod]&periodo={data_agenda}','{titulo_agenda}')"
							
					GridDate = "dd/mm/yyyy"
					GridPoss = "center"
					'GridURLS = 1
					'// Chama a Função
							
					GridForm Gkeys,15
				end if
					
	%>
	    </div>	  </td>
    </tr>
	<%
		case "del"
			
			db "normal",,sql,"delete from tb_agenda where cod_agenda = "&formus("id",3),conn
			voltar request.ServerVariables("SCRIPT_NAME")&"?aspid="&aspid&"&op=list_agenda&periodo="&formus("periodo",3),"","_self"
			
		case else
		'//Menu Principal ------------------------------------
	%>
	<tr>
      <td id="menutab">&nbsp;</td>
    </tr>
    <tr>
      <td id="conteudo_page"><p class="texto_size_10">Selecione qual o tipo de informa&ccedil;&atilde;o deseja visualizar:</p>
        <ul class="texto_size_10">
          <li><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=cad_agenda&aspid=<%=aspid%>" target="_self">Criar novo Agendamento </a></li>
          <li><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=list_agenda&aspid=<%=aspid%>" target="_self">Listar Agenda </a></li>
          <li><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=cad_album&aspid=<%=aspid%>" target="_self">Visualizar Calend&aacute;rio </a></li>
        </ul>
      <p>&nbsp;</p></td>
    </tr>
	<%
		end select
	%>
  </table>
  <br />
 <% if request("op")<>"list_foto" then %>
</form>
<% end if %>
		
</body>
</html>
