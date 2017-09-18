<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp"-->
<%
	protec("2")
	cache 1,"true"
%>
<html>
<head>
<title>Conteudo</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/estilo.css" rel="stylesheet" type="text/css">
<%
	if isempty(request.form("submit")) and request.QueryString("op")="cad" then
	
		if request.QueryString("id")<>"" then
	 	db "leitura",rs,sql,"select id,sessao,texto from matrix_cont where id="&formus("id",3),conn
		er(144)
	 else
	 	db "leitura",rs,sql,"select id,sessao,texto from matrix_cont where criado=0",conn
		er(135)
		if rs.eof and rs.bof then
			db "normal",,sql,"insert into dbo.matrix_cont(criado) values(0)",conn
			er(138)
			db "leitura",rs,sql,"select id,sessao,texto from matrix_cont where criado=0",conn
			er(140)
		end if
	 end if
%>
<script type="text/javascript">
  _editor_url = "../editor/";
  _editor_lang = "en";
</script>
<script type="text/javascript" src="../editor/htmlarea.asp?tela=conteudo&id=<%=rs("id")%>"></script>

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

</head>

<body onload="initEditor()">
<% else %>
</head>

<body>

<% end if %>
<%
dim rs,sql,x
select case formus("op",3) 
case "cad"

if isempty(request.form("submit")) then


%>
<form name="form1" method="post" action="<%=request.ServerVariables("SCRIPT_NAME")%>?op=cad&aspid=<%=aspid()%>">
  <table width="98%" border="0" align="center" cellpadding="0" cellspacing="2">
    <tr>
      <td class="fundo1">&nbsp;<span class="fonte"><strong>Editar Conte&uacute;do </strong></span></td>
    </tr>
    <tr>
      <td><%=erros%></td>
    </tr>
    <tr>
      <td class="fonte"><strong> <br>
        &nbsp;Sess&atilde;o<br>
        </strong>(Edite no nome da Sess&atilde;o)<strong><br>
        </strong></td>
    </tr>
    <tr>
      <td>&nbsp;<span class="fonte">
	  <% if nivelaspid<>1 then%>
	  <%=rs("sessao")%>&nbsp; <input type="hidden"  name="sessao" value="<%=rs("sessao")%>">
	  <% else %>
        <input name="sessao" type="text" class="formtextarea" id="titulo" size="50" value="<%=rs("sessao")%>">
		<% end if %>
        <br>
        <br>
</span></td>
    </tr>
    <tr>
      <td><strong>&nbsp;<span class="fonte">Editor de Texto <br>
        </span></strong><span class="fonte">(Edite o conte&uacute;do  do site ) </span><strong><span class="fonte"><br>
        </span></strong></td>
    </tr>
    <tr>
      <td><table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="bordasa">
  <tr>
    <td><textarea id="ta" name="ta" style="width:100%" rows="25" cols="80"><%=rs("texto")%></textarea></td>
  </tr>
</table>
</td>
    </tr>
    <tr>
      <td><div align="left">
    
  &nbsp;<input name="id" type="hidden" value="<%=rs("id")%>">
    <input name="Submit" type="submit" class="botao" value="Salvar">
    
  <input name="Submit2" type="button" class="botao" value="Cancelar" onClick="window.open('<%=request.ServerVariables("SCRIPT_NAME")%>?op=listar&aspid=<%=aspid()%>','_self')">
      </div></td>
    </tr>
  </table>
</form>
<%
elseif not isempty(request.form("submit")) and trim(request.form("ta"))<>"" then
		
		db "normal",,sql,"update dbo.matrix_cont set sessao='"&formus("sessao",1)&"', texto='"&replace(request.form("ta"),"'","&#39")&"', criado=1 where id="&formus("id",1),conn
		er(226)
		voltar request.ServerVariables("SCRIPT_NAME")&"?op=listar&id="&formus("id",1)&"&aspid="&aspid(),"","_self"
else
	voltar request.ServerVariables("SCRIPT_NAME")&"?op=cad&id="&formus("id",1)&"&aspid="&aspid(),"É necessário preencher com algum texto.","_self"
end if

case "listar"
prep_deletar()
%>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="2">
  <tr>
    <td class="fundo1">&nbsp;<span class="fonte"><strong>Listar Conteudos</strong></span></td>
  </tr>
  <tr>
    <td class="fonte"><% if nivelaspid=1 then %>[ <a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=cad&aspid=<%=aspid()%>">Nova Sess&atilde;o</a> ] <%end if%>&nbsp;</td>
  </tr>
  <tr>
    <td><table width="80%" border="0" align="center" cellpadding="0" cellspacing="1">
      <tr class="fundo1">
        <td width="6%"><div align="center" class="fonte"><strong>Id</strong></div></td>
        <td width="65%"><div align="center" class="fonte"><strong>Sess&atilde;o</strong></div></td>
        <td width="15%"><div align="center" class="fonte"><strong>Alterar</strong></div></td>
        <td width="14%"><div align="center" class="fonte"><strong>Deletar</strong></div></td>
      </tr>
	  <%
	  	db "leitura",rs,sql,"select id,sessao from matrix_cont where criado=1",conn
		er(236)
		while not rs.eof 
	  %>
	     <tr class="fonte">
	    <td><div align="center"><%=rs("id")%></div></td>
        <td><div align="left">&nbsp;&nbsp;<%=rs("sessao")%></div></td>
        <td><div align="center"><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=cad&aspid=<%=aspid()%>&id=<%=rs("id")%>">Alterar</a></div></td>
        <td><div align="center"><% if nivelaspid<>1 then %>Deletar<%else%><a href="<%  deletar request.ServerVariables("SCRIPT_NAME")&"?op=del&aspid="&aspid()&"&id="&rs("id"),rs("sessao")%>">Deletar</a><%end if%></div></td>
      </tr>
	  <%
	  		rs.movenext
		wend
		er(244)
		if rs.eof and rs.bof then
	%>
      <tr>
        <td colspan="4"><div align="center" class="fonte"><strong>Nenhuma Sess&atilde;o Cadastrada</strong></div></td>
        </tr>
	<%
	end if
%>
    </table></td>
  </tr>
</table>
<%
fdb(rs)
case "del"

	db "leitura",rs,sql,"select id from matrix_cont where id="&formus("id",3),conn
	er(259)
	if rs.eof and rs.bof then
		voltar request.ServerVariables("SCRIPT_NAME")&"?op=listar&id="&formus("id",1)&"&aspid="&aspid(),"","_self"
	else
		db "normal",,sql,"delete from matrix_cont where id="&formus("id",3),conn
		er(264)
		voltar request.ServerVariables("SCRIPT_NAME")&"?op=listar&id="&formus("id",1)&"&aspid="&aspid(),"","_self"
	end if
	fdb(rs)
end select
fdb(conn)
%>
</body>
</html>
