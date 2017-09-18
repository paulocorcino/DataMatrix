<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp"-->
<%
	protec("2")
	cache 1,"true"
	dim rs2,sql2
%>
<html>
<head>
<title>Infomativo</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/estilo.css" rel="stylesheet" type="text/css">
<%
	function image(t)
		image = t
		image = replace(t,fotog_root,sites&fotog_root)
	end function
	'alert(sites&fotog_root)
		if isempty(request.form("submit")) and request.QueryString("op")="cad" then
			if isempty(request.form("submit")) then

		if request.QueryString("id")<>"" then
			db "leitura", rs,sql,"select id, titulo, conteudo from matrix_info where id="&formus("id",3),conn
			
			id = rs("id")
			conteud = rs("conteudo")
			titul = trim(rs("titulo"))
			
		else
			
			db  "leitura",rs,sql,"select id from matrix_info where criado=0",conn
			
			if rs.eof and rs.bof then
				db "normal",,sql,"insert into matrix_info(criado) values (0)",conn
				db "leitura",rs,sql,"select id from matrix_info where criado=0",conn
			end if
			
			id = rs("id")
			
		end if
		end if
%>
<script type="text/javascript">
  _editor_url = "../editor/";
  _editor_lang = "en";
</script>
<script type="text/javascript" src="../editor/htmlarea.asp?tela=info&id=<%=id%>"></script>

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
	dim rs, sql, x, id, conteud, titul
	select case request.QueryString("op")
	case "cad"
		if isempty(request.form("submit")) then


%>
<form name="form1" method="post" action="">
  <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="2">
    <tr>
      <td height="20" class="fundo1"><strong class="fonte"> &nbsp;Editar Inform&aacute;tivo </strong></td>
    </tr>
    <tr>
      <td height="20"><%=erros%></td>
    </tr>
    <tr>
      <td height="20" class="fonte"><strong>Titulo do Informativo</strong><br>
        (Digite o Titulo do informativo)</td>
    </tr>
    <tr>
      <td height="20"><input name="info" type="text" class="formtextarea" id="info" size="50" value="<%=titul%>">
      <br>
      <br>
      <span class="fonte"><strong>Texto Informativo</strong><br>
      (Digite o texto do informativo)</span> </td>
    </tr>
    <tr>
      <td height="100%"><textarea id="ta" name="ta" style="width:100%;height=80%" rows="25" cols="80"><%=conteud%></textarea></td>
    </tr>
    <tr><input type="hidden" name="id" value="<%=id%>">
      <td height="20"><input name="Submit" type="submit" class="botao" value="Salvar"></td>
    </tr>
    <tr>
      <td height="20">&nbsp;</td>
    </tr>
  </table>
</form>

<%
		elseif not isempty(request.Form("submit")) and trim(formus("info",1))<>"" then
			db "normal",,sql,"update matrix_info set titulo='"&formus("info",1)&"',conteudo='"&image(replace(request.form("ta"),"'","&#39"))&"',criado=1 where id="&formus("id",1),conn
			voltar request.ServerVariables("SCRIPT_NAME")&"?op=list&id="&formus("id",1)&"&aspid="&aspid(),"","_self"
		else
			voltar request.ServerVariables("SCRIPT_NAME")&"?op=cad&id="&formus("id",1)&"&aspid="&aspid(),"Todos os campos devem estar preenchidos","_self"
		end if
	case "list"
	prep_deletar()
%>
<table width="100%"  border="0" cellspacing="2" cellpadding=" 0">
  <tr>
    <td class="fundo1">&nbsp;<strong class="fonte">Listar Informativos</strong></td>
  </tr>
  <tr>
    <td><%=erros%></td>
  </tr>
  <tr>
    <td><table width="70%"  border="0" align="center" cellpadding=" 0" cellspacing="2">
      <tr class="fundo2">
        <td width="30%"><div align="center" class="fonte"><strong>Informativo</strong></div></td>
        <td width="10%"><div align="center" class="fonte"><strong>Status</strong></div></td>
        <td width="10%"><div align="center" class="fonte"><strong>Total / Lidos</strong></div></td>
        <td width="10%"><div align="center" class="fonte"><strong>Enviar</strong></div></td>
        <td width="10%"><div align="center" class="fonte"><strong>Alterar</strong></div></td>
        <td width="10%"><div align="center" class="fonte"><strong>Deletar</strong></div></td>
      </tr>
	  <%
	  	db "leitura",rs,sql,"select id, titulo, enviado, criado from matrix_info where criado=1", conn
		dim tReg, lReg
		tReg = empty
		lReg = empty
	  	while not rs.eof
			db "leitura",rs2,sql2,"SELECT COUNT_BIG (id_i) AS total FROM dbo.matrix_info_stat WHERE (id_i = "&rs("id")&")",conn
			tReg = rs2("total")
			db "leitura",rs2,sql2,"SELECT COUNT_BIG (ok) AS lidos FROM dbo.matrix_info_stat WHERE (ok = 1) AND (id_i = "&rs("id")&")",conn
			lReg = rs2("lidos")
	  %>
      <tr class="fonte">
        <td>&nbsp;<%=trim(rs("titulo"))%></td>
        <td><div align="center"><% if rs("enviado")=0 then %>Na Fila<%else%>Enviado<%end if%></div></td>
        <td><div align="center"><%=tReg%>/<%=lReg%></div></td>
        <td><div align="center"><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=enviar&id=<%=rs("id")%>&aspid=<%=aspid()%>">Enviar</a></div></td>
        <td><div align="center"><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=cad&id=<%=rs("id")%>&aspid=<%=aspid()%>">Alterar</a></div></td>
        <td><div align="center"><a href="<%deletar request.ServerVariables("SCRIPT_NAME")&"?op=del&id="&rs("id")&"&aspid="&aspid(),trim(rs("titulo"))%>">Deletar</a></div></td>
      </tr>
	  <%
	  	rs.movenext
	wend
	if rs.eof and rs.bof then
	  %>
      <tr>
        <td colspan="6"><div align="center" class="fonte">Nenhum informativo</div></td>
        </tr>
		<%	
			end if
		%>
    </table></td>
  </tr>
</table>
<%
	case "del"
		db "leitura",rs,sql,"select id from matrix_info where id="&formus("id",3),conn
		if not rs.eof and not rs.bof then
			db "normal", rs,sql,"delete from matrinx_info where id="&formus("id",3),conn
		end if
		voltar request.ServerVariables("SCRIPT_NAME")&"?op=list&id="&formus("id",3)&"&aspid="&aspid(),"","_self"
	case "enviar"
	
	dim strsql,objRS,objConn,iMaxRecords, iRow, iCol, iRecordsPerPage, iTd, iBlk, iLoop, iTw, iTh
	
	'How many records per page do we want to show?
	iRecordsPerPage = 10


	Dim currentPage	 'what page are we on??
	Dim bolLastPage	 'are we on the last page?
	
	if len(Request.QueryString("page")) = 0 then
		currentPage = 1
	else
		currentPage = CInt(Request.QueryString("page"))
	end if

	'Show the paged results
	'strSQL = "sp_PagedItems " & currentPage & "," & iRecordsPerPage &","& request.QueryString("id")
	'objRS.Open strSQL, objConn
	db "leitura",objRS,strSQL,"sp_EmailSend " & currentPage & "," & iRecordsPerPage ,conn
	
	'alert(objRS("TotalF"))
	 
	'See if we're on the last page
	if Not objRS.EOF then
		if CInt(objRS("MoreRecords")) > 0 then
			bolLastPage = False
		else
			bolLastPage = True
		end if
	end if
	
	dim iTot,msg
	'if not objRS.eof then
		iTot = cdbl(objRS("MoreRecords")/iRecordsPerPage)
		if cdbl(iTot) > cint(iTot) then
			iTot = cint(iTot) + 1
		end if			
	'else
		'iTot = 1
	'end if
	'-------------- Criando a tabela de imagens --------------------
	db "leitura",rs2,sql2,"select id,titulo,conteudo from matrix_info where id="&formus("id",3),conn
	Do While Not objRS.EOF 
	
'aquiiiiiiiii
		db "normal",,sql,"insert into matrix_info_stat(id_i,ok) values ("&formus("id",3)&",0)",conn
		db "leitura",rs,sql,"select uid from matrix_info_stat where id_i="&formus("id",3)&" order by id desc",conn
		msg = "<br><br><img src="""&site&"/matrix/info/count.asp?uid="&rs("uid")&""" width=""1"" height=""1"">"
		emails "sac@tonazoeira.com.br",trim(objRS("emails")),"Inforzoeira - "& trim(rs2("titulo")),rs2("conteudo")&msg,"html"
	
		
		
		objRS.MoveNext
	Loop 
	
	

	'-------------- Criando a tabela de imagens --------------------
	
	%>
	 <%if Not bolLastPage then %>
	<script>
var tempo = 10
var militempo = 1000
var x = tempo * militempo
window.setTimeout("pular()",x)
function pular(){
window.open("<%=request.ServerVariables("SCRIPT_NAME")%>?op=enviar&page=<%=currentPage+1%>&id=<%=request.querystring("id")%>&aspid=<%=aspid()%>","_self")
}
</script>
<%end if%>
	<table width="100%"  border="0" cellspacing="2" cellpadding=" 0">
      <tr>
        <td class="fundo1">&nbsp;<strong class="fonte">Enviando</strong></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="296"><div align="center" class="fonte">
		 <%if Not bolLastPage then %>
          <p>Enviando os Informativos aguarde... Faltam <%=iTot%> grupos<br>Tempo Estimado <%=cint(iTot*10)%>sec</p>
		  <%else
			  db "normal",,sql,"update matrix_info set enviado=1 where id="&formus("id",3),conn
		  %>
          <p>Informativo enviado com sucesso!</p>
          <p><a href="<%=request.ServerVariables("SCRIPT_NAME")&"?op=list&id="&formus("id",3)&"&aspid="&aspid()%>">Voltar</a></p>
		  <%end if%>
        </div></td>
      </tr>
    </table>
<%
	end select
%>
</body>
</html>
