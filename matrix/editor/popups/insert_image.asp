<!--#include virtual="/bymidia/matrix/config/config.asp"-->
<%
	cache 1,"true"
%>
<html>

<head>
  <title>Inserir Imagem</title>

<script type="text/javascript" src="popup.js"></script>

<script type="text/javascript">

window.resizeTo(400, 150);

function Init() {
  __dlg_init();
  var param = window.dialogArguments;
  if (param) {
  
      document.getElementById("f_url").value = param["f_url"];
      document.getElementById("f_alt").value = param["f_alt"];
	   document.getElementById("f_float").value = param["f_float"];
      document.getElementById("f_border").value = param["f_border"];
      document.getElementById("f_align").value = param["f_align"];
      document.getElementById("f_vert").value = param["f_vert"];
      document.getElementById("f_horiz").value = param["f_horiz"];
      window.ipreview.location.replace(param.f_url);
  }
  document.getElementById("f_url").focus();
};

function onOK() {
  var required = {
    "f_url": "Voc� deve escolher uma URL"
  };
  for (var i in required) {
    var el = document.getElementById(i);
    if (!el.value) {
      alert(required[i]);
      el.focus();
      return false;
    }
  }
  // pass data back to the calling window
  var fields = ["f_url", "f_alt", "f_align", "f_border", "f_float",
                "f_horiz", "f_vert"];
  var param = new Object();
  for (var i in fields) {
    var id = fields[i];
    var el = document.getElementById(id);
    param[id] = el.value;
  }
  __dlg_close(param);
  return false;
};

function onCancel() {
  __dlg_close(null);
  return false;
};

function onPreview() {
  var f_url = document.getElementById("f_url");
  var url = f_url.value;
  if (!url) {
    alert("Voc� deve inserir uma URL primeiro");
    f_url.focus();
    return false;
  }
  window.ipreview.location.replace(url);
  return false;
};
</script>

<style type="text/css">
html, body {
  background: ButtonFace;
  color: ButtonText;
  font: 11px Tahoma,Verdana,sans-serif;
  margin: 0px;
  padding: 0px;
}
body { padding: 5px; }
table {
  font: 11px Tahoma,Verdana,sans-serif;
}
form p {
  margin-top: 5px;
  margin-bottom: 5px;
}
.fl { width: 7em; float: left; padding: 2px 5px; text-align: right; }
.fr { width: 6em; float: left; padding: 2px 5px; text-align: right; }
fieldset { padding: 0px 10px 5px 5px; }
select, input, button { font: 11px Tahoma,Verdana,sans-serif; }
button { width: 70px; }
.space { padding: 2px; }

.title { background: #ddf; color: #000; font-weight: bold; font-size: 120%; padding: 3px 10px; margin-bottom: 10px;
border-bottom: 1px solid black; letter-spacing: 2px;
}
form { padding: 0px; margin: 0px; }
.atual {
	font-weight: bold;
	color: #000000;
	background-color: #FFCC00;
}
.mais {
	font-weight: bold;
	color: #000000;
	background-color: #0099CC;
}
</style>

</head>

<body onload="Init()">

<div class="title">Inserir Imagem</div>
<form action="" method="get">
<table border="0" width="100%" style="padding: 0px; margin: 0px">
  <tbody>

  <tr>
    <td style="width: 7em; text-align: right">URL Imagem :</td>
    <td><!--input type="text" name="url" id="f_url" style="width:75%"
      title="Enter the image URL here" /-->
	  <select name="url" id="f_url">
	    <option value="" selected class="atual">Imagens de <%=request.QueryString("area")%></option>
	<%
		dim rs,sql
		db "leitura",rs,sql,"SELECT id_a, area, foto, nome FROM dbo.myimage WHERE (id_a = "&formus("id",3)&") AND (area LIKE N'"&formus("area",3)&"')",conn
		while not rs.eof
	%>
	    <option value="<%=sites&fotog_root&rs("foto")%>"><%=rs("nome")%></option>
	<%
			rs.movenext
		wend
		if rs.eof and rs.bof then
	%>
	    <option value="">Esta �rea n�o cont�m imagem</option>
	<%
		end if
	%>
	<option class="mais" value="">+ Mais Imagens</option>
	<%
		 
		db "leitura",rs,sql,"SELECT area, foto, nome FROM dbo.myimage",conn
		while not rs.eof
	%>
	   <option value="<%=sites&fotog_root&rs("foto")%>"><%=rs("nome")%></option>
	 <%
	 		rs.movenext
		wend
			if rs.eof and rs.bof then
	%>
	    <option>Nenuma Imagem cadastrada</option>
	<%
		end if
	%>
	    </select>
      <button name="preview" onclick="return onPreview();"
      title="Preview the image in a new window">Visualizar</button>
	  <button name="import" onclick="window.open('../../myimage/myimage.asp?area=<%=request.QueryString("area")%>&id=<%=request.QueryString("id")%>','_self')"
      title="Envie imagem">Importar</button>
    </td>
  </tr>
  <tr>
    <td style="width: 7em; text-align: right">Texto Alter.:</td>
    <td><input type="text" name="alt" id="f_alt" style="width:100%"
      title="For browsers that don't support images" /></td>
  </tr>

  </tbody>
</table>

<p />

<fieldset style="float: left; margin-left: 5px;">
<legend>Layout</legend>

<div class="space"></div>

<div class="fl">Alinhamento:</div>
<p>
  <select size="1" name="align" id="f_align"
  title="Positioning of this image">
    <option value="" selected>N&atilde;o Alterar</option>
    <option value="left">Esquerda</option>
    <option value="right">Direita</option>
    <option value="texttop">Texto no Topo</option>
    <option value="absmiddle">Exatam. Meio</option>
    <option value="baseline">Linha Base</option>
    <option value="absbottom">Exatam. Embaixo</option>
    <option value="bottom">Embaixo</option>
    <option value="middle">Meio</option>
    <option value="top">Topo</option>
  </select>
  </p>
<div class="fl">Flutuar: </div>
    <select name="float" id="f_float">
      <option value="none" selected>Nada</option>
      <option value="right; margin: 3px;">Esquerda</option>
      <option value="left; margin: 3px;">Direita</option>
    </select>
  

<p></p>
<div class="fl">Borda:</div>
<input type="text" name="border" id="f_border" size="5"
title="Leave empty for no border" />

<div class="space"></div>

</fieldset>

<fieldset style="float:right; margin-right: 5px;">
<legend>Espa&ccedil;amento</legend>

<div class="space"></div>

<div class="fr">Horizontal:</div>
<input type="text" name="horiz" id="f_horiz" size="3"
title="Horizontal padding" />

<p />

<div class="fr">Vertical:</div>
<input type="text" name="vert" id="f_vert" size="3"
title="Vertical padding" />

<div class="space"></div>

</fieldset>
<br clear="all" />
<table width="100%" style="margin-bottom: 0.2em">
 <tr>
  <td valign="bottom">
    Vizualiza&ccedil;&atilde;o de Imagem:<br />
    <iframe name="ipreview" id="ipreview" frameborder="0" style="border : 1px solid gray;" height="200" width="300" src=""></iframe>
  </td>
  <td valign="bottom" style="text-align: right">
    <button type="button" name="ok" onclick="return onOK();">OK</button><br>
    <button type="button" name="cancel" onclick="return onCancel();">Cancelar</button>
  </td>
 </tr>
</table>
</form>
</body>
</html>
