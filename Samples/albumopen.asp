<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<%
	'Cache
	cache 1, true
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<link href="albumestilo.css" rel="stylesheet" type="text/css" />
</head>

<body>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td colspan="3"><div id="aAlbumShowFoto">&nbsp;</div></td>
  </tr>
  <tr>
    <td width="33%"><div id="aAlbumNomeFoto">&nbsp;nome foto </div></td>
    <td width="33%"><div id="aAlbumClick" style="display:none">cliques</div></td>
    <td width="33%"><div id="aAlbumEnviar">&nbsp;<a href="#">enviar para amigo</a> </div></td>
  </tr>
  <tr>
    <td colspan="3"><div id="aAlbumThumb">&nbsp;</div></td>
  </tr>
  <tr>
    <td><div id="aAlbumVoltar">&lt;&nbsp;voltar</div></td>
    <td><div id="aAlbumListPage">&nbsp;</div></td>
    <td><div id="aAlbumAvancar">&nbsp;Avan&ccedil;ar &gt; </div></td>
  </tr>
</table>
<%
	imgAlbumVoltar = "Voltar"
	imgAlbumAvanca = "Avançar"
	imgAlbumVoltarIn = "Voltar"
	imgAlbumAvancaIn = "Avançar"
	imgAlbumEnviar = "Enviar Foto"
	imgAlbumNumberItens = 8
	imgAlbumNumberColum = 4
	imgAlbumFotoDestaq = 1
	imgAlbumMiniComent = 1
	imgAlbumComentStart = "Iniciar%20Album"
	imgAlbumComentOrdem = "data,album,local,evento,fotografo,start"
	AlbumId = 1
	AlbumW = 88
	AlbumH = 65

	aAlbum()
%>

</body>
</html>
