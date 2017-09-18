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
</head>

<body>
<div id="aAlbumThumb">&nbsp;</div>
<div id="aAlbumVoltar">&nbsp;</div>
<div id="aAlbumAvancar">&nbsp;</div>	
<div id="aAlbumListPage">&nbsp;</div>		
<div id="aAlbumNomeFoto">&nbsp;</div>	
<div id="aAlbumClick">&nbsp;</div>	
<div id="aAlbumEnviar">&nbsp;</div>
<div id="aAlbumShowFoto">&nbsp;</div>

<%
	imgAlbumVoltar = "Voltare"
	imgAlbumAvanca = "Avançare"
	imgAlbumVoltarIn = "Voltare"
	imgAlbumAvancaIn = "Avançare"
	imgAlbumEnviar = "EnViare Fotos"
	imgAlbumNumberItens = 14
	imgAlbumNumberColum = 7
	imgAlbumFotoDestaq = 1
	imgAlbumMiniComent = 1
	imgAlbumComentStart = "Iniciar Album"
	imgAlbumComentOrdem = "start,data,album,local,evento,fotografo,start"
	AlbumId = 1
	AlbumW = 0
	AlbumH = 0

	aAlbum()
%>

</body>
</html>
