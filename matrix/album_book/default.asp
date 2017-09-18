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
<title>Album Book</title>
<link href="../css/estilos.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" type="text/javascript" src="../js/script.js"></script>
<script language="JavaScript" type="text/javascript" src="../js/form_album_book.js"></script>
<script language="JavaScript" type="text/javascript" src="../js/form.js"></script>
<script>AspId = '<%=aspid%>'</script>

</head>
<!--[Info File]-->
<!--[!]
	Album Book
	Criado: Paulo Corcino		
	@Ano: 2006 - @Version: 1.5
[!]-->
<!--[Info File]-->
<body>
<form method="post" name="form" id="form">
  <table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td id="titulo_pagina"><div id="matrix_titulo" style="float:left">Album Book</div></td>
    </tr>
	<%
		dim rs,sql,id, criado, x
		select case request.QueryString("op")
		case "cad_area"
		dim area
		'//Nova Area ----------------------------------
		if lcase(request.ServerVariables("REQUEST_METHOD")) <> "post" then
			
			id = formus("id",3)
			
			if id<>"" then
				subTitle("> Alterar Área")
				'//Alterar Dados
				db "leitura",rs,sql,"select top 1 * from album_area where id="&id,conn
				area = trim(rs("area"))
				
			else
				subTitle("> Nova Área")
				'//nova area
				db "leitura",rs,sql,"select top 1 * from album_area where criado=0",conn
				if rs.eof and rs.bof then
					'//não existe registro
					db "normal",,sql,"insert into album_area(criado) values(0)",conn
					db "leitura",rs,sql,"select top 1 * from album_area where criado=0",conn
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
		AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?aspid=' + AspId)"
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
			db "leitura",rs,sql,"select * from album_area where id<>"&id&" and area='"&area&"'",conn
			'verifica se existe algo cadastrado com este Nome
				if rs.eof and rs.bof then 
				
					db "normal",,sql,"update album_area set area='"&area&"', criado=1 where id="&id,conn
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
				
				
					GridSQL = "select * from album_area where criado=1;"
					
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
		case "cad_album"
		dim id_ar, album, dest, quando, aonde, quem, coment, fotografo, img_patrocinio
		'//Novo Album -------------------------------------
		'if lcase(request.ServerVariables("REQUEST_METHOD")) <> "post" then
			
			id = formus("id",3)
			
			if id<>"" then
				subTitle("> Alterar Álbum")
				'//Alterar Dados
				db "leitura",rs,sql,"select top 1 * from album_album where id="&id,conn
				
				id_ar = cint(trim(rs("id_ar")))
				album = trim(rs("album"))
				dest = trim(rs("dest"))
				quando = trim(rs("quando"))
				aonde = trim(rs("aonde"))
				quem = trim(rs("quem"))
				coment = trim(rs("coment"))
				fotografo = trim(rs("fotografo"))
				img_patrocinio = trim(rs("img_patrocinio"))
											
			else
				subTitle("> Novo Álbum")
				'//nova area
				
				db "leitura",rs,sql,"select top 1 * from album_album where criado=0",conn
				if rs.eof and rs.bof then
					'//não existe registro
					db "normal",,sql,"insert into album_album(criado) values(0)",conn
					db "leitura",rs,sql,"select top 1 * from album_album where criado=0",conn
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
		if id_ar <> "" then
			AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?op=list_album&id_ar="&id_ar&"&aspid=' + AspId)"
		else
			AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?aspid=' + AspId)"
		end if
		
		AspMenuOp(1) = "Salvar|3|miniMenu_album()"
		'//----------------------------------------
		AspMenu() 'gera menu
		'// Gera Menus ---------------------------- >
		%>
	  </td>
    </tr>
    <tr>
      <td id="conteudo_page"><table width="396" border="0" cellspacing="2" cellpadding="0">
            <tr>
              <td width="50" valign="top" class="texto_size_10">C&oacute;d<br />
                <input name="id" type="text" id="id" value="<%=id%>" size="3" /></td>
              <td valign="top" class="texto_size_10">&Aacute;rea<br />
                <select name="id_ar" id="id_ar">
				<%
					db "leitura",rs,sql,"select id, area from album_area where criado=1",conn
					while not rs.eof
						response.write "<option value='"&rs("id")&"' "
						
						if id_ar = rs("id") then
							response.write "selected" 
						end if 
						
						response.write ">"&trim(rs("area"))&"</option>"
						rs.movenext
					wend
				%>
                </select></td>
            </tr>
          </table>
        <span class="texto_size_10">Nome do &Aacute;lbum<br />
        <input name="album" type="text" id="album" value="<%=album%>" size="40" maxlength="255" />
        <br />
        Evento<br />
        <input name="quem" type="text" id="quem" value="<%=quem%>" size="40" maxlength="255" />
        <br />
        Data do Evento <br />
        <input name="quando" type="text" id="quando" value="<%=quando%>" size="30" maxlength="255" />
        <br />
        Local<br />
        <input name="aonde" type="text" id="aonde" value="<%=aonde%>" size="30" maxlength="255" />
        <br />
        Coment&aacute;rios<br />
        <textarea name="coment" cols="40" rows="5" onkeypress="CharMax(this,255)"  id="coment"><%=coment%></textarea>
        <br />
        Fotografo<br />
        <input name="fotografo" type="text" id="fotografo" value="<%=fotografo%>" size="30" maxlength="255" />
        <br />
        Patrocinador<br />
        </span>
		<%
			if img_patrocinio <> "" then
		%>
        <table width="300" border="0" cellspacing="3" cellpadding="0">
        <tr>
          <td width="20" class="texto_size_10">
            <input name="img_pat" type="radio" value="1" checked="checked" />
          </td>
          <td class="texto_size_10">&nbsp;Manter Imagem </td>
          <td width="20" class="texto_size_10">
            <input name="img_pat" type="radio" value="0" />
          </td>
          <td class="texto_size_10">Remover Imagem </td>
        </tr>
      </table>
	  <%
	  		end if
	  %>
        <span class="texto_size_10">
        <input type="file" name="img_patrocinio" onchange="previewImg(document.form.img_patrocinio,'foto_loja_select','gif')"/>
        <input name="criado" type="hidden" id="criado" value="<%=criado%>" />
		<br />
		<div id="foto_loja_select">&nbsp;</div>
		<%
			if img_patrocinio <> "" then
		%>
			<script> document.getElementById("foto_loja_select").innerHTML = "<img src='<%=pImg&img_patrocinio%>'>";</script>
		<%
			end if
		%>
      </span></td>
    </tr>
	<%
		case "list_album"
		
		'//Listar Album ------------------------------------
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
      <td id="conteudo_page"><span class="texto_size_10">Selecione &Aacute;rea</span><br />
        <select name="id_ar" id="id_ar">
		<%
			id_ar = formus("id_ar",3)
			
			db "leitura",rs,sql,"select id, area from album_area where criado=1",conn
			while not rs.eof
				response.write "<option value='"&rs("id")&"' "
				
				if cstr(id_ar) = cstr(rs("id")) then
					response.write "selected" 
				end if 
				
				response.write ">"&trim(rs("area"))&"</option>"
				rs.movenext
			wend
		%>
        </select> <input name="aspid" type="hidden" value="<%=aspid%>" /><input name="op" type="hidden" value="list_album" />
	  	   <input name="Button" type="button" value="Abrir" onclick="area_abrir()"/>
	  	   <div align="center">
	    <%
		  subTitle("> Listar Álbum")
		  reDim GridData(7), GridName(7)
				
				
					GridSQL = "SELECT TOP 100 PERCENT dbo.album_album.id, dbo.album_album.album, dbo.album_album.data, dbo.total_fotos.Total_foto, dbo.total_fotos.clicks, dbo.album_album.id_ar FROM dbo.album_album  LEFT OUTER JOIN dbo.total_fotos ON dbo.album_album.id = dbo.total_fotos.id_a WHERE     (dbo.album_album.id_ar = "&request.QueryString("id_ar")&") and criado=1 order by id desc;"
					
					GridName(0) = "Cod;35"
					GridName(1) = "Album;370"
					GridName(2) = "Fotos;40"
					GridName(3) = "Publicação;75"
					GridName(4) = "Clicks;60"
					GridName(5) = "Alterar;60"
					GridName(6) = "Editar Fotos;80"
					GridName(7) = "Deletar;60"
									
					GridData(0) = "id;1;0"
					GridData(1) = "album;0;0"
					GridData(2) = "total_foto;1;0"
					GridData(3) = "data;1;0"
					GridData(4) = "clicks;1;0"
					GridData(5) = "*Alterar;1;default.asp?op=cad_album&id=[col0]&aspid=[aspid]"
					GridData(6) = "*Editar Foto;1;default.asp?op=list_foto&id_ar={id_ar}&id=[col0]&aspid=[aspid]"
					GridData(7) = "*Deletar;1;javascript: deletar('"&request.ServerVariables("SCRIPT_NAME")&"?op=del&tp=album&aspid=[aspid]&id=[col0]','[col1]')"
							
					GridDate = "dd/mm/yyyy"
					GridPoss = "center"
					'GridURLS = 1
					'// Chama a Função
					
					if id_ar <> "" then
						GridForm Gkeys,15
					end if
					
	%>
	    </div>	  </td>
    </tr>
	<%
		case "send_foto_tp1"
		'//Enviar Foto Tipo 1 ------------------------------------
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
		AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?op=list_foto&id_ar="&formus("id_ar",3)&"&id="&formus("id",3)&"&aspid=' + AspId)"
		AspMenuOp(1) = "Salvar|3|miniMenu_enviar()"
		'//----------------------------------------
		AspMenu() 'gera menu
		'// Gera Menus ---------------------------- >
		subTitle("> Enviar Foto")
		%>
	  </td>
    </tr>
    <tr>
      <td class="texto_size_10" id="conteudo_page">Selecione as Fotos<br />
	  	  <%
	  	for x=1 to 10
	  %>
		<input name="<%=x%>" type="file" id="<%=x%>_item" size="60"><br />
     <%
		next
	%><input name="id_ar" type="hidden" value="<%=formus("id_ar",3)%>" />
	<input name="id" type="hidden" value="<%=formus("id",3)%>" />
	</td>
    </tr>
	<%
		case "send_foto_tp2"
		'//Enviar Foto Tipo 2 ------------------------------------
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
		AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?op=list_foto&id_ar="&formus("id_ar",3)&"&id="&formus("id",3)&"&aspid=' + AspId)"
		AspMenuOp(1) = "Salvar|3|miniMenu_enviar_ft()"
		'//----------------------------------------
		AspMenu() 'gera menu
		'// Gera Menus ---------------------------- >
		subTitle("> Enviar Foto - Modo Beta")
		%>
	  </td>
    </tr>
    <tr>
      <td class="texto_size_10" id="conteudo_page">
<!--form action="" method="post" enctype="multipart/form-data" name="form1" -->
  <p class="texto_size_10">Selecione a Pasta<br>
  <input name="id_ar" type="hidden" value="<%=formus("id_ar",3)%>" />
	<input name="id" type="hidden" value="<%=formus("id",3)%>" />
      <input type="file" name="files" id="files">
	  <input name="arqSel" id="arqSel" type="hidden" value="" />
	  <input name="inicio" type="hidden" value="1" />
      <br>
      <input name="lst_tp" type="radio" value="1" checked onClick="tp.value=this.value">
    Adicionar este Arquivo 
    <input name="lst_tp" type="radio" value="0" onClick="tp.value=this.value">
	<input type="hidden" value="1" name="tp" id="tp">
    Adicionar todos os arquivos desta pasta.<br>
    <input type="button" name="Button" value="Adicione Agora" onClick="opadd(tp.value,files.value)">
  <div id="msg" ></div>
  <div id="resultado"  style="float:left; margin-right: 10px"></div>
  <div id="resum" style="float:left"></div>
<!--/form-->	  </td>
    </tr>
	<%
		case "list_foto"
		'//Listar Fotos ------------------------------------
		dim patro
		db "leitura",rs,sql,"select top 1 img_patrocinio from album_album where id="&formus("id",3),conn
		if not rs.eof and not rs.bof then
			if rs("img_patrocinio") <> "" then
				patro = 1
			else
				patro = 0
			end if
		else
			patro = 0
		end if
	%>
    <tr>
      <td id="menutab">
	  
	  	<%
		id_ar = request("id_ar")
	  	'// Gera Menus ----------------------------- >
		'// 0-menu;1-voltar;2-imprimir;3-confirmar;4-incluir;
		'// 5-apagar;6-editar;7-reportar;8-mostrarall;
		'// 9-ocultarall;10-item.
		'//----------------------------------------- >
		
		if patro=1 then
			redim AspMenuOp(6) 'Define o numero de itens
		else
			redim AspMenuOp(5) 'Define o numero de itens
		end if
		
		'//campos ---------------------------------
		AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?op=list_album&id_ar="&request.QueryString("id_ar")&"&aspid=' + AspId)"
		AspMenuOp(1) = "Destacar|3|destaq_foto(document.getElementById('fotosel').value,'"&request.QueryString("id")&"')"
		AspMenuOp(2) = "Excluir|5|excluir_foto(document.getElementById('fotosel').value,'"&request.QueryString("id")&"')"
		AspMenuOp(3) = "Inc.Foto|4|WindowOpen('default.asp?op=send_foto_tp1&id="&request.QueryString("id")&"&id_ar="&id_ar&"&aspid=' + AspId)"
		AspMenuOp(4) = "Inc.Beta|4|WindowOpen('default.asp?op=send_foto_tp2&id="&request.QueryString("id")&"&id_ar="&id_ar&"&aspid=' + AspId)"

		if patro = 1 then
			AspMenuOp(5) = "Inc.Patr.|10|incpat('"&request.QueryString("id")&"','"&id_ar&"','"&aspid&"')"
		end if
		
		'//----------------------------------------
		AspMenu() 'gera menu
		'// Gera Menus ---------------------------- >
		 subTitle("> Editar Foto do Álbum")
		%>
	  
	  </td>
    </tr>
    <tr>
      <td id="conteudo_page">
	  <input name="fotosel" id="fotosel" type="hidden" value="" />
	  <% response.write "</form>" %>
	  <link href="../gui/ajax_estilo.css" rel="stylesheet" type="text/css" />
	  <div align="center" id="AlbumAberto">	 		
	  </div>
	  
	  <script>AjaxLoad('inc_album_abertos.asp?id=<%=request.QueryString("id")%>&aspid=<%=aspid%>','','AlbumAberto');</script>
	  </td>
    </tr>
	<%
		case "del"
		
			select case request("tp")
			case "area"
				
				'// Listar todos as fotos desta área antes de apagar a area
				db "leitura",rs,sql,"select * from AlbumListFotoArea where id_a="&formus("id",3),conn
				while not rs.eof
						
						delfile foto_root&trim(rs("foto"))
						delfile fotog_root&trim(rs("foto"))				
				
					rs.movenext
				wend
				
				'// Deleta todos os albuns desta area
				db "leitura",rs,sql,"select * from album_album where id_ar="&formus("id",3),conn
				while not rs.eof 
					
						db "normal",,sql,"delete from album_fotos where id_a="&rs("id"),conn
						db "normal",,sql,"delete from album_album where id="&rs("id"),conn
				
					rs.movenext
				wend 
				
				db "normal",,sql,"delete from album_area where id="&formus("id",3),conn
				voltar request.ServerVariables("SCRIPT_NAME")&"?op=list_area&aspid="&aspid,"","_self"
				
			case "album"
			
				'// Listar todos as fotos desta área antes de apagar a area
				db "leitura",rs,sql,"select * from AlbumListFotoArea where id_album="&request("id"),conn
				while not rs.eof
						
						delfile foto_root&trim(rs("foto"))
						delfile fotog_root&trim(rs("foto"))				
				
					rs.movenext
				wend
				
				db "normal",,sql,"delete from album_fotos where id_a="&request("id"),conn
				db "normal",,sql,"delete from album_album where id="&request("id"),conn
				voltar request.ServerVariables("SCRIPT_NAME")&"?op=list_album&aspid="&aspid,"","_self"
						
			case else
				voltar request.ServerVariables("SCRIPT_NAME")&"?aspid="&aspid,"","_self"
			end select			
		case "inc_pat"
			
			db "leitura",rs,sql,"select foto, id_a from album_fotos where id_a="&formus("id",3),conn
			while not rs.eof
			 
					dim jpeg_zoeira_new, jpeg_zoeira_new_Logo, jpeg_zoeira_new_LogoPath
					
					imga = server.MapPath(fotog_root&trim(rs("foto")))
					Set jpeg_zoeira_new = Server.CreateObject("Persits.Jpeg")
					
					jpeg_zoeira_new.Open(imga)
					
					'//Tamanho da imagem 
					jpeg_zoeira_new.Width = 378 
					jpeg_zoeira_new.Height = 283 

					'//nivel de nitidez
					'jpeg_zoeira_new.Sharpen 1, 130 '130
					
					
					jpeg_zoeira_new.Interpolation = 1
					jpeg_zoeira_new.Quality = 90 '//compactação
					
					'jpeg_zoeira_new.Canvas.Pen.Color = &HFF6600
					'jpeg_zoeira_new.Canvas.Pen.Width = 40
					'jpeg_zoeira_new.Canvas.Line 0, jpeg_zoeira_new.Height, jpeg_zoeira_new.Width, jpeg_zoeira_new.Height
					
					'//--------------------------------------------------------------------------------					
					Set jpeg_zoeira_new_Logo = Server.CreateObject("Persits.Jpeg")
					
					'//Imagem de Patrocinio
					'if not rs.eof and not rs.bof then
						'if trim(rs("img_patrocinio"))<>"" then
							'jpeg_zoeira_new_LogoPath = server.MapPath(pImg&"logos_p.gif") 'selo tonazoeira para patrocinio
						'else
							'jpeg_zoeira_new_LogoPath = server.MapPath(pImg&"logos.gif") 'selo tonazoeira
						'end if
					'end if
					
					'jpeg_zoeira_new_Logo.Open jpeg_zoeira_new_LogoPath
					'jpeg_zoeira_new.DrawImage 0, jpeg_zoeira_new.Height-100, jpeg_zoeira_new_Logo, , &H0000FF 
					'//----------------------------------------------------------------------------------
					
					'//Fixando imagem do patrocinador
					'if not rs.eof and not rs.bof then
					'	if trim(rs("img_patrocinio"))<>"" then
							Set jpeg_zoeira_new_Logo = Server.CreateObject("Persits.Jpeg")
					
							'//Imagem de Patrocinio
							jpeg_zoeira_new_LogoPath = server.MapPath(pImg&"patrocinio_"&rs("id_a")&".gif") 'selo patrocinador
							jpeg_zoeira_new_Logo.Open jpeg_zoeira_new_LogoPath
							jpeg_zoeira_new.DrawImage 0, 0, jpeg_zoeira_new_Logo, , &H0000FF 
							
					'	end if
					'end if
					'//----------------------------------------------------------------------------------
					
					
					'jpeg_zoeira_new.Canvas.Font.Color = &HFFFFFF' red
					'jpeg_zoeira_new.Canvas.Font.Family = "Tahoma"
					'jpeg_zoeira_new.Canvas.Font.Bold = True
					'jpeg_zoeira_new.Canvas.Font.Size = 14
					'jpeg_zoeira_new.Canvas.Print 65, jpeg_zoeira_new.Height-18, "w w w . t o n a z o e i r a . c o m . b r"
					
					jpeg_zoeira_new.Save imga
					
					rs.movenext
				wend
				
					set jpeg_zoeira_new = nothing
					
				alert("Operação concluída com sucesso!")
				voltar request.ServerVariables("SCRIPT_NAME")&"?op=list_foto&id="&formus("id",3)&"&aspid="&aspid&"&id_ar="&formus("id_ar",3),"","_self"
			
		case else
		'//Menu Principal ------------------------------------
	%>
	<tr>
      <td id="menutab">&nbsp;</td>
    </tr>
    <tr>
      <td id="conteudo_page"><p class="texto_size_10">Selecione qual o tipo de informa&ccedil;&atilde;o deseja visualizar:</p>
        <ul class="texto_size_10">
          <li><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=cad_area&aspid=<%=aspid%>" target="_self">Criar nova &Aacute;rea de Fotos</a></li>
          <li><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=list_area&aspid=<%=aspid%>" target="_self">Listar &Aacute;reas Criadas</a></li>
          <li><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=cad_album&aspid=<%=aspid%>" target="_self">Criar novo &Aacute;lbum</a></li>
          <li><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=list_album&aspid=<%=aspid%>" target="_self">Listar &Aacute;lbuns</a></li>
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
