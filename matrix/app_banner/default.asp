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
<title>Banners</title>
<link href="../css/estilos.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" type="text/javascript" src="../js/script.js"></script>
<script language="JavaScript" type="text/javascript" src="../js/form_banner.js"></script>
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
      <td id="titulo_pagina"><div id="matrix_titulo" style="float:left">Banners</div></td>
    </tr>
	<%
		dim rs,sql,id, criado, x
		dim area, chave, nban, wban, hban, posban, tban
		select case request.QueryString("op")
		case "cad_area"
		
		'//Nova Area ----------------------------------
		if lcase(request.ServerVariables("REQUEST_METHOD")) <> "post" then
			
			id = formus("id",3)
			
			if id<>"" then
				subTitle("> Alterar Área")
				'//Alterar Dados
				db "leitura",rs,sql,"select top 1 * from tb_area_sysban where cod_area_sysban="&id,conn
				area = trim(rs("area_area_sysban"))
				tban = rs("t_area_sysban")
				chave = trim(rs("chave_area_sysban"))
				nban = trim(rs("n_area_sysban"))
				wban = trim(rs("w_area_sysban"))
				hban = trim(rs("h_area_sysban"))
				posban = trim(rs("pos_area_sysban"))
				
			else
				subTitle("> Nova Área")
				'//nova area
				db "leitura",rs,sql,"select top 1 * from tb_area_sysban where criado=0",conn
				if rs.eof and rs.bof then
					'//não existe registro
					db "normal",,sql,"insert into tb_area_sysban(criado) values(0)",conn
					db "leitura",rs,sql,"select top 1 * from tb_area_sysban where criado=0",conn
				end if						
				
			end if
			
			id = rs("cod_area_sysban")
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
		%></td>
    </tr>
    <tr>
      <td id="conteudo_page"><p><span class="texto_size_10">Cod</span><br />
          <input name="id" type="text" id="id" value="<%=id%>" size="3" readonly="true" />
          <br />
          <span class="texto_size_10">Nome da &Aacute;rea</span><br />
          <input name="area" type="text" id="area" value="<%=area%>" size="30" maxlength="255" />
          <br />
          <span class="texto_size_10">Chave</span><br />
          <input name="chave" type="text" id="chave" value="<%=chave%>"  onkeyup="Clear(this)" size="10" maxlength="10" />
          <br />
          <span class="texto_size_10">Visualizar quantos banners?</span><br />
          <input name="nban" type="text" id="nban" value="<%=nban%>" size="3" maxlength="2" />
          <br />
          <span class="texto_size_10">Posi&ccedil;&atilde;o do Banner</span><br />
          <select name="posban" id="posban">
                    <option value="1" <% if posban=1 then %>selected<% end if %>>Horizontal</option>
                    <option value="0" <% if posban=0 then %>selected<% end if %>>Vertical</option>
          </select>
          <br />
          <span class="texto_size_10">Tempo (segundos)</span><br />
<input name="tban" type="text" id="tban" value="<%=tban%>" size="3" maxlength="2" />  
<br />
  <span class="texto_size_10">Dimens&atilde;o</span> <span class="texto_size_10">Largura</span><br />
  <input name="wban" type="text" id="wban" value="<%=wban%>" size="3" maxlength="3" />
  <br />
  <span class="texto_size_10">Dimens&atilde;o Altura</span><br />
  <input name="hban" type="text" id="hban" value="<%=hban%>" size="3" maxlength="3" />
  <br />
  <input name="criado" type="hidden" value="<%=criado%>" />
          <script>
	//if(isIE()){
	
		oDateMasks = new Mask("#", "number");
		oDateMasks.attach(document.form.wban);	
		oDateMasks.attach(document.form.hban);	
		oDateMasks.attach(document.form.nban);	
		oDateMasks.attach(document.form.tban);	
		
	//}
	      </script>
        </p>
      </td>
    </tr>
    <%
		else
		
			id = formus("id",0)
			area = formus("area",0)
			criado = cint(formus("criado",0))
			chave = formus("chave",0)
			nban = formus("nban",0)
			wban = formus("wban",0)
			hban = formus("hban",0)
			tban = formus("tban",0)
			posban  = formus("posban",0)
			
			'///--------- [salva dados] ---->
			db "leitura",rs,sql,"select * from tb_area_sysban where cod_area_sysban<>"&id&" and chave_area_sysban='"&chave&"'",conn
			'verifica se existe algo cadastrado com este Nome
				if rs.eof and rs.bof then 
				
					db "normal",,sql,"update tb_area_sysban set area_area_sysban='"&area&"', chave_area_sysban='"&chave&"',n_area_sysban="&nban&",w_area_sysban="&wban&", h_area_sysban="& hban &", t_area_sysban="& tban &" ,pos_area_sysban="&posban&", criado=1 where cod_area_sysban="&id,conn
					'---------------------------------------------------------------------------------------------------------------------------------------------------------------------^
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
		'AspMenuOp(1) = "Novo|4|WindowOpen('default.asp?op=cad_banner&aspid=' + AspId)"
		'//----------------------------------------
		AspMenu() 'gera menu
		'// Gera Menus ---------------------------- >
		%>	  </td>
    </tr>
    <tr>
      <td id="conteudo_page">
	   <div align="center">
	    <%
		  subTitle("> Listar Área")
		  Dim GridData(4), GridName(4), GridSQL, GridDate, GridPoss
				
				
					GridSQL = "select * from tb_area_sysban where criado=1;"
					
					GridName(0) = "Cod;35"
					GridName(1) = "Área;370"
					GridName(2) = "Chave;100"
					GridName(3) = "Alterar;60"
					GridName(4) = "Deletar;60"
									
					GridData(0) = "cod_area_sysban;1;0"
					GridData(1) = "area_area_sysban;0;0"
					GridData(2) = "chave_area_sysban;0;0"
					GridData(3) = "*Alterar;1;default.asp?op=cad_area&id=[col0]&aspid=[aspid]&m=[mod]"
					GridData(4) = "*Deletar;1;javascript: deletar('"&request.ServerVariables("SCRIPT_NAME")&"?op=del&tp=area&aspid=[aspid]&id=[col0]&m=[mod]','[col1]')"
							
					GridDate = "dd/mm/yyyy"
					GridPoss = "center"
					'GridURLS = 1
					'// Chama a Função
							
					GridForm Gkeys,15
					
	%>
	    </div>	  </td>
    </tr>
	<%
		case "cad_banner"
		'dim id_ar, album, dest, quando, aonde, quem, coment, fotografo, img_patrocinio
		dim cod_area, fileban, datae, datasb, urlban, extban
		'//Novo Album -------------------------------------
		'if lcase(request.ServerVariables("REQUEST_METHOD")) <> "post" then
			
			id = formus("id",3)
			
			if id<>"" then
				subTitle("> Alterar Banner")
				'//Alterar Dados
				db "leitura",rs,sql,"select top 1 * from tb_ban_sysban where cod_ban_sysban="&id,conn
				'response.write sql
				cod_area = rs("cod_area_sysban")
				fileban = trim(rs("file_ban_sysban"))
				datae = fdate(rs("datae_ban_sysban"),"dd/mm/yyyy")
				datasb = fdate(rs("datas_ban_sysban"),"dd/mm/yyyy")
				urlban = trim(rs("url_ban_sysban"))
				extban = trim(rs("tpfile_ban_sysban"))
				
											
			else
				subTitle("> Novo Banner")
				'//nova area
				
				db "leitura",rs,sql,"select top 1 * from tb_ban_sysban where criado=0",conn
				if rs.eof and rs.bof then
					'//não existe registro
					db "normal",,sql,"insert into tb_ban_sysban(criado) values(0)",conn
					db "leitura",rs,sql,"select top 1 * from tb_ban_sysban where criado=0",conn
				end if						
				
			end if
			
			id = rs("cod_ban_sysban")
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
		if cod_area <> "" then
			AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?op=list_banner&cod_area="&cod_area&"&aspid=' + AspId)"
		else
			AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?aspid=' + AspId)"
		end if
		
		AspMenuOp(1) = "Salvar|3|miniMenu_banner()"
		'//----------------------------------------
		AspMenu() 'gera menu
		'// Gera Menus ---------------------------- >
		%>	  </td>
    </tr>
    <tr>
      <td id="conteudo_page"><table width="396" border="0" cellspacing="2" cellpadding="0">
            <tr>
              <td width="50" valign="top" class="texto_size_10">C&oacute;d<br />
                <input name="id" type="text" id="id" value="<%=id%>" size="3" /></td>
              <td valign="top" class="texto_size_10">&Aacute;rea<br />
			  <%if extban = "" then %>
			       <select name="cod_area" id="cod_area">
				<%
					db "leitura",rs,sql,"select cod_area_sysban, area_area_sysban from tb_area_sysban where criado=1",conn
					while not rs.eof
						response.write "<option value='"&rs("cod_area_sysban")&"' "
						
						if cod_area = rs("cod_area_sysban") then
							response.write "selected" 
						end if 
						
						response.write ">"&trim(rs("area_area_sysban"))&"</option>"
						rs.movenext
					wend
				%>
                </select>
				<% 
				else
				db "leitura",rs,sql,"select cod_area_sysban, area_area_sysban from tb_area_sysban where cod_area_sysban="&cod_area,conn
				 %>
				<%=trim(rs("area_area_sysban"))%>
				<input name="cod_area" type="hidden" value="<%=rs("cod_area_sysban")%>" />
				<% end if %></td>
            </tr>
          </table>
		 <%
		 	if extban <> "" then
		%>
			<input name="noimg" type="hidden" value="1" />
		<%
				if extban = "swf" then
		%>
		<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0" >
              <param name="movie" value="<%=ban_g&fileban%>" />
              <param name="quality" value="high" />
              <embed src="<%=ban_g&fileban%>" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" ></embed>
	    </object>
		<%
				else 
		%>
			<img src="<%=ban_g&fileban%>" />
		<%		
				end if
		%>
		<% else %>
		<input name="noimg" type="hidden" value="0" />
        <span class="texto_size_10">Selecione o Banner <br />
        <input name="filebanner" type="file" id="filebanner" onchange="previewImg(document.form.filebanner,'foto_loja_select','gif,jpg,jpeg,swf')"/>
        <br />
        Preview<br />
		
        </span>
        <div id="foto_loja_select">&nbsp;</div>
		<% end if %>
        <span class="texto_size_10"><br />
		
        Data Entrada <br />
        <input name="datae" type="text" id="datae" value="<%=datae%>"  onBlur="if(this.value!=''){dataVal(this);};"  size="10" maxlength="10" />
        <br />
        Data Sa&iacute;da <br />
        <input name="datas" type="text" id="datas" value="<%=datasb%>"  onBlur="if(this.value!=''){dataVal(this);};"  size="10" maxlength="10" />
        <br />
        URL<br />
        <input name="urlban" type="text" id="urlban" value="<%=urlban%>" size="30" maxlength="255" />
        <br />
				<script>
	//if(isIE()){
	
		oDateMasks = new Mask("dd/mm/yyyy", "date");
		oDateMasks.attach(document.form.datae);	
		oDateMasks.attach(document.form.datas);	

		
	//}
	</script>	
        </span></td>
    </tr>
	<%
		dim id_ar
		case "list_banner"
		
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
		redim AspMenuOp(2) 'Define o numero de itens
		'//campos ---------------------------------
		AspMenuOp(0) = "Voltar|1|WindowOpen('default.asp?aspid=' + AspId)"
		AspMenuOp(1) = "Novo|4|WindowOpen('default.asp?op=cad_banner&aspid=' + AspId)"
		'//----------------------------------------
		AspMenu() 'gera menu
		'// Gera Menus ---------------------------- >
		%>	  </td>
    </tr>
    <tr>
      <td id="conteudo_page"><span class="texto_size_10">Selecione &Aacute;rea</span><br />
        <select name="cod_area" id="cod_area">
		<%
			id_ar = formus("cod_area",3)
			
			db "leitura",rs,sql,"select * from tb_area_sysban where criado=1",conn
			while not rs.eof
				response.write "<option value='"&rs("cod_area_sysban")&"' "
				
				if cstr(id_ar) = cstr(rs("cod_area_sysban")) then
					response.write "selected" 
				end if 
				
				response.write ">"&trim(rs("area_area_sysban"))&"</option>"
				rs.movenext
			wend
		%>
        </select> <input name="aspid" type="hidden" value="<%=aspid%>" /><input name="op" type="hidden" value="list_banner" />
	  	   <input name="Button" type="button" value="Abrir" onclick="area_abrir()"/>
	  	   <div align="center">
	    <%
		  subTitle("> Listar Banner")
		  reDim GridData(5), GridName(5)
				
				
					GridSQL = "SELECT * from BannerList where cod_area_sysban="&formus("cod_area",3)
					
					GridName(0) = "Banner;150"
					GridName(1) = "Data Entrada;80"
					GridName(2) = "Data Saida;80"
					GridName(3) = "Cliques;75"
					GridName(4) = "Editar;80"
					GridName(5) = "Deletar;60"
									
					GridData(0) = "$tpfile_ban_sysban;1;0;$case({tpfile_ban_sysban}=swf):Imagem Flash:<img src='"&ban_root&"{file_ban_sysban}'>"
					GridData(1) = "datae_ban_sysban;1;0"
					GridData(2) = "datas_ban_sysban;1;0"
					GridData(3) = "click_ban_sysban;1;0"
					GridData(4) = "*Editar;1;default.asp?op=cad_banner&cod_area={cod_area_sysban}&id={cod_ban_sysban}&aspid=[aspid]"
					GridData(5) = "*Deletar;1;javascript: deletar('"&request.ServerVariables("SCRIPT_NAME")&"?op=del&tp=banner&cod_area={cod_area_sysban}&aspid=[aspid]&id={cod_ban_sysban}','Nº {cod_ban_sysban}')"
							
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
		case "del"
			select case request.QueryString("tp")
			case "banner"
				
				db "leitura",rs,sql,"select * from tb_ban_sysban where cod_ban_sysban="&formus("id",3),conn
					
					delfile ban_root&trim(rs("file_ban_sysban"))
					
					delfile ban_g&trim(rs("file_ban_sysban"))
					
				db "normal",,sql,"delete from tb_ban_sysban where cod_ban_sysban="&formus("id",3),conn
				voltar "default.asp?op=list_banner&cod_area="&formus("cod_area",3)&"&aspid="&aspid,"","_self"
			case "area"
				
				db "leitura",rs,sql,"select * from tb_ban_sysban where cod_area_sysban="&formus("id",3),conn
				while not rs.eof 
					delfile ban_root&trim(rs("file_ban_sysban"))
					delfile ban_g&trim(rs("file_ban_sysban"))
					db "normal",,sql,"delete from tb_ban_sysban where cod_ban_sysban="&rs("cod_ban_sysban"),conn
					rs.movenext
				wend
				
				db "normal",,sql,"delete from tb_area_sysban where cod_area_sysban="&formus("id",3),conn
				voltar "default.asp?op=list_area&aspid="&aspid,"","_self"
			end select
		case "salvar_banner"
			db "normal",,sql,"update tb_ban_sysban set datae_ban_sysban='"&fdate(formus("datae",0),"yyyymmdd")&"',datas_ban_sysban='"&fdate(formus("datas",0),"yyyymmdd")&"',url_ban_sysban='"&formus("urlban",0)&"', criado=1 where cod_ban_sysban="&formus("id",0),conn
			db "normal",,sql,"update tb_ban_sysban set show=0 where cod_area_sysban="&formus("cod_area",0),conn
			alert("Dados alterados com sucesso")
			voltar "default.asp?op=cad_banner&id="&formus("id",0)&"&aspid="&aspid,"","_self"
		case else
		'//Menu Principal ------------------------------------
	%>
	<tr>
      <td id="menutab">&nbsp;</td>
    </tr>
    <tr>
      <td id="conteudo_page"><p class="texto_size_10">Selecione qual o tipo de informa&ccedil;&atilde;o deseja visualizar:</p>
        <ul class="texto_size_10">
          <li><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=cad_area&aspid=<%=aspid%>" target="_self">Criar nova &Aacute;rea de Banners </a></li>
          <li><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=list_area&aspid=<%=aspid%>" target="_self">Listar &Aacute;reas Criadas</a></li>
          <li><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=cad_banner&aspid=<%=aspid%>" target="_self">Criar novo Banner </a></li>
          <li><a href="<%=request.ServerVariables("SCRIPT_NAME")%>?op=list_banner&aspid=<%=aspid%>" target="_self">Listar Banners </a></li>
        </ul>
      <p>&nbsp;</p></td>
    </tr>
	<%
		end select
	%>
  </table>
  <br />

</form>

		
</body>
</html>
