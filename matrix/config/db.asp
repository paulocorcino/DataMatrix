<%
	Dim db_oledb,db_oledbs,db_oledbcs, sites	
' Altere somentes estas linhas =======================================================================================

	db_oledb = "Provider=sqloledb; Data Source=127.0.0.1; Initial Catalog=matrix_bymidia; User Id=corcino; Password=cino577#"
			
	' Coloca as datas e moedas no padrão brasileiro
	Session.LCID = 1046
	
	'URL
	sites = "http://"&request.ServerVariables("HTTP_HOST")
	
	
	'Informações do site Boaluz
	dim teste, cotacao_data, f, cotacao_venda, cotacao_compra, i, Gkeys, email_tec
	dim site, roots, tema_root,ban_root,ban_g, banner_root, selo,foto_root,fotoo_root,tempo_root, xml_root, swf_root, nome_site, fotog_root, noticia_root,xmlnoticia_root,fotonoticia_root,dh,dw,htmlnoticia_root,dfotonoticia_root,dwm
		nome_site = "ByMidia"
		email_tec = "sac@bymidia.com.br"
		
		Gkeys = "ç"
		
		select case left(request.ServerVariables("HTTP_HOST"),3)
		case "loc","192","10.","127","200"
			roots = "/bymidia"
		case else
			roots = ""
		end select
		
		site = trim("http://"&request.ServerVariables("HTTP_HOST")&roots)
		
		tema_root 	= roots&"/tema/"
		banner_root	= roots&"/banner/"
		foto_root	= roots&"/fotos/"
		fotog_root	= roots&"/fotos/gr/"
		fotoo_root	= roots&"/fotos/or/"
		ban_root	= roots&"/banner/"
		ban_g	= roots&"/banner/gr/"
		swf_root	= roots&"/swf/"
		xml_root	= roots&"/xml/"
		tempo_root	= roots&"/tempo/"
		noticia_root = roots&"/matrix/noticia/foto/"
		xmlnoticia_root	= roots&"/matrix/noticia/xml/"
		htmlnoticia_root = roots&"/matrix/noticia/html/"
		fotonoticia_root = roots&"/matrix/noticia/foto/"
		fotonoticia_root = roots&"/matrix/noticia/foto/"
		dfotonoticia_root = roots&"/matrix/noticia/dfoto/"

		dim pImg, pImg_temp
		pImg = roots&"/img/"
		pImg_temp = roots&"/fotos/temp/"
		
			
		selo = "<img src="&roots&"/matrix/img/selocs.gif border=0 />"

	
	'dimensão das fotos de destaque
	'dim dh, dw
	'//Fotos do sistema de Noticias
	dh = 63
	dw = 86
	dwm = 60
	
	'//Dimensão Fotos do Album
	dim wAlbumThumb, hAlbumThumb
	dim wAlbumFull, hAlbumFull
	dim seloAlbum, seloAlbumPatrocinio
	
	wAlbumThumb = 86 'largura
	hAlbumThumb = 63 'altura
	wAlbumFull = 378 'largura
	hAlbumFull = 283 'altura
	seloAlbum = "logos.gif" 'imagem selo do album
	seloAlbumPatrocinio = "logos_p.gif" 'imagem selo do album quando existe imagem do patrocinador
	
' Fim ================================================================================================================

'Voltando a página 
'Usa-se: voltar "pagina.asp?op=al","Erro aqui","_self"
	function voltar(pag,erro,rota)
		response.write "<script language=""JavaScript"">"&vbcrlf
		response.write "<!--"&vbcrlf
		response.write "window.open('"&pag&"&erro="&server.URLEncode(erro)&"','"&rota&"');"&vbcrlf
		response.write "-->"&vbcrlf
		response.write "</script>"&vbcrlf
		response.write "<font size=""2"" face=""Arial, Helvetica, sans-serif"">Se a p&aacute;gina n&atilde;o abrir "&vbcrlf
		response.write "automaticamente <a href="""&pag&"&erro="&server.URLEncode(erro)&""" target="""&rota&""">clique aqui</a></font>"&vbcrlf
	end function
'captura o ip do usuário
dim tpip
	tpip = 1

function ip()
	if tpip=1 then
		ip = request.ServerVariables("REMOTE_ADDR")
	else
		ip = "127.0.0.5"
	end if	
end function

'Valida o usuário
	function seguranca(nivel)
		
		'verifica se o cookie amarazenado pertence ao site
		if request.Cookies("naweb")("site") = nome_site then
			
			'se sim indica que ele suporte cookie, neste caso vamos analizar se a identidade dele é igal a registrada
			if request.Cookies("naweb")("ident") = session.SessionID then
				
				'vamos verificar seu nivel é diferente do exigido
				if request.Cookies("naweb")("nivel") > nivel then
					
					'erro, envia para página inical
					voltar site&"/matrix/senha/no.asp?0=1","Seu nivel não permite o acesso a página requerida. ","_self"
					response.end()
					
				end if
			
			else
			
				'erro, envia para página inical
				voltar site&"/matrix/senha/login.asp?pag=0","Seu acesso expirou, efetue o login. ","_self"
				response.end()
			
			end if
		
		else
		
				'erro, envia para página inical
				voltar site&"/matrix/senha/login.asp?pag=0","Seu acesso expirou, efetue o login. ","_self"
				response.end()
		
		end if
	
	end function
	
	'Valida o usuário
	dim rsseg,sqlseg,iduseg,nomeaspid,nivelaspid,idaspid
	function protec(nivel)
	
	if request.QueryString("aspid") <> "" then
		
		db "leitura",rsseg,sqlseg,"SELECT TOP 100 PERCENT nivel,nome,id_u FROM dbo.auditoria WHERE (sessionid = '"&aspid()&"') AND (logado = 1)",conn

		
		if rsseg.eof and rsseg.bof then
			'erro, envia para página inical
			voltar site&"/matrix/senha/no.asp?0=1","Seu acesso expirou, efetue o login. ","_self"
			fdb(rsseg)
			fdb(conn)
			response.End()
		else
			if rsseg("nivel") > cint(nivel) then
					
					voltar site&"/matrix/senha/login.asp?pag=0","Seu nivel não permite o acesso a página requerida.","_self"
					fdb(conn)
					response.End()
			else
			
			nomeaspid = rsseg("nome")
			nivelaspid = rsseg("nivel")
			idaspid = rsseg("id_u")
			
					fdb(rsseg)
			end if
		end if
	else
			voltar site&"/matrix/senha/login.asp?pag=0","Seu acesso expirou, efetue o login. ","_self"
			fdb(conn)
			response.End()
	end if
	
	end function
	
	dim unome, unick, univel, usexo, uid
	univel=10
	function user(nivel)
	
		if aspid<>"" then
		db "leitura",rsseg,sqlseg,"SELECT dbo.cad_list.id, dbo.cad_list.nome, dbo.cad_list.nivel, dbo.cad_list.nick, dbo.cad_list.sexo, dbo.cad_auditoria.aspid, dbo.cad_auditoria.id_u, dbo.cad_list.atv, dbo.cad_auditoria.logado FROM dbo.cad_list INNER JOIN dbo.cad_auditoria ON dbo.cad_list.id = dbo.cad_auditoria.id_u WHERE (dbo.cad_list.atv = 1) AND (dbo.cad_auditoria.aspid = '"&aspid()&"') AND (dbo.cad_auditoria.logado = 1)",conn
			if rsseg.eof and rsseg.bof then
				'erro, envia para página inical
				alert("Seu acesso expirou, autentique-se novamente!!!")
				voltar site&"?0=0","","_self"
				fdb(rsseg)
				fdb(conn)
				response.End()
			else
				if nivel <> 0 then
						if rsseg("nivel") > cint(nivel) then
								
								alert("Area de acesso restrito.[n]Cadastre-se Gratuitamente e aproveite!!!")
								voltar site&"?0=0","","_self"
								fdb(conn)
								response.End()
						else
						
							unome = rsseg("nome")
							unick = rsseg("nick")
							univel = rsseg("nivel")
							usexo = rsseg("sexo")
							uid = rsseg("id")
							fdb(rsseg)
						end if
				else
					unome = rsseg("nome")
					unick = rsseg("nick")
					univel = rsseg("nivel")
					usexo = rsseg("sexo")
					uid = rsseg("id")
					fdb(rsseg)
				end if
			end if
		else
			if nivel <> 0 then
				alert("Area de acesso restrito.[n]Cadastre-se Gratuitamente e aproveite!!!")
				voltar site&"?0=0","","_self"
				fdb(conn)
				response.End()
			end if
		end if
	end function
	
	
	Dim AspIdRs, AspIdSql, AspidValid 
	function AspidValids(id)
		
		db "leitura", AspIdRs,AspIdSql,"select dbo.cad_auditoria.id, dbo.cad_auditoria.aspid, dbo.cad_auditoria.logado from dbo.cad_auditoria where dbo.cad_auditoria.aspid='"&replace(id,"'","")&"'",conn
		if AspIdRs.eof and AspIdRs.bof then
			AspidValid = 0
		else
			if AspIdRs("logado") = 0 then
				AspidValid = 0
			else
				AspidValid = 1
			end if
		end if
		fdb(AspIdRs)
	
	end function
	
%>