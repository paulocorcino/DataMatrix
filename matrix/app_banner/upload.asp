<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<!--#include virtual="/bymidia/matrix/config/upload.asp" -->
<%
	'protec("2")
	server.ScriptTimeout = 60000
	
	'//------------------------------------------------------
	'Cria pastas caso não exista
	md(pImg_temp) 'cria a pasta caso não exista
	md(ban_root) 'cria a pasta caso não exista
	md(ban_g) 'cria a pasta caso não exista
	'/-------------------------------------------------------
%>
<p><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><div id="message">Aguarde.....</div></font>
  <%

	Dim owNoOverwrite, owOverwrite, owUnique, imgerror, id_ar
	Dim objUpload, objFile, sql, rs, texto_eng, texto_esp, x, foton
	Dim raiz, cabecalho, registro, ID, TEXTO, TITULO,FONTE, TIPO, FOTO
	dim ImgExist
	
	ImgExist = false
	
		Set objUpload = New Upload

			objUpload.MaxUploadSize = 7168
			objUpload.ValidExtensions = "jpg,gif,bmp,png,swf"
			owNoOverwrite = 0
			owOverwrite = 1
			owUnique = 2

			objUpload.OverwriteMode = owUnique
			objUpload.UploadPath = server.MapPath (pImg_temp)
			
			'Enviar os arquivos
			'On Error Resume Next
			objUpload.ProcessRequest
			
	
	'Validando Upload -------------------------------------------------------------------------
	If Err.Number<>0 then
		imgerror = "O sistema detectou um erro do servidor, provavelmente limitação no tamanho do Upload." & chr(10) & "Sujestão: Envie apenas uma imagem de no máximo 200Kb, caso o envio seja efetivado com sucesso entre em contato com o provedor para aumentar a upload maximo permitido em servidores windows 2003"
	else
	'cadastra banner
	dim extfiles, data_i, data_f,jpeg_zoeira_new,jpeg_zoeira_new_Logo,jpeg_zoeira_new_LogoPath, tpc, u_image, arq, exten, filehtml, folderhtml, nPageFile

		

		'//salvar dados 
		db "normal",,sql,"update tb_ban_sysban set cod_area_sysban="&formus("cod_area",2)&",datae_ban_sysban='"&fdate(formus("datae",2),"yyyymmdd")&"',datas_ban_sysban='"&fdate(formus("datas",2),"yyyymmdd")&"',url_ban_sysban='"&formus("urlban",2)&"', criado=1 where cod_ban_sysban="&formus("id",2),conn
				
		
		
	
		'//Preparando pra enviar fotoas
		'for x=1 to 10
			
		if objUpload.File(1).Filename<>"" then
			if objUpload.File(1).Saved=true then
				
				'db "normal",,sql,"INSERT INTO album_fotos(foto) VALUES ('"&x&"')",conn
				'db "leitura",rs,sql,"select id from album_fotos where foto='"&x&"' order by id desc",conn
				
				set objFile = objUpload.File(1)
				
				'registro = rs("cod_")
				extfiles = right(objFile.Filename,3)
				arq = formus("id",2)&"_sysban."&extfiles
				
				delfile foto_root&arq  '//apaga uma suposta imagem nesta pasta
				delfile fotog_root&arq '//apaga uma suposta imagem nesta pasta
				
				renfile pImg_temp&objFile.Filename, ban_root&arq '//copia a imagem para a nova pasta
				delfile pImg_temp&objFile.Filename '//apaga a imagem do temp
				'alert(extfiles)
				
				'//----------------------------------------------------------------------------
				'Salva a imagem dentro dos padrões do area banner
				'//----------------------------------------------------------------------------
				if extfiles <> "swf" and extfiles <> "gif" then
					'alert("aqui")
					db "leitura",rs,sql,"select * from tb_area_sysban where cod_area_sysban="&formus("cod_area",2),conn
				
					imga = server.MapPath(ban_root&arq)
					Set jpeg_zoeira_new = Server.CreateObject("Persits.Jpeg")
					
					jpeg_zoeira_new.Open(imga)
					
					'//Tamanho da imagem 
					jpeg_zoeira_new.Width = rs("w_area_sysban") 
					jpeg_zoeira_new.Height = rs("h_area_sysban")  

					'//nivel de nitidez
					jpeg_zoeira_new.Sharpen 1, 130 '130
					
					
					jpeg_zoeira_new.Interpolation = 1
					jpeg_zoeira_new.Quality = 90 '//compactação
					
					jpeg_zoeira_new.Save server.MapPath(ban_g&arq)
					
					set jpeg_zoeira_new = nothing
					
					'Cria thumbmail
					jpeg ban_root,ban_root,arq,"autored",150,60,100,110
				'//----------------------------------------------------------------------------
				elseif extfiles <> "swf" then
					'//cria thumbmail
					copyfile ban_root&arq, ban_g&arq
					jpeg ban_root,ban_root,arq,"autored",150,60,100,110
				else
					'formato swf
					copyfile ban_root&arq, ban_g&arq
				end if
				'cadastrando imagem no banco ----------------------------		
				db "normal",,sql,"update tb_ban_sysban set file_ban_sysban='"&arq&"',tpfile_ban_sysban='"&extfiles&"' where cod_ban_sysban="&formus("id",2),conn
				db "normal",,sql,"update tb_ban_sysban set show=0 where cod_area_sysban="&formus("cod_area",2),conn
						
			else
				delfile pImg_temp&objUpload.File(1).Filename '//emcaso de erro apaga a imagem do temp
				imgerror = imgerror & "Erro ao enviar a imagem "&objUpload.File(1).Filename&" , devido ao tamanho da imagem" & chr(10)			
			end if
		end if
			
		

	end if

	
				
	

	if imgerror = "" then
			imgerror = "Todas as imagens foram enviadas com sucesso"
	end if 
	
	alert(imgerror)
	voltar "default.asp?op=cad_banner&aspid="&aspid,"","_self"
	
	set objUpload = nothing '//cancela envio de foto
	'Fechando banco de dados
	fdb(conn)

%>

