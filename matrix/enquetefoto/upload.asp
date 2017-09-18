<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<!--#include virtual="/bymidia/matrix/config/upload.asp" -->
<%
	'protec("2")
	cache 1, false
	server.ScriptTimeout = 60000
	
	'//------------------------------------------------------
	'Cria pastas caso não exista
	md(pImg_temp) 'cria a pasta caso não exista
	md(pImg) 'cria a pasta caso não exista
	'/-------------------------------------------------------
%>
<p><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><div id="message">Aguarde.....</div></font>
  <%

	Dim owNoOverwrite, owOverwrite, owUnique, imgerror, id_ar
	Dim objUpload, objFile, sql, rs, texto_eng, texto_esp, x, foton
	Dim raiz, cabecalho, registro, ID, TEXTO, TITULO,FONTE, TIPO, FOTO, rs2
	dim ImgExist, nrAlt, voto, sqls, votos, imgarray, titarray, imgsplit, titsplit, votoarray, votosplit, voto_val, totalvotos, repeteuser

	
	ImgExist = true
	
		Set objUpload = New Upload

			objUpload.MaxUploadSize = 7168
			objUpload.ValidExtensions = "jpg,gif,bmp,png,jpeg"
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
	
		'armazena as varaiaveis

		
		dim extfiles, data_i, data_f,jpeg_zoeira_new,jpeg_zoeira_new_Logo,jpeg_zoeira_new_LogoPath, tpc, u_image, arq, exten, filehtml, folderhtml, nPageFile
		
		for x = 1 to objUpload.form("alt")
		if objUpload.form("isupdate") <> "1" then
			if objUpload.File(x).Filename<>"" then
				if objUpload.File(x).Saved=true then
						
						'//imagem salva
						if ImgExist then
							'//caso exista um erro o sistema não mas ira confirmar o cadastro da enquete
							ImgExist = true		
						end if
						
						set objFile = objUpload.File(x)
						
							arq = "enquete"&x&"-"&formus("cod_area",2)&".jpg"

							delfile foto_root&arq  '//apaga uma suposta imagem nesta pasta
							delfile fotog_root&arq '//apaga uma suposta imagem nesta pasta
						
							renfile pImg_temp&objFile.Filename, foto_root&arq '//copia a imagem para a nova pasta
							delfile pImg_temp&objFile.Filename '//apaga a imagem do temp
								
							jpeg foto_root,fotog_root,arq,"autored",wAlbumFull,hAlbumFull,80,110
							jpeg foto_root,foto_root,arq,"prop",wAlbumThumb,hAlbumThumb,80,110 'mini imagem

							
							imgarray = arq&","&imgarray
							titarray = formus("titulo"&x,2)&","&titarray
							voto_val = formus("votos"&x,2)
							
							if isnumeric(voto_val) then
								voto_val = cint(voto_val)
								if voto_val = 0 then
									voto_val = 1
								else
									voto_val = voto_val + 1
								end if
							else
								voto_val = 1
							end if
							
							votoarray = voto_val&","&votoarray
						
				else
					ImgExist = false
					imgerror = "A imagem já existe no servidor ou seu tamanho é superior a 2mb"
				end if
			else
				ImgExist = false
				imgerror = "A imagem já existe no servidor ou seu tamanho é superior a 2mb"
			end if 
		else
				titarray = formus("titulo"&x,2)&","&titarray
				voto_val = formus("votos"&x,2)
				
				if isnumeric(voto_val) then
					voto_val = cint(voto_val)
					if voto_val = 0 then
						voto_val = 1
					else
						voto_val = voto_val + 1
					end if
				else
					voto_val = 1
				end if
				
				votoarray = voto_val&","&votoarray
		end if
		next
	END IF
	
	'se tudo estiver ok fazer proximo processo
	if ImgExist then
	
		if objUpload.form("isupdate") <> "1" then
			imgsplit = split(left(imgarray,len(imgarray)-1),",")
		end if
		
		titsplit = split(left(titarray,len(titarray)-1),",")
		votosplit = split(left(votoarray,len(votoarray)-1),",")
		
		'//antes de cadastrar as imagens deve-se saber quantas vagas e quais estão abertas no banco.
		db "leitura",rs,sql,"select id from enquetefoto_respostas where id_p="&formus("id",2),conn
		
		for x=lbound(titsplit) to ubound(titsplit)
			
			'//cadastra as imagens
			if rs.eof and rs.bof or rs.eof then
			
				'//caso não exista nada crie
				response.write "aqui 1"
				db "normal",,sql,"insert into enquetefoto_respostas(id_p,fotos,votos,titulo) values("&formus("id",2)&",'"&imgsplit(x)&"',"&votosplit(x)&",'"&titsplit(x)&"')",conn
				
				db "leitura",rs2,sql,"select id from enquetefoto_respostas where fotos = '"&imgsplit(x)&"'",conn
				
				ID = rs2("id")
				FOTO = "fotoenquete_"&ID&".jpg"
				renfile foto_root&imgsplit(x), foto_root&FOTO
				renfile fotog_root&imgsplit(x), fotog_root&FOTO
				
				db "normal",,sql,"update enquetefoto_respostas set fotos='"&FOTO&"',votos="&votosplit(x)&",titulo='"&titsplit(x)&"' where id="&ID,conn
				
			else
				
				'//case existe salve (update)
				ID = rs("id")
				'FOTO = "fotoenquete_"&ID&".jpg"
				'renfile foto_root&imgsplit(x), foto_root&FOTO
				'renfile fotog_root&imgsplit(x), fotog_root&FOTO
				'response.write "aqui 2" & titarray
				db "normal",,sql,"update enquetefoto_respostas set votos="&votosplit(x)&",titulo='"&titsplit(x)&"' where id="&ID,conn
				
				
				rs.movenext
			end if
			'response.write 
			
		next
		
		totalvotos = formus("totalvotos",2)
		
		repeteuser = formus("repeteuser",2)
		
		if repeteuser <> "1" then
			repeteuser = 0
		end if
		
		if totalvotos <> "1" then
			totalvotos = 0
		end if
		
		db "normal",,sql,"update enquetefoto_perguntas set id_a="&formus("cod_area",2)&",pergunta='"&formus("pergunta",2)&"',criado=1,entrada='"&fdate(formus("entrada",2),"yyyymmdd")&"',saida='"&fdate(formus("saida",2),"yyyymmdd")&"',alt='"&formus("alt",2)&"',totalvotos="&totalvotos&",repeteuser="&repeteuser&" where id="&formus("id",2),conn
		' salvo com sucesso
		'response.write "ok"
		alert("Informações Salvas com sucesso!")
		voltar "default.asp?op=cad_enquete&aspid="&aspid()&"&id="&formus("id",2),"","_self"
	else
		'deu problema erro
		alert(imgerror)
		voltar "default.asp?op=cad_enquete&aspid="&aspid()&"&id="&formus("id",2),"","_self"
		
	end if		

	
	
	set objUpload = nothing '//cancela envio de foto
	
	'Fechando banco de dados
	fdb(conn)

%>

