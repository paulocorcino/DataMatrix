<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<!--#include virtual="/bymidia/matrix/config/upload.asp" -->
<%
	'protec("2")
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
			On Error Resume Next
			objUpload.ProcessRequest
			
	
	'Validando Upload -------------------------------------------------------------------------
	If Err.Number<>0 then
		imgerror = "O sistema detectou um erro do servidor, provavelmente limitação no tamanho do Upload." & chr(10) & "Sujestão: Envie apenas uma imagem de no máximo 200Kb, caso o envio seja efetivado com sucesso entre em contato com o provedor para aumentar a upload maximo permitido em servidores windows 2003"
	else
	'cadastra banner
	dim extfiles, data_i, data_f,jpeg_zoeira_new,jpeg_zoeira_new_Logo,jpeg_zoeira_new_LogoPath, tpc, u_image, arq, exten, filehtml, folderhtml, nPageFile
	select case request.QueryString("op")
	case "album"
		

		'//salvar dados 
		db "normal",,sql,"update album_album set id_ar="&formus("id_ar",2)&",album='"&formus("album",2)&"',quando='"&formus("quando",2)&"',aonde='"&formus("aonde",2)&"',quem='"&formus("quem",2)&"',coment='"&formus("coment",2)&"',fotografo='"&formus("fotografo",2)&"',criado=1 where id="&formus("id",2),conn
				
		if formus("img_pat",2) = "0" then
			'//Exclui imagem patrocinador
					
			db "normal",,sql,"update album_album set img_patrocinio=null where id="&formus("id",2),conn
			arq = "patrocinio_"&formus("id",2)&".gif"
			delfile pImg&arq 'se existir alguma foto com o mesmo nome ela deve ser deletada
			
			
						
		end if
		
		if objUpload.File(1).Filename<>"" then
			if objUpload.File(1).Saved=true then
				set objFile = objUpload.File(1)
				
				arq = "patrocinio_"&formus("id",2)&"."&right(objFile.Filename,3)
				delfile pImg&arq 'se existir alguma foto com o mesmo nome ela deve ser deletada
				renfile pImg_temp&objFile.Filename, pImg&arq
				delfile pImg&objFile.Filename	
				alert("Dados e Imagem salvos com sucesso!")
				
				db "normal",,sql,"update album_album set img_patrocinio='"&arq&"' where id="&formus("id",2),conn
				
			else
				alert("A foto está em um formato inválido ou o tamanho é acima do permitido")									
			end if
		else
			alert("Dados salvos com sucesso!")		
		end if
	
		if formus("criado",2) = "0" then
			voltar "default.asp?aspid="&aspid&"&op=cad_album","","_self"
		else
			voltar "default.asp?aspid="&aspid&"&op=cad_album&id="&formus("id",2),"","_self"
		end if
	
	case "fotos"
	
		'//Preparando pra enviar fotoas
		for x=1 to 10
			
		if objUpload.File(x).Filename<>"" then
			if objUpload.File(x).Saved=true then
				
				db "normal",,sql,"INSERT INTO album_fotos(foto) VALUES ('"&x&"')",conn
				db "leitura",rs,sql,"select id from album_fotos where foto='"&x&"' order by id desc",conn
				
				set objFile = objUpload.File(x)
				
				registro = rs("id")
				arq = rs("id")&".jpg"
				
				delfile foto_root&arq  '//apaga uma suposta imagem nesta pasta
				delfile fotog_root&arq '//apaga uma suposta imagem nesta pasta
				
				renfile pImg_temp&objFile.Filename, foto_root&arq '//copia a imagem para a nova pasta
				delfile pImg_temp&objFile.Filename '//apaga a imagem do temp
				
				
				'//----------------------------------------------------------------------------
				'Salva a imagem dentro dos padrões do album
				'//----------------------------------------------------------------------------
				
				db "leitura",rs,sql,"select img_patrocinio from album_album where id="&formus("id",2),conn
				
					imga = server.MapPath(foto_root&arq)
					Set jpeg_zoeira_new = Server.CreateObject("Persits.Jpeg")
					
					jpeg_zoeira_new.Open(imga)
					
					'//Tamanho da imagem 
					jpeg_zoeira_new.Width = wAlbumFull 
					jpeg_zoeira_new.Height = hAlbumFull 

					'//nivel de nitidez
					jpeg_zoeira_new.Sharpen 1, 130 '130
					
					
					jpeg_zoeira_new.Interpolation = 1
					jpeg_zoeira_new.Quality = 90 '//compactação
					
					'jpeg_zoeira_new.Canvas.Pen.Color = &HFF6600
					'jpeg_zoeira_new.Canvas.Pen.Width = 40
					'jpeg_zoeira_new.Canvas.Line 0, jpeg_zoeira_new.Height, jpeg_zoeira_new.Width, jpeg_zoeira_new.Height
					
					'//--------------------------------------------------------------------------------	
					SET fso = Server.CreateObject("Scripting.FileSystemObject")
					
									
						Set jpeg_zoeira_new_Logo = Server.CreateObject("Persits.Jpeg")
						
						'//Imagem de Patrocinio
						if not rs.eof and not rs.bof then
							if trim(rs("img_patrocinio"))<>"" then
								If fso.FileExists(server.MapPath(pImg&seloAlbumPatrocinio)) Then
									jpeg_zoeira_new_LogoPath = server.MapPath(pImg&seloAlbumPatrocinio) 'selo tonazoeira para patrocinio
									ImgExist = true
								end if
							else
								If fso.FileExists(server.MapPath(pImg&seloAlbum)) Then
									jpeg_zoeira_new_LogoPath = server.MapPath(pImg&seloAlbum) 'selo tonazoeira
									ImgExist = true
								end if
							end if
						end if
						
						if ImgExist then
							jpeg_zoeira_new_Logo.Open jpeg_zoeira_new_LogoPath
							jpeg_zoeira_new.DrawImage 0,0, jpeg_zoeira_new_Logo, , &H0000FF 
							ImgExist = false
						end if
						
					
					'//----------------------------------------------------------------------------------
					
					'//Fixando imagem do patrocinador
					if not rs.eof and not rs.bof then
						if trim(rs("img_patrocinio"))<>"" then
							Set jpeg_zoeira_new_Logo = Server.CreateObject("Persits.Jpeg")
					
							'//Imagem de Patrocinio
							If fso.FileExists(server.MapPath(pImg&trim(rs("img_patrocinio")))) Then
								jpeg_zoeira_new_LogoPath = server.MapPath(pImg&trim(rs("img_patrocinio"))) 'selo patrocinador
								jpeg_zoeira_new_Logo.Open jpeg_zoeira_new_LogoPath
								jpeg_zoeira_new.DrawImage 0, 0, jpeg_zoeira_new_Logo, , &H0000FF 
							end if
							
						end if
					end if
					'//----------------------------------------------------------------------------------
					
					
					'jpeg_zoeira_new.Canvas.Font.Color = &HFFFFFF' red
					'jpeg_zoeira_new.Canvas.Font.Family = "Tahoma"
					'jpeg_zoeira_new.Canvas.Font.Bold = True
					'jpeg_zoeira_new.Canvas.Font.Size = 14
					'jpeg_zoeira_new.Canvas.Print 65, jpeg_zoeira_new.Height-18, "w w w . t o n a z o e i r a . c o m . b r"
					
					jpeg_zoeira_new.Save server.MapPath(fotog_root&arq)
					
					set jpeg_zoeira_new = nothing
				
				'//----------------------------------------------------------------------------
				
				'//cria thumbmail
				jpeg foto_root,foto_root,arq,"paulo",wAlbumThumb,hAlbumThumb,100,110
				
				'cadastrando imagem no banco ----------------------------		
				db "normal",,sql,"update album_fotos set id_a="&objUpload.Form("id")&" ,foto='"&arq&"' where id="&registro,conn
				
				
				db "leitura",rs,sql,"select id from album_fotos where id_a="&objUpload.Form("id")&" and dest=1",conn
				if rs.eof and rs.bof then
					db "normal",,sql,"update album_fotos set dest=1 where foto='"&arq&"'",conn
				end if
				'----------------------------------------------------------
				
			else
				delfile pImg_temp&objUpload.File(x).Filename '//emcaso de erro apaga a imagem do temp
				imgerror = imgerror & "Erro ao enviar a imagem "&objUpload.File(x).Filename&" , devido ao tamanho da imagem" & chr(10)			
			end if
		end if
			
		next
	case "fotosbeta"
	
		'//Prepara lista de arquivo para seren enviados
		if formus("inicio",2)="1" then
			registro = split(lcase(formus("FileValid",2)),"jpg,")
			for i=lbound(registro) to ubound(registro) 'formus("FileValid",2)
				session("file_"&i) = registro(i)&"jpg"
			next
			session("final") = ubound(registro)
		end if
		
		%>
		<!--[Upload multiplos Arquivos]-->
		<form action="<%=request.ServerVariables("SCRIPT_NAME")%>" method="post" enctype="multipart/form-data" name="form">
			<input type="hidden" id="first_file_element" value="<%=session("file_1")%>" />
		</form>
		Enviando arquivo 1 de <%=session("final")%>
		<div id="files_list"></div>
		<script>
			//var multi_selector = new MultiSelector( document.getElementById( 'files_list' ), 1 );
			//multi_selector.addElement( document.getElementById( 'first_file_element' ) );
			
			var new_element = document.createElement( 'input' );
			new_element.type = 'file';
			new_element.setAttribute('id','teste');
			new_element.setAttribute('value','oioioi');
			document.getElementById('files_list').appendChild(new_element);
			
		</script>
		<!--[Upload]-->		
		<%
		
		
	case else
		alert("Opção inválida")
		voltar "default.asp?aspid="&aspid,"","_self"
	end select
	end if

	
				
	

	if imgerror = "" then
			imgerror = "Todas as imagens foram enviadas com sucesso"
	end if 
	%>
</p>
<form id="form1" name="form1" method="post" action="">
  <div align="center">
    <textarea name="textarea" cols="60" rows="10" readonly="readonly"><%=imgerror%></textarea>
  </div>
  
  <div align="center">
    <input type="button" name="Button" value="Retornar" onclick="window.open('default.asp?op=send_foto_tp1&id=<%=formus("id",2)%>&id_ar=<%=formus("id_ar",2)%>&aspid=<%=formus("aspid",3)%>','_self');"/>
  </div>
</form>
<%
	set objUpload = nothing '//cancela envio de foto
	'Fechando banco de dados
	fdb(conn)

%>
<p>
<script>document.getElementById("message").innerHTML = "Resumo da Operação";</script>  
</p>
