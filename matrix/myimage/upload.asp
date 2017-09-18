<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<!--#include virtual="/bymidia/matrix/config/upload.asp" -->
<font size="2" face="Verdana, Arial, Helvetica, sans-serif">Aguarde.....</font>
<%

	Dim owNoOverwrite, owOverwrite, owUnique
	Dim objUpload, objFile, sql, rs, foton, nom
		Set objUpload = New Upload

			objUpload.MaxUploadSize = 9216
			objUpload.ValidExtensions = "jpg,jpeg,bmp,gif,png"
			owNoOverwrite = 0
			owOverwrite = 1
			owUnique = 2

			objUpload.OverwriteMode = owNoOverwrite
			objUpload.UploadPath = server.MapPath (foto_root)
			objUpload.ProcessRequest
	
	if objUpload.File(1).Filename<>"" then
		if objUpload.File(1).Saved=false then

			'erro
			voltar site&"/matrix/myimage/myimage.asp?id="&formus("id",2)&"&area="&formus("area",2),"Esta foto tem um formato inválido ou já existe!!!!","_self"
			set objUpload = nothing

			'Fechando banco de dados
			fdb(conn)
			response.End()
		
		else
			
				set objFile = objUpload.File(1)
				
				'Renomeado a imagem 
				foton = "tb"&dia&"_"&mes&"_"&ano&"__"&horas&"_"&minuto&"_"&second(now)&"__"&objUpload.Form("t")&"_"&objUpload.Form("id")&".jpg"
				renfile foto_root&objFile.Filename, foto_root&foton
				
				'Redimensiona para foto original
				jpeg foto_root,fotog_root,foton,"prop",250,0,60,130
				
				
				'Redimensiona a foto para tubmail
				jpeg foto_root,foto_root,foton,"prop",dwm,0,50,130
				
				'response.Write server.MapPath(foto_root)&"\"&foton
				if trim(formus("id",2))="" then
					nom = "Nova Imagem"
				else
					nom = formus("nome",2)
				end if
				'cadastrando os dados e o tema				
				db "normal",,sql,"INSERT INTO myimage(id_a,foto,nome,area,dst) VALUES ("&formus("id",2)&",'"&foton&"','"&nom&"','"&formus("area",2)&"',0)",conn

				'Volta a listar os arquivos
				voltar site&"/matrix/myimage/myimage.asp?id="&formus("id",2)&"&area="&formus("area",2),"","_self"
				
			
		end if
	else
		
			voltar site&"/matrix/myimage/myimage.asp?id="&formus("id",2)&"&area="&formus("area",2),"escolha um arquivo!!!!","_self"
			set objUpload = nothing
				
			'Fechando banco de dados
			fdb(conn)
			response.End()
				
		
			
	end if
	
	set objUpload = nothing
	
	'Fechando banco de dados
	fdb(conn)
%>