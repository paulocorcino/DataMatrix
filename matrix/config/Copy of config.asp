<%
'******************************************************************
'	      		     Matrix2004
'		             Versão 2.0
'         desenvolvido por: Paulo Corcino (pecjs@msn.com)
'******************************************************************
'			Página de configuração
'******************************************************************
'Parametros de inicialização
	Option Explicit
	

'Abilita analise de erro
	'On Error Resume Next
	
'Função relacionado a cache da página
	function cache(tipo,buffers)
		select case tipo
			case 1
			'Sem cache
				Response.Buffer = buffers
				Response.Expires = 0	
				Response.Expiresabsolute = Now() - 1
				Response.AddHeader "pragma","no-cache"
				Response.AddHeader "cache-control","private"
				Response.CacheControl = "no-cache"
				
			case 2
			'Cache validade 24hs
				Response.Buffer = buffers
				Response.Expiresabsolute = Now() + 1
				
			case 3
			'Com cache
				Response.Buffer = buffers
				
			end select
			
	end function

'Analizes de erros


' Valores especiais do site
	dim sessoes_max, subsessoes_max, fotos_max, not_q
	sessoes_max = 5
	subsessoes_max = 12
	fotos_max = 10 
	not_q = 5
	
%>
<!--#include virtual="/bymidia/matrix/config/db.asp" -->
<%
'//-------------------------------------------------
function er(p)
	'//Codigo Inativado
end function

'Conectando ao Banco de Dados
	Dim dbconect, conn
		'dbconect  = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & db_oledb
		dbconect  =  db_oledb
		Set conn = Server.CreateObject("ADODB.Connection")
		conn.Open dbconect

'Verificando hora atual
	Dim horas, minuto, hora
		horas = hour(now)
		minuto = minute(now)
		hora = horas&":"&minuto		
		
'Verificando a data atual
	Dim data, dia, mes, ano
		dia = Day(Now)
		mes = Month(Now)
		ano = Year(Now)
		data = dia&"/"&mes&"/"&ano
	
'Verificando a data atual padrão norte-americano
	Dim data_usa
		data_usa = mes&"/"&dia&"/"&ano

'Correção de datas para Banco de dados
'Para datas do tipo: dd/mm/aaaa hh:mm:ss usa-se CampoData( cDate( date() & " " & time() ), true )
'Para datas do tipo: dd/mm/aaaa usa-se CampoData( cDate( date() & " " & time() ), false )
'Para Consulta usa-se o próprio. Ex.: validade <"&CampoData( cDate( date() & " " & time() ), false )
'Não é necessário colocar #
	Function CampoData( Data, TimeStamp )
		If TimeStamp Then
 		CampoData = "#" & month(Data) & "/" & day(Data) & "/" & year(Data) & " " & Hour(Data) & ":" & Minute(Data) & ":" & Second(Data) & "#"
	Else
		CampoData = "#" & month(Data) & "/" & day(Data) & "/" & year(Data) & "#"
	End If
	End Function

'Fechando banco de dados
'Usa-se: fdb(rs) ou fdb(conn)
	dim fdbo
	function fdb(r)
		set fdbo = r
		fdbo.close()
		set fdbo=nothing
	end function

'Mes por extensso
	dim mess
	function mesext(mes)
		mess = Array("Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro")
		mesext = mess(cint(mes)-1)
	end function
	function mesexts(mes)
		mess = Array("Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez")
		mesexts = mess(cint(mes)-1)
	end function

	'dia semana por extensso
	dim diass
	function diaext(dia,t)
		if t = 1 then
			diass = Array("Domingo","Segunda-Feira","Terça-Feira","Quarta-Feira","Quinta-Feira","Sexta-Feira","Sabado")
		else
			diass = Array("Dom","Seg","Ter","Qua","Qui","Sex","Sab")
		end if
		diaext = diass(cint(dia)-1)
	end function
	
'E-mail Rapido V2.0 ------------------------------------------------------------------------
'Usa-se: emails "de@site.com.br","para@site.com.br","assunto","mensagem","html ou texto"
	Dim email_r, emailsms, sch, cdoConfig, cdoMessage
	function emails(de,para,assunto,mensagem,op)
	
		if Instr(trim(Request.ServerVariables("HTTP_USER_AGENT")),"Windows NT 5.0") <> 0 then
			Set email_r = Server.CreateObject("CDONTS.NewMail")
				email_r.From = de
				email_r.Subject = assunto
				email_r.To = para
				email_r.body = mensagem
			if op = "html" then
				email_r.BodyFormat = 0
				email_r.MailFormat = 0
			else
				email_r.BodyFormat = 1
				email_r.MailFormat = 1 
			end if
				email_r.send
			Set email_r = nothing
			
		else 
%>
<!--
	METADATA
	TYPE="typelib"
	UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"
	NAME="CDO for Windows 2000 Libary"
-->
<%
			
			set cdoConfig = CreateObject("CDO.Configuration")
			
			with cdoConfig.Fields
				'.Item(cdoSendUsingMethod) = 110'cdoSendUsingPort
				
				if roots = "" then
					.Item(cdoSMTPServer)      = "mail.infonet.com.br"
				else
					.Item(cdoSMTPServer)      = "localhost"
				end if
				
				.Update
			end with
			
			set cdoMessage = CreateObject("CDO.Message")
			
			with cdoMessage 
				set .configuration = cdoConfig 
				.From              = de
				.To                = para
				.Subject           = assunto
				
					if op = "html" then
						.HTMLBody = mensagem
					else
						.TextBody = mensagem
					end if
					
				On Error Resume Next
				.Send
				'Analizar este erro!!!!!
				
			end with
			
			set cdoMessage = nothing
			set cdoConfig  = nothing
					
		end if
	end function
'--------------------------------------------------------------------------

'Deletar dados
'Usa-se: <a href="#" [%=deletar("pagina.asp","nome_item")%] >

	function prep_deletar()
	response.write "<script language=""JavaScript"">"&vbcrlf
	response.write "<!--"&vbcrlf
	response.write "//Matrix2004 -  desenvolvido por Paulo Corcino - pecjs@msn.com"&vbcrlf
	response.write "function deletar(pag,nom){"&vbcrlf
	response.write "if(confirm('Você realmente deseja deletar ""'+nom+'""?')){"&vbcrlf
	response.write "window.open(pag,'_self')"&vbcrlf
	response.write "}"&vbcrlf
	response.write "}"&vbcrlf
	response.write "-->"&vbcrlf
	response.write "</script>"&vbcrlf
	end function

	function deletar(pag,nom)
		nom = replace(nom,"\","\\")
		nom = replace(nom,"&#39;","\'")
		nom = replace(nom,"""","\""")
		response.write "javascript: deletar('"&pag&"','"&nom&"')"
	end function

'Erros aviso
	dim erros
	if not request.QueryString("erro")<>"" or not request.QueryString("erro")<>chr(32) then
	erros = "&nbsp;"
	else
	erros = "<table width=100%  border=0 cellpadding=0 cellspacing=0 style='filter:alpha(opacity=60)'>"&_
    "<tr bgcolor=#00FFCC>"&_
    "<td width=30><div align=center><img src='"&site&"/matrix/img/atencao.gif' width=16 height=15></div></td>"&_
    "<td width=1130 valign=middle>&nbsp;<font face='MS Sans Serif, sans-serif, Arial, Helvetica' size='1' color='#FF0000'>"&request.QueryString("erro")&"</font> </td>"&_
    "</tr>"&_
    "</table>"
	end if
'envia para página de erro



'Conectando ao banco da dados
	function db(modo,v_rs,v_sql,instrucao,v_db)
		select case modo
			case "leitura"
				'Modo somente leitura - o banco de dados abre em modo protegido
					Set v_rs = Server.CreateObject("ADODB.Recordset")
					Set v_rs.ActiveConnection = v_db
					v_sql = instrucao
					v_rs.CursorLocation = adUseClient
					v_rs.CursorType = adOpenForwardOnly
					v_rs.LockType = adLockReadOnly
					v_rs.Open v_sql
			case "normal"
				'Modo Normal, permite leitura, gravação, atualização e deletar
				
						v_sql = instrucao
						v_db.execute(v_sql)			

			end select	
	end function

	'-----------------------------------------------------------------------------------------
	'Remove Acentos - Paulo Corcino - pecjs@msn.com
	'Versão 1.1 - Favor não remover os creditos. 
	'-----------------------------------------------------------------------------------------
	
	dim CAcento, Pos_Acento, COR_Texto, SAcento, COR_X, COR_apos, Letra
	
	Function ac(modo,palavra)
	
	
	
	CAcento = "àáâãäèéêëìíîïòóôõöùúûüÀÁÂÃÄÈÉÊËÌÍÎÒÓÔÕÖÙÚÛÜçÇñÑŽŠžšŸÅÏås/.-~*=+-_^""']}[{`?/;:>,<\|!@#$%¨¨&()"
	SAcento = "aaaaaeeeeiiiiooooouuuuAAAAAEEEEIIIOOOOOUUUUcCnNZSzsYAIas                                     "
	COR_Texto = ""
	
	if modo = "completo" then
		
	if palavra <> "" then
	For COR_X = 1 to Len(SAcento)
	Letra = mid(palavra,COR_X,1)
	Pos_Acento = inStr(CAcento,Letra)
	if COR_Texto = "" then
	COR_Texto = replace(palavra, mid(CAcento,COR_X,1), mid(SAcento,COR_X,1))
	else
	COR_Texto = replace(COR_Texto, mid(CAcento,COR_X,1), mid(SAcento,COR_X,1))
	end if 
	next
	end if
	
	end if
	
	if modo = "senha" then
		COR_apos = ""
	end if
	
	if modo = "html" then
		COR_apos = "&#39;"
	end if
	
	'Corrige o problema do acento apostrofo (') em Bancos de Dados
	'Escolha como deve aparecer o apostrofo. ë necessário que um esteja ativado, 
	'para ativar basta remover o apostrofo antes do COR_apos, e desative o outro
	
	if COR_Texto = "" then
	COR_Texto = Replace(palavra,"'",COR_apos)
	else
	COR_Texto = Replace(COR_Texto,"'",COR_apos)
	end if
	
	ac = COR_Texto
		
	
	end function
	'----------------------------------------------------------------------------------------
'Função de saida anti-erro
	function invalido(cessao,db,r)	
		'voltar a página
		voltar request.ServerVariables("SCRIPT_NAME")&cessao,"","_self"
						
		'encerra banco de dados
		if db=1 then
		fdb(r)
		end if
		fdb(conn)		
		'Cancelar processamento da página, isso evitará erros
		response.end()
	end function

'Deletar arquivos do servidor
	dim fso, fsotxt
	function delfile(file)
		
		SET fso = Server.CreateObject("Scripting.FileSystemObject")
		
		If fso.FileExists(server.MapPath(file)) Then
			fso.DeleteFile(server.MapPath(file))
		End If
		
		SET fso = nothing
			
	end function
	
	'Criar Pastas
		function md(root)
			Set fso = CreateObject("Scripting.FileSystemObject")
			
			if not fso.FolderExists(Server.mapPath(root)) then
				fso.CreateFolder(Server.mapPath(root))
			end if
			
			Set fso = Nothing
		end function
	
	'Deletar Folder
		function rd(root)
			Set fso = CreateObject("Scripting.FileSystemObject")
			
			if fso.FolderExists(Server.mapPath(root)) then
				fso.DeleteFolder(Server.mapPath(root))
			end if
			
			Set fso = Nothing
		end function

	'Renomear arquivo
		function renfile(file,newfile)
		
			Set fso = CreateObject("Scripting.FileSystemObject") 
	
			if fso.FileExists(server.MapPath(file)) Then
			   fso.CopyFile server.MapPath(file), server.MapPath(newfile)
			   fso.DeleteFile(server.MapPath(file))
			end if
			
			Set fso = nothing
			
		end function
	
	'Renomear arquivo
		function copyfile(file,newfile)
		
			Set fso = CreateObject("Scripting.FileSystemObject") 
	
			if fso.FileExists(server.MapPath(file)) Then
			   fso.CopyFile server.MapPath(file), server.MapPath(newfile)
			end if
			
			Set fso = nothing
			
		end function
	
'cria arquivo de texto
function criatxt(files, news, cont)
	
	Const ForReading = 1, ForWriting =2, ForAppending =8
	Const TristateUsedefault = -2 
	Const TristateTrue = -1
	Const TristateFalse = 0
	
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(server.MapPath(files)) = true then
		if news = 1 then
			fso.DeleteFile(server.MapPath(files))
			Set fsotxt = fso.CreateTextFile(server.MapPath(files),false,false)
			'fsotxt.WriteLine "<!-- DataMatrixs2004 vr. 2.5 - pecjs@msn.com -->"&vbcrlf
			fsotxt.WriteLine cont&vbcrlf
			
			fsotxt.close
			set fsotxt = nothing
			
		elseif news = 2 then
			if fso.FileExists(server.MapPath(files)) = false then
				Set fsotxt = fso.CreateTextFile(server.MapPath(files),false,false)
				'fsotxt.WriteLine "<!-- DataMatrixs2004 vr. 2.5 - pecjs@msn.com -->"&vbcrlf
				fsotxt.WriteLine cont&vbcrlf
				
				fsotxt.close
				set fsotxt = nothing
				
			end if			
		else
			set fsotxt = fso.OpenTextFile(server.MapPath(files), ForAppending, False, Tristatefalse)
			'fsotxt.WriteLine "<!-- DataMatrixs2004 vr. 2.5 - pecjs@msn.com -->"&vbcrlf
			fsotxt.WriteLine cont&vbcrlf
				
			fsotxt.close
			set fsotxt = nothing
			
		end if 
	else
		Set fsotxt = fso.CreateTextFile(server.MapPath(files),false,false)
		'fsotxt.WriteLine "<!-- DataMatrixs2004 vr. 2.5 - pecjs@msn.com -->"&vbcrlf
		fsotxt.WriteLine cont&vbcrlf
		
		fsotxt.close
		set fsotxt = nothing
		
	end if
	
	

	set fso = nothing
	
end function
'Leitura de arquivos de Texto
function txtleia(file)

	set fso = Server.CreateObject("Scripting.FileSystemObject")
	set fsotxt = fso.OpenTextFile(server.MapPath(file))
	
	'while not fsotxt.eof
		txtleia = fsotxt.ReadAll

	'wend 
	
	
	fsotxt.close
	set fsotxt = nothing
	set fso = nothing
	
end function
	
'redimensionando a imagem
dim imga
function jpeg(arquivo,salvar,im,prop,larg,alt,quali,sharp)
			imga = server.MapPath(arquivo&im)
			Set jpeg = Server.CreateObject("Persits.Jpeg")
			
			jpeg.Open(imga)
			
			select case prop
			case "original"
			'if prop = "original" then
				jpeg.Width = jpeg.OriginalWidth '160
				jpeg.Height = jpeg.OriginalHeight  'jpeg.OriginalHeight * jpeg.Width / 
			'elseif prop="prop" then
			case "prop"
			
				if alt <> "0" then
					
					jpeg.Height = alt 'jpeg.OriginalHeight * jpeg.Width / jpeg.OriginalWidth
					jpeg.Width = jpeg.OriginalWidth * jpeg.Height / jpeg.OriginalHeight '160
					
				elseif larg<>"0" then
				
					jpeg.Width = larg '160
					jpeg.Height = jpeg.OriginalHeight * jpeg.Width / jpeg.OriginalWidth 'jpeg.OriginalHeight * jpeg.Width / jpeg.OriginalWidth
				
				end if
			case "autored"
			
				if cint(jpeg.OriginalHeight)>cint(jpeg.OriginalWidth) then
					'imagem horizontal
					
					jpeg.Height = alt 'jpeg.OriginalHeight * jpeg.Width / jpeg.OriginalWidth
					jpeg.Width = jpeg.OriginalWidth * jpeg.Height / jpeg.OriginalHeight '160
					
				else
					'imagem vertical
					jpeg.Width = larg '160
					jpeg.Height = jpeg.OriginalHeight * jpeg.Width / jpeg.OriginalWidth 'jpeg.OriginalHeight * 
					
				end if
			
			case else
				jpeg.Width = larg '160
				jpeg.Height = alt 'jpeg.OriginalHeight * jpeg.Width / jpeg.OriginalWidth
			end select
			
			if sharp <> "0" or sharp > "100" then
				Jpeg.Sharpen 1, sharp '130
			end if
			
			Jpeg.Interpolation = 1
			Jpeg.Quality = quali '40
			
			Jpeg.Save server.MapPath(salvar&im)
			
			set jpeg = nothing
			
end function


dim  Jpeg_zoeira_Logo,Jpeg_zoeira_LogoPath
function jpeg_zoeira(arquivo,salvar,im,prop,larg,alt,quali,sharp)
			imga = server.MapPath(arquivo&im)
			Set jpeg_zoeira = Server.CreateObject("Persits.Jpeg")
			
			jpeg_zoeira.Open(imga)
			
			if prop = "original" then
				jpeg_zoeira.Width = jpeg_zoeira.OriginalWidth '160
				jpeg_zoeira.Height = jpeg_zoeira.OriginalHeight  'jpeg.OriginalHeight * jpeg.Width / 
			elseif prop="prop" then
			
				if alt <> "0" then
					
					jpeg_zoeira.Height = alt 'jpeg.OriginalHeight * jpeg.Width / jpeg.OriginalWidth
					jpeg_zoeira.Width = jpeg_zoeira.OriginalWidth * jpeg_zoeira.Height / jpeg_zoeira.OriginalHeight '160
					
				elseif larg<>"0" then
				
					jpeg_zoeira.Width = larg '160
					jpeg_zoeira.Height = jpeg_zoeira.OriginalHeight * jpeg_zoeira.Width / jpeg_zoeira.OriginalWidth 'jpeg.OriginalHeight * jpeg.Width / jpeg.OriginalWidth
				
				end if
			
			else
				jpeg_zoeira.Width = larg '160
				jpeg_zoeira.Height = alt 'jpeg.OriginalHeight * jpeg.Width / jpeg.OriginalWidth
			end if
			
			if sharp <> "0" or sharp > "100" then
				Jpeg_zoeira.Sharpen 1, sharp '130
			end if
			
			Jpeg_zoeira.Interpolation = 1
			Jpeg_zoeira.Quality = quali '40
			
'			Jpeg_zoeira.Canvas.Pen.Color = &HFF6600
'			Jpeg_zoeira.Canvas.Pen.Width = 40
'			Jpeg_zoeira.Canvas.Line 0, Jpeg_zoeira.Height, Jpeg_zoeira.Width, Jpeg_zoeira.Height
'			
'			Jpeg_zoeira.Canvas.Pen.Color = &HFF6600' black
'			Jpeg_zoeira.Canvas.Pen.Width = 2
'			Jpeg_zoeira.Canvas.Brush.Solid = False ' or a solid bar would be drawn
'			Jpeg_zoeira.Canvas.Bar 1, 1, Jpeg_zoeira.Width, Jpeg_zoeira.Height
'			

			
			Set Jpeg_zoeira_Logo = Server.CreateObject("Persits.Jpeg")
			
			Jpeg_zoeira_LogoPath = Server.MapPath("/bymidia/img/") & "\logos.gif"
			
			Jpeg_zoeira_Logo.Open Jpeg_zoeira_LogoPath
			
			Jpeg_zoeira.DrawImage 0, Jpeg_zoeira.Height-100, Jpeg_zoeira_Logo, , &H0000FF 
			
			Jpeg_zoeira.Canvas.Font.Color = &HFFFFFF' red
			Jpeg_zoeira.Canvas.Font.Family = "Tahoma"
			Jpeg_zoeira.Canvas.Font.Bold = True
			Jpeg_zoeira.Canvas.Font.Size = 14
			Jpeg_zoeira.Canvas.Print 65, Jpeg_zoeira.Height-18, "w w w . t o n a z o e i r a . c o m . b r"
			
			Jpeg_zoeira.Save server.MapPath(salvar&im)
			
			set jpeg_zoeira = nothing
			
end function


'Data por periodo

	Dim dia_hoje,mese,anno,tmese,day1,primog,day2, mese2, anno2, dayoff, ultimog, dia_atual_
	  
	function datas(estilo)
	
	 'Declarando variaveis 
	 	 dia_hoje = day(now())
	 	 mese = month(now())
	 	 anno = year(now())


	  If mese = 12 then
      	mese2 = 1
	  	anno2 = anno+1
	  else 
	 	 mese2 = mese+1
	  	anno2 = anno
	  End If
	   

	  'corrige os dia 31
	  'devido que nem todos os meses tem dia 31 este código consegue detectar o dia e pula para o proximo.
	  Dim a, b, c
	  a = day(now)
	  if  a = 31 then
	 	 b = 30
	 	 c = 1
	  else
		 b = a
	 	 c = 0
	  end if
	  
	  
	  day1 = Cdate(day(now)&"/"+Cstr(mese)+"/"+Cstr(anno))+c
	  ' b indica o numero de dias que vai pular, por exemplo hoje é 1/1/2003 se o valor de b for 3 a data inicial será 4/1/2003
      dayoff = Cdate(b&"/"+Cstr(mese2)+"/"+Cstr(anno2))-1
      primog = Weekday(Cdate(day(now)&"/"+Cstr(mese)+"/"+Cstr(anno)))-1
	  ultimog = Weekday(dayoff)-1
	  
	  If ultimog = 0 then
	  ultimog = 7
	  End If
	  
	  If primog = 0 then
	  primog = 7
	  End If

	  x=1
		dia_atual_ = (day1-primog+x)
		Do While dia_atual_ <= (dayoff+7-ultimog)
		if dia_atual_ >= day1 and dia_atual_ <= dayoff then
		
			select case weekday(day1-primog+x)
		'abaixo veja o uso das variaveis da semana
		' se o dia naum for indicado o seu valor , ele não será utilisado pelo case
		'case dom,seg,ter,qua,qui,sex,sab
			case 1,2,3,4,5,6,7
		
				select case estilo
					case "texto"
						response.write day(day1-primog+x)&"/"&month(day1-primog+x)&"/"&year(day1-primog+x)&"<br>"&vbcrlf
					case "lista"
						response.write "<option value="""&month(day1-primog+x)&"/"&day(day1-primog+x)&"/"&year(day1-primog+x)&""">"&day(day1-primog+x)&"/"&month(day1-primog+x)&"/"&year(day1-primog+x)&"</option>"&vbcrlf
					end select

			end select
		End If
		
			dia_atual_ = dia_atual_ + 1
			x=x+1
			loop
	end function

'Função para incluir conteudo em uma cessão é necessário estar utilizando o sistema de conteudo
	function conteudo(numero)
	
		dim conteudo_rs,conteudo_sql
		db "leitura",conteudo_rs,conteudo_sql,"select texto from matrix_cont where id="&numero,conn
		response.write cHTML(conteudo_rs("texto"))
		fdb(conteudo_rs)
		
	end function

'Função Data por extensso
	function hoje(cidade)
			response.Write cidade&",&nbsp;"&day(now)&"&nbsp;de&nbsp;"&mesext(month(now))&"&nbsp;de&nbsp;"&year(now)
	end function

'Marca de Desenvolvimento
	dim criadoano

	function marca(ano)
		ano = cint(ano)
	
			if year(now)=ano then
				criadoano = ano
			else
				criadoano = ano&"&nbsp;-&nbsp;"&year(now)
			end if
			
			response.Write "<b>Paulo&nbsp;Corcino</b>&nbsp;||&nbsp;"&criadoano&",&nbsp;cdlestancia.com.br&nbsp;.&nbsp;E-Mail:&nbsp;pecjs@msn.com"
			
	end function

'Contador
function contador(pagina,id)
response.write " <script type=""text/javascript"" language=""JavaScript"">"&vbcrlf
 
response.write "var file='"&site&"/matrix/contador/count.asp';"&vbcrlf
response.write "var d=new Date(); "&vbcrlf
response.write "var s=d.getSeconds(); "&vbcrlf
response.write "var m=d.getMinutes();"&vbcrlf
response.write "var x=s*m;"&vbcrlf
'response.write "function contador(pagina){"&vbcrlf
response.write "f='' + escape(document.referrer);"&vbcrlf
response.write "if (navigator.appName=='Netscape'){b='NS';} "&vbcrlf
response.write "if (navigator.appName=='Microsoft Internet Explorer'){b='MSIE';} "&vbcrlf
response.write "if (navigator.appVersion.indexOf('MSIE 3')>0) {b='MSIE';}"&vbcrlf
response.write "u='' + escape('"&pagina&"'); w=screen.width; h=screen.height; "&vbcrlf
response.write "v=navigator.appName; "&vbcrlf
response.write "ip='"&Request.ServerVariables("REMOTE_ADDR")&"'"&vbcrlf
response.write "fs = window.screen.fontSmoothingEnabled;"&vbcrlf
response.write "if (v != 'Netscape') {c=screen.colorDepth;}"&vbcrlf
response.write "else {c=screen.pixelDepth;}"&vbcrlf
response.write "j=navigator.javaEnabled();"&vbcrlf
response.write "info='cods="&id&"&ip='+ip+'&w=' + w + '&h=' + h + '&c=' + c + '&r=' + f + '&u='+ u + '&fs=' + fs + '&b=' + b + '&x=' + x;"&vbcrlf
response.write "document.write('img src='+file + '?'+info+'  width=1 height=1 border=0>')"&vbcrlf
'response.write "document.write('oi')"&vbcrlf
'response.write "}"&vbcrlf

response.write "</script>"&vbcrlf
end function


'Função para redução de código Request.form("valor") para formu("valor") usa-se: formu("teste",1)
dim formu_v
function formus(forms,filtro)
'Filtra os caracteres inválidos para banco de dados
	select case filtro
	case 1
	'Filtra caracteres
		formu_v = replace(request.form(forms),"'","&#39")
		formu_v = replace(formu_v,"""","&#34")
		formu_v = replace(formu_v,"*","&#42")
		formu_v = replace(formu_v,"<","&#60")
		formu_v = replace(formu_v,">","&#62")
		formu_v = replace(formu_v,"[n]","<b>")
		formu_v = replace(formu_v,"[/n]","</b>")
		formu_v = replace(formu_v,"[i]","<i>")
		formu_v = replace(formu_v,"[/i]","</i>")
		formu_v = replace(formu_v,"[s]","<u>")
		formu_v = replace(formu_v,"[/s]","</u>")
		formu_v = replace(formu_v,chr(10),"<br>")
	case 2
	'Filtra caracteres para sistema upload
		formu_v = replace(objUpload.Form(forms),"'","&#39")
		formu_v = replace(formu_v,"""","&#34")
		formu_v = replace(formu_v,"*","&#42")
		formu_v = replace(formu_v,"<","&#60")
		formu_v = replace(formu_v,">","&#62")
		formu_v = replace(formu_v,"[n]","<b>")
		formu_v = replace(formu_v,"[/n]","</b>")
		formu_v = replace(formu_v,"[i]","<i>")
		formu_v = replace(formu_v,"[/i]","</i>")
		formu_v = replace(formu_v,"[s]","<u>")
		formu_v = replace(formu_v,"[/s]","</u>")
		formu_v = replace(formu_v,chr(10),"<br>")
	case 3
		formu_v = request.QueryString(forms)
		formu_v = replace(formu_v,"'","&#39")
	case else
		formu_v = request.form(forms)
		formu_v = replace(formu_v,"'","&#39")
	end select
	
	formus = formu_v
end function

'converte de  html para acsii
function rstxt(rcd,forms)


		formu_v = replace(rcd(forms),"&#39","'")
		formu_v = replace(formu_v,"&#34","""")
		formu_v = replace(formu_v,"&#42","*")
		formu_v = replace(formu_v,"&#60","<")
		formu_v = replace(formu_v,"&#62",">")
		formu_v = replace(formu_v,"<b>","[n]")
		formu_v = replace(formu_v,"</b>","[/n]")
		formu_v = replace(formu_v,"<i>","[i]")
		formu_v = replace(formu_v,"</i>","[/i]")
		formu_v = replace(formu_v,"<u>","[s]")
		formu_v = replace(formu_v,"</u>","[/s]")
		formu_v = replace(formu_v,"<br>",chr(10))

		
		rstxt = formu_v
end function
function rsv(rcd,forms)


		formu_v = replace(rcd(forms),chr(10),"<br>")
		formu_v = replace(formu_v,chr(32),"")

		rsv = formu_v
end function
function rtxt(rcd,forms)


		formu_v = replace(rcd(forms),">","--> ")
		formu_v = replace(formu_v,"<"," <!--")
		
		rtxt = formu_v
end function

dim num_col, num_lin
num_col = 4

'Alert em javascript
dim javaalert
	function alert(msg)
		response.write "<script>"&vbcrlf
		response.write "//ASP Alerta - Paulo Corcino pecjs@msn.com"&vbcrlf
		msg = replace(msg,"\","\\")
		msg = replace(msg,"&#39;","\'")
		msg = replace(msg,"'","\'")
		msg = replace(msg,"""","\""")
		msg = replace(msg,"[n]","\n")
		response.write "alert('"&msg&"')"&vbcrlf
		response.Write "</script>"&vbcrlf
	end function
'----------------------------------------------------------------------
Function Encrypt(Message, EncryptionKey)
		Dim Temp
		Dim Num
		Dim i
		Dim Char
		Message = RC4(Message,EncryptionKey)
		Temp=""
		For i = 1 to Len(Message)
			Char = Mid(Message,i,1)
			Num = Asc(Char)
			If i <> 1 Then
				Temp = Temp & "." & Num
			Else
				Temp = Temp & Num
			End If
		Next
		Encrypt=Temp
End Function
Function Decrypt(Message, EncryptionKey)
	Dim Nums
	Dim i
	Dim Temp
	Temp=""
	Nums = Split(Message,".",-1,1)
	For i = 0 to uBound(Nums)
		Temp = Temp & chr(Nums(i))
	Next
	Message = RC4(Temp,EncryptionKey)				
	Decrypt = Message
End Function
'-----------------------------------------------------------------
'All Credit for the Code below goes to Lewis Moten (http://www.lewismoten.com)
'Check out his submission to Planet-Source-Code for more info
'Lewis Moten's RC4 Encryption/Decryption Algorithm code:
Function RC4(ByRef pStrMessage, ByRef pStrKey)

	Dim lBytAsciiAry(255)
	Dim lBytKeyAry(255)
	Dim lLngIndex
	Dim lBytJump
	Dim lBytTemp
	Dim lBytY
	Dim lLngT
	Dim lLngX
	Dim lLngKeyLength
	
	' Validate data
	If Len(pStrKey) = 0 Then Exit Function
	If Len(pStrMessage) = 0 Then Exit Function

	' transfer repeated key to array
	lLngKeyLength = Len(pStrKey)
	For lLngIndex = 0 To 255
	    lBytKeyAry(lLngIndex) = Asc(Mid(pStrKey, ((lLngIndex) Mod (lLngKeyLength)) + 1, 1))
	Next

	' Initialize S
	For lLngIndex = 0 To 255
	    lBytAsciiAry(lLngIndex) = lLngIndex
	Next

	' Switch values of S arround based off of index and Key value
	lBytJump = 0
	For lLngIndex = 0 To 255
	
		' Figure index to switch
	    lBytJump = (lBytJump + lBytAsciiAry(lLngIndex) + lBytKeyAry(lLngIndex)) Mod 256
	    
	    ' Do the switch
	    lBytTemp				= lBytAsciiAry(lLngIndex)
	    lBytAsciiAry(lLngIndex)	= lBytAsciiAry(lBytJump)
	    lBytAsciiAry(lBytJump)	= lBytTemp
	    
	Next

	
	lLngIndex = 0
	lBytJump = 0
	For lLngX = 1 To Len(pStrMessage)
	    lLngIndex = (lLngIndex + 1) Mod 256 ' wrap index
	    lBytJump = (lBytJump + lBytAsciiAry(lLngIndex)) Mod 256 ' wrap J+S()
	    
		' Add/Wrap those two	    
	    lLngT = (lBytAsciiAry(lLngIndex) + lBytAsciiAry(lBytJump)) Mod 256
	    
	    ' Switcheroo
	    lBytTemp				= lBytAsciiAry(lLngIndex)
	    lBytAsciiAry(lLngIndex)	= lBytAsciiAry(lBytJump)
	    lBytAsciiAry(lBytJump)	= lBytTemp

	    lBytY = lBytAsciiAry(lLngT)
	
		' Character Encryption ...    
	    RC4 = RC4 & Chr(Asc(Mid(pStrMessage, lLngX, 1)) Xor lBytY)
	Next
	
End Function
'------------------------------------
'criptografia
Dim sbox(255)
   Dim key(255)
   
   ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
   ':::                                                             :::
   ':::  This script performs 'RC4' Stream Encryption               :::
   ':::  (Based on what is widely thought to be RSA's RC4           :::
   ':::  algorithm. It produces output streams that are identical   :::
   ':::  to the commercial products)                                :::
   ':::                                                             :::
   ':::  This script is Copyright © 1999 by Mike Shaffer            :::
   ':::  ALL RIGHTS RESERVED WORLDWIDE                              :::
   ':::                                                             :::
   ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

   Sub RC4Initialize(strPwd)
   ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
   ':::  This routine called by EnDeCrypt function. Initializes the :::
   ':::  sbox and the key array)                                    :::
   ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

      dim tempSwap
      dim a
      dim b
	  dim intLength

      intLength = len(strPwd)
      For a = 0 To 255
         key(a) = asc(mid(strpwd, (a mod intLength)+1, 1))
         sbox(a) = a
      next

      b = 0
      For a = 0 To 255
         b = (b + sbox(a) + key(a)) Mod 256
         tempSwap = sbox(a)
         sbox(a) = sbox(b)
         sbox(b) = tempSwap
      Next
   
   End Sub
   
   Function EnDeCrypt(plaintxt, psw)
   ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
   ':::  This routine does all the work. Call it both to ENcrypt    :::
   ':::  and to DEcrypt your data.                                  :::
   ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

      dim temp
      dim a
      dim i
      dim j
      dim k
      dim cipherby
      dim cipher

      i = 0
      j = 0

      RC4Initialize psw

      For a = 1 To Len(plaintxt)
         i = (i + 1) Mod 256
         j = (j + sbox(i)) Mod 256
         temp = sbox(i)
         sbox(i) = sbox(j)
         sbox(j) = temp
   
         k = sbox((sbox(i) + sbox(j)) Mod 256)

         cipherby = Asc(Mid(plaintxt, a, 1)) Xor k
         cipher = cipher & Chr(cipherby)
      Next

      EnDeCrypt = cipher

   End Function
   
   	
	
	Private Function lZeros(byVal lValue, byVal iLength) 
		lZeros = Right(String(iLength, "0") & Trim(cStr(lValue)), iLength) 
	End Function 

Private Function fDate(byVal oDate, byVal sFormat) 
	Dim iDay, iMonth, iYear 

	iDay = DatePart("d", oDate) 
	iMonth = DatePart("m", oDate) 
	iYear = DatePart("yyyy", oDate) 

		fDate = Replace(lcase(sFormat), "w", WeekDayName(DatePart("w", oDate))) 
		fDate = Replace(fDate, "hh", lZeros(DatePart("h", oDate),2)) 
		fDate = Replace(fDate, "nn", lZeros(DatePart("n", oDate),2)) 
		fDate = Replace(fDate, "ss", lZeros(DatePart("s", oDate),2)) 
		fDate = Replace(fDate, "dd", lZeros(iDay,2)) 
		fDate = Replace(Replace(fDate,"mmmm", MonthName(iMonth)),"mmm", MonthName(iMonth, True)) 
		fDate = Replace(fDate, "mm", lZeros(iMonth,2)) 
		fDate = Replace(Replace(fDate, "yyyy", iYear), "yy", right(cStr(iYear),2)) 
End Function 

'Gera aspid

	'function aspid(val,op)
	
		'select case op
		'case 1
			'criptografia
			'aspid = server.URLEncode(EnDeCrypt(val,fDate(date(),"ddmmyyyy")))
		'case 2
			'aspid = request.QueryString("aspid")
		'case else
			'descriptografar
			'aspid = EnDeCrypt(val,fDate(date(),"ddmmyyyy")) 
		'end select
		
	'end function
	
	function aspid()
	
		aspid = replace(request.QueryString("aspid"),"'","''")
				
	end function
	
	
	'Xingamento ================================================================
	Const BAD_WORDS_FILTER = "bicha, buceta, xoxota, viado, puta, cu, pica, cunhão, xibiu, chibiu, boquete, boquet, boket, pedófilo, pedofilo, cuzinho, cusinho, ku, fdp, caralho, corno, corninho, peste, fuder, putinha, fudida, maconheiro, MACONHEIRO, cú, Cú, "
	
	Function padBadWordWithAsterisks(s)
		If (Len(s) < 3) Then
			padBadWordWithAsterisks = s
			Exit Function
		End If
		
		padBadWordWithAsterisks = Left(s, 1) & String(Len(Mid(s, 2, Len(s)-2)), "*") '& Right(s, 1)
		'padBadWordWithAsterisks = String(Len(Mid(s, 2, Len(s))), "*")
		
	End Function
	
	
	'
	' Replace bad words from a string with first and last character in word 
	' being replaced.
	'
	Function replaceBadWords(s, badWordsFilter)
		
		Dim badWords, badWord
		badWords = Split(badWordsFilter, ",")
		For Each badWord In badWords
			badWord = Trim(badWord)
			
			Dim regEx
			Set regEx = New RegExp
			regEx.Pattern = badWord
			regEx.IgnoreCase = True
			regEx.Global = True
			s = regEx.Replace(s, padBadWordWithAsterisks(badWord))
			
		Next
		
		replaceBadWords = s
		
	End Function	
	
	'=======================================================================================

'Fatiamento de Texto=======================================================
 dim SplitLen, SplitY, SplitX, SplitWord, SplitWords
 function splittxt(i, msg)
	'separa o texto apartir dos espaços
	SplitWord = Split(msg,chr(32))
	for each SplitWords in SplitWord
		'caso a palavra tenha um tamanho acima do aceitavel no site ele fatia o texto
		if cint(len(SplitWords))>=cint(i) then
		
			SplitLen=len(SplitWords)
			SplitY = 1
			
			for SplitX = 1 to SplitLen
			
				splittxt = splittxt + mid(SplitWords,SplitX,1)
				if SplitY < i then
					SplitY = SplitY + 1
				else
					splittxt = splittxt & " "
					SplitY = 1
				end if
				
			next
		else
			'caso contrario o texto é mostrado sem fatiamento
			
			splittxt = splittxt & trim(SplitWords) & " "
		
		end if
	
	next 

 
 end function
 '======================================================================
 
	
	Function QuantosDiasTemOMes(Mes,Ano)
  Select Case Mes
    Case 1,3,5,7,8,10,12: QuantosDiasTemOMes = 31
    Case 4,6,9,11: QuantosDiasTemOMes = 30
    Case Else
      If Ano Mod 4 = 0 And (Ano Mod 100 <> 0 Or Ano Mod 400 = 0) Then
        QuantosDiasTemOMes = 29
      Else
        QuantosDiasTemOMes = 28
      End If
  End Select
End Function

dim nickv,rsns,sqlns
function nicks(asp)
	db "leitura",rsns,sqlns,"SELECT     dbo.cad_auditoria.aspid, dbo.cad_list.nick, dbo.cad_list.id FROM  dbo.cad_auditoria INNER JOIN  dbo.cad_list ON dbo.cad_auditoria.id_u = dbo.cad_list.id WHERE (dbo.cad_auditoria.aspid = '"&asp&"')",conn
	nicks = trim(rsns("nick"))
	fdb(rsns)
end function

	dim br
	br = chr(10)

'//---- Anivesário ------------------------------
	dim NiverId, NiverRS, NiverRS2, NiverSql, NiverData, NiverNick, NiverEmail, NiverMsg
	function Matrixniver(dbo)
	
		db "leitura",NiverRS,NiverSql,"SELECT TOP 100 PERCENT id, email, nick, msg_niver, atv FROM dbo.cad_list WHERE (criado = 1) AND (MONTH(datan) = MONTH(GETDATE())) AND (DAY(datan) = DAY(GETDATE())) AND (atv = 1) ORDER BY nick DESC",dbo
		while not NiverRS.eof 
			
			NiverId = NiverRS("id")
			NiverData = NiverRS("msg_niver")
			NiverNick = chtml(trim(NiverRS("nick")))
			NiverEmail = trim(NiverRS("email"))
			
			Matrixniver = Matrixniver & NiverNick & "<br>"
			
			
			
			if not isdate(NiverRS("msg_niver")) then
			
				NiverData = Dateadd("d",-1,date())
			
			end if
			
			if fDate(NiverData,"dd/mm/yyyy")<>fDate(now(),"dd/mm/yyyy") then
			
				db "leitura",NiverRS2, NiverSql, "select top 1 msg from cad_nivermsg where criado=1 order by newid()",dbo
				
				if not NiverRS2.eof and not NiverRS2.bof then 
					
					NiverMsg = replace(NiverRS2("msg"),"[Nome]",NiverNick)
					NiverMsg = replace(NiverMsg,"[nome]",NiverNick)
					
					
					emails "tonazoeira@hotmail.com",NiverEmail,"ToNaZoeira - Feliz Anivesário!!!",NiverMsg,"html"
								
					db "normal",,NiverSql,"update cad_list set msg_niver='"&fdate(now(),"yyyymmdd hh:nn")&"' where id="&NiverId,dbo
				
				end if
				
				fdb(NiverRS2)
				
			end if
				
			NiverRS.movenext
			
		wend
		
		if NiverRS.eof and NiverRS.bof then
		
			Matrixniver = ""
			
		end if 
		
		fdb(NiverRS)
		
	end function

'//----------------------------------------------
'//----------------- Cadastro Limpa Não Cadastro
	
	dim CadRs, CadSql, CadFile, CadFilen
	
	function cadclear()
	
		CadFile = roots&"/cadlog-"&fDate(date(),"dd-mm-yyyy")&".txt"
		CadFilen = roots&"/cadlog-"&fDate(dateadd("d",-1,date()),"dd-mm-yyyy")&".txt"
		
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		If fso.FileExists(server.MapPath(CadFile)) = false then
			
			db "normal",,CadSql,"delete FROM dbo.cad_list WHERE (atv <> 1) AND (MONTH(data_del) = "&month(now)&") AND (DAY(data_del) = "&day(now)&") AND (YEAR(data_del) = "&year(now)&")",conn
			
			criatxt CadFile, 1, "Registro Limpo - "& now()
			delfile CadFilen
				
		end if
	end function
'//-----------------------------------------------------
'//----------------- Lista Estados
	dim UfList, UfNome, UfNomes, UfSelect

	function uf(selected, tp)
		
		if tp=1 then
			UfList = "AC, AL, AM, BA, CE, DF, ES, GO, MA, MG, MS, MT, PA, PB, PE, PI, PR, RJ, RN, RO, RR, RS, SC, SE, SP, TO"
		else
			UfList = "Acre, Alagoas, Amazônia, Bahia, Ceará, Distrito Federal, Espiríto Santo, Goiaís, Maranhão, Minas Gerais, Mato Grosso do Sul, Mato Grosso, Pará, Paraíba, Pernambuco, Piauí, Paraná, Rio de Janeiro, Rio Grande do Norte, Rondônia, Roraima, Rio Grande do Sul, Santa Catarina, Sergipe, São Paulo, Tocantins"
		end if
		
	
		UfNomes = Split(UfList, ",")
		For Each UfNome In UfNomes
		
			UfNome = Trim(UfNome)
			
			if UfNome = selected then
				UfSelect = "selected"
			else
				UfSelect = ""
			end if
			
			uf = uf & "<option value='"&UfNome&"' "&UfSelect&">"&UfNome&"</option>"&vbcrlf
		Next
	
	end function

'//-----------------------------------------------------
			
				
'Fiiltro JavaScript
function filtrojava(txt)
		filtrojava = txt
		filtrojava = replace(filtrojava,"\","\\")
		filtrojava = replace(filtrojava,"&#39;","\'")
		filtrojava = replace(filtrojava,"'","\'")
		filtrojava = replace(filtrojava,"\'+","'+")
		filtrojava = replace(filtrojava,"+\'","+'")
		filtrojava = replace(filtrojava,"""","\""")
		filtrojava = replace(filtrojava,"[n]","\n")
	
end function

'//-----------------------------------------------------
' E-Mail sistema
'-------------------------------------------------------
dim SMailC
function SysMail(email,msg)
	SMailC = SMailC &"<pre>"&vbcrlf
	SMailC = SMailC & "=================================================================================="&vbcrlf
	SMailC = SMailC & "                      - - -  ToNaZoeira - - -"&vbcrlf
	SMailC = SMailC & "=================================================================================="&vbcrlf
	SMailC = SMailC & "Data: "&fdate(now,"dd/mm/yyyy")&" - Hora: "&fdate(now,"hh:nn")&vbcrlf
	SMailC = SMailC & "**********************************************************************************"&vbcrlf
	SMailC = SMailC & "Mensagem:_________________________________________________________________________"&vbcrlf&vbcrlf
	'SMailC = SMailC & splitline(83, chr(10), replace(msg,"/n",chr(10)))&vbcrlf
	SMailC = SMailC & replace(msg,"/n",chr(10))&vbcrlf
	SMailC = SMailC & "__________________________________________________________________________________"&vbcrlf
	SMailC = SMailC & "=================================================================================="&vbcrlf
	SMailC = SMailC & "             Sistema DataMatrix v.3 / Webmaster: Paulo Corcino - "&year(now)&vbcrlf
	SMailC = SMailC & "=================================================================================="&vbcrlf
	SMailC = SMailC & "</pre>"
	SysMail = SMailC
	'response.Write SMailC
	emails email_tec,email,"ToNaZoeira - Informe",SMailC,"html"
end function
'=======================================================
'Texto de Boas Vindas
'=======================================================
function Boas()
	if hour(now)>=19 then
		Boas = "Boa Noite"
	elseif hour(now)>=12 then
		Boas = "Boa Tarde"
	else
		Boas = "Bom Dia"
	end if
end function
'//----------------------------------------------------------------->
'//                 *** Data Grid  *** v.1 - 10/02/2006
'//----------------------------------------------------------------->
'Documentação da manipulação de AspGrid.
'Criação de ASPGrid
'					
'					'//Declarar as variaveis é obrigatório
'					'//GridData(n) - Array onde indica o dados das colunas da tabela SQL
'					'//GridName(n) - Array onde indica os nomes das colunas a exibir
'					'//Obs.: n indica o numero de campos, deve ser igual nas duas.
'					'//GridSQL - variavel onde armazeará a query sql
'					'//GridDate  - variavel da formatação da data caso exista
'					Dim GridData(5), GridName(5), GridSQL, GridDate 
'				
'					'// rotina sql
'					GridSQL = "SELECT id, nome, nivel FROM senha"
'					
'					GridName(0) = "Cod;35" '// o valor estre as aspas "Cod;35" o primeiro valor COD é o nome da coluna que ira aparecer na tabela, o segundo valor é a sua largura em pixel
'					GridName(1) = "Nome;370"
'					GridName(2) = "Nivel;60"
'					GridName(3) = "Publicação;90"
'					GridName(4) = "Clicks;90"
'					GridName(5) = "Alterar;90"
'									
'					GridData(0) = "id;1;0" '// o primeiro valor indica o nome do campo a exibir coletado do sql, o segundo valor o alinhamento (0 - left, 1 - center, 2 - rigth), o terceiro valor um link, caso não tenha link insira o valor 0
'					GridData(1) = "nome;0;0"
'					GridData(2) = "$nivel;1;0; $case({nivel}=1):Master|$case({nivel}=2):Operador" ' Neste exemplo temos uma macro q analiza o valor a exibir e substitui por outro valor de equivalencia
'					'// Logica Macro
'					'// A coluna deve iniciar com o sinal $ para indicar a existencia de uma macro
'					'// Cria-se obrigatoriamente o terceiro campo onde existe a macro a ser executada
'					'// $case - indica a macro case onde faz a analize das situações
'					'//$case({nivel}=1) - caso o valor da coluna nivel for igual a 1
'					'o nome da coluna a ser analizada dever ser indicada da maneira apresentada {nomecol}
'					'//separdo por : temos o valor verdadeiro caso precise de um valor falso acrecente mais o sinal : exemplo($case({nivel}=1):verdadeiro:falso) 
'					'para incluir outra analize separe a primeira da segunda com o sinal de | exemplo( $case({nivel}=1):Master|$case({nivel}=2):Operador )
'					'As regras são as mesmas, e não existe limites para a analize
'					GridData(3) = "*Alterar;1;session.asp?op=alts&id=[col0]&aspid=[aspid]" 'nos links caso precise de um valor de uma determinada coluna indique desta maneira a coluna do valor necesario [col2]  o numero indica a coluna
'					GridData(4) = "*Editar Foto;1;fotos.asp?op=alt&id=[col0]&aspid=[aspid]"
'					GridData(5) = "*Deletar;1;javascript: deletar('"&request.ServerVariables("SCRIPT_NAME")&"?op=del&aspid=[aspid]&id=[col0]','[col1]')"
'							
'					GridDate = "dd/mm/yyyy"
'								
'					GridForm "ç",15


'//------------------------------------------------------------------>
	Dim Gv
	function GValue(n)
			Gv = replace(n,"[","")
			Gv = replace(Gv,"]","")
			Gv = trim(Gv)
			Gv = split(GridData(Gv),";")
			GValue = Gv(0)
	
	end function


Function GridASP(Gsql,Gbase,rp)
	
	
	Dim CONN_STRING
	Dim CONN_USER
	Dim CONN_PASS
	Dim iPageSize       
	Dim iPageCount      
	Dim iPageCurrent   
	Dim strOrderBy   
	Dim strSQL        
	Dim objPagingConn    
	Dim rs2
	Dim iRecordsShown 
	Dim Grid_semr
	Dim Grid_pgi,Grid_pgf, GFinal, GridMacro
	Dim Gridvar, GridX, Glink, Glinkx, GlinkVar, Gridvars, GField
	
	CONN_STRING = Gbase
	CONN_USER = ""
	CONN_PASS = ""
	Grid_semr = 0
	Grid_pgi = empty
	Grid_pgf = empty
'Dim I               

' Numero de Resultados por página
	if rp = 0 then
		iPageSize = 10000000000
	else
		iPageSize = rp
	end if
	
	

			if GridDate="" then
				GridDate = "dd/mm/yyyy hh:nn"
			end if

	If Gpage = "" Then
		iPageCurrent = 1
	Else
		iPageCurrent = Gpage
	End If
	
	
	
	strSQL = Gsql
	
	if request.form("orderby")<>"" then
		strSQL = strSQL & " order by "&formus("orderby",3)
		if request.form("order")="d" then
			strSQL = strSQL & " desc"
			GridString = GridString & "&order=a"
		else
			GridString = GridString & "&order=d"
		end if	
	end if 'WHERE     (dbo.album_album.id_ar = "&request.QueryString("ar")&")"
	'strSQL = "SELECT     id, id_ar, dest, quando FROM         dbo.album_album"
	'response.write strSQL 
	Set objPagingConn = Server.CreateObject("ADODB.Connection")

	
	objPagingConn.Open CONN_STRING, CONN_USER, CONN_PASS
	objPagingConn.CursorLocation = 3

	
	Set Grs = Server.CreateObject("ADODB.Recordset")
	'set Gfield = Server.CreateObject("ADODB.Field")
	Grs.CursorLocation = adUseClient
	
	' order by crecente ASC crecente DESC decrecente
	'Grs.sort = "id ASC"
	
	Grs.PageSize = iPageSize
	
	Grs.CacheSize = iPageSize


	Grs.Open strSQL, objPagingConn, adOpenKeyset, adLockOptimistic


	iPageCount = Grs.PageCount
	
	'response.write iPageCount &" "& iPageCurrent
	
	If cint(iPageCurrent) > cint(iPageCount) Then iPageCurrent = cint(iPageCount)
	If cint(iPageCurrent) < 1 Then iPageCurrent = 1
	
	If iPageCount = 0 Then
		Grid_semr = 1
	else
		Grs.AbsolutePage = cint(iPageCurrent)
		Grid_pgi = cint(iPageCurrent)
		Grid_pgf = cint(iPageCount)
		
	end if
	
	'Calcular Tamanho da Tabela
	Glink = 0
	for Glinkx = LBound(GridName) to UBound(GridName)
		GlinkVar = split(GridName(Glinkx),";")
		Glink = cint(GlinkVar(1))+Glink	
	next
	'------------------------------------------/

GridASP ="<table border=""0"" cellpadding=""3"" cellspacing=""2"" class=""borda_table"" width="&Glink&">"&_
 "<tr>"&_
   " <td><table border=""0"" cellpadding=""0"" cellspacing=""0"" class=""texto"">"&_
      "<tr>"
	  
	 'Limpando variavel da largura da página
	 Glink = empty
	 GlinkVar = empty
	 Glinkx = empty
	 '------------------------------------------------
	 
	 'Listar Colunas
	 for GridX = LBound(GridName) to UBound(GridName)
	   Gridvar = split(GridName(GridX),";")
	   Gridvars = split(GridData(GridX),";")
	   
	   GridASP = GridASP & "<td height=24 width="""&Gridvar(1)&""" align=""center"" "
	   
	   if GridX = UBound(GridName) then
      		GridASP = GridASP & "class=""titulo_td_fim"""
				if left(Gridvars(0),1)<>"*" then
			 		GridASP = GridASP & "onmouseover=""this.className='titulo_td_fim_o'"" onmouseout=""this.className='titulo_td_fim'"" onclick=""this.className='titulo_td_fim_c;orders('"&Gridvar(0)&"');'"""
				end if
		else 
			GridASP = GridASP & "class=""titulo_td"" "
				if left(Gridvars(0),1)<>"*" then
					GridASP = GridASP & "onmouseover=""this.className='titulo_td_o'"" onmouseout=""this.className='titulo_td'"" onclick=""this.className='titulo_td_c';orders('"&Gridvars(0)&"');"""
				end if
		end if
		
		GridASP = GridASP & ">"&Gridvar(0)&"</td>"
		
	 next
	  
	  GridASP = GridASP & "</tr>"
	
	  if Grid_semr <> 1 then
	  
	  	iRecordsShown = 0
		Do While iRecordsShown < iPageSize And Not Grs.EOF

    		GridASP = GridASP & " <tr onmouseover=""this.className='td_select'"" onmouseout=""this.className='td_none'"" >"
	 	
		'Listar Dados das Colunas ========================================================
				for GridX = LBound(GridData) to UBound(GridData)
					Gridvar = split(GridData(GridX),";")
					GridASP = GridASP & "<td class="""
					
					
					if GridX = UBound(GridData) then
						GridASP = GridASP & "td_space_fim"
					else 
						GridASP = GridASP & "td_space"
					end if
					
					GridASP = GridASP &  """ align="""
					
					select case cstr(Gridvar(1))
					case "0"
						GridASP = GridASP & "left"
					case "1"
						GridASP = GridASP & "center"
					case "2"
						GridASP = GridASP & "right"
					case else
						GridASP = GridASP & "left"
					end select
					
					GridASP = GridASP & """>"	
					
					'imprime valor em link caso diferente de 0
					if Gridvar(2)<>"0" or Gridvar(2)<> null then	
					
						Glink = Gridvar(2)
						Glink = replace(Glink,"[aspid]", request.form("aspid"))
						for Glinkx = LBound(GridData) to UBound(GridData)
							
								GlinkVar = split(GridData(Glinkx),";")
								if left(GlinkVar(0),1)<>"*" then
									GridMacro = replace(GlinkVar(0),"$","")
									if trim(Grs(GridMacro)) <> "" then	
										
											Glink = replace(Glink,"[col"&Glinkx&"]", trim(Grs(GridMacro))) '//modo antigo
											'// Metodo novo
											for each Gfield in Grs.Fields
												'analiza se foi digitado o nome de outras colunas da tabela
													
													On Error Resume Next
													Glink = replace(Glink,"{"&Gfield.Name&"}",trim(Grs(Gfield.Name)))								
													
											next
											
									else
										Glink = Glink
									end if
									
								end if
							
						next
								
						GridASP = GridASP & "<a href="""&Glink&""" target=""_self"">"
						
					end if
					
					'Imprimie valor na linha
					if cstr(left(Gridvar(0),1)) = "*" then
						GridMacro = replace(Gridvar(0),"*","") 
						'verifica a existencia de outros valores
							GridMacro = replace(GridMacro,"[aspid]", request.form("aspid"))
							for Glinkx = LBound(GridData) to UBound(GridData)
							
								GlinkVar = split(GridData(Glinkx),";")
								if left(GlinkVar(0),1)<>"*" then
									Glink = replace(GlinkVar(0),"$","") 
									if trim(Grs(Glink)) <> "" then
										
										GridMacro = replace(GridMacro,"[col"&Glinkx&"]", trim(Grs(Glink))) '//metodo antigo
										
										'// Metodo novo
										for each Gfield in Grs.Fields
											'analiza se foi digitado o nome de outras colunas da tabela
											On Error Resume Next
											GridMacro = replace(GridMacro,"{"&Gfield.Name&"}",trim(Grs(Gfield.Name)))								
										next
										
									else
										GridMacro = GridMacro
									end if
									
								end if
							
							next
						
						GridASP = GridASP & GridMacro
						
					elseif cstr(left(Gridvar(0),1)) = "$" then
						'existe analize de valor
						GridMacro = replace(Gridvar(0),"$","")
						if Gridvar(3)<>"" then
							if trim(Grs(GridMacro)) <> "" then 
								
								'analiza outros campos
								GridMacro = Gridvar(3)
								for each Gfield in Grs.Fields
									'analiza se foi digitado o nome de outras colunas da tabela
									On Error Resume Next
									GridMacro = replace(GridMacro,"{"&Gfield.Name&"}",trim(Grs(Gfield.Name)))								
								next
								
								GridMacro = AspLogica(GridMacro)
							else
								GridMacro = "<center>-</center>"
							end if
							
						end if
						GridASP = GridASP & GridMacro
					
					else
					
						
						
						if isdate(Grs(Gridvar(0))) then
							GridASP = GridASP & fdate(Grs(Gridvar(0)),GridDate)
						else
							if trim(Grs(Gridvar(0))) <> "" then
								GridASP = GridASP & trim(Grs(Gridvar(0)))
							else
								GridASP = GridASP & "<center>-</center>"
							end if
						end if
						
						
					end if
					
						if Gridvar(2) <> "0" or Gridvar(2)<> null then
							GridASP = GridASP & "</a>"
						end if
						
						GridASP = GridASP & "</td>"
						Gridvar = empty
					
					next
				
				'Fim ==============================================================================
				
			 GridASP = GridASP & "</tr>"
			
	
	  
				iRecordsShown = iRecordsShown + 1
				
				 
			
			Grs.MoveNext
			
		Loop
		
		GFinal = iRecordsShown
		
		if GFinal < rp then
		
			for Glinkx = GFinal to rp-1
				
				 GridASP = GridASP &" <tr onmouseover=""this.className='td_select'"" onmouseout=""this.className='td_none'"" >"
				 	
					for GridX = LBound(GridData) to UBound(GridData)
						 GridASP = GridASP &"<td class="
						 
							if GridX = UBound(GridData) then
								GridASP = GridASP & "td_space_fim"
							else 
								GridASP = GridASP & "td_space"
							end if
							
						GridASP = GridASP & ">&nbsp;</td>"
					next
					
				 GridASP = GridASP &"</tr>"
				
			next 
		
		end if
	
	
	Grs.Close
	Set Grs = Nothing
	objPagingConn.Close
	Set objPagingConn = Nothing
	
	end if 
	

	
	  	if Grid_semr = 1 then
	
    GridASP = GridASP &  "<tr><td colspan="&UBound(GridData)&" align=""center"" valign=""middle"" class=""texto_11""><p>&nbsp;</p>"&_
         " <p>Nenhuma informa&ccedil;&atilde;o encontrada! </p>"&_
          "<p>&nbsp;</p></td>"&_
     " </tr>"
	 
	  	end if
	 
  GridASP = GridASP &   " </table></td>"&_
 " </tr>  <tr>    <td>   " 
 if rp<>0 then
 GridASP = GridASP &   "   <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"" class=""texto"">"&_
       " <tr class=""fundo_base"">"&_
         " <td width=""33%"" align=""center"">"&_
		 
		"<button class=""button"""
		 
		 If cint(iPageCurrent) > 1 Then
		 
		 		GridASP = GridASP &  " onclick=""Gpages("& cint(iPageCurrent) - 1 &");"">"
		 
		 else
		 
			   GridASP = GridASP &   "  disabled=""disabled"">"
		
		 end if
		 	
		 GridASP = GridASP & "&lt;&nbsp;Anterior</button>"
		
 GridASP = GridASP &  "		 </td>          <td width=""33%"" align=""center"" valign=""middle"" class=""texto_11"">"
 			
			'GridASP = GridASP &  "<form>"
			GridASP = GridASP &  "<select name='pagesgr' id='pagesgr' onchange='Gpages(this.value);' class=""texto_11"">"
			for GridX = 1 to grid_pgf
				 	
					
				  	GridASP = GridASP &  "<option value='"&GridX&"'"
					
					 	if cint(iPageCurrent) = cint(GridX) then
							GridASP = GridASP & " selected "
						end if
					
					GridASP = GridASP & ">"&GridX&"</option>"
				 
			next
			GridASP = GridASP &  "</select>"
			' GridASP = GridASP &  "</form>"
			 
           GridASP = GridASP &   "de&nbsp;<b>"&_
           grid_pgf&_
            "</b></td>"&_
          "<td width=""33%"" align=""center"" valign=""middle"">"&_
		 
		  "<button class=""button"""
		  
		  If cint(iPageCurrent) < cint(iPageCount) Then
		  
		 	 
			GridASP = GridASP &  " onclick=""Gpages("& cint(iPageCurrent) + 1 &");"">"
			 
		  
		  else
		  
		  	GridASP = GridASP &  "disabled=""disabled"">"
		  
		  end if
		  
		 	 GridASP = GridASP &  "Pr&oacute;ximo&nbsp;&gt; </button>"
			 
		 GridASP = GridASP &    " </td>        </tr>      </table> "
		end if
		GridASP = GridASP &    " </td>  </tr></table>"
		'GridASP = filtrojava(GridASP)
		GridASP = server.URLEncode(GridASP)
		
	end function
	
	'//----------------------------------------------->
		Dim GridX, GridCount, GridURLS, objBCap
			function GridForm(key,r)
				GridCount = 0
				GridForm = "<form>"&br
				
				for GridX = LBound(GridName) to UBound(GridName)
					
					GridForm = GridForm & "<input name='GridName_"&GridX&"' id='GridName_"&GridX&"' type='hidden' value='"&Encrypt(GridName(GridX),key)&"'>"&br
						
				next
				
				for GridX = LBound(GridData) to UBound(GridData)
					
					GridForm = GridForm & "<input name='GridData_"&GridX&"' id='GridData_"&GridX&"' type='hidden' value='"&Encrypt(GridData(GridX),key)&"'>"&br
					GridCount = GridCount + 1	
				next

				GridForm = GridForm & "<input name='GridDate' id='GridDate' type='hidden' value='"&Encrypt(GridDate,key)&"' />"&br
				GridForm = GridForm & "<input name='GridN' id='GridN' type='hidden' value='"&GridCount&"' />"&br
				GridForm = GridForm & "<input name='GridR' id='GridR' type='hidden' value='"&r&"' />"&br
				GridForm = GridForm & "<input name='GridID' id='GridID'  type='hidden' value='"&Encrypt(GridSQL,key)&" '/>"&br
				
				GridForm = GridForm & "</form>"&br
				
				
				GridForm = GridForm & "<link href='"&site&"/matrix/gui/grid/estilo.css' rel='stylesheet' type='text/css'>"&br
				GridForm = GridForm & "<div id='tablegrid'></div>"&br
								
			GridForm = GridForm & "<script>"&br
			
			GridForm = GridForm & "	function Gpages(n)"&br
			GridForm = GridForm & "	{"&br
		    GridForm = GridForm & "	var http = null;"&br
			GridForm = GridForm & "	  if(window.XMLHttpRequest)"&br
			GridForm = GridForm & "		http = new XMLHttpRequest();"&br

			GridForm = GridForm & "	  else if (window.ActiveXObject)"&br
			GridForm = GridForm & "		http = new ActiveXObject('Microsoft.XMLHTTP');"&br
				
			GridForm = GridForm & "		web = '"&site&"/matrix/gui/grid/grid.asp';"&br
			
			GridForm = GridForm & "	  http.onreadystatechange = function()"&br
			GridForm = GridForm & "	  {"&br
			GridForm = GridForm & "		if(http.readyState == 4)"&br
			GridForm = GridForm & "		// alert(http.responseText);"&br
			GridForm = GridForm & "		  document.getElementById('tablegrid').innerHTML = unescape(http.responseText.replace(/\+/g,' '));"&br		  
			GridForm = GridForm & "	  };"&br
			
			'-- Detecte o navegador
			Set objBCap = Server.CreateObject("MSWC.BrowserType") 
			If objBCap.Browser = "IE" And CInt(objBCap.Version) >= 3 Then
			
				GridForm = GridForm & "	  http.open('POST', web, false);"&br
			
			ElseIf InStr(Request.ServerVariables("HTTP_USER_AGENT"), "MSIE 6") then
			
				GridForm = GridForm & "	  http.open('POST', web, false);"&br
			
			else
			
				GridForm = GridForm & "	  http.open('POST', web, true);"&br
			
			end if
			set objBCap = nothing
			
			GridForm = GridForm & "		  document.getElementById('tablegrid').innerHTML = '<div aling=left>Carregando....</div>';"&br
			GridForm = GridForm & "	  http.setRequestHeader('Content-Type','application/x-www-form-urlencoded');"&br  
			GridForm = GridForm & "	  GridX = document.getElementById('GridN').getAttribute('value');"&br
			GridForm = GridForm & "	  for(i=0;i<GridX;i++)"&br
			GridForm = GridForm & "		{"&br
			GridForm = GridForm & "			if(i<1){"&br
			GridForm = GridForm & "				inc = 'ASPID="&aspid&"&GPAGE='+n+'&GRIDR='+document.getElementById('GridR').getAttribute('value')+'&GRIDN=' + document.getElementById('GridN').getAttribute('value') + '&GridName_' + i + '=' + document.getElementById('GridName_'+i).getAttribute('value') + '&GridData_' + i + '=' + document.getElementById('GridData_'+i).getAttribute('value');"&br
			GridForm = GridForm & "			}else{"&br
			GridForm = GridForm & "				inc = inc + '&GridName_' + i + '=' + document.getElementById('GridName_'+i).getAttribute('value') + '&GridData_' + i + '=' + document.getElementById('GridData_'+i).getAttribute('value');"&br
			GridForm = GridForm & "			}"&br
			GridForm = GridForm & "		}"&br
				
			GridForm = GridForm & "	  inc = inc + '&GridDate='+document.getElementById('GridDate').getAttribute('value')+'&GridID=' + document.getElementById('GridID').getAttribute('value')"&br
				
			GridForm = GridForm & "	  http.send(inc);"&br
			GridForm = GridForm & "	}"&br
			
			GridForm = GridForm & "		Gpages(1);"&br
			
			GridForm = GridForm & "</script>"&br
				
				
				response.write GridForm
			end function

	'//------------------------------------------------->		
 '======================================================================
 'Quebra de linha automatica
 '======================================================================
 dim SplitLens, SplitText, SplitZ
 function splitline(i,t, msg)
	'separa o texto apartir dos espaços
	SplitLen = len(msg)
	SplitY = 1

			for SplitX = 1 to SplitLen
			
				splitline = splitline + mid(msg,SplitX,1)
				if SplitY < i then
					SplitY = SplitY + 1
				else
					if cint(asc(right(splitline,1)))=32 then
						splitline = splitline & t
						SplitY = 1
					else
						SplitLens = len(splitline)
						SplitWord = Split(splitline,chr(32))
						for SplitZ = Lbound(SplitWord) to Ubound(SplitWord)-1
							SplitText = SplitText & SplitWord(SplitZ) & chr(32) 'response.write(Ubound(SplitWord))
						next
						splitline = replace(splitline,SplitText,SplitText & t)
					end if
				end if
				
			next
		
	
	 
 end function
 '=======================================================================
 '   ASPLogica
 '=======================================================================
 	Function VarClear(palavra)
	
	
	
	CAcento = "0123456789qwertyuiopasdfghjklçzxcvbnmàáâãäèéêëìíîïòóôõöùúûüÀÁÂÃÄÈÉÊËÌÍÎÒÓÔÕÖÙÚÛÜçÇñÑŽŠžšŸÅÏås/.-~*+-_^""']}[{`?/;:,\|@#$%¨¨&()"
	SAcento = "                                                                                                                              "
	COR_Texto = ""
	
	
		
	if palavra <> "" then
		For COR_X = 1 to Len(SAcento)
			Letra = mid(palavra,COR_X,1)
			Pos_Acento = inStr(CAcento,Letra)
			if COR_Texto = "" then
				COR_Texto = replace(palavra, mid(CAcento,COR_X,1), mid(SAcento,COR_X,1))
			else
				COR_Texto = replace(COR_Texto, mid(CAcento,COR_X,1), mid(SAcento,COR_X,1))
			end if 
		next
	end if
	
	
		
	VarClear = COR_Texto
		
	
	end function
	

	'valor = request.QueryString("valor")
	't = replace("$case([v]=1):Masculino;$case([v]=2):Duvidoso;$case([v]=3):Feminino;$case([v]=4):Macho","[v]",valor)
	dim Lvar, Lx, Svar, Ly,Sclear, Lz, Stp, Lvalor1, Lvalor2
	function AspLogica(v)
		Lvar = split(v,"|")
		for Lx = Lbound(Lvar) to Ubound(Lvar)
			'analiza se o valor é uma situação
			if left(trim(Lvar(Lx)),5) = "$case" then
				'- Seprando situações
				
				'response.write var(x) & "<br>"
				svar = split(Lvar(Lx),":")
				for Ly = Lbound(svar) to Ubound(svar)
				
					if Ly=0 then
						Sclear = replace(svar(0),"$case(","")
						Sclear = replace(Sclear,")","")
						select case trim(VarClear(svar(0)))
						case "="
							Lz = split(Sclear,"=")
							Stp = 0	
						case "<"
							Lz = split(Sclear,"<")
							Stp = 1
						case ">"
							Lz = split(Sclear,">")
							Stp = 2
						case "!"
							Lz = split(Sclear,"!")
							Stp = 3					
						case else
							'response.write "!erro"
							'response.end()					
						end select
						
						Lvalor1 = Lz(0)
						Lvalor2 = Lz(1)
						
						if isnumeric(Lvalor1) then
							Lvalor1 = cint(Lvalor1)
						end if
						
						if isnumeric(Lvalor2) then
							Lvalor2 = cint(Lvalor2)
						end if
						
						'//Valor vazio
						if Lvalor1 = "null" then
							Lvalor1 = " "
						end if
						
						if Lvalor2 = "null" then
							Lvalor2 = " "
						end if
						
						Select case Stp
						case 0
							if Lvalor1 = Lvalor2 then
								AspLogica = svar(1) & AspLogica
							elseif Ubound(svar)=2 then
								AspLogica = svar(2) & AspLogica
							end if
						case 1
							if Lvalor1 < Lvalor2 then
								AspLogica = svar(1) & AspLogica
							elseif Ubound(svar)=2 then
								AspLogica = svar(2) & AspLogica
							end if
						case 2
							if Lvalor1 > Lvalor2 then
								AspLogica = svar(1) & AspLogica
							elseif Ubound(svar)=2 then
								AspLogica = svar(2) & AspLogica
							end if
						case 3
							if Lvalor1 <> Lvalor2 then
								AspLogica = svar(1) & AspLogica
							elseif Ubound(svar)=2 then
								AspLogica = svar(2) & AspLogica
							end if
						end select
						
					end if
						
				next
				
			else
				AspLogica = AspLogica & Lvar(Lx)
			end if
			
		next	
	end function
 '=======================================================================
 'Sys Mural =============================================================
 'Traduz Conteudo do mural
 
 dim  tagIco, tagImg
 'Filtrar tags como <script> <img> <font> <a>
	
	tagIco  = "::-#,8o|,8-|,^o),:-*,+o(,(sn),(pl),(||),(pi),(so),(au),(ap),(um),(ip),(co),(ce),(ch),(di),:^),*-),(li),:o),8-),|-),:@,:[,(b),(u),(^),(p),(@),(o),(c),:s,*(,(6),(&),(e),(~),(l),(k),(i),(d),(S),(du),(8),:o,(t),(g),(f),(h),(*),:p,:|,(pis)"
	tagImg = "47_47,48_48,49_49,50_50,51_51,52_52,53_53,55_55,56_56,57_57,58_58,59_59,60_60,61_61,62_62,63_63,64_64,66_66,69_69,71_71,72_72,73_73,74_74,75_75,77_77,angry_smile,bat,beer_mug,broken_heart,cake,camera,cat,clock,coffee,confused_smile,cry_smile,devil_smile,dog,envelope,film,heart,kiss,lightbulb,martini,moon,newToIM,note,omg_smile,phone,present,rose,shades_smile,star,tongue_smile,what_smile,wink_smile"
 
 function sysMural(n)
 	dim nMural,icos, imgs, ix, imgUrls, tags,tagDead
	
	nMural = replaceBadWords(n, BAD_WORDS_FILTER)
	nMural = cHTML(nMural)
	tagDead = "<script, <img, <iframe, <a, <h, <f, <l, <d, <br, <p, <sound, <in"
	imgUrls = site&"/img/icones_mural/"
	
 	
	
	
	icos = Split(tagIco, ",")
	imgs = Split(tagImg, ",")
	tags = Split(tagDead, ",")
	
	'Remove Tags Indevidas
	For ix=0 to Ubound(tags)
		nMural = replace(nMural,tags(ix),"")
	next
		
	'Subistitui os codigos pelos icones	
	For ix=0 to Ubound(icos)
		nMural = replace(nMural,icos(ix),"<img src='"&imgUrls&imgs(ix)&".gif'>")
	next
	
	sysMural = nMural
	
 end function
 
 '=======================================================================
 function cHTML(n)
 
 	dim acents, cx, acentslen, acCommand
	
	cHTML = n
	acents = "á,Á,à,À,ã,Ã,â,Â,ä,Ä,ç,Ç,é,É,è,È,ê,Ê,ë,Ë,í,Í,ì,Ì,î,Î,ï,Ï,ó,Ó,ò,Ò,ô,Ô,ö,Ö,õ,Õ,ú,Ú,ù,Ù,û,Û,ü,Ü,',ª,º,´,`,¨,ñ,Ñ,$data,$hora,$email,$usuario,$emailusuario,$ip"
	acentslen = split(acents,",")
	
	'Remove Tags Indevidas
	For cx=0 to Ubound(acentslen)
		if left(acentslen(cx),1) <> "$" then
		
			On Error Resume Next
				cHTML = replace(cHTML,acentslen(cx),"&#"&asc(acentslen(cx))&";")
			If Err.Number<>0 then
				cHTML = cHTML
			end if
			
		else
		
			acCommand = replace(acentslen(cx),"$","")
			select case acCommand
				case "data"
					cHTML = replace(cHTML,acentslen(cx),fdate(date(),"dd/mm/yyyy"))
				case "hora"
					cHTML = replace(cHTML,acentslen(cx),fdate(now(),"hh:nn"))
				case "ip"
					cHTML = replace(cHTML,acentslen(cx),request.ServerVariables("REMOTE_ADDR"))
				case else
					cHTML = cHTML
			end select
			 
		end if
	next
	
 end function
 '================================================================================
 ' AspTempo
 '================================================================================
 
 dim txtTmp, tTmp, tlTmp, cabTmp, XMLHttp
 dim tempoEstado, imgTempo, timerTempo
 
	function AspTempo(selCli)
		'//verifica se a pasta existe
		md tempo_root
		
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		
		If fso.FileExists(server.MapPath(tempo_root&fDate(DateAdd("D", 1, Date),"ddmmyyyy")&".js")) = false then
		
			Set XMLHttp = Server.CreateObject("Microsoft.XMLHTTP")
			if tempoEstado <> "" then
				XMLHttp.open "GET", "http://www.tempoagora.com.br/climaflash/"&tempoEstado&".php", false
			else
				XMLHttp.open "GET", "http://www.tempoagora.com.br/climaflash/se.php", false
			end if
			'XMLHttp.open "GET", "http://localhost/se.txt", false

			On Error Resume Next
				XMLHttp.send()
			If Err.Number<>0 then
				response.write "<script>document.getElementById('imgtmp').innerHTML='erro';</script>"
			else
			
				tTmp = trim(XMLHttp.ResponseText)
				
		
				tTmp = replace(tTmp,";","','")
				tTmp = replace(tTmp,chr(10),"")
				
				
				if tempoEstado = "br" then
					tTmp = replace(tTmp,"&atualizado=","var atualizado = ")
					tTmp = replace(tTmp,"&total=",";"&vbcrlf&"var Total = new Array('")
					tTmp = replace(tTmp,"&estado=","');"&vbcrlf&"var Estado = new Array('")
				else
					tTmp = replace(tTmp,"&total=","var Total = new Array('")
				end if
				
				tTmp = replace(tTmp,"','&cidade=","');"&vbcrlf&"var Cidade = new Array('")
				tTmp = replace(tTmp,"','&clima1=","');"&vbcrlf&"var Clima1 = new Array('")
				tTmp = replace(tTmp,"','&clima2=","');"&vbcrlf&"var Clima2 = new Array('")
				tTmp = replace(tTmp,"','&clima3=","');"&vbcrlf&"var Clima3 = new Array('")
				tTmp = replace(tTmp,"','&tpmin1=","');"&vbcrlf&"var Tpmin1 = new Array('")
				tTmp = replace(tTmp,"','&tpmin2=","');"&vbcrlf&"var Tpmin2 = new Array('")
				tTmp = replace(tTmp,"','&tpmin3=","');"&vbcrlf&"var Tpmin3 = new Array('")
				tTmp = replace(tTmp,"','&tpmax1=","');"&vbcrlf&"var Tpmax1 = new Array('")
				tTmp = replace(tTmp,"','&tpmax2=","');"&vbcrlf&"var Tpmax2 = new Array('")
				tTmp = replace(tTmp,"','&tpmax3=","');"&vbcrlf&"var Tpmax3 = new Array('")
				tTmp = replace(tTmp,"','&tpatual=","');"&vbcrlf&"var Tpatual= new Array('")
				tTmp = replace(tTmp,"','&iconatual=","');"&vbcrlf&"var Iconatual= new Array('")
				tTmp = replace(tTmp,"','&link=","');"&vbcrlf&"var Link= new Array('")
				tlTmp= len(tTmp) - 2
				tTmp = left(tTmp,tlTmp)
				tTmp = tTmp + ");"
				' determine a array que deseja obter a temperatura
				
				if selCli <> 0 then
					tTmp = tTmp + vbcrlf&"var x="&selCli&";"&vbcrlf
				else
				
					tTmp = tTmp + vbcrlf&"var x=0;"&vbcrlf
					
					if timerTempo <> "" then
						tTmp = tTmp + vbcrlf&"var timers = "&timerTempo&"; //tempo em segundos"&vbcrlf
					else
						tTmp = tTmp + vbcrlf&"var timers = 7; //tempo em segundos"&vbcrlf
					end if
					
					tTmp = tTmp + vbcrlf&"timers = timers * 1000;"&vbcrlf
					
					if tempoEstado = "br" then
						tTmp = tTmp + vbcrlf&"var y = Cidade.length - 2;"&vbcrlf
					else
						tTmp = tTmp + vbcrlf&"var y = Cidade.length - 1;"&vbcrlf
					end if
					
					tTmp = tTmp + vbcrlf&"function cid(){"&vbcrlf
					tTmp = tTmp +	vbcrlf&"if (x==y){"&vbcrlf
					tTmp = tTmp +	vbcrlf&"x=0;"&vbcrlf
					tTmp = tTmp +	vbcrlf&"}else{"&vbcrlf
					tTmp = tTmp +	vbcrlf&"x++;"&vbcrlf
					tTmp = tTmp + vbcrlf&"}"&vbcrlf
					
				end if
				
				'=================================================================
				tTmp = tTmp + "	try{document.getElementById('cidade').innerHTML = Cidade[x]}catch(e){}"&vbcrlf
				tTmp = tTmp + "	try{document.getElementById('temMax').innerHTML = Tpmax2[x]+'&ordm;'}catch(e){};"&vbcrlf
				tTmp = tTmp + "	try{document.getElementById('temMin').innerHTML = Tpmin2[x]+'&ordm;'}catch(e){};"&vbcrlf
				
				if tempoEstado = "br" then
					tTmp = tTmp + "	try{document.getElementById('estado').innerHTML = Estado[x].toUpperCase()}catch(e){};"&vbcrlf
				end if
				
				tTmp = tTmp + "		var g = new String(Clima2[x])"&vbcrlf
				tTmp = tTmp + "		if(g=='cc'){"&vbcrlf
				tTmp = tTmp + "			tTmp = 'Céu Claro';"&vbcrlf
				tTmp = tTmp + "			it = 'ceuclaro.gif';"&vbcrlf
				tTmp = tTmp + "		}"&vbcrlf
				tTmp = tTmp + "		else"&vbcrlf
				tTmp = tTmp + "		{"&vbcrlf
				tTmp = tTmp + "			if(g=='ch'){"&vbcrlf
				tTmp = tTmp + "				t = 'Chuvendo'"&vbcrlf
				tTmp = tTmp + "				it = 'chvendo.gif';"&vbcrlf
				tTmp = tTmp + "			}"&vbcrlf
				tTmp = tTmp + "			else"&vbcrlf
				tTmp = tTmp + "			{"&vbcrlf
				tTmp = tTmp + "				if(g=='cv'){"&vbcrlf
				tTmp = tTmp + "					t = 'Chuviscos'"&vbcrlf
				tTmp = tTmp + "					it = 'chuvisco.gif';"&vbcrlf
				tTmp = tTmp + "				}"&vbcrlf
				tTmp = tTmp + "				else"&vbcrlf
				tTmp = tTmp + "				{"&vbcrlf
				tTmp = tTmp + "					if(g=='en'){"&vbcrlf
				tTmp = tTmp + "						t = 'Encoberto'"&vbcrlf
				tTmp = tTmp + "						it = 'encoberto.gif';"&vbcrlf
				tTmp = tTmp + "					}"&vbcrlf
				tTmp = tTmp + "					else"&vbcrlf
				tTmp = tTmp + "					{"&vbcrlf
				tTmp = tTmp + "						if(g=='ge'){"&vbcrlf
				tTmp = tTmp + "							t = 'Geadas'"&vbcrlf
				tTmp = tTmp + "							it = 'geadas.gif';"&vbcrlf
				tTmp = tTmp + "						}"&vbcrlf
				tTmp = tTmp + "						else"&vbcrlf
				tTmp = tTmp + "						{"&vbcrlf
				tTmp = tTmp + "							if(g=='nb'){"&vbcrlf
				tTmp = tTmp + "								t = 'Nublado'"&vbcrlf
				tTmp = tTmp + "								it = 'nublado.gif'"&vbcrlf
				tTmp = tTmp + "							}"&vbcrlf
				tTmp = tTmp + "							else"&vbcrlf
				tTmp = tTmp + "							{"&vbcrlf
				tTmp = tTmp + "								if(g=='ne'){"&vbcrlf
				tTmp = tTmp + "									t = 'Neve'"&vbcrlf
				tTmp = tTmp + "									it = 'neve.gif'"&vbcrlf
				tTmp = tTmp + "								}"&vbcrlf
				tTmp = tTmp + "								else"&vbcrlf
				tTmp = tTmp + "								{"&vbcrlf
				tTmp = tTmp + "									if(g=='pc'){"&vbcrlf
				tTmp = tTmp + "										t = 'Panc. de Chuva'"&vbcrlf
				tTmp = tTmp + "										it = 'pancadas.gif'"&vbcrlf
				tTmp = tTmp + "									}"&vbcrlf
				tTmp = tTmp + "									else"&vbcrlf
				tTmp = tTmp + "									{"&vbcrlf
				tTmp = tTmp + "										if(g=='pi'){"&vbcrlf
				tTmp = tTmp + "											t= 'Chuvas Rápidas'"&vbcrlf
				tTmp = tTmp + "											it = 'chavarapida.gif'"&vbcrlf
				tTmp = tTmp + "										}"&vbcrlf
				tTmp = tTmp + "										else"&vbcrlf
				tTmp = tTmp + "										{"&vbcrlf
				tTmp = tTmp + "											if(g=='pn'){"&vbcrlf
				tTmp = tTmp + "												t = 'Poucas Nuvens'"&vbcrlf
				tTmp = tTmp + "												it = 'poucasnuv.gif'"&vbcrlf
				tTmp = tTmp + "											}"&vbcrlf
				tTmp = tTmp + "											else"&vbcrlf
				tTmp = tTmp + "											{"&vbcrlf
				tTmp = tTmp + "												if(g=='nc'){"&vbcrlf
				tTmp = tTmp + "													t = 'Nub. c/ Chuva'"&vbcrlf
				tTmp = tTmp + "													it = 'nubchuva.gif'"&vbcrlf
				tTmp = tTmp + "												}"&vbcrlf
				tTmp = tTmp + "												else"&vbcrlf
				tTmp = tTmp + "												{"&vbcrlf
				tTmp = tTmp + "													t = '&nbsp;'"&vbcrlf
				tTmp = tTmp + "													it = '&nbsp;'"&vbcrlf
				tTmp = tTmp + "												}"&vbcrlf
				tTmp = tTmp + "											}"&vbcrlf
				tTmp = tTmp + "										}"&vbcrlf
				tTmp = tTmp + "									}"&vbcrlf
				tTmp = tTmp + "								}"&vbcrlf
				tTmp = tTmp + "							}"&vbcrlf
				tTmp = tTmp + "						}"&vbcrlf
				tTmp = tTmp + "					}"&vbcrlf
				tTmp = tTmp + "				}"&vbcrlf
				tTmp = tTmp + "			}"&vbcrlf
				tTmp = tTmp + "		}"&vbcrlf
				
				if imgTempo = "" then
					imgTempo = site&"/matrix/img/tempo/"
				end if
				
				tTmp = tTmp + "		try{document.getElementById('imgtmp').innerHTML=""<img src='"&imgTempo&"""+it+""'>""}catch(e){};"&vbcrlf
				tTmp = tTmp + "		try{document.getElementById('comenttmp').innerHTML=t}catch(e){};"&vbcrlf
				
				'cab = cab + "/* **********************************"&vbcrlf
				'cab = cab + "ASPTempo tecnologia Microlabs"&vbcrlf
				'cab = cab + "************************************ */"&vbcrlf
				'cab = cab + "tUrl = location.protocol + '//' + location.hostname;"&vbcrlf
				'cab = cab + "if(tUrl!='"&sites&"'){"&vbcrlf
				'cab = cab + "alert('Microlabs - Sistema Tempo Agora\nO uso do sistema é de exclusividade dos sites desenvolvidos pela Microlabs.\n\nAo clicar em OK você será redirecionado um site da Microlabs. Obrigado!');"&vbcrlf
				'cab = cab + "window.open('"&sites&"','_self');"&vbcrlf
				'cab = cab + "}"&vbcrlf
				
				if selCli = 0 then
					tTmp = tTmp + vbcrlf&"setTimeout('cid()', timers);"&vbcrlf
					tTmp = tTmp + vbcrlf&"}"&vbcrlf
					tTmp = tTmp + vbcrlf&"cid();"
				end if
				
			
			
				criatxt tempo_root&fDate(DateAdd("d", 1, Date),"ddmmyyyy")&".js", 2, cabTmp&vbcrlf&tTmp
				delfile tempo_root&fDate(DateAdd("d", -1, Date),"ddmmyyyy")&".js"
				
				set fso = nothing
				'response.write t
			end if
			
		end if
		response.write  "<script src="""&sites&tempo_root&fDate(DateAdd("D", 0, Date),"ddmmyyyy")&".js""></script>"
	end function
'/* **************************************************
' *    Name: TextField Clock
' *    Created by: Tucows (http://www.tucows.com)
' *    Date: 7/30/1999
' *    Description:
' *          This is a clock that will update every
' *          second.  The time will be displayed
' *          within the text field that you must add
' *          into your page.  It is capable of
' *          displaying either 12 or 24 hour time.
' * --- Adaptado por Paulo Corcino 
' * --- www.microlabsse.com.br
' * **************************************************
' Variaveis para implatar o tempo
' 	<div id="temMax"></div> - temperatura maxima
'	<div id="temMin"></div> - temperatura minima
'	<div id="horas"></div> - data e hora
'	<div id="imgtmp"></div> - imagem
'*/
dim zeroLen, zeroConut
function incZero(n,i)
	zeroLen = len(n)
	n = cint(n)
	if n<10 then
		for zeroConut = 1 to i
			n = "0"&n
		next
	end if	
	incZero = n	
end function

dim UrlLen, UrlCod, UrlReal
function sysUrl(n,s)
	
	UrlCod  = split("[album],[info],[contato],[bandas],[mural]",",")
	UrlReal = split("inc/inc_album.asp?id=,inc/inc_noticia.asp?id=,inc/inc_bandas.asp?,inc/inc_mural.asp?",",")
	sysUrl  = n
	
	if s=0 then
		
		for i=lBound(UrlCod) to UBound(UrlCod)
		
			On Error Resume Next
				sysUrl = replace(sysUrl,UrlCod(i),UrlReal(i))
			
		
		next
		
	else
		for i=lBound(UrlCod) to UBound(UrlCod)
		
			On Error Resume Next
				sysUrl = replace(sysUrl,UrlReal(i),UrlCod(i))
		
		next
	end if


end function

'//------------------------------------------------------------------------------
'Menu Intranet
'//------------------------------------------------------------------------------
	
	
	function AspMenu()
	
		dim AspMenuImg(11)
		dim AspMenuItens
		dim AspMenuLarg
		dim AspMenuParan
		
		AspMenuImg(0) = "menu.gif"   'menu
		AspMenuImg(1) = "voltar.gif" 'voltar
		AspMenuImg(2) = "print.gif"  'imprimir
		AspMenuImg(3) = "ok.gif"     'confirmar
		AspMenuImg(4) = "incluir.gif"'Incluir
		AspMenuImg(5) = "exluir.gif" 'apagar
		AspMenuImg(6) = "edit.gif"   'editar
		AspMenuImg(7) = "report.gif" 'reportar
		AspMenuImg(8) = "most.gif"   'mostrar tudo
		AspMenuImg(9) = "hide.gif"   'ocultar tudo
		AspMenuImg(10) = "item.gif"  'ocultar tudo
		
		AspMenuItens = UBound(AspMenuOp)
		AspMenuLarg = 45 * AspMenuItens 
		
		response.write "<table width="""& AspMenuLarg &""" border=""0"" cellspacing=""2"" cellpadding=""0"">"&_
      				   "<tr>"
					
		for i = 0 to UBound(AspMenuOp) -  1
		
			AspMenuParan = split(AspMenuOp(i),"|",-1,1)
   			response.write "<td width=""45""><a  href=""javascript:" &   AspMenuParan(2) & ";""><div id=""menustyle""><img src="""& site &"/matrix/img/"& AspMenuImg(AspMenuParan(1)) &""" name=""imgmenu"" width=""16"" height=""16"" border=""0"" id=""imgmenu"" /><br />" & AspMenuParan(0) & "</div></a></td>"

		next
		
			response.write "</tr>" &_
		     		       "</table>"
		
				
	end function
	'//----------------------------------------------------------------------
	function subTitle(n)
		response.write "<script>subTitle('"&filtrojava(cHtml(n))&"')</script>"
	end function

'//-------------------------------------------------------------------------------
'//Parametros
	dim NoticiaImgComent, NoticiaDeSize, NoticiaEmailSize, NoticiaMsgSize
	NoticiaImgComent = ""
function NoticiaComent(cod)

	startjs() '//Inicia o script js

	response.write "<form name='nformComent'>"
	response.write "<div id='nComent'></div>"&_
	"</form>" &_
	"<script>" &_
	"AjaxLoad("""&site&"/matrix/noticia/inc_not_coment.asp?id="&cod&"&html="&request("html")&"&img="&NoticiaImgComent&""",null,""nComent"")" &_
	"</script>"
end function

	dim ncountrs, ncountsql, ncountern
function NoticiaCount(cod)
	db "leitura",ncountrs,ncountsql,"select counter from not_noticia where id="&cod,conn
	if ncountrs("counter")<>"" then
		if isnumeric(ncountrs("counter")) then
			ncountern = cint(ncountrs("counter"))
		else
			ncountern = 0
		end if
	else
		ncountern = 0
	end if
	
	ncountern = ncountern + 1
	
	db "normal",,ncountsql,"update not_noticia set counter="&ncountern&" where id="&cod,conn
	
end function

dim imgEnqV, imgEnqR

function eEnquete(cod)

startjs() '//Inicia o script js
response.write "<script language=""JavaScript"" type=""text/javascript"" src="""&site&"/matrix/enquete/nEnquete.js""></script>"
response.write "<form name='enqForm'>" &_
				"<div id='eEnquete'></div>" &_
				"</form>"

	dim isresult, enqtrs, enqtsql
	isresult = 0
	
	db "leitura",enqtrs,enqtsql,"select top 1 * from enqueteshow where id_a="&cod,conn
	
	if request.Cookies("eEnquete"&enqtrs("id")) = "votado" then
		isresult = 1
	end if
	
	response.write "<script>"
	if isresult = 0 then
		response.write "AjaxLoad("""&site&"/matrix/enquete/inc_enquete.asp?op=&ida="&cod&"&html="&request("html")&"&imgv="&imgEnqV&"&imgr="&imgEnqR&""",null,""eEnquete"")"
	else
		response.write "AjaxLoad("""&site&"/matrix/enquete/inc_enquete.asp?op=result&ida="&cod&"&html="&request("html")&"&imgv="&imgEnqV&"&imgr="&imgEnqR&""",null,""eEnquete"")"
	end if
	
	response.write "</script>"

end function

dim aAlbumText, imgAlbumVoltar, imgAlbumAvanca, imgAlbumVoltarIn, imgAlbumAvancaIn
dim imgAlbumEnviar, imgAlbumNumberItens, imgAlbumNumberColum, AlbumId
dim AlbumW, AlbumH, aSql, aRs, aAlbumCod_foto, imgAlbumFotoDestaq, imgAlbumMiniComent, imgAlbumComentOrdem, imgAlbumComentStart

function aAlbum()

	startjs() '//Inicia o script js
	aAlbumText = "<div id=""aAlbumSendFoto"" style=""display:none"">&nbsp;</div>" &_
				 "<script language=""JavaScript"" type=""text/javascript"" src="""&site&"/matrix/album_book/aAlbum.js""></script>" &_
				 "<script>" &_
					"var AlbumId;"&_
					"var AlbumW;"&_
					"var AlbumH;"&_
					"var imgAlbumVoltar;"&_
					"var imgAlbumVoltarIn;"&_
					"var imgAlbumAvanca;"&_
					"var imgAlbumAvancaIn;"&_
					"var imgAlbumEnviar;"&_
					"var imgAlbumNumberItens;"&_
					"var imgAlbumNumberColum;"&_
					"var AlbumLinkDefault;"&_
					"var AlbumLinkDefaultC;"&_
					"var AlbumUrl;" &_
					"var AlbumSendEmailText;"&_
					"AlbumUrl = '"&filtrojava(site)&"';"

		
		if imgAlbumVoltar = "" then
			aAlbumText = aAlbumText & "imgAlbumVoltar   = ""Voltar"";"
		else 
			aAlbumText = aAlbumText & "imgAlbumVoltar   = """&filtrojava(imgAlbumVoltar)&""";"
		end if
		
		if imgAlbumAvanca = "" then
			aAlbumText = aAlbumText & "imgAlbumAvanca   = ""Avançar"";"
		else
			aAlbumText = aAlbumText & "imgAlbumAvanca   = """&filtrojava(imgAlbumAvanca)&""";"
		end if
		
		if imgAlbumVoltarIn = "" then
			aAlbumText = aAlbumText & "imgAlbumVoltarIn = ""Voltar"";"
		else
			aAlbumText = aAlbumText & "imgAlbumVoltarIn = """&filtrojava(imgAlbumVoltarIn)&""";"
		end if 
		
		if imgAlbumAvancaIn = "" then
			aAlbumText = aAlbumText & "imgAlbumAvancaIn = ""Avançar"";"
		else
			aAlbumText = aAlbumText & "imgAlbumAvancaIn = """&filtrojava(imgAlbumAvancaIn)&""";"
		end if
		
		if imgAlbumEnviar =  "" then
			aAlbumText = aAlbumText & "imgAlbumEnviar   = ""Enviar Foto"";"
		else
			aAlbumText = aAlbumText & "imgAlbumEnviar   = """&filtrojava(imgAlbumEnviar)&""";"
		end if
		
		if imgAlbumNumberItens = "" then
			aAlbumText = aAlbumText & "imgAlbumNumberItens = 5;"
		else
			aAlbumText = aAlbumText & "imgAlbumNumberItens = "&filtrojava(imgAlbumNumberItens)&";"
		end if
		
		if imgAlbumNumberColum = "" then
			aAlbumText = aAlbumText & "imgAlbumNumberColum = 5;"
		else
			aAlbumText = aAlbumText & "imgAlbumNumberColum = "&filtrojava(imgAlbumNumberColum)&";"
		end if 
		
		if AlbumId = "" then
			alert("Este album não pode ser aberto pois, não foi determidado o seu ID")
			response.end()
		else
			if imgAlbumFotoDestaq <> "" and imgAlbumFotoDestaq<>0  then
				'//Detecta o ID da foto destaque
				db "leitura",aRs,aSql,"select id from album_fotos where dest=1 and id_a = " & AlbumId,conn
				aAlbumCod_foto = aRs("id")
			end if
			
			aAlbumText = aAlbumText & "AlbumId = "&filtrojava(AlbumId)&";"
		end if 
		
		if AlbumW = "" then
			aAlbumText = aAlbumText & "AlbumW  = 86;"
		else
			aAlbumText = aAlbumText & "AlbumW  = "&filtrojava(AlbumW)&";"
		end if 
		
		if AlbumH = "" then
			aAlbumText = aAlbumText & "AlbumH  = 60;"
		else
			aAlbumText = aAlbumText & "AlbumH  = "&filtrojava(AlbumH)&";"
		end if 
		
	
	if imgAlbumMiniComent <> "" and imgAlbumMiniComent <> 0 then
		aAlbumText = aAlbumText & "AlbumLinkDefaultC = """&site&"/matrix/album_book/inc_coment_album.asp?itencol=""+imgAlbumNumberColum+""&itenpg=""+imgAlbumNumberItens+""&w=""+AlbumW+""&h=""+AlbumH+""&id=""+AlbumId+""&gpage=&itens="&server.URLEncode(imgAlbumComentOrdem)&""";"
		aAlbumText = aAlbumText & "AlbumLinkDefault = """&site&"/matrix/album_book/inc_thumb_album.asp?itencol=""+imgAlbumNumberColum+""&itenpg=""+imgAlbumNumberItens+""&w=""+AlbumW+""&h=""+AlbumH+""&id=""+AlbumId+""&gpage="";"
		
	else
		aAlbumText = aAlbumText & "AlbumLinkDefault = """&site&"/matrix/album_book/inc_thumb_album.asp?itencol=""+imgAlbumNumberColum+""&itenpg=""+imgAlbumNumberItens+""&w=""+AlbumW+""&h=""+AlbumH+""&id=""+AlbumId+""&gpage="";"
	end if
	
	
	aAlbumText = aAlbumText & "if(ShowAlbum){"	

		if imgAlbumMiniComent <> "" and imgAlbumMiniComent <> 0 then
			 aAlbumText = aAlbumText & "AjaxLoadAlbum(AlbumLinkDefaultC,'imgstart="&filtrojava(imgAlbumComentStart)&"','aAlbumThumb',1);"
		else
			aAlbumText = aAlbumText & "AjaxLoadAlbum(AlbumLinkDefault,null,'aAlbumThumb');"
		end if
			
			if aAlbumCod_foto <> "" then
				 aAlbumText = aAlbumText & "document.getElementById('aAlbumShowFoto').innerHTML = '<img id=\'aAlbumShowFotoImg\' src=\'"&site&"/matrix/album_book/inc_album_foto.asp?id="&aAlbumCod_foto&"\'>';"
			end if
			
	aAlbumText = aAlbumText & "}" &_	
				 "</script>"
	
	response.write aAlbumText

end function

dim Jststart

	Jststart = false
function startjs()
	
	if Jststart = false then
		response.write "<script language=""JavaScript"" type=""text/javascript"" src="""&site&"/matrix/js/script.js""></script>"
		Jststart = true
	end if
	
end function

dim OptNews, NewsGrupo
function AddNews()

	startjs()
	
	if NewsGrupo = "" then
		NewsGrupo = "NewsSite"
	end if
	
	
	response.write "<form name=""NewsAdm""><div id=""NewsAddRem""></div></form>" &_
	"<script>" &_
	"var WebSiteNews = '"&site&"';" &_ 
	"AjaxLoad('"&site&"/matrix/mala/inc_mala_add.asp?op="&OptNews&"&nome_grupo="&NewsGrupo&"',null,'NewsAddRem');" &_
	"</script>"&_
	"<script language=""JavaScript"" type=""text/javascript"" src="""&site&"/matrix/mala/nNews.js""></script>"

end function

dim asrs, assql
function AddSessionsNoticia(obj)
	
	response.write "<script>function addOption(obj,text,value,selected) { "&_
			   	   "if (obj!=null && obj.options!=null) { "&_
      			   "obj.options[obj.options.length] = new Option(text, value, false, selected); "&_
	      		   "}};" &_
				   "try{ document."&obj&".value = 1;"
				   response.write "addOption("&obj&",'Todas as áreas','0',true);" 
				   db "leitura",asrs,assql,"select * from NoticiaAreaNoBlank",conn
				   while not asrs.eof 
				   		
						response.write "addOption("&obj&",'"&trim(asrs("area"))&"','"&asrs("id")&"',false);" 
						
				   		asrs.movenext
				   wend				   
				   
	response.write "} catch(e){" &_
					"alert('Este campo não existe, verifique se o nome está correto.');"&_
					"}</script>"
				   
end function


%>

<!--#include virtual="/bymidia/matrix/config/adovbs.asp" -->

