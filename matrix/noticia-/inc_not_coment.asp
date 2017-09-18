<!--#include virtual ="/bymidia/matrix/config/config.asp" -->
<%
	cache 1, true
	dim rs,sql, img, objsText, ObjsTextComent

		
	img = request("img")
	
	if img = "" then
		img = "/matrix/img/edit.gif"
	end if



		objsText     =     "<p>" &_
					       "<div id='nComentFundo'>"

		db "leitura",rs,sql,"select * from NoticiaComent where id_n="&request("id"),conn
		while not rs.eof

			objsText  =     objsText & "<div id='nComentFundoItem' style='margin-bottom:5px'>" &_
							"<div id='nComentItemTopo'>" &_
							"<font id='nComentItemTituloDe'>De:</font> <font id='nComentItemTextoDe'>"&chtml(rs("de"))&"</font> <br>"
		
			if rs("email") <> "" then
			
				objsText  =      objsText & "<font id='nComentItemTituloEmail'>E-Mail:</font> <font id='nComentItemTextoEmail'>"&chtml(rs("email"))&"</font><br>"

			end if
			
			objsText = objsText & "<font id='nComentItemTituloData'>Data:</font> <font id='nComentItemTextoData'>"&fdate(rs("data"),"dd/mm/yyyy")&" - "&fdate(rs("data"),"hh:nn")&"</font>" &_
					   "</div><div id='nComentItemMsg'>"&_
					   chtml(rs("msg"))&"</div></div>"

		rs.movenext
	wend

		objsText = objsText &"</div></p>"




	if request("post")<>"" then
		db "leitura",rs,sql,"select id from not_coment where criado=0",conn
		db "normal",,sql,"update not_coment set id_n="&formus("nid_n",0)&",de='"&formus("nDeComent",1)&"',email='"&formus("nEmailComent",1)&"',ip='"&request.ServerVariables("REMOTE_ADDR")&"',msg='"&formus("nMsgComent",1)&"',criado=1,atv=0 where id="&rs("id"),conn

			ObjsTextComent = "<div id='nComentMsgEnviada'>Coment&aacute;rio postado com sucesso!<img src='"&site&"/matrix/malaaction/addmail.asp?grupo=E-Mail Comentario&email="&formus("nEmailComent",1)&"'></div>"
	else 
			ObjsTextComent = "<div id='nComentMsgBoaVindas'>"&_
							  "Deixe Aqui seu coment&aacute;rio"&_
								"</div>"&_
								"<div id='nDeTexto'>De</div> "&_
								"<input name='nDeComent' type='text' id='nDeComent' maxlength='30'>"&_
								"<br>" &_
								"<div id='nEmailTexto'>E-Mail</div>" &_
								"<input name='nEmailComent' type='text' id='nEmailComent' maxlength='255' onBlur=""if(!isEmail(this.value)){this.value=''}"">" &_
								"<br>"&_
								"<div id='nMsgTexto'>Mensagem</div><input name='nid_n' type='hidden' value='"&request("id")&"' />" &_
								"<textarea name='nMsgComent' id='nMsgComent' onKeyUp='CharMax(this,255)'></textarea><input name='post' type='hidden' value='1' />" &_
								"<br />"&_
								"<br />"&_
								"<a href='javascript:void(0)' onclick=""AjaxLoad('"&site&"/matrix/noticia/inc_not_coment.asp?id="&request("id")&"&img="&img&"',allCampQuery('nformComent'),'nComent')""><img src='"&site&img&"' name='nSalvarComent' border='0' id='nSalvarComent' /></a>"

	end if
	
	if request("html")<>"" then
		
		ObjsTextComent = replace(ObjsTextComent,"<input","&#60input")
		ObjsTextComent = replace(ObjsTextComent,"<img","&#60img")
		ObjsTextComent = replace(ObjsTextComent,">","&#62")
		response.write 	"<textarea cols=60 rows=15>"&chtml(objsText&ObjsTextComent)&"</textarea>"
		
	else
		response.write objsText
		response.write ObjsTextComent
	end if
	

%>