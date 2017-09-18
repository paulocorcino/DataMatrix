<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<%
	
	'Cache
	cache 1, true
	
	dim rs, sql, textomsg
	
			dim snome, semail, anome, aemail, msg
		
		snome = request("albumMyname")
		semail = request("albumMyemail")
		anome = request("albumFriendnome")
		aemail = request("albumFriendemail")
		
	db "leitura",rs,sql,"select foto from album_fotos where id="&request("albumFotoId"),conn
	
	if not rs.eof and not rs.bof then		
		
		msg = "<HR><BR>"
		msg = msg + "<p><CENTER>Indicação de fotografia</CENTER></p>"
		msg = msg + "<HR><BR><BR>"
		msg = msg & now() & "<p>"
		msg = msg + "<img src="&sites&fotog_root&trim(rs("foto"))&"></p>"
		msg = msg + "<p>Caro "&anome& "<BR>"
		msg = msg + "Seu amigo, "&snome&" lhe indicou a foto acima.  Do site "& site &"."&"<BR>"
	'	msg = msg + "<hr><br>Mensagem:<br>"&request.form("msg")
	'	msg = msg + "<br><HR><BR>"
	'	msg = msg + "<CENTER>ToNaZoeira .COM VC</CENTER><BR>"
		msg = msg + "<HR><BR>"
		'msg = msg + "Power by Corcino&Sousa"&"<BR>"
		
		emails semail,aemail,snome & " indica uma foto",msg,"html"
		
	end if
	
	
%>
<img src="<%=site%>/matrix/malaaction/addmail.asp?grupo=Album Enviar Foto&email=<%=semail%>,<%=aemail%>" />