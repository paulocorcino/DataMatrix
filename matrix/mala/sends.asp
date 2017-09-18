<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<%
	cache 1,true
		
	   	Dim strSQL,listemail,sql,rs,objPagingConn,Grs,iPageSize,Gpage,iPageCurrent,CONN_STRING, CONN_USER, CONN_PASS, iPageCount, Grid_pgi,Grid_pgf,Grid_semr,iRecordsShown
		
			iPageSize = 1
			Gpage = request.QueryString("Gpage")
			
			If Gpage = "" Then
				iPageCurrent = 1
			Else
				iPageCurrent = Gpage
			End If
			
			CONN_STRING = conn
	
			strSQL = "select * from mala_email where criado=1 "
			
			if formus("sexo_email",0)<>"2" then
				strSQL = strSQL & " and sexo_email = " & formus("sexo_email",0)
			end if 
			
			if formus("atv_email",0)<>"2" then
				strSQL = strSQL & " and atv_email = " & formus("atv_email",0)
			end if 
			
			'if request.QueryString("cod_grupo")<>"0" then
			'	strSQL = strSQL & " and  cod_grupo = " & request.QueryString("cod_grupo")
			'end if 
			dim splitcodmail
			if formus("cod_grupo",0)<>"" then
				splitcodmail = split(formus("cod_grupo",0),",")
				for i = lbound(splitcodmail) to ubound(splitcodmail)
					if i = 0 then
						strSQL = strSQL & " and  cod_grupo = " & splitcodmail(i)
					else
						strSQL = strSQL & " or  cod_grupo = " & splitcodmail(i)
					end if
				next
			else
				strSQL = strSQL & " and  cod_grupo = 0"
			end if 
			
			
			strSQL = left(strSQL,len(strSQL))
			'response.write strSQL
			Set objPagingConn = Server.CreateObject("ADODB.Connection")
			objPagingConn.Open CONN_STRING, CONN_USER, CONN_PASS
			objPagingConn.CursorLocation = 3
			Set Grs = Server.CreateObject("ADODB.Recordset")
			Grs.CursorLocation = adUseClient
			
			Grs.PageSize = iPageSize
	
			Grs.CacheSize = iPageSize
			
			
			Grs.Open strSQL, objPagingConn, adOpenKeyset, adLockOptimistic
			
			
			iPageCount = Grs.PageCount
			
			If cint(iPageCurrent) > cint(iPageCount) Then iPageCurrent = cint(iPageCount)
			If cint(iPageCurrent) < 1 Then iPageCurrent = 1
			
			If iPageCount = 0 Then
				Grid_semr = 1
			else
				Grs.AbsolutePage = cint(iPageCurrent)
				Grid_pgi = cint(iPageCurrent)
				Grid_pgf = cint(iPageCount)
				
			end if
				
			iRecordsShown = 0
			Do While iRecordsShown < iPageSize And Not Grs.EOF
				
				
				
				
				listemail = split(formus("cod_msg",3),",")
				
				for i = lbound(listemail) to ubound(listemail)
					db "leitura",rs,sql,"select * from mala_msg where cod_msg="&listemail(i),conn
					if not rs.eof and not rs.bof then
						html_msg = trim(rs("html_msg"))
						html_msg = replace(html_msg,"\nome\",trim(Grs("nome_email")))
						html_msg = html_msg & "<p><img src='"&site&"/matrix/malaaction/count.asp?id="&listemail(i)&"' width='1' height='1' ><font face='Verdana, Arial, Helvetica, sans-serif' size='2'>Para remover o seu nome da lista <a href='"&site&"/matrix/malaaction/removeremail.asp?email="&trim(Grs("email_email"))&"' target='_blank'>clique aqui</a></font></p>"
						'//email a enviar
						emails trim(rs("de_msg")),trim(Grs("email_email")),replace(trim(rs("titulo_msg")),"\nome\",trim(Grs("nome_email"))),html_msg,"html"
					end if
				next
				listemail = empty
				
				
				iRecordsShown = iRecordsShown + 1				
				Grs.MoveNext				
			Loop
			
	   %>

<% If cint(iPageCurrent) < cint(iPageCount) Then %>
<%=iPageCurrent+1%>::<p class="texto_size_10">Aguarde Enviando e-mails.....</p>
       <p class="texto_size_10">Enviando grupo <%=Grid_pgi%> de <%=Grid_pgf%>.... </p>
<% 
	else
		
				listemail = split(formus("cod_msg",0),",")
				for i = lbound(listemail) to ubound(listemail)
					db "normal",,sql,"update mala_msg set enviado_msg=1, send_msg="&Grid_pgf&", datasend_msg='"&fdate(now(),"yyyymmdd")&"' where cod_msg="&listemail(i),conn
				next
		

 %>
<%="0"%>::<span class="texto_size_10">Todos os e-mails enviados com sucesso! <br />
	Clique em abortar para concluir esta tarefa.</span>
    <% end if %>	