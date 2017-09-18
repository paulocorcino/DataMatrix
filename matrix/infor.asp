<%
cache 1,"false"
dim infor_rs,infor_sql
 db "leitura",infor_rs,infor_sql,"select nome_comp,email,inf,id from senha where inf=1 and inf_data<>"&CampoData( cDate( date() & " " & time() ), false ),conn
 	while not infor_rs.eof 
		'response.write rs("email")&"<br>"
		emails "pecjs@msn.com",infor_rs("email"),"Data Matrixs - Estatisticas "&nome_site,informativo(infor_rs("nome_comp"),date()),"html"
		db "normal",,infor_sql,"update senha set inf_data="&CampoData( cDate( date() & " " & time() ), false )&" where id="&rs("id"),conn
	infor_rs.movenext
	wend
%>