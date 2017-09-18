<!--#include virtual ="/bymidia/matrix/config/config.asp" -->
<%
	cache 1, true
	dim rs, sql, x, totvot, op, imgv, imgr, y
	
	imgv = request("imgv")
	imgr = request("imgr")
	
	if imgv = "" then
		imgv = site & "/matrix/img/ok.gif"
	end if
	
	if imgr = "" then
		imgr = site & "/matrix/img/report.gif"
	end if
	
	totvot = empty
	'cookiename = "enquete"
	
	
	
	if request("op") = "" then
		op = "votar"
	else
		op = request("op")
	end if
	
	if request.form("rdoAlt")<>"" then
		
		db "leitura",rs,sql,"select votos, id from enquete_respostas where id="&request.form("rdoAlt"),conn
		db "normal",,sql,"update enquete_respostas set votos="&cint(rs("votos"))+1&" where id="&rs("id"),conn
		op = "result"
		
	end if
	
	'On Error Resume Next
	'ifVoto = cint(request.Cookies(cookiename)(cstr(enId)))
	select case op
	case "votar"
	db "leitura",rs,sql,"select * from enqueteshow where id_a="&request("ida"),conn
%>

  <font id="eEnquetePergunta"><%=chtml(rs("pergunta"))%></font><br>
  <%
  	x = 1
  	while not rs.eof
  %>
  <input name="rdoAlt" type="radio" value="<%=rs("id_r")%>" <% if x = 1 then%>checked<% end if %>> 
  <font id="eEnqueteResposta"><%=chtml(rs("resposta"))%></font><br>
  <%
  		x = x + 1
  		rs.movenext
	wend
	rs.movefirst
  %>	
  <a href="javascript:void(0)" onClick="setCookie('eEnquete<%=rs("id")%>', 'votado', now);  AjaxLoad('<%=site%>/matrix/enquete/inc_enquete.asp?op=result&ida=<%=request("ida")%>&imgv=<%=request("imgv")%>&imgr=<%=request("imgr")%>',allCampQuery('enqForm'),'eEnquete')"><img src="<%=imgv%>" id="eEnqueteVotarimg" border="0"></a>
  <a href="javascript:void(0)" onClick="AjaxLoad('<%=site%>/matrix/enquete/inc_enquete.asp?op=result&ida=<%=request("ida")%>&imgv=<%=request("imgv")%>&imgr=<%=request("imgr")%>',null,'eEnquete')"><img src="<%=imgr%>" border="0" id="eEnqueteResultadoimg"></a>

<%
	x = empty
	case "result"
	db "leitura",rs,sql,"select * from enqueteshow where id_a="&request("ida"),conn
%>
<font id="eEnquetePerguntaResult"><%=chtml(rs("pergunta"))%></font><br>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<%
	if cstr(rs("showvotos")) = "1" then
		totvot = rs("totalvotos")
	end if
	y = 1
	while not rs.eof
			
		x = (1/(rs("totalvotos")/rs("votos")))*100
%>
  <tr>
    <td><div id="alt<%=y%>" style="width:<%=cint(x)%>%; height:10px"><img src="<%=sites%>/matrix/img/blank.gif" width="1" height="10"></div>
    <font id="eEnqueteRespostaResult"><%=chtml(rs("resposta"))%> - </font><font id="eEnquetePorcentResult"><%=cint(x)%>%</font> </td>
  </tr>
  <%
  		y = y + 1
  		rs.movenext
	wend
  %>
</table>
<div id="eEnqueteVotosResult">
<%=totvot%> votos
</div>
<%
	end select
%>
