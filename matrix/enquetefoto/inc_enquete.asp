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
		
		db "leitura",rs,sql,"select votos, id from enquetefoto_respostas where id="&request.form("rdoAlt"),conn
		db "normal",,sql,"update enquetefoto_respostas set votos="&cint(rs("votos"))+1&" where id="&rs("id"),conn
		op = "result"
		
		
	end if
	
	'On Error Resume Next
	'ifVoto = cint(request.Cookies(cookiename)(cstr(enId)))
	select case op
	case "votar"
	db "leitura",rs,sql,"select * from EnqueteFotoShow where id_a="&request("ida"),conn
%>
<table width="100%" border="0" cellspacing="3" cellpadding="4">
  <tr>
    <td colspan="6"><font id="feEnquetePergunta"><%=chtml(rs("pergunta"))%></font></td>
  </tr>
  <tr>  
  <%
  	x = 1
  	while not rs.eof
  %>
    <td width="<%=100/rs.RecordCount%>%"><p>
      <input name="rdoAlt" type="radio" value="<%=rs("id_r")%>" <% if x = 1 then%>checked<% end if %>>   <font id="feEnqueteResposta"><%=chtml(rs("titulo"))%></font></p>
    <p align="center"><img src="<%=sites&foto_root&trim(chtml(rs("fotos")))%>"/>
    <div id="feAmpliar">[<a href="javascript: void(0);" onclick="MM_openBrWindow('<%=site&"/matrix/enquetefoto/fotoshow.asp?url="&sites&fotog_root&chtml(rs("fotos"))%>','fotoenquete','width=<%=wAlbumFull+5%>,height=<%=hAlbumFull+5%>')">Ampliar</a>]</div></p></td>
  <%
  		x = x + 1
  		rs.movenext
	wend
	rs.movefirst
  %>
  </tr>
  <tr>
    <td colspan="6">  <a href="javascript:void(0)" onClick="setCookie('feEnquete<%=rs("id")%>', 'votado', now);  AjaxLoad('<%=site%>/matrix/enquetefoto/inc_enquete.asp?op=result&ida=<%=request("ida")%>&imgv=<%=request("imgv")%>&imgr=<%=request("imgr")%>',allCampQuery('enqfForm'),'feEnquete')"><img src="<%=imgv%>" id="feEnqueteVotarimg" border="0"></a>
  <a href="javascript:void(0)" onClick="AjaxLoad('<%=site%>/matrix/enquetefoto/inc_enquete.asp?op=result&ida=<%=request("ida")%>&imgv=<%=request("imgv")%>&imgr=<%=request("imgr")%>',null,'feEnquete')"><img src="<%=imgr%>" border="0" id="feEnqueteResultadoimg"></a></td>
  </tr>
</table>
<%
	x = empty
	case "result"
	db "leitura",rs,sql,"select * from EnqueteFotoShow where id_a="&request("ida"),conn
%>
<table width="100%" border="0" cellspacing="3" cellpadding="4">
  <tr>
    <td colspan="6"><font id="feEnquetePerguntaResult"><%=chtml(rs("pergunta"))%></font></td>
  </tr>
    <tr>  
  <%
  	if cstr(rs("showvotos")) = "1" then
		totvot = rs("totalvotos") - rs.RecordCount
	end if
	y = 1
	
  	while not rs.eof
		x = (1/(rs("totalvotos")/rs("votos")))*100
  %>
    <td width="<%=100/rs.RecordCount%>%"><p>
      <font id="feEnqueteResposta"><%=chtml(rs("titulo"))%></font> - <font id="feEnquetePorcentResult"><%=cint(x)%>%</font><div id="alt<%=y%>" style="width:<%=cint(x)%>%; height:10px"><img src="<%=sites%>/matrix/img/blank.gif" width="1" height="10"></div></p>
    <p align="center"><img src="<%=sites&foto_root&trim(chtml(rs("fotos")))%>"/><br />
    <div id="feAmpliar">[<a href="javascript: void(0);" onclick="MM_openBrWindow('<%=site&"/matrix/enquetefoto/fotoshow.asp?url="&sites&fotog_root&chtml(rs("fotos"))%>','fotoenquete','width=<%=wAlbumFull+5%>,height=<%=hAlbumFull+5%>')">Ampliar</a>]</div></p></td>
  <%
  		y = y + 1
  		rs.movenext
	wend

  %>
  </tr>
</table>
<div id="feEnqueteVotosResult">
<%=totvot%> votos
</div>
<%
	end select
%>
