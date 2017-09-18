<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<%
	'Cache
	cache 1, true
	
'Definindo Variaveis da Página
	Dim x, rs, sql, strSQL,objPagingConn,Grs,iPageSize,Gpage,iPageCurrent,CONN_STRING, CONN_USER, CONN_PASS, iPageCount, Grid_pgi,Grid_pgf,Grid_semr,iRecordsShown
		
		dim checked
		
		
		select case request("op")
		case "dest"
			'//Aplicando Destaques
			db "normal",,sql,"update album_fotos set dest=0 where id_a="&formus("id",0),conn
			db "normal",,sql,"update album_fotos set dest=1 where id="&formus("ids",0),conn
			
		case "del"
			'//Excluindo Imagens
			checked = split(request("ids"),",")
			for i=lbound(checked) to ubound(checked)
				'//Apagando Imagens
				delfile fotog_root&checked(i)&".jpg"
				delfile foto_root&checked(i)&".jpg"
				delfile fotoo_root&checked(i)&".jpg"
				
				db "normal",,sql,"delete from album_fotos where id="&checked(i),conn
				
			next
			
			db "leitura",rs,sql,"select top 1 id from album_fotos where dest=1 and id_a="&formus("id",0),conn
			'//Caso não tenha mais destaque automaticamente o banco decreta o novo destaque
			if rs.eof and rs.bof then
				db "leitura",rs,sql,"select top 1 id from album_fotos where id_a="&formus("id",0)&" order by newid()",conn
				db "normal",,sql,"update album_fotos set dest=1 where id="&rs("id"),conn
			end if
			
		end select

		  	'Lista Mensagens...
			
			iPageSize = 18
			Gpage = request("Gpage")
			
			If Gpage = "" Then
				iPageCurrent = 1
			Else
				iPageCurrent = Gpage
			End If
			
			CONN_STRING = conn
	
			strSQL = "SELECT album_album.id,album_fotos.dest, album_album.album, album_fotos.id AS fotoid, album_fotos.foto FROM album_album LEFT JOIN album_fotos ON album_album.id = album_fotos.id_a WHERE (((album_album.id)="&request("id")&"))  AND (dbo.album_fotos.foto IS NOT NULL)"
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
			
			dim img_temp, link_temp, colmax
		  %>
	    <table width="368" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="368"><table width="1%" border="0" cellspacing="3" cellpadding="2">
<%
			colmax = 6
			x = 0
			iRecordsShown = 0
			Do While iRecordsShown < iPageSize And Not Grs.EOF
				
				if x = 0 then
					response.write "<tr>"
					img_temp  = empty
					link_temp = empty
				end if
				
				img_temp  = img_temp & "<td "
				
					if Grs("dest") then 
						img_temp  = img_temp & " bgcolor='#CCCC00' "
					end if 
					
				img_temp  = img_temp & "><img src='"&sites&foto_root&Grs("fotoid")&".jpg' width='86' height='63' id='img_list_ajax' /></td>" & chr(10)
				link_temp = link_temp & "<td id='ajax_img_info'"
					
					if Grs("dest") then 
						link_temp = link_temp & " bgcolor='#CCCC00' "
					end if 
					
				link_temp = link_temp & "><div style='float:left'><input type='checkbox' name='foto_id' value='"&Grs("fotoid")&"' onclick=""incV('"&Grs("fotoid")&"', 'fotosel'); ActualP('"&iPageCurrent&"');"" "
				
					checked = split(request("sel"),",")
					for i=lbound(checked) to ubound(checked)
						if cint(checked(i)) = Grs("fotoid") then
							link_temp = link_temp & " checked "
						end if
					next
				
				link_temp = link_temp & " /></div><div id='ajav_img_nome'>&nbsp;"&Grs("fotoid")&"</div></td>" & chr(10)
				
				x = x + 1
				
				if x = colmax then
					response.write img_temp
						response.write "</tr>"
					response.write "<tr>"
						response.write link_temp
					response.write "</tr>"
					x = 0
				end if
				
				iRecordsShown = iRecordsShown + 1
				Grs.MoveNext
			Loop
			
			if x < colmax and x<>0 then
				response.write img_temp
				response.write "</tr>"
				response.write "<tr>"
				response.write link_temp
				response.write "</tr>"
			end if	
%>
            </table></td>
          </tr>
          <tr>
            <td >
			<form name="form">
				<div id="left" align="center">
				<%
					If cint(iPageCurrent) > 1 Then		 
		 		%>
					<button onclick="AjaxLoad('inc_album_abertos.asp?Gpage=<%=cint(iPageCurrent) - 1%>&id=<%=request("id")%>&aspid=<%=aspid%>','','AlbumAberto');">< Voltar</button>
				<%
					end if			
					
				%>&nbsp;
				</div>
				<% if not Grs.eof and not Grs.bof or grid_pgf > 1 then %>
				<div id="left" align="center">
					<select name="Murpages" class="fonte_10" onchange="AlbumPage(this.value,'<%=request("id")%>');">
<%
					for GridX = 1 to grid_pgf
				 	
					
						response.write  "<option value='"&GridX&"'"
						
							if cint(iPageCurrent) = cint(GridX) then
								response.write " selected "
							end if
						
						response.write  ">"&GridX&"</option>"
				 
					next
%>
					</select>
					
				</div>
				<% end if %>
				<div id="left" align="center">
				<%
					If cint(iPageCurrent) < cint(iPageCount) Then
				%>
					<button onclick="AjaxLoad('inc_album_abertos.asp?Gpage=<%=cint(iPageCurrent) + 1%>&id=<%=request("id")%>&aspid=<%=aspid%>','','AlbumAberto');">Proximo ></button>
				<%
					end if
				%>&nbsp;
				</div>
			  </form>
			</td>
          </tr>
        </table>
