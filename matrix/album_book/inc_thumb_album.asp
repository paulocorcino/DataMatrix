<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<%
	cache 1,true
	'protec(2)
	
	dim colmax, rs, sql, x, objTable
	Dim strSQL,objPagingConn,Grs,iPageSize,Gpage,iPageCurrent,CONN_STRING, CONN_USER, CONN_PASS, iPageCount, Grid_pgi,Grid_pgf,Grid_semr,iRecordsShown
	
	
	colmax = cint(request.QueryString("itencol"))

		objTable = "<table width='100%' border='0' cellspacing='0' cellpadding='0' id='aAlbumTable'>"
	
			iPageSize = cint(request.QueryString("itenpg"))
			Gpage = request("Gpage")
			
			If Gpage = "" Then
				iPageCurrent = 1
			Else
				iPageCurrent = Gpage
			End If
			
			CONN_STRING = conn
	
			strSQL = "SELECT album_album.id,album_fotos.dest, album_album.album,album_fotos.click, album_fotos.id AS fotoid, album_fotos.foto FROM album_album LEFT JOIN album_fotos ON album_album.id = album_fotos.id_a WHERE (((album_album.id)="&request("id")&"))  AND (dbo.album_fotos.foto IS NOT NULL)"
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
			
			
			x = 0
			iRecordsShown = 0
			Do While iRecordsShown < iPageSize And Not Grs.EOF
				
				if x = 0 then
					objTable = objTable & "<tr>"
				end if
				
   					objTable = objTable & "<td id='aAlbumThumbAreaFoto'><a href=""javascript:imgOpenFoto('"&Grs("fotoid")&"','"&Grs("fotoid")&"','"&Grs("click")&"')"" ><img src='"&sites&foto_root&Grs("foto")&"' id='aAlbumThumbImg'"

					if cstr(request("w")) <> "0" then						
						objTable = objTable & " width="""&request("w")&"""" 
					end if
					
					if cstr(request("h")) <> "0" then
						objTable = objTable & " height="""&request("h")&""""
					end if
					
					objTable = objTable & " border=0></a></td>"
  					x = x + 1
					
				if x = colmax then
					objTable = objTable & "</tr>"
					x = 0
				end if
  
  			iRecordsShown = iRecordsShown + 1
				Grs.MoveNext
			Loop
			
			if iRecordsShown < iPageSize then
			
				for i = iRecordsShown+1 to iPageSize
				
					if x = 0 then
						objTable = objTable & "<tr>"
					end if
					
						objTable = objTable & "<td id='aAlbumThumbAreaFotoVazio'><img src="""&site&"/matrix/img/blank.gif"" id='aAlbumThumbImg'"
						
					if cstr(request("w")) <> "0" then						
						objTable = objTable & " width="""&request("w")&"""" 
					end if
					
					if cstr(request("h")) <> "0" then
						objTable = objTable & " height="""&request("h")&""""
					end if
					
						objTable = objTable & " border=0></td>"
						x = x + 1
						
					if x = colmax then
						objTable = objTable & "</tr>"
						x = 0
					end if
				
				next
			
			end if

			
			objTable = objTable & "</table>"
			
			response.write "<input type='hidden' name='AlbumStat' id='AlbumStat' value='"&Grid_pgi&","&Grid_pgf&"'>" & objTable
  %>
