<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp"-->
<%
	cache 1,"true"
%>
<html>
<head>
<title>Importar Imagem</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
html, body {
  background: ButtonFace;
  color: ButtonText;
  font: 11px Tahoma,Verdana,sans-serif;
  margin: 0px;
  padding: 0px;
}
body { padding: 5px; }
.texto {
	font-family: tahoma, verdana, arial, helvetica, sans-serif;
}
.style1 {font-size: 11px}
.form {
	font-family: tahoma, verdana, arial, helvetica, sans-serif;
	font-size: 11px;
	color: #000000;
}
.style3 {font-family: tahoma, verdana, arial, helvetica, sans-serif; font-size: 10px; }
.foto {
	background-color: #000000;
	border: 1px solid #333333;
}
.style5 {font-family: tahoma, verdana, arial, helvetica, sans-serif; font-size: 11px; color: #000000; font-weight: bold; }
.style6 {
	font-weight: bold;
	font-size: 10px;
}
.style7 {color: #006699}
.style8 {color: #FF0000}
-->
</style>
</head>

<body>
<form action="upload.asp" method="post" enctype="multipart/form-data" name="form">
  <table width="100%"  border="0" cellpadding="0">
    <tr>
      <td class="texto style1"><strong>Upload de Imagem</strong><br>
        <span class="style1">[ Voc&ecirc; pode importar para o banco de dados imagens JPG, PNG, GIF, BMP com at&eacute; 9 Mb ] <%=erros%>
        </span></td>
    </tr>
    <tr>
      <td><hr size="1" noshade></td>
    </tr>
    <tr>
      <td><span class="style5">Nome da imagem:</span>        <input name="nome" type="text" class="form" id="nome" value="Nova Imagem" size="18" maxlength="14">        
      <br>        <input name="1" type="file" class="form" id="1" size="40">
      <input name="Submit" type="submit" class="form" id="Submit" value="Importar"></td>
    </tr>
    <tr>
      <td><span class="texto style1"><strong>Minhas Imagens </strong><br>
          <span class="style1">[Ger&ecirc;ncie as imagens desta area e selecione a imagem que pode ser destacada em algumas areas]<br>
          <br> 
		  <script>
		  
		  function valor(i){
		  var id_i =   
		  }
		  </script>
		  <input name="area" type="hidden" value="<%=request.QueryString("area")%>">
		  <input name="id" type="hidden" value="<%=request.QueryString("id")%>">
</span></span></td>
    </tr>
    <tr>
      <td><hr size="1" noshade></td>
    </tr>
    <tr>
      <td height="198" valign="top">
	  <table width="1%"  border="0" align="center" cellpadding="0" cellspacing="0">
	  <%
	  dim colunas, rs, sql, x
	  db "leitura",rs,sql,"SELECT id, area, foto, nome, dst FROM dbo.myimage WHERE (area LIKE N'"&request.QueryString("area")&"') and id_a="&request.QueryString("id"),conn
	  x = 1
	  
	  colunas = 8
	  while not rs.eof
	  
		
		 if x = 1 then
			response.write "<tr>"
			response.write "<td width=25"&chr(37)&">"
		else
			response.write "<td width=25"&chr(37)&">"
		end if
		
		%>
       
         
		  
		  <table width="11%"  border="0" align="center" cellpadding="0" cellspacing="4">
            <tr>
              <td><table width="90"  border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="foto">
                  <tr>
                    <td width="90" height="87" bgcolor="#FFFFFF"><div align="center"><img src="<%=sites%><%=foto_root&rs("foto")%>" width="70" height="70"></div></td>
                  </tr>
                  <tr>
                    <td bgcolor=<%if trim(rs("dst"))<>0 then%>"#66FF99"<%else%>"#ffffff"<% end if %>>
                      <div align="left">
                      <span class="style6"><a href="#" class="style7" onclick="window.open('op.asp?area=<%=request.QueryString("area")%>&id=<%=request.QueryString("id")%>&id_i=<%=rs("id")%>&op=dst','_self')">D</a> <a href="#" class="style8" onclick="window.open('op.asp?area=<%=request.QueryString("area")%>&id=<%=request.QueryString("id")%>&id_i=<%=rs("id")%>&op=del','_self')">X</a></span>                        <span class="style3"><%=rs("nome")%></span></div></td>
                  </tr>
              </table></td>
            </tr>
          </table>
		  
		
	  <%
	  		
			if x=colunas then
				response.write "</td>"
				response.write "</tr>"
				x = 1
			else
				response.write "</td>"
				x = x + 1
			end if
			
			
	  		rs.movenext
		wend
		if rs.eof and rs.bof then
	  %>
	  <tr>
	  <td valign="middle" width="100%"> <div align="center">Nenhum imagem na cadastrado</div></td>
	  </tr>
	 <%
	 end if 
	%>
	  </table>
	  </td>
    </tr>
    <tr>
      <td><input name="Button" type="button" class="form" value="Voltar" onclick="window.open('<%=site%>/matrix/editor/popups/insert_image.asp?area=<%=request.QueryString("area")%>&id=<%=request.QueryString("id")%>','_self')"></td>
    </tr>
  </table>
  <p>&nbsp;</p>
</form>
</body>
</html>
