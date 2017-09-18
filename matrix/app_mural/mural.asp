<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/tnz/matrix/config/config.asp"-->
<%
	cache 1,"false"
	Dim SplitId, sql
%>
<html>
<head>
<title>Promoções</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/estilo.css" rel="stylesheet" type="text/css">
</head>

<body>
<script>
	function SubmitForm(){
		if(document.form.itemsels.value==''){
			alert("Você deve selecionar algum item antes de deletar");
		}else{
			document.form.submit();
		}

	}
	
	function incV(n){
		var ValInc = document.form.itemsels.value;
		
			if(ValInc.indexOf(n)>0){
				document.form.itemsels.value = document.form.itemsels.value.replace(","+n,"");
			}else{
			
				if(ValInc.indexOf(n)==0){
					document.form.itemsels.value = document.form.itemsels.value.replace(n,"");
					if(document.form.itemsels.value.slice(0,1)==","){
						document.form.itemsels.value = document.form.itemsels.value.slice(1,document.form.itemsels.value.length)
					}
				}else{
					if(ValInc==''){
						document.form.itemsels.value = n;
					}else{
						document.form.itemsels.value = ValInc +","+ n;
					}
				}
				
			}
			
		
		}
</script>

<%
	select case request.QueryString("op")
	case "mural"
%>
<table width="95%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="fundo2XP"> &nbsp;<span class="fonte1"><strong>Mural </strong></span></td>
  </tr>
  <tr>
    <td align="center" valign="middle"><br />
	  <p align="center">
	    <%
			 Dim GridData(4), GridName(4), GridSQL, GridDate, GridPoss
				
				
					GridSQL = "SELECT TOP 800 id, para, de, mensagem  FROM dbo.tb_mural ORDER BY id DESC" 
										
					GridName(0) = "-;30"
					GridName(1) = "Cod;30"
					GridName(2) = "De;180"
					GridName(3) = "Para;100"
					GridName(4) = "Mensagem;300"
									
					GridData(0) = "$id;1;0;$case({id}!0):<input type='checkbox' name='codmural' value='{id}' onclick='incV(this.value)'>"
					GridData(1) = "id;0;0"
					GridData(2) = "de;0;0"
					GridData(3) = "para;0;0"
					GridData(4) = "mensagem;0;0"
				
							
					GridDate = "dd/mm/yyyy"
								
					GridForm Gkeys,15
				
%>
        </p></td>
  </tr>
</table>

<%
	case "scrap"
%>
<table width="95%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="fundo2XP"> &nbsp;<span class="fonte1"><strong>Scraps </strong></span></td>
  </tr>
  <tr>
    <td align="center" valign="middle"><br />
	  <p align="center">
	    <%
			 ReDim GridData(6), GridName(6)
				
				
					GridSQL = "SELECT TOP 100 PERCENT dbo.album_coment_foto.id, dbo.album_coment_foto.id_u, dbo.cad_list.nome, dbo.album_coment_foto.foto, dbo.album_coment_foto.coment, dbo.album_coment_foto.apelido FROM dbo.album_coment_foto LEFT OUTER JOIN dbo.cad_list ON dbo.album_coment_foto.id_u = dbo.cad_list.id ORDER BY dbo.album_coment_foto.id DESC" 
					
					GridName(0) = "-;35"
					GridName(1) = "Cod;35"
					GridName(2) = "Cod User;35"
					GridName(3) = "Nome;160"
					GridName(4) = "Apelido;160"
					GridName(5) = "Foto;62"
					GridName(6) = "Comentario;270"

									
					GridData(0) = "$id;1;0;$case({id}!0):<input type='checkbox' name='codscrap' value='{id}' onclick='incV(this.value)'>"
					GridData(1) = "id;1;0"
					GridData(2) = "id_u;1;0"
					GridData(3) = "nome;0;0"
					GridData(4) = "apelido;0;0"
					GridData(5) = "$foto;1;0;$case({foto}!0):<img src=../../fotos/{foto}>"
					GridData(6) = "coment;0;0"



							
					GridDate = "dd/mm/yyyy"
								
					GridForm Gkeys,15
				
%>
        </p></td>
  </tr>
</table>

<%
	case "del"
	
		SplitId = split(formus("itemsels",3),",")
		
		select case formus("tp",3)
		case "scrap"
			'scrap -----------
			
			for i=Lbound(SplitId) to ubound(SplitId)			
				db "normal",,sql,"delete from album_coment_foto where id="&SplitId(i),conn			
			next
			
		case "mural"
			'Mural -----------
			
			for i=Lbound(SplitId) to ubound(SplitId)
				db "normal",,sql,"delete from tb_mural where id="&SplitId(i),conn
			next
			
		end select
		
		voltar request.ServerVariables("SCRIPT_NAME")&"?op="&formus("tp",3)&"&aspid="&aspid(),"","_self"
	
	end select
%>
<form action="<%=request.ServerVariables("SCRIPT_NAME")%>" method="get" name="form">
<input name="op" type="hidden" value="del" />
<input name="aspid" type="hidden" value="<%=aspid%>" />
<input name="tp" type="hidden" value="<%=request.QueryString("op")%>" />
<input name="itemsels" type="hidden" value="" />
<input name="Deletar" type="button" value="Deletar"  onclick="SubmitForm()"/>
</form>