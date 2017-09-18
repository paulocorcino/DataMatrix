<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<%
	protec(2)
	cache 1,false
	
%>
<html>
<head>
<title>Enquete</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/estilo.css" rel="stylesheet" type="text/css">
<script type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_validateForm() { //v4.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
      } else if (test!='R') { num = parseFloat(val);
        if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (num<min || max<num) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' é orbigatório.\n'; }
  } if (errors) alert('Alguns erros foram encontrados:\n'+errors);
  document.MM_returnValue = (errors == '');
}
//-->
</script>
</head>
<script language="JavaScript" type="text/javascript" src="../js/form.js"></script>
<script language="JavaScript" type="text/javascript" src="../js/script.js"></script>

<body>
<%
	dim sql,rs,x,alter,id_p,alt, rs2, votos, totalvotos
	dim cor

	'numero maximo de alternativas
	alter = 6
	x = empty
	select case request.QueryString("op")
		case "cad1"
		
		'verifica se já existe alguma informação cadastrada
		alt = 0
		if request.QueryString("alt")<>"" and request.QueryString("alt")<>0 then
			
			if request.QueryString("id")<>"" and request.QueryString("id")<>0 then
			
				db "leitura",rs,sql,"select id, id_a, criado, entrada, saida, pergunta,alt from dbo.enquetefoto_perguntas where id="&request.QueryString("id"),conn
				er(30)
				x = trim(rs("pergunta"))
				
				alt = 1
			
			else
			
				db "leitura",rs,sql,"select id, id_a, criado, entrada, saida, pergunta,alt from dbo.enquetefoto_perguntas where criado=0 ",conn
				er(37)
				x = trim(rs("pergunta"))
			
			end if
			
			
		else
	
			db "leitura",rs,sql,"select id, id_a, criado, entrada, saida, pergunta,alt from enquetefoto_perguntas where criado=0",conn
			er(47)
				if rs.eof and rs.bof then
				
					db "normal",,sql,"insert into enquetefoto_perguntas(criado,alt,totalvotos,repeteuser) values (0,2,1,0)",conn
					db "leitura",rs,sql,"select id, id_a, entrada, saida, criado,pergunta,alt from enquetefoto_perguntas where criado=0",conn
					er(50)
			
				end if
	
		end if
		session("id_a") = empty
		
		

%>
<table width="98%"  border="0" cellspacing="2" cellpadding="0">
  <tr>
    <td class="fundo2">&nbsp;&nbsp;<span class="fonte"><strong>Nova Foto Enquete - Passo 1 / 4 </strong></span></td>
  </tr>
  <tr>
    <td><%=erros%></td>
  </tr>
  <tr>
    <td><form action="<%=request.ServerVariables("SCRIPT_NAME")%>?op=cad2&aspid=<%=aspid()%>&alt=0&aspid=<%=aspid()%>" method="post" name="form" onSubmit="MM_validateForm('pergunta','','R');return document.MM_returnValue">
      <table width="98%"  border="0" align="center" cellpadding="0" cellspacing="2">
        <tr>
          <td><span class="fonte"><strong>Categoria</strong></span><br>
            <span class="fonte">Selecione a categoria que esta enquete pertence </span></td>
        </tr>
        <tr>
          <td>
		  <select name="id_a" class="formtextarea" id="id_a">
		  	<%
				db "leitura",rs2,sql,"select * from enquetefoto_areas",conn
				while not rs2.eof
			%>
					<option value="<%=rs2("id")%>" 
					<% 
					if rs("id_a")<>"" then
						if cstr(rs("id_a")) = cstr(rs2("id")) then 
					%> selected <%
						end if
					end if
					%> ><%=trim(rs2("areas"))%></option>
			<%
					rs2.movenext
				wend
			%>		
          </select>
          </td>
        </tr>
        <tr>
          <td><span class="fonte"><strong><br>
            Pergunta</strong></span><br>
            <span class="fonte">Digite a pergunta de sua enquete. </span></td>
        </tr>
        <tr>
          <td><input name="pergunta" type="text" class="formtextarea" id="pergunta" size="50" maxlength="255" value="<%=x%>"></td>
        </tr>
        <tr>
          <td><br>
            <span class="fonte"><strong>Selecione o Peri&oacute;do</strong><br>
            Selecione o peri&oacute;do em que sua enquete ira entrar no ar </span></td>
        </tr>
        <tr>
          <td>
            <input name="entrada" type="text" class="formtextarea" id="entrada" value="<%=fdate(rs("entrada"),"dd/mm/yyyy")%>" onBlur="if(this.value==''){this.value='<%=fdate(rs("entrada"),"dd/mm/yyyy")%>'}else{dataVal(this)}" onFocus="this.value=''"  size="13" maxlength="10"> 
            <span class="fonte">at&eacute;</span> <input name="saida" type="text" class="formtextarea" id="saida" value="<%=fdate(rs("saida"),"dd/mm/yyyy")%>" onBlur="if(this.value==''){this.value='<%=fdate(rs("saida"),"dd/mm/yyyy")%>'}else{dataVal(this)}" onFocus="this.value=''" size="13" maxlength="10">
            &nbsp;
            <input name="Submit2" type="button" class="botao" onClick="MM_openBrWindow('check.asp?incio='+document.getElementById('entrada').value+'&fim='+document.getElementById('saida').value+'&id_a='+document.getElementById('id_a').value,'testdata','width=150,height=90')" value="Checar Data"></td>
        </tr>
        <tr>
          <td><br>
            <span class="fonte"><strong>Alternativas</strong><br>
            Selecione o numero de alternativas</span></td>
        </tr>
        <tr>
          <td><select name="alt" class="formtextarea" id="alt">
		  <%
		  	x = empty
			for x=2 to alter 
		  %>
            <option value="<%=x%>" <% if x=rs("alt") then %> selected <% end if %>><%=x%></option>
		<%
			next
		%>
          </select></td>
        </tr>
        <tr>
          <td><input type="hidden" name="id" value="<%=rs("id")%>"><input type="hidden" name="cad" value="1">
			             <br>
            <input name="Submit" type="submit" class="botao" value="Avan&ccedil;ar &gt;">
          &nbsp;
          <input name="button" type="button" class="botao" id="button" value="Cancelar" onClick="window.open('<%=request.ServerVariables("SCRIPT_NAME")%>?aspid=<%=aspid()%>','_self')"></td>
        </tr>
      </table>
    </form></td>
  </tr>
</table>
		<script>
	if(isIE()){
		oDateMasks = new Mask("dd/mm/yyyy", "date");
		oDateMasks.attach(document.form.entrada);
		oDateMasks.attach(document.form.saida)
		
	}
	</script>
<%
	fdb(rs)
	
	case "cad2"
	
	if request.QueryString("alt")<>"" and request.QueryString("alt")<>0 then
		
			db "leitura",rs,sql,"SELECT dbo.enquetefoto_perguntas.id, dbo.enquetefoto_perguntas.pergunta,dbo.enquetefoto_perguntas.alt, dbo.enquetefoto_respostas.id AS id_r, dbo.enquetefoto_respostas.fotos, dbo.enquetefoto_respostas.votos FROM dbo.enquetefoto_perguntas INNER JOIN dbo.enquetefoto_respostas ON dbo.enquetefoto_perguntas.id = dbo.enquetefoto_respostas.id_p WHERE (dbo.enquetefoto_perguntas.id = "&request.querystring("id")&request.form("id")&")",conn
			er(122)
			alt=1
	else
	'response.Write "teste"
		db "leitura",rs,sql,"SELECT id, pergunta, entrada, ativo FROM dbo.enquetefoto_perguntas WHERE (id_a = "&formus("id_a",0)&") AND (ativo = 1) AND (criado = 1) and (id <> "&request("id")&") and (entrada >= CONVERT(DATETIME, '"&fdate(formus("entrada",0),"yyyy-mm-dd")&" 00:00:00', 102)) AND (saida <= CONVERT(DATETIME, '"&fdate(formus("saida",0),"yyyy-mm-dd")&" 00:00:00', 102)) OR (entrada <= CONVERT(DATETIME, '"&fdate(formus("saida",0),"yyyy-mm-dd")&" 00:00:00', 102)) AND (saida >= CONVERT(DATETIME, '"&fdate(formus("entrada",0),"yyyy-mm-dd")&" 00:00:00', 102)) and (id <> "&request("id")&") and (id_a = "&formus("id_a",0)&")",conn
		er(127)
		if rs.eof and rs.bof then
			
			if not trim(request.form("pergunta"))<>"" then
			
				voltar request.ServerVariables("SCRIPT_NAME")&"?op=cad1&aspid="&aspid()&"&alt=1&id="&formus("id",0),"É necessário que você digite uma pergunta.","_self"
				fdb(conn)
				response.End()

			else
				'x ira armazenar a data
				x = fDate(formus("entrada",0),"yyyymmdd")
				'alert(x)
				'x = "Convert(Datetime,'"&year(x)&"-"&month(x)&"-"&day(x)&"',102)"
				'alert(x)
				session("id_a") = formus("id_a",0)
				db "normal",,sql,"update enquetefoto_perguntas set pergunta='"&formus("pergunta",1)&"',entrada='"&x&"',saida='"&fdate(formus("saida",0),"yyyymmdd")&"',ativo=1,alt="&formus("alt",0)&" where id="&formus("id",0),conn
				er(143)
				' x esvasia
				x = empty
				
				db "leitura",rs,sql,"SELECT dbo.enquetefoto_perguntas.id, dbo.enquetefoto_perguntas.pergunta, dbo.enquetefoto_perguntas.alt, dbo.enquetefoto_respostas.id AS id_r, dbo.enquetefoto_respostas.fotos, dbo.enquetefoto_respostas.votos FROM dbo.enquetefoto_perguntas INNER JOIN dbo.enquetefoto_respostas ON dbo.enquetefoto_perguntas.id = dbo.enquetefoto_respostas.id_p WHERE (dbo.enquetefoto_perguntas.id = "&request.querystring("id")&request.form("id")&")",conn
				er(148)
				'numero do id da pergunta
				id_p=formus("id",0)
				alt = 1
					if rs.eof and rs.bof then
						for x=1 to formus("alt",0)
				
							db "normal",,sql,"insert into enquetefoto_respostas(id_p,votos) values ("&id_p&",1)",conn
							er(156)
						next
						
							db "leitura",rs,sql,"SELECT dbo.enquetefoto_perguntas.id, dbo.enquetefoto_perguntas.pergunta,dbo.enquetefoto_perguntas.alt, dbo.enquetefoto_respostas.id AS id_r, dbo.enquetefoto_respostas.fotos, dbo.enquetefoto_respostas.votos FROM dbo.enquetefoto_perguntas INNER JOIN dbo.enquetefoto_respostas ON dbo.enquetefoto_perguntas.id = dbo.enquetefoto_respostas.id_p WHERE (dbo.enquetefoto_perguntas.id = "&request.querystring("id")&request.form("id")&")",conn
							er(160)
							alt = 0
						end if
			end if
			
		else
		
			'response.Write sql
			voltar request.ServerVariables("SCRIPT_NAME")&"?op=cad1&aspid="&aspid()&"&alt=1&id="&formus("id",0),"Já existe uma enquete programada para este periódo. Escolha outra data.","_self"
			fdb(rs)
			fdb(conn)
			response.End()
			
		end if
	end if
	
%>
<table width="98%"  border="0" cellspacing="2" cellpadding="0">
  <tr>
    <td class="fundo2">&nbsp;&nbsp;<span class="fonte"><strong>Nova Foto Enquete - Passo 2 / 4 </strong></span></td>
  </tr>
  <tr>
    <td><%=erros%></td>
  </tr>
  <tr>
    <td><form action="upload.asp"  method="post" name="formsa" enctype="multipart/form-data">
      <table width="98%"  border="0" align="center" cellpadding="0" cellspacing="2">
        <tr>
          <td class="fonte"><strong>Pergunta</strong></td>
        </tr>
        <tr>
          <td class="fonte"><%=trim(rs("pergunta"))%></td>
        </tr>
        <tr>
          <td class="fonte">&nbsp;</td>
        </tr>
        <tr>
          <td class="fonte"><strong>Alternativas</strong><br>
            Descreva abaixo as alternativas</td>
        </tr>
        <tr>
          <td class="fonte"><table width="95%"  border="0" align="center" cellpadding="0" cellspacing="2">
		  <%
		 ' x = 1
		  	for x=1 to rs("alt")
		%>  
		<tr class="fonte">
              <td>Alternativa <%=x%> </td>
              <td colspan="2">Votos</td>
            </tr>
            <tr>
		<%
			
			if rs.eof and rs.bof or rs.eof then
		%>
				<td width="17%"><input name="alt<%=x%>" class="fonte" id="alt<%=x%>" size="30" type="file"></td>
              	<td width="5%"><input name="votos<%=x%>" type="text" class="fonte" id="votos<%=x%>"  size="5" value="0">
				
				</td>
		<%
				else
		%>
 				<td width="17%">
				<%
					if rs("fotos") <> "" then
				%>
				<%=trim(rs("fotos"))%><input name="calt<%=x%>" type="checkbox" value="alterar" onClick="if(!this.checked){document.getElementById('alt<%=x%>').disabled=false}else{document.getElementById('alt<%=x%>').disabled=true}" checked="checked">
 				  <span class="fonte">N&atilde;o</span> <span class="fonte">Alterar Imagem</span><br>
				 <% end if %>
 				<input name="alt<%=x%>" class="fonte" id="alt<%=x%>" size="30" type="file" onChange="previewImg(this,'img<%=x%>','jpeg,jpg,png,gif,bmp','0,90')" <% if rs("fotos") <> "" then %> disabled="disabled" <% end if %>></td>
              	<td width="5%"><input name="votos<%=x%>" type="text" class="fonte" id="votos<%=x%>"  size="5" value="<%=cint(rs("votos")) - 1%>">
								<input type="hidden" name="id_r<%=x%>" value="<%=rs("id_r")%>">
				</td>
             
		<%
				rs.movenext
			end if
			er(229)
			
			
		%>

              <td width="78%"><div id="img<%=x%>">&nbsp;</div></td>
            </tr>
			<%
			next
			%>
          </table></td>
        </tr>
		        <tr>
          <td ><br><input type="hidden" name="id" value="<%=request.QueryString("id")&formus("id",0)%>">
            <input name="Submit3" type="button" class="botao" value="&lt; Voltar" onClick="window.open('<%=request.ServerVariables("SCRIPT_NAME")%>?op=cad1&aspid=<%=aspid()%>&alt=1&id=<%=request.form("id")%><%=request.QueryString("id")%>','_self')">
            <input name="Submit" type="submit" class="botao" id="Submit" value="Avan&ccedil;ar &gt;"></td>
		        </tr>
      </table>
    </form></td>
  </tr>
</table>
<%
	case "cad3"

	db "leitura",rs,sql,"select * from enquetefoto_perguntas where id="&request("id"),conn
	'db "leitura",rs,sql,"SELECT enquetefoto_perguntas.id, enquetefoto_perguntas.pergunta, enquetefoto_perguntas.alt,respostas.id, respostas.resposta, totalfoto_votos.Votos FROM (enquetefoto_perguntas INNER JOIN respostas ON enquetefoto_perguntas.id = respostas.id_p) INNER JOIN totalfoto_votos ON respostas.id = totalfoto_votos.id WHERE (((enquetefoto_perguntas.id)="&request.querystring("id")&formus("id",0)&"))",conn
	'//-------------------------------------------------------------------------------------------------------------------------------------------
	
%>
<table width="98%"  border="0" cellspacing="2" cellpadding="0">
  <tr>
    <td class="fundo2">&nbsp;&nbsp;<span class="fonte"><strong>Nova Foto Enquete - Passo 3 / 4 </strong></span></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><form name="form3" method="post" action="<%=request.ServerVariables("SCRIPT_NAME")%>?op=cad4&aspid=<%=aspid()%>">
      <table width="98%"  border="0" align="center" cellpadding="0" cellspacing="2">
        <tr>
          <td class="fonte"><strong>Pergunta</strong></td>
        </tr>
        <tr>
          <td class="fonte"><%=rs("pergunta")%><input type="hidden" name="id" value="<%=rs("id")%>"></td>
        </tr>
        <tr>
          <td class="fonte">&nbsp;</td>
        </tr>
        <tr>
          <td class="fonte"><strong>Deseja que apare&ccedil;a o numero total de votos ? </strong></td>
        </tr>
        <tr>
          <td class="fonte"><select name="totalvoto" class="formtextarea" id="totalvoto">
            <option value="0" <%if trim(rs("totalvotos"))=0 then%>selected<%end if%>>N&atilde;o</option>
            <option value="1" <%if trim(rs("totalvotos"))=1 then%>selected<%end if%>>Sim</option>
                    </select></td>
        </tr>
        <tr>
          <td class="fonte"><strong>Deseja permitir que a mesma pessoa vote mais de uma vez? </strong></td>
        </tr>
        <tr>
          <td height="21" class="fonte"><select name="repeteuser" class="formtextarea" id="repeteuser">
            <option value="0" <%if trim(rs("repeteuser"))=0 then%>selected<%end if%>>N&atilde;o</option>
            <option value="1" <%if trim(rs("repeteuser"))=1 then%>selected<%end if%>>Sim</option>
          </select></td>
        </tr>
        <tr>
          <td><br>
            <input name="button" type="button" class="botao" id="button" value="&lt; Voltar" onClick="window.open('<%=request.ServerVariables("SCRIPT_NAME")%>?op=cad2&aspid=<%=aspid()%>&alt=1&id=<%=request.form("id")%><%=request.QueryString("id")%>','_self')">
            <input name="Submit" type="submit" class="botao" id="Submit" value="Concluir"></td>
        </tr>
      </table>
    </form></td>
  </tr>
</table>
<%
	case "cad4"
	
	db "normal",,sql,"update enquetefoto_perguntas set criado=1, id_a="&session("id_a")&", totalvotos="&formus("totalvoto",0)&", repeteuser="&formus("repeteuser",0)&" where id="&formus("id",0),conn
	
%>
<table width="98%"  border="0" cellspacing="2" cellpadding="0">
  <tr>
    <td class="fundo2">&nbsp;&nbsp;<span class="fonte"><strong>Nova Foto Enquete - Passo 4 / 4 </strong></span></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><form name="form3" method="post" action="">
      <table width="98%"  border="0" align="center" cellpadding="0" cellspacing="2">
        <tr>
          <td height="141" class="fonte"><div align="center">Sua enquete foi criada com Exito!<br>
            </div></td>
        </tr>
        <tr>
          <td><br>
            <input name="button" type="button" class="botao" id="button" value="Fechar" onClick="window.open('<%=request.ServerVariables("SCRIPT_NAME")%>?aspid=<%=aspid()%>','_self')"></td>
        </tr>
      </table>
    </form></td>
  </tr>
</table>
<%

	case "atv"
	
		x = request.QueryString("atv")
		
		if x = 1 then
	
			db "normal",rs,sql,"update enquetefoto_perguntas set ativo=0 where id="&request.QueryString("id"),conn
	
		else
	
			db "normal",rs,sql,"update enquetefoto_perguntas set ativo=1 where id="&request.QueryString("id"),conn
	
		end if
		voltar request.ServerVariables("SCRIPT_NAME")&"?id_a="&formus("id_a",3)&"&aspid="&aspid(),"","_self"
	
	case "del"
	
		db "normal",,sql,"delete from enquetefoto_perguntas where id="&request.QueryString("id"),conn
		voltar request.ServerVariables("SCRIPT_NAME")&"?id_a="&formus("id_a",3)&"&aspid="&aspid(),"","_self"
	case "result"
	
	%>
<table width="100%" border="0" cellspacing="2" cellpadding="0">
      <tr>
        <td class="fundo2">&nbsp;&nbsp;<span class="fonte"><strong>Resultado Percial  </strong></span></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>
		<%
	

	',x,cookiename,totalvotos,votos,cor
	'------------^



	'db "leitura",rs,sql,"SELECT respostas.id, perguntas.pergunta, perguntas.alt, respostas.resposta, Sum(totalfoto_votos.SomaDevotos) AS totalvotos, (1/(totalvotos/votos))*100 AS porcenta FROM totalfoto_votos INNER JOIN (perguntas INNER JOIN respostas ON perguntas.id = respostas.id_p) ON totalfoto_votos.id = perguntas.id GROUP BY respostas.id, perguntas.id, perguntas.pergunta, perguntas.alt, respostas.resposta, respostas.votos, (1/(totalvotos/votos))*100 HAVING (((perguntas.id)="&request.QueryString("id")&")) ORDER BY respostas.id",conn
	db "leitura",rs,sql,"SELECT TOP 100 PERCENT dbo.enquetefoto_perguntas.id, dbo.enquetefoto_respostas.id AS id_r, dbo.enquetefoto_perguntas.pergunta, dbo.enquetefoto_respostas.fotos, dbo.enquetefoto_perguntas.alt, dbo.totalfoto_votos.SomaDevotos AS totalvotos, dbo.enquetefoto_respostas.votos FROM dbo.enquetefoto_perguntas INNER JOIN  dbo.enquetefoto_respostas ON dbo.enquetefoto_perguntas.id = dbo.enquetefoto_respostas.id_p INNER JOIN dbo.totalfoto_votos ON dbo.enquetefoto_perguntas.id = dbo.totalfoto_votos.id WHERE(dbo.enquetefoto_perguntas.id = "&request.QueryString("id")&")",conn
	er(15)
%>
<link rel="stylesheet" type="text/css" href="../css/enquete.css">
<table width="100%" height="220" border="0" align="center" cellpadding="0" cellspacing="2" <% if request.QueryString("janela")=1 then%><%end if%>>
  <tr>
    <td align="center" valign="top"><form name="form1" method="post" action="">
      <table width="80%" height="100%" border="0" cellpadding="0" cellspacing="1">
        <tr>
          <td height="31" valign="top"><div class="titulo">
            <%=rs("pergunta")%>
          </div></td>
        </tr>
        <tr>
          <td height="108" valign="top"><table width="100%"  border="0" align="center" cellpadding="0" cellspacing="3">
		<%
			cor = 0
			
			while not rs.eof 
			
			totalvotos = cint(rs("totalvotos"))
			votos = cint(rs("votos"))
			x = (1/(totalvotos/votos))*100
			cor = cor + 1
		%>
            <tr>
              <td align="center" valign="middle" class="alternativas"><div align="left"><%=rs("fotos")%></div></td>
            </tr>
            <tr>
              <td align="center" valign="top" class="alternativas"><div align="left">
                <table width="100%"  border="0" cellpadding="0" cellspacing="1">
                  <tr>
                    <td width="<%if cint(x)=0 then%>1<%else%><%=cint(x)%><%end if%>%" class="cor<%if request.QueryString("cor")<>0 then%><%=cor%><%else%>1<%end if%>"><img src="../img/blank.gif" width="1%" height="10"></td>
                    <td width="1" class="alternativas"><div align="left">&nbsp;<%=formatnumber(x,1)%>%</div></td>
                  </tr>
                </table>
              </div></td>
            </tr>
			<%
				rs.movenext
			wend
			rs.movefirst
			%>
          </table></td>
        </tr>
        <tr>
          <td height="20" valign="middle"><div align="center" class="votos"><%=rs("totalvotos") - rs("alt")%> votos</div></td>
        </tr>
      </table>
    </form></td>
  </tr>
</table>
<%
	er("acima")
	fdb(rs)
	

		%>
		</td>
      </tr>
    </table>
<%
	
	case else
	
	'db "leitura",rs,sql,"SELECT TOP 100 PERCENT dbo.enquetefoto_perguntas.id, dbo.enquetefoto_perguntas.pergunta, dbo.enquetefoto_perguntas.criado, dbo.enquetefoto_perguntas.ativo, dbo.enquetefoto_perguntas.alt, dbo.enquetefoto_perguntas.entrada, dbo.totalfoto_votos.SomaDevotos AS totalvotos FROM dbo.enquetefoto_perguntas INNER JOIN dbo.totalfoto_votos ON dbo.enquetefoto_perguntas.id = dbo.totalfoto_votos.id WHERE (dbo.enquetefoto_perguntas.criado = 1) ORDER BY dbo.enquetefoto_perguntas.entrada DESC",conn 
	
	prep_deletar()
%>
<table width="98%"  border="0" cellspacing="2" cellpadding="0">
  <tr>
    <td class="fundo2">&nbsp;&nbsp;<span class="fonte"><strong>Listar Foto Enquetes </strong></span></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>
	<span class="fonte">Selecione a &Aacute;rea</span><br>
	<form name="form1" method="post">
	  <select name="id_a" class="formtextarea" onChange="window.open('<%=request.ServerVariables("SCRIPT_NAME")%>?aspid=<%=aspid%>&id_a='+this.value,'_self');">
	  <option value="" selected>Selecione a Area</option>
	  <%
	  	db "leitura",rs,sql,"select * from enquetefoto_areas",conn
		while not rs.eof
	  %>
	  	<option value="<%=rs("id")%>"><%=rs("areas")%></option>
	<%
			rs.movenext
		wend
	%>
	    </select>
	  </form>
	  <div align="center">
	    <%
				  Dim GridData(6), GridName(6), GridSQL, GridDate, GridPoss
				
				
					GridSQL = "select * from show_enquetefoto where id_a = "&formus("id_a",3)&";"
					
					GridName(0) = "Pergunta;250"
					GridName(1) = "Entrada;80"
					GridName(2) = "Saída;80"
					GridName(3) = "Votos;80"
					GridName(4) = "Ativado;60"
					GridName(5) = "Alterar;60"
					GridName(6) = "Deletar;60"
									
					GridData(0) = "pergunta;0;"&request.ServerVariables("SCRIPT_NAME")&"?op=result&aspid=[aspid]&id_a={id_a}&id={id}"
					GridData(1) = "entrada;0;0"
					GridData(2) = "saida;0;0"
					GridData(3) = "totalvotos;1;0"
					GridData(4) = "$ativo;1;"&request.ServerVariables("SCRIPT_NAME")&"?op=atv&aspid=[aspid]&id_a={id_a}&atv={ativo}&id={id};$case({ativo}=1):On-Line:Off-Line"
					GridData(5) = "*Alterar;1;enquete.asp?op=cad1&aspid=[aspid]&alt=1&id={id}"
					GridData(6) = "*Deletar;1;javascript: deletar('"&request.ServerVariables("SCRIPT_NAME")&"?op=del&aspid=[aspid]&id={id}','[col0]')"
							
					GridDate = "dd/mm/yyyy"
					GridPoss = "center"
					'GridURLS = 1
					'// Chama a Função
					if request.QueryString("id_a")<>"" then
						GridForm Gkeys,15
					end if
	%>	
      </div></td>
  </tr>
</table>
<%
	
	end select
	fdb(conn)
%>
</body>
</html>
