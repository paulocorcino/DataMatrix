//Album Book Script
// 2.07.2006
function miniMenu_area(){ //menu incluir area

			if(document.form.nome_area.value==''){
				alert("Digite o nome da Área");
			}else{
				document.form.method = "post"
				document.form.action = 'default.asp?op=cad_area&aspid=' + AspId
				document.form.submit();
			}
}

function miniMenu_listemail(){ //menu incluir area
			var emails;
			var sqmail;
			var mailfail = "";
			
			if(document.form.list_mail.value==''){
				alert("Digite os e-mails antes de enviar");
			}else{
				emails = document.form.list_mail.value;
				sqmail = emails.split(";");
				
				for(var i=0;i<sqmail.length;i++){
					
					if(!isEmail(sqmail[i])){
						mailfail = mailfail +", " + sqmail[i];
					}
					
				}
				
				if(mailfail!=""){
					alert("Nesta lista foi encontrados e-mails inválidos:\n" + mailfail);
				}else{
					document.form.method = "post"
					document.form.action = 'default.asp?op=cad_list_email&aspid=' + AspId
					document.form.submit();
				}
				
			}
}

function miniMenu_grupo(){ //menu incluir area

			if(document.form.nome_grupo.value==''){
				alert("Digite o nome do Grupo");
			}else{
				document.form.method = "post"
				document.form.action = 'default.asp?op=cad_grupo&aspid=' + AspId
				document.form.submit();
			}
}

function miniMenu_email(){ //menu incluir area

			if(document.form.nome_email.value==''){
				alert("Digite o nome do usuário");
			}else{
				if(document.form.email_email.value==''){
					alert("Digite o e-mail");
				}else{
					if(document.form.cod_grupo.value==''){
						alert("Selecione um grupo, cadastre os grupos antes de cadastar os e-mails");
					}else{
						document.form.method = "post"
						document.form.action = 'default.asp?op=cad_email&aspid=' + AspId
						document.form.submit();
					}
				}
			}
}

function area_abrir(){
	var area_id = document.form.id_ar.value;
	window.open("default.asp?op=list_album&id_ar="+area_id+"&aspid=" + AspId,"_self");
}

function miniMenu_msg(){
	
			if(document.form.titulo_msg.value==''){
				alert("Digite o assunto da mensagem");
			}else{
				if(document.form.nome_msg.value=='')
				{
					alert("Digite o nome da mensagem");
				}else{
					
					document.form.ta.value = editor.getHTML();
					document.form.method = "post"
					//document.form.encoding = "multipart/form-data"
					document.form.action = 'default.asp?op=cad_msg&aspid=' + AspId
					document.form.submit();
					//alert(editor.getHTML());					
					
				}
			}
}

function stat_noticia(n,i)
{
	var obj = document.getElementById('sCod_vale').value;
	if(obj == ""){
		alert("É necessário selecionar alguma noticia antes de mudar o status");
	}
	else
	{

			document.form.method = "post"
			document.form.action = 'default.asp?op=stat&id_s='+n+'&periodo='+i+'&aspid=' + AspId
			document.form.submit();
		
	}
}

function send_msg()
{
	if(chkValid(document.form.cod_grupo)){
		if(confirm("Todas as opções estão diacordo com sua necessidade?")){
				document.form.method = "get"
				document.form.action = 'default.asp'
				document.form.submit();
		}
	}else{
		alert("Selecione algum grupo");
	}

}

function enviar_msg()
{
	var obj = document.getElementById('sCod_vale').value;
	if(obj == ""){
		alert("É necessário selecionar alguma mensagem antes de enviar");
	}
	else
	{

			document.form.method = "post"
			document.form.action = 'default.asp?op=send&aspid=' + AspId
			document.form.submit();
		
	}
}





function alt_statmala(n)
{
	var obj = document.getElementById('sCod_vale').value;
	if(obj == ""){
		alert("É necessário selecionar algum e-mail antes de alterar o status");
	}
	else
	{
		//if(confirm("Você realmente quer excluir estes comentários?"))
		//{
			//page = 'inc_album_abertos.asp';
			//AjaxLoad(page,'ids='+document.getElementById('fotosel').value+"&op=del&Gpage=1&id="+i,'AlbumAberto');
			//document.getElementById('fotosel').value = "";
			document.form.method = "post"
			//document.form.encoding = "multipart/form-data"
			document.form.action = 'default.asp?op=alt_stats&idgrupo='+n+'&aspid=' + AspId
			document.form.submit();
		//}
	}
}

function excluir_emailmala(n)
{
	var obj = document.getElementById('sCod_vale').value;
	if(obj == ""){
		alert("É necessário selecionar algum email antes de exluir");
	}
	else
	{
		if(confirm("Você realmente quer excluir estes emails?"))
		{
			//page = 'inc_album_abertos.asp';
			//AjaxLoad(page,'ids='+document.getElementById('fotosel').value+"&op=del&Gpage=1&id="+i,'AlbumAberto');
			//document.getElementById('fotosel').value = "";
			document.form.method = "post"
			//document.form.encoding = "multipart/form-data"
			document.form.action = 'default.asp?op=del&tp=email&idgrupo='+n+'&aspid=' + AspId
			document.form.submit();
		}
	}
}

function excluir_msg(n,y,z)
{
	var obj = document.getElementById('sCod_vale').value;
	if(obj == ""){
		alert("É necessário selecionar algum email antes de exluir");
	}
	else
	{
		if(confirm("Você realmente quer excluir estes emails?"))
		{
			//page = 'inc_album_abertos.asp';
			//AjaxLoad(page,'ids='+document.getElementById('fotosel').value+"&op=del&Gpage=1&id="+i,'AlbumAberto');
			//document.getElementById('fotosel').value = "";
			document.form.method = "post"
			//document.form.encoding = "multipart/form-data"
			document.form.action = 'default.asp?op=del&tp=msg&idarea='+n+'&enviado='+y+'&modelo='+z+'&aspid=' + AspId
			document.form.submit();
		}
	}
}

function filtrarmail(obj){
		var emails = obj;
		var process = '';
		var validemail;
		var mailvalid = ' ';
		
		process = emails.value;
		process = process.replace(/["'"]/g," ");
		process = process.replace(/[","]/g," ");
		process = process.replace(/["\""]/g," ");
		process = process.replace(/["<"]/g," ");
		process = process.replace(/[">"]/g," ");
		//process = process.replace(/["http:"]/g," ");
		//process = process.replace(/["ftp:"]/g," ");
		process = process.replace(/["\/"]/g," ");
		process = process.replace(/[\n\r]/g," ");
		process = process.replace(/[\s]/g,";");
		
		validemail = process.split(";");
		emails.value = " ";
		
		for(var i=0;i<validemail.length;i++){
			if(isEmail(validemail[i])){
				emails.value =  validemail[i] + ";" + emails.value;
			}
		}
		
		
		emails.value = emails.value.replace(/[\s]/g,"");
		emails.value = emails.value.substring(0,emails.value.length-1);
		//alert(emails.value.length)
}

function QueryEmail()
{
	var qGrupo	=	document.getElementById("Grupos").value;
	var qSexo	=	document.getElementById("Sexo").value;
	var qStatus	=	document.getElementById("Status").value;
	var qQuery	=	document.getElementById("Querys").value;
	var qTpQuery	=	document.getElementById("Por").value;
	//op=list_email&aspid='+ AspId +'&idgrupo=' + document.getElementById('cod_grupo').value
	var qUrl	=	'default.asp?op=list_email&aspid=' + AspId;
	
	if(qGrupo!="")
	{
		qUrl = qUrl + "&idgrupo=" + qGrupo;
	}
	
	if(qSexo!="")
	{
		qUrl = qUrl + "&idsexo=" + qSexo;
	}
	
	if(qStatus!="")
	{
		qUrl = qUrl + "&idstatus=" + qStatus;
	}
	
	if(qQuery!="")
	{
		qUrl = qUrl + "&query=" + qQuery;
	}
	
	if(qTpQuery!="")
	{
		qUrl = qUrl + "&tpquery=" + qTpQuery;
	}
	
	window.open(qUrl,"_self");
}

