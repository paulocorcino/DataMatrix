//Album Book Script
// 2.07.2006
function miniMenu_area(){ //menu incluir area

			if(document.form.area.value==''){
				alert("Digite o nome da Área");
			}else{
				document.form.method = "post"
				document.form.action = 'default.asp?op=cad_area&aspid=' + AspId
				document.form.submit();
			}
}

function area_abrir(){
	var area_id = document.form.id_ar.value;
	window.open("default.asp?op=list_album&id_ar="+area_id+"&aspid=" + AspId,"_self");
}

function miniMenu_noticia(){
	
			if(document.form.ntitulo.value==''){
				alert("Digite o titulo da noticia");
			}else{
				if(document.form.ndata.value=='')
				{
					alert("Digite a data de publicação");
				}else{
					if(document.form.ncoment.value==''){
						alert("Digite um comentário");
					}else{
							document.form.ta.value = editor.getHTML();
							document.form.method = "post"
							//document.form.encoding = "multipart/form-data"
							document.form.action = 'default.asp?op=cad_noticia&aspid=' + AspId
							document.form.submit();
							//alert(editor.getHTML());
						
					}
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

function stat_noticiaC(n,i,z)
{
	var obj = document.getElementById('sCod_vales').value;
	if(obj == ""){
		alert("É necessário selecionar algum comentário antes de mudar o status");
	}
	else
	{

			document.form.method = "post"
			document.form.action = 'default.asp?op=comentstat&id_s='+n+'&sCod_vale='+z+'&periodo='+i+'&aspid=' + AspId
			document.form.submit();
		
	}
}

function destaq_noticia(n,i)
{
	var nValor = document.getElementById('sCod_vale').value.split(",");
	if(nValor.length!=1)
	{
		alert("Para destacar é necessário apenas selecionar uma noticia");
	}
	else
	{
		if(nValor[0]==""){
			alert("É necessário selecionar algum item");
		}
		else
		{
			//page = 'inc_album_abertos.asp';
			//AjaxLoad(page,'ids='+document.getElementById('fotosel').value+"&op=dest&Gpage="+ActualPage+"&id="+i,'AlbumAberto');
			//document.getElementById('fotosel').value = "";
			document.form.method = "post"
			//document.form.encoding = "multipart/form-data"
			document.form.action = 'default.asp?op=dest&id_s='+n+'&periodo='+i+'&aspid=' + AspId
			document.form.submit();
		}
	}
	/*for(i=0;i<nValor.length;i++){
		
		//tCObj.value = tCObjValue.replace(tCClear[i],"");
	}*/
}

function excluir_noticia(n,i)
{
	var obj = document.getElementById('sCod_vale').value;
	if(obj == ""){
		alert("É necessário selecionar alguma noticia antes de exluir");
	}
	else
	{
		if(confirm("Você realmente quer excluir estas noticias?"))
		{
			//page = 'inc_album_abertos.asp';
			//AjaxLoad(page,'ids='+document.getElementById('fotosel').value+"&op=del&Gpage=1&id="+i,'AlbumAberto');
			//document.getElementById('fotosel').value = "";
			document.form.method = "post"
			//document.form.encoding = "multipart/form-data"
			document.form.action = 'default.asp?op=del&tp=nt&id_s='+n+'&periodo='+i+'&aspid=' + AspId
			document.form.submit();
		}
	}
}

function excluir_noticiaC(n,i,z)
{
	var obj = document.getElementById('sCod_vales').value;
	if(obj == ""){
		alert("É necessário selecionar algum comentário antes de exluir");
	}
	else
	{
		if(confirm("Você realmente quer excluir estes comentários?"))
		{
			//page = 'inc_album_abertos.asp';
			//AjaxLoad(page,'ids='+document.getElementById('fotosel').value+"&op=del&Gpage=1&id="+i,'AlbumAberto');
			//document.getElementById('fotosel').value = "";
			document.form.method = "post"
			//document.form.encoding = "multipart/form-data"
			document.form.action = 'default.asp?op=del&tp=coment&id_s='+n+'&sCod_vale='+z+'&periodo='+i+'&aspid=' + AspId
			document.form.submit();
		}
	}
}

function coment_noticia(n,i)
{
	var nValor = document.getElementById('sCod_vale').value.split(",");
	if(nValor.length!=1)
	{
		alert("Para visualizar os comentários é necessário apenas selecionar uma noticia");
	}
	else
	{
		if(nValor[0]==""){
			alert("É necessário selecionar algum item");
		}
		else
		{
			//page = 'inc_album_abertos.asp';
			//AjaxLoad(page,'ids='+document.getElementById('fotosel').value+"&op=dest&Gpage="+ActualPage+"&id="+i,'AlbumAberto');
			//document.getElementById('fotosel').value = "";
			document.form.method = "post"
			//document.form.encoding = "multipart/form-data"
			document.form.action = 'default.asp?op=coment&id_s='+n+'&periodo='+i+'&aspid=' + AspId
			document.form.submit();
		}
	}
}