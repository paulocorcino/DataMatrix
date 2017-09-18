// JavaScript Document
function Clear(obj){
	//var resp = t;
	var results;
	var splitstring;
	var objv = obj.value;
	var tstring='';
	
	resp = objv.toLowerCase()
	resp = '' + objv;
	resp = resp.replace("á","a");
	resp = resp.replace("à","a");
	resp = resp.replace("â","a");
	resp = resp.replace("ã","a");
	resp = resp.replace("ä","a");
	resp = resp.replace("Á","a");
	resp = resp.replace("À","a");
	resp = resp.replace("A","a");
	resp = resp.replace("Â","a");
	resp = resp.replace("Ã","a");
	resp = resp.replace("Ä","a");
	resp = resp.replace("é","e");
	resp = resp.replace("è","e");
	resp = resp.replace("ê","e");
	resp = resp.replace("ë","e");
	resp = resp.replace("É","e");
	resp = resp.replace("È","e");
	resp = resp.replace("Ê","e");
	resp = resp.replace("E","e");
	resp = resp.replace("Ë","e");
	resp = resp.replace("í","i");
	resp = resp.replace("ì","i");
	resp = resp.replace("î","i");
	resp = resp.replace("ï","i");
	resp = resp.replace("Í","i");
	resp = resp.replace("Ì","i");
	resp = resp.replace("Î","i");
	resp = resp.replace("I","i");
	resp = resp.replace("Ï","i");
	resp = resp.replace("ó","o");
	resp = resp.replace("ò","o");
	resp = resp.replace("ô","o");
	resp = resp.replace("õ","o");
	resp = resp.replace("ö","o");
	resp = resp.replace("Ó","o");
	resp = resp.replace("Ò","o");
	resp = resp.replace("Ô","o");
	resp = resp.replace("Õ","o");
	resp = resp.replace("O","o");
	resp = resp.replace("Ö","o");
	resp = resp.replace("ú","u");
	resp = resp.replace("ù","u");
	resp = resp.replace("û","u");
	resp = resp.replace("ü","u");
	resp = resp.replace("U","u");
	resp = resp.replace("Ú","u");
	resp = resp.replace("Ù","u");
	resp = resp.replace("Û","u");
	resp = resp.replace("Ü","u");
	resp = resp.replace("ç","c");
	resp = resp.replace("C","c");
	resp = resp.replace("~","");
	resp = resp.replace("^","");
	resp = resp.replace("-","");
	resp = resp.replace("´","");
	resp = resp.replace("`","");
	resp = resp.replace("'","");
	resp = resp.replace("&","");
	resp = resp.replace(",","");
	resp = resp.replace("@","");
	
	splitstring = resp.split(" ");
	for(i = 0; i < splitstring.length; i++)
	tstring += splitstring[i];
	
	resp = tstring;
	obj.value = resp;
	//alert(obj.value);
}

function miniMenu_area(){
	
			if(document.form.area.value==''){
				alert("Digite o nome da Área");
			}else{
				if(document.form.chave.value=='')
				{
					alert("Digite a chave");
				}else{
					if(document.form.tban.value=='')
					{
						alert("Digite o Tempo de troca de imagem");
					}else{
					
						if(document.form.nban.value=='' && document.form.nban.value<1){
							alert("Digite o numero de banners");
						}else{
							if(document.form.wban.value=='' && document.form.hban.value==''){
									alert("Digite as dimenções");
								}else{
									document.form.method = "post"
								//	document.form.encoding = "multipart/form-data"
									document.form.action = 'default.asp?op=cad_area&aspid=' + AspId
									document.form.submit();
							}
						}
					}
				}
			}
}

function miniMenu_banner(){
	
			if(document.form.datae.value==''){
				alert("Digite a Data de Entrada");
			}else{
				if(document.form.datas.value=='')
				{
					alert("Digite a Data de Saida");
				}else{
					
						try{
							var tests = document.form.filebanner.value;
						}
						catch(e){
							ImgValid = true;
						}
					
						if(!ImgValid){
							alert("Selecione um arquivo válido");
						}else{
							if(document.form.cod_area.value=='')
							{
								alert("Cadastre as areas antes");
							}else{
								document.form.method = "post"
								if(document.form.noimg.value != "1")
								{
									document.form.encoding = "multipart/form-data"
									document.form.action = 'upload.asp?aspid=' + AspId
								}else{
									document.form.action = 'default.asp?op=salvar_banner&aspid=' + AspId
								}
								
								document.form.submit();
							}
						}
					}
				}
			
}

function area_abrir(){
	var area_id = document.form.cod_area.value;
	window.open("default.asp?op=list_banner&cod_area="+area_id+"&aspid=" + AspId,"_self");
}
