// JavaScript Document
function Clear(obj){
	//var resp = t;
	var results;
	var splitstring;
	var objv = obj.value;
	var tstring='';
	
	resp = objv.toLowerCase()
	resp = '' + objv;
	resp = resp.replace("�","a");
	resp = resp.replace("�","a");
	resp = resp.replace("�","a");
	resp = resp.replace("�","a");
	resp = resp.replace("�","a");
	resp = resp.replace("�","a");
	resp = resp.replace("�","a");
	resp = resp.replace("A","a");
	resp = resp.replace("�","a");
	resp = resp.replace("�","a");
	resp = resp.replace("�","a");
	resp = resp.replace("�","e");
	resp = resp.replace("�","e");
	resp = resp.replace("�","e");
	resp = resp.replace("�","e");
	resp = resp.replace("�","e");
	resp = resp.replace("�","e");
	resp = resp.replace("�","e");
	resp = resp.replace("E","e");
	resp = resp.replace("�","e");
	resp = resp.replace("�","i");
	resp = resp.replace("�","i");
	resp = resp.replace("�","i");
	resp = resp.replace("�","i");
	resp = resp.replace("�","i");
	resp = resp.replace("�","i");
	resp = resp.replace("�","i");
	resp = resp.replace("I","i");
	resp = resp.replace("�","i");
	resp = resp.replace("�","o");
	resp = resp.replace("�","o");
	resp = resp.replace("�","o");
	resp = resp.replace("�","o");
	resp = resp.replace("�","o");
	resp = resp.replace("�","o");
	resp = resp.replace("�","o");
	resp = resp.replace("�","o");
	resp = resp.replace("�","o");
	resp = resp.replace("O","o");
	resp = resp.replace("�","o");
	resp = resp.replace("�","u");
	resp = resp.replace("�","u");
	resp = resp.replace("�","u");
	resp = resp.replace("�","u");
	resp = resp.replace("U","u");
	resp = resp.replace("�","u");
	resp = resp.replace("�","u");
	resp = resp.replace("�","u");
	resp = resp.replace("�","u");
	resp = resp.replace("�","c");
	resp = resp.replace("C","c");
	resp = resp.replace("~","");
	resp = resp.replace("^","");
	resp = resp.replace("-","");
	resp = resp.replace("�","");
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
				alert("Digite o nome da �rea");
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
									alert("Digite as dimen��es");
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
							alert("Selecione um arquivo v�lido");
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
