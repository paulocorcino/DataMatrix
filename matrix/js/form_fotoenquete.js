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

function miniMenu_fotoenq(){
	
			var sendt = false;
			var nerror = 0;
			var updat = 0;
			
			if(document.form.cod_area.value==''){
				alert("Cadastre alguma área antes de criar a enquete.");
				nerror = 1 + nerror;
			}else{
				if(document.form.pergunta.value=='')
				{
					alert("Digite alguma pergunta.");
					nerror = 1 + nerror;
				}else{
					if(document.form.entrada.value=='' || document.form.saida.value==''){
						alert("Digite o periódo de funcionamento da enquete.");
						nerror = 1 + nerror;
					}else{
						if(document.form.dtainvalid.value=='true'){
							alert("O periódo selecionado é inválido.");
							nerror = 1 + nerror;
						}else{
						//teste
							var naltr = document.form.alt.value;
							for(i=1;i<=naltr;i++)
							{
								try
								{
									updat = document.getElementById("nalt" + i).value;
									if(!document.getElementById("nalt" + i).checked)
									{
										//imagem será atualizada
										if(document.getElementById("titulo" + i).value=="")
										{
											alert("O titulo da alternativa " + i + " está vazia.");
											nerror = 1 + nerror;
											
										}else{
											
											if(document.getElementById("img" + i).value == "")
											{
												//erro
												alert("Você deve selecionar uma imagem da alternativa " + i );
												nerror = 1 + nerror;
											}
							
										}
									}
								}
								catch(e)
								{
									//imagem será atualizada
									if(document.getElementById("titulo" + i).value=="")
									{
										alert("O titulo da alternativa " + i + " está vazia.");
										nerror = 1 + nerror;
										
									}else{
										try
										{
												if(document.getElementById("img" + i).value == "")
											{
												//erro
												alert("Você deve selecionar uma imagem da alternativa " + i );
												nerror = 1 + nerror;
											}
										}
										catch(e)
										{
											//
										}
											
									}
								}
							}
						}
					}
				}
			}
			
		if(nerror < 1){
			
			
			document.form.method = "post"
			document.form.encoding = "multipart/form-data"
			document.form.action = 'upload.asp?op=album&aspid=' + AspId
			
			try
			{
				document.form.submit();
			}
			catch(e)
			{
				var urlpage = '?ok';
				urlpage = location + urlpage;				
				var urlpages = urlpage.split("?");
				//alert(urlpages[0]);
				document.form.action = urlpages[0] + '?op=album&aspid=' + AspId
				window.onerror = function() { return (false); }
				setTimeout('document.form.submit()',500);
				window.onerror = function() { return (false); }
			}
			
		}
}

function destaq_foto(n,i)
{
	var nValor = n.split(",");
	if(nValor.length!=1)
	{
		alert("Para destacar é necessário apenas selecionar uma fotografia");
	}
	else
	{
		if(nValor[0]==""){
			alert("É necessário selecionar algum item");
		}
		else
		{
			page = 'inc_album_abertos.asp';
			AjaxLoad(page,'ids='+document.getElementById('fotosel').value+"&op=dest&Gpage="+ActualPage+"&id="+i,'AlbumAberto');
			document.getElementById('fotosel').value = "";
		}
	}
	/*for(i=0;i<nValor.length;i++){
		
		//tCObj.value = tCObjValue.replace(tCClear[i],"");
	}*/
}

function excluir_foto(n,i)
{
	var obj = document.getElementById('fotosel').value;
	if(obj == ""){
		alert("É necessário selecionar alguma foto antes de exluir");
	}
	else
	{
		if(confirm("Você realmente quer excluir estas fotos?"))
		{
			page = 'inc_album_abertos.asp';
			AjaxLoad(page,'ids='+document.getElementById('fotosel').value+"&op=del&Gpage=1&id="+i,'AlbumAberto');
			document.getElementById('fotosel').value = "";
		}
	}
}

function miniMenu_enviar(){
	var imgyes = false
	for(var i=0;i<=9;i++)
	{
		if(document.forms[0].elements[i].value != "")
		{
			imgyes = true
		}
		//alert(document.forms[0].elements[i].value + " " + document.forms[0].elements[i].name)
	}
	if(!imgyes){
		alert("Selecione alguma imagem");
	}else{
		document.form.method = "post"
		document.form.encoding = "multipart/form-data"
		document.form.action = 'upload.asp?op=fotos&aspid=' + AspId
		document.form.submit();
	}
}

//--------------------------------------------------------------------------

var MaxSize = 300; //valor em KB
var ValidExt = "jpg"; //extenção liberada

//Convertendo em byte
MaxSize = MaxSize * 1024;

// -------------------------------------------------------------------
// Seleciona forma de adicionar arquivos
// 
// 
// -------------------------------------------------------------------
function opadd(tp,fln){
	if(fln!=""){
		if(tp==0){
			ShowFolderList(fln);
		}else{
			ShowFile(fln);
		}
	}else{
		alert("Selecione algum arquivo");
		}
}

// -------------------------------------------------------------------
// removeSelectedOptions(select_object)
//  Remove all selected options from a list
//  (Thanks to Gene Ninestein)
// -------------------------------------------------------------------
function removeSelectedOptions(from) { 
	//if (!hasOptions(from)) { return; }
	 document.getElementById("resum").innerHTML = '';
	if (from.type=="select-one") {
		from.options[from.selectedIndex] = null;
		}
	else {
		for (var i=(from.options.length-1); i>=0; i--) { 
			var o=from.options[i]; 
			if (o.selected) { 
				from.options[i] = null; 
				} 
			}
		}
	from.selectedIndex = -1; 
	} 
// -------------------------------------------------------------------
// Resumo
// Add an option to a list
// -------------------------------------------------------------------
	function resumo(file,n)
	{
		  var s, fso, fsoFile, ext, size, modi ;
		  s = '<br><br><table style="font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 11px;border: 1px solid #009999;" width="314" height="193" border="0" cellpadding="5" cellspacing="5"><tr><td valign="top" bgcolor="#00FFCC">'
		  	
			   fso = new ActiveXObject("Scripting.FileSystemObject");
  			   ext = fso.GetExtensionName(file);
			   
			   //se imagem abre o arquivo
			   if(ext=="jpg" || ext=="gif" || ext=="jpeg" || ext=="png" || ext=="bmp")
			   		s += '<center><img src="'+file+'" width=120></center><hr><br>';
				
				fsoFile = fso.GetFile(file);
				
				size = fsoFile.Size/1024;
				if(size<=1024){
					size = size.toFixed(1) +"Kb";
				}else{
					size = size/1024;
					size = size.toFixed(2) +"Mb";
					}
				
				//modi = fsoFile.DateLastModified;
				var modi = new Date(fsoFile.DateLastModified);
				
				s += "<b>Nome do Arquivo</b><br>"+fsoFile.Name;
				s += "<br><br><b>Tamanho</b><br>"+size;
				s += "<br><br><b>Modificado</b><br>"+modi.getDate()+"/"+(parseInt(modi.getMonth())+1)+"/"+modi.getYear()+" "+modi.getHours()+":"+modi.getMinutes()+":"+modi.getSeconds();
				s += "<br><br><b>Tipo do Arquivo</b><br>"+fsoFile.Type;
				s += "<p><input type='button' name='bt_apagar' value='Deletar' onclick='removeSelectedOptions(document.forms[0][\""+n+"\"]);'></p>"
				
		  s += '</td></tr></table>'
		  document.getElementById("resum").innerHTML = s;
	}
// -------------------------------------------------------------------
// addOption(select_object,display_text,value,selected)
//  Add an option to a list
// -------------------------------------------------------------------
function addOption(obj,text,value,selected) {
	if (obj!=null && obj.options!=null) {
		obj.options[obj.options.length] = new Option(text, value, false, selected);
		}
	}
// -------------------------------------------------------------------
// Adicione todos um unico arquivo
//  
// -------------------------------------------------------------------
function ShowFile(folderspec)
{
   var fso, fsoFile, fsoExt,f, fc, s, xy, xn, Fload;
   var Fname, Ffolder, Fsize, Fext, Fpatch;
   
   document.getElementById("resum").innerHTML = '';
   document.getElementById("msg").innerHTML = "Aguarde procurando arquivos....";
   
   fso = new ActiveXObject("Scripting.FileSystemObject");
   fsoFile = fso.GetFile(folderspec);
   
   fsoExt = fso.GetExtensionName(folderspec);
	  
	  Fpatch = folderspec; //camiho completo do arquivo
	  Ffolder = fsoFile.ParentFolder; //pasta onde esta os arquivos
	  Fname = fsoFile.Name; //nome do arquivo
	  Fsize = fsoFile.size; //tamanho do arquiivo em byte
	  Fext = fsoExt.toLowerCase(); //estenção do arquivo
	  
	try
		{
			document.getElementById("FileValid").name;
	
		}
	catch(e)
		{
			
		   	s += "Arquivos Validos<br> <select id='FileValid' style='font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 11px;width: 230px;' name='FileValid' size='10' multiple='multiple'  onchange='resumo(this.value,\"FileValid\");'> </select><br>";
		 
		 	document.getElementById("resultado").innerHTML = s;
		}
		
	
	
	try
		{
			document.getElementById("FileInvalid").name;
	
		}
	catch(e)
		{	   
			s += "<br><br>Arquivos Inválidos<br> <select id='FileInvalid' style='font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 11px;width: 230px;' name='FileInvalid' size='10' multiple='multiple'  onchange='resumo(this.value,\"FileInvalid\");'> </select><br>";
		 
		 	document.getElementById("resultado").innerHTML = s;
		}
		
		
    	
		if(Fext==ValidExt) //filtrando a extenção
		{
			
			if(Fsize <= MaxSize){
			
					fsoFile = fso.GetFile(Fpatch);
					addOption(document.forms[0]['FileValid'],fsoFile.Name,Fpatch,false);
					
				}else{
				
					fsoFile = fso.GetFile(Fpatch);
					addOption(document.forms[0]['FileInvalid'],fsoFile.Name,Fpatch,false);
				
				}
				
				s = "";
				
		}
   
}
//-------------------------------------------------------------------------------------

function validperiodo(dini, dfim, darea, did)
{
	var startfunction = false;
	
	try
	{
		var dtini = dini.value;
		var dtfim = dfim.value;
		var codarea = darea.value;
		var codid = did.value;
		
		if(dtini!=''){
			
			if(dtfim!='')
			{
				if(codarea!='')
				{
					if(codid!='')
					{
						startfunction = true;
					}
					else
					{
						startfunction = false;
					}
				}
				else
				{
					startfunction = false;
				}
			}	
			else
			{
				startfunction = false;
			}
		}
		else
		{
			startfunction = false;
		}
		
	}
	catch(e)
	{
		//erro
		alert("Algum objeto não está presente.");
	}
	
	if(startfunction)
	{
		//inicia a função
		
				var http = null; //variavel
			
			 //--------------------------------------------------------------
			 // Detecta o componentes XML
			 //---------------------------------------------------------------
			
			if (typeof XMLHttpRequest != "undefined") {
				http = new XMLHttpRequest();
			}else if (window.ActiveXObject){
			  var aVersions = ["MSXML2.XMLHttp.5.0",
				"MSXML2.XMLHttp.4.0","MSXML2.XMLHttp.3.0",
				"MSXML2.XMLHttp","Microsoft.XMLHttp"];
		
			  for (var i = 0; i < aVersions.length; i++) {
				try{
					var http = new ActiveXObject(aVersions[i]);
				}
				catch(oError){
					//Do nothing
				}
			  }
			}
			
			  //-----------------------------------------------------------------------
			  // Página do Post
			  //-----------------------------------------------------------------------
			  web = 'check.asp';
				  
			  //------------------------------------------------------------------------
			  // Envio dos Dados
			  //------------------------------------------------------------------------
			  http.open('post', web, true);
			  document.getElementById('priodo').className = '';
			   document.getElementById('priodo').value = 'Analisando...';
			  http.setRequestHeader('Content-Type','application/x-www-form-urlencoded');
			
			  //-----------------------------------------------------------------------
			  // Leitura de Dados
			  //-----------------------------------------------------------------------
			  http.onreadystatechange = function()
			  {
					if(http.readyState == 4){
						var results = unescape(http.responseText.replace(/\+/g,' '));
						//document.write(results);
						if(results=='1'){
									document.form.dtainvalid.value = 'false';
									document.getElementById('priodo').className = 'green';
									document.getElementById('priodo').value = 'Periódo Válido';
						}else{
									document.form.dtainvalid.value = 'true';
									document.getElementById('priodo').className = 'red';
									document.getElementById('priodo').value = 'Periódo Inválido';
						}
					}		 
			  };
			  
			//------------------------------------------------------------------------
			// Captura a Pagina
			//------------------------------------------------------------------------			
			iTimeoutId = setTimeout(function(){http.send('fim='+dtfim+'&incio='+dtini+'&id_a='+codarea+'&id='+codid);},800);
			
		
		
	}else{
		
		document.form.dtainvalid.value = 'true';
		document.getElementById('priodo').className = 'red';
		document.getElementById('priodo').value = 'Periódo Inválido';
	}
	
	
	
}


//--------------------------------------------------------------------------------------
function miniMenu_enviar_ft()
{
	try{
		if(document.form.FileValid.options.length<1){
			alert("Selecione alguma imagem");
		}else{
			//selectAllOptions(document.form.FileValid);
			var obj = document.form.FileValid;
			var imgs = "";
			for (var i=0; i<obj.options.length; i++) { 
      				obj.options[i].selected = true; 
					//alert(obj.options[i].value)
					imgs = imgs + "|" + obj.options[i].value
			} 
			document.getElementById("arqSel").value = imgs;
			//alert(document.getElementById("arqSel").value)
			document.form.method = "post"
			document.form.encoding = "multipart/form-data"
			document.form.action = 'upload.asp?op=fotosbeta&aspid=' + AspId
			document.form.submit();
			
		}
	}
	catch(e)
	{
		alert("Selecione alguma imagem");
	}
}

function incpat(i,x,y)
{
	if(confirm("Você realmente deseja inserir o logo do patrocinado no canto superior da fotografias deste album?\nLembre-se que esta operação não poderá ser desfeita, certifique-se que já tenha testado o visual da imagem nesta posição.")){
		WindowOpen('default.asp?op=inc_pat&id='+i+'&id_ar='+x+'&aspid=' + y)
	}
}

//---------------------------------------------------------------
// Cria Lista de objectos da página 
//---------------------------------------------------------------

function listaUpload(obj){
	
	var stratfunction = false;
	var htmlt = '';
	try
	{
		var nitens = obj.value;
		stratfunction = true;
	}
	catch(e)
	{
		alert("Objeto inválido");
	}
	
	if(stratfunction){
		for(i=1;i <= obj.value;i++)
		{
			var resptitulo = '';
			var respvotos  = '0';
			var respfoto   = '';
			var respid     = '';
					
			//Caso tenha as respostas colocar em seus campos
			try
			{
				var respitem = resp.split(";");
				//for(y=0;y < respitem.length;y++)
				//{
					try
					{
						var respitemop = respitem[obj.value - i].split(",");
						if(respitemop[3]!=undefined){
							resptitulo = respitemop[3];
							respvotos  = respitemop[2] - 1;
							respfoto   = "<img src='"+respitemop[1]+"'>";
							respid     = respitemop[0];
							enquete_tp = 'upt';
						}
						else
						{
							respfoto = '';
							enquete_tp = 'new';
						}
					}
					catch(e)
					{
						respfoto = '';
						enquete_tp = 'new';						
					}
					
				//}
			}
			catch(e)
			{
				//nada a fazer
				/*respfoto = '';
				enquete_tp = 'new';*/
			}
			
			    htmlt = htmlt + '<table width="50%" border="0" cellspacing="0" cellpadding="3">'
				htmlt = htmlt +  '	<tr><td width="56%">'
				htmlt = htmlt +    '	<span class="texto_size_10">Alternativa #'+i+'</span><br />'
              	htmlt = htmlt + 	'	<hr size="1" noshade="noshade" />'
           		htmlt = htmlt + 	'	<span class="texto_size_10">Titulo</span><br />'
            	htmlt = htmlt + 	'	<input name="titulo'+i+'" type="text" id="titulo'+i+'" size="40" maxlength="255" value="'+resptitulo+'" />'
           		htmlt = htmlt + 	'<br />'
				
				if(enquete_tp!='new'){
					/*htmlt = htmlt + 	'	<input name="nalt'+i+'" type="checkbox" id="nalt'+i+'" value="1" checked="checked" onclick="disabledImg(this,document.form.img'+i+')" />&nbsp;'
					htmlt = htmlt + 	'	<span class="texto_size_10">N&atilde;o Alterar imagem</span>'
					htmlt = htmlt + 	'<br />'*/
				}else{
				
            	htmlt = htmlt + 	'	<span class="texto_size_10">Imagem</span>&nbsp;<br />'
				htmlt = htmlt + 	'	<input name="img'+i+'" type="file" id="img'+i+'" size="40" onChange="previewImg(this,\'imgs'+i+'\',\'jpeg,jpg,png,gif,bmp\',\'0,90\')"'
				
					if(enquete_tp=='new')
					{
						htmlt = htmlt + '';
					}
					else
					{
						htmlt = htmlt + ' disabled="true" ';
					}
				
					htmlt = htmlt + ' />'
				}
				htmlt = htmlt + '	<span class="texto_size_10"><br />'
				htmlt = htmlt + 	'	Votos Iniciais</span><br />'
				htmlt = htmlt + 		'<input name="votos'+i+'" type="text" id="votos'+i+'" size="3" maxlength="3" onBlur="if(this.value==\'\'){this.value==\'0\';}" value="'+respvotos+'" /></td>'
				htmlt = htmlt + 	'<td width="44%" align="center" valign="middle"><div id="imgs'+i+'">'+respfoto+'</div></td></tr></table>'
     			htmlt = htmlt + ' <br />'
		}
		
		document.getElementById("divalt").innerHTML = htmlt;
		
		
	
		
	}
}

function disabledImg(iten,obj)
{
	if(iten.checked)
	{
		obj.disabled = true;
	}else{
		obj.disabled = false;
	}
}