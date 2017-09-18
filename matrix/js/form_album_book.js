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

function miniMenu_album(){
	
			if(document.form.album.value==''){
				alert("Digite o nome do Álbum");
			}else{
				if(document.form.quem.value=='')
				{
					alert("Digite o nome do Evento");
				}else{
					if(document.form.quando.value==''){
						alert("Digite a data do evento");
					}else{
						if(document.form.coment.value==''){
							alert("Digite um comentário");
						}else{
							document.form.method = "post"
							document.form.encoding = "multipart/form-data"
							document.form.action = 'upload.asp?op=album&aspid=' + AspId
							document.form.submit();
						}
					}
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


// -------------------------------------------------------------------
// Adicione todos os arquivo de uma pasta
//  
// -------------------------------------------------------------------
function ShowFolderList(folderspec)
{

   var fso, fsoFile, fsoExt,f, fc, s, xy, xn, Fload;
   var Fname, Ffolder, Fsize, Fext, Fpatch;

	//variaveis para a filtragens de arquivos validos
   var FileN = new Array(); //Arquivos não validos 
   var FileY = new Array(); //Arquivos Validos
   
   //document.getElementById("resultado").innerHTML = '';
   document.getElementById("resum").innerHTML = '';
   document.getElementById("msg").innerHTML = "Aguarde procurando arquivos....";
	
	try{
 	  fso = new ActiveXObject("Scripting.FileSystemObject");
	  fsoFile = fso.GetFile(folderspec);
	}
	catch(e){
		document.getElementById("msg").innerHTML = "<br>[!] O sistema não pode funcionar corretamente pois você deve liberar o nivel de segurança, para este site.<br><br>Solução: No Internet Explorer clique no meunu ferramentas > opçoes da internet > segurança > clique em Sites Confiáveis > Clique no botão Sites > Digite '*.tonazoeira.com.br' e desmarque a opção 'exigir verificação...' em seguida clique em OK.";
	}
  	 
   
   f = fso.GetFolder(fsoFile.ParentFolder);

   fc = new Enumerator(f.Files);
   
   s = "Nenhum arquivo encontrado!<br>[!] Valido somente arquivo "+ValidExt+" com o tamanho maximo de "+MaxSize/1024+" Kb.";
   
   xy = 0;
   xn = 0;
 
    
	
   for (; !fc.atEnd(); fc.moveNext())

   {
   	  //Inicia coletagem de dados dos arquivos
	  fsoFile = fso.GetFile(fc.item());
	  fsoExt = fso.GetExtensionName(fc.item());
	  
	  Fpatch = fc.item(); //camiho completo do arquivo
	  Ffolder = fsoFile.ParentFolder; //pasta onde esta os arquivos
	  Fname = fsoFile.Name; //nome do arquivo
	  Fsize = fsoFile.size; //tamanho do arquiivo em byte
	  Fext = fsoExt.toLowerCase(); //estenção do arquivo
    	
		if(Fext==ValidExt) //filtrando a extenção
		{
			
			if(Fsize <= MaxSize){
				FileY[xy] = Fpatch;
				xy = xy + 1;
				}else{
				FileN[xn] = Fpatch;
				xn = xn + 1;
				}
				
				s = "";
				
		}
	
   }
   	

		

	try
		{
			document.getElementById("FileValid").name;
	
		}
	catch(e)
		{
			
		   if(xy!=0)
				s += "Arquivos Validos<br> <select id='FileValid' style='font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 11px;width: 230px;' name='FileValid' size='10' multiple='multiple'  onchange='resumo(this.value,\"FileValid\");'> </select><br>";
		   if (xn!=0)
				s += "<br><br>Arquivos Inválidos<br> <select id='FileInvalid' style='font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 11px;width: 230px;' name='FileInvalid' size='10' multiple='multiple'  onchange='resumo(this.value,\"FileInvalid\");'> </select><br>";
		 
		 	document.getElementById("resultado").innerHTML = s;
		}
		
		
   
   
  for(VFile in FileY){
   
   		fsoFile = fso.GetFile(FileY[VFile]);
		addOption(document.forms[0]['FileValid'],fsoFile.Name,FileY[VFile],false);
   		
   }
    
	for(NFile in FileN){
   
   		fsoFile = fso.GetFile(FileN[NFile]);
		addOption(document.forms[0]['FileInvalid'],fsoFile.Name,FileN[NFile],false);
   		
   }

	 document.getElementById("msg").innerHTML = '';

   //return(s);

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
