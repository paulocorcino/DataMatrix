// JavaScript Document

function ispocket(){
	if(screen.width<640)
		return true;
	else
		return false;
}
//---------------------------------------------------------------------------
function selectAllOptions(obj) { 
   for (var i=0; i<obj.options.length; i++) { 
      obj.options[i].selected = true; 
   } 
} 
//---------------------------------------------------------------------------
function isEmail(n){
	var email = n;
	var emailsufix;
	
	//valida se contem o arroba
	if(email.lastIndexOf("@") > 0)
	{
		//verifica se o primeiro caracter � um @
		if(email.slice(0,1)!="@")
		{
			   //verfica se existe algo ap�s o @
			   emailsufix = email.slice(email.lastIndexOf("@")+1,email.length); //dominio do e-mail
			   if(emailsufix.length > 0)
			   {
				   //verifica a existencia de um . para deduzir que foi digitado o dominio
				   if(emailsufix.lastIndexOf(".") > 0)
				   {
					   email = true;
					}
					else
					{
						email = false;
					}
			   }
			   else
			   {
				   	email = false;
			   }
			   
		}
		else
		{
			email = false;
		}
	}
	else
	{
		email = false;
	}
		
	return email;
}
//---------------------------------------------------------------------------

function isIE(){
	if(navigator.appName=='Microsoft Internet Explorer')
		return true;
	else
		return false;
}
//---------------------------------------------------------------------------
function subTitle(n){
	try{
		document.getElementById("matrix_titulo").innerHTML = document.getElementById("matrix_titulo").innerHTML + " " + n;
	}
	catch(e)
	{
		alert("O objeto matrix_titulo_sub n�o existe contate o administrador do site");
	}
}
//---------------------------------------------------------------------------
function CarregaFrames(){

	var html = "";
	var Fdim = 0;
	
	if(ispocket())
		Fdim = 25;
	else
		Fdim = 40;	
		
		 html = html + '<frameset rows="'+Fdim+',100%" cols="*" framespacing="1" frameborder="no" border="1" bordercolor="#FFFFFF">';
		 html = html + '<frame src="topo.asp" frameborder="No" scrolling="No" noresize="noresize" id="topo" />';
		 html = html + '<frameset rows="21,100%" cols="*" framespacing="1" frameborder="no" border="1" bordercolor="#FFFFFF">';
		 html = html + '<frame src="about:blank" frameborder="No" scrolling="No" noresize="noresize" id="_nome" name="_nome" />';
		 html = html + '<frame src="login.asp" frameborder="no" scrolling="auto" id="_centro" />';
		 html = html + '</frameset>';
		 html = html + '</frameset>';
		 
	return html;
}
//---------------------------------------------------------------------------
function WindowOpen(url){
	if(ispocket())
		window.location.href = url;
	else
		window.open(url,"_self");
}
//-------------------------------------------------------------------------------

function miniMenu_dup(){ //menu incluir pedido

			if(document.form.nome_empresa.value==''){
				alert("Digite o nome da empresa");
			}else{
				document.form.action = 'default.asp?op=cad&aspid=' + AspId + '&m=' + ModId
				document.form.submit();
			}
}

//-----------------------------------------------------------------------------------

function miniMenu_mod(){ //menu incluir pedido

			if(document.form.nome_modulo.value==''){
				alert("Digite o nome do m�dulo");
			}else{
				if(document.form.url_modulo.value==''){
					alert("Digite a url do m�dulo");
				}else{
					if(document.form.ico_modulo.value==''){
						alert("Digite o nome do icone do m�dulo");
					}else{
						document.form.action = 'default.asp?op=cad&aspid=' + AspId + '&m=' + ModId
						document.form.submit();
					}
				}
			}
}
//-------------------------------------------------------------------------------------
function miniMenu_user(){ //menu incluir pedido
	var nEmpresas;
	var nModulos;
	var nEChecked = false;
	var nMChecked = false;
	
			try
			{
				
				if(document.forms[0].nEmpresa.length != undefined)
				{
					nEmpresas = document.forms[0].nEmpresa.length;
				}
				else
				{
					nEmpresas = 1;
				}
				
								
			}
			catch(e)
			{
				nEmpresas = 0;
			}
			
			try
			{
				
				if(document.forms[0].nModulo.length != undefined)
				{
					nModulos = document.forms[0].nModulo.length;
				}
				else
				{
					nModulos = 1;
				}
				
			}
			catch(e)
			{
				nModulos = 0;
			}


		if(document.form.nome_user.value==''){
				alert("Digite o nome do Usu�rio");
			}else{
				if(document.form.email_user.value==''){
					alert("Digite o E-Mail");
				}else{
					if(document.form.user_user.value==''){
						alert("Digite o nome do usu�rio do login.");
					}else{
						if(document.form.pass_user.value==''){
							alert("Digite a senha do usu�rio!");
						}else{
							if(nModulos > 0 && nEmpresas > 0){
								for(i=0;i<=nEmpresas - 1;i++)
								{
									if(document.forms[0].nEmpresa[i].checked)
									{
										nEChecked = true;
									}
								}
								
								for(i=0;i<=nModulos - 1;i++)
								{
									if(document.forms[0].nModulo[i].checked)
									{
										nMChecked = true;
									}
								}
								
								if(nMChecked)
								{
									if(nEChecked){
										document.form.action = 'default.asp?op=cad&aspid=' + AspId + '&m=' + ModId
										document.form.submit();
									}
									else
									{
										alert("Selecione uma empresa!");	
									}
								}
								else
								{
									alert("Selecione um m�dulo!");
								}
							}
							else
							{
								alert("N�o � poss�vel cadastrar o usu�rio sem cadastrar empresas e os modulos!");
							}
						}
					}
				}
			}
}
//-----------------------------------------------------

function miniMenu_estab(){ //menu incluir pedido

			if(document.form.nome_vale_estab.value==''){
				alert("Digite o nome do Estabelecimento");
			}else{
				document.form.action = 'default.asp?op=cad&aspid=' + AspId + '&m=' + ModId
				document.form.submit();
			}
}
//--------------------------------------------------------
function miniMenu_login(){ //menu incluir pedido

			if(document.form.user_user.value==''){
				alert("Digite o nome do Usu�rio");
			}else{
				if(document.form.user_user.value==''){
					alert("Digite sua Senha.");
				}else{
					
					document.form.action = 'login.asp?op=valid'
					document.form.submit();
				
				}
			}
}
//---------------------------------------------------------
function toplogin(){
	if(ispocket()){
		abrir("http://" + location.host + "/intranet/inc/inc_nome.asp?aspid=" + AspId);
	}else{
		abrir("inc/inc_nome.asp?aspid=" + AspId);
	}
	
}
function abrir(url){
	if(ispocket())
		parent._nome.location.href = url;
	else
		window.open(url,"_nome");
}//---------------------------------------------------------------------------
function User_name(name){
		var nome = "";
		
		if(ispocket()){
			nome = name.split(' ');
			return nome[0];
		}
		else{
			nome = name;
			return nome;
		}
			
		
}
//---------------------------------------------------------------------------
function MsgDia(){
	
	var hora = new Date();
	var Msg = "";
	
	if(hora.getHours()>=18)
		Msg = "Boa Noite";
	else if(hora.getHours()>=12)
		Msg = "Boa Tarde";
	else
		Msg = "Bom Dia";
		
	return Msg;
}
//----------------------------------------------------------------------------

	function OpGerarSel(v)
	{
		
		if(v==0)
		{
			document.form.qtsPrint.value = '';
			document.form.qtsPrint.disabled = true;
		}
		else
		{
			document.form.qtsPrint.disabled = false;
			document.form.qtsPrint.value = 0;
		}
	}

//----------------------------------------------------------------------------
var lojSend = false;
function miniMenu_loj(){ //menu incluir pedido

			if(document.form.cod_loja.value==''){
				alert("Selecione uma categoria");
			}else{
				if(document.form.nome_loja.value==''){
					alert("Digite o nome da Loja");
				}else{
					if(document.form.diretor_loja.value==''){
						alert("Digite o nome do diretor");
					}else{
						if(document.form.end_loja.value==''){
							alert("Digite o endere�o da loja");
						}else{
							if(document.form.cidade_loja.value==''){
								alert("Digite o nome da cidade");
							}else{
								if(document.form.www_loja.value==''){
									alert("Digite o dominio da Loja");
								}else{

									lojSend = true;	
									
									//------------------------------------------------
									if(document.form.email_loja.value!=''){
										
										if(!isEmail(document.form.email_loja.value)){
											alert("Voc� digitou um e-mail inv�lido!");
											lojSend = false;	
										}else{
											lojSend = true;	
										}
										
									}
									//-------------------------------------------------
									if(document.form.foto_loja.value!=''){
										if(ImgValid){
											lojSend = true;		
										}else{
											lojSend = false;
											alert("A imagem que voc� selecionou n�o � valida");
										}
									}else{
										lojSend = true;	
									}
									//--------------------------------------------------
								}
							}
						}
					}
				}
			}
			
	if(lojSend){
		document.form.method = "post"
		document.form.action = 'upload.asp?op=cad_loj&aspid=' + AspId
		document.form.submit();
	}
				

}
//--------------------------------------------------------------------------------

function miniMenu_cat(){ //menu incluir pedido

			if(document.form.nome_cat.value==''){
				alert("Digite o nome da categoria");
			}else{
				document.form.method = "post"
				document.form.action = 'default.asp?op=cad_cat&aspid=' + AspId
				document.form.submit();
			}
}

//-------------------------------------------------------------------------------
var ImgValid = false;
function previewImg(f,l,ext,dim){

	var pImageObj     = f; //objeto form
	var pImageFile    = pImageObj.value.substr(pImageObj.value.lastIndexOf("\\")+1,pImageObj.value.length); //Nome do arquivo
	var pImageFileExt = pImageFile.substr(pImageFile.lastIndexOf(".")+1,pImageFile.length);  //exten��o do arquivo
	var pImageLocal   = l; //Local da preview
	var pImageExt     = ext; //exten��o valida
	var pImageExtOk   = false;
	var pImageError   = false;
	var pImageDim	  = dim;
	var pImageHeight  = "";
	var pImageWidth   = "";
	
	if(pImageDim!="" && pImageDim!=null){ //este campo n�o � obrigat�rio
		if(pImageDim.lastIndexOf(",")!=-1){
			var pImaged = pImageDim.split(",");
			for(i=0;i<pImaged.length;i++)
			{
				if(i==0)
				{
					//largura
					if(pImaged[i]!="0" && pImaged[i]!=0 && pImaged[i]!="0%"){
						pImageWidth = pImaged[i];
					}
					
				}else{
					if(i==1)
					{
						//altura
						if(pImaged[i]!="0" && pImaged[i]!=0 && pImaged[i]!="0%"){
							pImageHeight = pImaged[i];
						}
					}
				}
			}
		}else{
			pImageError = true;
			alert("Voc� deve definir duas dimens�es, caso deseja expressar apenas uma dimens�o coloque o valor 0 na dimens�o que deseja anular");
		}
	}
	
	
	if(pImageExt.lastIndexOf(",")!=-1){
		var pImagext = pImageExt.split(",");
		for(i=0;i<pImagext.length;i++){
			if(pImageFileExt==pImagext[i]){
				pImageExtOk = true;
			}
		}
	}else{
		if(pImageFileExt==pImageExt){
			pImageExtOk = true;
		}
	}
	
	if(!pImageExtOk){
		pImageError = true;
		alert("Imagem tem formato inv�lido");
		ImgValid = false;
	}
	

	try
	{
		if(pImageError){
			document.getElementById(pImageLocal).innerHTML = ""
		}else{
			document.getElementById(pImageLocal).innerHTML = "Carregando..."
		}
		
	}
	catch(e)
	{
		pImageError = true;
		alert("A area de exibi��o da imagem n�o existe");
	}
	
	
	
	if(!pImageError){
		//Executa a imagem
		var pImageShow;
		pImageShow = "<img ";
		 if (pImageHeight!="")
		 {
			pImageShow = pImageShow + " height='"+ pImageHeight +"' ";
		 }
		 
		 if (pImageWidth!="")
		 {
			pImageShow = pImageShow + " Width='" + pImageWidth + "' ";
		 }
		 
		 pImageShow = pImageShow + " src='" + pImageObj.value + "' >";
		 
		 document.getElementById(pImageLocal).innerHTML = pImageShow;
		 
		ImgValid = true;
	}
	
}

//-------------------------------------------------------------------
function thisChar(n,chr){
	
	var tCChar     = chr;
	var tCObj      = n;
	var tCObjValue = tCObj.value;
	
	var tCClear = tCChar.split(",");
	
		for(i=0;i<tCClear.length;i++){
			tCObj.value = tCObjValue.replace(tCClear[i],"");
		}
	
}
//-------------------------------------------------------------------------
function CharMax(obj,n){
	var cmObj = obj;
	var cmN   = n;
	var cmObjLen = cmObj.value.length;
	
	if(cmObjLen>cmN){
		cmObj.value = cmObj.value.substring(0,cmN)
		alert("Voc� atingiu o limite de "+ cmN +" caracters!");
	}	
	
}
//--------------------------------------------------------------------------
function miniMenu_prod(){ //menu incluir pedido

			if(document.form.nome_produto.value==''){
				alert("Digite o nome do Produto");
			}else{
				if(document.form.preco_produto.value==''){
					alert("Digite o seu valor");
				}else{
					if(document.form.desc_produto.value==''){
						alert("Digite a descri��o do produto.");
					}else{
					
						if(document.form.cod_loja.value==''){
							alert("O c�digo da loja n�o foi encontrado favor recarregue o sistema");
						}else{
							
							//---------------------------------------------------------
									if(document.form.foto_produto.value!=''){
										if(ImgValid){
											lojSend = true;		
										}else{
											lojSend = false;
											alert("A imagem que voc� selecionou n�o � valida");
										}
									}else{
										lojSend = true;	
									}
							//-----------------------------------------------------
						}
					}
				}
			}
			
	if(lojSend){
		document.form.method = "post"
		document.form.action = 'upload.asp?op=cad_prod&aspid=' + AspId
		document.form.submit();
	}
}

//--------------------------------------------------------------------------------------------------
var nObj = "";

function nList(n,s)
{
	var nI;
	
	if(nObj != "") 
	{
		if(nObj.indexOf(n)<0)
		{
			nObj =  nObj + "," + n;
		}
		else
		{
			if(nObj.indexOf(n)==0){
				
				nI = n + ",";
				nObj = nObj.replace(nI,"");
				nObj = nObj.replace(n,"");
				
			}
			else
			{
				nI = "," + n;
				nObj = nObj.replace(nI,"");
			}
		}
	}else{
		nObj = n;
	}
		
	s.value = nObj;
		
}
//--------------------------------------------------------------------------------------------------
	function incV(n, obj){
		var Vobj   = obj;
		var ValInc = Vobj.value;
		
			if(ValInc.indexOf(n)>0){
				Vobj.value = Vobj.value.replace(","+n,"");
			}else{
			
				if(ValInc.indexOf(n)==0){
					Vobj.value = Vobj.value.replace(n,"");
					if(Vobj.value.slice(0,1)==","){
						Vobj.value = Vobj.value.slice(1,Vobj.value.length)
					}
				}else{
					if(ValInc==''){
						Vobj.value = n;
					}else{
						Vobj.value = ValInc +","+ n;
					}
				}
				
			}
			
		
		}
//--------------------------------------------------------------------------------------------------------
function AlbumPage(n,i){
	page = 'inc_album_abertos.asp?Gpage=' + n + '&id=' + i;
	AjaxLoad(page,'sel='+document.getElementById('fotosel').value,'AlbumAberto');
}
//------------------------------------------------------------------------- >
var ActualPage;
function ActualP(n){
	ActualPage = n;
}
//------------------------------------------------------------------------- >
function AjaxLoad(pg,query,camada)
{
	//--------------------------------
	// Ajax - Paulo Corcino - 1.5
	//--------------------------------
	
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
	  // P�gina do Post
	  //-----------------------------------------------------------------------
	  web = pg;
	  	  
	  //------------------------------------------------------------------------
	  // Envio dos Dados
	  //------------------------------------------------------------------------
  	  http.open('post', web, true);
	  //document.getElementById(camada).innerHTML = '<div aling=left class=fonte_10_b>Carregando....</div>';
	  http.setRequestHeader('Content-Type','application/x-www-form-urlencoded');
	
	  //-----------------------------------------------------------------------
	  // Leitura de Dados
	  //-----------------------------------------------------------------------
	  http.onreadystatechange = function()
	  {
			if(http.readyState == 4){
			//	if (http.status == 200) {					
					document.getElementById(camada).innerHTML = unescape(http.responseText.replace(/\+/g,' '));
				//}else{
				//	document.getElementById(camada).innerHTML = "<div aling=left class=fonte_10_b>Erro na abertura da p�gina: " + http.status + " - " + http.statusText + "</div>";
				//}
			}		 
	  };
	  
	//------------------------------------------------------------------------
	// Captura a Pagina
	//------------------------------------------------------------------------

	
	if(query!=''){  
		iTimeoutId = setTimeout(function(){http.send(query);},800);
	}else{
		//alert("aqui")
		if(document.getElementById('fotosel').value != "")
		{
		//ocasi�o especial
			iTimeoutId = setTimeout(function(){http.send('sel=' + document.getElementById('fotosel').value);},800);
			
		}else{
			
			iTimeoutId = setTimeout(function() {http.send(null);},800);
			
		}
	}
	  
}
		
//--------------------------------------------------------------------------------------------------------

function miniMenu_del(i,obj){	
	if(obj.value!=''){
		if(confirm("Voc� realmente deseja excluir estes itens selecionados?")){
			document.form.method = "post"
			document.form.action = 'default.asp?op=del&tp='+i+'&aspid=' + AspId
			document.form.submit();
		}
	}else{
		alert("Favor selecione algum item!");
	}
}
//-----------------------------------------------------------------------------------------------------------
function LayerImg(l,img){
		document.getElementById(l).innerHTML = "<img src='" + img + "'  width='150'>";
}
//------------------------------------------------------------------------------------------------------------
	function incV(n, obj){
	
			var Vobj   = document.getElementById(obj);
			var ValInc = Vobj.value;

		
			if(ValInc.indexOf(n)>0){
				Vobj.value = Vobj.value.replace(","+n,"");
			}else{
			
				if(ValInc.indexOf(n)==0){
					Vobj.value = Vobj.value.replace(n,"");
					if(Vobj.value.slice(0,1)==","){
						Vobj.value = Vobj.value.slice(1,Vobj.value.length)
					}
				}else{
					if(ValInc==''){
						Vobj.value = n;
					}else{
						Vobj.value = ValInc +","+ n;
					}
				}
				
			}
		
		}
		
//--------------------------------------------------------------------------------------
	function setCookie(name, value, expires, path, domain, secure) {
  var curCookie = name + "=" + escape(value) +
      ((expires) ? "; expires=" + expires.toGMTString() : "") +
      ((path) ? "; path=" + path : "") +
      ((domain) ? "; domain=" + domain : "") +
      ((secure) ? "; secure" : "");
  document.cookie = curCookie;
}

function getCookie(name) {
  var dc = document.cookie;
  var prefix = name + "=";
  var begin = dc.indexOf("; " + prefix);
  if (begin == -1) {
    begin = dc.indexOf(prefix);
    if (begin != 0) return null;
  } else
    begin += 2;
  var end = document.cookie.indexOf(";", begin);
  if (end == -1)
    end = dc.length;
  return unescape(dc.substring(begin + prefix.length, end));
}

function deleteCookie(name, path, domain) {
  if (getCookie(name)) {
    document.cookie = name + "=" +
    ((path) ? "; path=" + path : "") +
    ((domain) ? "; domain=" + domain : "") +
    "; expires=Thu, 01-Jan-70 00:00:01 GMT";
  }
}


function fixDate(date) {
  var base = new Date(0);
  var skew = base.getTime();
  if (skew > 0)
    date.setTime(date.getTime() - skew);
}

//-----------------------------------------------------------------------------------------

function allCampQuery(n)
{
	var nForms = document.forms.length - 1;
	var nForm;
	for(i=0; i<=nForms; i++){
		if(document.forms[i].name == n){
			nForm = i;
		}
	}

	var nFormElements = document.forms[nForm].elements.length - 1;
	var nFormElement = document.forms[nForm].elements;
	var nQuery='';
	
	for(i=0;i<=nFormElements;i++)
	{
		//analiza se o elemento � um checkbox ou um radio
	
		if(nFormElement[i].name.slice(0,3) != "chk" && nFormElement[i].name.slice(0,3) != "rdo")
		{
			//campos comuns
			if(i!=0)
				nQuery = nQuery + "&"+nFormElement[i].name+"="+escape(nFormElement[i].value);
			else
				nQuery = nFormElement[i].name+"="+escape(nFormElement[i].value);
		}
		else
		{
			//check box ou radio
			if(nFormElement[i].checked)
			{
				if(i!=0)
					nQuery = nQuery + "&"+nFormElement[i].name+"="+escape(nFormElement[i].value);
				else
					nQuery = nFormElement[i].name+"="+escape(nFormElement[i].value);
			}
		}
	
	
	}
	//alert(nQuery.substring(0,1));
	if(nQuery.substring(0,1)=="&")
	{
		//alert(nQuery.substring(1,nQuery.length));
		nQuery = nQuery.substring(1,nQuery.length);
	}
	
	return nQuery;
}
//---------------------------------------------------------
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}

///---------------------------------------------------------------------
//nocharacter special no space
function charclear(obj){

	var resp = obj.value;
	var results;
	var splitstring;
	var tstring='';
	
	resp = resp.toLowerCase()
	resp = '' + resp;
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
	resp = resp.replace(" ","_");
	
	splitstring = resp.split(" ");
	for(i = 0; i < splitstring.length; i++)
	tstring += splitstring[i];
	
	obj.value = tstring;
}
//-----------------------------------------------------------------
// JavaScript Document
/**
 * @author M�rcio d'�vila
 * @version 1.01, 2004
 *
 * PROT�TIPOS:
 * m�todo String.lpad(int pSize, char pCharPad)
 * m�todo String.trim()
 *
 * String unformatNumber(String pNum)
 * String formatCpfCnpj(String pCpfCnpj, boolean pUseSepar, boolean pIsCnpj)
 * String dvCpfCnpj(String pEfetivo, boolean pIsCnpj)
 * boolean isCpf(String pCpf)
 * boolean isCnpj(String pCnpj)
 * boolean isCpfCnpj(String pCpfCnpj)
 */


var NUM_DIGITOS_CPF  = 11;
var NUM_DIGITOS_CNPJ = 14;
var NUM_DGT_CNPJ_BASE = 8;


/**
 * Adiciona m�todo lpad() � classe String.
 * Preenche a String � esquerda com o caractere fornecido,
 * at� que ela atinja o tamanho especificado.
 */
String.prototype.lpad = function(pSize, pCharPad)
{
	var str = this;
	var dif = pSize - str.length;
	var ch = String(pCharPad).charAt(0);
	for (; dif>0; dif--) str = ch + str;
	return (str);
} //String.lpad


/**
 * Adiciona m�todo trim() � classe String.
 * Elimina brancos no in�cio e fim da String.
 */
String.prototype.trim = function()
{
	return this.replace(/^\s*/, "").replace(/\s*$/, "");
} //String.trim


/**
 * Elimina caracteres de formata��o e zeros � esquerda da string
 * de n�mero fornecida.
 * @param String pNum
 *      String de n�mero fornecida para ser desformatada.
 * @return String de n�mero desformatada.
 */
function unformatNumber(pNum)
{
	return String(pNum).replace(/\D/g, "").replace(/^0+/, "");
} //unformatNumber


/**
 * Formata a string fornecida como CNPJ ou CPF, adicionando zeros
 * � esquerda se necess�rio e caracteres separadores, conforme solicitado.
 * @param String pCpfCnpj
 *      String fornecida para ser formatada.
 * @param boolean pUseSepar
 *      Indica se devem ser usados caracteres separadores (. - /).
 * @param boolean pIsCnpj
 *      Indica se a string fornecida � um CNPJ.
 *      Caso contr�rio, � CPF. Default = false (CPF).
 * @return String de CPF ou CNPJ devidamente formatada.
 */
function formatCpfCnpj(pCpfCnpj, pUseSepar, pIsCnpj)
{
	if (pIsCnpj==null) pIsCnpj = false;
	if (pUseSepar==null) pUseSepar = true;
	var maxDigitos = pIsCnpj? NUM_DIGITOS_CNPJ: NUM_DIGITOS_CPF;
	var numero = unformatNumber(pCpfCnpj);

	numero = numero.lpad(maxDigitos, '0');
	if (!pUseSepar) return numero;

	if (pIsCnpj)
	{
		reCnpj = /(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/;
		numero = numero.replace(reCnpj, "$1.$2.$3/$4-$5");
	}
	else
	{
		reCpf  = /(\d{3})(\d{3})(\d{3})(\d{2})$/;
		numero = numero.replace(reCpf, "$1.$2.$3-$4");
	}
	return numero;
} //formatCpfCnpj


/**
 * Calcula os 2 d�gitos verificadores para o n�mero-efetivo pEfetivo de
 * CNPJ (12 d�gitos) ou CPF (9 d�gitos) fornecido. pIsCnpj � booleano e
 * informa se o n�mero-efetivo fornecido � CNPJ (default = false).
 * @param String pEfetivo
 *      String do n�mero-efetivo (SEM d�gitos verificadores) de CNPJ ou CPF.
 * @param boolean pIsCnpj
 *      Indica se a string fornecida � de um CNPJ.
 *      Caso contr�rio, � CPF. Default = false (CPF).
 * @return String com os dois d�gitos verificadores.
 */
function dvCpfCnpj(pEfetivo, pIsCnpj)
{
	if (pIsCnpj==null) pIsCnpj = false;
	var i, j, k, soma, dv;
	var cicloPeso = pIsCnpj? NUM_DGT_CNPJ_BASE: NUM_DIGITOS_CPF;
	var maxDigitos = pIsCnpj? NUM_DIGITOS_CNPJ: NUM_DIGITOS_CPF;
	var calculado = formatCpfCnpj(pEfetivo, false, pIsCnpj);
	calculado = calculado.substring(2, maxDigitos);
	var result = "";

	for (j = 1; j <= 2; j++)
	{
		k = 2;
		soma = 0;
		for (i = calculado.length-1; i >= 0; i--)
		{
			soma += (calculado.charAt(i) - '0') * k;
			k = (k-1) % cicloPeso + 2;
		}
		dv = 11 - soma % 11;
		if (dv > 9) dv = 0;
		calculado += dv;
		result += dv
	}

	return result;
} //dvCpfCnpj


/**
 * Testa se a String pCpf fornecida � um CPF v�lido.
 * Qualquer formata��o que n�o seja algarismos � desconsiderada.
 * @param String pCpf
 *      String fornecida para ser testada.
 * @return <code>true</code> se a String fornecida for um CPF v�lido.
 */
function isCpf(pCpf)
{
	var numero = formatCpfCnpj(pCpf, false, false);
	var base = numero.substring(0, numero.length - 2);
	var digitos = dvCpfCnpj(base, false);
	var algUnico, i;

	// Valida d�gitos verificadores
	if (numero != base + digitos) return false;

	/* N�o ser�o considerados v�lidos os seguintes CPF:
	 * 000.000.000-00, 111.111.111-11, 222.222.222-22, 333.333.333-33, 444.444.444-44,
	 * 555.555.555-55, 666.666.666-66, 777.777.777-77, 888.888.888-88, 999.999.999-99.
	 */
	algUnico = true;
	for (i=1; i<NUM_DIGITOS_CPF; i++)
	{
		algUnico = algUnico && (numero.charAt(i-1) == numero.charAt(i));
	}
	return (!algUnico);
} //isCpf


/**
 * Testa se a String pCnpj fornecida � um CNPJ v�lido.
 * Qualquer formata��o que n�o seja algarismos � desconsiderada.
 * @param String pCnpj
 *      String fornecida para ser testada.
 * @return <code>true</code> se a String fornecida for um CNPJ v�lido.
 */
function isCnpj(pCnpj)
{
	var numero = formatCpfCnpj(pCnpj, false, true);
	var base = numero.substring(0, NUM_DGT_CNPJ_BASE);
	var ordem = numero.substring(NUM_DGT_CNPJ_BASE, 12);
	var digitos = dvCpfCnpj(base + ordem, true);
	var algUnico;

	// Valida d�gitos verificadores
	if (numero != base + ordem + digitos) return false;

	/* N�o ser�o considerados v�lidos os CNPJ com os seguintes n�meros B�SICOS:
	 * 11.111.111, 22.222.222, 33.333.333, 44.444.444, 55.555.555,
	 * 66.666.666, 77.777.777, 88.888.888, 99.999.999.
	 */
	algUnico = numero.charAt(0) != '0';
	for (i=1; i<NUM_DGT_CNPJ_BASE; i++)
	{
		algUnico = algUnico && (numero.charAt(i-1) == numero.charAt(i));
	}
	if (algUnico) return false;

	/* N�o ser� considerado v�lido CNPJ com n�mero de ORDEM igual a 0000.
	 * N�o ser� considerado v�lido CNPJ com n�mero de ORDEM maior do que 0300
	 * e com as tr�s primeiras posi��es do n�mero B�SICO com 000 (zeros).
	 * Esta cr�tica n�o ser� feita quando o no B�SICO do CNPJ for igual a 00.000.000.
	 */
	if (ordem == "0000") return false;
	return (base == "00000000"
		|| parseInt(ordem, 10) <= 300 || base.substring(0, 3) != "000");
} //isCnpj


/**
 * Testa se a String pCpfCnpj fornecida � um CPF ou CNPJ v�lido.
 * Se a String tiver uma quantidade de d�gitos igual ou inferior
 * a 11, valida como CPF. Se for maior que 11, valida como CNPJ.
 * Qualquer formata��o que n�o seja algarismos � desconsiderada.
 * @param String pCpfCnpj
 *      String fornecida para ser testada.
 * @return <code>true</code> se a String fornecida for um CPF ou CNPJ v�lido.
 */
function isCpfCnpj(pCpfCnpj)
{
	var numero = pCpfCnpj.replace(/\D/g, "");
	if (numero.length > NUM_DIGITOS_CPF)
		return isCnpj(pCpfCnpj)
	else
		return isCpf(pCpfCnpj);
} //isCpfCnpj

//-------------------------------------------------------
function CharMax(obj,n){
	var cmObj = obj;
	var cmN   = n;
	var cmObjLen = cmObj.value.length;
	
	if(cmObjLen>cmN){
		cmObj.value = cmObj.value.substring(0,cmN)
		alert("Voc� atingiu o limite de "+ cmN +" caracters!");
	}	
	
}
//------------------------------------------------------------
function chkDefault(v,obj,disabled){
		var objchk = obj
		var objvalid = false;
		var objnvalid = 0;
		var objvaluedefault = v.split(",");
		
		try{
			
			objchk.length;
			objvalid = true;
			objnvalid = objchk.length;
			
			if (objnvalid == undefined){
				objnvalid = 1;
			}
					
		}
		catch(e){
			objvalid = false;
		}
		
		if(objvalid){
			//existe objetos e a valida��o inicia
			for(i=0;i<objnvalid;i++){
				for (y = 0; y < objvaluedefault.length; y++){
					//alert(objvaluedefault[y] + "\n" + objchk[i].value)
					if (objchk[i].value == objvaluedefault[y]){ 
						objchk[i].checked = true;
					}
				}
				if(disabled == 1){
					objchk[i].disabled = true;
				}
			}
		}
		
	}
//----------------------------------------------------------------------
function chkValid(obj){
		var objchk = obj
		var objvalid = false;
		var objnvalid = 0;
		
		//var objvaluedefault = v.split(",");
		
		try{
			
			objchk.length;
			objvalid = true;
			objnvalid = objchk.length;
			
			if (objnvalid == undefined){
				objnvalid = 1;
			}
					
		}
		catch(e){
			objvalid = false;
			return false;
		}
		
		if(objvalid){
			//existe objetos e a valida��o inicia
			for(i=0;i<objnvalid;i++){
				//for (y = 0; y < objvaluedefault.length; y++){
					//alert(objvaluedefault[y] + "\n" + objchk[i].value)
					//if (objchk[i].value == objvaluedefault[y]){ 
						//objchk[i].checked = true;
					//}
				//}
				//alert(objchk[i].checked)
				if (objchk[i].checked){
					chkValid = true;
					return true;
				}
				//if(disabled == 1){
					//objchk[i].disabled = true;
				//}
			}
		}
		
		
		
	}
//----------------------------------------------------------------------