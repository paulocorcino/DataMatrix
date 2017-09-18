// JavaScript Document

	var ShowAlbum = true;

	//testa a presença dos componentes antes de executar o script
	var aDivObjError = "";
	var aDivObjMsg;
	var aDivObj = new Array();
	//este objetos são obrigatórios
	aDivObj[0] = "aAlbumThumb" //Local dos Thumbmail
	aDivObj[1] = "aAlbumVoltar" //botão voltar
	aDivObj[2] = "aAlbumAvancar" //botao avancar
	aDivObj[3] = "aAlbumListPage" //lista de numero das paginas
	aDivObj[4] = "aAlbumNomeFoto" //nome da foto
	aDivObj[5] = "aAlbumClick" //numero de clicks
	aDivObj[6] = "aAlbumEnviar" //enviar foto
	aDivObj[7] = "aAlbumShowFoto" //local exibir foto

 for(var objAlbumList=0; objAlbumList <  aDivObj.length; objAlbumList++){
	
	try{
		
		document.getElementById(aDivObj[objAlbumList]).innerHTML = "";
	}
	catch(e){
		aDivObjError = aDivObj[objAlbumList] + ", " + aDivObjError
	}
	
 }
 
 if(aDivObjError!="")
 {
	 aDivObjMsg = "Erro no carregamento do album:\nVeja abaixo as div que estão faltando nesta página.\n\n" + aDivObjError + "\n\nContate o administrador para solucionar o problema. O album não será carregado.";
	 alert(aDivObjMsg);
	 ShowAlbum = false;
 }

function AjaxLoadAlbum(pg,query,camada,nostart)
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
	  // Página do Post
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
					if(nostart!=1){
						AlbumStartButtons();
					}
				//}else{
				//	document.getElementById(camada).innerHTML = "<div aling=left class=fonte_10_b>Erro na abertura da página: " + http.status + " - " + http.statusText + "</div>";
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
		//ocasião especial
			iTimeoutId = setTimeout(function(){http.send('sel=' + document.getElementById('fotosel').value);},800);
			
		}else{
			
			iTimeoutId = setTimeout(function() {http.send(null);},800);
			
		}
	}
	  
}
	

	

function AlbumStartButtons()
	{
		var AlbumPgActual;
		var AlbumPgFinal;
		var AlbumPg;
		var AlbumObjPages;
		
		AlbumObjPages = "<select name=Gpages id='aAlbumSelectPage' onchange='AjaxLoadAlbum(\""+AlbumLinkDefault+"\"+this.value,null,\"aAlbumThumb\")'>";
		
		AlbumPg = document.getElementById('AlbumStat').value.split(",");
		
		AlbumPgActual = AlbumPg[0];
		AlbumPgFinal = AlbumPg[1];
					
			for(var i=1; i <= AlbumPgFinal; i++){
				AlbumObjPages = AlbumObjPages + "<option value='"+i+"'";
				
					if(AlbumPgActual==i){
						
						AlbumObjPages = AlbumObjPages + " selected ";
					}
				
				AlbumObjPages = AlbumObjPages + ">"+i+"</option>";
			}
		AlbumObjPages = AlbumObjPages + "</select>"
		
		document.getElementById('aAlbumListPage').innerHTML = AlbumObjPages;
		
		
		if(AlbumPgActual>1){
			document.getElementById('aAlbumVoltar').innerHTML   = "<a href='javascript:void(0)' onclick='AjaxLoadAlbum(\"" + AlbumLinkDefault + (parseInt(AlbumPgActual)-1) +"\",null,\"aAlbumThumb\")'>" + imgAlbumVoltar + "</a>";
		}
		else
		{
			//nada
			document.getElementById('aAlbumVoltar').innerHTML  = imgAlbumVoltarIn;
		}
		
		if(AlbumPgActual<AlbumPgFinal){
			document.getElementById('aAlbumAvancar').innerHTML  = "<a href='javascript:void(0)' onclick='AjaxLoadAlbum(\"" + AlbumLinkDefault + (parseInt(AlbumPgActual)+1) +"\",null,\"aAlbumThumb\")'>" + imgAlbumAvanca + "</a>";
		}
		else
		{
			document.getElementById('aAlbumAvancar').innerHTML = imgAlbumAvancaIn;
			//nada
		}		

		
	}
	
	function imgOpenFoto(n,name,clic)
	{
		//alert(i);
		var cookieFoto = getCookie('eFoto'+n);
		
		if(cookieFoto!=null){
			
			setCookie('eFoto'+n, '1');
			
		}else{
		
			setCookie('eFoto'+n, '0');
			
		}
		//alert(getCookie('eFoto'+n))
		document.getElementById('aAlbumShowFoto').innerHTML = "<img src='"+AlbumUrl+"/matrix/album_book/inc_album_foto.asp?id="+n+"' border=0 id='aAlbumShowFotoImg'>";
		document.getElementById('aAlbumNomeFoto').innerHTML = name;
		document.getElementById('aAlbumClick').innerHTML = clic;
		document.getElementById('aAlbumEnviar').innerHTML = "<a href='javascript:void(0)' onclick='imgSendFoto(\""+n+"\")'>"+imgAlbumEnviar+"</a>";
	}
	
	function imgOpenFotoNot(n)
	{
		//alert(i);
		var cookieFoto = getCookie('eFoto'+n);
		
		if(cookieFoto!=null){
			
			setCookie('eFoto'+n, '1');
			
		}else{
		
			setCookie('eFoto'+n, '0');
			
		}
		//alert(getCookie('eFoto'+n))
		document.getElementById('aAlbumShowFoto').innerHTML = "<img src='"+AlbumUrl+"/matrix/album_book/inc_album_foto.asp?id="+n+"' border=0 id='aAlbumShowFotoImg'>";
	}
	
	function imgSendFoto(n)
	{
		  
		 AlbumSendEmailText = "<form name='AlbumEmail'><input name=\"albumFotoId\" type=\"hidden\" id=\"albumFotoId\" value="+n+"><p><font id='aAlbumEmailMyName'>Seu Nome</font><br>"
    	 AlbumSendEmailText = AlbumSendEmailText + "<input name=\"albumMyname\" type=\"text\" id=\"albumMyname\">" 
    	 AlbumSendEmailText = AlbumSendEmailText + "<br>"
		 AlbumSendEmailText = AlbumSendEmailText + "<font id='aAlbumEmailMyEmail'>Seu E-Mail</font><br>"
    	 AlbumSendEmailText = AlbumSendEmailText + "<input name=\"albumMyemail\" type=\"text\" id=\"albumMyemail\" onBlur=\"if(!isEmail(this.value)){this.value=''}\">"
    	 AlbumSendEmailText = AlbumSendEmailText + "<br>"
    	 AlbumSendEmailText = AlbumSendEmailText + "<font id='aAlbumEmailFriendName'>Nome do Amigo</font><br>"
    	 AlbumSendEmailText = AlbumSendEmailText + "<input name=\"albumFriendnome\" type=\"text\" id=\"albumFriendnome\">"
    	 AlbumSendEmailText = AlbumSendEmailText + "<br>"
    	 AlbumSendEmailText = AlbumSendEmailText + "<font id='aAlbumEmailFriendEmail'>E-Mail do Amigo</font><br>"
    	 AlbumSendEmailText = AlbumSendEmailText + "<input name=\"albumFriendemail\" type=\"text\" id=\"albumFriendemail\" onBlur=\"if(!isEmail(this.value)){this.value=''}\">"
    	 AlbumSendEmailText = AlbumSendEmailText + "<br>"
    	 AlbumSendEmailText = AlbumSendEmailText + "<input type=\"button\" name=\"Button\" id='aAlbumEmailButton' value=\"Enviar\" onclick=\"ValidSendMail('"+n+"')\">"
  		 AlbumSendEmailText = AlbumSendEmailText + "</p><p><img src=\""+AlbumUrl+"/matrix/album_book/inc_album_foto.asp?id="+n+"\"";
		 
		 if(AlbumW != "0" && AlbumW != "")
		 {
			  AlbumSendEmailText = AlbumSendEmailText + " width=\""+AlbumW+"\" "
		 }
		 else
		 {
			  AlbumSendEmailText = AlbumSendEmailText + " width=\"15%\" "
		 }
		 
		 if(AlbumH != "0" && AlbumH != "")
		 {
			 AlbumSendEmailText = AlbumSendEmailText + " height=\""+AlbumH+"\" "
		 }
		 
		  AlbumSendEmailText = AlbumSendEmailText + " id='aAlbumEmailImg'> </p></form>"
		 
  
		setTimeout("document.getElementById('aAlbumShowFoto').innerHTML = AlbumSendEmailText",500);
		
	}
	
	function ValidSendMail(n){
		
		var ValidOK = false;
		
		if(document.getElementById('albumMyname').value!="")
		{
			if(document.getElementById('albumMyemail').value!="")
			{
				if(document.getElementById('albumFriendnome').value!="")
				{
					if(document.getElementById('albumFriendemail').value!="")
					{
						ValidOK = true
						AjaxLoadAlbum(AlbumUrl+'/matrix/album_book/inc_album_send.asp',allCampQuery('AlbumEmail'),'aAlbumSendFoto',1);
						alert('Fotografia enviada com sucesso!');
						imgOpenFotoNot(n);
					}
				}
			}
		}
		
		if(!ValidOK)
		{
			alert("Todos os campos são obrigatórios!!!");
		}
		
		
	}