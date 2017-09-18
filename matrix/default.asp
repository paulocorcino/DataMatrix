<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>

<html>
<head>
<title>DataMatrixs2004</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="css/estilo.css" rel="stylesheet" type="text/css">
<META NAME='ROBOTS' CONTENT='INDEX,NOFOLLOW'>
<meta http-equiv="imagetoolbar" content="no">
<script language="JavaScript">
<!--
var msg="Direito Não";
function disableIE() {if (document.all) {return false;}
}
function disableNS(e) {
  if (document.layers||(document.getElementById&&!document.all)) {
    if (e.which==2||e.which==3) {return false;}
  }
}
if (document.layers) {
  document.captureEvents(Event.MOUSEDOWN);document.onmousedown=disableNS;
} else {
  document.onmouseup=disableNS;document.oncontextmenu=disableIE;
}
document.oncontextmenu=new Function("return false")

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
<style type="text/css">
<!--
.tabele_invi {
	filter: Alpha(Opacity=90);
	border: 1px solid #006633;
}
.style1 {color: #666666}
body {
	background-color: #FFFFFF;
}
.style2 {color: #FF0000}
-->
</style>
</head>

<body  oncontextmenu="return false" onselectstart="return false" ondragstart="return false">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="middle"><div align="center">
      <table width="457" height="252"  border="0" cellpadding="0" cellspacing="0" background="img/logo.gif">
        <tr>
          <td><div align="center">
            <form name="form1" method="post" action="check.asp">
              <table width="213"  border="0" cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="tabele_invi">
			    <tr bgcolor="#F3F3F3">
                  <td width="203" height="68" valign="middle" class="fonte style2"><div align="center" id="texto">
                    <p>Seu navegador está sem suporte a JavaScript, favor ative-o ou seu navegador n&atilde;o &eacute; o Internet Explorer 6.0 ou superior. </p>
                  </div>
				    <script>
<!--
	//detecta componentes necessários.
	//document.all.texto.innerHTML = navigator.onLine
	var tmpcookie = new Date();
	var chkcookie = (tmpcookie.getTime() + '');
	var version=0
	if (navigator.appVersion.indexOf("MSIE")!=-1){
		temp=navigator.appVersion.split("MSIE")
		version=parseFloat(temp[1])
	}

	if (version<=5.5){
		
		//Avisa que o usuário necessita do internet explorer 6
		document.all.texto.innerHTML = "Para usar o DataMatrixs2004 é necessário ter o Internet Explorer 6.0 ou +"
	}else{
		
	   //verifica se o usuário tem os cookie ativados
	   document.cookie = "chkcookie=" + chkcookie + "; path=/";
	   if (document.cookie.indexOf(chkcookie,0) < 0) {
	   
			document.all.texto.innerHTML = "Para usar o DataMatrixs2004 é necessário ativar o suporte a Cookie. Duvidas consulte o Webmaster do site."
    	
		}else{
			
			//Aparece texto
			document.all.texto.innerHTML = "Aguarde inicializando o DataMatrixs2004..."
			
			//abre o programa
			MM_openBrWindow('index_online.asp','datamatrixs','status=yes,resizable=yes,width=640,height=480');
			
			//Aparece texto
			document.all.texto.innerHTML = '# DataMatrixs2004 #<br><a href="javascript: history.go(-1)"><span class="linksimples">Voltar a Página</span> '
		
		}
	}
	//-->
</script>				</a></td>

                  </tr>
				</table>
            </form>
          </div></td>
        </tr>
      </table>
    </div></td>
  </tr>
</table>
</body>
</html>
