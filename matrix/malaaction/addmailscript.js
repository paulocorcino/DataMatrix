// JavaScript Document
function AddEmail(obj,objn,grupo,domain)
{
	var nDomain = domain;
	var nGrupo = grupo;
	
	try
	{
		if(objn!="" || objn!=null )
		{
			var nObj2 = document.getElementById(objn).value;
		}
	}
	catch(e)
	{
		//
	}
	
	try
	{
		var nObj = document.getElementById(obj).value;	
	}
	catch(e)
	{
		//
	}
	
	
	
	try
	{
		var imagem = new Image();
 		imagem.src = nDomain + "/matrix/malaaction/addmail.asp?grupo=" + nGrupo + "&email=" + nObj + "&nome=" + nObj2;
	}	
	catch(e)
	{
		//
	}
	
	
	//document[img].src = "imgs/estrela_on.gif";
	
}