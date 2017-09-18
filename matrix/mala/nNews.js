// JavaScript Document


	function nFormNewAdd()
	{
		var validForm = false;
		if(document.getElementById('nomeMala').value!="")
		{
			if(document.getElementById('emailMala').value!="")
			{
				//envia dados para cadastro
				validForm = true;
				AjaxLoad(WebSiteNews + '/matrix/mala/inc_mala_add.asp',allCampQuery('NewsAdm'),'NewsAddRem');
				alert("Seu E-Mail foi adiconado em nossa lista com sucesso!");
			}
		}
		if(!validForm)
		{
			alert("Todos os campos são obrigatórios!");
		}
	}
	
	function nFormNewRem()
	{
		var validForm = false;
		if(document.getElementById('emailMalaRemove').value!="")
		{
				//envia dados para cadastro
				validForm = true;
				AjaxLoad(WebSiteNews + '/matrix/mala/inc_mala_add.asp',allCampQuery('NewsAdm'),'NewsAddRem');
				alert("Seu E-Mail foi removido da nossa lista com sucesso!");
		}
		
		if(!validForm)
		{
			alert("Todos os campos são obrigatórios!");
		}
	}
