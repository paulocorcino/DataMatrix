
var atualizado = 1163156683;
var Total = new Array('26');
var Cidade = new Array('Rio Branco','Maceió','Macapá','Manaus','Salvador','Fortaleza','Vitória','Goiânia','São Luís','Cuiabá','Campo Grande','Belo Horizonte','Belém','João Pessoa','Curitiba','Recife','Teresina','Rio de Janeiro','Natal','Porto Alegre','Porto Velho','Boa Vista','Florianópolis','São Paulo','Aracaju','Palmas','');
var Estado = new Array('ac','al','ap','am','ba','ce','es','go','ma','mt','ms','mg','pa','pb','pr','pe','pi','rj','rn','rs','ro','rr','sc','sp','se','to');
var Clima1 = new Array('pc','pn','pc','ch','pc','pn','ch','ch','nb','nb','pc','nc','pi','pn','pn','pn','pi','ch','pn','pn','pc','pc','pn','pc','pn','pc');
var Clima2 = new Array('ch','pn','pc','pc','pc','pn','ch','nb','pn','pn','pn','nc','pn','pn','cc','pn','pn','nb','pn','nb','ch','pc','pi','pn','nb','pc');
var Clima3 = new Array('pc','pn','pn','pc','pc','pn','ch','nb','pn','pn','cc','nc','nb','cc','pi','pn','pn','nb','pn','pn','nb','ch','nb','pn','pn','ch');
var Tpmin1 = new Array('23','23','25','25','24','26','20','20','26','25','17','17','26','23','12','27','25','17','26','12','22','25','10','13','25','25');
var Tpmin2 = new Array('22','24','24','24','24','26','19','18','24','22','18','16','25','23','8','27','25','15','26','13','23','23','13','13','25','23');
var Tpmin3 = new Array('22','24','24','24','25','26','18','18','25','20','18','17','25','24','14','26','26','16','25','16','22','23','13','13','23','23');
var Tpmax1 = new Array('29','33','32','28','26','31','22','25','32','30','30','23','33','31','22','31','35','21','31','26','33','31','22','19','31','32');
var Tpmax2 = new Array('28','32','32','30','28','31','21','25','33','30','31','19','35','31','22','31','37','22','30','22','32','30','21','22','30','29');
var Tpmax3 = new Array('30','32','33','30','30','30','22','25','33','30','30','20','33','31','18','31','38','24','31','25','32','27','20','23','32','27');
var Tpatual= new Array('22','25','24','27','24','28','22','23','27','23','19','18','25','27','13','29','28','18','29','14','24','25','15','13','26','26');
var Iconatual= new Array('pn','pn','nb','pn','ch','nb','en','en','pn','pn','nb','nb','pn','pn','nb','pn','pn','ch','pn','pn','pn','pn','pn','pn','pn','nb');
var Link= new Array('RioBranco-AC','Maceio-AL','Macapa-AP','Manaus-AM','Salvador-BA','Fortaleza-CE','Vitoria-ES','Goiania-GO','SaoLuis-MA','Cuiaba-MT','CampoGrande-MS','BeloHorizonte-MG','Belem-PA','JoaoPessoa-PB','Curitiba-PR','Recife-PE','Teresina-PI','RiodeJaneiro-RJ','Natal-RN','PortoAlegre-RS','PortoVelho-RO','BoaVista-RR','Florianopolis-SC','SaoPaulo-SP','Aracaju-SE','Palmas-TO');
var x=0;

var timers = 7; //tempo em segundos

timers = timers * 1000;

var y = Cidade.length - 2;

function cid(){

if (x==y){

x=0;

}else{

x++;

}
	try{document.getElementById('cidade').innerHTML = Cidade[x]}catch(e){}
	try{document.getElementById('temMax').innerHTML = Tpmax2[x]+'&ordm;'}catch(e){};
	try{document.getElementById('temMin').innerHTML = Tpmin2[x]+'&ordm;'}catch(e){};
	try{document.getElementById('estado').innerHTML = Estado[x].toUpperCase()}catch(e){};
		var g = new String(Clima2[x])
		if(g=='cc'){
			tTmp = 'Céu Claro';
			it = 'ceuclaro.gif';
		}
		else
		{
			if(g=='ch'){
				t = 'Chuvendo'
				it = 'chvendo.gif';
			}
			else
			{
				if(g=='cv'){
					t = 'Chuviscos'
					it = 'chuvisco.gif';
				}
				else
				{
					if(g=='en'){
						t = 'Encoberto'
						it = 'encoberto.gif';
					}
					else
					{
						if(g=='ge'){
							t = 'Geadas'
							it = 'geadas.gif';
						}
						else
						{
							if(g=='nb'){
								t = 'Nublado'
								it = 'nublado.gif'
							}
							else
							{
								if(g=='ne'){
									t = 'Neve'
									it = 'neve.gif'
								}
								else
								{
									if(g=='pc'){
										t = 'Panc. de Chuva'
										it = 'pancadas.gif'
									}
									else
									{
										if(g=='pi'){
											t= 'Chuvas Rápidas'
											it = 'chavarapida.gif'
										}
										else
										{
											if(g=='pn'){
												t = 'Poucas Nuvens'
												it = 'poucasnuv.gif'
											}
											else
											{
												if(g=='nc'){
													t = 'Nub. c/ Chuva'
													it = 'nubchuva.gif'
												}
												else
												{
													t = '&nbsp;'
													it = '&nbsp;'
												}
											}
										}
									}
								}
							}
						}
					}
				}
			}
		}
		try{document.getElementById('imgtmp').innerHTML="<img src='http://127.0.0.1/bymidia/matrix/img/tempo/"+it+"'>"}catch(e){};
		try{document.getElementById('comenttmp').innerHTML=t}catch(e){};

setTimeout('cid()', timers);

}

cid();

