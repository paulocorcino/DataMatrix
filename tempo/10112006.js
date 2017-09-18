
var Total = new Array('2');
var Cidade = new Array('Aracaju','Estância');
var Clima1 = new Array('nb','nb');
var Clima2 = new Array('nb','nb');
var Clima3 = new Array('nb','nb');
var Tpmin1 = new Array('26','25');
var Tpmin2 = new Array('25','24');
var Tpmin3 = new Array('25','24');
var Tpmax1 = new Array('30','29');
var Tpmax2 = new Array('30','30');
var Tpmax3 = new Array('31','30');
var Tpatual= new Array('');
var Iconatual= new Array('pn','');
var Link= new Array('Aracaju-SE','Estancia-SE');
var x=0;

var timers = 7; //tempo em segundos

timers = timers * 1000;

var y = Cidade.length - 1;

function cid(){

if (x==y){

x=0;

}else{

x++;

}
	try{document.getElementById('cidade').innerHTML = Cidade[x]}catch(e){}
	try{document.getElementById('temMax').innerHTML = Tpmax2[x]+'&ordm;'}catch(e){};
	try{document.getElementById('temMin').innerHTML = Tpmin2[x]+'&ordm;'}catch(e){};
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

