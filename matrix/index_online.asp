<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<%
	'Cache
	cache 1,"false"
	'response.Write aspid()
	'response.end()
%>
<html>
<head>
<title>Paulo Corcino - DataMatrixs 2004 - Vers&atilde;o 2.5.0</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="css/estilo.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_displayStatusMsg(msgStr) { //v1.0
  status=msgStr;
  document.MM_returnValue = true;
}
//-->
</script>
<script type="text/javascript" language="JavaScript1.2" src="stm31.js"></script>
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_displayStatusMsg('DataMatrixs - 2004 vr. 2.5 / <%=nome_site%> --> <% if trim(aspid())<>"" then %>Usuário: <%=formus("nome",3)%> - <%if trim(formus("nivel",3))=1 then%>Administrador<%else%>Operador<%end if%><% else %> Você está desconectado do DataMatrixs2004 <%end if%>');return document.MM_returnValue" oncontextmenu="return false" onselectstart="return false" ondragstart="return false">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="18" valign="top" bgcolor="#DEDBD6" class="fundo2XP"><script type="text/javascript" language="JavaScript1.2">
<!--
stm_bm(["uueoehr",400,"","blank.gif",0,"","",0,0,0,0,0,1,0,0,""],this);
stm_bp("p0",[0,4,0,0,3,4,0,0,100,"",-2,"",-2,100,0,0,"#000000","transparent","",3,0,0,"#000000"]);
stm_ai("p0i0",[0,"Arquivo","","",-1,-1,0,"","_self","","","","",0,0,0,"","",0,0,0,0,1,"#00ff33",1,"#93a070",0,"","",3,3,0,0,"#fffff7","#000000","#000000","#ffffff","8pt 'Tahoma','Arial'","8pt 'Tahoma','Arial'",0,0]);
stm_bp("p1",[1,4,0,0,2,3,6,4,100,"progid:DXImageTransform.Microsoft.Wipe(GradientSize=1.0,wipeStyle=0,motion=forward,enabled=0,Duration=0.10)",6,"progid:DXImageTransform.Microsoft.Wipe(GradientSize=1.0,wipeStyle=1,motion=forward,enabled=0,Duration=0.10)",5,100,0,0,"#8c8888","#ffffff","",3,1,1,"#aca899"]);
stm_aix("p1i0","p0i0",[0,"Efetuar Logon","","",-1,-1,0,"senha/login.asp","_area","","Efetue o logon. Para poder ter acesso aos menus do site.","","",6,0,0,"","",0,0,0,0,1,"#ffffff",0]);
stm_aix("p1i1","p1i0",[0,"Editar Minha Conta","","",-1,-1,0,"senha/conta_ind.asp?op=alt&aspid=<%=aspid()%>","_area","","Efetua alterações em sua conta de acesso.","","",0]);
stm_aix("p1i2","p1i0",[0,"Gerenciar Contas","","",-1,-1,0,"","_area","","","","",6,0,0,"setdir.gif","setdir.gif",4,7]);
stm_bpx("p2","p1",[1,2,0,0,2,3,6,0]);
stm_aix("p2i0","p1i0",[0,"Nova Conta                           ","","",-1,-1,0,"senha/contas.asp?op=cad&aspid=<%=aspid()%>","_area","","Crie as contas dos usuários do DataMatrix2004. É necessário ter nivel de Administrador."]);
stm_aix("p2i1","p1i1",[0,"Listar Contas","","",-1,-1,0,"senha/contas.asp?op=listar&aspid=<%=aspid()%>","_area","","Liste e edite as contas dos usuários do DataMatrix2004. É necessário ter nivel de Administrador."]);
stm_ep();
stm_aix("p1i3","p1i1",[0,"Página Inicial","","",-1,-1,0,"intro.asp?aspid=<%=aspid()%>","_area","","Página Inicial do DataMatrixs2004"]);
stm_ai("p1i4",[6,1,"#aca899","",0,0,0]);
stm_aix("p1i5","p1i0",[0,"Fechar                               ALT+F4","","",-1,-1,0,"sair.asp?aspid=<%=aspid()%>","_self","","Fecha o DataMatrixs2004"]);
stm_ep();
/*
stm_aix("p0i1","p0i0",[0,"Estatísticas"]);
stm_bpx("p3","p1",[]);
stm_aix("p3i0","p1i0",[0,"Sumário                                      ","","",-1,-1,0,"contador/reports.asp?aspid=<%=aspid()%>&cods={376D1B34-D87F-4FB8-AA0C-312076ADE4E6}","_area","","Tenha em mãos a atual situação dasvisitas de seu site."]);
stm_aix("p3i1","p1i2",[0,"Estatísticas Detalhadas","","",-1,-1,0,"","_self","","Todas as informações do site"]);
stm_bpx("p4","p2",[]);
stm_aix("p4i0","p1i0",[0,"Relatório Detalhado                    ","","",-1,-1,0,"contador/reportd.asp?aspid=<%=aspid()%>&cods={376D1B34-D87F-4FB8-AA0C-312076ADE4E6}","_area","","Veja todas as estatísticas das visitas do site."]);
stm_aix("p4i1","p1i0",[0,"Page Views","","",-1,-1,0,"contador/reportpath.asp?aspid=<%=aspid()%>&cods={376D1B34-D87F-4FB8-AA0C-312076ADE4E6}","_area","","Vejas as páginas mais acessadas do seu site."]);
stm_aix("p4i2","p1i0",[0,"Referências","","",-1,-1,0,"contador/reportref.asp?aspid=<%=aspid()%>&cods={376D1B34-D87F-4FB8-AA0C-312076ADE4E6}","_area","","Veja as referência de sites que indicaram seu site."]);
stm_aix("p4i3","p1i0",[0,"Histórico","","",-1,-1,0,"contador/reportpathy.asp?aspid=<%=aspid()%>&cods={376D1B34-D87F-4FB8-AA0C-312076ADE4E6}","_area","","Veja Antigas informações das visitas de seu site."]);
stm_aix("p4i4","p1i0",[0,"Estatísticas IP","","",-1,-1,0,"contador/ips.asp?aspid=<%=aspid()%>&cods={376D1B34-D87F-4FB8-AA0C-312076ADE4E6}","_area","","Veja os ips dos visitantes e que páginas ele visitou."]);
stm_ep();
stm_aix("p3i2","p1i2",[0,"Gráficos","","",-1,-1,0,"","_self","","Gráficos detalhados mostam todas as infomrções necessárias sobre as visitas."]);
stm_bpx("p5","p2",[]);
stm_aix("p5i0","p1i0",[0,"Horarios                                      ","","",-1,-1,0,"contador/graphs.asp?Type=Hour&aspid=<%=aspid()%>&cods={376D1B34-D87F-4FB8-AA0C-312076ADE4E6}","_area","","Veja o gráfico das visitas de seu site tomando como referência os Harários."]);
stm_aix("p5i1","p1i0",[0,"Dias da Semana","","",-1,-1,0,"contador/graphs.asp?Type=DOW&aspid=<%=aspid()%>&cods={376D1B34-D87F-4FB8-AA0C-312076ADE4E6}","_area","","Veja o gráfico das visitas de seu site tomando como referência os Dias da Semana."]);
stm_aix("p5i2","p1i0",[0,"Dias do Mês","","",-1,-1,0,"contador/graphs.asp?Type=DOM&aspid=<%=aspid()%>&cods={376D1B34-D87F-4FB8-AA0C-312076ADE4E6}","_area","","Veja o gráfico das visitas de seu site tomando como referência os Dias do Mês."]);
stm_aix("p5i3","p1i0",[0,"Semanal","","",-1,-1,0,"contador/graphs.asp?Type=Week&aspid=<%=aspid()%>&cods={376D1B34-D87F-4FB8-AA0C-312076ADE4E6}","_area","","Veja o gráfico das visitas de seu site tomando como referência as Semanas do Ano."]);
stm_aix("p5i4","p1i0",[0,"Mensal","","",-1,-1,0,"contador/graphs.asp?Type=Month&aspid=<%=aspid()%>&cods={376D1B34-D87F-4FB8-AA0C-312076ADE4E6}","_area","","Veja o gráfico das visitas de seu site tomando como referência os Meses."]);
stm_aix("p5i5","p1i0",[0,"Anual","","",-1,-1,0,"contador/graphs.asp?Type=Year&aspid=<%=aspid()%>&cods={376D1B34-D87F-4FB8-AA0C-312076ADE4E6}","_area","","Veja o gráfico das visitas de seu site tomando como referência os Anos."]);
stm_ep();
stm_ep();
*/
stm_aix("p0i2","p0i0",[0,"Editar Site"]);
stm_bpx("p6","p1",[]);
stm_aix("p6i0","p1i2",[0,"Gerencia Enquetes                        "]);
stm_bpx("p7","p2",[]);
stm_aix("p7i0","p1i0",[0,"Categorias                          ","","",-1,-1,0,"enquete/area.asp?aspid=<%=aspid()%>","_area","","Cria as Sessões do Site."]);
stm_aix("p7i0","p1i0",[0,"Criar Categoria                     ","","",-1,-1,0,"enquete/area.asp?op=cad&aspid=<%=aspid()%>","_area","","Cria as Sessões do Site."]);
stm_aix("p7i0","p1i0",[0,"Criar Enquete                          ","","",-1,-1,0,"enquete/enquete.asp?op=cad1&aspid=<%=aspid()%>","_area","","Cria as Sessões do Site."]);
stm_aix("p7i0","p1i0",[0,"Listar Enquetes                          ","","",-1,-1,0,"enquete/enquete.asp?op=listar&aspid=<%=aspid()%>","_area","","Cria as Sessões do Site."]);
stm_ep();
/*
stm_aix("p6i0","p1i2",[0,"Gerencia Foto Enquete"]);
stm_bpx("p7","p2",[]);
stm_aix("p7i0","p1i0",[0,"Categorias                          ","","",-1,-1,0,"enquetefoto/area.asp?aspid=<%=aspid()%>","_area","","Cria as Sessões do Site."]);
stm_aix("p7i0","p1i0",[0,"Criar Categoria                     ","","",-1,-1,0,"enquetefoto/area.asp?op=cad&aspid=<%=aspid()%>","_area","","Cria as Sessões do Site."]);
stm_aix("p7i0","p1i0",[0,"Criar Enquete                          ","","",-1,-1,0,"enquetefoto/enquete.asp?op=cad1&aspid=<%=aspid()%>","_area","","Cria as Sessões do Site."]);
stm_aix("p7i0","p1i0",[0,"Listar Enquetes                          ","","",-1,-1,0,"enquetefoto/enquete.asp?op=listar&aspid=<%=aspid()%>","_area","","Cria as Sessões do Site."]);
stm_ep();

stm_aix("p6i1","p1i2",[0,"Gerenciar Banner"]);
stm_bpx("p8","p2",[]);
stm_aix("p8i0","p1i0",[0,"Nova Área                        ","","",-1,-1,0,"banner/gen_area.asp?op=cad&aspid=<%=aspid()%>","_area","","Cria novas área para banner."]);
stm_aix("p8i0","p1i0",[0,"Listar Área                        ","","",-1,-1,0,"banner/gen_area.asp?op=list&aspid=<%=aspid()%>","_area","","Listagem de área para banner."]);
stm_aix("p8i0","p1i0",[0,"Novo Cliente                       ","","",-1,-1,0,"banner/gen_cliente.asp?op=cad&aspid=<%=aspid()%>","_area","","Cria novas Clientes para banner."]);
stm_aix("p8i0","p1i0",[0,"Listar Cliente                        ","","",-1,-1,0,"banner/gen_cliente.asp?op=list&aspid=<%=aspid()%>","_area","","Listagem de Clientes para banner."]);
stm_aix("p8i1","p1i0",[0,"Listar Temas","","",-1,-1,0,"bomfim/temas/temas.asp?op=listar&aspid=<%=aspid()%>","_area","","Liste os temas do Site."]);
stm_ep();

stm_aix("p6i1","p1i2",[0,"Gerenciar Materias"]);
stm_bpx("p8","p2",[]);

stm_aix("p8i0","p1i0",[0,"Gerenciar Temas                          ","","",-1,-1,0,"materias/sessao.asp?op=tabelas&aspid=<%=aspid()%>","_area","","Envia os temas do Site."]);
stm_aix("p8i1","p1i0",[0,"Nova Materia","","",-1,-1,0,"materias/materias.asp?op=cad&aspid=<%=aspid()%>","_area","","Liste os temas do Site."]);
stm_aix("p8i1","p1i0",[0,"Listar Materia","","",-1,-1,0,"materias/materias.asp?op=listar&aspid=<%=aspid()%>","_area","","Liste os temas do Site."]);
stm_ep();

stm_aix("p6i1","p1i2",[0,"Gerenciar Album"]);
stm_bpx("p9","p2",[]);
stm_aix("p9i0","p1i0",[0,"Gerenciar Areas                      ","","",-1,-1,0,"album/area.asp?op=tabelas&aspid=<%=aspid()%>","_area","","Cria uma novas Areas de Fotos"]);
stm_aix("p9i0","p1i0",[0,"Nova Sessão                          ","","",-1,-1,0,"album/session.asp?op=cads&aspid=<%=aspid()%>","_area","","Cria uma nova sessão de fotos"]);
stm_aix("p9i1","p9i0",[0,"Listar Sessão","","",-1,-1,0,"album/session.asp?op=listar&aspid=<%=aspid()%>","_area","","Liste todas as sessões de fotos"]);
stm_ep();

stm_aix("p6i1","p1i2",[0,"Gerenciar Promoções"]);
stm_bpx("p9","p2",[]);
stm_aix("p9i0","p1i0",[0,"Nova Promoção                      ","","",-1,-1,0,"promo/promo.asp?op=cad&aspid=<%=aspid()%>","_area","","Cria uma novas Areas de Fotos"]);
stm_aix("p9i0","p1i0",[0,"Listar Promoções                          ","","",-1,-1,0,"promo/promo.asp?op=list&aspid=<%=aspid()%>","_area","","Cria uma nova sessão de fotos"]);

stm_ep();
stm_aix("p6i1","p1i2",[0,"Gerenciar Destaques"]);
stm_bpx("p9","p2",[]);
stm_aix("p9i0","p1i0",[0,"Novo Destaque                      ","","",-1,-1,0,"destaques/dest.asp?op=cad&aspid=<%=aspid()%>","_area","","Cria uma novas Areas de Fotos"]);
stm_aix("p9i0","p1i0",[0,"Listar Destques                         ","","",-1,-1,0,"destaques/dest.asp?op=list&aspid=<%=aspid()%>","_area","","Cria uma nova sessão de fotos"]);
stm_aix("p9i0","p1i0",[0,"Gerar XML                         ","","",-1,-1,0,"destaques/dest.asp?op=xml&aspid=<%=aspid()%>","_area","","Cria uma nova sessão de fotos"]);

stm_ep();
stm_aix("p6i1","p1i2",[0,"Gerenciar Mural e Scraps"]);
stm_bpx("p9","p2",[]);
stm_aix("p9i0","p1i0",[0,"Mural                        ","","",-1,-1,0,"mural/mural.asp?op=mural&aspid=<%=aspid()%>","_area","","Cria uma nova sessão de fotos"]);
stm_aix("p9i0","p1i0",[0,"Scrap                         ","","",-1,-1,0,"mural/mural.asp?op=scrap&aspid=<%=aspid()%>","_area","","Cria uma nova sessão de fotos"]);
stm_ep();

stm_aix("p6i0","p1i2",[0,"Gerencia Anuncios                        "]);
stm_bpx("p7","p2",[]);
stm_aix("p7i0","p1i0",[0,"Criar Area                          ","","",-1,-1,0,"ban/session.asp?op=cads&aspid=<%=aspid()%>","_area","","Cria as Sessões do Site."]);
stm_aix("p7i0","p1i0",[0,"Listar Anuncios                          ","","",-1,-1,0,"ban/session.asp?op=listar&aspid=<%=aspid()%>","_area","","Cria as Sessões do Site."]);
stm_ep();
stm_aix("p6i1","p1i2",[0,"Informativos"]);
stm_bpx("p9","p2",[]);
stm_aix("p9i0","p1i0",[0,"Nova InfoZoeira                      ","","",-1,-1,0,"info/info.asp?op=cad&aspid=<%=aspid()%>","_area","","Cria uma nova sessão de fotos"]);
stm_aix("p9i1","p9i0",[0,"Listar InfoZoeira","","",-1,-1,0,"info/info.asp?op=list&aspid=<%=aspid()%>","_area","","Liste todas as sessões de fotos"]);
stm_ep();*/
/*
stm_aix("p6i2","p1i2",[0,"Gerenciar Noticias","","",-1,-1,0,"","_self","","","","",0]);
stm_bpx("p10","p2",[]);
stm_aix("p10i0","p1i1",[0,"Gerenciar Sessão                ","","",-1,-1,0,"noticia/sessao.asp?op=tabelas&aspid=<%=aspid()%>","_area","",""]);
stm_aix("p10i0","p1i1",[0,"Nova Noticia                 ","","",-1,-1,0,"noticia/noticias.asp?op=cad&aspid=<%=aspid()%>","_area","",""]);
stm_aix("p10i1","p10i0",[0,"Listar Noticias","","",-1,-1,0,"noticia/noticias.asp?op=listar&aspid=<%=aspid()%>"]);
stm_ep();

stm_aix("p6i2","p1i2",[0,"Gerenciar Mala Direta","","",-1,-1,0,"","_self","","","","",0]);
stm_bpx("p10","p2",[]);
stm_aix("p6i2","p1i0",[0,"Editor de Conteúdo","","",-1,-1,0,"conteudo/conteudo.asp?op=listar&aspid=<%=aspid()%>","_area","",""]);
stm_ep();

stm_aix("p6i2","p1i2",[0,"Gerenciar Usuários","","",-1,-1,0,"","_self","","","","",0]);
stm_bpx("p10","p2",[]);
stm_aix("p10i0","p1i1",[0,"Novo Usuário        ","","",-1,-1,0,"cad/session.asp?op=alt&aspid=<%=aspid()%>","_area","",""]);
stm_aix("p10i0","p1i1",[0,"Mensagens de Anivesário        ","","",-1,-1,0,"niver/area.asp?op=tabelas&aspid=<%=aspid()%>","_area","",""]);
stm_aix("p10i0","p1i1",[0,"Listar Usuários        ","","",-1,-1,0,"cad/session.asp?op=list&aspid=<%=aspid()%>","_area","",""]);
stm_ep();
*/
stm_aix("p6i2","p1i0",[0,"Banner","","",-1,-1,0,"app_banner/default.asp?aspid=<%=aspid()%>","_area","",""]);
stm_aix("p6i2","p1i0",[0,"Agenda","","",-1,-1,0,"agenda/default.asp?aspid=<%=aspid()%>","_area","",""]);
stm_aix("p6i2","p1i0",[0,"Foto Enquete","","",-1,-1,0,"enquetefoto/default.asp?aspid=<%=aspid()%>","_area","",""]);
stm_aix("p6i2","p1i0",[0,"Mala Direta","","",-1,-1,0,"mala/default.asp?aspid=<%=aspid()%>","_area","",""]);
stm_aix("p6i2","p1i0",[0,"Noticias","","",-1,-1,0,"noticia/default.asp?aspid=<%=aspid()%>","_area","",""]);
stm_aix("p6i2","p1i0",[0,"Álbum Book","","",-1,-1,0,"album_book/default.asp?aspid=<%=aspid()%>","_area","",""]);
stm_aix("p6i2","p1i0",[0,"Editor de Conteúdo","","",-1,-1,0,"conteudo/conteudo.asp?op=listar&aspid=<%=aspid()%>","_area","",""]);
stm_ep();
stm_ep();

stm_aix("p0i3","p0i0",[0,"Ajuda"]);
stm_bpx("p11","p2",[1,4]);
stm_aix("p11i0","p1i0",[0,"Ajuda                                     ","","",-1,-1,0,"","_area","","Ajuda sobre o DataMatrixs2004. Em desenvolvimento."]);
stm_aix("p11i1","p1i4",[]);
stm_aix("p11i2","p1i0",[0,"Sobre","","",-1,-1,0,"sobre.asp","_area","","Sobre o DataMatrixs2004."]);
stm_ep();
stm_ep();
stm_em();
//-->
</script>






    </td>
  </tr>
  <tr> 
    <td height="95%" valign="top" class="planoarea">
<table width="100%" height="100%" border="0" align="center" cellpadding="6" cellspacing="3">
        <tr> 
          <td height="100%" valign="top" bgcolor="#FFFFFF" class="fonte"><iframe src="<% if request.QueryString("pag")<>"" then %> <%=request.QueryString("pag")%> <%else%>blank.asp<%end if%>" name="_area" width="100%" marginwidth="0" height="100%" marginheight="0" scrolling="auto" frameborder="0" ></iframe></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="19" valign="top" class="fundo3XP">&nbsp;</td>
  </tr>
</table>
</body>
</html>
