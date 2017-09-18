<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<script language="JavaScript" type="text/javascript" src="matrix/js/script.js"></script>
<script language="JavaScript" type="text/javascript" src="matrix/agenda/nAgenda.js"></script>
<style type="text/css">
<!--
div#qTip {
	padding: 3px;
	border-right-width: 2px;
	border-bottom-width: 2px;
	color: #000000;
	text-align: left;
	position: absolute;
	z-index: 1000;
	display:none;
	background-color: #99CC00;
	border-top-width: 1px;
	border-left-width: 1px;
	border-top-style: solid;
	border-right-style: solid;
	border-bottom-style: solid;
	border-left-style: solid;
	border-top-color: #996600;
	border-right-color: #996600;
	border-bottom-color: #996600;
	border-left-color: #996600;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 10px;
	font-weight: normal;
	filter: Alpha(Opacity=90);
}

.fontedias {
	font-family: Tahoma, sans-serif;
	font-size: 11px;
	color: #FFFFFF;
}
.borda {
	border: 1px solid #00CCCC;
}
.diasfonte {
	font-family: Tahoma, sans-serif;
	font-size: 10px;
	cursor: default;
}
.mesanofonte {
	font-family: Tahoma, sans-serif;
	font-size: 11px;
	font-weight: normal;
	color: #000000;
}
.buttonnacv {
	font-family: Tahoma, sans-serif;
	font-size: 11px;
	color: #0066CC;
	border: 1px solid #00CCCC;
	background-color: #FFFFFF;
}
.borda_dia {
}
.mess {font-size: 11px}
.mark {
	border: 1px solid #FF0000;
	background-color: #99FF99;
}
.markno {
	
	background-color: #00FFFF;
}
.marksel {
	
	background-color: #00ECEC;
}
.markselmk {
	border: 1px solid #FF0000;
	background-color: #00ECEC;
}
-->
</style>
<div id="eAgenda"></div>
<script>
	AjaxLoad('<%=site%>/matrix/agenda/agenda.asp',null,'eAgenda');
</script>
