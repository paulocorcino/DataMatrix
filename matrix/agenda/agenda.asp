<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->
<%
	cache 1,true

'Corcino Calendário  v1.0 - pecjs@msn.com ===========================================
	
	dim cal_mes, cal_ano, cal_dia, agentext, ims, imy, imi,mark,cComp, primeiro_dia, corrente_dia, celulas, linha, x, coluna, data_g, rs, sql
	dim hmtlagenda
	Function LastDay(MyMonth, MyYear)
	'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=6363&lngWId=4
    ' Returns the last day of the month. Takes into account leap years
    ' Usage: LastDay(Month, Year)
    ' Example: LastDay(12,2000) or LastDay(12) or Lastday
    

    Select Case MyMonth
        Case 1, 3, 5, 7, 8, 10, 12
            LastDay = 31
            
        Case 4, 6, 9, 11
            LastDay = 30
            
        Case 2
            If IsDate(MyYear & "-" & MyMonth & "-" & "29") Then LastDay = 29 Else LastDay = 28
            
        Case Else
            LastDay = 0
    
    	End Select
	End Function
		
		'Coleta querystring ou se query em branco pega o dia ou mes selecionado.
		if request.QueryString("cal_mes")="" then
			cal_mes = month(now)
			cal_ano = year(now)
		else
			cal_mes = cint(request.QueryString("cal_mes"))
			cal_ano = cint(request.QueryString("cal_ano"))
		end if
		
		if request.QueryString("cal_dia")="" then
			cal_dia = day(now)
		else
			cal_dia = cint(request.QueryString("cal_dia"))
		end if
		
	function nav_mes(sinal)
		if sinal = "menos" then
					
				if cal_mes = "1" then
					nav_mes = request.ServerVariables("SCRIPT_NAME")&"?op="&request.QueryString("op")&"&id="&request.QueryString("id")&"&aspid="&aspid&"&id_d="&request.QueryString("id_d")&"&cal_dia="&cal_dia&"&cal_mes=12&cal_ano="&cal_ano-1	
				else
					nav_mes = request.ServerVariables("SCRIPT_NAME")&"?op="&request.QueryString("op")&"&id="&request.QueryString("id")&"&aspid="&aspid&"&id_d="&request.QueryString("id_d")&"&cal_dia="&cal_dia&"&cal_mes="&cal_mes-1&"&cal_ano="&cal_ano
				end if		
		
		elseif sinal = "mais" then
		
				if cal_mes = "12" then
					nav_mes = request.ServerVariables("SCRIPT_NAME")&"?op="&request.QueryString("op")&"&id="&request.QueryString("id")&"&aspid="&aspid&"&id_d="&request.QueryString("id_d")&"&cal_dia="&cal_dia&"&cal_mes=1&cal_ano="&cal_ano+1	
				else
					nav_mes = request.ServerVariables("SCRIPT_NAME")&"?op="&request.QueryString("op")&"&id="&request.QueryString("id")&"&aspid="&aspid&"&id_d="&request.QueryString("id_d")&"&cal_dia="&cal_dia&"&cal_mes="&cal_mes+1&"&cal_ano="&cal_ano
				end if		
		
		else
		
		end if 
	end function
	


hmtlagenda = "<table width=""172"" height=""1"" border=""0"" cellpadding=""2"" cellspacing=""0"" bgcolor=""#FFFFFF"" class=""borda"">"
hmtlagenda = hmtlagenda & "<tr>    <td align=""center"" valign=""top"">      <table width=""170""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">        <tr>"
hmtlagenda = hmtlagenda & "<td height=""23"" colspan=""7"">"
		  		 if aspid="" then
				 	hmtlagenda = hmtlagenda & "<div id=""Layer1"">"
				 end if
           hmtlagenda = hmtlagenda & " <div align=""center"">"
          hmtlagenda = hmtlagenda & "    <select name=""cal_mes"" class=""mesanofonte"" id=""cal_mes"" onChange=""window.open('"&request.ServerVariables("SCRIPT_NAME")&"?op="&request.QueryString("op")&"&aspid="&aspid&"&id="&request.QueryString("id")&"&id_d="&request.QueryString("id_d")&"&cal_dia="&cal_dia&"&cal_mes='+this.value+'&cal_ano="&cal_ano&"','_self')"">"
		    
						
						for ims=1 to 12
			
            hmtlagenda = hmtlagenda & "<option value="""&ims&""" "
			 if ims=cal_mes then
			 	hmtlagenda = hmtlagenda & " selected "
			end if
			hmtlagenda = hmtlagenda & ">"&mesext(ims)&"</option>"

					 	next

					hmtlagenda = hmtlagenda & " </select>"
                 hmtlagenda = hmtlagenda & " <select name=""cal_ano"" class=""mesanofonte"" id=""cal_ano"" onChange=""window.open('"&request.ServerVariables("SCRIPT_NAME")&"?op="&request.QueryString("op")&"&aspid="&aspid&"&id="&request.QueryString("id")&"&id_d="&request.QueryString("id_d")&"&cal_dia="&cal_dia&"&cal_mes="&cal_mes&"&cal_ano='+this.value,'_self')"">"

						imy = cal_ano - 7
						imi = cal_ano + 3
						for ims = imy to imi

			hmtlagenda = hmtlagenda & " <option value="""&ims&""""
			 if ims=cal_ano then
			 hmtlagenda = hmtlagenda & 	"selected"
			end if
			hmtlagenda = hmtlagenda & ">"&ims&"</option>"
	                       
			next

hmtlagenda = hmtlagenda & "                          </select>"
				         
hmtlagenda = hmtlagenda & "          </div>"
 if aspid="" then
 hmtlagenda = hmtlagenda & "</div>"
 end if
                      
             
       hmtlagenda = hmtlagenda & "   </td>"
hmtlagenda = hmtlagenda & "        </tr>"
hmtlagenda = hmtlagenda & "        <tr bgcolor=""#D5AFAA"" class=""fontedias"">"
hmtlagenda = hmtlagenda & "          <td width=""16%"" height=""16""><div align=""center"" >D</div></td>"
hmtlagenda = hmtlagenda & "          <td width=""14%""><div align=""center"" >S</div></td>"
hmtlagenda = hmtlagenda & "          <td width=""14%""><div align=""center"" >T</div></td>"
hmtlagenda = hmtlagenda & "          <td width=""14%""><div align=""center"">Q</div></td>"
hmtlagenda = hmtlagenda & "          <td width=""14%""><div align=""center"" >Q</div></td>"
hmtlagenda = hmtlagenda & "          <td width=""14%""><div align=""center"" >S</div></td>"
hmtlagenda = hmtlagenda & "          <td width=""14%""><div align=""center"" >S</div></td>"
hmtlagenda = hmtlagenda & "        </tr>"
hmtlagenda = hmtlagenda & "        <tr>"
hmtlagenda = hmtlagenda & "          <td colspan=""7"">"
hmtlagenda = hmtlagenda & "              <table width=""103%""  border=""0"" cellpadding=""1"" cellspacing=""1"">"

		
		'Dia em que dia da semana foi o  primeiro dia do mes
		primeiro_dia = weekday(dateserial(cal_ano,cal_mes, 1)) - 1 
		corrente_dia = 1
		celulas = 0
		
		for linha=0 to 5

 hmtlagenda = hmtlagenda & "<tr bgcolor=""#00FFFF"">"

			x = 0
			
			for coluna=0 to 6
				if celulas >= primeiro_dia and corrente_dia <=  LastDay(cal_mes,cal_ano) then
				
					
					data_g= cdate(corrente_dia&"/"&cal_mes&"/"&cal_ano)
					'alert(data_g)
					db "leitura",rs,sql,"SELECT * from AgendaDia where (data_agenda = CONVERT(DATETIME, '"&fDate(data_g,"yyyy-mm-dd")&" 00:00:00', 102))",conn
					                     
					if rs.bof and rs.eof then
						mark = ""
						cComp = "&comp=0"
					else
						mark = "class=mark"
						cComp = "&comp=1"
						db "leitura",rs,sql,"select * from AgendamemoResum where (data_agenda=CONVERT(DATETIME, '"&fDate(data_g,"yyyy-mm-dd")&" 00:00:00', 102))",conn
						agentext = "<b>Dia "&data_g&"</b>:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>"
						while not rs.eof 
							agentext = agentext & " - "&trim(rs("titulo_agenda"))&"<br>"
							rs.movenext
						wend
						agentext = agentext & "<br><br>"
					end if

hmtlagenda = hmtlagenda & "<td "&mark&" width="
 if x < 6 then
  hmtlagenda = hmtlagenda & "14% "
  else
  hmtlagenda = hmtlagenda & "15% "
  end if
  if cal_dia = corrente_dia then
 hmtlagenda = hmtlagenda & " bgcolor=""#00CCFF"""
  else
  hmtlagenda = hmtlagenda & " onMouseOver=""bgOn('"&corrente_dia&"')"" onMouseOut=""bgOff('"&corrente_dia&"')"""


  end if
  hmtlagenda = hmtlagenda & " onClick='go("""&request.ServerVariables("SCRIPT_NAME")&"?op="&request.QueryString("op")&"&aspid="&aspid&"&id="&request.QueryString("id")&"&id_d="&request.QueryString("id_d")&"&cal_dia="&corrente_dia&"&cal_mes="&cal_mes&"&cal_ano="&cal_ano&cComp&""")' id="""&corrente_dia&"""  ><div align=""center"" class=""diasfonte"">"
   if agentext<>"" then
  hmtlagenda = hmtlagenda & " <a title="""&agentext&""">"
   end if
   hmtlagenda = hmtlagenda & corrente_dia
   if agentext<>"" then
  hmtlagenda = hmtlagenda & " </a>"
   end if
   
 hmtlagenda = hmtlagenda & "  </div></td>"

				corrente_dia = corrente_dia + 1
				else

hmtlagenda = hmtlagenda & "					<td width="
if x < 6 then
hmtlagenda = hmtlagenda & "14%"
else
hmtlagenda = hmtlagenda & "14%"
end if
hmtlagenda = hmtlagenda & "	><div align=""center"" class=""diasfonte"">&nbsp;</div></td>"

				end if
				x = x + 1 
				celulas = celulas +1
				
			next

hmtlagenda = hmtlagenda & "                </tr>"

		next

hmtlagenda = hmtlagenda & "              </table><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
hmtlagenda = hmtlagenda & "                <tr>"
hmtlagenda = hmtlagenda & "                  <td><div align=""left"">"
hmtlagenda = hmtlagenda & "                    </div>     "               
hmtlagenda = hmtlagenda & "                    <div align=""center"" class=""diasfonte"">Hoje &eacute; "&day(now)&" de "&mesext(month(now))&" de "&year(now)&" </div>                  <div align=""right"">"
hmtlagenda = hmtlagenda & "                    </div></td>"
hmtlagenda = hmtlagenda & "                </tr>"
hmtlagenda = hmtlagenda & "              </table>"
hmtlagenda = hmtlagenda & "          </td>"
hmtlagenda = hmtlagenda & "        </tr>"
hmtlagenda = hmtlagenda & "    </table>    </td>"
hmtlagenda = hmtlagenda & "  </tr>"
hmtlagenda = hmtlagenda & "</table>"

response.write server.URLEncode(hmtlagenda)
%>
