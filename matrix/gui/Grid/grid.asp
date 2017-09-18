<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bymidia/matrix/config/config.asp" -->

<%
	cache 1, false
	'Response.CodePage = 1254
	
	dim Txts, Gx, GridN, GridSQL, GridData(), GridName(), Grs, GridDate, GPage, GridR, GridPoss, GridSqlOr, GridSqlTransf
	
	'//--------------------------------------------------------------------------------->
	
		Gpage = request.form("gpage") 'Numero da página aatual	
		GridN = cint(request.form("Gridn")) - 1 'Numero de campos
		GridR = cint(request.form("gridr")) 'numero de intens na página
		GridN = cint(GridN)
		
		Redim GridData(GridN)
		Redim GridName(GridN)
		
		GridSQL = Decrypt(request.form("gridid"),Gkeys) 'Query SQL
		GridSqlOr = GridSQL
		'response.write gridsql
		'GridName ----------------------------->
		for Gx = 0 to GridN
			GridName(Gx) = Decrypt(request.form("gridname_"&gx),Gkeys) 'Nome Campos
			'response.write GridName(Gx)
		next
		
		'GridData ----------------------------->
		for Gx = 0 to GridN
			GridData(Gx) = Decrypt(request.form("griddata_"&gx),Gkeys) 'Campos na Tabela
		next
		
		GridDate =  Decrypt(request.form("griddate"),Gkeys) 'Formatação da Data
		GridPoss =  Decrypt(request.form("gridposs"),Gkeys) 'Formatação da Data
		
		
		'//Filtros -------------------------------------------------------------------
		
				
			'//[Incluir filtro] ----------------------------------------------------------
			if request("GRIDFilter") <> "" then
			GridSqlTransf = instr(GridSQL,"where")
			if GridSqlTransf <> 0 then
				GridSQL = left(GridSQL,GridSqlTransf+4) & " "&request("GRIDFilterF")
				
				'if isnumeric(request("GRIDFilter")) then
				'	GridSQL = GridSQL & " = " & request("GRIDFilter") & " and "
				'else
					GridSQL = GridSQL & " like " & " '"& replace(request("GRIDFilter"),"?","%") & "' and "
				'end if 
				
				  GridSQL = GridSQL & right(GridSqlOr,len(GridSqlOr) - (instr(GridSqlOr,"where")+4))
			else
			
				GridSqlTransf = instr(GridSQL,"having")
				if GridSqlTransf <> 0 then
					GridSQL =  left(GridSQL,GridSqlTransf+5) & " "&request("GRIDFilterF")
					
				'if isnumeric(request("GRIDFilter")) then
				'	GridSQL = GridSQL & " = " & request("GRIDFilter") & " and "
				'else
					GridSQL = GridSQL & " like " & " '"& replace(request("GRIDFilter"),"?","%")  & "' and "
				'end if 
					
					GridSQL = GridSQL & right(GridSqlOr,len(GridSqlOr) - (instr(GridSqlOr,"having")+5))
					
				else
				
					GridSqlTransf = instr(GridSQL,"order")
					if GridSqlTransf <> 0 then
					
						GridSQL = left(GridSQL,GridSqlTransf-1) & " where  "&request("GRIDFilterF") 
						
						'if isnumeric(request("GRIDFilter")) then
							'GridSQL = GridSQL & " = " & request("GRIDFilter") & " and "
						'else
							GridSQL = GridSQL & " like " & " '"& replace(request("GRIDFilter"),"?","%")  & "' and "
						'end if 
						
						 GridSQL = GridSQL & right(GridSqlOr,len(GridSqlOr) - (instr(GridSqlOr,"order")-1))
					
					else
						
						GridSQL =  GridSQL & " where  "&request("GRIDFilterF")
						
						'if isnumeric(request("GRIDFilter")) then
							'GridSQL = GridSQL & " = " & request("GRIDFilter") 
						'else
							GridSQL = GridSQL & " like " & " '"& replace(request("GRIDFilter"),"?","%")  & "' "
						'end if 	
						
					end if 		
				end if
			
			end if
			end if
			'//----------------------------------------------------------------------------
			
			'//[incluir order by] ----------------------------------------------------------
			if request("GRIDORDER") <> "" then
			
			GridSqlTransf = instr(GridSQL,"order")
				if GridSqlTransf <> 0 then
					
						GridSQL =  left(GridSQL,GridSqlTransf-1) & " order by " & request("GRIDORDER") & " " & request("GRIDORDERF")
					
				else
			
					GridSQL =  GridSQL & " order by " & request("GRIDORDERF") & " " & request("GRIDORDER")
				
				end if
				
			end if
			'//-----------------------------------------------------------------------------
		
		'-----------------------------------------------------------------------------
		
		'response.write GridSQL
		response.write GridASP(GridSQL,db_oledb,GridR)
	'//--------------------------------------------------------------------------------->	
	

%>