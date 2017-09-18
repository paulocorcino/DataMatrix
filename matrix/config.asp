<!--#include virtual="/bymidia/matrix/config/db.asp" -->

<%


	sConnStats = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & db_oledb
	
	set conn = Server.CreateObject("ADODB.Connection")
	set rs = Server.CreateObject("ADODB.Recordset")

	'ADO Constants
	'---- CursorTypeEnum Values ----
	Const adOpenForwardOnly = 0
	Const adOpenKeyset = 1
	Const adOpenDynamic = 2
	Const adOpenStatic = 3

	'---- CursorLocationEnum Values ----
	Const adUseServer = 2
	Const adUseClient = 3

	'---- CommandTypeEnum Values ----
	Const adCmdUnknown = &H0008
	Const adCmdText = &H0001
	Const adCmdTable = &H0002
	Const adCmdStoredProc = &H0004
	Const adCmdFile = &H0100
	Const adCmdTableDirect = &H0200
	
	sub OpenDB(sConn)
		'Opens the given connection and initializes the recordset
		conn.open sConn
		set rs.ActiveConnection = conn
		rs.CursorType = adOpenStatic
	end sub
	
	sub CloseDB()
		'Closes and destroys the connection and recordset objects
		rs.close
		conn.close
		set rs = nothing
		set conn = nothing
	end sub
	
	sub w(sText)
		'A Quickie ;)
		response.write sText & vbCrLf
	end sub

	'Load Config from DB
	
	'Open Connection
	conn.open sConnStats
		
	'Build SqlString
	strSql = "SELECT C_ImageLoc "
	strSql = strSql & ", C_FilterIP "
	strSql = strSql & ", C_ShowLinks "
	strSql = strSql & ", C_RefThisServer "
	strSql = strSql & ", C_StripPathParameters "
	strSql = strSql & ", C_StripPathProtocol "
	strSql = strSql & ", C_StripRefParameters "
	strSql = strSql & ", C_StripRefProtocol "
	strSql = strSql & ", C_StripRefFile "
	strSql = strSql & "FROM Config WHERE ID = 1"
				
	'Open RecordSet
	set rs = conn.Execute(strSql)
		
	'Get Variables
	sImageLocation			= rs.Fields("C_ImageLoc")
	sFilterIps				= rs.Fields("C_FilterIP")
	bShowLinks				= rs.Fields("C_ShowLinks")
	bRefThisServer			= rs.Fields("C_RefThisServer")
	bStripPathParameters	= rs.Fields("C_StripPathParameters")
	bStripPathProtocol		= rs.Fields("C_StripPathProtocol")
	bStripRefParameters 	= rs.Fields("C_StripRefParameters")
	bStripRefProtocol		= rs.Fields("C_StripRefProtocol")
	bStripRefFile			= rs.Fields("C_StripRefFile")
	
	
	'Terminate database connection
	rs.Close
	conn.close

%>