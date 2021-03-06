<!--#include virtual="header_conn_info.asp"-->	
<%
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open ConnectionString

	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = Conn
	
	SQLstmt = "ALTER SESSION SET NLS_DATE_FORMAT = 'MM/DD/YYYY'"
	Set RSSESSION = Conn.Execute(SQLstmt)
	Set RSSESSION = Nothing 
	
	SQLstmt = "WITH PULSE_DATA AS " & _
	"( " & _
		"SELECT " & _
		"CONNECT_BY_ROOT(OPS_USD_OPS_USR_ID) PULSE_USR_ID, " & _
		"CONNECT_BY_ROOT(OPS_USD_TYPE) ROOT_DEPARTMENT, " & _
		"CONNECT_BY_ROOT(OPS_USD_TEAM) ROOT_TEAM, " & _
		"CASE " & _
			"WHEN CONNECT_BY_ROOT(OPS_USD_LOCATION) IN ('MOT','WFD','WFH') AND CONNECT_BY_ROOT(OPS_USD_JOB) = 'ADM' AND CONNECT_BY_ROOT(OPS_USD_TYPE) = 'OPS' THEN 6 " & _
			"WHEN CONNECT_BY_ROOT(OPS_USD_LOCATION) IN ('MOT','WFD','WFH') AND CONNECT_BY_ROOT(OPS_USD_JOB) = 'ANL' AND CONNECT_BY_ROOT(OPS_USD_TYPE) = 'OPS' THEN 5 " & _
			"WHEN CONNECT_BY_ROOT(OPS_USD_LOCATION) IN ('MOT','WFD','WFH') AND CONNECT_BY_ROOT(OPS_USD_JOB) = 'DIR' OR (CONNECT_BY_ROOT(OPS_USD_JOB) IN ('ADM','ANL') AND CONNECT_BY_ROOT(OPS_USD_TYPE) = 'HRA') THEN 4 " & _
			"WHEN CONNECT_BY_ROOT(OPS_USD_LOCATION) IN ('MOT','WFD','WFH') AND CONNECT_BY_ROOT(OPS_USD_JOB) = 'MGR' THEN 3 " & _
			"WHEN CONNECT_BY_ROOT(OPS_USD_LOCATION) IN ('MOT','WFD','WFH') AND CONNECT_BY_ROOT(OPS_USD_JOB) = 'SUP' THEN 2 " & _
			"WHEN CONNECT_BY_ROOT(OPS_USD_LOCATION) IN ('MOT','WFD','WFH') AND CONNECT_BY_ROOT(OPS_USD_JOB) = 'LED' THEN 1 " & _
			"ELSE 0 " & _
		"END PULSE_SECURITY, " & _
		"TO_DATE(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE)) - MOD(TO_DATE(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE))-TO_DATE('5/5/2019','MM/DD/YYYY'),14) PULSE_PAYPERIOD_START, " & _
		"OPS_USD_TYPE, " & _
		"OPS_USD_LOCATION, " & _
		"OPS_USD_CLASS, " & _
		"OPS_USD_PAY_RATE " & _
		"FROM OPS_USER " & _
		"JOIN OPS_USER_DETAIL " & _
		"ON OPS_USR_ID = OPS_USD_OPS_USR_ID " & _
		"AND TO_DATE(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE)) BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE " & _
		"WHERE CONNECT_BY_ISLEAF = 1 " & _
		"START WITH UPPER(OPS_USR_NT_ID) = UPPER(?) " & _
 		"CONNECT BY OPS_USD_SUPERVISOR = PRIOR OPS_USD_OPS_USR_ID " & _
	") " & _
	"SELECT DISTINCT " & _
	"PULSE_USR_ID, " & _
	"CASE " & _
		"WHEN PULSE_SECURITY >= 5 THEN 'RES' " & _
		"WHEN PULSE_SECURITY = 4 THEN 'ACC,CRT,DOC,GRP,OPS,OSS,POP,RES' " & _
		"ELSE PULSE_DEPARTMENT " & _
	"END PULSE_DEPARTMENT, " & _
	"PULSE_SECURITY, " & _
	"PULSE_PAYPERIOD_START " & _
	"FROM PULSE_DATA " & _
	"CROSS JOIN " & _
	"( " & _
		"SELECT LISTAGG(OPS_USD_TYPE,',') WITHIN GROUP (ORDER BY OPS_USD_TYPE) PULSE_DEPARTMENT " & _
		"FROM " & _
		"( " & _
			"SELECT OPS_USD_TYPE " & _
			"FROM PULSE_DATA " & _
			"WHERE OPS_USD_LOCATION IN ('MOT','WFD','WFH') " & _
			"AND OPS_USD_CLASS IN ('RGFT','RGPT','PT<30','LEAVE') " & _
			"AND OPS_USD_PAY_RATE > 0 " & _
			"UNION " & _
			"SELECT CASE WHEN PULSE_SECURITY >= 2 AND ROOT_DEPARTMENT = 'RES' AND ROOT_TEAM = 'SPT' THEN 'CRT' ELSE ROOT_DEPARTMENT END FROM PULSE_DATA " & _
		") " & _
	")"
	cmd.CommandText = SQLstmt
	cmd.Parameters(0).value = Request.ServerVariables("LOGON_USER")
	Set RSPULSE = cmd.Execute
	IF Not RSPULSE.EOF Then
		PULSE_USR_ID = RSPULSE("PULSE_USR_ID")
		PULSE_DEPARTMENT = RSPULSE("PULSE_DEPARTMENT")
		PULSE_SECURITY = CInt(RSPULSE("PULSE_SECURITY"))
		PULSE_PAYPERIOD_START = CDate(RSPULSE("PULSE_PAYPERIOD_START"))
	Else
		PULSE_USR_ID = "-1"
		PULSE_DEPARTMENT = "RES"
		PULSE_SECURITY = 0
		PULSE_PAYPERIOD_START = Date
	End If
	If PULSE_USR_ID = "9861" Then
		'PULSE_USR_ID = "6326"
		'PULSE_DEPARTMENT = "OSS"
		'PULSE_SECURITY = 3
	End If
	Set RSPULSE = Nothing
%>