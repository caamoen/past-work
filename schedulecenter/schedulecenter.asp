<!--#include virtual="header.asp"-->
<script type="text/javascript" src="/stickytable.js"></script>
<link rel="stylesheet" type="text/css" href="/scheduleCenterCSS.css?v=2.11">



<%
	Server.ScriptTimeout = 6000

	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = Conn

	If Request.Form("USE_LAYOUT") <> "" Then
		Response.Cookies("SCHEDSQUATCH_LAYOUT") = Request.Form("USE_LAYOUT")
		Response.Cookies("SCHEDSQUATCH_LAYOUT").Expires = DateAdd("y",1,Date)
	End If
	If Request.Form("STPLUS_ENABLED") <> "" Then
		Response.Cookies("STPLUS_ENABLED") = Request.Form("STPLUS_ENABLED")
		Response.Cookies("STPLUS_ENABLED").Expires = DateAdd("y",1,Date)
	End If
	If Request.Form("SUBMIT_FLAG") = "0" Then
		Session("FILTER_AGENT") = CInt(Request.Form("FILTER_AGENT"))
		Session("FILTER_DATE") = CDate(Request.Form("FILTER_DATE"))
		Response.Redirect(Request.ServerVariables("SCRIPT_NAME"))
		Response.End
	End If

	If Request.Cookies("SCHEDSQUATCH_LAYOUT") <> "" Then
		SCHEDSQUATCH_LAYOUT = Request.Cookies("SCHEDSQUATCH_LAYOUT")
	Else
		SCHEDSQUATCH_LAYOUT = "V"
	End If

	If Request.Cookies("STPLUS_ENABLED") <> "" Then 
		STPLUS_ENABLED = Request.Cookies("STPLUS_ENABLED")
	Else
		STPLUS_ENABLED = "N"
	End If

	If Request.Form("FILTER_AGENT") <> "" Then
		SCHEDULE_USR_ID = CInt(Request.Form("FILTER_AGENT"))
	Elseif Session.Contents("FILTER_AGENT") <> "" Then
		SCHEDULE_USR_ID = CInt(Session.Contents("FILTER_AGENT"))
	Else
		SCHEDULE_USR_ID = OPS_USR_ID
	End If
	Session.Contents.Remove("FILTER_AGENT")

	SQLstmt = "SELECT " & _
	"AGENT_NAME, " & _
	"AGENT_TYPE, " & _
	"AGENT_DEPT, " & _
	"AGENT_TEAM, " & _
	"AGENT_JOB, " & _
	"AGENT_HOURS, " & _
	"AGENT_ROUTING, " & _
	"MIN(PICK_START_DATE) PICK_START_DATE, " & _
	"MAX(PICK_END_DATE) PICK_END_DATE, " & _
	"NVL(STPLUS.OPS_PAR_VALUE,'N') STPLUS_AVAILABLE " & _
	"FROM " & _
	"( " & _
		"SELECT " & _
		"OPS_USR_NAME AGENT_NAME, " & _
		"CASE " & _
			"WHEN OPS_USD_TYPE = 'RES' OR (OPS_USD_TYPE = 'GRP' AND OPS_USD_TEAM = 'SPT') THEN CASE WHEN TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) <= TO_DATE('12/2/2018','MM/DD/YYYY') THEN DECODE(RES_RTE_RES_RTG_ID,1,'SLS',2,'SLS',3,'SLS',4,'SPT',5,'SPT',10,'SRV',13,'OSR','RES') ELSE DECODE(OPS_USD_TEAM,'SRV','SPT','SPT','SPT','OSR','OSR','RES') END " & _
			"ELSE OPS_USD_TYPE " & _
		"END AGENT_TYPE, " & _
		"CASE WHEN OPS_USD_TYPE = 'GRP' AND OPS_USD_TEAM = 'SPT' THEN 'RES' ELSE OPS_USD_TYPE END AGENT_DEPT, " & _
		"OPS_USD_TEAM AGENT_TEAM, " & _
		"OPS_USD_JOB AGENT_JOB, " & _
		"NVL(OPS_USD_SCH_HOURS,0) AGENT_HOURS, " & _
		"NVL(DECODE(RES_RTE_RES_RTG_ID,3,1,2,1,RES_RTE_RES_RTG_ID),0) AGENT_ROUTING, " & _
		"COALESCE(NVL2(RES_SCM_ID,TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)),NULL),GREATEST(TO_DATE(REGEXP_SUBSTR(SYS_CDD_VALUE,'[^_]+',1,1),'MM/DD/YYYY'),TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE))),TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)-1)) PICK_START_DATE, " & _
		"COALESCE(NVL2(RES_SCM_ID,TO_DATE('12/31/2040','MM/DD/YYYY'),NULL),TO_DATE(REGEXP_SUBSTR(SYS_CDD_VALUE,'[^_]+',1,2),'MM/DD/YYYY'),TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)-1)) PICK_END_DATE " & _
		"FROM OPS_USER_DETAIL " & _
		"JOIN OPS_USER " & _
		"ON OPS_USD_OPS_USR_ID = OPS_USR_ID " & _
		"LEFT JOIN RES_ROUTING " & _
		"ON RES_RTE_OPS_USR_ID = OPS_USD_OPS_USR_ID " & _
		"AND RES_RTE_YEAR = TO_CHAR(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)-(6/24),'YYYY') " & _
		"AND RES_RTE_MONTH = TO_CHAR(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)-(6/24),'MM') " & _
		"LEFT JOIN RES_SCHEDULE_MASTER " & _
		"ON TO_CHAR(OPS_USD_OPS_USR_ID) = RES_SCM_NAME " & _
		"AND RES_SCM_STATUS = 'ACT' " & _
		"AND TRIM(REPLACE(RES_SCM_TYPE,'NA',0)) > 0 " & _
		"LEFT JOIN SYS_CODE_DETAIL " & _
		"ON SYS_CDD_SYS_CDM_ID = 459 " & _
		"AND OPS_USD_OPS_USR_ID = SYS_CDD_NAME " & _
		"AND TO_DATE(REGEXP_SUBSTR(SYS_CDD_VALUE,'[^_]+',1,2),'MM/DD/YYYY') >= TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) " & _
		"AND REGEXP_SUBSTR(SYS_CDD_VALUE,'[^_]+',1,4) > 0 " & _
		"WHERE TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE " & _
		"AND OPS_USD_OPS_USR_ID = ? " & _
	") " & _
    "JOIN OPS_PARAMETER PICK " & _
    "ON AGENT_TYPE = PICK.OPS_PAR_PARENT_TYPE " & _
    "AND PICK.OPS_PAR_CODE = 'PICK_WINDOW' " & _
    "AND TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) BETWEEN PICK.OPS_PAR_EFF_DATE AND PICK.OPS_PAR_DIS_DATE " & _
    "LEFT JOIN OPS_PARAMETER STPLUS " & _
    "ON AGENT_TYPE = STPLUS.OPS_PAR_PARENT_TYPE " & _
    "AND STPLUS.OPS_PAR_CODE = 'SELFTRADE_PLUS' " & _
    "AND TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) BETWEEN STPLUS.OPS_PAR_EFF_DATE AND STPLUS.OPS_PAR_DIS_DATE " & _
    "GROUP BY AGENT_NAME,AGENT_TYPE,AGENT_DEPT,AGENT_TEAM,AGENT_JOB,AGENT_HOURS,AGENT_ROUTING,NVL(STPLUS.OPS_PAR_VALUE,'N') "
	cmd.CommandText = SQLstmt
	cmd.Parameters(0).value = SCHEDULE_USR_ID
	Set RSAGT = cmd.Execute
	If Not RSAGT.EOF Then
		AGENT_NAME = RSAGT("AGENT_NAME")
		AGENT_TYPE = RSAGT("AGENT_TYPE")
		AGENT_DEPT = RSAGT("AGENT_DEPT")
		AGENT_TEAM = RSAGT("AGENT_TEAM")
		AGENT_JOB = RSAGT("AGENT_JOB")
		AGENT_HOURS  = CDbl(RSAGT("AGENT_HOURS"))
		AGENT_ROUTING = CDbl(RSAGT("AGENT_ROUTING"))
		PICK_START_DATE = CDate(RSAGT("PICK_START_DATE"))
		PICK_END_DATE = CDate(RSAGT("PICK_END_DATE"))
		STPLUS_AVAILABLE = RSAGT("STPLUS_AVAILABLE")
	Else
		AGENT_NAME = OPS_USR_NAME
		AGENT_TYPE = OPS_USR_TYPE
		AGENT_DEPT = OPS_USR_TYPE
		AGENT_TEAM = OPS_USR_TEAM
		AGENT_JOB = OPS_USR_JOB
		AGENT_HOURS = 0
		AGENT_ROUTING = 0
		PICK_START_DATE = Date-1
		PICK_END_DATE = Date-1
		STPLUS_AVAILABLE = "N"
	End If
	Set RSAGT = Nothing

	SQLstmt = "SELECT " & _
	"SCHEDULE_START_DATE - TO_CHAR(SCHEDULE_START_DATE,'D') + 1 SCHEDULE_START_DATE, " & _
	"HOUR_DELAY " & _
	"FROM " & _
	"( " & _
		"SELECT " & _
        "LEAST(NVL(MAX(DECODE(OPS_PAR_CODE,'DROP_WINDOW',CASE " & _
			"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),':') > 0 THEN TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24) - TO_CHAR(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24,'D') + 7*REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'[^:]+',1,1) + REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'[^:]+',1,2) " & _
			"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'/') > 0 THEN TO_DATE(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'MM/DD/YYYY') " & _
			"ELSE TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24) + REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1) " & _
        "END)),TO_DATE('12/31/2040','MM/DD/YYYY')), " & _
        "NVL(MAX(DECODE(OPS_PAR_CODE,'ADD_WINDOW',CASE " & _
			"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),':') > 0 THEN TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24) - TO_CHAR(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24,'D') + 7*REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'[^:]+',1,1) + REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'[^:]+',1,2) " & _
			"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'/') > 0 THEN TO_DATE(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'MM/DD/YYYY') " & _
			"ELSE TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24) + REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1) " & _
        "END)),TO_DATE('12/31/2040','MM/DD/YYYY')), " & _
        "NVL(MAX(DECODE(OPS_PAR_CODE,'SELFTRADE_WINDOW',CASE " & _
			"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),':') > 0 THEN TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24) - TO_CHAR(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24,'D') + 7*REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'[^:]+',1,1) + REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'[^:]+',1,2) " & _
			"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'/') > 0 THEN TO_DATE(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'MM/DD/YYYY') " & _
			"ELSE TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24) + REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1) " & _
        "END)),TO_DATE('12/31/2040','MM/DD/YYYY')), " & _
        "NVL(MAX(DECODE(OPS_PAR_CODE,'PICK_WINDOW',CASE " & _
			"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),':') > 0 THEN TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24) - TO_CHAR(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24,'D') + 7*REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'[^:]+',1,1) + REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'[^:]+',1,2) " & _
			"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'/') > 0 THEN TO_DATE(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'MM/DD/YYYY') " & _
			"ELSE TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24) + REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1) " & _
        "END)),TO_DATE('12/31/2040','MM/DD/YYYY'))) SCHEDULE_START_DATE, " & _
		"NVL(MAX(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3)),0) HOUR_DELAY " & _
		"FROM OPS_PARAMETER PAR " & _
		"WHERE OPS_PAR_CODE IN ('DROP_WINDOW','ADD_WINDOW','SELFTRADE_WINDOW','PICK_WINDOW') " & _
		"AND TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) BETWEEN OPS_PAR_EFF_DATE AND OPS_PAR_DIS_DATE " & _
		"AND OPS_PAR_PARENT_TYPE = ? " & _
	")"
	cmd.CommandText = SQLstmt
	cmd.Parameters(0).value = AGENT_TYPE
	Set RSSTART = cmd.Execute
	If CDate(Request.Form("FILTER_DATE")) >= CDate(RSSTART("SCHEDULE_START_DATE")) or Request.Form("SUBMIT_FLAG") = "1" Then
		SCHEDULE_START_DATE = CDate(Request.Form("FILTER_DATE"))
	Elseif CDate(Session.Contents("FILTER_DATE")) >= CDate(RSSTART("SCHEDULE_START_DATE")) Then
		SCHEDULE_START_DATE = CDate(Session.Contents("FILTER_DATE"))
	Else
		SCHEDULE_START_DATE = CDate(RSSTART("SCHEDULE_START_DATE"))
	End If
	Session.Contents.Remove("FILTER_DATE")
	HOUR_DELAY = RSSTART("HOUR_DELAY")
	SCHEDULE_END_DATE = SCHEDULE_START_DATE + 6
	Set RSSTART = Nothing

	If Request.Form("SUBMIT_FLAG") = "1" Then
		SCHEDSQLstmt = ""
		For Each FIELD in Request.Form
			If Instr(FIELD,"SCHEDDATA") > 0 and Request.Form(FIELD) <> "" Then
				SCHEDDATA_ARRAY = Split(FIELD,"_")

				SCHEDSQLstmt = SCHEDSQLstmt & "SELECT TO_DATE('" & SCHEDULE_START_DATE + SCHEDDATA_ARRAY(1) & "','MM/DD/YYYY') SS_DATE, " & _
				"TO_CHAR(TO_DATE('" & SCHEDDATA_ARRAY(2) & "','HH24MI'),'HH24:MI') SS_START, " & _
				"TO_CHAR(TO_DATE('" & SCHEDDATA_ARRAY(2) & "','HH24MI') + 29/1440,'HH24:MI') SS_END, " & _
				"'" & Request.Form(FIELD) & "' SS_TYPE " & _
				"FROM DUAL " & _
				"UNION ALL "
			End If
		Next
		If SCHEDSQLstmt <> "" Then
			SCHEDSQLstmt = Left(SCHEDSQLstmt,Len(SCHEDSQLstmt)-11)

			SQLstmt = "SELECT MIN(SS_DATE) SCHEDSQUATCH_START, " & _
			"MAX(SS_DATE) SCHEDSQUATCH_END " & _
			"FROM " & _
			"( " & _
				SCHEDSQLstmt & _
			") " & _
			"HAVING MIN(SS_DATE) > TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)-((?+(5/60))/24))"
			cmd.CommandText = SQLstmt
			cmd.Parameters(0).value = HOUR_DELAY
			Set RSDATE = cmd.Execute
			If Not RSDATE.EOF Then
				SCHEDSQUATCH_START = RSDATE("SCHEDSQUATCH_START")
				SCHEDSQUATCH_END = RSDATE("SCHEDSQUATCH_END")
			Else
				Session("FILTER_AGENT") = SCHEDULE_USR_ID
				Session("FILTER_DATE") = SCHEDULE_START_DATE
				Response.Redirect(Request.ServerVariables("SCRIPT_NAME"))
				Response.End
			End If
			Set RSDATE = Nothing

			SQLstmt = "SELECT OPS_SCI_ID " & _
			"FROM OPS_SCHEDULE_INFO " & _
			"WHERE TO_DATE(OPS_SCI_START) BETWEEN TO_DATE(?,'MM/DD/YYYY') AND TO_DATE(?,'MM/DD/YYYY') " & _
			"AND OPS_SCI_OPS_USR_ID = ? " & _
			"AND OPS_SCI_STATUS = 'APP' " & _
			"AND OPS_SCI_TYPE NOT LIKE 'HOL_'"
			cmd.CommandText = SQLstmt
			cmd.Parameters(0).value = SCHEDSQUATCH_START
			cmd.Parameters(1).value = SCHEDSQUATCH_END
			cmd.Parameters(2).value = SCHEDULE_USR_ID
			Set RSIDS = cmd.Execute

			SQLstmt = "WITH SS_DATA " & _
			"AS " & _
			"( " & _
				"SELECT DISTINCT " & _
				"SS_DATE, " & _
				"SS_INTERVAL, " & _
				"SUBSTR(SS_TYPE,1,2) SS_SYSTEM, " & _
				"SUBSTR(SS_TYPE,3) SS_TYPE " & _
				"FROM " & _
				"( " & _
					SCHEDSQLstmt & _
				") " & _
				"JOIN " & _
				"( " & _
					"SELECT TO_CHAR(TO_DATE('00:00','HH24:MI') + (ROWNUM - 1) / 240,'HH24:MI') SS_INTERVAL " & _
					"FROM DUAL " & _
					"CONNECT BY ROWNUM <= 240 " & _
				") " & _
				"ON SS_INTERVAL BETWEEN SS_START AND SS_END " & _
			") " & _
			"SELECT " & _
			"? OPS_SCI_OPS_USR_ID, " & _
			"'APP' OPS_SCI_STATUS, " & _
			"CONNECT_BY_ROOT(USE_START) OPS_SCI_START, " & _
			"USE_END OPS_SCI_END, " & _
			"USE_TYPE OPS_SCI_TYPE, " & _
			"? OPS_SCI_OPS_USR_TYPE, " & _
			"CONNECT_BY_ROOT(USE_NOTES) OPS_SCI_NOTES " & _
			"FROM " & _
			"( " & _
				"SELECT " & _
				"USE_START, " & _
				"USE_END, " & _
				"USE_TYPE, " & _
				"CASE " & _
					"WHEN LAG(USE_END) OVER (PARTITION BY USE_TYPE, TO_DATE(USE_START) ORDER BY USE_START) IS NULL " & _
					"OR LAG(USE_END) OVER (PARTITION BY USE_TYPE, TO_DATE(USE_START) ORDER BY USE_START) <> USE_START " & _
					"OR LAG(USE_TYPE) OVER (PARTITION BY USE_TYPE, TO_DATE(USE_START) ORDER BY USE_START) <> USE_TYPE THEN 1 " & _
				"END START_FLAG, " & _
				"USE_NOTES " & _
				"FROM " & _
				"( " & _
					"SELECT " & _
					"TO_DATE(USE_DATE || ' ' || USE_INTERVAL,'YYYY-MM-DD HH24:MI') USE_START, " & _
					"TO_DATE(USE_DATE || ' ' || USE_INTERVAL,'YYYY-MM-DD HH24:MI')+1/240 USE_END, " & _
					"CASE " & _
						"WHEN APP.SS_TYPE = 'ADDT' AND SCI_TYPE IS NULL THEN APP.SS_TYPE " & _
						"WHEN APP.SS_TYPE IN ('SRPT','SRUN') AND SCI_TYPE IN ('BASE','PICK','HOLW','ADDT','EXTD') THEN APP.SS_TYPE " & _
						"WHEN APP.SS_TYPE = 'LNCH' AND APP.SS_SYSTEM = 'AD' AND SCI_TYPE IS NULL THEN APP.SS_TYPE " & _
						"WHEN APP.SS_TYPE = 'LNCH' AND APP.SS_SYSTEM = 'ST' AND SCI_TYPE IS NULL THEN APP.SS_TYPE " & _
						"WHEN APP.SS_TYPE = 'STLNCH' AND APP.SS_SYSTEM = 'ST' AND (SCI_TYPE IS NULL OR SCI_TYPE IN ('BASE','PICK','HOLW','ADDT','EXTD')) THEN LTRIM(APP.SS_TYPE,'ST') " & _
						"WHEN APP.SS_TYPE = 'OFF' AND APP.SS_SYSTEM = 'ST' AND SCI_TYPE IN ('BASE','PICK','HOLW','ADDT','EXTD') THEN NULL " & _
						"WHEN APP.SS_TYPE = 'BASE' AND APP.SS_SYSTEM = 'ST' AND SCI_TYPE IS NULL THEN APP.SS_TYPE " & _
						"WHEN APP.SS_SYSTEM = 'PK' AND (SCI_TYPE IS NULL OR SCI_TYPE IN ('LNCH','PICK')) THEN NULLIF(APP.SS_TYPE,'OFF') " & _
						"ELSE SCI_TYPE " & _
					"END USE_TYPE, " & _
					"CASE " & _
						"WHEN APP.SS_TYPE = 'ADDT' AND SCI_TYPE IS NULL THEN 'Time Added in Schedule Center' " & _
						"WHEN APP.SS_TYPE IN ('SRPT','SRUN') AND SCI_TYPE IN ('BASE','PICK','HOLW','ADDT','EXTD') THEN 'Time Dropped in Schedule Center' " & _
						"WHEN APP.SS_TYPE = 'LNCH' AND APP.SS_SYSTEM = 'AD' AND SCI_TYPE IS NULL THEN 'Lunch Added in Schedule Center' " & _
						"WHEN APP.SS_TYPE = 'LNCH' AND APP.SS_SYSTEM = 'ST' AND SCI_TYPE IS NULL THEN 'Lunch Added in Schedule Center' " & _
						"WHEN APP.SS_TYPE = 'STLNCH' AND APP.SS_SYSTEM = 'ST' AND (SCI_TYPE IS NULL OR SCI_TYPE IN ('BASE','PICK','HOLW','ADDT','EXTD')) THEN '(Self-Trade) Lunch Added in Schedule Center' " & _
						"WHEN APP.SS_TYPE = 'OFF' AND APP.SS_SYSTEM = 'ST' AND SCI_TYPE IN ('BASE','PICK','HOLW','ADDT','EXTD') THEN NULL " & _
						"WHEN APP.SS_TYPE = 'BASE' AND APP.SS_SYSTEM = 'ST' AND SCI_TYPE IS NULL THEN 'Self-Trade in Schedule Center' " & _
						"WHEN APP.SS_SYSTEM = 'PK' AND (SCI_TYPE IS NULL OR SCI_TYPE IN ('LNCH','PICK')) THEN DECODE(APP.SS_TYPE,'OFF',NULL,'Hours Picked in Schedule Center') " & _
						"ELSE SCI_NOTES " & _
					"END USE_NOTES " & _
					"FROM " & _
					"( " & _
						"SELECT " & _
						"USE_DATE, " & _
						"USE_INTERVAL " & _
						"FROM " & _
						"( " & _
							"SELECT TO_DATE(?,'MM/DD/YYYY') + (ROWNUM - 1) USE_DATE " & _
							"FROM DUAL " & _
							"CONNECT BY ROWNUM < TO_DATE(?,'MM/DD/YYYY') - TO_DATE(?,'MM/DD/YYYY') + 2 " & _
						") " & _
						"CROSS JOIN " & _
						"( " & _
							"SELECT TO_CHAR(TO_DATE('00:00','HH24:MI') + (ROWNUM - 1) / 240,'HH24:MI') USE_INTERVAL " & _
							"FROM DUAL " & _
							"CONNECT BY ROWNUM <= 240 " & _
						") " & _
					") " & _
					"LEFT JOIN " & _
					"( " & _
						"SELECT " & _
						"TO_DATE(OPS_SCI_START) SCI_DATE, " & _
						"SCI_INTERVAL, " & _
						"MAX(OPS_SCI_TYPE) KEEP (DENSE_RANK LAST ORDER BY INSERT_DATE, OPS_SCI_TYPE) SCI_TYPE, " & _
						"MAX(OPS_SCI_NOTES) KEEP (DENSE_RANK LAST ORDER BY INSERT_DATE, OPS_SCI_TYPE) SCI_NOTES " & _
						"FROM " & _
						"( " & _
							"SELECT OPS_SCI_START, " & _
							"OPS_SCI_END, " & _
							"OPS_SCI_TYPE, " & _
							"OPS_SCI_NOTES, " & _
							"INSERT_DATE " & _
							"FROM OPS_SCHEDULE_INFO " & _
							"WHERE TO_DATE(OPS_SCI_START) BETWEEN TO_DATE(?,'MM/DD/YYYY') AND TO_DATE(?,'MM/DD/YYYY') " & _
							"AND OPS_SCI_OPS_USR_ID = ? " & _
							"AND OPS_SCI_STATUS = 'APP' " & _
							"AND OPS_SCI_TYPE NOT LIKE 'HOL_' " & _
						") " & _
						"JOIN " & _
						"( " & _
							"SELECT TO_CHAR(TO_DATE('00:00','HH24:MI') + (ROWNUM - 1) / 240,'HH24:MI') SCI_INTERVAL " & _
							"FROM DUAL " & _
							"CONNECT BY ROWNUM <= 240 " & _
						") " & _
						"ON SCI_INTERVAL BETWEEN TO_CHAR(OPS_SCI_START,'HH24:MI') AND TO_CHAR(OPS_SCI_END-(1/1440),'HH24:MI') " & _
						"GROUP BY TO_DATE(OPS_SCI_START), SCI_INTERVAL " & _
					") " & _
					"ON USE_DATE = SCI_DATE " & _
					"AND USE_INTERVAL = SCI_INTERVAL " & _
					"LEFT JOIN " & _
					"( " & _
						"SELECT " & _
						"SS_DATE, " & _
						"SS_INTERVAL, " & _
						"SS_SYSTEM, " & _
						"DECODE(SS_TYPE,'STLNCH','STLNCH',REGEXP_REPLACE(SS_TYPE,'ST(.*)','\1')) SS_TYPE " & _
						"FROM SS_DATA " & _
					") APP " & _
					"ON USE_DATE = APP.SS_DATE " & _
					"AND USE_INTERVAL = APP.SS_INTERVAL " & _
				") " & _
				"WHERE USE_TYPE IS NOT NULL " & _
			") " & _
			"WHERE CONNECT_BY_ISLEAF = 1 " & _
			"START WITH START_FLAG = 1 " & _
			"CONNECT BY TO_DATE(USE_START) = PRIOR TO_DATE(USE_START) " & _
			"AND USE_START = PRIOR USE_END " & _
			"AND USE_TYPE = PRIOR USE_TYPE " & _
			"UNION ALL " & _
			"SELECT " & _
			"?, " & _
			"'COM', " & _
			"CONNECT_BY_ROOT(USE_START), " & _
			"USE_END, " & _
			"USE_TYPE, " & _
			"?, " & _
			"'' " & _
			"FROM " & _
			"( " & _
				"SELECT " & _
				"USE_START, " & _
				"USE_END, " & _
				"USE_TYPE, " & _
				"CASE " & _
					"WHEN LAG(USE_END) OVER (PARTITION BY USE_TYPE, TO_DATE(USE_START) ORDER BY USE_START) IS NULL " & _
					"OR LAG(USE_END) OVER (PARTITION BY USE_TYPE, TO_DATE(USE_START) ORDER BY USE_START) <> USE_START " & _
					"OR LAG(USE_TYPE) OVER (PARTITION BY USE_TYPE, TO_DATE(USE_START) ORDER BY USE_START) <> USE_TYPE THEN 1 " & _
				"END START_FLAG " & _
				"FROM " & _
				"( " & _
					"SELECT " & _
					"TO_DATE(SS_DATE || ' ' || SS_INTERVAL,'YYYY-MM-DD HH24:MI') USE_START, " & _
					"TO_DATE(SS_DATE || ' ' || SS_INTERVAL,'YYYY-MM-DD HH24:MI')+1/240 USE_END, " & _
					"DECODE(SS_TYPE,'STBASE','SLTU','SLTD') USE_TYPE " & _
					"FROM SS_DATA " & _
					"WHERE SUBSTR(SS_TYPE,1,2) = 'ST' " & _
				") SLT " & _
			") " & _
			"WHERE CONNECT_BY_ISLEAF = 1 " & _
			"START WITH START_FLAG = 1 " & _
			"CONNECT BY TO_DATE(USE_START) = PRIOR TO_DATE(USE_START) " & _
			"AND USE_START = PRIOR USE_END " & _
			"AND USE_TYPE = PRIOR USE_TYPE"
			cmd.CommandText = SQLstmt
			cmd.Parameters(0).value = SCHEDULE_USR_ID
			cmd.Parameters(1).value = AGENT_DEPT
			cmd.Parameters(2).value = SCHEDSQUATCH_START
			cmd.Parameters(3).value = SCHEDSQUATCH_END
			cmd.Parameters(4).value = SCHEDSQUATCH_START
			cmd.Parameters(5).value = SCHEDSQUATCH_START
			cmd.Parameters(6).value = SCHEDSQUATCH_END
			cmd.Parameters(7).value = SCHEDULE_USR_ID
			cmd.Parameters(8).value = SCHEDULE_USR_ID
			cmd.Parameters(9).value = AGENT_DEPT
			Set RSSSIMMORTAL = cmd.Execute

			If Not RSSSIMMORTAL.EOF Then
				SQLstmt = "INSERT ALL "
				Do While Not RSSSIMMORTAL.EOF
					SQLstmt = SQLstmt & "INTO OPS_SCHEDULE_INFO(OPS_SCI_OPS_USR_ID, OPS_SCI_STATUS, OPS_SCI_START, OPS_SCI_END, OPS_SCI_TYPE, OPS_SCI_OPS_USR_TYPE, OPS_SCI_NOTES, OPS_SCI_INS_USER, INSERT_DATE) VALUES (" & RSSSIMMORTAL("OPS_SCI_OPS_USR_ID") & ",'" & RSSSIMMORTAL("OPS_SCI_STATUS") & "',TO_DATE('" & RSSSIMMORTAL("OPS_SCI_START") & "','MM/DD/YYYY HH:MI:SS AM'),TO_DATE('" & RSSSIMMORTAL("OPS_SCI_END") & "','MM/DD/YYYY HH:MI:SS AM'),'" & RSSSIMMORTAL("OPS_SCI_TYPE") & "','" & RSSSIMMORTAL("OPS_SCI_OPS_USR_TYPE") & "','" & RSSSIMMORTAL("OPS_SCI_NOTES") & "'," & OPS_USR_ID & ", CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) "
					RSSSIMMORTAL.MoveNext
				Loop
				SQLstmt = SQLstmt & "SELECT * FROM DUAL"
				cmd.CommandText = SQLstmt
				Set RSI = cmd.Execute
				Set RSI = Nothing
			End If
			Set RSSSIMMORTAL = Nothing

			Do While Not RSIDS.EOF
				SQLstmt = "UPDATE OPS_SCHEDULE_INFO " & _
				"SET OPS_SCI_STATUS = 'DEL' " & _
				"WHERE OPS_SCI_ID = ?"
				cmd.CommandText = SQLstmt
				cmd.Parameters(0).value = RSIDS("OPS_SCI_ID")
				Set RSD = cmd.Execute
				Set RSD = Nothing

				RSIDS.MoveNext
			Loop
			Set RSIDS = Nothing

			'If AGENT_TYPE = "SLS" or AGENT_TYPE = "SRV" or AGENT_TYPE = "RES" or AGENT_TYPE = "SPT" or AGENT_TYPE = "OSR" Then
			'	Call UPDATE_OPS_SCHEDULE_NEED(AGENT_TYPE, SCHEDSQUATCH_START, SCHEDSQUATCH_END)
			'End If
		End If
		If Request.Form("LUNCH_WAIVER") <> "" Then
			MergeSQLstmt = ""
			WAIVER_LIST = Split(Request.Form("LUNCH_WAIVER"),";")
			If AGENT_TYPE = "SPT" or AGENT_TYPE = "OSR" Then
				PULSE_DEPARTMENT = "SPT"
			Else
				PULSE_DEPARTMENT = AGENT_DEPT
			End If

			For Each WAIVER in WAIVER_LIST
				WAIVER_DETAILS = Split(WAIVER,"_")
				MergeSQLstmt = MergeSQLstmt & "SELECT " & SCHEDULE_USR_ID & " RES_DLN_OPS_USR_ID, TO_DATE('" & WAIVER_DETAILS(0) & "','MM/DD/YYYY') RES_DLN_DATE, CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) RES_DLN_TIME, 'LWAV' RES_DLN_TYPE, '" & SCHEDULE_USR_ID & "' RES_DLN_TEXT, '" & PULSE_DEPARTMENT & "' RES_DLN_OPS_USR_TYPE, " & WAIVER_DETAILS(1) & " WAIVER_FLAG FROM DUAL UNION ALL "
			Next
			MergeSQLstmt = Left(MergeSQLstmt,Len(MergeSQLstmt)-11)

			SQLstmt = "MERGE INTO RES_DAILY_STATS_NOTES ORI " & _
			"USING " & _
			"( " & _
				MergeSQLstmt & _
			") CUR " & _
			"ON " & _
			"( " & _
				"ORI.RES_DLN_OPS_USR_ID = CUR.RES_DLN_OPS_USR_ID " & _
				"AND ORI.RES_DLN_DATE = CUR.RES_DLN_DATE " & _
				"AND ORI.RES_DLN_TYPE = CUR.RES_DLN_TYPE " & _
			") " & _
			"WHEN MATCHED THEN UPDATE " & _
				"SET ORI.RES_DLN_TIME = CUR.RES_DLN_TIME " & _
				"DELETE WHERE CUR.WAIVER_FLAG = 0 " & _
			"WHEN NOT MATCHED THEN INSERT " & _
				"(ORI.RES_DLN_OPS_USR_ID, ORI.RES_DLN_DATE, ORI.RES_DLN_TIME, ORI.RES_DLN_TYPE, ORI.RES_DLN_TEXT, ORI.RES_DLN_OPS_USR_TYPE) " & _
				"VALUES " & _
				"(CUR.RES_DLN_OPS_USR_ID, CUR.RES_DLN_DATE, CUR.RES_DLN_TIME, CUR.RES_DLN_TYPE, CUR.RES_DLN_TEXT, CUR.RES_DLN_OPS_USR_TYPE) " & _
				"WHERE CUR.WAIVER_FLAG = 1"
			cmd.CommandText = SQLstmt
			Set RSM = cmd.Execute
			Set RSM = Nothing
		End If
		Session("FILTER_AGENT") = SCHEDULE_USR_ID
		Session("FILTER_DATE") = SCHEDULE_START_DATE
		Response.Redirect(Request.ServerVariables("SCRIPT_NAME"))
		Response.End
	End If
%>
<%
	'Response.Write("Schedule center is currently unavailable. It will be available again shortly.")
	'Response.End
%>
<div id="main-copy">
	<h1 id="first-item">
		<a href="/communication/Schedule Center Guidelines.pdf" target="_blank" style="text-decoration:underline;font-weight:900;color:#FFF;">Schedule Center</a>
		<a href="/communication/Schedule Center Guidelines.pdf" target="_blank"><i class="fa fa-link" style="color:#FFF;font-size:16pt;" aria-hidden="true"></i></a>
	</h1>
	<br/>
	<div id="SCHEDSQUATCH_LEGEND">
		<div style="font-weight:900;font-size:14pt;margin-bottom:10px;">Schedule Center Legend</div>
		<div style="margin-bottom:10px;">
			<form id="FILTER_FORM" name="FILTER_FORM" method="post">
				<input type="hidden" name="USE_LAYOUT" value="<%=SCHEDSQUATCH_LAYOUT%>"/>
				<input type="hidden" name="STPLUS_ENABLED" value="<%=STPLUS_ENABLED%>"/>
				<% If SECURITY_LEVEL >= 2 and OPS_USR_JOB <> "LED" Then %>
					<span style="font-weight:900;">Agent:</span>
					<select style="margin:0px 15px;font-size:10pt;font-weight:900;font-family:Calibri;" id="FILTER_AGENT" name="FILTER_AGENT">
						<option value="-1">Not Selected</option>
						<%
							SQLstmt = "SELECT * FROM " & _
							"( " & _
								"SELECT OPS_USR_ID, OPS_USR_NAME " & _
								"FROM OPS_USER " & _
								"JOIN OPS_USER_DETAIL " & _
								"ON OPS_USR_ID = OPS_USD_OPS_USR_ID " & _
								"AND TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE " & _
								"WHERE OPS_USD_TYPE = 'RES' " & _
								"AND OPS_USD_JOB IN ('AGT','LED') " & _
								"AND OPS_USD_OPS_USR_ID NOT IN (2026,6068,9606) " & _
								"UNION " & _
								"SELECT OPS_USR_ID, OPS_USR_NAME " & _
								"FROM RES_SCHEDULE_MASTER " & _
								"JOIN OPS_USER_DETAIL " & _
								"ON RES_SCM_NAME = TO_CHAR(OPS_USD_OPS_USR_ID) " & _
								"AND TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE " & _
								"JOIN OPS_USER " & _
								"ON OPS_USD_OPS_USR_ID = OPS_USR_ID " & _
								"JOIN OPS_PARAMETER " & _
								"ON NULLIF(OPS_USD_TYPE,'RES') = OPS_PAR_PARENT_TYPE " & _
								"AND OPS_PAR_CODE = 'PICK_WINDOW' " & _
								"AND TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) BETWEEN OPS_PAR_EFF_DATE AND OPS_PAR_DIS_DATE " & _
								"LEFT JOIN SYS_CODE_DETAIL " & _
								"ON SYS_CDD_SYS_CDM_ID = 459 " & _
								"AND RES_SCM_NAME = SYS_CDD_NAME " & _
								"AND TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) BETWEEN TO_DATE(REGEXP_SUBSTR(SYS_CDD_VALUE,'[^_]+',1,1),'MM/DD/YYYY') AND TO_DATE(REGEXP_SUBSTR(SYS_CDD_VALUE,'[^_]+',1,2),'MM/DD/YYYY') " & _
								"WHERE RES_SCM_STATUS = 'ACT' " & _
								"AND TO_NUMBER(NVL(REGEXP_SUBSTR(SYS_CDD_VALUE,'[^_]+',1,4),TRIM(REPLACE(RES_SCM_TYPE,'NA',0)))) > 0 " & _
							") " & _
							"ORDER BY OPS_USR_NAME"
							Set RSAGT = Conn.Execute(SQLstmt)
						%>
						<% Do While Not RSAGT.EOF %>
							<option <% If SCHEDULE_USR_ID = CInt(RSAGT("OPS_USR_ID")) Then %>selected="selected" <% End If %>value="<%=RSAGT("OPS_USR_ID")%>"><%=RSAGT("OPS_USR_NAME")%></option>
							<% RSAGT.MoveNext %>
						<% Loop %>
						<% Set RSAGT = Nothing %>
					</select>
				<% End If %>
				<span style="font-weight:900;">Week Of:</span>
				<select style="margin:0px 15px;font-size:10pt;font-weight:900;font-family:Calibri;" id="FILTER_DATE" name="FILTER_DATE">
					<% If SECURITY_LEVEL >= 2 and OPS_USR_JOB <> "LED" Then %>
					<%
						SQLstmt = "SELECT START_DATE + 7*(ROWNUM - 1) WEEK_START " & _
						"FROM " & _
						"( " & _
							"SELECT TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)+1) - TO_CHAR(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)+1,'D') + 1 START_DATE, MAX(TO_DATE(OPS_SCI_START) - TO_CHAR(OPS_SCI_START,'D') + 1) END_DATE " & _
							"FROM OPS_SCHEDULE_INFO " & _
							"WHERE OPS_SCI_STATUS = 'APP' " & _
							"AND OPS_SCI_TYPE IN ('BASE','PICK') " & _
							"AND OPS_SCI_OPS_USR_TYPE = 'RES' " & _
							"AND TO_DATE(OPS_SCI_START) BETWEEN TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) AND TO_DATE('3/31/'||TO_CHAR(ADD_MONTHS(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE),12),'YYYY'),'MM/DD/YYYY') - TO_CHAR(TO_DATE('3/31/'||TO_CHAR(ADD_MONTHS(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE),12),'YYYY'),'MM/DD/YYYY'),'D') + 7 " & _
						") " & _
						"CONNECT BY ROWNUM < (END_DATE - START_DATE)/7 + 2"
						Set RSSD = Conn.Execute(SQLstmt)
					%>
					<% Else %>
					<%
						SQLstmt = "SELECT DISTINCT USE_DATE - TO_CHAR(USE_DATE,'D') + 1 WEEK_START " & _
						"FROM " & _
						"( " & _
							"SELECT TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) + ROWNUM - 1 USE_DATE, " & _
							"CASE WHEN TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) + ROWNUM - 1 BETWEEN TO_DATE(?,'MM/DD/YYYY') AND TO_DATE(?,'MM/DD/YYYY') THEN 'Y' ELSE 'N' END PICK_ENABLED " & _
							"FROM " & _
							"( " & _
								"SELECT " & _
								"MAX(TO_DATE(OPS_SCI_START)) MAX_DATE " & _
								"FROM OPS_SCHEDULE_INFO " & _
								"WHERE OPS_SCI_STATUS = 'APP' " & _
								"AND REGEXP_LIKE(OPS_SCI_TYPE,'BASE|PICK|HOLW|ADDT|EXTD|VACA|SLIP|RESH|RCHG|ROUT|MEET|PRES|PROJ|TRAN|FAMP|WFHU|MLTU|OTRG|NEWH|BRVT|JURY|HOLU|HOLR|PP$|PT$|[^FM]UN$') " & _
								"AND TO_DATE(OPS_SCI_START) BETWEEN TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) AND TO_DATE('3/31/'||TO_CHAR(ADD_MONTHS(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE),12),'YYYY'),'MM/DD/YYYY') - TO_CHAR(TO_DATE('3/31/'||TO_CHAR(ADD_MONTHS(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE),12),'YYYY'),'MM/DD/YYYY'),'D') + 7 " & _
								"AND OPS_SCI_OPS_USR_ID = ? " & _
							") " & _
							"CONNECT BY ROWNUM < MAX_DATE - TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) + 2 " & _
						") " & _
						"JOIN " & _
						"( " & _
							"SELECT " & _
							"MAX(DECODE(OPS_PAR_CODE,'DROP_WINDOW',CASE " & _
								"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),':') > 0 THEN TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24) - TO_CHAR(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24,'D') + 7*REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'[^:]+',1,1) + REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'[^:]+',1,2) " & _
								"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'/') > 0 THEN TO_DATE(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'MM/DD/YYYY') " & _
								"ELSE TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24) + REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1) " & _
							"END)) DROP_START_DATE, " & _
							"MAX(DECODE(OPS_PAR_CODE,'DROP_WINDOW',CASE " & _
								"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),':') > 0 THEN TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) - TO_CHAR(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE),'D') + 7*REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'[^:]+',1,1) + REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'[^:]+',1,2) " & _
								"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'/') > 0 THEN TO_DATE(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'MM/DD/YYYY') " & _
								"ELSE TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) + REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) " & _
							"END)) DROP_END_DATE, " & _
							"MAX(DECODE(OPS_PAR_CODE,'ADD_WINDOW',CASE " & _
								"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),':') > 0 THEN TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24) - TO_CHAR(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24,'D') + 7*REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'[^:]+',1,1) + REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'[^:]+',1,2) " & _
								"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'/') > 0 THEN TO_DATE(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'MM/DD/YYYY') " & _
								"ELSE TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24) + REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1) " & _
							"END)) ADD_START_DATE, " & _
							"MAX(DECODE(OPS_PAR_CODE,'ADD_WINDOW',CASE " & _
								"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),':') > 0 THEN TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) - TO_CHAR(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE),'D') + 7*REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'[^:]+',1,1) + REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'[^:]+',1,2) " & _
								"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'/') > 0 THEN TO_DATE(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'MM/DD/YYYY') " & _
								"ELSE TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) + REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) " & _
							"END)) ADD_END_DATE, " & _
							"MAX(DECODE(OPS_PAR_CODE,'SELFTRADE_WINDOW',CASE " & _
								"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),':') > 0 THEN TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24) - TO_CHAR(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24,'D') + 7*REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'[^:]+',1,1) + REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'[^:]+',1,2) " & _
								"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'/') > 0 THEN TO_DATE(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'MM/DD/YYYY') " & _
								"ELSE TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24) + REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1) " & _
							"END)) SELFTRADE_START_DATE, " & _
							"MAX(DECODE(OPS_PAR_CODE,'SELFTRADE_WINDOW',CASE " & _
								"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),':') > 0 THEN TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) - TO_CHAR(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE),'D') + 7*REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'[^:]+',1,1) + REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'[^:]+',1,2) " & _
								"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'/') > 0 THEN TO_DATE(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'MM/DD/YYYY') " & _
								"ELSE TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) + REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) " & _
							"END)) SELFTRADE_END_DATE, " & _
							"MAX(DECODE(OPS_PAR_CODE,'PICK_WINDOW',CASE " & _
								"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),':') > 0 THEN TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24) - TO_CHAR(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24,'D') + 7*REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'[^:]+',1,1) + REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'[^:]+',1,2) " & _
								"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'/') > 0 THEN TO_DATE(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'MM/DD/YYYY') " & _
								"ELSE TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24) + REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1) " & _
							"END)) PICK_START_DATE, " & _
							"MAX(DECODE(OPS_PAR_CODE,'PICK_WINDOW',CASE " & _
								"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),':') > 0 THEN TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) - TO_CHAR(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE),'D') + 7*REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'[^:]+',1,1) + REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'[^:]+',1,2) " & _
								"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'/') > 0 THEN TO_DATE(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'MM/DD/YYYY') " & _
								"ELSE TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) + REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) " & _
							"END)) PICK_END_DATE " & _
							"FROM OPS_PARAMETER PAR " & _
							"WHERE OPS_PAR_CODE IN ('DROP_WINDOW','ADD_WINDOW','SELFTRADE_WINDOW','PICK_WINDOW') " & _
							"AND TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) BETWEEN OPS_PAR_EFF_DATE AND OPS_PAR_DIS_DATE " & _
							"AND OPS_PAR_PARENT_TYPE = ? " & _
						") " & _
						"ON " & _
						"( " & _
							"USE_DATE BETWEEN DROP_START_DATE AND DROP_END_DATE " & _
							"OR USE_DATE BETWEEN ADD_START_DATE AND ADD_END_DATE " & _
							"OR USE_DATE BETWEEN SELFTRADE_START_DATE AND SELFTRADE_END_DATE " & _
							"OR " & _
							"( " & _
								"USE_DATE BETWEEN PICK_START_DATE AND PICK_END_DATE " & _
								"AND PICK_ENABLED = 'Y' " & _
							") " & _
						") " & _
						"ORDER BY WEEK_START"
						cmd.CommandText = SQLstmt
						cmd.Parameters(0).value = PICK_START_DATE
						cmd.Parameters(1).value = PICK_END_DATE
						cmd.Parameters(2).value = SCHEDULE_USR_ID
						cmd.Parameters(3).value = AGENT_TYPE
						Set RSSD = cmd.Execute
					%>
					<% End If %>

					<% Do While Not RSSD.EOF %>
						<option <% If SCHEDULE_START_DATE = CDate(RSSD("WEEK_START")) Then %>selected="selected" <% End If %>value="<%=RSSD("WEEK_START")%>"><%=RSSD("WEEK_START")%></option>
						<% RSSD.MoveNext %>
					<% Loop %>
					<% Set RSSD = Nothing %>
				</select>
				<input style="font-size:10pt;font-weight:900;font-family:Calibri;" type="button" id="FILTER_SUBMIT" name="FILTER_SUBMIT" value="Filter"/>
			</form>
		</div>
		<div id="SCHEDSQUATCH_TIMER" style="position:absolute;right:18px;top:9px;display:none;font-size:10pt;"></div>
		<div id="SCHEDSQUATCH_KEY_CONTAINER" style="margin-bottom:10px;display:none;">
			<i id="SCHEDSQUATCH_KEY_VIEW" class="fa fa-plus-square" style="cursor:pointer;" aria-hidden="true"></i><label for="SCHEDSQUATCH_KEY_VIEW" style="font-weight:900;margin-left:5px;cursor:pointer;">Schedule Key</label>
			<table id="SCHEDSQUATCH_KEY" style="width:100%;display:none;text-align:center;">
				<tr class="AD" style="display:none;">
					<td class="classicScheduleTd" style="background-color:#1b94d1;"></td>
					<td>
						Phone Time (Drop Time Not Available)
					</td>
					<td class="classicScheduleTd" style="background-color:#F7F7F7;text-align:center;"><input type="checkbox"/></td>
					<td>
						Lunch Time Available
					</td>
				</tr>
				<tr class="AD" style="display:none;">
					<td class="classicScheduleTd" style="background-color:#1b94d1;text-align:center;"><select class="classicInput" style="width:40px;background-color: #1b94d1;"><option></option></select></td>
					<td>
						Phone Time (Drop Time Available)
					</td>
					<td class="classicScheduleTd" style="background-color:#ff7f7f;text-align:center;"><input type="checkbox"/></td>
					<td>
						Additional/Lunch Time Available
					</td>
				</tr>
				<tr class="AD" style="display:none;">
					<td class="classicScheduleTd" style="background-color:#ab29b2;"></td>
					<td>
						Approved Time Off (Vacation, Appointments, etc.)
					</td>
					<td class="classicScheduleTd" style="background-color:#e5b219;"></td>
					<td>
						Dropped Time
					</td>
				</tr>
				<tr class="AD" style="display:none;">
					<td class="classicScheduleTd" style="background-color:#b2b2b2;"></td>
					<td>
						Training/Meeting/Presentation
					</td>
					<td></td>
					<td></td>
				</tr>
				<tr class="ST SP" style="display:none;">
					<td class="classicScheduleTd" style="background-color:#52a8d3;"></td>
					<td>
						Phone Time (Self-Trade Not Available)
					</td>
					<td class="classicScheduleTd" style="background-color:#F7F7F7;text-align:center;"><input type="checkbox"/></td>
					<td>
						Lunch Time Available
					</td>
				</tr>
				<tr class="ST" style="display:none;">
					<td class="classicScheduleTd" style="background-color:#52a8d3;text-align:center;"><input type="checkbox"/></td>
					<td>
						Phone Time (Self-Trade Available)
					</td>
					<td class="classicScheduleTd" style="background-color:#ff93a5;text-align:center;"><input type="checkbox"/></td>
					<td>
						Self-Trade/Lunch Time Available
					</td>
				</tr>
				<% If STPLUS_AVAILABLE = "Y" Then %>
					<tr class="SP" style="display:none;">
						<td class="classicScheduleTd" style="background:linear-gradient(to bottom right,#0074cc 50%,#52a8d3 50%);text-align:center;"><input type="checkbox"/></td>
						<td>
							Phone Time (Self-Trade <span style="color:#0074cc;font-weight:900;">Plus</span> Available)
						</td>
						<td class="classicScheduleTd" style="background:linear-gradient(to bottom right,#ff003b 50%,#ff93a5 50%);text-align:center;"><input type="checkbox"/></td>
						<td>
							Self-Trade <span style="color:#ff003b;font-weight:900;">Plus</span>/Lunch Time Available
						</td>
					</tr>
				<% End If %>
				<tr class="ST SP" style="display:none;">
					<td class="classicScheduleTd" style="background-color:#ab29b2;"></td>
					<td>
						Approved Time Off (Vacation, Appointments, etc.)
					</td>
					<td class="classicScheduleTd" style="background-color:#e5b219;"></td>
					<td>
						Dropped Time
					</td>
				</tr>
				<tr class="ST SP" style="display:none;">
					<td class="classicScheduleTd" style="background-color:#b2b2b2;"></td>
					<td>
						Training/Meeting/Presentation
					</td>
					<td></td>
					<td></td>
				</tr>
				<% If STPLUS_AVAILABLE = "Y" Then %>
					<tr class="SP" style="display:none;">
						<td style="text-align:center;" colspan="4">
							<div class="classicScheduleTd" style="display:inline-block;background-color:#0074cc;text-align:center;"><input type="checkbox" style="position:relative;top:15px;"/></div>
							<span style="position:relative;top:12px;font-weight:900;">Can be traded for</span>
							<div class="classicScheduleTd" style="display:inline-block;background:linear-gradient(to bottom right,#ff003b 50%,#ff93a5 50%);text-align:center;"><input type="checkbox" style="position:relative;top:15px;"/></div>
						</td>
					</tr>
					<tr class="SP" style="display:none;">
						<td style="text-align:center;" colspan="4">
							<div class="classicScheduleTd" style="display:inline-block;background-color:#52a8d3;text-align:center;"><input type="checkbox" style="position:relative;top:15px;"/></div>
							<span style="position:relative;top:12px;font-weight:900;">Can be traded for</span>
							<div class="classicScheduleTd" style="display:inline-block;background-color:#ff93a5;text-align:center;"><input type="checkbox" style="position:relative;top:15px;"/></div>
						</td>
					</tr>
				<% End If %>
				<tr class="PK" style="display:none;">
					<td class="classicScheduleTd" style="background-color:#1b94d1;"></td>
					<td>
						Phone Time (Base Hours)
					</td>
					<td class="classicScheduleTd" style="background-color:#aeedef;text-align:center;"><input type="checkbox" checked="checked"/></td>
					<td>
						Lunch Time
					</td>
				</tr>
				<tr class="PK" style="display:none;">
					<td class="classicScheduleTd" style="background-color:#1b94d1;text-align:center;"><input type="checkbox" checked="checked"/></td>
					<td>
						Phone Time (Pick Hours)
					</td>
					<td class="classicScheduleTd" style="background-color:#ffff46;text-align:center;"><input type="checkbox"/></td>
					<td>
						Available Time
					</td>
				</tr>
				<tr class="PK" style="display:none;">
					<td class="classicScheduleTd" style="background-color:#ab29b2;"></td>
					<td>
						Approved Time Off (Vacation, Appointments, etc.)
					</td>
					<td class="classicScheduleTd" style="background-color:#b2b2b2;"></td>
					<td>
						Training/Meeting/Presentation
					</td>
				</tr>
			</table>
		</div>
		<a id="LUNCH_FLEX_LINK" href="/general/lunchflex.asp" target="_blank" style="display:none;text-decoration:underline;font-weight:900;">Lunch Flex here!</a>
		<table id="SCHEDSQUATCH_STATS" style="width:80%;margin:10px auto;font-weight:900;">
			<tr class="AD" style="display:none;">
				<td id="UNP_REMAINING_DESCRIPTION" style="width:40%;text-align:center;">
					Unpaid Drop Time Remaining:
				</td>
				<td id="UNP_REMAINING_STAT" style="width:10%;"></td>
				<td id="ADD_REMAINING_DESCRIPTION" style="width:40%;text-align:center;">
					Add Time Remaining:
				</td>
				<td id="ADD_REMAINING_STAT" style="width:10%;"></td>
			</tr>
			<% If 1=1 Then %>
			<tr class="AD" style="display:none;margin-top:10px">
				<td id="SR_DISCLAIMER" colspan="4" style="text-align:center;padding-top:10px;font-size:6pt;">
					*Schedule reductions may be modified and/or canceled based on operational needs during times of high call volume or other unforeseen staffing events.
				</td>
			</tr>
			<% End If %>
			<tr class="ST SP" style="display:none;">
				<td id="SLT_DESCRIPTION" colspan="4" style="text-align:center;">
					Self Trade Alert:
					<span id="SLT_STAT" style="margin-left:10px;">None</span>
				</td>
			</tr>
			<% If STPLUS_AVAILABLE = "Y" Then %>
				<tr class="ST SP" style="display:none;">
					<td id="SLT_DESCRIPTION" colspan="4" style="text-align:center;">
						Self Trade Plus:
						<label class="switch" style="top:3px;left:2px;">
							<input id="STPLUS_SWITCH" type="checkbox" value="SP"<% If STPLUS_ENABLED = "Y" Then %> checked="checked"<% End If %>>
							<span class="slider round"></span>
						</label>
					</td>
				</tr>
			<% End If %>
			<tr class="PK" style="display:none;">
				<td style="width:40%;text-align:center;">
					Base Hours:
				</td>
				<td id="BASE_HOURS_STAT" style="width:10%;"></td>
				<td style="width:40%;text-align:center;">
					Pick Hours:
				</td>
				<td id="PICK_HOURS_STAT" style="width:10%;"></td>
			</tr>
			<tr class="PK" style="display:none;">
				<td id="TOTAL_SCHEDULED_DESCRIPTION" style="text-align:center;">
					Total Scheduled Hours:
				</td>
				<td id="TOTAL_SCHEDULED_STAT"></td>
				<td id="REMAINING_NEED_DESCRIPTION" style="text-align:center;">
					Remaining Need:
				</td>
				<td id="REMAINING_NEED_STAT"></td>
			</tr>
			<tr id="PICK_HOLIDAY_ROW" class="PK" style="display:none;">
				<td style="text-align:center;">
					Holiday Hours:
				</td>
				<td id="HOLIDAY_STAT"></td>
				<td id="REMAINING_POSSIBLE_DESCRIPTION" style="text-align:center;">
					Remaining Possible:
				</td>
				<td id="REMAINING_POSSIBLE_STAT"></td>
			</tr>
		</table>
	</div>
	<%
		Dim DROP_DATE(6)
		Dim ADD_DATE(6)
		Dim SELFTRADE_DATE(6)
		Dim SELFTRADEPLUS_DATE(6)
		Dim PICK_DATE(6)
		Dim DAY_CHECK(6)

		SQLstmt = "SELECT " & _
		"CASE WHEN NVL(TO_NUMBER(DT.OPS_PAR_VALUE),12) < 12 THEN NVL2(LOAFLAG.SYS_CDD_ID,10,TO_NUMBER(DT.OPS_PAR_VALUE)) ELSE NVL(TO_NUMBER(DT.OPS_PAR_VALUE),12) END THRESHOLD_HOURS, " & _
		"GREATEST(CASE WHEN NVL(TO_NUMBER(DT.OPS_PAR_VALUE),12) < 12 THEN NVL2(LOAFLAG.SYS_CDD_ID,10,TO_NUMBER(DT.OPS_PAR_VALUE)) ELSE NVL(TO_NUMBER(DT.OPS_PAR_VALUE),12) END - NVL(DAILY_THRESHOLD,0),0) DAILY_THRESHOLD, " & _
		"NVL(LUNCH_WORKED,0) LUNCH_WORKED, " & _
		"NVL(LUNCH_HOURS,0) LUNCH_HOURS, " & _
		"NVL(DROP_END_DATE,TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE))) DROP_END_DATE, " & _
		"CASE WHEN SCHEDULE_DATE BETWEEN DROP_START_DATE AND DROP_END_DATE THEN 1 ELSE 0 END DROP_DATE, " & _
		"CASE WHEN SCHEDULE_DATE BETWEEN ADD_START_DATE AND ADD_END_DATE THEN 1 ELSE 0 END ADD_DATE, " & _
		"CASE WHEN SCHEDULE_DATE BETWEEN SELFTRADE_START_DATE AND SELFTRADE_END_DATE THEN 1 ELSE 0 END SELFTRADE_DATE, " & _
		"CASE WHEN SCHEDULE_DATE BETWEEN SELFTRADEPLUS_START_DATE AND SELFTRADEPLUS_END_DATE THEN 1 ELSE 0 END SELFTRADEPLUS_DATE, " & _
		"CASE WHEN SCHEDULE_DATE BETWEEN PICK_START_DATE AND PICK_END_DATE AND PICK_HOURS > 0 THEN 1 ELSE 0 END PICK_DATE, " & _
		"NVL(MIN(CASE " & _
			"WHEN SCHEDULE_DATE BETWEEN DROP_START_DATE AND DROP_END_DATE OR SCHEDULE_DATE BETWEEN ADD_START_DATE AND ADD_END_DATE OR SCHEDULE_DATE BETWEEN SELFTRADE_START_DATE AND SELFTRADE_END_DATE THEN 300 " & _
			"WHEN SCHEDULE_DATE BETWEEN PICK_START_DATE AND PICK_END_DATE AND PICK_HOURS > 0 THEN 900 " & _
		"END) OVER (),30) USE_TIMER, " & _
        "MAX(CASE " & _
            "WHEN SCHEDULE_DATE BETWEEN DROP_START_DATE AND DROP_END_DATE OR SCHEDULE_DATE BETWEEN ADD_START_DATE AND ADD_END_DATE THEN 'M' " & _
        "END) OVER () || " & _
        "MAX(CASE " & _
            "WHEN SCHEDULE_DATE BETWEEN SELFTRADE_START_DATE AND SELFTRADE_END_DATE THEN 'S' " & _
        "END) OVER () || " & _
        "MAX(CASE " & _
            "WHEN SCHEDULE_DATE BETWEEN PICK_START_DATE AND PICK_END_DATE AND PICK_HOURS > 0 THEN 'P' " & _
        "END) OVER () OPEN_DATE, " & _
		"BASE_HOURS, " & _
		"PICK_HOURS, " & _
		"HOLIDAY_HOURS, " & _
		"HOLIDAY_DEDUCTION, " & _
		"DECODE(?,'LED',0,FLOOR(2*GREATEST(NVL(REDUCEL.REDUCE_HOURS,0) - UNPAID_HOURS,0))/2) UNP_REMAINING, " & _
		"FLOOR(2*GREATEST(40 + NVL(ADDL.OPS_PAR_VALUE,0) - PAID_HOURS - UNPAID_HOURS,0))/2 ADD_REMAINING, " & _
		"PAID_HOURS + UNPAID_HOURS TOTAL_SCHEDULED, " & _
		"GREATEST(BASE_HOURS + PICK_HOURS, PAID_HOURS + UNPAID_HOURS) TOTAL_EXPECTED, " & _
		"GREATEST(BASE_HOURS + PICK_HOURS - HOLIDAY_DEDUCTION - PAID_HOURS - UNPAID_HOURS,0) REMAINING_NEED, " & _
		"GREATEST(GREATEST(BASE_HOURS + PICK_HOURS, PAID_HOURS + UNPAID_HOURS) - PAID_HOURS - UNPAID_HOURS,0) REMAINING_POSSIBLE, " & _
		"GREATEST(PTO_BALANCE,0) PTO_BALANCE, " & _
		"CASE " & _
			"WHEN SCHEDULE_DATE < NEWH_DATE " & _
			"OR RESTRICTDROP.SYS_CDD_ID IS NOT NULL " & _
			"THEN -1 " & _
			"ELSE 1 " & _
		"END DAY_CHECK " & _
		"FROM " & _
		"( " & _
			"SELECT TO_DATE(?,'MM/DD/YYYY') + (ROWNUM - 1) SCHEDULE_DATE " & _
			"FROM DUAL " & _
			"CONNECT BY ROWNUM <= 7 " & _
		") " & _
		"CROSS JOIN " & _
		"( " & _
			"SELECT " & _
			"ROUND(SUM(DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'SLIP|HOLU|[^FM]UN$'),0,0,24*(OPS_SCI_END-OPS_SCI_START))),1) UNPAID_HOURS, " & _
			"ROUND(SUM(DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'BASE|PICK|HOLW|ADDT|EXTD|VACA|RESH|RCHG|ROUT|MEET|PRES|PROJ|TRAN|FAMP|WFHU|MLTU|OTRG|NEWH|BRVT|JURY|HOLR|PP$|PT$'),0,0,CASE WHEN OPS_SCI_TYPE = 'HOLR' AND RES_BUE_ID IS NOT NULL THEN 0 ELSE 24*(OPS_SCI_END-OPS_SCI_START) END)),1) PAID_HOURS, " & _
			"ROUND(SUM(DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'HOLU|HOLR'),0,0,24*(OPS_SCI_END-OPS_SCI_START))),1) HOLIDAY_HOURS, " & _
			"ROUND(SUM(CASE WHEN (TO_CHAR(OPS_SCI_START,'MM') = 1 AND TO_CHAR(OPS_SCI_START,'DD') > 3) OR TO_CHAR(OPS_SCI_START,'MM') IN (2,3,4) THEN 0 ELSE DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'HOLU|HOLR'),0,0,NVL2(RES_BUE_ID,24*(OPS_SCI_END-OPS_SCI_START),0)) END),1) HOLIDAY_DEDUCTION " & _
			"FROM OPS_SCHEDULE_INFO " & _
			"LEFT JOIN RES_BUDGET_EXCEPTION " & _
			"ON TO_DATE(OPS_SCI_START) = RES_BUE_DATE " & _
			"AND RES_BUE_TYPE = 'NOR' " & _
			"WHERE OPS_SCI_STATUS = 'APP' " & _
			"AND REGEXP_LIKE(OPS_SCI_TYPE,'BASE|PICK|HOLW|ADDT|EXTD|VACA|SLIP|RESH|RCHG|ROUT|MEET|PRES|PROJ|TRAN|FAMP|WFHU|MLTU|OTRG|NEWH|BRVT|JURY|HOLU|HOLR|PP$|PT$|[^FM]UN$') " & _
			"AND TO_DATE(OPS_SCI_START) BETWEEN TO_DATE(?,'MM/DD/YYYY') AND TO_DATE(?,'MM/DD/YYYY') " & _
			"AND OPS_SCI_OPS_USR_ID = ? " & _
		") " & _
		"CROSS JOIN " & _
		"( " & _
			"SELECT NVL(ROUND(AVG(OPS_QUM_SCORE),3),3) EVAL_SCORE " & _
			"FROM " & _
			"( " & _
				"SELECT OPS_QUM_SCORE, " & _
				"ROW_NUMBER() OVER (ORDER BY OPS_QUM_YEAR DESC, OPS_QUM_MONTH DESC) EVAL_NUM " & _
				"FROM OPS_QUALITY_MASTER " & _
				"JOIN OPS_FORM_MASTER " & _
				"ON OPS_QFM_ID = OPS_QUM_OPS_QFM_ID " & _
				"AND UPPER(OPS_QFM_NAME) LIKE '%MONTHLY EVALUATION' " & _
				"WHERE OPS_QUM_STATUS = 'COM' " & _
				"AND OPS_QUM_AGT_OPS_USR_ID = ? " & _
			") " & _
			"WHERE EVAL_NUM <= 3 " & _
		") " & _
		"CROSS JOIN " & _
		"( " & _
			"SELECT USE_DATE + NEWH_INCREMENT NEWH_DATE " & _
			"FROM " & _
			"( " & _
				"SELECT NVL(MAX(TO_DATE(OPS_SCI_START)),TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)-365)) USE_DATE " & _
				"FROM OPS_SCHEDULE_INFO " & _
				"WHERE OPS_SCI_OPS_USR_ID = ? " & _
				"AND OPS_SCI_STATUS = 'APP' " & _
				"AND OPS_SCI_TYPE IN ('SLIPP','NEWH') " & _
			") " & _
			"CROSS JOIN " & _
			"( " & _
				"SELECT NVL(MAX(OPS_PAR_VALUE),0) NEWH_INCREMENT " & _
				"FROM OPS_PARAMETER " & _
				"WHERE OPS_PAR_CODE = 'NEW_HIRE' " & _
				"AND OPS_PAR_PARENT_TYPE = ? " & _
				"AND TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) BETWEEN OPS_PAR_EFF_DATE AND OPS_PAR_DIS_DATE " & _
			") " & _
		") " & _
		"CROSS JOIN " & _
		"( " & _
			"SELECT TO_NUMBER(NVL(REGEXP_SUBSTR(SYS_CDD_VALUE,'[^_]+',1,3),TRIM(RES_SCM_SHIFT_LENGTH))) BASE_HOURS, TO_NUMBER(NVL(REGEXP_SUBSTR(SYS_CDD_VALUE,'[^_]+',1,4),TRIM(REPLACE(RES_SCM_TYPE,'NA',0)))) PICK_HOURS " & _
			"FROM RES_SCHEDULE_MASTER " & _
			"LEFT JOIN SYS_CODE_DETAIL " & _
			"ON SYS_CDD_SYS_CDM_ID = 459 " & _
			"AND RES_SCM_NAME = SYS_CDD_NAME " & _
			"AND TO_DATE(?,'MM/DD/YYYY') BETWEEN TO_DATE(REGEXP_SUBSTR(SYS_CDD_VALUE,'[^_]+',1,1),'MM/DD/YYYY') AND TO_DATE(REGEXP_SUBSTR(SYS_CDD_VALUE,'[^_]+',1,2),'MM/DD/YYYY') " & _
			"WHERE RES_SCM_NAME = ? " & _
			"AND RES_SCM_STATUS = 'ACT' " & _
			"AND 'SUP' <> ? " & _
		") " & _
		"CROSS JOIN " & _
		"( " & _
			"SELECT " & _
			"MAX(DECODE(OPS_PAR_CODE,'DROP_WINDOW',CASE " & _
				"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),':') > 0 THEN TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24) - TO_CHAR(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24,'D') + 7*REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'[^:]+',1,1) + REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'[^:]+',1,2) " & _
				"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'/') > 0 THEN TO_DATE(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'MM/DD/YYYY') " & _
				"ELSE TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24) + REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1) " & _
			"END)) DROP_START_DATE, " & _
			"MAX(DECODE(OPS_PAR_CODE,'DROP_WINDOW',CASE " & _
				"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),':') > 0 THEN TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) - TO_CHAR(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE),'D') + 7*REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'[^:]+',1,1) + REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'[^:]+',1,2) " & _
				"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'/') > 0 THEN TO_DATE(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'MM/DD/YYYY') " & _
				"ELSE TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) + REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) " & _
			"END)) DROP_END_DATE, " & _
			"MAX(DECODE(OPS_PAR_CODE,'ADD_WINDOW',CASE " & _
				"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),':') > 0 THEN TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24) - TO_CHAR(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24,'D') + 7*REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'[^:]+',1,1) + REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'[^:]+',1,2) " & _
				"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'/') > 0 THEN TO_DATE(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'MM/DD/YYYY') " & _
				"ELSE TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24) + REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1) " & _
			"END)) ADD_START_DATE, " & _
			"MAX(DECODE(OPS_PAR_CODE,'ADD_WINDOW',CASE " & _
				"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),':') > 0 THEN TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) - TO_CHAR(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE),'D') + 7*REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'[^:]+',1,1) + REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'[^:]+',1,2) " & _
				"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'/') > 0 THEN TO_DATE(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'MM/DD/YYYY') " & _
				"ELSE TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) + REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) " & _
			"END)) ADD_END_DATE, " & _
			"MAX(DECODE(OPS_PAR_CODE,'SELFTRADE_WINDOW',CASE " & _
				"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),':') > 0 THEN TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24) - TO_CHAR(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24,'D') + 7*REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'[^:]+',1,1) + REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'[^:]+',1,2) " & _
				"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'/') > 0 THEN TO_DATE(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'MM/DD/YYYY') " & _
				"ELSE TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24) + REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1) " & _
			"END)) SELFTRADE_START_DATE, " & _
			"MAX(DECODE(OPS_PAR_CODE,'SELFTRADE_WINDOW',CASE " & _
				"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),':') > 0 THEN TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) - TO_CHAR(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE),'D') + 7*REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'[^:]+',1,1) + REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'[^:]+',1,2) " & _
				"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'/') > 0 THEN TO_DATE(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'MM/DD/YYYY') " & _
				"ELSE TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) + REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) " & _
			"END)) SELFTRADE_END_DATE, " & _
			"MAX(DECODE(OPS_PAR_CODE,'SELFTRADEPLUS_WINDOW',CASE " & _
				"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),':') > 0 THEN TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24) - TO_CHAR(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24,'D') + 7*REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'[^:]+',1,1) + REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'[^:]+',1,2) " & _
				"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'/') > 0 THEN TO_DATE(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'MM/DD/YYYY') " & _
				"ELSE TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24) + REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1) " & _
			"END)) SELFTRADEPLUS_START_DATE, " & _
			"MAX(DECODE(OPS_PAR_CODE,'SELFTRADEPLUS_WINDOW',CASE " & _
				"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),':') > 0 THEN TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) - TO_CHAR(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE),'D') + 7*REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'[^:]+',1,1) + REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'[^:]+',1,2) " & _
				"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'/') > 0 THEN TO_DATE(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'MM/DD/YYYY') " & _
				"ELSE TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) + REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) " & _
			"END)) SELFTRADEPLUS_END_DATE, " & _
			"MAX(DECODE(OPS_PAR_CODE,'PICK_WINDOW',CASE " & _
				"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),':') > 0 THEN TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24) - TO_CHAR(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24,'D') + 7*REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'[^:]+',1,1) + REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'[^:]+',1,2) " & _
				"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'/') > 0 THEN TO_DATE(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),'MM/DD/YYYY') " & _
				"ELSE TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) - NVL(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),0)/24) + REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1) " & _
			"END)) PICK_START_DATE, " & _
			"MAX(DECODE(OPS_PAR_CODE,'PICK_WINDOW',CASE " & _
				"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),':') > 0 THEN TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) - TO_CHAR(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE),'D') + 7*REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'[^:]+',1,1) + REGEXP_SUBSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'[^:]+',1,2) " & _
				"WHEN INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'/') > 0 THEN TO_DATE(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2),'MM/DD/YYYY') " & _
				"ELSE TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) + REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) " & _
			"END)) PICK_END_DATE " & _
			"FROM OPS_PARAMETER PAR " & _
			"WHERE OPS_PAR_CODE IN ('DROP_WINDOW','ADD_WINDOW','SELFTRADE_WINDOW','SELFTRADEPLUS_WINDOW','PICK_WINDOW') " & _
			"AND TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) BETWEEN OPS_PAR_EFF_DATE AND OPS_PAR_DIS_DATE " & _
			"AND OPS_PAR_PARENT_TYPE = ? " & _
		")WINDOW " & _
		"LEFT JOIN OPS_PARAMETER DT " & _
		"ON DT.OPS_PAR_PARENT_TYPE = 'STF' " & _
		"AND DT.OPS_PAR_CODE = 'DAILY_THRESHOLD' " & _
		"AND SCHEDULE_DATE BETWEEN DT.OPS_PAR_EFF_DATE AND DT.OPS_PAR_DIS_DATE " & _
		"LEFT JOIN SYS_CODE_DETAIL LOAFLAG " & _
		"ON LOAFLAG.SYS_CDD_SYS_CDM_ID = 583 " & _
		"AND LOAFLAG.SYS_CDD_NAME = ? " & _
		"AND SCHEDULE_DATE > TO_DATE(REGEXP_SUBSTR(LOAFLAG.SYS_CDD_VALUE,'[^_]+',1,2),'MM/DD/YYYY') " & _
		"LEFT JOIN SYS_CODE_DETAIL RESTRICTDROP " & _
		"ON RESTRICTDROP.SYS_CDD_SYS_CDM_ID = 561 " & _
		"AND RESTRICTDROP.SYS_CDD_NAME = ? " & _
		"AND SCHEDULE_DATE BETWEEN TO_DATE(REGEXP_SUBSTR(RESTRICTDROP.SYS_CDD_VALUE,'[^;]+',1,1),'MM/DD/YYYY') AND TO_DATE(REGEXP_SUBSTR(RESTRICTDROP.SYS_CDD_VALUE,'[^;]+',1,2),'MM/DD/YYYY') " & _
		"LEFT JOIN " & _
		"( " & _
			"SELECT TO_DATE(OPS_SCI_START) SCI_DATE, " & _
			"ROUND(SUM(DECODE(OPS_SCI_TYPE,'LNCH',0,24*(OPS_SCI_END-OPS_SCI_START))),1) DAILY_THRESHOLD, " & _
			"ROUND(SUM(DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'BASE|PICK|ADDT|EXTD|HOLW|MEET|TRAN|PROJ|NEWH|WFHU'),0,0,24*(OPS_SCI_END-OPS_SCI_START))),1) LUNCH_WORKED, " & _
			"ROUND(SUM(DECODE(OPS_SCI_TYPE,'LNCH',24*(OPS_SCI_END-OPS_SCI_START),0)),1) LUNCH_HOURS " & _
			"FROM OPS_SCHEDULE_INFO " & _
			"WHERE OPS_SCI_STATUS = 'APP' " & _
			"AND OPS_SCI_TYPE NOT IN ('LNFL','HOLR','HOLU') " & _
			"AND TO_DATE(OPS_SCI_START) BETWEEN TO_DATE(?,'MM/DD/YYYY') AND TO_DATE(?,'MM/DD/YYYY') " & _
			"AND OPS_SCI_OPS_USR_ID = ? " & _
			"GROUP BY TO_DATE(OPS_SCI_START) " & _
		") " & _
		"ON SCHEDULE_DATE = SCI_DATE " & _
		"LEFT JOIN OPS_PARAMETER ADDL " & _
		"ON ADDL.OPS_PAR_PARENT_TYPE = ? " & _
		"AND ADDL.OPS_PAR_CODE = 'ADD_LIMIT' " & _
		"AND SCHEDULE_DATE BETWEEN ADDL.OPS_PAR_EFF_DATE AND ADDL.OPS_PAR_DIS_DATE " & _
		"LEFT JOIN " & _
		"( " & _
			"SELECT " & _
			"REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1) SCHEDULED_HOURS_MIN, " & _
			"REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) SCHEDULED_HOURS_MAX, " & _
			"REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3) EVAL_MIN, " & _
			"REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,4) EVAL_MAX, " & _
			"REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,5) REDUCE_HOURS, " & _
			"OPS_PAR_EFF_DATE, " & _
			"OPS_PAR_DIS_DATE " & _
			"FROM OPS_PARAMETER " & _
			"WHERE OPS_PAR_CODE = 'REDUCE_LIMIT' " & _
			"AND OPS_PAR_PARENT_TYPE = ? " & _
		")REDUCEL " & _
		"ON GREATEST(?,0.1) > REDUCEL.SCHEDULED_HOURS_MIN " & _
		"AND GREATEST(?,0.1) <= REDUCEL.SCHEDULED_HOURS_MAX " & _
		"AND LEAST(EVAL_SCORE,4.999) >= REDUCEL.EVAL_MIN " & _
		"AND LEAST(EVAL_SCORE,4.999) < REDUCEL.EVAL_MAX " & _
		"AND SCHEDULE_DATE BETWEEN REDUCEL.OPS_PAR_EFF_DATE AND REDUCEL.OPS_PAR_DIS_DATE " & _
		"LEFT JOIN " & _
		"( " & _
			"SELECT " & _
			"BALANCE_START_DATE, " & _
			"BALANCE_END_DATE, " & _
			"MIN(CASE WHEN BALANCE_END_DATE >= TO_DATE(?,'MM/DD/YYYY') THEN PTO_BALANCE END) OVER (ORDER BY BALANCE_END_DATE RANGE BETWEEN CURRENT ROW AND UNBOUNDED FOLLOWING) PTO_BALANCE " & _
			"FROM " & _
			"( " & _
				"SELECT " & _
				"BALANCE_DATE BALANCE_START_DATE, " & _
				"NVL(LEAD(BALANCE_DATE) OVER (ORDER BY BALANCE_DATE) - 1,TO_DATE('12/31/2040','MM/DD/YYYY')) BALANCE_END_DATE, " & _
				"PTO_BALANCE " & _
				"FROM " & _
				"( " & _
					"SELECT " & _
					"DISTINCT " & _
					"BALANCE_DATE, " & _
					"TRUNC(SUM(PTO_BALANCE) OVER (ORDER BY BALANCE_DATE, USE_ORDER RANGE UNBOUNDED PRECEDING),1) PTO_BALANCE " & _
					"FROM " & _
					"( " & _
						"SELECT " & _
						"1 USE_ORDER, " & _
						"OPS_ACC_DATE BALANCE_DATE, " & _
						"SUM(TRUNC(OPS_ACC_BALANCE,1)) PTO_BALANCE " & _
						"FROM OPS_ACCRUAL " & _
						"WHERE OPS_ACC_OPS_USR_ID = ? " & _
						"AND OPS_ACC_CODE IN ('VACA','PPTV') " & _
						"GROUP BY OPS_ACC_DATE " & _
						"UNION ALL " & _
						"SELECT " & _
						"2, " & _
						"TO_DATE('4/1/'||TO_CHAR(ADD_MONTHS(OPS_ACC_DATE,9),'YYYY'),'MM/DD/YYYY'), " & _
						"TRUNC(OPS_ACC_BALANCE,1) " & _
						"FROM OPS_ACCRUAL " & _
						"WHERE OPS_ACC_OPS_USR_ID = ? " & _
						"AND OPS_ACC_CODE = 'ACVC' " & _
						"UNION ALL " & _
						"SELECT " & _
						"2, " & _
						"TO_DATE(OPS_SCI_START), " & _
						"-1*SUM(ROUND(24*(OPS_SCI_END-OPS_SCI_START),1)) " & _
						"FROM OPS_SCHEDULE_INFO " & _
						"WHERE OPS_SCI_OPS_USR_ID = ? " & _
						"AND REGEXP_LIKE(OPS_SCI_TYPE,'^VA|PT$|PP$') " & _
						"AND OPS_SCI_STATUS = 'APP' " & _
						"AND TO_DATE(OPS_SCI_START) >= (SELECT MAX(OPS_ACC_DATE) FROM OPS_ACCRUAL) " & _
						"GROUP BY TO_DATE(OPS_SCI_START) " & _
						"UNION ALL " & _
						"SELECT " & _
						"2, " & _
						"USE_DATE, " & _
						"CASE " & _
							"WHEN " & _
								"(SELECT MAX(OPS_ACC_DATE) FROM OPS_ACCRUAL) BETWEEN TO_DATE('3/18/'||TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'YYYY'),'MM/DD/YYYY') AND TO_DATE('3/31/'||TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'YYYY'),'MM/DD/YYYY') " & _
								"AND USE_DATE BETWEEN TO_DATE('3/18/'||TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'YYYY'),'MM/DD/YYYY') AND TO_DATE('3/31/'||TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'YYYY'),'MM/DD/YYYY') " & _
							"THEN 0 " & _
							"WHEN OPS_USD_SCH_HOURS >= 35.1 THEN 1.08 " & _
							"WHEN OPS_USD_SCH_HOURS >= 30.1 THEN .94 " & _
							"WHEN OPS_USD_SCH_HOURS >= 25.1 THEN .81 " & _
							"WHEN OPS_USD_SCH_HOURS >= 20.1 THEN .67 " & _
							"WHEN OPS_USD_SCH_HOURS > 0 THEN .54 " & _
							"ELSE 0 " & _
						"END " & _
						"FROM OPS_USER_DETAIL " & _
						"JOIN " & _
						"( " & _
							"SELECT MAX_DATE + (7*(ROWNUM-1)) USE_DATE " & _
							"FROM " & _
							"( " & _
								"SELECT MAX(OPS_ACC_DATE) + 12 MAX_DATE " & _
								"FROM OPS_ACCRUAL " & _
							") " & _
							"CONNECT BY ROWNUM < FLOOR((TO_DATE('3/31/'||TO_CHAR(ADD_MONTHS(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE),12),'YYYY'),'MM/DD/YYYY') - MAX_DATE)/7) + 2 " & _
						") " & _
						"ON USE_DATE BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE " & _
						"AND OPS_USD_OPS_USR_ID = ? " & _
					") " & _
				") " & _
			") " & _
		") " & _
		"ON SCHEDULE_DATE BETWEEN BALANCE_START_DATE AND BALANCE_END_DATE " & _
		"ORDER BY SCHEDULE_DATE"
		cmd.CommandText = SQLstmt
		cmd.Parameters(0).value = AGENT_JOB
		cmd.Parameters(1).value = SCHEDULE_START_DATE
		cmd.Parameters(2).value = SCHEDULE_START_DATE
		cmd.Parameters(3).value = SCHEDULE_END_DATE'
		cmd.Parameters(4).value = SCHEDULE_USR_ID
		cmd.Parameters(5).value = SCHEDULE_USR_ID
		cmd.Parameters(6).value = SCHEDULE_USR_ID
		cmd.Parameters(7).value = AGENT_TYPE
		cmd.Parameters(8).value = SCHEDULE_START_DATE
		cmd.Parameters(9).value = SCHEDULE_USR_ID
		cmd.Parameters(10).value = AGENT_JOB
		cmd.Parameters(11).value = AGENT_TYPE
		cmd.Parameters(12).value = SCHEDULE_USR_ID
		cmd.Parameters(13).value = SCHEDULE_USR_ID
		cmd.Parameters(14).value = SCHEDULE_START_DATE
		cmd.Parameters(15).value = SCHEDULE_END_DATE
		cmd.Parameters(16).value = SCHEDULE_USR_ID
		cmd.Parameters(17).value = AGENT_TYPE
		cmd.Parameters(18).value = AGENT_TYPE
		cmd.Parameters(19).value = AGENT_HOURS
		cmd.Parameters(20).value = AGENT_HOURS
		cmd.Parameters(21).value = SCHEDULE_START_DATE
		cmd.Parameters(22).value = SCHEDULE_USR_ID
		cmd.Parameters(23).value = SCHEDULE_USR_ID
		cmd.Parameters(24).value = SCHEDULE_USR_ID
		cmd.Parameters(25).value = SCHEDULE_USR_ID
		Set RSDAY = cmd.Execute
	%>
	<% If Not RSDAY.EOF Then %>
		<% n = 0 %>
		<% OPEN_DATE = RSDAY("OPEN_DATE") %>
		<% DROP_MAX_INTERVAL = CDate(RSDAY("DROP_END_DATE") & " " & Time) %>
		<input type="hidden" id="UNP_REMAINING" value="<%=RSDAY("UNP_REMAINING")%>"/>
		<input type="hidden" id="ADD_REMAINING" value="<%=RSDAY("ADD_REMAINING")%>"/>
		<input type="hidden" id="BASE_HOURS" value="<%=RSDAY("BASE_HOURS")%>"/>
		<input type="hidden" id="PICK_HOURS" value="<%=RSDAY("PICK_HOURS")%>"/>
		<input type="hidden" id="HOLIDAY_HOURS" value="<%=RSDAY("HOLIDAY_HOURS")%>"/>
		<input type="hidden" id="HOLIDAY_DEDUCTION" value="<%=RSDAY("HOLIDAY_DEDUCTION")%>"/>
		<input type="hidden" id="TOTAL_SCHEDULED" value="<%=RSDAY("TOTAL_SCHEDULED")%>"/>
		<input type="hidden" id="TOTAL_EXPECTED" value="<%=RSDAY("TOTAL_EXPECTED")%>"/>
		<input type="hidden" id="REMAINING_NEED" value="<%=RSDAY("REMAINING_NEED")%>"/>
		<input type="hidden" id="REMAINING_POSSIBLE" value="<%=RSDAY("REMAINING_POSSIBLE")%>"/>
		<input type="hidden" id="SELFTRADE_COUNTER" value="0"/>
		<input type="hidden" id="PLUS_COUNTER" value="0"/>
		<input type="hidden" id="USE_TIMER" value="<%=RSDAY("USE_TIMER")%>"/>
		<% Do While Not RSDAY.EOF %>
			<input type="hidden" id="THRESHOLD_HOURS_<%=n%>" value="<%=RSDAY("THRESHOLD_HOURS")%>"/>
			<input type="hidden" id="DAILY_THRESHOLD_<%=n%>" value="<%=RSDAY("DAILY_THRESHOLD")%>"/>
			<input type="hidden" id="PTO_BALANCE_<%=n%>" value="<%=RSDAY("PTO_BALANCE")%>"/>
			<input type="hidden" id="LUNCH_WORKED_<%=n%>" value="<%=RSDAY("LUNCH_WORKED")%>"/>
			<input type="hidden" id="LUNCH_HOURS_<%=n%>" value="<%=RSDAY("LUNCH_HOURS")%>"/>
			<% DROP_DATE(n) = RSDAY("DROP_DATE") %>
			<% ADD_DATE(n) = RSDAY("ADD_DATE") %>
			<% SELFTRADE_DATE(n) = RSDAY("SELFTRADE_DATE") %>
			<% SELFTRADEPLUS_DATE(n) = RSDAY("SELFTRADEPLUS_DATE") %>
			<% PICK_DATE(n) = RSDAY("PICK_DATE") %>
			<% DAY_CHECK(n) = RSDAY("DAY_CHECK") %>
			<% RSDAY.MoveNext %>
			<% n = n + 1 %>
		<% Loop %>
		<%
			SQLstmt = "SELECT " & _
			"OPS_SCN_INTERVAL, " & _
			"TO_CHAR(TO_DATE(OPS_SCN_INTERVAL,'HH24:MI'),'HH:MI AM') CUR_DISP_INTERVAL, " & _
			"TO_CHAR(TO_DATE(OPS_SCN_INTERVAL,'HH24:MI')+1/48,'HH:MI AM') NEXT_DISP_INTERVAL, " & _
			"REGEXP_SUBSTR(DAY_0,'[^;]+',1,1) TYPE_0, " & _
			"REGEXP_SUBSTR(DAY_0,'[^;]+',1,2) NOTES_0, " & _
			"REGEXP_SUBSTR(DAY_0,'[^;]+',1,3) PLUS_MINUS_0, " & _
			"REGEXP_SUBSTR(DAY_0,'[^;]+',1,4) PLUS_0, " & _
			"REGEXP_SUBSTR(DAY_0,'[^;]+',1,5) LUNCH_0, " & _
			"REGEXP_SUBSTR(DAY_0,'[^;]+',1,6) INT_DROP_0, " & _
			"REGEXP_SUBSTR(DAY_0,'[^;]+',1,7) INT_ADD_0, " & _
			"REGEXP_SUBSTR(DAY_0,'[^;]+',1,8) INT_PICK_0, " & _
			"REGEXP_SUBSTR(DAY_0,'[^;]+',1,9) INT_SRUN_0, " & _
			"REGEXP_SUBSTR(DAY_0,'[^;]+',1,10) INT_TRADE_0, " & _
			"REGEXP_SUBSTR(DAY_1,'[^;]+',1,1) TYPE_1, " & _
			"REGEXP_SUBSTR(DAY_1,'[^;]+',1,2) NOTES_1, " & _
			"REGEXP_SUBSTR(DAY_1,'[^;]+',1,3) PLUS_MINUS_1, " & _
			"REGEXP_SUBSTR(DAY_1,'[^;]+',1,4) PLUS_1, " & _
			"REGEXP_SUBSTR(DAY_1,'[^;]+',1,5) LUNCH_1, " & _
			"REGEXP_SUBSTR(DAY_1,'[^;]+',1,6) INT_DROP_1, " & _
			"REGEXP_SUBSTR(DAY_1,'[^;]+',1,7) INT_ADD_1, " & _
			"REGEXP_SUBSTR(DAY_1,'[^;]+',1,8) INT_PICK_1, " & _
			"REGEXP_SUBSTR(DAY_1,'[^;]+',1,9) INT_SRUN_1, " & _
			"REGEXP_SUBSTR(DAY_1,'[^;]+',1,10) INT_TRADE_1, " & _
			"REGEXP_SUBSTR(DAY_2,'[^;]+',1,1) TYPE_2, " & _
			"REGEXP_SUBSTR(DAY_2,'[^;]+',1,2) NOTES_2, " & _
			"REGEXP_SUBSTR(DAY_2,'[^;]+',1,3) PLUS_MINUS_2, " & _
			"REGEXP_SUBSTR(DAY_2,'[^;]+',1,4) PLUS_2, " & _
			"REGEXP_SUBSTR(DAY_2,'[^;]+',1,5) LUNCH_2, " & _
			"REGEXP_SUBSTR(DAY_2,'[^;]+',1,6) INT_DROP_2, " & _
			"REGEXP_SUBSTR(DAY_2,'[^;]+',1,7) INT_ADD_2, " & _
			"REGEXP_SUBSTR(DAY_2,'[^;]+',1,8) INT_PICK_2, " & _
			"REGEXP_SUBSTR(DAY_2,'[^;]+',1,9) INT_SRUN_2, " & _
			"REGEXP_SUBSTR(DAY_2,'[^;]+',1,10) INT_TRADE_2, " & _
			"REGEXP_SUBSTR(DAY_3,'[^;]+',1,1) TYPE_3, " & _
			"REGEXP_SUBSTR(DAY_3,'[^;]+',1,2) NOTES_3, " & _
			"REGEXP_SUBSTR(DAY_3,'[^;]+',1,3) PLUS_MINUS_3, " & _
			"REGEXP_SUBSTR(DAY_3,'[^;]+',1,4) PLUS_3, " & _
			"REGEXP_SUBSTR(DAY_3,'[^;]+',1,5) LUNCH_3, " & _
			"REGEXP_SUBSTR(DAY_3,'[^;]+',1,6) INT_DROP_3, " & _
			"REGEXP_SUBSTR(DAY_3,'[^;]+',1,7) INT_ADD_3, " & _
			"REGEXP_SUBSTR(DAY_3,'[^;]+',1,8) INT_PICK_3, " & _
			"REGEXP_SUBSTR(DAY_3,'[^;]+',1,9) INT_SRUN_3, " & _
			"REGEXP_SUBSTR(DAY_3,'[^;]+',1,10) INT_TRADE_3, " & _
			"REGEXP_SUBSTR(DAY_4,'[^;]+',1,1) TYPE_4, " & _
			"REGEXP_SUBSTR(DAY_4,'[^;]+',1,2) NOTES_4, " & _
			"REGEXP_SUBSTR(DAY_4,'[^;]+',1,3) PLUS_MINUS_4, " & _
			"REGEXP_SUBSTR(DAY_4,'[^;]+',1,4) PLUS_4, " & _
			"REGEXP_SUBSTR(DAY_4,'[^;]+',1,5) LUNCH_4, " & _
			"REGEXP_SUBSTR(DAY_4,'[^;]+',1,6) INT_DROP_4, " & _
			"REGEXP_SUBSTR(DAY_4,'[^;]+',1,7) INT_ADD_4, " & _
			"REGEXP_SUBSTR(DAY_4,'[^;]+',1,8) INT_PICK_4, " & _
			"REGEXP_SUBSTR(DAY_4,'[^;]+',1,9) INT_SRUN_4, " & _
			"REGEXP_SUBSTR(DAY_4,'[^;]+',1,10) INT_TRADE_4, " & _
			"REGEXP_SUBSTR(DAY_5,'[^;]+',1,1) TYPE_5, " & _
			"REGEXP_SUBSTR(DAY_5,'[^;]+',1,2) NOTES_5, " & _
			"REGEXP_SUBSTR(DAY_5,'[^;]+',1,3) PLUS_MINUS_5, " & _
			"REGEXP_SUBSTR(DAY_5,'[^;]+',1,4) PLUS_5, " & _
			"REGEXP_SUBSTR(DAY_5,'[^;]+',1,5) LUNCH_5, " & _
			"REGEXP_SUBSTR(DAY_5,'[^;]+',1,6) INT_DROP_5, " & _
			"REGEXP_SUBSTR(DAY_5,'[^;]+',1,7) INT_ADD_5, " & _
			"REGEXP_SUBSTR(DAY_5,'[^;]+',1,8) INT_PICK_5, " & _
			"REGEXP_SUBSTR(DAY_5,'[^;]+',1,9) INT_SRUN_5, " & _
			"REGEXP_SUBSTR(DAY_5,'[^;]+',1,10) INT_TRADE_5, " & _
			"REGEXP_SUBSTR(DAY_6,'[^;]+',1,1) TYPE_6, " & _
			"REGEXP_SUBSTR(DAY_6,'[^;]+',1,2) NOTES_6, " & _
			"REGEXP_SUBSTR(DAY_6,'[^;]+',1,3) PLUS_MINUS_6, " & _
			"REGEXP_SUBSTR(DAY_6,'[^;]+',1,4) PLUS_6, " & _
			"REGEXP_SUBSTR(DAY_6,'[^;]+',1,5) LUNCH_6, " & _
			"REGEXP_SUBSTR(DAY_6,'[^;]+',1,6) INT_DROP_6, " & _
			"REGEXP_SUBSTR(DAY_6,'[^;]+',1,7) INT_ADD_6, " & _
			"REGEXP_SUBSTR(DAY_6,'[^;]+',1,8) INT_PICK_6, " & _
			"REGEXP_SUBSTR(DAY_6,'[^;]+',1,9) INT_SRUN_6, " & _
			"REGEXP_SUBSTR(DAY_6,'[^;]+',1,10) INT_TRADE_6 " & _
			"FROM " & _
			"( " & _
				"SELECT " & _
				"OPS_SCN_DATE, " & _
				"OPS_SCN_INTERVAL, " & _
				"DECODE(STATUS,'PICK','PHONE','ADDT','PHONE',STATUS) " & _
				"|| ';' || " & _
				"NOTES " & _
				"|| ';' || " & _
				"CASE " & _
					"WHEN TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)) <= TO_DATE('12/2/2018','MM/DD/YYYY') THEN " & _
						"CASE " & _
							"WHEN INT_CHECK = 0 THEN 0 " & _
							"WHEN SPECIALTY_CHECK = 0 THEN -1 " & _
							"ELSE SIGN(PLUS_MINUS) " & _
						"END " & _
					"ELSE " & _
						"CASE " & _
							"WHEN STATUS = 'OFF' AND SPECIALTY_CHECK = 0 THEN -1 " & _
							"WHEN STATUS = 'OFF' THEN SIGN(PLUS_MINUS+4) " & _
							"WHEN STATUS IN ('PICK','PHONE','ADDT') AND INT_CHECK = 0 THEN 0 " & _
							"WHEN STATUS IN ('PICK','PHONE','ADDT') AND SPECIALTY_CHECK = 0 THEN -1 " & _
							"ELSE SIGN(PLUS_MINUS) " & _
						"END " & _
				"END " & _
				"|| ';' || " & _
				"CASE " & _
					"WHEN INT_CHECK = 0 THEN 0.5 " & _
					"WHEN SPECIALTY_CHECK = 0 THEN -1 " & _
					"ELSE SIGN(CEILING) " & _
				"END " & _
				"|| ';' || " & _
				"LUNCH_CHECK " & _
				"|| ';' || " & _
				"CASE " & _
					"WHEN " & _
						"STATUS IN ('PICK','PHONE') " & _
						"AND INT_CHECK = 1 " & _
						"AND 'LED' <> ? " & _
						"AND " & _
						"( " & _
							"DPOPEN.DPO_DOTW IS NOT NULL " & _
							"OR " & _
							"( " & _
								"DPCLOSED.DPC_DOTW IS NULL " & _
								"AND BUFFER > 0 " & _
								"AND SPECIALTY_CHECK = 1 " & _
							") " & _
						") " & _
					"THEN 1 " & _
					"ELSE -1 " & _
				"END " & _
				"|| ';' || " & _
				"CASE " & _
					"WHEN " & _
						"STATUS = 'OFF' " & _
						"AND INT_CHECK = 1 " & _
						"AND " & _
						"( " & _
							"ADDOPEN.ADO_DOTW IS NOT NULL " & _
							"OR " & _
							"( " & _
								"ADDCLOSED.ADC_DOTW IS NULL " & _
								"AND " & _
								"( " & _
									"PLUS_MINUS <= 0 " & _
									"OR SPECIALTY_CHECK = 0 " & _
								") " & _
							") " & _
						") " & _
					"THEN 1 " & _
					"ELSE -1 " & _
				"END " & _
				"|| ';' || " & _
				"CASE " & _
					"WHEN " & _
						"STATUS = 'PICK' " & _
						"AND INT_CHECK = 1 " & _
					"THEN 1 " & _
					"WHEN " & _
						"STATUS IN ('LUNCH','OFF')" & _
						"AND INT_CHECK = 1 " & _
						"AND PICKCLOSED.PICK_DOTW IS NULL " & _
					"THEN 1 " & _
					"ELSE -1 " & _
				"END " & _
				"|| ';' || " & _
				"NVL2(SRUNAVAIL.SRA_DOTW,0,1) " & _
				"|| ';' || " & _
				"NVL2(SELFTRADEOFF.STO_DOTW,0,1) " & _
				"INTERVAL_DATA " & _
				"FROM " & _
				"( " & _
					"SELECT " & _
					"OPS_SCN_DATE, " & _
					"OPS_SCN_INTERVAL, " & _
					"PLUS_MINUS, " & _
					"BUFFER, " & _
					"CEILING, " & _
					"STATUS, " & _
					"CASE " & _
						"WHEN STATUS = 'PHONE' AND NOTES = 'Off' THEN 'Phone Time' " & _
						"WHEN STATUS = 'PICK' AND NOTES = 'Off' THEN 'Pick Hours' " & _
						"WHEN STATUS = 'ADDT' AND NOTES = 'Off' THEN 'Additional Time' " & _
						"WHEN STATUS = 'TRAIN' AND NOTES = 'Off' THEN 'Training/Meeting/Presentation' " & _
						"WHEN STATUS = 'SRED' AND NOTES = 'Off' THEN 'Schedule Reduction' " & _
						"WHEN STATUS = 'LUNCH' AND NOTES = 'Off' THEN 'Lunch' " & _
						"WHEN STATUS = 'VACA' AND NOTES = 'Off' THEN 'Approved Time Off' " & _
						"ELSE REPLACE(NOTES,';',':') " & _
					"END NOTES, " & _
					"INT_CHECK," & _
					"CASE " & _
						"WHEN " & _
							"STATUS = 'LUNCH' " & _
							"AND LAST_VALUE(NULLIF(STATUS,'LUNCH') IGNORE NULLS) OVER (PARTITION BY OPS_SCN_DATE ORDER BY OPS_SCN_INTERVAL ROWS BETWEEN UNBOUNDED PRECEDING AND 1 PRECEDING) IN ('PHONE','TRAIN','SRED','VACA') " & _
							"AND FIRST_VALUE(NULLIF(STATUS,'LUNCH') IGNORE NULLS) OVER (PARTITION BY OPS_SCN_DATE ORDER BY OPS_SCN_INTERVAL ROWS BETWEEN 1 FOLLOWING AND UNBOUNDED FOLLOWING) IN ('PHONE','TRAIN','SRED','VACA') " & _
						"THEN 0 " & _
						"ELSE 1 " & _
					"END LUNCH_CHECK, " & _
					"NVL(SPECIALTY_CHECK,1) SPECIALTY_CHECK " & _
					"FROM " & _
					"( " & _
						"SELECT " & _
						"OPS_SCN_DATE, " & _
						"OPS_SCN_INTERVAL, " & _
						"CASE " & _
							"WHEN RD.OPS_PAR_VALUE = 'MAX' THEN MAX(OPS_SCN_STAFFED-OPS_SCN_PROJECTION) " & _
							"ELSE MIN(OPS_SCN_STAFFED-OPS_SCN_PROJECTION) " & _
						"END PLUS_MINUS, " & _
						"CASE " & _
							"WHEN RD.OPS_PAR_VALUE = 'MAX' THEN MAX(OPS_SCN_STAFFED-OPS_SCN_PROJECTION-OPS_SCN_BUFFER) " & _
							"ELSE MIN(OPS_SCN_STAFFED-OPS_SCN_PROJECTION-OPS_SCN_BUFFER) " & _
						"END BUFFER, " & _
						"CASE " & _
							"WHEN RD.OPS_PAR_VALUE = 'MAX' THEN MAX(OPS_SCN_STAFFED-OPS_SCN_PROJECTION-OPS_SCN_CEILING) " & _
							"ELSE MIN(OPS_SCN_STAFFED-OPS_SCN_PROJECTION-OPS_SCN_CEILING) " & _
						"END CEILING, " & _
						"MAX(CASE " & _
							"WHEN OPS_SCI_TYPE = 'PICK' THEN 'PICK' " & _
							"WHEN OPS_SCI_TYPE = 'ADDT' THEN 'ADDT' " & _
							"WHEN OPS_SCI_TYPE IN ('BASE','HOLW','EXTD') THEN 'PHONE' " & _
							"WHEN OPS_SCI_TYPE IN ('MEET','PRES','PROJ','TRAN','FAMP','WFHU','MLTU','OTRG','NEWH') THEN 'TRAIN' " & _
							"WHEN OPS_SCI_TYPE IN ('SRPT','SRUN') THEN 'SRED' " & _
							"WHEN OPS_SCI_TYPE IN ('LNCH','LNFL') THEN 'LUNCH' " & _
							"WHEN REGEXP_LIKE(OPS_SCI_TYPE,'^VAC|UN$|PT$|PP$|HOLU|SLIP|RESH|RCHG|ROUT|JURY|BRVT') OR (OPS_SCI_TYPE = 'HOLR' AND RES_BUE_ID IS NULL) THEN 'VACA' " & _
							"WHEN OPS_SCI_TYPE IS NULL OR (OPS_SCI_TYPE = 'HOLR' AND RES_BUE_ID IS NOT NULL) THEN 'OFF' " & _
						"END) KEEP (DENSE_RANK FIRST ORDER BY NULLIF(OPS_SCI_TYPE,'HOLR') NULLS LAST, OPS_SCI_START) STATUS, " & _
						"MAX(NVL(OPS_SCI_NOTES,'Off')) KEEP (DENSE_RANK FIRST ORDER BY NULLIF(OPS_SCI_TYPE,'HOLR') NULLS LAST, OPS_SCI_START) NOTES, " & _
						"DECODE(MAX(ROUND(NVL(1440*(LEAST(TO_DATE(TO_CHAR(OPS_SCI_END-(1/1440),'HH24:MI'),'HH24:MI'),TO_DATE(OPS_SCN_INTERVAL,'HH24:MI')+29/1440) - GREATEST(TO_DATE(TO_CHAR(OPS_SCI_START,'HH24:MI'),'HH24:MI'), TO_DATE(OPS_SCN_INTERVAL,'HH24:MI')))+1,0))) KEEP (DENSE_RANK FIRST ORDER BY NULLIF(OPS_SCI_TYPE,'HOLR') NULLS LAST, OPS_SCI_START),0,1,30,1,0) INT_CHECK " & _
						"FROM OPS_SCHEDULE_NEED " & _
						"JOIN " & _
						"( " & _
							"SELECT " & _
							"OPS_PAR_VALUE SCHEDULE_TYPE " & _
							"FROM " & _
							"( " & _
								"SELECT * FROM OPS_PARAMETER " & _
								"WHERE OPS_PAR_CODE = 'PARENT_WORKGROUP' " & _
								"AND TO_DATE(?,'MM/DD/YYYY') BETWEEN OPS_PAR_EFF_DATE AND OPS_PAR_DIS_DATE " & _
							")PAR " & _
							"START WITH OPS_PAR_PARENT_TYPE = ? " & _
							"CONNECT BY OPS_PAR_PARENT_TYPE = PRIOR OPS_PAR_VALUE " & _
							"AND OPS_PAR_ID <> PRIOR OPS_PAR_ID " & _
							"UNION " & _
							"SELECT ? FROM DUAL " & _
						") WG " & _
						"ON OPS_SCN_TYPE = WG.SCHEDULE_TYPE " & _
						"LEFT JOIN OPS_PARAMETER RD " & _
						"ON RD.OPS_PAR_PARENT_TYPE = ? " & _
						"AND RD.OPS_PAR_CODE = 'ROLLUP_DIRECTION' " & _
						"AND OPS_SCN_DATE BETWEEN RD.OPS_PAR_EFF_DATE AND RD.OPS_PAR_DIS_DATE " & _
						"LEFT JOIN OPS_SCHEDULE_INFO " & _
						"ON OPS_SCN_DATE = TO_DATE(OPS_SCI_START) " & _
						"AND OPS_SCI_STATUS = 'APP' " & _
						"AND TO_CHAR(OPS_SCI_START,'HH24:MI') <= TO_CHAR(TO_DATE(OPS_SCN_INTERVAL,'HH24:MI')+29/1440,'HH24:MI') " & _
						"AND TO_CHAR(OPS_SCI_END-(1/1440),'HH24:MI') >= OPS_SCN_INTERVAL " & _
						"AND OPS_SCI_OPS_USR_ID = ? " & _
						"LEFT JOIN RES_BUDGET_EXCEPTION " & _
						"ON OPS_SCN_DATE = RES_BUE_DATE " & _
						"AND RES_BUE_TYPE = 'NOR' " & _
						"WHERE OPS_SCN_DATE BETWEEN TO_DATE(?,'MM/DD/YYYY') AND TO_DATE(?,'MM/DD/YYYY') " & _
						"GROUP BY OPS_SCN_DATE, OPS_SCN_INTERVAL, RD.OPS_PAR_VALUE " & _
					") " & _
					"LEFT JOIN " & _
					"( " & _
						"SELECT " & _
						"OPS_SCN_DATE SPECIALTY_DATE, " & _
						"OPS_SCN_INTERVAL SPECIALTY_INTERVAL, " & _
						"DECODE(SIGN(COUNT(RES_RTE_ID)),1,1,0) SPECIALTY_CHECK " & _
						"FROM OPS_SCHEDULE_NEED " & _
						"JOIN OPS_PARAMETER " & _
						"ON OPS_PAR_CODE = 'ROUTING_CHECK' " & _
						"AND OPS_SCN_TYPE = OPS_PAR_PARENT_TYPE " & _
						"AND OPS_SCN_DATE BETWEEN OPS_PAR_EFF_DATE AND OPS_PAR_DIS_DATE " & _
						"LEFT JOIN OPS_SCHEDULE_INFO " & _
						"ON OPS_SCN_DATE = TO_DATE(OPS_SCI_START) " & _
						"AND OPS_SCI_STATUS = 'APP' " & _
						"AND OPS_SCI_TYPE IN ('BASE','HOLW','ADDT','EXTD','PICK') " & _
						"AND TO_CHAR(OPS_SCI_START,'HH24:MI') <= TO_CHAR(TO_DATE(OPS_SCN_INTERVAL,'HH24:MI')+29/1440,'HH24:MI') " & _
						"AND TO_CHAR(OPS_SCI_END-(1/1440),'HH24:MI') >= OPS_SCN_INTERVAL " & _
						"AND OPS_SCI_OPS_USR_ID <> ? " & _
						"LEFT JOIN RES_ROUTING " & _
						"ON OPS_SCI_OPS_USR_ID = RES_RTE_OPS_USR_ID " & _
						"AND OPS_PAR_VALUE = DECODE(RES_RTE_RES_RTG_ID,3,1,2,1,RES_RTE_RES_RTG_ID) " & _
						"AND RES_RTE_YEAR = TO_CHAR(TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE))-(6/24),'YYYY') " & _
						"AND RES_RTE_MONTH = TO_CHAR(TO_DATE(CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE))-(6/24),'MM') " & _
						"WHERE OPS_SCN_TYPE = ? " & _
						"AND OPS_PAR_VALUE = ? " & _
						"AND OPS_SCN_DATE BETWEEN TO_DATE(?,'MM/DD/YYYY') AND TO_DATE(?,'MM/DD/YYYY') " & _
						"GROUP BY OPS_SCN_DATE, OPS_SCN_INTERVAL " & _
					") " & _
					"ON OPS_SCN_DATE = SPECIALTY_DATE " & _
					"AND OPS_SCN_INTERVAL = SPECIALTY_INTERVAL " & _
				") " & _
				"LEFT JOIN " & _
				"( " & _
					"SELECT " & _
					"REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1) DPC_DOTW, " & _
					"REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) DPC_START, " & _
					"REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3) DPC_END, " & _
					"OPS_PAR_EFF_DATE, " & _
					"OPS_PAR_DIS_DATE " & _
					"FROM OPS_PARAMETER " & _
					"WHERE OPS_PAR_CODE = 'DROP_CLOSED' " & _
					"AND OPS_PAR_PARENT_TYPE = DECODE(?,'SRV','RES','SLS','RES',?) " & _
				")DPCLOSED " & _
				"ON INSTR(DPCLOSED.DPC_DOTW,TO_CHAR(OPS_SCN_DATE,'D')) > 0 " & _
				"AND OPS_SCN_DATE BETWEEN DPCLOSED.OPS_PAR_EFF_DATE AND DPCLOSED.OPS_PAR_DIS_DATE " & _
				"AND OPS_SCN_INTERVAL BETWEEN DPCLOSED.DPC_START AND DPCLOSED.DPC_END " & _
				"AND OPS_SCN_INTERVAL <> DPCLOSED.DPC_END " & _
				"LEFT JOIN " & _
				"( " & _
					"SELECT " & _
					"REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1) DPO_DOTW, " & _
					"REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) DPO_START, " & _
					"REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3) DPO_END, " & _
					"OPS_PAR_EFF_DATE, " & _
					"OPS_PAR_DIS_DATE " & _
					"FROM OPS_PARAMETER " & _
					"WHERE OPS_PAR_CODE = 'DROP_OPEN' " & _
					"AND OPS_PAR_PARENT_TYPE = ? " & _
				")DPOPEN " & _
				"ON INSTR(DPOPEN.DPO_DOTW,TO_CHAR(OPS_SCN_DATE,'D')) > 0 " & _
				"AND OPS_SCN_DATE BETWEEN DPOPEN.OPS_PAR_EFF_DATE AND DPOPEN.OPS_PAR_DIS_DATE " & _
				"AND OPS_SCN_INTERVAL BETWEEN DPOPEN.DPO_START AND DPOPEN.DPO_END " & _
				"AND OPS_SCN_INTERVAL <> DPOPEN.DPO_END " & _
				"LEFT JOIN " & _
				"( " & _
					"SELECT " & _
					"REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1) SRA_DOTW, " & _
					"REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) SRA_START, " & _
					"REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3) SRA_END, " & _
					"OPS_PAR_EFF_DATE, " & _
					"OPS_PAR_DIS_DATE " & _
					"FROM OPS_PARAMETER " & _
					"WHERE OPS_PAR_CODE = 'SRUN_UNAVAILABLE' " & _
					"AND OPS_PAR_PARENT_TYPE = DECODE(?,'SRV','RES','SLS','RES',?) " & _
				")SRUNAVAIL " & _
				"ON INSTR(SRUNAVAIL.SRA_DOTW,TO_CHAR(OPS_SCN_DATE,'D')) > 0 " & _
				"AND OPS_SCN_DATE BETWEEN SRUNAVAIL.OPS_PAR_EFF_DATE AND SRUNAVAIL.OPS_PAR_DIS_DATE " & _
				"AND OPS_SCN_INTERVAL BETWEEN SRUNAVAIL.SRA_START AND SRUNAVAIL.SRA_END " & _
				"AND OPS_SCN_INTERVAL <> SRUNAVAIL.SRA_END " & _
				"LEFT JOIN " & _
				"( " & _
					"SELECT " & _
					"REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1) ADO_DOTW, " & _
					"REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) ADO_START, " & _
					"REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3) ADO_END, " & _
					"OPS_PAR_EFF_DATE, " & _
					"OPS_PAR_DIS_DATE " & _
					"FROM OPS_PARAMETER " & _
					"WHERE OPS_PAR_CODE = 'ADD_OPEN' " & _
					"AND OPS_PAR_PARENT_TYPE = ? " & _
				")ADDOPEN " & _
				"ON INSTR(ADDOPEN.ADO_DOTW,TO_CHAR(OPS_SCN_DATE,'D')) > 0 " & _
				"AND OPS_SCN_DATE BETWEEN ADDOPEN.OPS_PAR_EFF_DATE AND ADDOPEN.OPS_PAR_DIS_DATE " & _
				"AND OPS_SCN_INTERVAL BETWEEN ADDOPEN.ADO_START AND ADDOPEN.ADO_END " & _
				"AND OPS_SCN_INTERVAL <> ADDOPEN.ADO_END " & _
				"LEFT JOIN " & _
				"( " & _
					"SELECT " & _
					"REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1) ADC_DOTW, " & _
					"REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) ADC_START, " & _
					"REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3) ADC_END, " & _
					"OPS_PAR_EFF_DATE, " & _
					"OPS_PAR_DIS_DATE " & _
					"FROM OPS_PARAMETER " & _
					"WHERE OPS_PAR_CODE = 'ADD_CLOSED' " & _
					"AND OPS_PAR_PARENT_TYPE = ? " & _
				")ADDCLOSED " & _
				"ON INSTR(ADDCLOSED.ADC_DOTW,TO_CHAR(OPS_SCN_DATE,'D')) > 0 " & _
				"AND OPS_SCN_DATE BETWEEN ADDCLOSED.OPS_PAR_EFF_DATE AND ADDCLOSED.OPS_PAR_DIS_DATE " & _
				"AND OPS_SCN_INTERVAL BETWEEN ADDCLOSED.ADC_START AND ADDCLOSED.ADC_END " & _
				"AND OPS_SCN_INTERVAL <> ADDCLOSED.ADC_END " & _
				"LEFT JOIN " & _
				"( " & _
					"SELECT " & _
					"REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1) PICK_DOTW, " & _
					"REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) PICK_START, " & _
					"REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3) PICK_END, " & _
					"OPS_PAR_EFF_DATE, " & _
					"OPS_PAR_DIS_DATE " & _
					"FROM OPS_PARAMETER " & _
					"WHERE OPS_PAR_CODE = 'PICK_CLOSED' " & _
					"AND OPS_PAR_PARENT_TYPE = ? " & _
				")PICKCLOSED " & _
				"ON INSTR(PICKCLOSED.PICK_DOTW,TO_CHAR(OPS_SCN_DATE,'D')) > 0 " & _
				"AND OPS_SCN_DATE BETWEEN PICKCLOSED.OPS_PAR_EFF_DATE AND PICKCLOSED.OPS_PAR_DIS_DATE " & _
				"AND OPS_SCN_INTERVAL BETWEEN PICKCLOSED.PICK_START AND PICKCLOSED.PICK_END " & _
				"AND OPS_SCN_INTERVAL <> PICKCLOSED.PICK_END " & _
				"LEFT JOIN " & _
				"( " & _
					"SELECT " & _
					"REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1) STO_DOTW, " & _
					"REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) STO_START, " & _
					"REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3) STO_END, " & _
					"OPS_PAR_EFF_DATE, " & _
					"OPS_PAR_DIS_DATE " & _
					"FROM OPS_PARAMETER " & _
					"WHERE OPS_PAR_CODE = 'SELFTRADE_OFF' " & _
					"AND OPS_PAR_PARENT_TYPE = ? " & _
				")SELFTRADEOFF " & _
				"ON INSTR(SELFTRADEOFF.STO_DOTW,TO_CHAR(OPS_SCN_DATE,'D')) > 0 " & _
				"AND OPS_SCN_DATE BETWEEN SELFTRADEOFF.OPS_PAR_EFF_DATE AND SELFTRADEOFF.OPS_PAR_DIS_DATE " & _
				"AND OPS_SCN_INTERVAL BETWEEN SELFTRADEOFF.STO_START AND SELFTRADEOFF.STO_END " & _
				"AND OPS_SCN_INTERVAL <> SELFTRADEOFF.STO_END " & _
			") " & _
			"PIVOT (MAX(INTERVAL_DATA) FOR OPS_SCN_DATE IN (TO_DATE('" & SCHEDULE_START_DATE & "','MM/DD/YYYY') DAY_0, TO_DATE('" & SCHEDULE_START_DATE + 1 & "','MM/DD/YYYY') DAY_1, TO_DATE('" & SCHEDULE_START_DATE + 2 & "','MM/DD/YYYY') DAY_2, TO_DATE('" & SCHEDULE_START_DATE + 3 & "','MM/DD/YYYY') DAY_3, TO_DATE('" & SCHEDULE_START_DATE + 4 & "','MM/DD/YYYY') DAY_4, TO_DATE('" & SCHEDULE_START_DATE + 5 & "','MM/DD/YYYY') DAY_5, TO_DATE('" & SCHEDULE_START_DATE + 6 & "','MM/DD/YYYY') DAY_6)) " & _
			"ORDER BY OPS_SCN_INTERVAL"
			cmd.CommandText = SQLstmt
			cmd.Parameters(0).value = AGENT_JOB
			cmd.Parameters(1).value = SCHEDULE_START_DATE
			cmd.Parameters(2).value = AGENT_TYPE
			cmd.Parameters(3).value = AGENT_TYPE
			cmd.Parameters(4).value = AGENT_TYPE
			cmd.Parameters(5).value = SCHEDULE_USR_ID
			cmd.Parameters(6).value = SCHEDULE_START_DATE
			cmd.Parameters(7).value = SCHEDULE_END_DATE
			cmd.Parameters(8).value = SCHEDULE_USR_ID
			cmd.Parameters(9).value = AGENT_TYPE
			cmd.Parameters(10).value = AGENT_ROUTING
			cmd.Parameters(11).value = SCHEDULE_START_DATE
			cmd.Parameters(12).value = SCHEDULE_END_DATE
			cmd.Parameters(13).value = AGENT_TYPE
			cmd.Parameters(14).value = AGENT_TYPE
			cmd.Parameters(15).value = AGENT_TYPE
			cmd.Parameters(16).value = AGENT_TYPE
			cmd.Parameters(17).value = AGENT_TYPE
			cmd.Parameters(18).value = AGENT_TYPE
			cmd.Parameters(19).value = AGENT_TYPE
			cmd.Parameters(20).value = AGENT_TYPE
			cmd.Parameters(21).value = AGENT_TYPE
			Set RSINT = cmd.Execute
		%>
		<% If Not RSINT.EOF Then %>
			<form id="SCHEDULE_FORM" name="SCHEDULE_FORM" method="post">
				<input type="hidden" id="FILTER_DATE" name="FILTER_DATE" value="<%=SCHEDULE_START_DATE%>"/>
				<input type="hidden" id="FILTER_AGENT" name="FILTER_AGENT" value="<%=SCHEDULE_USR_ID%>"/>
				<input type="hidden" id="AGENT_TYPE" name="AGENT_TYPE" value="<%=AGENT_TYPE%>"/>
				<input type="hidden" id="SUBMIT_FLAG" name="SUBMIT_FLAG" value="1"/>
				<input type="hidden" id="LUNCH_WAIVER" name="LUNCH_WAIVER" value=""/>
				<input type="hidden" name="USE_LAYOUT" value="<%=SCHEDSQUATCH_LAYOUT%>"/>
				<input type="hidden" name="STPLUS_ENABLED" value="<%=STPLUS_ENABLED%>"/>
				<% If Len(OPEN_DATE) >= 1 Then %>
					<input type="hidden" id="SCHEDULE_SELECT" value="<%=Replace(Replace(Replace(Left(OPEN_DATE,1),"M","AD"),"S","ST"),"P","PK")%>">
					<% USE_SCHEDULE = Replace(Replace(Replace(Left(OPEN_DATE,1),"M","AD"),"S","ST"),"P","PK") %>
					<div id="SCHEDSQUATCH_TABS" style="display:none;height:30px;margin-left:5px;">
						<% For i = 1 to Len(OPEN_DATE) %>
							<% SCHEDULE_TYPE_PREFIX = Replace(Replace(Replace(Mid(OPEN_DATE,i,1),"M","AD"),"S","ST"),"P","PK") %>
							<% SCHEDULE_TYPE_NAME = Replace(Replace(Replace(Mid(OPEN_DATE,i,1),"M","Add/Drop"),"S","Self Trade"),"P","Pick Hours") %>
							<div id="<%=SCHEDULE_TYPE_PREFIX%>_TAB" class="tab<% If i = 1 Then %>selected<% End If %>"><%=SCHEDULE_TYPE_NAME%></div>
						<% Next %>
					</div>
				<% Else %>
					<input type="hidden" id="SCHEDULE_SELECT" value="">
					<% USE_SCHEDULE = "" %>
				<% End If %>

				<table id="SCHEDSQUATCH" style="display:none;">
					<thead>
						<tr>
							<th style="text-align:center;font-weight:900;width:48px;">
								<i id="SCHEDSQUATCH_LAYOUT" class="fa fa-arrows-<%=Replace(Replace(SCHEDSQUATCH_LAYOUT,"V","h"),"H","v")%>" style="cursor:pointer;font-size:22pt;color:#00284E;" aria-hidden="true"></i>
							</th>
						<% For i = 0 to 6 %>
							<% USE_CLASS = "" %>
							<% If DROP_DATE(i) = 0 and ADD_DATE(i) = 0 Then %>
								<% USE_CLASS = " ADCLOSED" %>
							<% End If %>
							<% If SELFTRADE_DATE(i) = 0 Then %>
								<% USE_CLASS = USE_CLASS & " STCLOSED" %>
								<% If STPLUS_AVAILABLE = "Y" Then %>
									<% USE_CLASS = USE_CLASS & " SPCLOSED" %>
								<% End If %>
							<% End If %>
							<% If PICK_DATE(i) = 0 Then %>
								<% USE_CLASS = USE_CLASS & " PKCLOSED" %>
							<% End If %>
							<th id="SCHEDCLEAR_<%=i%>" class="scheduleHeading center <%=USE_SCHEDULE%><%=USE_CLASS%>">
								<span style="white-space:nowrap;">
									<%=Month(SCHEDULE_START_DATE + i) & "/" & Day(SCHEDULE_START_DATE + i) & " " & WeekdayName(DatePart("w",SCHEDULE_START_DATE + i),True)%>
								</span>
							</th>
						<% Next %>
						</tr>
					</thead>
					<% PREVIOUS_INTERVAL = "" %>
					<tbody>
						<% Do While Not RSINT.EOF %>
							<% CURRENT_INTERVAL = RSINT("OPS_SCN_INTERVAL") %>
							<% CURRENT_INTERVAL_DISPLAY = RSINT("CUR_DISP_INTERVAL") %>
							<% NEXT_INTERVAL_DISPLAY = RSINT("NEXT_DISP_INTERVAL") %>

							<% If PREVIOUS_INTERVAL <> "" Then %>
								<% If DateDiff("n",CDate(PREVIOUS_INTERVAL),CDate(CURRENT_INTERVAL)) > 30 Then %>
									<tr>
										<td style="background-color:#000;height:10px;border-radius:0.6em;" colspan="8"></td>
									</tr>
								<% End If %>
							<% End If %>
							<tr>
								<td id="SCHEDCLEAR_<%=Replace(CURRENT_INTERVAL,":","")%>" class="scheduleHeading center">
									<span style="white-space:nowrap;"><%=CURRENT_INTERVAL_DISPLAY%></span>
									<br/>
									<span style="white-space:nowrap;"><%=NEXT_INTERVAL_DISPLAY%></span>
								</td>
								<% For i = 0 to 6 %>
									<% USE_CLASS = "" %>
									<% SCHEDULE_TYPE = RSINT("TYPE_" & i) %>
									<% PLUS_MINUS = RSINT("PLUS_MINUS_" & i) %>
									<% ADD_FLAG = RSINT("INT_ADD_" & i) %>
									<% DROP_FLAG = RSINT("INT_DROP_" & i) %>
									<% SRUN_AVAILABLE = RSINT("INT_SRUN_" & i) %>
									<% TRADE_AVAILABLE = RSINT("INT_TRADE_" & i) %>
									<% PICK_FLAG = RSINT("INT_PICK_" & i) %>
									<% PLUS_FLAG = RSINT("PLUS_" & i) %>
									<% LUNCH_FLAG = RSINT("LUNCH_" & i) %>
									<% CURRENT_NOTES = RSINT("NOTES_" & i) %>

									<% If (DROP_DATE(i) = 0 and ADD_DATE(i) = 0) or CDate(SCHEDULE_START_DATE + i & " " & CURRENT_INTERVAL) <= DateAdd("h",1,Now) Then %>
										<% USE_CLASS = USE_CLASS & " ADCLOSED" %>
									<% End If %>
									<% If SELFTRADE_DATE(i) = 0 or CDate(SCHEDULE_START_DATE + i & " " & CURRENT_INTERVAL) <= DateAdd("h",1,Now) Then %>
										<% USE_CLASS = USE_CLASS & " STCLOSED" %>
										<% If STPLUS_AVAILABLE = "Y" Then %>
											<% USE_CLASS = USE_CLASS & " SPCLOSED" %>
										<% End If %>
									<% End If %>
									<% If PICK_DATE(i) = 0 or CDate(SCHEDULE_START_DATE + i & " " & CURRENT_INTERVAL) <= DateAdd("h",1,Now) Then %>
										<% USE_CLASS = USE_CLASS & " PKCLOSED" %>
									<% End If %>
									<% If SCHEDULE_TYPE <> "TRAIN" and SCHEDULE_TYPE <> "VACA" and SCHEDULE_TYPE <> "SRED" Then %>
										<% If ADD_DATE(i) = 1 and ADD_FLAG = 1 and AGENT_TEAM <> "NEW" Then %>
											<% USE_CLASS = USE_CLASS & " ADADD" %>
										<% Else %>
											<% USE_CLASS = USE_CLASS & " AD" & SCHEDULE_TYPE %>
										<% End If %>

										<% If SELFTRADE_DATE(i) = 1 and PLUS_MINUS <= 0 and TRADE_AVAILABLE = 1 and AGENT_TEAM <> "NEW" and SCHEDULE_TYPE = "OFF" Then %>
											<% USE_CLASS = USE_CLASS & " STADD" %>
										<% Else %>
											<% USE_CLASS = USE_CLASS & " ST" & SCHEDULE_TYPE %>
										<% End If %>

										<% If STPLUS_AVAILABLE = "Y" Then %>
											<% If Date <= CDate("12/2/2018") and SELFTRADEPLUS_DATE(i) = 1 and PLUS_FLAG <= 0 and PLUS_MINUS = 1 and TRADE_AVAILABLE = 1 and AGENT_TEAM <> "NEW" and SCHEDULE_TYPE = "OFF" Then %>
												<% USE_CLASS = USE_CLASS & " SPPLUSADD" %>
											<% Elseif Date <= CDate("12/2/2018") and SELFTRADEPLUS_DATE(i) = 1 and PLUS_FLAG = 1 and TRADE_AVAILABLE = 1 and AGENT_TEAM <> "NEW" and SCHEDULE_TYPE = "PHONE" Then %>
												<% USE_CLASS = USE_CLASS & " SPPLUSPHONE" %>
											<% Elseif SELFTRADE_DATE(i) = 1 and PLUS_MINUS <= 0 and TRADE_AVAILABLE = 1 and AGENT_TEAM <> "NEW" and SCHEDULE_TYPE = "OFF" Then %>
												<% USE_CLASS = USE_CLASS & " SPADD" %>
											<% Else %>
												<% USE_CLASS = USE_CLASS & " SP" & SCHEDULE_TYPE %>
											<% End If %>
										<% End If %>

										<% USE_CLASS = USE_CLASS & " PK" & SCHEDULE_TYPE %>
									<% Else %>
										<% USE_CLASS = USE_CLASS & " " & SCHEDULE_TYPE %>
									<% End If %>

									<td id="TD_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" class="classicScheduleTd center <%=USE_SCHEDULE%> <%=USE_CLASS%>" title="<%=CURRENT_NOTES%>">
										<!--Only these inputs have names, making them the only submitted inputs -->
											<input id="SCHEDDATA_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" name="SCHEDDATA_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" type="hidden" value=""/>
										<!---->
										<!--AD-->
											<% If ADD_DATE(i) = 1 and ADD_FLAG = 1 and AGENT_TEAM <> "NEW" and CDate(SCHEDULE_START_DATE + i & " " & CURRENT_INTERVAL) > DateAdd("h",1,Now) Then %>
												<input id="ADVALUE_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" type="hidden" value="ADADD;<%=CURRENT_NOTES%>"/>
												<input id="ADADD_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" class="<%=USE_SCHEDULE%> ADADD" type="checkbox" value="ADADD"/>
											<% Elseif Date <= CDate("12/2/2018") and ADD_DATE(i) = 1 and SCHEDULE_TYPE = "OFF" and CDate(SCHEDULE_START_DATE + i & " " & CURRENT_INTERVAL) > DateAdd("h",1,Now) and (ADD_FLAG <> 1 or AGENT_TEAM = "NEW") Then %>
												<input id="ADVALUE_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" type="hidden" value="ADOFF;<%=CURRENT_NOTES%>"/>
												<input id="ADOFF_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" class="<%=USE_SCHEDULE%> ADOFF" type="checkbox" value="ADOFF"/>
											<% Elseif DROP_DATE(i) = 1 and CDate(SCHEDULE_START_DATE + i & " " & CURRENT_INTERVAL) > DateAdd("h",1,Now) and DROP_FLAG = 1 and DAY_CHECK(i) = 1 and Instr(USE_CLASS,"PKCLOSED") > 0 and CDate(SCHEDULE_START_DATE + i & " " & CURRENT_INTERVAL) <= DROP_MAX_INTERVAL Then %>
												<input id="ADVALUE_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" type="hidden" value="ADPHONE;<%=CURRENT_NOTES%>"/>
												<select id="ADSRED_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" style="font-family:Calibri;font-size:8pt;font-weight:900;" class="<%=USE_SCHEDULE%> ADPHONE">
													<option value="ADPHONE" class="<%=USE_SCHEDULE%> ADPHONE"></option>
													<% If SRUN_AVAILABLE = 1 Then %>
														<option id="ADSRUN_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" value="SRUN" class="<%=USE_SCHEDULE%> ADSRED">UNP</option>
													<% End If %>
													<option id="ADSRPT_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" value="SRPT" class="<%=USE_SCHEDULE%> ADSRED">PTO</option>
												</select>
											<% End If %>
										<!---->
										<!--ST-->
											<% If SELFTRADE_DATE(i) = 1 and PLUS_MINUS <= 0 and TRADE_AVAILABLE = 1 and SCHEDULE_TYPE = "OFF" and AGENT_TEAM <> "NEW" and CDate(SCHEDULE_START_DATE + i & " " & CURRENT_INTERVAL) > DateAdd("h",1,Now) Then %>
												<input id="STVALUE_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" type="hidden" value="STADD;<%=CURRENT_NOTES%>"/>
												<input id="STADD_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" class="<%=USE_SCHEDULE%> STADD" type="checkbox" value="STADD"/>
											<% Elseif SELFTRADE_DATE(i) = 1 and PLUS_MINUS = 1 and TRADE_AVAILABLE = 1 and SCHEDULE_TYPE = "PHONE" and AGENT_TEAM <> "NEW" and Instr(USE_CLASS,"PKCLOSED") > 0 and CDate(SCHEDULE_START_DATE + i & " " & CURRENT_INTERVAL) > DateAdd("h",1,Now) Then %>
												<input id="STVALUE_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" type="hidden" value="STPHONE;<%=CURRENT_NOTES%>"/>
												<input id="STDROP_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" class="<%=USE_SCHEDULE%> STPHONE" type="checkbox" value="STPHONE"/>
											<% Elseif Date <= CDate("12/2/2018") and SELFTRADE_DATE(i) = 1 and (PLUS_MINUS = 1 or AGENT_TEAM = "NEW" or TRADE_AVAILABLE = 0) and SCHEDULE_TYPE = "OFF" and CDate(SCHEDULE_START_DATE + i & " " & CURRENT_INTERVAL) > DateAdd("h",1,Now) Then %>
												<input id="STVALUE_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" type="hidden" value="STOFF;<%=CURRENT_NOTES%>"/>
												<input id="STOFF_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" class="<%=USE_SCHEDULE%> STOFF" type="checkbox" value="STOFF"/>
											<% End If %>
										<!---->
										<!--SP-->
											<% If STPLUS_AVAILABLE = "Y" Then %>
												<% If Date <= CDate("12/2/2018") and SELFTRADEPLUS_DATE(i) = 1 and PLUS_FLAG <= 0 and PLUS_MINUS = 1 and TRADE_AVAILABLE = 1 and SCHEDULE_TYPE = "OFF" and AGENT_TEAM <> "NEW" and CDate(SCHEDULE_START_DATE + i & " " & CURRENT_INTERVAL) > DateAdd("h",1,Now) Then %>
													<input id="SPVALUE_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" type="hidden" value="SPPLUSADD;<%=CURRENT_NOTES%>"/>
													<input id="SPPLUSADD_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" class="<%=USE_SCHEDULE%> SPPLUSADD" type="checkbox" value="SPPLUSADD"/>
												<% Elseif Date <= CDate("12/2/2018") and SELFTRADEPLUS_DATE(i) = 1 and PLUS_FLAG = 1 and TRADE_AVAILABLE = 1 and SCHEDULE_TYPE = "PHONE" and AGENT_TEAM <> "NEW" and Instr(USE_CLASS,"PKCLOSED") > 0 and CDate(SCHEDULE_START_DATE + i & " " & CURRENT_INTERVAL) > DateAdd("h",1,Now) Then %>
													<input id="SPVALUE_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" type="hidden" value="SPPLUSPHONE;<%=CURRENT_NOTES%>"/>
													<input id="SPPLUSDROP_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" class="<%=USE_SCHEDULE%> SPPLUSPHONE" type="checkbox" value="SPPLUSPHONE"/>
												<% Elseif SELFTRADE_DATE(i) = 1 and PLUS_MINUS <= 0 and TRADE_AVAILABLE = 1 and SCHEDULE_TYPE = "OFF" and AGENT_TEAM <> "NEW" and CDate(SCHEDULE_START_DATE + i & " " & CURRENT_INTERVAL) > DateAdd("h",1,Now) Then %>
													<input id="SPVALUE_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" type="hidden" value="SPADD;<%=CURRENT_NOTES%>"/>
													<input id="SPADD_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" class="<%=USE_SCHEDULE%> SPADD" type="checkbox" value="SPADD"/>
												<% Elseif SELFTRADE_DATE(i) = 1 and PLUS_MINUS = 1 and TRADE_AVAILABLE = 1 and SCHEDULE_TYPE = "PHONE" and AGENT_TEAM <> "NEW" and Instr(USE_CLASS,"PKCLOSED") > 0 and CDate(SCHEDULE_START_DATE + i & " " & CURRENT_INTERVAL) > DateAdd("h",1,Now) Then %>
													<input id="SPVALUE_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" type="hidden" value="SPPHONE;<%=CURRENT_NOTES%>"/>
													<input id="SPDROP_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" class="<%=USE_SCHEDULE%> SPPHONE" type="checkbox" value="SPPHONE"/>
												<% Elseif Date <= CDate("12/2/2018") and SELFTRADE_DATE(i) = 1 and (PLUS_MINUS = 1 or AGENT_TEAM = "NEW" or TRADE_AVAILABLE = 0) and SCHEDULE_TYPE = "OFF" and CDate(SCHEDULE_START_DATE + i & " " & CURRENT_INTERVAL) > DateAdd("h",1,Now) Then %>
													<input id="SPVALUE_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" type="hidden" value="SPOFF;<%=CURRENT_NOTES%>"/>
													<input id="SPOFF_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" class="<%=USE_SCHEDULE%> SPOFF" type="checkbox" value="SPOFF"/>
												<% End If %>
											<% End If %>
										<!---->
										<!--PK-->
											<% If PICK_DATE(i) = 1 and PICK_FLAG = 1 and SCHEDULE_TYPE = "OFF" and CDate(SCHEDULE_START_DATE + i & " " & CURRENT_INTERVAL) > DateAdd("h",1,Now) Then %>
												<input id="PKVALUE_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" type="hidden" value="PKOFF;<%=CURRENT_NOTES%>"/>
												<input id="PKPICK_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" class="<%=USE_SCHEDULE%> PKOFF" type="checkbox" value="PKOFF"/>
											<% Elseif PICK_DATE(i) = 1 and PICK_FLAG = 1 and SCHEDULE_TYPE = "PHONE" and CDate(SCHEDULE_START_DATE + i & " " & CURRENT_INTERVAL) > DateAdd("h",1,Now) Then %>
												<input id="PKVALUE_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" type="hidden" value="PKPHONE;<%=CURRENT_NOTES%>"/>
												<input id="PKPICK_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" class="<%=USE_SCHEDULE%> PKPHONE PKCHECK" type="checkbox" checked="checked" value="PKPHONE"/>
											<% Elseif PICK_DATE(i) = 1 and PICK_FLAG = 1 and LUNCH_FLAG = 1 and SCHEDULE_TYPE = "LUNCH" and CDate(SCHEDULE_START_DATE + i & " " & CURRENT_INTERVAL) > DateAdd("h",1,Now) Then %>
												<input id="PKVALUE_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" type="hidden" value="PKLUNCH;<%=CURRENT_NOTES%>"/>
												<input id="PKPICK_<%=i%>_<%=Replace(CURRENT_INTERVAL,":","")%>_" class="<%=USE_SCHEDULE%> PKLUNCH PKCHECK" type="checkbox" checked="checked" value="PKLUNCH"/>
											<% End If %>
										<!---->
									</td>
								<% Next %>
							</tr>
							<% PREVIOUS_INTERVAL = CURRENT_INTERVAL %>
							<% RSINT.MoveNext %>
						<% Loop %>
					</tbody>
					<tfoot>
						<tr>
							<td style="text-align:center;font-size:10pt;font-weight:900;font-family:Calibri;width:100%;" colspan="8">
								<div id="ERROR_SECTION" style="display:none;color:#c90c35;"></div>
								<div id="LUNCH_WAIVER_AGREEMENT" style="display:none;">
									A minimum 30-minute lunch break is required on all schedules over 5 hours in length. Click "I Agree" below to waive your lunch on <span id="WAIVE_LUNCH_DATES"></span>. In doing so, you are indicating that you are aware of this requirement and are voluntarily waiving a lunch break.
									<div style="margin-bottom:10px;">
										<input type="checkbox" id="LUNCH_WAIVER_AGREEMENT_CHBX" value="N"/> I Agree
									</div>
								</div>
								<div id="SUBMIT_SECTION">
									<input style="font-size:10pt;font-weight:900;font-family:Calibri;display:inline;" type="button" id="SCHEDULE_SUBMIT" value="Submit"/>
								</div>
							</td>
						</tr>
					</tfoot>
				</table>
			</form>
		<% End If %>
		<% Set RSINT = Nothing %>
	<% End If %>
	<% Set RSDAY = Nothing %>

</div>
<script>
	Date.prototype.addDays = function(days) {
		var dat = new Date(this.valueOf());
		dat.setDate(dat.getDate() + days);
		return dat;
	}
    $(document).ready(function() {
		if($("#SCHEDULE_SELECT").val() !== undefined && $("#SCHEDULE_SELECT").val() != ""){
			var curTime = new Date();
			var PAGE_TIME = encodeURIComponent(("0" + (curTime.getMonth()+1)).slice(-2) + "/" + ("0" + curTime.getDate()).slice(-2) + "/" + curTime.getFullYear() + " " + ("0" + curTime.getHours()).slice(-2) + ":" + ("0" + curTime.getMinutes()).slice(-2) + ":" + ("0" + curTime.getSeconds()).slice(-2));
			var START_DATE = new Date($("#FILTER_DATE").val());
			var USE_PATH = encodeURIComponent("<%=Request.ServerVariables("SCRIPT_NAME")%>");

			if ($("input[name=USE_LAYOUT]").val() == "V"){
				$("#SCHEDSQUATCH").stickyTableHeaders({cacheHeaderHeight: true});
			}
			else{
				flipTable($("#SCHEDSQUATCH"));
			}
			$("#SCHEDSQUATCH").show();
			$("#SCHEDSQUATCH_TABS").show();
			$("#SCHEDSQUATCH_TIMER").show();
			$("#SCHEDSQUATCH_KEY_CONTAINER").show();
			$("#SCHEDSQUATCH_KEY").find("tr." + $("#SCHEDULE_SELECT").val()).show();
			$("#SCHEDSQUATCH_STATS").find("tr." + $("#SCHEDULE_SELECT").val()).show();
			if($("#SCHEDULE_SELECT").val() == "AD"){
				$("#LUNCH_FLEX_LINK").show();
			}

			$("#UNP_REMAINING_STAT").html($("#UNP_REMAINING").val());
			$("#ADD_REMAINING_STAT").html($("#ADD_REMAINING").val());
			$("#BASE_HOURS_STAT").html($("#BASE_HOURS").val());
			$("#PICK_HOURS_STAT").html($("#PICK_HOURS").val());
			$("#TOTAL_SCHEDULED_STAT").html($("#TOTAL_SCHEDULED").val());
			$("#REMAINING_NEED_STAT").html($("#REMAINING_NEED").val());
			if ($("#HOLIDAY_HOURS").val() != 0){
				$("#HOLIDAY_STAT").html($("#HOLIDAY_HOURS").val());
				if ($("#HOLIDAY_DEDUCTION").val() != 0){
					$("#REMAINING_POSSIBLE_STAT").html($("#REMAINING_POSSIBLE").val());
				}
				else{
					$("#REMAINING_POSSIBLE_DESCRIPTION").empty();
				}
			}
			else{
				$("#PICK_HOLIDAY_ROW").remove();
			}

			$("#SCHEDSQUATCH tbody").find(":input:not([id^="+ $("#SCHEDULE_SELECT").val() + "])[type!=hidden]").prop("disabled",true);
			if ($("#PK_TAB").length){
				$("#SCHEDSQUATCH tbody").find("input:checkbox[id^=PKPICK][value!='PKOFF']").parent().addClass("PKCHECK");
				$("#SCHEDSQUATCH tbody").find("input:checkbox[id^=PKPICK][value!='PKOFF']").siblings(":input[type!=hidden]").addClass("PKCHECK");
				$("#SCHEDSQUATCH tbody").find("input:checkbox[id^=PKPICK][value='PKPHONE']").siblings("[id^=SCHEDDATA]").val("PKPICK");
				$("#SCHEDSQUATCH tbody").find("input:checkbox[id^=PKPICK][value='PKLUNCH']").siblings("[id^=SCHEDDATA]").val("PKLNCH");
				$("#SCHEDSQUATCH tbody").find("input:checkbox[id^=PKPICK][value='PKOFF']").siblings("[id^=SCHEDDATA]").val("PKOFF");
			}

			if($("#STPLUS_SWITCH").prop("checked") == true){
				$("#ST_TAB").prop("id","SP_TAB");
			}
			errorChecks();
			lunchWaiverAction();

			var timerEnd = new Date();
			timerEnd.setSeconds(timerEnd.getSeconds() - (-1)*$("#USE_TIMER").val());
			var myTimer = setInterval(function(){updateTimer();}, 1000);
		}

		$("#SCHEDSQUATCH_KEY_VIEW").on("click",function() {
			if($("#SCHEDSQUATCH_KEY").is(":visible") == true){
				$("#SCHEDSQUATCH_KEY").hide();
				$("#SCHEDSQUATCH_KEY_VIEW").removeClass("fa-minus-square").addClass("fa-plus-square");
			}
			else{
				$("#SCHEDSQUATCH_KEY").show();
				$("#SCHEDSQUATCH_KEY_VIEW").removeClass("fa-plus-square").addClass("fa-minus-square");
			}
		});

		$("#FILTER_SUBMIT").on("click",function() {
			$("#SCHEDULE_SUBMIT").remove();
			$("#ERROR_SECTION").hide();
			$("#LUNCH_WAIVER_AGREEMENT").hide();

			$("#SCHEDSQUATCH tbody").find("input[type!=hidden], select[type!=hidden]").prop("disabled",true);
			$("#STPLUS_SWITCH").prop("disabled",true);
			$("#AD_TAB, #ST_TAB, #SP_TAB, #PK_TAB").off();
			$("#SCHEDSQUATCH").off("click","#SCHEDSQUATCH_LAYOUT");
			$("#SCHEDSQUATCH").off("click","td[id^=SCHEDCLEAR],th[id^=SCHEDCLEAR]");

			$("#FILTER_SUBMIT").val("Processing...");
			$("#FILTER_FORM").submit();
		});
		$("#STPLUS_SWITCH").on("change",function() {
			if($("#STPLUS_SWITCH").prop("checked") == true){
				$("input[name=STPLUS_ENABLED]").val("Y");
				var OLD_VALUE = "ST";
				var NEW_VALUE = "SP";
			}
			else{
				$("input[name=STPLUS_ENABLED]").val("N");
				var OLD_VALUE = "SP";
				var NEW_VALUE = "ST";
			}
			$("#" + OLD_VALUE + "_TAB").prop("id",NEW_VALUE + "_TAB");
			$("#SCHEDSQUATCH tbody").find("input:checkbox[id^=" + OLD_VALUE  + "]." + OLD_VALUE + "CHECK").each(function(){
				if(this.value != "SPPLUSPHONE"){
					var $that = $(this);
					var $sibling = $that.siblings("input:checkbox[id^=" + NEW_VALUE + "]");
					var value_array = $that.siblings("input[id^=" + OLD_VALUE + "VALUE]").val().split(";");
					var reset_value = value_array[0];
					var cur_value = "";
					if(this.value == "STADD" && $sibling.prop("id").indexOf("PLUS") > -1){
						cur_value = "SPPLUSADD";
					}
					else if(this.value.indexOf("PLUS") > -1){
						cur_value = "ST" + this.value.substr(6);
					}
					else{
						cur_value = NEW_VALUE + this.value.substr(2);
					}
					$that.parent().removeClass(OLD_VALUE + "CHECK " + this.value + " " + $sibling.val()).addClass(NEW_VALUE + "CHECK " + reset_value + " " + cur_value);
					$sibling.prop("checked",true).val(cur_value);
					$that.siblings(":input[type!=hidden]").removeClass(OLD_VALUE + "CHECK").addClass(NEW_VALUE + "CHECK");
					$that.removeClass(OLD_VALUE + "CHECK").addClass(NEW_VALUE + "CHECK").prop("checked",false).val(reset_value);
				}

			});
			$("#SCHEDULE_SELECT").val(NEW_VALUE);
			$("#SCHEDSQUATCH_KEY").find("tr." + OLD_VALUE).hide();
			$("#SCHEDSQUATCH_KEY").find("tr." + NEW_VALUE).show();
			$("#SCHEDSQUATCH tbody").find("." + OLD_VALUE).removeClass(OLD_VALUE).addClass(NEW_VALUE);
			$("#SCHEDSQUATCH thead").find("th[id^=SCHEDCLEAR]").removeClass(OLD_VALUE).addClass(NEW_VALUE);
			$("#SCHEDSQUATCH tbody").find(":input:not([id^="+ NEW_VALUE + "])[type!=hidden]").prop("disabled",true);
			$("#SCHEDSQUATCH tbody").find(":input[id^="+ NEW_VALUE + "][type!=hidden]").prop("disabled",false);
		});

		$("#AD_TAB, #ST_TAB, #SP_TAB, #PK_TAB").on("click",function() {
			var SCHEDULE_SELECT = $("#SCHEDULE_SELECT");
			var OLD_VALUE = SCHEDULE_SELECT.val();
			var NEW_VALUE = this.id.substr(0,2);

			if(OLD_VALUE != NEW_VALUE){
				$("#AD_TAB, #ST_TAB, #SP_TAB, #PK_TAB").removeClass("tabselected").addClass("tab");
				$(this).removeClass("tab").addClass("tabselected");

				$("#SCHEDSQUATCH_KEY").find("tr." + OLD_VALUE).hide();
				$("#SCHEDSQUATCH_KEY").find("tr." + NEW_VALUE).show();
				$("#SCHEDSQUATCH_STATS").find("tr." + OLD_VALUE).hide();
				$("#SCHEDSQUATCH_STATS").find("tr." + NEW_VALUE).show();

				$("#SCHEDSQUATCH tbody").find("." + OLD_VALUE).removeClass(OLD_VALUE).addClass(NEW_VALUE);
				$("#SCHEDSQUATCH thead").find("th[id^=SCHEDCLEAR]").removeClass(OLD_VALUE).addClass(NEW_VALUE);
				$("#SCHEDSQUATCH tbody").find(":input:not([id^="+ NEW_VALUE + "])[type!=hidden]").prop("disabled",true);
				$("#SCHEDSQUATCH tbody").find(":input[id^="+ NEW_VALUE + "][type!=hidden]").prop("disabled",false);
				SCHEDULE_SELECT.val(NEW_VALUE);
			}

		});
		$("#SCHEDSQUATCH").on("click","#SCHEDSQUATCH_LAYOUT",function() {
			if ($("input[name=USE_LAYOUT]").val() == "V"){
				$("input[name=USE_LAYOUT]").val("H");
				$("#SCHEDSQUATCH").stickyTableHeaders("destroy");
				$("#SCHEDSQUATCH_LAYOUT").removeClass("fa-arrows-h").addClass("fa-arrows-v");
				flipTable($("#SCHEDSQUATCH"));
			}
			else{
				$("input[name=USE_LAYOUT]").val("V");
				$("#SCHEDSQUATCH_LAYOUT").removeClass("fa-arrows-v").addClass("fa-arrows-h");
				flipTable($("#SCHEDSQUATCH"));
				$("#SCHEDSQUATCH").stickyTableHeaders({cacheHeaderHeight: true});
			}
		});
		$("#SCHEDSQUATCH").on("change","input:checkbox[id^=ADADD]",function() {
			var $that = $(this);
			switch($that.val()) {
				case "ADADD":
					$that.parent().removeClass($that.val()).addClass("ADPHONE ADCHECK").prop("title","Add/Drop - Added Phone Time");
					$that.siblings(":input[type!=hidden]").addClass("ADCHECK");
					$that.addClass("ADCHECK").val("ADPHONE");
					$that.siblings("[id^=SCHEDDATA]").val("ADADDT");
					addDropCheck(this.id.substr(this.id.indexOf("_")+1,1),"OFF","ADDT",1);
					lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),-.5,0,1);
					break;
				<% If Date <= CDate("12/2/2018") Then %>
					case "ADPHONE":
						$that.parent().removeClass($that.val()).addClass("ADLUNCH").prop("title","Add/Drop - Added Lunch Time");
						$that.prop("checked", true).val("ADLUNCH");
						$that.siblings("[id^=SCHEDDATA]").val("ADLNCH");
						addDropCheck(this.id.substr(this.id.indexOf("_")+1,1),"ADDT","OFF",1);
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),.5,-.5,1);
						break;
					case "ADLUNCH":
						$that.parent().removeClass("ADCHECK " + $that.val()).addClass("ADADD").prop("title","Off");
						$that.siblings(":input[type!=hidden]").removeClass("ADCHECK");
						$that.removeClass("ADCHECK").val("ADADD");
						$that.siblings("[id^=SCHEDDATA]").val("");
						addDropCheck(this.id.substr(this.id.indexOf("_")+1,1),"OFF","ADOFF",1);
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),0,.5,1);
						break;
				<% Else %>
					case "ADPHONE":
						$that.parent().removeClass("ADCHECK " + $that.val()).addClass("ADADD").prop("title","Off");
						$that.siblings(":input[type!=hidden]").removeClass("ADCHECK");
						$that.removeClass("ADCHECK").val("ADADD");
						$that.siblings("[id^=SCHEDDATA]").val("");
						addDropCheck(this.id.substr(this.id.indexOf("_")+1,1),"ADDT","ADOFF",1);
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),.5,0,1);
						break;
				<% End If %>
			}
		});

		$("#SCHEDSQUATCH").on("change","input:checkbox[id^=ADOFF]",function() {
			var $that = $(this);
			switch($that.val()) {
				case "ADOFF":
					$that.parent().removeClass($that.val()).addClass("ADLUNCH ADCHECK").prop("title","Add/Drop - Added Lunch Time");
					$that.siblings(":input[type!=hidden]").addClass("ADCHECK");
					$that.addClass("ADCHECK").val("ADLUNCH");
					$that.siblings("[id^=SCHEDDATA]").val("ADLNCH");
					lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),0,-.5,1);
					break;
				case "ADLUNCH":
					$that.parent().removeClass("ADCHECK " + $that.val()).addClass("ADOFF").prop("title","Off");
					$that.siblings(":input[type!=hidden]").removeClass("ADCHECK");
					$that.removeClass("ADCHECK").val("ADOFF");
					$that.siblings("[id^=SCHEDDATA]").val("");
					lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),0,.5,1);
					break;
			}
		});

		$("#SCHEDSQUATCH").on("focus","select[id^=ADSRED]",function() {
			first_ad_value = $(this).val();
		});
		$("#SCHEDSQUATCH").on("change","select[id^=ADSRED]",function() {
			var $that = $(this);
			switch($that.val()) {
				case "SRUN":
					$that.parent().removeClass("ADPHONE").addClass("ADCHECK ADSRED").prop("title","Add/Drop - Dropped Time (Unpaid)")
					$that.siblings(":input[type!=hidden]").addClass("ADCHECK");
					$that.removeClass("ADPHONE").addClass("ADCHECK ADSRED");
					$that.siblings("[id^=SCHEDDATA]").val("ADSRUN");
					break;
				case "SRPT":
					$that.parent().removeClass("ADPHONE").addClass("ADCHECK ADSRED").prop("title","Add/Drop - Dropped Time (Paid)");
					$that.siblings(":input[type!=hidden]").addClass("ADCHECK");
					$that.removeClass("ADPHONE").addClass("ADCHECK ADSRED");
					$that.siblings("[id^=SCHEDDATA]").val("ADSRPT");
					break;
				default:
					$that.parent().removeClass("ADCHECK ADSRED").addClass("ADPHONE").prop("title","Phone Time");
					$that.siblings(":input[type!=hidden]").removeClass("ADCHECK");
					$that.removeClass("ADCHECK ADSRED").addClass("ADPHONE");
					$that.siblings("[id^=SCHEDDATA]").val("");
					break;
			}
			addDropCheck(this.id.substr(this.id.indexOf("_")+1,1),first_ad_value,$that.val(),1);
			if(first_ad_value == "ADPHONE"){
				lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),.5,0,1);
			}
			else if($that.val() == "ADPHONE"){
				lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),-.5,0,1);
			}
			else{
				lunchWaiver(0,0,0,1);
			}
			first_ad_value = $that.val();
		});

		$("#SCHEDSQUATCH").on("change","input:checkbox[id^=STADD]",function() {
			var $that = $(this);
			switch($that.val()) {
				case "STADD":
					$that.parent().removeClass($that.val()).addClass("STCHECK STPHONE").prop("title","Self Trade - Added Phone Time");
					$that.siblings(":input[type!=hidden]").addClass("STCHECK");
					$that.addClass("STCHECK").val("STPHONE");
					$that.siblings("[id^=SCHEDDATA]").val("STSTBASE");
					selfTradeCheck("",this.id.substr(this.id.indexOf("_")+1,1),.5,1);
					lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),-.5,0,1);
					break;
				<% If Date <= CDate("12/2/2018") Then %>
					case "STPHONE":
						$that.parent().removeClass($that.val()).addClass("STLUNCH").prop("title","Self Trade - Added Lunch Time");
						$that.prop("checked", true).val("STLUNCH");
						$that.siblings("[id^=SCHEDDATA]").val("STLNCH");
						selfTradeCheck("",this.id.substr(this.id.indexOf("_")+1,1),-.5,1);
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),.5,-.5,1);
						break;
					case "STLUNCH":
						$that.parent().removeClass("STCHECK " + $that.val()).addClass("STADD").prop("title","Off");
						$that.siblings(":input[type!=hidden]").removeClass("STCHECK");
						$that.removeClass("STCHECK").val("STADD");
						$that.siblings("[id^=SCHEDDATA]").val("");
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),0,.5,1);
						break;
				<% Else %>
					case "STPHONE":
						$that.parent().removeClass("STCHECK " + $that.val()).addClass("STADD").prop("title","Off");
						$that.siblings(":input[type!=hidden]").removeClass("STCHECK");
						$that.removeClass("STCHECK").val("STADD");
						$that.siblings("[id^=SCHEDDATA]").val("");
						selfTradeCheck("",this.id.substr(this.id.indexOf("_")+1,1),-.5,1);
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),.5,0,1);
						break;
				<% End If %>
			}
		});

		$("#SCHEDSQUATCH").on("change","input:checkbox[id^=STDROP]",function() {
			var $that = $(this);
			var PLUS_STRING = "";
			if($that.siblings("input:checkbox[id^=SPPLUSDROP]").length){
				PLUS_STRING = "PLUS";
			}
			switch($that.val()) {
				case "STPHONE":
					$that.parent().removeClass($that.val()).addClass("STCHECK STADD").prop("title","Self Trade - Dropped Time");
					$that.siblings(":input[type!=hidden]").addClass("STCHECK");
					$that.addClass("STCHECK").val("STADD");
					$that.siblings("[id^=SCHEDDATA]").val("STSTOFF");
					selfTradeCheck(PLUS_STRING,this.id.substr(this.id.indexOf("_")+1,1),-.5,1);
					lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),.5,0,1);
					break;
				<% If Date <= CDate("12/2/2018") Then %>
					case "STADD":
						$that.parent().removeClass($that.val()).addClass("STLUNCH").prop("title","Self Trade - Added Lunch Time");
						$that.prop("checked", true).val("STLUNCH");
						$that.siblings("[id^=SCHEDDATA]").val("STSTLNCH");
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),0,-.5,1);
						break;
					case "STLUNCH":
						$that.parent().removeClass("STCHECK " + $that.val()).addClass("STPHONE").prop("title","Phone Time");
						$that.siblings(":input[type!=hidden]").removeClass("STCHECK");
						$that.removeClass("STCHECK").val("STPHONE");
						$that.siblings("[id^=SCHEDDATA]").val("");
						selfTradeCheck(PLUS_STRING,this.id.substr(this.id.indexOf("_")+1,1),.5,1);
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),-.5,.5,1);
						break;
				<% Else %>
					case "STADD":
						$that.parent().removeClass("STCHECK " + $that.val()).addClass("STPHONE").prop("title","Phone Time");
						$that.siblings(":input[type!=hidden]").removeClass("STCHECK");
						$that.removeClass("STCHECK").val("STPHONE");
						$that.siblings("[id^=SCHEDDATA]").val("");
						selfTradeCheck(PLUS_STRING,this.id.substr(this.id.indexOf("_")+1,1),.5,1);
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),-.5,0,1);
						break;
				<% End If %>
			}
		});

		$("#SCHEDSQUATCH").on("change","input:checkbox[id^=STOFF]",function() {
			var $that = $(this);
			switch($that.val()) {
				case "STOFF":
					$that.parent().removeClass($that.val()).addClass("STCHECK STLUNCH").prop("title","Self Trade - Added Lunch Time");
					$that.siblings(":input[type!=hidden]").addClass("STCHECK");
					$that.addClass("STCHECK").val("STLUNCH");
					$that.siblings("[id^=SCHEDDATA]").val("STLNCH");
					lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),0,-.5,1);
					break;
				case "STLUNCH":
					$that.parent().removeClass("STCHECK " + $that.val()).addClass("STOFF").prop("title","Off");
					$that.siblings(":input[type!=hidden]").removeClass("STCHECK");
					$that.siblings("[id^=SCHEDDATA]").val("");
					$that.removeClass("STCHECK").val("STOFF");
					lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),0,.5,1);
					break;
			}
		});

		$("#SCHEDSQUATCH").on("change","input:checkbox[id^=SPPLUSADD]",function() {
			var $that = $(this);
			switch($that.val()) {
				case "SPPLUSADD":
					$that.parent().removeClass($that.val()).addClass("SPCHECK SPPLUSPHONE").prop("title","Self Trade - Added Phone Time");
					$that.siblings(":input[type!=hidden]").addClass("SPCHECK");
					$that.addClass("SPCHECK").val("SPPLUSPHONE");
					$that.siblings("[id^=SCHEDDATA]").val("STSTBASE");
					selfTradeCheck("PLUS",this.id.substr(this.id.indexOf("_")+1,1),.5,1);
					lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),-.5,0,1);
					break;
				<% If Date <= CDate("12/2/2018") Then %>
					case "SPPLUSPHONE":
						$that.parent().removeClass($that.val()).addClass("SPLUNCH").prop("title","Self Trade - Added Lunch Time");
						$that.prop("checked", true).val("SPLUNCH");
						$that.siblings("[id^=SCHEDDATA]").val("STLNCH");
						selfTradeCheck("PLUS",this.id.substr(this.id.indexOf("_")+1,1),-.5,1);
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),.5,-.5,1);
						break;
					case "SPLUNCH":
						$that.parent().removeClass("SPCHECK " + $that.val()).addClass("SPPLUSADD").prop("title","Off");
						$that.siblings(":input[type!=hidden]").removeClass("SPCHECK");
						$that.removeClass("SPCHECK").val("SPPLUSADD");
						$that.siblings("[id^=SCHEDDATA]").val("");
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),0,.5,1);
						break;
				<% Else %>
					case "SPPLUSPHONE":
						$that.parent().removeClass("SPCHECK " + $that.val()).addClass("SPPLUSADD").prop("title","Off");
						$that.siblings(":input[type!=hidden]").removeClass("SPCHECK");
						$that.removeClass("SPCHECK").val("SPPLUSADD");
						$that.siblings("[id^=SCHEDDATA]").val("");
						selfTradeCheck("PLUS",this.id.substr(this.id.indexOf("_")+1,1),-.5,1);
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),.5,0,1);
						break;
				<% End If %>
			}
		});

		$("#SCHEDSQUATCH").on("change","input:checkbox[id^=SPPLUSDROP]",function() {
			var $that = $(this);
			switch($that.val()) {
				case "SPPLUSPHONE":
					$that.parent().removeClass($that.val()).addClass("SPCHECK SPPLUSADD").prop("title","Self Trade - Dropped Time");
					$that.siblings(":input[type!=hidden]").addClass("SPCHECK");
					$that.addClass("SPCHECK").val("SPPLUSADD");
					$that.siblings("[id^=SCHEDDATA]").val("STSTOFF");
					selfTradeCheck("PLUS",this.id.substr(this.id.indexOf("_")+1,1),-.5,1);
					lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),.5,0,1);
					break;
				<% If Date <= CDate("12/2/2018") Then %>
					case "SPPLUSADD":
						$that.parent().removeClass($that.val()).addClass("SPLUNCH").prop("title","Self Trade - Added Lunch Time");
						$that.prop("checked", true).val("SPLUNCH");
						$that.siblings("[id^=SCHEDDATA]").val("STSTLNCH");
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),0,-.5,1);
						break;
					case "SPLUNCH":
						$that.parent().removeClass("SPCHECK " + $that.val()).addClass("SPPLUSPHONE").prop("title","Phone Time");
						$that.siblings(":input[type!=hidden]").removeClass("SPCHECK");
						$that.removeClass("SPCHECK").val("SPPLUSPHONE");
						$that.siblings("[id^=SCHEDDATA]").val("");
						selfTradeCheck("PLUS",this.id.substr(this.id.indexOf("_")+1,1),.5,1);
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),-.5,.5,1);
						break;
				<% Else %>
					case "SPPLUSADD":
						$that.parent().removeClass("SPCHECK " + $that.val()).addClass("SPPLUSPHONE").prop("title","Phone Time");
						$that.siblings(":input[type!=hidden]").removeClass("SPCHECK");
						$that.removeClass("SPCHECK").val("SPPLUSPHONE");
						$that.siblings("[id^=SCHEDDATA]").val("");
						selfTradeCheck("PLUS",this.id.substr(this.id.indexOf("_")+1,1),.5,1);
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),-.5,0,1);
						break;
				<% End If %>
			}
		});

		$("#SCHEDSQUATCH").on("change","input:checkbox[id^=SPADD]",function() {
			var $that = $(this);
			switch($that.val()) {
				case "SPADD":
					$that.parent().removeClass($that.val()).addClass("SPCHECK SPPHONE").prop("title","Self Trade - Added Phone Time");
					$that.siblings(":input[type!=hidden]").addClass("SPCHECK");
					$that.addClass("SPCHECK").val("SPPHONE");
					$that.siblings("[id^=SCHEDDATA]").val("STSTBASE");
					selfTradeCheck("",this.id.substr(this.id.indexOf("_")+1,1),.5,1);
					lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),-.5,0,1);
					break;
				<% If Date <= CDate("12/2/2018") Then %>
					case "SPPHONE":
						$that.parent().removeClass($that.val()).addClass("SPLUNCH").prop("title","Self Trade - Added Lunch Time");
						$that.prop("checked", true).val("SPLUNCH");
						$that.siblings("[id^=SCHEDDATA]").val("STLNCH");
						selfTradeCheck("",this.id.substr(this.id.indexOf("_")+1,1),-.5,1);
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),.5,-.5,1);
						break;
					case "SPLUNCH":
						$that.parent().removeClass("SPCHECK " + $that.val()).addClass("SPADD").prop("title","Off");
						$that.siblings(":input[type!=hidden]").removeClass("SPCHECK");
						$that.removeClass("SPCHECK").val("SPADD");
						$that.siblings("[id^=SCHEDDATA]").val("");
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),0,.5,1);
						break;
				<% Else %>
					case "SPPHONE":
						$that.parent().removeClass("SPCHECK " +$that.val()).addClass("SPADD").prop("title","Off");
						$that.siblings(":input[type!=hidden]").removeClass("SPCHECK");
						$that.removeClass("SPCHECK").val("SPADD");
						$that.siblings("[id^=SCHEDDATA]").val("");
						selfTradeCheck("",this.id.substr(this.id.indexOf("_")+1,1),-.5,1);
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),.5,0,1);
						break;
				<% End If %>
			}
		});

		$("#SCHEDSQUATCH").on("change","input:checkbox[id^=SPDROP]",function() {
			var $that = $(this);
			switch($that.val()) {
				case "SPPHONE":
					$that.parent().removeClass($that.val()).addClass("SPCHECK SPADD").prop("title","Self Trade - Dropped Time");
					$that.siblings(":input[type!=hidden]").addClass("SPCHECK");
					$that.addClass("SPCHECK").val("SPADD");
					$that.siblings("[id^=SCHEDDATA]").val("STSTOFF");
					selfTradeCheck("",this.id.substr(this.id.indexOf("_")+1,1),-.5,1);
					lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),.5,0,1);
					break;
				<% If Date <= CDate("12/2/2018") Then %>
					case "SPADD":
						$that.parent().removeClass($that.val()).addClass("SPLUNCH").prop("title","Self Trade - Added Lunch Time");
						$that.prop("checked", true).val("SPLUNCH");
						$that.siblings("[id^=SCHEDDATA]").val("STSTLNCH");
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),0,-.5,1);
						break;
					case "SPLUNCH":
						$that.parent().removeClass("SPCHECK " + $that.val()).addClass("SPPHONE").prop("title","Phone Time");
						$that.siblings(":input[type!=hidden]").removeClass("SPCHECK");
						$that.removeClass("SPCHECK").val("SPPHONE");
						$that.siblings("[id^=SCHEDDATA]").val("");
						selfTradeCheck("",this.id.substr(this.id.indexOf("_")+1,1),.5,1);
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),-.5,.5,1);
						break;
				<% Else %>
					case "SPADD":
						$that.parent().removeClass("SPCHECK " + $that.val()).addClass("SPPHONE").prop("title","Phone Time");
						$that.siblings(":input[type!=hidden]").removeClass("SPCHECK");
						$that.removeClass("SPCHECK").val("SPPHONE");
						$that.siblings("[id^=SCHEDDATA]").val("");
						selfTradeCheck("",this.id.substr(this.id.indexOf("_")+1,1),.5,1);
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),-.5,0,1);
						break;
				<% End If %>
			}
		});

		$("#SCHEDSQUATCH").on("change","input:checkbox[id^=SPOFF]",function() {
			var $that = $(this);
			switch($that.val()) {
				case "SPOFF":
					$that.parent().removeClass($that.val()).addClass("SPCHECK SPLUNCH").prop("title","Self Trade - Added Lunch Time");
					$that.siblings(":input[type!=hidden]").addClass("SPCHECK");
					$that.addClass("SPCHECK").val("SPLUNCH");
					$that.siblings("[id^=SCHEDDATA]").val("STLNCH");
					lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),0,-.5,1);
					break;
				case "SPLUNCH":
					$that.parent().removeClass("SPCHECK " + $that.val()).addClass("SPOFF").prop("title","Off");
					$that.siblings(":input[type!=hidden]").removeClass("SPCHECK");
					$that.siblings("[id^=SCHEDDATA]").val("");
					$that.removeClass("SPCHECK").val("SPOFF");
					lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),0,.5,1);
					break;
			}
		});

		$("#SCHEDSQUATCH").on("change","input:checkbox[id^=PKPICK]",function() {
			var $that = $(this);
			switch($that.val()) {
				case "PKOFF":
					$that.parent().removeClass("ADOFF STOFF SPOFF " + $that.val()).addClass("PKCHECK ADPHONE STPHONE SPPHONE PKPHONE").prop("title","Pick Hours - Added Phone Time");
					$that.siblings(":input[type!=hidden]").addClass("PKCHECK");
					$that.addClass("PKCHECK").val("PKPHONE");
					$that.siblings("[id^=SCHEDDATA]").val("PKPICK");
					pickCheck(this.id.substr(this.id.indexOf("_")+1,1),.5,1);
					lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),-.5,0,1);
					break;
				case "PKPHONE":
					$that.parent().removeClass("ADPHONE STPHONE SPPHONE " + $that.val()).addClass("ADLUNCH STLUNCH SPLUNCH PKLUNCH").prop("title","Pick Hours - Added Lunch Time");
					$that.prop("checked", true).val("PKLUNCH");
					$that.siblings("[id^=SCHEDDATA]").val("PKLNCH");
					pickCheck(this.id.substr(this.id.indexOf("_")+1,1),-.5,1);
					lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),.5,-.5,1);
					break;
				case "PKLUNCH":
					$that.parent().removeClass("PKCHECK ADLUNCH STLUNCH SPLUNCH " + $that.val()).addClass("ADOFF STOFF SPOFF PKOFF").prop("title","Off");
					$that.siblings(":input[type!=hidden]").removeClass("PKCHECK");
					$that.removeClass("PKCHECK").val("PKOFF");
					$that.siblings("[id^=SCHEDDATA]").val("PKOFF");
					lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),0,.5,1);
					break;
			}
		});

		$("#SCHEDSQUATCH").on("mouseenter","td[id^=SCHEDCLEAR],th[id^=SCHEDCLEAR]",function() {
			var schedarray = this.id.split("_");
			$("#SCHEDSQUATCH tbody").find("td[id*=_"+schedarray[1]+"_]:not(." + $("#SCHEDULE_SELECT").val() + "CLOSED)").css("box-shadow","0px 0px 10px rgba(0, 40, 78, 0.6)");
		});
		$("#SCHEDSQUATCH").on("mouseleave","td[id^=SCHEDCLEAR],th[id^=SCHEDCLEAR]",function() {
			var schedarray = this.id.split("_");
			$("#SCHEDSQUATCH tbody").find("td[id*=_"+schedarray[1]+"_]:not(." + $("#SCHEDULE_SELECT").val() + "CLOSED)").css("box-shadow","none");
		});
		$("#SCHEDSQUATCH").on("click","td[id^=SCHEDCLEAR],th[id^=SCHEDCLEAR]",function() {
			var schedarray = this.id.split("_");
			var current_view = $("#SCHEDULE_SELECT").val();
			$("#SCHEDSQUATCH tbody").find("input[id^=" + current_view + "VALUE][id*=_"+schedarray[1]+"_]").each(function(){
				var $that = $(this);
				var values_array = this.value.split(";");
				var use_value = values_array[0];
				var use_title = values_array[1];
				var class_string;
				if(current_view == "AD"){
					class_string = "ADPHONE ADADD ADOFF ADLUNCH ADSRED".replace(use_value,"");
					if($that.siblings("input:checkbox[id^=ADADD]").val() == "ADPHONE"){
						addDropCheck(this.id.substr(this.id.indexOf("_")+1,1),"ADDT","OFF",0);
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),.5,0,0);
					}
					else if($that.siblings("input:checkbox[id^=ADADD],input:checkbox[id^=ADOFF]").val() == "ADLUNCH"){
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),0,.5,0);
					}
					else if($that.siblings("select").val() == "SRUN" || $that.siblings("select").val() == "SRPT"){
						addDropCheck(this.id.substr(this.id.indexOf("_")+1,1),$that.siblings("select").val(),use_value,0);
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),-.5,0,0);
					}
				}
				else if (current_view == "ST"){
					class_string = "STPHONE STADD STOFF STLUNCH".replace(use_value,"");
					if($that.siblings("input:checkbox[id^=STADD]").val() == "STPHONE"){
						selfTradeCheck("",this.id.substr(this.id.indexOf("_")+1,1),-.5,0);
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),.5,0,0);
					}
					else if($that.siblings("input:checkbox[id^=STDROP]").val() == "STADD"){
						var PLUS_STRING = "";
						if($that.siblings("input:checkbox[id^=SPPLUSDROP]").length){
							PLUS_STRING = "PLUS";
						}
						selfTradeCheck(PLUS_STRING,this.id.substr(this.id.indexOf("_")+1,1),.5,0);
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),-.5,0,0);
					}
					else if($that.siblings("input:checkbox[id^=STDROP]").val() == "STLUNCH"){
						var PLUS_STRING = "";
						if($that.siblings("input:checkbox[id^=SPPLUSDROP]").length){
							PLUS_STRING = "PLUS";
						}
						selfTradeCheck(PLUS_STRING,this.id.substr(this.id.indexOf("_")+1,1),.5,0);
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),-.5,.5,0);
					}
					else if($that.siblings("input:checkbox[id^=STOFF],input:checkbox[id^=STADD]").val() == "STLUNCH"){
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),0,.5,0);
					}
				}
				else if (current_view == "SP"){
					class_string = "SPPHONE SPADD SPPLUSPHONE SPPLUSADD SPOFF SPLUNCH".replace(use_value,"");
					if($that.siblings("input:checkbox[id^=SPPLUSADD]").val() == "SPPLUSPHONE"){
						selfTradeCheck("PLUS",this.id.substr(this.id.indexOf("_")+1,1),-.5,0);
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),.5,0,0);
					}
					else if($that.siblings("input:checkbox[id^=SPPLUSDROP]").val() == "SPPLUSADD"){
						selfTradeCheck("PLUS",this.id.substr(this.id.indexOf("_")+1,1),.5,0);
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),-.5,0,0);
					}
					else if($that.siblings("input:checkbox[id^=SPPLUSDROP]").val() == "SPLUNCH"){
						selfTradeCheck("PLUS",this.id.substr(this.id.indexOf("_")+1,1),.5,0);
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),-.5,.5,0);
					}
					else if($that.siblings("input:checkbox[id^=SPADD]").val() == "SPPHONE"){
						selfTradeCheck("",this.id.substr(this.id.indexOf("_")+1,1),-.5,0);
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),.5,0,0);
					}
					else if($that.siblings("input:checkbox[id^=SPDROP]").val() == "SPADD"){
						selfTradeCheck("",this.id.substr(this.id.indexOf("_")+1,1),.5,0);
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),-.5,0,0);
					}
					else if($that.siblings("input:checkbox[id^=SPDROP]").val() == "SPLUNCH"){
						selfTradeCheck("",this.id.substr(this.id.indexOf("_")+1,1),.5,0);
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),-.5,.5,0);
					}
					else if($that.siblings("input:checkbox[id^=SPOFF],input:checkbox[id^=SPADD]").val() == "SPLUNCH"){
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),0,.5,0);
					}
				}
				else if (current_view == "PK"){
					class_string = "PKPHONE PKLUNCH";
					if($that.siblings("input:checkbox[id^=PKPICK]").val() == "PKPHONE"){
						pickCheck(this.id.substr(this.id.indexOf("_")+1,1),-.5,0);
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),.5,0,0);
					}
					else if($that.siblings("input:checkbox[id^=PKPICK]").val() == "PKLUNCH"){
						lunchWaiver(this.id.substr(this.id.indexOf("_")+1,1),0,.5,0);
					}
				}
				if (current_view == "AD" || current_view == "ST" || current_view == "SP"){
					$that.siblings("[id^=SCHEDDATA]").val("");
					$that.siblings("input[id^=" + current_view + "][type!=hidden]").prop("checked",false).val(use_value);
					$that.siblings("select[id^=" + current_view + "][type!=hidden]").removeClass(class_string).addClass(use_value).removeAttr("selected").val(use_value);
					$that.parent().removeClass(current_view + "CHECK " + class_string).addClass(use_value).prop("title",use_title);
					$that.siblings("input[type!=hidden], select[type!=hidden]").removeClass(current_view + "CHECK");
				}
				else if (current_view == "PK"){
					if($that.siblings("input[id^=PK][type!=hidden]").val() != "PKOFF"){
						$that.siblings("[id^=SCHEDDATA]").val("PKOFF");
						$that.siblings("input[id^=PK][type!=hidden]").prop("checked",false).val("PKOFF");
						$that.parent().removeClass("PKCHECK ADPHONE STPHONE ADLUNCH STLUNCH " + class_string).addClass("PKOFF ADOFF STOFF").prop("title","Off");
						$that.siblings("input[type!=hidden], select[type!=hidden]").removeClass("PKCHECK");
					}
				}
			});
			if(current_view == "AD"){
				addDropCheck(0,"","",1);
			}
			else if (current_view == "ST" || current_view == "SP"){
				selfTradeCheck("",0,0,1);
			}
			else if (current_view == "PK"){
				pickCheck(0,0,1);
			}
			lunchWaiver(0,0,0,1);
		});

		$("#SCHEDSQUATCH").on("change","#LUNCH_WAIVER_AGREEMENT_CHBX",function() {
			if($(this).val() == "N"){
				$(this).val("Y");
				$("#SCHEDULE_SUBMIT").removeClass("LNCHDISABLED");
				if($("#SCHEDULE_SUBMIT").is(":visible") == true){
					$("#SCHEDULE_SUBMIT").prop("disabled",false);
				}
			}
			else{
				$(this).val("N");
				$("#SCHEDULE_SUBMIT").addClass("LNCHDISABLED").prop("disabled",true);
			}
		});

		$("#SCHEDSQUATCH").on("click","#SCHEDULE_SUBMIT",function() {
			if($("#SUBMIT_FLAG").val() == "0"){
				$("#FILTER_SUBMIT").remove();
				$("#SCHEDULE_SUBMIT").remove();
				$("#SUBMIT_SECTION").html("Processing...");
				$("#SCHEDULE_FORM").submit();
			}
			else{
				var LUNCH_WAIVER = $("#LUNCH_WAIVER");

				clearInterval(myTimer);
				$("#FILTER_SUBMIT").remove();
				$("#SCHEDULE_SUBMIT").remove();
				$("#LUNCH_WAIVER_AGREEMENT_CHBX").prop("disabled",true);
				$("#SUBMIT_SECTION").html("Processing...");
				$("#SCHEDSQUATCH tbody").find("[id^=SCHEDDATA][value='PKOFF']").each(function(){
					if($(this).siblings("input[id^=PKVALUE]").val().indexOf("PKOFF") > -1){
						$(this).val("");
					}
				});
				$("#SCHEDSQUATCH tbody").find("input[type!=hidden], select[type!=hidden]").prop("disabled",true);
				$("#STPLUS_SWITCH").prop("disabled",true);
				$("#AD_TAB, #ST_TAB, #SP_TAB, #PK_TAB").off();
				$("#AD_TAB, #ST_TAB, #SP_TAB, #PK_TAB").off();
				$("#SCHEDSQUATCH").off("click","td[id^=SCHEDCLEAR],th[id^=SCHEDCLEAR]");

				for(var i = 0;i <= 6;i++){
					if(!($("#SCHEDCLEAR_" + i).hasClass("ADCLOSED") && $("#SCHEDCLEAR_" + i).hasClass("STCLOSED") && $("#SCHEDCLEAR_" + i).hasClass("PKCLOSED"))){
						if($("#LUNCH_WORKED_" + i).val() > 5 && $("#LUNCH_HOURS_" + i).val() < 0.5){
							LUNCH_WAIVER.val(LUNCH_WAIVER.val() + formatDate(START_DATE.addDays(i)) + "_1;")
						}
						else{
							LUNCH_WAIVER.val(LUNCH_WAIVER.val() + formatDate(START_DATE.addDays(i)) + "_0;")
						}
					}
				}
				LUNCH_WAIVER.val(LUNCH_WAIVER.val().slice(0,-1));

				$.ajax({
					url: "lastsubmit.asp?USR_ID=" + <%=SCHEDULE_USR_ID%> + "&USE_WEEK=" + encodeURIComponent(formatDate(START_DATE)) + "&REFER_PAGE=" + USE_PATH + "&PAGE_TIME=" + PAGE_TIME,
					cache: false,
					success: function(result){
						$("#SUBMIT_FLAG").val(result.trim());
						$("#SCHEDULE_FORM").submit();
					}
				});
			}
		});

		function addDropCheck(dotw, first_value, second_value, error_bool){
			if (first_value == "SRUN"){
				$("#UNP_REMAINING").val(parseFloat(($("#UNP_REMAINING").val() - (-.5)).toFixed(1)));
			}
			else if (first_value == "SRPT"){
				for(var i = dotw;i <= 6;i++){
					$("#PTO_BALANCE_" + i).val(parseFloat(($("#PTO_BALANCE_" + i).val() - (-.5)).toFixed(1)));
				}
			}
			else if (first_value == "ADDT"){
				$("#ADD_REMAINING").val(parseFloat(($("#ADD_REMAINING").val() - (-.5)).toFixed(1)));
				$("#DAILY_THRESHOLD_" + dotw).val(parseFloat(($("#DAILY_THRESHOLD_" + dotw).val() - (-.5)).toFixed(1)));
			}

			if (second_value == "SRUN"){
				$("#UNP_REMAINING").val(parseFloat(($("#UNP_REMAINING").val() - .5).toFixed(1)));
			}
			else if (second_value == "SRPT"){
				for(var i = dotw;i <= 6;i++){
					$("#PTO_BALANCE_" + i).val(parseFloat(($("#PTO_BALANCE_" + i).val() - .5).toFixed(1)));
				}
			}
			else if (second_value == "ADDT"){
				$("#ADD_REMAINING").val(parseFloat(($("#ADD_REMAINING").val() - .5).toFixed(1)));
				$("#DAILY_THRESHOLD_" + dotw).val(parseFloat(($("#DAILY_THRESHOLD_" + dotw).val() - .5).toFixed(1)));
			}

			if(error_bool == 1){
				$("#UNP_REMAINING_STAT").html($("#UNP_REMAINING").val());
				$("#ADD_REMAINING_STAT").html($("#ADD_REMAINING").val());
				errorChecks();
			}
		}
		function selfTradeCheck(st_mode, dotw, increment_value, error_bool){
			$("#SELFTRADE_COUNTER").val($("#SELFTRADE_COUNTER").val() - increment_value);
			$("#DAILY_THRESHOLD_" + dotw).val(($("#DAILY_THRESHOLD_" + dotw).val() - increment_value).toFixed(1));
			if(st_mode == "PLUS"){
				$("#PLUS_COUNTER").val($("#PLUS_COUNTER").val() - increment_value);
			}
			if(error_bool == 1){
				var SLT_STAT = $("#SLT_STAT");
				var SLT_COUNTER = $("#SELFTRADE_COUNTER").val();
				var PLUS_COUNTER = $("#PLUS_COUNTER").val();

				if(SLT_COUNTER < 0){
					SLT_STAT.html("Trade down " + -2*SLT_COUNTER + " interval(s)");
				}
				else if (SLT_COUNTER > 0){
					SLT_STAT.html("Trade up " + 2*SLT_COUNTER + " interval(s)");
				}
				else if (PLUS_COUNTER < 0){
					SLT_STAT.html("You've traded up " + -2*$("#PLUS_COUNTER").val() + " too many Self Trade Plus intervals");
				}
				else{
					SLT_STAT.html("None");
				}

				errorChecks();
			}
		}
		function pickCheck(dotw, increment_value, error_bool){
			$("#DAILY_THRESHOLD_" + dotw).val(($("#DAILY_THRESHOLD_" + dotw).val() - increment_value).toFixed(1));
			$("#TOTAL_SCHEDULED").val(parseFloat(($("#TOTAL_SCHEDULED").val() - (-1*increment_value)).toFixed(1)));
			$("#REMAINING_POSSIBLE").val(parseFloat(($("#REMAINING_POSSIBLE").val() - increment_value).toFixed(1)));

			if($("#TOTAL_SCHEDULED").val() >= $("#BASE_HOURS").val() - (-1*$("#PICK_HOURS").val()) - $("#HOLIDAY_DEDUCTION").val() && $("#TOTAL_SCHEDULED").val() <= $("#TOTAL_EXPECTED").val()){
				$("#REMAINING_NEED").val("0")
			}
			else{
				$("#REMAINING_NEED").val(parseFloat(($("#REMAINING_NEED").val() - increment_value).toFixed(1)));
			}

			if(error_bool == 1){
				$("#TOTAL_SCHEDULED_STAT").html($("#TOTAL_SCHEDULED").val());
				if($("#HOLIDAY_HOURS").val() != 0){
					$("#REMAINING_NEED_STAT").html(Math.max($("#REMAINING_NEED").val(),Math.min($("#REMAINING_POSSIBLE").val(),0)));
					if($("#HOLIDAY_DEDUCTION").val() != 0){
						$("#REMAINING_POSSIBLE_STAT").html($("#REMAINING_POSSIBLE").val());
					}
				}
				else{
					$("#REMAINING_NEED_STAT").html($("#REMAINING_NEED").val());
				}
				errorChecks();
			}
		}
		function errorChecks(){
			var ERROR_SECTION = $("#ERROR_SECTION");
			ERROR_SECTION.hide().empty();

			var SUBMIT_BUTTON = $("#SCHEDULE_SUBMIT");
			SUBMIT_BUTTON.show();
			if(!SUBMIT_BUTTON.hasClass("LNCHDISABLED")){
				SUBMIT_BUTTON.prop("disabled",false);
			}

			if ($("#AD_TAB").length){
				if($("#UNP_REMAINING").val() < 0){
					ERROR_SECTION.show().append("You have exceeded the allowable unpaid drop limit this week.<br/>");
					SUBMIT_BUTTON.prop("disabled",true).hide();
					$("#UNP_REMAINING_DESCRIPTION").css("color","#c90c35");
					$("#UNP_REMAINING_STAT").css("color","#c90c35");
				}
				else{
					$("#UNP_REMAINING_DESCRIPTION").css("color","#000");
					$("#UNP_REMAINING_STAT").css("color","#000");
				}

				for(var i = 0;i <= 6;i++){
					if($("#PTO_BALANCE_" + i).val() < 0){
						SUBMIT_BUTTON.prop("disabled",true).hide();
						ERROR_SECTION.show().append("You have exceeded your allocated PTO amount.<br/>");
						break;
					}
				}

				if($("#ADD_REMAINING").val() < 0){
					ERROR_SECTION.show().append("You have exceeded the allowable overtime limit this week.<br/>");
					SUBMIT_BUTTON.prop("disabled",true).hide();
					$("#ADD_REMAINING_DESCRIPTION").css("color","#c90c35");
					$("#ADD_REMAINING_STAT").css("color","#c90c35");
				}
				else{
					$("#ADD_REMAINING_DESCRIPTION").css("color","#000");
					$("#ADD_REMAINING_STAT").css("color","#000");
				}
			}
			if ($("#ST_TAB").length || $("#SP_TAB").length){
				if($("#SELFTRADE_COUNTER").val() != 0){
					ERROR_SECTION.show().append("Your trade hours do not match.<br/>");
					SUBMIT_BUTTON.prop("disabled",true).hide();
					$("#SLT_DESCRIPTION").css("color","#c90c35");
					$("#SLT_STAT").css("color","#c90c35");
				}
				if($("#PLUS_COUNTER").val() < 0){
					ERROR_SECTION.show().append("Fix Self Trade Plus imbalance.<br/>");
					SUBMIT_BUTTON.prop("disabled",true).hide();
				}
				if($("#SELFTRADE_COUNTER").val() == 0 && $("#PLUS_COUNTER").val() >= 0){
					$("#SLT_DESCRIPTION").css("color","#000");
					$("#SLT_STAT").css("color","#000");
				}
			}
			if ($("#PK_TAB").length){
				if($("#REMAINING_POSSIBLE").val() < 0){
					ERROR_SECTION.show().append("You have exceeded " + $("#TOTAL_EXPECTED").val() + " scheduled hours this week.<br/>");
					SUBMIT_BUTTON.prop("disabled",true).hide();
					$("#TOTAL_SCHEDULED_STAT").css("color","#c90c35");
					$("#TOTAL_SCHEDULED_DESCRIPTION").css("color","#c90c35");
					$("#REMAINING_NEED_DESCRIPTION").css("color","#c90c35");
					$("#REMAINING_POSSIBLE_DESCRIPTION").css("color","#c90c35");
					$("#REMAINING_NEED_STAT").css("color","#c90c35");
					$("#REMAINING_POSSIBLE_STAT").css("color","#c90c35");
				}
				else if($("#REMAINING_NEED").val() > 0){
					ERROR_SECTION.show().append("Please pick up " + $("#REMAINING_NEED").val() + " more hour(s) this week.<br/>");
					SUBMIT_BUTTON.prop("disabled",true).hide();
					$("#TOTAL_SCHEDULED_STAT").css("color","#c90c35");
					$("#TOTAL_SCHEDULED_DESCRIPTION").css("color","#c90c35");
					$("#REMAINING_NEED_DESCRIPTION").css("color","#c90c35");
					$("#REMAINING_NEED_STAT").css("color","#c90c35");
					$("#REMAINING_POSSIBLE_DESCRIPTION").css("color","#000");
					$("#REMAINING_POSSIBLE_STAT").css("color","#000");
				}
				else{
					$("#TOTAL_SCHEDULED_STAT").css("color","#000");
					$("#TOTAL_SCHEDULED_DESCRIPTION").css("color","#000");
					$("#REMAINING_NEED_DESCRIPTION").css("color","#000");
					$("#REMAINING_POSSIBLE_DESCRIPTION").css("color","#000");
					$("#REMAINING_NEED_STAT").css("color","#000");
					$("#REMAINING_POSSIBLE_STAT").css("color","#000");
				}

				var AGENT_TYPE = $("#AGENT_TYPE").val();
				if(AGENT_TYPE == "SLS" || AGENT_TYPE == "SRV" || AGENT_TYPE == "RES" || AGENT_TYPE == "SPT" || AGENT_TYPE == "OSR"){
					for(var i = 0;i <= 6;i++){
						var PRIOR_STATUS = ""
						var CURRENT_STATUS = "";
						var GAP_COUNTER = 0;
						var LUNCH_COUNTER = 0;
						var PHONE_BOOL = 0;
						var LAST_ID = $("td[id^=TD_" + i + "_]").last().prop("id");
						$("td[id^=TD_" + i + "_]").each(function(){
							var $that = $(this);
							if($that.hasClass("PKLUNCH") && PHONE_BOOL == 1){
								LUNCH_COUNTER += 1;
								CURRENT_STATUS = "GAP";
							}
							else if($that.hasClass("PKOFF") && PHONE_BOOL == 1){
								CURRENT_STATUS = "GAP";
							}
							else if(PHONE_BOOL == 0 && ($that.hasClass("PKPHONE") || $that.hasClass("PKTRAIN"))){
								PHONE_BOOL = 1;
							}

							if (CURRENT_STATUS == "GAP" && PRIOR_STATUS != "GAP"){
								GAP_COUNTER += 1;
							}
							if(CURRENT_STATUS == "GAP" && $that.prop("id") == LAST_ID){
								GAP_COUNTER -= 1;
							}

							PRIOR_STATUS = CURRENT_STATUS;
							CURRENT_STATUS = "";
						});
						if($("#DAILY_THRESHOLD_" + i).val() > 7 && GAP_COUNTER > 0){
							SUBMIT_BUTTON.prop("disabled",true).hide();
							ERROR_SECTION.show().append("You need five hours scheduled in order to have a gap on " + formatDate(START_DATE.addDays(i)) + ".<br/>");
						}
						if(GAP_COUNTER > 2){
							SUBMIT_BUTTON.prop("disabled",true).hide();
							ERROR_SECTION.show().append("You have exceeded two gaps on " + formatDate(START_DATE.addDays(i)) + ".<br/>");
						}
						if(LUNCH_COUNTER > 4){
							SUBMIT_BUTTON.prop("disabled",true).hide();
							ERROR_SECTION.show().append("You have exceeded two hours of lunch on " + formatDate(START_DATE.addDays(i)) + ".<br/>");
						}
					}
				}
			}

			for(var i = 0;i <= 6;i++){
				if($("#DAILY_THRESHOLD_" + i).val() < 0){
					SUBMIT_BUTTON.prop("disabled",true).hide();
					ERROR_SECTION.show().append("You have exceeded " + $("#THRESHOLD_HOURS_" + i).val() + " hours on " + formatDate(START_DATE.addDays(i)) + ".<br/>");
				}
			}
			if (SUBMIT_BUTTON.is(":visible") == true){
				$("#SCHEDSQUATCH_LEGEND").css("background-color","#f7f7f7").css("box-shadow","none");
				$("#SCHEDSQUATCH").css("background-color","#f7f7f7").css("box-shadow","none");
			}
			else{
				$("#SCHEDSQUATCH_LEGEND").css("background-color","#f9e0e9").css("box-shadow","0px 0px 8px rgba(200, 0, 0, 0.9)");
				$("#SCHEDSQUATCH").css("background-color","#f9e0e9").css("box-shadow","0px 0px 8px rgba(200, 0, 0, 0.9)");
			}
		}
		function lunchWaiver(dotw,work_increment,lunch_increment, error_bool){
			$("#LUNCH_WORKED_" + dotw).val(($("#LUNCH_WORKED_" + dotw).val() - work_increment).toFixed(1));
			$("#LUNCH_HOURS_" + dotw).val(($("#LUNCH_HOURS_" + dotw).val() - lunch_increment).toFixed(1));

			if(error_bool == 1){
				$("#LUNCH_WAIVER_AGREEMENT_CHBX").prop("checked",false).val("N");
				lunchWaiverAction();
			}
		}
		function lunchWaiverAction(){
			var SUBMIT_BUTTON = $("#SCHEDULE_SUBMIT");
			var WAIVER_DATES = [];

			SUBMIT_BUTTON.removeClass("LNCHDISABLED");
			if (SUBMIT_BUTTON.is(":visible") == true){
				SUBMIT_BUTTON.prop("disabled",false);
			}

			for(var i = 0;i <= 6;i++){
				if($("#LUNCH_WORKED_" + i).val() > 5 && $("#LUNCH_HOURS_" + i).val() < 0.5 && !($("#SCHEDCLEAR_" + i).hasClass("ADCLOSED") && $("#SCHEDCLEAR_" + i).hasClass("STCLOSED") && $("#SCHEDCLEAR_" + i).hasClass("PKCLOSED"))){
					if(!SUBMIT_BUTTON.hasClass("LNCHDISABLED")){
						SUBMIT_BUTTON.addClass("LNCHDISABLED").prop("disabled",true);
					}
					WAIVER_DATES.push(formatDate(START_DATE.addDays(i)));
				}
			}
			if(WAIVER_DATES.length > 0 && SUBMIT_BUTTON.is(":visible") == true){
				$("#LUNCH_WAIVER_AGREEMENT").show();
				if (WAIVER_DATES.length == 1){
					$("#WAIVE_LUNCH_DATES").html(WAIVER_DATES[0]);
				}
				else if (WAIVER_DATES.length == 2){
					$("#WAIVE_LUNCH_DATES").html(WAIVER_DATES.join(" and "));
				}
				else{
					$("#WAIVE_LUNCH_DATES").html(WAIVER_DATES.join(", ").replace(/, (?![\s\S]*, )/, ", and "));
				}
			}
			else{
				$("#LUNCH_WAIVER_AGREEMENT").hide();
				$("#WAIVE_LUNCH_DATES").empty();
			}
		}
		function updateTimer(){
			var timerCurrent = new Date();
			var TIME_REMAINING = Math.ceil((timerEnd - timerCurrent)/1000);
			if (TIME_REMAINING < 0){
				$("#SUBMIT_FLAG").val("0");

				clearInterval(myTimer);
				$("#SCHEDSQUATCH_TIMER").html("Please refresh below.");

				$("#ERROR_SECTION").hide();
				$("#LUNCH_WAIVER_AGREEMENT").hide();
				$("#SUBMIT_SECTION").show();

				$("#FILTER_SUBMIT").remove();
				$("#SCHEDULE_SUBMIT").show().prop("disabled",false).val("Refresh");
				$("#AD_TAB, #ST_TAB, #SP_TAB, #PK_TAB").off();
				$("#SCHEDSQUATCH").off("click","#SCHEDSQUATCH_LAYOUT");
				$("#SCHEDSQUATCH").off("click","td[id^=SCHEDCLEAR],th[id^=SCHEDCLEAR]");
				$("#SCHEDSQUATCH tbody").find("input[type!=hidden], select[type!=hidden]").prop("disabled",true);
				$("#STPLUS_SWITCH").prop("disabled",true);

				$("#SCHEDSQUATCH_TIMER").css("color","#019601").css("font-weight","900");
				$("#SCHEDSQUATCH_LEGEND").css("background-color","#BEDABE").css("box-shadow","0px 0px 12px rgba(1,150,1,0.95)");
				$("#SCHEDSQUATCH").css("background-color","#BEDABE").css("box-shadow","0px 0px 12px rgba(1,150,1,0.95)");
			}
			else {
				var HH = ("0" + Math.floor((TIME_REMAINING/60) % 60)).substr(-2);
				var MM = ("0" + Math.floor(TIME_REMAINING % 60)).substr(-2);
				$("#SCHEDSQUATCH_TIMER").html("Time Left: " + HH + ":" + MM);
			}
		}
		function formatDate(dateIn) {
			var MM = dateIn.getMonth() - (-1);
			var DD = dateIn.getDate();
			var YYYY = dateIn.getFullYear();
			return MM + "/" + DD + "/" + YYYY;
		}
		function flipTable(useTable) {
			var $that = useTable;
			var newrows = [];
			var colcount = 0;

			$that.find("input:checkbox").each(function(){
				$(this).attr("checked", $(this).prop("checked"));
			});
			$that.find("select option").each(function(){
				if($(this).prop("selected") == true){
					$(this).attr("selected", "selected");
				}
				else{
					$(this).removeAttr("selected");
				}
			});

			$that.find("thead tr, tbody tr").each(function(){
				colcount++;
				var i = 0;
				$(this).find("td, th").each(function(){
					if(newrows[i] === undefined){
						newrows[i] = $("<tr></tr>");
					}

					if(i == 0){
						newrows[i].append(this.outerHTML.replace("<td","<th").replace("/td>","/th>"));
					}
					else{
						newrows[i].append(this.outerHTML.replace("<th","<td").replace("/th>","/td>"));
					}
					i++;
				});
			});

			$that.find("thead tr, tbody tr").remove();
			for(var i = 0; i < newrows.length; i++){
				if(i == 0){
					$that.find("thead").append(newrows[i]);
				}
				else{
					$that.find("tbody").append(newrows[i]);
				}
			}
			$that.find("tfoot td").attr("colSpan",colcount);
		}
	});
</script>
<!--#include virtual="footer.asp"-->
