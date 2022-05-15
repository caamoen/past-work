<!--#include file="pulseheader.asp"-->
<%	
	If Request.Querystring("DATE") <> "" then
		PARAMETER_DATE = CDate(Request.Querystring("DATE"))
	Else
		PARAMETER_DATE = Date
	End If
	If Request.Querystring("MODE") <> "" then
		PARAMETER_MODE = Request.Querystring("MODE")
	Else
		PARAMETER_MODE = "STAFFING"
	End If
	DEPT_ARRAY = Split(PULSE_DEPARTMENT,",")
	If PULSE_SECURITY >= 5 Then
		DEPT_COUNT = -1
	Else
		DEPT_COUNT = UBound(DEPT_ARRAY)
	End If
%>
<% If PARAMETER_MODE = "STAFFING" Then %>
	<optgroup label="Associate">
	<%
		ReDim CHOSEN_PARAMETER_ARRAY(DEPT_COUNT+1)
		i = 0
		SQLstmt = "SELECT OPS_USR_ID VALUE, OPS_USR_NAME DESCRIPTION " & _
		"FROM OPS_USER " & _
		"JOIN OPS_USER_DETAIL " & _
		"ON OPS_USR_ID = OPS_USD_OPS_USR_ID " & _
		"AND TO_DATE(?,'MM/DD/YYYY') BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE " & _
		"WHERE OPS_USD_LOCATION IN ('MOT','WFD','WFH') " & _
		"AND OPS_USD_CLASS IN ('RGFT','RGPT','PT<30','LEAVE') " & _
		"AND OPS_USD_TYPE <> 'HRA' " & _
		"AND OPS_USD_PAY_RATE > 0 "
		CHOSEN_PARAMETER_ARRAY(i) = PARAMETER_DATE
		i = i + 1
		If PULSE_SECURITY < 5 Then
			SQLstmt = SQLstmt & "AND OPS_USD_TYPE IN ("
			For n = 0 to UBound(DEPT_ARRAY)
				If n <> UBound(DEPT_ARRAY) Then
					SQLstmt = SQLstmt & "?,"
				Else
					SQLstmt = SQLstmt & "?) "
				End If
				CHOSEN_PARAMETER_ARRAY(i) = DEPT_ARRAY(n)
				i = i + 1
			Next
		End If
		SQLstmt = SQLstmt & "ORDER BY DESCRIPTION"
		cmd.CommandText = SQLstmt
		For n = 0 to UBound(CHOSEN_PARAMETER_ARRAY)
			cmd.Parameters(n).value = CHOSEN_PARAMETER_ARRAY(n)
		Next
		Set RSSELECT = cmd.Execute(SQLstmt)
	%>
	<% Do While Not RSSELECT.EOF %>
		<option value="AGT_<%=RSSELECT("VALUE")%>"><%=RSSELECT("DESCRIPTION")%><% If CInt(RSSELECT("VALUE")) = CInt(PULSE_USR_ID) Then %> (My Schedule)<% End If %></option>
		<% RSSELECT.MoveNext %>
	<% Loop %>
	<option value="AGT_ALL">All Associates</option>
	<% Set RSSELECT = Nothing %>
	</optgroup>
	<optgroup label="Supervisor">
	<%
		ReDim CHOSEN_PARAMETER_ARRAY(DEPT_COUNT+1)
		i = 0
		SQLstmt = "SELECT DISTINCT " & _
		"OPS_USR_ID VALUE, " & _
		"OPS_USR_NAME DESCRIPTION " & _
		"FROM OPS_USER " & _
		"JOIN OPS_USER_DETAIL " & _
		"ON OPS_USR_ID = OPS_USD_SUPERVISOR " & _
		"AND TO_DATE(?,'MM/DD/YYYY') BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE " & _
		"WHERE OPS_USD_LOCATION IN ('MOT','WFD','WFH') " & _
		"AND OPS_USD_CLASS IN ('RGFT','RGPT','PT<30','LEAVE') " & _
		"AND OPS_USD_TYPE <> 'HRA' " & _
		"AND OPS_USD_PAY_RATE > 0 "
		CHOSEN_PARAMETER_ARRAY(i) = PARAMETER_DATE
		i = i + 1
		If PULSE_SECURITY < 5 Then
			SQLstmt = SQLstmt & "AND OPS_USD_TYPE IN ("
			For n = 0 to UBound(DEPT_ARRAY)
				If n <> UBound(DEPT_ARRAY) Then
					SQLstmt = SQLstmt & "?,"
				Else
					SQLstmt = SQLstmt & "?) "
				End If
				CHOSEN_PARAMETER_ARRAY(i) = DEPT_ARRAY(n)
				i = i + 1
			Next
		End If
		SQLstmt = SQLstmt & "ORDER BY DESCRIPTION"
		cmd.CommandText = SQLstmt
		For n = 0 to UBound(CHOSEN_PARAMETER_ARRAY)
			cmd.Parameters(n).value = CHOSEN_PARAMETER_ARRAY(n)
		Next
		Set RSSELECT = cmd.Execute(SQLstmt)
	%>
		<% Do While Not RSSELECT.EOF %>
			<option value="SUP_<%=RSSELECT("VALUE")%>"><%=RSSELECT("DESCRIPTION")%><% If CInt(RSSELECT("VALUE")) = CInt(PULSE_USR_ID) Then %> (My Team)<% End If %></option>
			<% RSSELECT.MoveNext %>
		<% Loop %>
		<% Set RSSELECT = Nothing %>
	</optgroup>
	<optgroup label="Department">
	<%
		ReDim CHOSEN_PARAMETER_ARRAY(DEPT_COUNT+1)
		i = 0
		SQLstmt = "SELECT DISTINCT " & _
		"DECODE(OPS_USD_TEAM,'SPT','SPT','SRV','SPT','OSR','SPT',OPS_USD_TYPE) VALUE, " & _
		"DECODE(DECODE(OPS_USD_TEAM,'SPT','SPT','SRV','SPT','OSR','SPT',OPS_USD_TYPE),'ACC','Accounting','CRT','Customer Relations','DOC','Documents','GRP','Group','OPS','Operations','OSS','Operations Support','POP','Product Operations','RES','Reservations','SPT','Support Desk') || ' (' || DECODE(OPS_USD_TEAM,'SPT','SPT','SRV','SPT','OSR','SPT',OPS_USD_TYPE) || ')' DESCRIPTION " & _
		"FROM OPS_USER " & _
		"JOIN OPS_USER_DETAIL " & _
		"ON OPS_USR_ID = OPS_USD_SUPERVISOR " & _
		"AND TO_DATE(?,'MM/DD/YYYY') BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE " & _
		"WHERE OPS_USD_LOCATION IN ('MOT','WFD','WFH') " & _
		"AND OPS_USD_CLASS IN ('RGFT','RGPT','PT<30','LEAVE') " & _
		"AND OPS_USD_TYPE <> 'HRA' " & _
		"AND OPS_USD_PAY_RATE > 0 "
		CHOSEN_PARAMETER_ARRAY(i) = PARAMETER_DATE
		i = i + 1
		If PULSE_SECURITY < 5 Then
			SQLstmt = SQLstmt & "AND OPS_USD_TYPE IN ("
			For n = 0 to UBound(DEPT_ARRAY)
				If n <> UBound(DEPT_ARRAY) Then
					SQLstmt = SQLstmt & "?,"
				Else
					SQLstmt = SQLstmt & "?) "
				End If
				CHOSEN_PARAMETER_ARRAY(i) = DEPT_ARRAY(n)
				i = i + 1
			Next
		End If
		SQLstmt = SQLstmt & "ORDER BY DESCRIPTION"
		cmd.CommandText = SQLstmt
		For n = 0 to UBound(CHOSEN_PARAMETER_ARRAY)
			cmd.Parameters(n).value = CHOSEN_PARAMETER_ARRAY(n)
		Next
		Set RSSELECT = cmd.Execute(SQLstmt)
	%>
		<% Do While Not RSSELECT.EOF %>
			<option value="DEPT_<%=RSSELECT("VALUE")%>"><%=RSSELECT("DESCRIPTION")%></option>
			<% RSSELECT.MoveNext %>
		<% Loop %>
		<% Set RSSELECT = Nothing %>
	</optgroup>
	<%
		ReDim CHOSEN_PARAMETER_ARRAY(DEPT_COUNT+3)
		i = 0
		SQLstmt = "SELECT * " & _
		"FROM " & _
		"( " & _
			"SELECT DISTINCT " & _
			"CASE " & _
				"WHEN OPS_USD_TYPE = 'RES' OR OPS_USD_TEAM = 'SPT' THEN DECODE(RES_RTE_RES_RTG_ID,4,'SPT',5,'SPT',13,'SPT','RES') || DECODE(RES_RTE_RES_RTG_ID,1,DECODE(OPS_USD_TEAM,'SLS','SLS','RES'),4,'VSS',5,'ASR',10,'SRV',13,'OSR','SRV') " & _
				"ELSE OPS_USD_TYPE || DECODE(OPS_USD_TYPE,'GRP',OPS_USD_JOB,OPS_USD_TEAM) " & _
			"END VALUE, " & _
			"CASE " & _
				"WHEN OPS_USD_TYPE = 'RES' OR OPS_USD_TEAM = 'SPT' THEN DECODE(RES_RTE_RES_RTG_ID,1,DECODE(OPS_USD_TEAM,'SLS','Sales Specialty (SLS)','RES Sales (RES)'),4,'SPT VSS',5,'SPT ASR',10,'Elite Service (SES)',13,'SPT OSR') " & _
				"WHEN OPS_USD_TYPE = 'POP' AND OPS_USD_TEAM = 'POC' THEN 'Product Ops (POC)' " & _
				"WHEN OPS_USD_TYPE = 'POP' AND OPS_USD_TEAM = 'DOC' THEN 'Documents' " & _
				"WHEN OPS_USD_TYPE = 'OSS' AND OPS_USD_TEAM = 'SKD' THEN 'Schedule Changes (SKD)' " & _
				"WHEN OPS_USD_TYPE = 'OSS' AND OPS_USD_TEAM = 'PRD' THEN 'Product Support (PRD)' " & _
				"WHEN OPS_USD_TYPE = 'OSS' AND OPS_USD_TEAM = 'AIR' THEN 'Air Support' " & _
				"WHEN OPS_USD_TYPE = 'OSS' AND OPS_USD_TEAM = 'LED' THEN 'OSS Leads' " & _
				"WHEN OPS_USD_TYPE = 'OPS' THEN 'OPS Desk' " & _
				"WHEN OPS_USD_TYPE = 'GRP' AND OPS_USD_JOB = 'SPC' THEN 'Group Product (GPR)' " & _
				"WHEN OPS_USD_TYPE = 'GRP' AND OPS_USD_JOB = 'GSP' THEN 'Group Service (GSR)' " & _
				"WHEN OPS_USD_TYPE = 'GRP' AND OPS_USD_JOB = 'GSM' THEN 'Group Sales (GSA)' " & _
				"WHEN OPS_USD_TYPE = 'DOC' THEN 'Facilities' " & _
				"WHEN OPS_USD_TYPE = 'CRT' THEN 'Customer Relations' " & _
				"WHEN OPS_USD_TYPE = 'ACC' AND OPS_USD_TEAM = 'REC' THEN 'Account Receivable' " & _
				"WHEN OPS_USD_TYPE = 'ACC' AND OPS_USD_TEAM = 'PAY' THEN 'Account Payable' " & _
				"WHEN OPS_USD_TYPE = 'ACC' AND OPS_USD_TEAM = 'LDA' THEN 'Account Leads' " & _
			"END DESCRIPTION " & _
			"FROM OPS_USER " & _
			"JOIN OPS_USER_DETAIL " & _
			"ON OPS_USD_OPS_USR_ID = OPS_USR_ID " & _
			"AND TO_DATE(?,'MM/DD/YYYY') BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE " & _
			"LEFT JOIN " & _
			"( " & _
				"SELECT " & _
				"RES_RTE_OPS_USR_ID, " & _
				"MAX(DECODE(RES_RTE_RES_RTG_ID,0,NULL,RES_RTE_RES_RTG_ID)) KEEP (DENSE_RANK LAST ORDER BY RTE_ORDERING) RES_RTE_RES_RTG_ID " & _
				"FROM " & _
				"( " & _
					"SELECT " & _
					"RES_RTE_OPS_USR_ID, " & _
					"DECODE(RES_RTE_RES_RTG_ID,2,1,3,1,RES_RTE_RES_RTG_ID) RES_RTE_RES_RTG_ID, " & _
					"1 RTE_ORDERING " & _
					"FROM RES_ROUTING " & _
					"WHERE RES_RTE_YEAR = TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE)-(6/24),'YYYY') " & _
					"AND RES_RTE_MONTH = TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE)-(6/24),'MM') " & _
					"UNION ALL " & _
					"SELECT TO_NUMBER(SYS_CDD_NAME), " & _
					"DECODE(TO_NUMBER(REGEXP_SUBSTR(SYS_CDD_VALUE,'[^;]+',1,2)),2,1,3,1,TO_NUMBER(REGEXP_SUBSTR(SYS_CDD_VALUE,'[^;]+',1,2))), " & _
					"2 " & _
					"FROM SYS_CODE_DETAIL " & _
					"WHERE SYS_CDD_SYS_CDM_ID = 508 " & _
					"AND TO_DATE(?,'MM/DD/YYYY') >= TO_DATE(REGEXP_SUBSTR(SYS_CDD_VALUE,'[^;]+',1,1),'MM/DD/YYYY') " & _
					"AND TO_DATE(REGEXP_SUBSTR(SYS_CDD_VALUE,'[^;]+',1,1),'MM/DD/YYYY') >= TO_DATE(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE)) " & _
					"UNION ALL " & _
					"SELECT " & _
					"RES_STA_OPS_USR_ID, " & _
					"DECODE(RES_STA_RES_RTD_ID,2,1,3,1,RES_STA_RES_RTD_ID) RES_STA_RES_RTD_ID, " & _
					"2 " & _
					"FROM RES_STATS_INCENTIVE " & _
					"WHERE RES_STA_DATE = TO_DATE(?,'MM/DD/YYYY') " & _
				") " & _
				"GROUP BY RES_RTE_OPS_USR_ID " & _
			") " & _
			"ON RES_RTE_OPS_USR_ID = OPS_USD_OPS_USR_ID " & _
			"WHERE OPS_USD_LOCATION IN ('MOT','WFD','WFH') " & _
			"AND OPS_USD_CLASS IN ('RGFT','RGPT','PT<30','LEAVE') " & _
			"AND OPS_USD_PAY_RATE > 0 " & _
			"AND NOT " & _
			"( " & _
				"OPS_USD_TYPE = 'RES' " & _
				"AND RES_RTE_RES_RTG_ID IS NULL " & _
			") "
			CHOSEN_PARAMETER_ARRAY(i) = PARAMETER_DATE
			CHOSEN_PARAMETER_ARRAY(i+1) = PARAMETER_DATE
			CHOSEN_PARAMETER_ARRAY(i+2) = PARAMETER_DATE
			i = 3
			If PULSE_SECURITY < 5 Then
				SQLstmt = SQLstmt & "AND OPS_USD_TYPE IN ("
				For n = 0 to UBound(DEPT_ARRAY)
					If n <> UBound(DEPT_ARRAY) Then
						SQLstmt = SQLstmt & "?,"
					Else
						SQLstmt = SQLstmt & "?) "
					End If
					CHOSEN_PARAMETER_ARRAY(i) = DEPT_ARRAY(n)
					i = i + 1
				Next
			End If
		SQLstmt = SQLstmt & ") " & _
		"ORDER BY DESCRIPTION"
		cmd.CommandText = SQLstmt
		For n = 0 to UBound(CHOSEN_PARAMETER_ARRAY)
			cmd.Parameters(n).value = CHOSEN_PARAMETER_ARRAY(n)
		Next
		Erase CHOSEN_PARAMETER_ARRAY
		Set RSSELECT = cmd.Execute(SQLstmt)
	%>
	<% If Not RSSELECT.EOF Then %>
		<optgroup label="Workgroup">
			<% Do While Not RSSELECT.EOF %>
				<option value="WRK_<%=RSSELECT("VALUE")%>"><%=RSSELECT("DESCRIPTION")%></option>
				<% RSSELECT.MoveNext %>
			<% Loop %>
		</optgroup>
	<% End If %>
	<% Set RSSELECT = Nothing %>
	<%
		ReDim CHOSEN_PARAMETER_ARRAY(DEPT_COUNT+1)
		i = 0
		SQLstmt = "SELECT DISTINCT " & _
		"OPS_USD_CLASS VALUE, " & _
		"DECODE(OPS_USD_CLASS,'RGFT','Reg Full-Time (RGFT)','RGPT','Reg Part-Time (RGPT)','PT<30','PT Less Than 30 (PT<30)','LEAVE','On Leave', OPS_USD_CLASS) DESCRIPTION " & _
		"FROM OPS_USER " & _
		"JOIN OPS_USER_DETAIL " & _
		"ON OPS_USR_ID = OPS_USD_SUPERVISOR " & _
		"AND TO_DATE(?,'MM/DD/YYYY') BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE " & _
		"WHERE OPS_USD_LOCATION IN ('MOT','WFD','WFH') " & _
		"AND OPS_USD_CLASS IN ('RGFT','RGPT','PT<30','LEAVE') " & _
		"AND OPS_USD_TYPE <> 'HRA' " & _
		"AND OPS_USD_PAY_RATE > 0 "
		CHOSEN_PARAMETER_ARRAY(i) = PARAMETER_DATE
		i = i + 1
		If PULSE_SECURITY < 5 Then
			SQLstmt = SQLstmt & "AND OPS_USD_TYPE IN ("
			For n = 0 to UBound(DEPT_ARRAY)
				If n <> UBound(DEPT_ARRAY) Then
					SQLstmt = SQLstmt & "?,"
				Else
					SQLstmt = SQLstmt & "?) "
				End If
				CHOSEN_PARAMETER_ARRAY(i) = DEPT_ARRAY(n)
				i = i + 1
			Next
		End If
		SQLstmt = SQLstmt & "ORDER BY DESCRIPTION"
		cmd.CommandText = SQLstmt
		For n = 0 to UBound(CHOSEN_PARAMETER_ARRAY)
			cmd.Parameters(n).value = CHOSEN_PARAMETER_ARRAY(n)
		Next
		Set RSSELECT = cmd.Execute(SQLstmt)
	%>
	<% If Not RSSELECT.EOF Then %>
		<optgroup label="Classification">
			<% Do While Not RSSELECT.EOF %>
				<option value="CLASS_<%=RSSELECT("VALUE")%>"><%=RSSELECT("DESCRIPTION")%></option>
				<% RSSELECT.MoveNext %>
			<% Loop %>
		</optgroup>
	<% End If %>
	<% Set RSSELECT = Nothing %>
	<%
		ReDim CHOSEN_PARAMETER_ARRAY(DEPT_COUNT+1)
		i = 0
		SQLstmt = "SELECT DISTINCT " & _
		"OPS_USD_LOCATION VALUE, " & _
		"DECODE(OPS_USD_LOCATION,'MOT','In-House (MOT)','WFD','Work From Distance (WFD)','WFH','Work From Home (WFH)',OPS_USD_LOCATION) DESCRIPTION " & _
		"FROM OPS_USER " & _
		"JOIN OPS_USER_DETAIL " & _
		"ON OPS_USR_ID = OPS_USD_SUPERVISOR " & _
		"AND TO_DATE(?,'MM/DD/YYYY') BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE " & _
		"WHERE OPS_USD_LOCATION IN ('MOT','WFD','WFH') " & _
		"AND OPS_USD_CLASS IN ('RGFT','RGPT','PT<30','LEAVE') " & _
		"AND OPS_USD_TYPE <> 'HRA' " & _
		"AND OPS_USD_PAY_RATE > 0 "
		CHOSEN_PARAMETER_ARRAY(i) = PARAMETER_DATE
		i = i + 1
		If PULSE_SECURITY < 5 Then
			SQLstmt = SQLstmt & "AND OPS_USD_TYPE IN ("
			For n = 0 to UBound(DEPT_ARRAY)
				If n <> UBound(DEPT_ARRAY) Then
					SQLstmt = SQLstmt & "?,"
				Else
					SQLstmt = SQLstmt & "?) "
				End If
				CHOSEN_PARAMETER_ARRAY(i) = DEPT_ARRAY(n)
				i = i + 1
			Next
		End If
		SQLstmt = SQLstmt & "ORDER BY DESCRIPTION"
		cmd.CommandText = SQLstmt
		For n = 0 to UBound(CHOSEN_PARAMETER_ARRAY)
			cmd.Parameters(n).value = CHOSEN_PARAMETER_ARRAY(n)
		Next
		Set RSSELECT = cmd.Execute(SQLstmt)
	%>
	<% If Not RSSELECT.EOF Then %>
		<optgroup label="Location">
			<% Do While Not RSSELECT.EOF %>
				<option value="LOC_<%=RSSELECT("VALUE")%>"><%=RSSELECT("DESCRIPTION")%></option>
				<% RSSELECT.MoveNext %>
			<% Loop %>
		</optgroup>
	<% End If %>
	<% Set RSSELECT = Nothing %>
	<%
		ReDim CHOSEN_PARAMETER_ARRAY(DEPT_COUNT+1)
		i = 0
		SQLstmt = "SELECT DISTINCT " & _
		"DECODE(OPS_USD_TYPE,'GRP','AGT',DECODE(OPS_USD_TEAM,'LDA','LED',OPS_USD_JOB)) VALUE, " & _
		"DECODE(DECODE(OPS_USD_TYPE,'GRP','AGT',DECODE(OPS_USD_TEAM,'LDA','LED',OPS_USD_JOB)),'AGT','Associate (AGT)','LED','Lead (LED)','ANL','Analyst (ANL)',DECODE(OPS_USD_TYPE,'GRP','AGT',DECODE(OPS_USD_TEAM,'LDA','LED',OPS_USD_JOB))) DESCRIPTION " & _
		"FROM OPS_USER " & _
		"JOIN OPS_USER_DETAIL " & _
		"ON OPS_USR_ID = OPS_USD_SUPERVISOR " & _
		"AND TO_DATE(?,'MM/DD/YYYY') BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE " & _
		"WHERE OPS_USD_LOCATION IN ('MOT','WFD','WFH') " & _
		"AND OPS_USD_CLASS IN ('RGFT','RGPT','PT<30','LEAVE') " & _
		"AND OPS_USD_TYPE <> 'HRA' " & _
		"AND OPS_USD_PAY_RATE > 0 "
		CHOSEN_PARAMETER_ARRAY(i) = PARAMETER_DATE
		i = i + 1
		If PULSE_SECURITY < 5 Then
			SQLstmt = SQLstmt & "AND OPS_USD_TYPE IN ("
			For n = 0 to UBound(DEPT_ARRAY)
				If n <> UBound(DEPT_ARRAY) Then
					SQLstmt = SQLstmt & "?,"
				Else
					SQLstmt = SQLstmt & "?) "
				End If
				CHOSEN_PARAMETER_ARRAY(i) = DEPT_ARRAY(n)
				i = i + 1
			Next
		End If
		SQLstmt = SQLstmt & "ORDER BY DESCRIPTION"
		cmd.CommandText = SQLstmt
		For n = 0 to UBound(CHOSEN_PARAMETER_ARRAY)
			cmd.Parameters(n).value = CHOSEN_PARAMETER_ARRAY(n)
		Next
		Set RSSELECT = cmd.Execute(SQLstmt)
	%>
	<% If Not RSSELECT.EOF Then %>
		<optgroup label="Job">
			<% Do While Not RSSELECT.EOF %>
				<option value="JOB_<%=RSSELECT("VALUE")%>"><%=RSSELECT("DESCRIPTION")%></option>
				<% RSSELECT.MoveNext %>
			<% Loop %>
		</optgroup>
	<% End If %>
	<% Set RSSELECT = Nothing %>
	<optgroup label="Schedule Code">
	<%
		SQLstmt = "SELECT SYS_CDD_VALUE VALUE, SYS_CDD_NAME || ' (' || SYS_CDD_VALUE || ')' DESCRIPTION FROM SYS_CODE_DETAIL " & _
		"WHERE SYS_CDD_SYS_CDM_ID IN (33,34) " & _
		"AND SYS_CDD_VALUE NOT LIKE '%SK' " & _
		"ORDER BY DESCRIPTION"
		Set RSSELECT = Conn.Execute(SQLstmt)
	%>
	<% Do While Not RSSELECT.EOF %>
		<option value="SCH_<%=RSSELECT("VALUE")%>"><%=RSSELECT("DESCRIPTION")%></option>
		<% RSSELECT.MoveNext %>
	<% Loop %>
	<% Set RSSELECT = Nothing %>
	</optgroup>
	<optgroup label="Times">
	<%
		SQLstmt = "WITH INTERVAL_DATA AS " & _
		"( " & _
			"SELECT " & _
			"TO_DATE(?,'MM/DD/YYYY') USE_DATE, " & _
			"OPS_PAR_VALUE INTERVAL_LENGTH " & _
			"FROM OPS_PARAMETER " & _
			"WHERE OPS_PAR_PARENT_TYPE = 'STF' " & _
			"AND TO_DATE(?,'MM/DD/YYYY') BETWEEN OPS_PAR_EFF_DATE AND OPS_PAR_DIS_DATE " & _
			"AND OPS_PAR_CODE = 'INTERVAL_LENGTH' " & _
		") " & _
		"SELECT " & _
		"'CW' || INTERVAL_LENGTH VALUE, " & _
		"'Currently Working' DESCRIPTION " & _
		"FROM INTERVAL_DATA " & _
		"WHERE USE_DATE BETWEEN TO_DATE(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE) + ((60 * INTERVAL_LENGTH)*ROUND(TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS')/(60 * INTERVAL_LENGTH)) - TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS'))/86400) AND TO_DATE(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE) + (1/24) + ((60 * INTERVAL_LENGTH)*ROUND(TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS')/(60 * INTERVAL_LENGTH)) - TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS'))/86400) " & _
		"UNION ALL " & _
		"SELECT " & _
		"'SNH' || INTERVAL_LENGTH VALUE, " & _
		"'Start: Next Hour' DESCRIPTION " & _
		"FROM INTERVAL_DATA " & _
		"WHERE USE_DATE BETWEEN TO_DATE(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE) + ((60 * INTERVAL_LENGTH)*ROUND(TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS')/(60 * INTERVAL_LENGTH)) - TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS'))/86400) AND TO_DATE(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE) + (1/24) + ((60 * INTERVAL_LENGTH)*ROUND(TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS')/(60 * INTERVAL_LENGTH)) - TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS'))/86400) " & _
		"UNION ALL " & _
		"SELECT " & _
		"'S' ||  TO_CHAR(USE_DATE + (ROWNUM - 1) / 24,'MMDDYYYYHH24MI') || TO_CHAR(USE_DATE + (ROWNUM / 24) - (INTERVAL_LENGTH / 1440),'MMDDYYYYHH24MI'), " & _
		"'Start: ' || TO_CHAR(TO_DATE('00:00','HH24:MI') + (ROWNUM - 1) / 24,'HH24:MI') || ' - ' || TO_CHAR(TO_DATE('00:00','HH24:MI') + (ROWNUM / 24) - (INTERVAL_LENGTH / 1440),'HH24:MI') " & _
		"FROM INTERVAL_DATA " & _
		"CONNECT BY ROWNUM <= 24 " & _
		"UNION ALL " & _
		"SELECT " & _
		"'ENH' || INTERVAL_LENGTH, " & _
		"'End: Next Hour' " & _
		"FROM INTERVAL_DATA " & _
		"WHERE USE_DATE BETWEEN TO_DATE(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE) + ((60 * INTERVAL_LENGTH)*ROUND(TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS')/(60 * INTERVAL_LENGTH)) - TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS'))/86400) AND TO_DATE(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE) + (1/24) + ((60 * INTERVAL_LENGTH)*ROUND(TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS')/(60 * INTERVAL_LENGTH)) - TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS'))/86400) " & _
		"UNION ALL " & _
		"SELECT " & _
		"'E' || TO_CHAR(USE_DATE + (ROWNUM - 1)/24 + (INTERVAL_LENGTH/1440),'MMDDYYYYHH24MI') || TO_CHAR(USE_DATE + (ROWNUM / 24),'MMDDYYYYHH24MI'), " & _
		"'End: ' || TO_CHAR(TO_DATE('00:00','HH24:MI') + (ROWNUM-1)/24 + (INTERVAL_LENGTH/1440),'HH24:MI') || ' - ' || TO_CHAR(TO_DATE('00:00','HH24:MI') + (ROWNUM / 24),'HH24:MI') " & _
		"FROM INTERVAL_DATA " & _
		"CONNECT BY ROWNUM <= 24"
		cmd.CommandText = SQLstmt
		cmd.Parameters(0).value = PARAMETER_DATE
		cmd.Parameters(1).value = PARAMETER_DATE
		Set RSSELECT = cmd.Execute(SQLstmt)
	%>
	<% Do While Not RSSELECT.EOF %>
		<option value="TIMES_<%=RSSELECT("VALUE")%>"><%=RSSELECT("DESCRIPTION")%></option>
		<% RSSELECT.MoveNext %>
	<% Loop %>
	<% Set RSSELECT = Nothing %>
	</optgroup>
	<% If (InStr(PULSE_DEPARTMENT,"RES") <> 0 or PULSE_SECURITY >= 5) and PULSE_SECURITY >= 3 Then %>
		<% 
			SQLstmt = "SELECT DISTINCT OPS_USR_HIRE_DATE VALUE, OPS_USR_HIRE_DATE DESCRIPTION " & _
			"FROM OPS_USER " & _
			"JOIN OPS_USER_DETAIL " & _
			"ON OPS_USR_ID = OPS_USD_OPS_USR_ID " & _
			"AND TO_DATE(?,'MM/DD/YYYY') BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE " & _
			"WHERE OPS_USD_LOCATION IN ('MOT','WFD','WFH') " & _
			"AND OPS_USD_CLASS IN ('RGFT','RGPT','PT<30','LEAVE') " & _
			"AND OPS_USD_TYPE <> 'HRA' " & _
			"AND OPS_USD_PAY_RATE > 0 " & _
			"AND OPS_USR_HIRE_DATE BETWEEN ADD_MONTHS(TO_DATE(?,'MM/DD/YYYY'),-24) AND TO_DATE(?,'MM/DD/YYYY') " & _
			"ORDER BY DESCRIPTION"
			cmd.CommandText = SQLstmt
			cmd.Parameters(0).value = PARAMETER_DATE
			cmd.Parameters(1).value = PARAMETER_DATE
			cmd.Parameters(2).value = PARAMETER_DATE
			Set RSSELECT = cmd.Execute(SQLstmt)
		%>
		<% If Not RSSELECT.EOF Then %>
			<optgroup label="New Hire Class">
			<% Do While Not RSSELECT.EOF %>
				<option value="HIRE_<%=RSSELECT("VALUE")%>"><%=RSSELECT("DESCRIPTION")%></option>
				<% RSSELECT.MoveNext %>
			<% Loop %>
			</optgroup>
		<% End If %>
		<% Set RSSELECT = Nothing %>
		<optgroup label="Training">
			<option value="TRN_3">Caribbean</option>
			<option value="TRN_4">Hawaii</option>
			<option value="TRN_5">International</option>
			<option value="TRN_43">Dual</option>
			<option value="CHAT">Chat</option>
		</optgroup>
		<optgroup label="Routing Group">
			<option value="ROUT_1">Ruby</option>
			<option value="ROUT_2">Emerald</option>
			<option value="ROUT_3">Sapphire</option>
		</optgroup>
	<% End If %>
<% Elseif PARAMETER_MODE = "ADMIN" Then %>
	<optgroup label="Employee">
	<%
		SQLstmt = "SELECT " & _
		"OPS_USR_ID VALUE, OPS_USR_NAME DESCRIPTION, " & _
		"MIN(CASE WHEN TO_DATE(?,'MM/DD/YYYY') BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE THEN 0 ELSE 1 END) FUTURE_BOOL " & _
		"FROM OPS_USER " & _
		"JOIN OPS_USER_DETAIL " & _
		"ON OPS_USR_ID = OPS_USD_OPS_USR_ID " & _
		"AND OPS_USD_DIS_DATE >= TO_DATE(?,'MM/DD/YYYY') "
		If PULSE_SECURITY <= 5 Then
			SQLstmt = SQLstmt & "AND OPS_USD_LOCATION IN ('MOT','WFD','WFH') "
		End If 
		SQLstmt = SQLstmt & "AND OPS_USR_ID NOT IN (2026,9151,9618) " & _
		"GROUP BY OPS_USR_ID, OPS_USR_NAME " & _
		"ORDER BY DESCRIPTION"
		cmd.CommandText = SQLstmt
		cmd.Parameters(0).value = PARAMETER_DATE
		cmd.Parameters(1).value = PARAMETER_DATE
		Set RSSELECT = cmd.Execute(SQLstmt)
	%>
	<% Do While Not RSSELECT.EOF %>
		<option <% If RSSELECT("FUTURE_BOOL") = "1" Then %>class="new-entry-color" <% End If %> value="AGT_<%=RSSELECT("VALUE")%>"><%=RSSELECT("DESCRIPTION")%></option>
		<% RSSELECT.MoveNext %>
	<% Loop %>
	<option value="AGT_ALL">All Employees</option>
	<% Set RSSELECT = Nothing %>
	</optgroup>
	<optgroup label="Supervisor">
	<%
		SQLstmt = "SELECT DISTINCT " & _
		"OPS_USR_ID VALUE, " & _
		"OPS_USR_NAME DESCRIPTION " & _
		"FROM OPS_USER " & _
		"JOIN OPS_USER_DETAIL " & _
		"ON OPS_USR_ID = OPS_USD_SUPERVISOR " & _
		"AND OPS_USD_DIS_DATE >= TO_DATE(?,'MM/DD/YYYY') "
		If PULSE_SECURITY <= 5 Then
			SQLstmt = SQLstmt & "AND OPS_USD_LOCATION IN ('MOT','WFD','WFH') "
		End If
		SQLstmt = SQLstmt & "AND OPS_USR_ID NOT IN (2026,9151,9618) " & _
		"ORDER BY DESCRIPTION"
		cmd.CommandText = SQLstmt
		cmd.Parameters(0).value = PARAMETER_DATE
		Set RSSELECT = cmd.Execute(SQLstmt)
	%>
		<% Do While Not RSSELECT.EOF %>
			<option value="SUP_<%=RSSELECT("VALUE")%>"><%=RSSELECT("DESCRIPTION")%></option>
			<% RSSELECT.MoveNext %>
		<% Loop %>
		<% Set RSSELECT = Nothing %>
	</optgroup>
	<optgroup label="Department">
	<%
		SQLstmt = "SELECT DISTINCT " & _
		"DECODE(OPS_USD_TEAM,'SPT','SPT','SRV','SPT','OSR','SPT','TRN','TRN',OPS_USD_TYPE) VALUE, " & _
		"DECODE(DECODE(OPS_USD_TEAM,'SPT','SPT','SRV','SPT','OSR','SPT','TRN','TRN',OPS_USD_TYPE),'ACC','Accounting','CRT','Customer Relations','DOC','Documents','GRP','Group','OPS','Operations','OSS','Operations Support','POP','Product Operations','RES','Reservations','SPT','Support Desk','TRN','Training','HRA','Human Resources','MIS','Information Technology') || ' (' || DECODE(OPS_USD_TEAM,'SPT','SPT','SRV','SPT','OSR','SPT','TRN','TRN',OPS_USD_TYPE) || ')' DESCRIPTION " & _
		"FROM OPS_USER " & _
		"JOIN OPS_USER_DETAIL " & _
		"ON OPS_USR_ID = OPS_USD_SUPERVISOR " & _
		"AND OPS_USD_DIS_DATE >= TO_DATE(?,'MM/DD/YYYY') " & _
		"WHERE OPS_USD_LOCATION IN ('MOT','WFD','WFH') " & _
		"ORDER BY DESCRIPTION"
		cmd.CommandText = SQLstmt
		cmd.Parameters(0).value = PARAMETER_DATE
		Set RSSELECT = cmd.Execute(SQLstmt)
	%>
		<% Do While Not RSSELECT.EOF %>
			<option value="DEPT_<%=RSSELECT("VALUE")%>"><%=RSSELECT("DESCRIPTION")%></option>
			<% RSSELECT.MoveNext %>
		<% Loop %>
		<% Set RSSELECT = Nothing %>
	</optgroup>
	<%
		SQLstmt = "SELECT DISTINCT " & _
		"DECODE(OPS_USD_TYPE,'GRP','AGT',DECODE(OPS_USD_TEAM,'LDA','LED',OPS_USD_JOB)) VALUE, " & _
		"DECODE(DECODE(OPS_USD_TYPE,'GRP','AGT',DECODE(OPS_USD_TEAM,'LDA','LED',OPS_USD_JOB)),'AGT','Associate (AGT)','LED','Lead (LED)','ANL','Analyst (ANL)','ADM','Admin (ADM)','DIR','Director (DIR)','MGR','Manager (MGR)','SUP','Supervisor (SUP)',DECODE(OPS_USD_TYPE,'GRP','AGT',DECODE(OPS_USD_TEAM,'LDA','LED',OPS_USD_JOB))) DESCRIPTION " & _
		"FROM OPS_USER " & _
		"JOIN OPS_USER_DETAIL " & _
		"ON OPS_USR_ID = OPS_USD_SUPERVISOR " & _
		"AND OPS_USD_DIS_DATE >= TO_DATE(?,'MM/DD/YYYY') " & _
		"WHERE OPS_USD_LOCATION IN ('MOT','WFD','WFH') " & _
		"ORDER BY DESCRIPTION"
		cmd.CommandText = SQLstmt
		cmd.Parameters(0).value = PARAMETER_DATE
		Set RSSELECT = cmd.Execute(SQLstmt)
	%>
	<% If Not RSSELECT.EOF Then %>
		<optgroup label="Job">
			<% Do While Not RSSELECT.EOF %>
				<option value="JOB_<%=RSSELECT("VALUE")%>"><%=RSSELECT("DESCRIPTION")%></option>
				<% RSSELECT.MoveNext %>
			<% Loop %>
		</optgroup>
	<% End If %>
	<% Set RSSELECT = Nothing %>
	<optgroup label="Location">
		<option value="LOC_MOT">Minot (MOT)</option>
		<option value="LOC_WFH">Work From Home (WFH)</option>
		<option value="LOC_WFD">Work From Distance (WFD)</option>
		<% If PULSE_SECURITY > 5 Then %>
			<option value="LOC_MSP">Minneapolis/Edina (MSP)</option>
			<option value="LOC_ATL">Atlanta (ATL)</option>
		<% End If %>
	</optgroup>
	<% 
		SQLstmt = "SELECT DISTINCT OPS_USR_HIRE_DATE VALUE, OPS_USR_HIRE_DATE DESCRIPTION " & _
		"FROM OPS_USER " & _
		"JOIN OPS_USER_DETAIL " & _
		"ON OPS_USR_ID = OPS_USD_OPS_USR_ID " & _
		"AND TO_DATE(?,'MM/DD/YYYY') BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE " & _
		"WHERE OPS_USD_LOCATION IN ('MOT','WFD','WFH') " & _
		"AND OPS_USD_CLASS IN ('RGFT','RGPT','PT<30','LEAVE') " & _
		"AND OPS_USD_TYPE <> 'HRA' " & _
		"AND OPS_USD_PAY_RATE > 0 " & _
		"AND OPS_USR_HIRE_DATE BETWEEN ADD_MONTHS(TO_DATE(?,'MM/DD/YYYY'),-24) AND TO_DATE(?,'MM/DD/YYYY') " & _
		"ORDER BY DESCRIPTION"
		cmd.CommandText = SQLstmt
		cmd.Parameters(0).value = PARAMETER_DATE
		cmd.Parameters(1).value = PARAMETER_DATE
		cmd.Parameters(2).value = PARAMETER_DATE
		Set RSSELECT = cmd.Execute(SQLstmt)
	%>
	<% If Not RSSELECT.EOF Then %>
		<optgroup label="New Hire Class">
		<% Do While Not RSSELECT.EOF %>
			<option value="HIRE_<%=RSSELECT("VALUE")%>"><%=RSSELECT("DESCRIPTION")%></option>
			<% RSSELECT.MoveNext %>
		<% Loop %>
		</optgroup>
	<% End If %>
	<% Set RSSELECT = Nothing %>
<% End If %>
<% Set cmd = Nothing %>
<% Conn.Close %>
<% Set Conn = Nothing %>