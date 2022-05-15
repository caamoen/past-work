<%
	SQLstmt = "SELECT " & _
	"USE_AGENT, " & _
	"AGENT_NAME, " & _
	"SUPERVISOR_NAME, " & _
	"USE_WORKGROUP, " & _
	"USE_CLASS, " & _
	"USE_LOCATION, " & _
	"NVL2(NOTE_USR_ID,1,0) NOTE_BOOL, " & _
	"NVL2(WAIVER_USR_ID,1,0) WAIVER_BOOL, " & _
	"COUNT(*) OVER () AGENT_COUNT " & _
	"FROM " & _
	"( " & _
		"SELECT DISTINCT " & _
		"AGT.OPS_USR_ID USE_AGENT, " & _
		"AGT.OPS_USR_NAME AGENT_NAME, " & _
		"SUP.OPS_USR_NAME SUPERVISOR_NAME, " & _
		"CASE " & _
			"WHEN OPS_USD_TYPE = 'RES' OR OPS_USD_TEAM = 'SPT' THEN DECODE(RES_RTE_RES_RTG_ID,4,'SPT',5,'SPT',13,'SPT','RES') || DECODE(RES_RTE_RES_RTG_ID,1,DECODE(OPS_USD_TEAM,'SLS','SLS','RES'),2,DECODE(OPS_USD_TEAM,'SLS','SLS','RES'),3,DECODE(OPS_USD_TEAM,'SLS','SLS','RES'),4,'VSS',5,'ASR',10,'SRV',13,'OSR','SRV') " & _
			"ELSE OPS_USD_TYPE || DECODE(OPS_USD_TYPE,'GRP',OPS_USD_JOB,OPS_USD_TEAM) " & _
		"END WORKGROUP_VALUE, " & _
		"CASE " & _
			"WHEN OPS_USD_TYPE = 'RES' OR OPS_USD_TEAM = 'SPT' THEN DECODE(RES_RTE_RES_RTG_ID,1,DECODE(OPS_USD_TEAM,'SLS','Sales Specialty','RES Sales'),2,DECODE(OPS_USD_TEAM,'SLS','Sales Specialty','RES Sales'),3,DECODE(OPS_USD_TEAM,'SLS','Sales Specialty','RES Sales'),4,'SPT VSS',5,'SPT ASR',10,'Elite Service',13,'SPT OSR') " & _
			"WHEN OPS_USD_TYPE = 'POP' AND OPS_USD_TEAM = 'POC' THEN 'Product Ops' " & _
			"WHEN OPS_USD_TYPE = 'POP' AND OPS_USD_TEAM = 'DOC' THEN 'Documents' " & _
			"WHEN OPS_USD_TYPE = 'OSS' AND OPS_USD_TEAM = 'SKD' THEN 'Schedule Changes' " & _
			"WHEN OPS_USD_TYPE = 'OSS' AND OPS_USD_TEAM = 'PRD' THEN 'Product Support' " & _
			"WHEN OPS_USD_TYPE = 'OSS' AND OPS_USD_TEAM = 'AIR' THEN 'Air Support' " & _
			"WHEN OPS_USD_TYPE = 'OSS' AND OPS_USD_TEAM = 'LED' THEN 'OSS Leads' " & _
			"WHEN OPS_USD_TYPE = 'OPS' THEN 'OPS Desk' " & _
			"WHEN OPS_USD_TYPE = 'GRP' AND OPS_USD_JOB = 'SPC' THEN 'Group Product' " & _
			"WHEN OPS_USD_TYPE = 'GRP' AND OPS_USD_JOB = 'GSP' THEN 'Group Service' " & _
			"WHEN OPS_USD_TYPE = 'GRP' AND OPS_USD_JOB = 'GSM' THEN 'Group Sales' " & _
			"WHEN OPS_USD_TYPE = 'DOC' THEN 'Facilities' " & _
			"WHEN OPS_USD_TYPE = 'CRT' THEN 'Customer Relations' " & _
			"WHEN OPS_USD_TYPE = 'ACC' AND OPS_USD_TEAM = 'REC' THEN 'Account Receivable' " & _
			"WHEN OPS_USD_TYPE = 'ACC' AND OPS_USD_TEAM = 'PAY' THEN 'Account Payable' " & _
			"WHEN OPS_USD_TYPE = 'ACC' AND OPS_USD_TEAM = 'LDA' THEN 'Account Leads' " & _
		"END || " & _
		"CASE " & _
			"WHEN OPS_USD_TYPE = 'RES' AND RES_RTE_RES_RTG_ID = 5 AND TRN.OPS_ASN_OPS_ASM_ID = 41 THEN ' - Jira' " & _
			"WHEN OPS_USD_TYPE = 'RES' AND RES_RTE_RES_RTG_ID = 5 AND TRN.OPS_ASN_OPS_ASM_ID = 45 THEN ' - TAS' " & _
			"WHEN OPS_USD_TYPE = 'RES' AND RES_RTE_RES_RTG_ID IN (1,2,3) AND TRN.OPS_ASN_OPS_ASM_ID = 43 THEN ' - Dual' " & _
			"WHEN OPS_USD_TYPE = 'RES' AND RES_RTE_RES_RTG_ID IN (1,2,3) AND TRN.OPS_ASN_OPS_ASM_ID = 5 THEN ' - Intl' " & _
			"WHEN OPS_USD_TYPE = 'RES' AND RES_RTE_RES_RTG_ID IN (1,2,3) AND TRN.OPS_ASN_OPS_ASM_ID = 4 THEN ' - Haw' " & _
			"WHEN OPS_USD_TYPE = 'RES' AND RES_RTE_RES_RTG_ID IN (1,2,3) AND TRN.OPS_ASN_OPS_ASM_ID = 3 THEN ' - Car' " & _
		"END USE_WORKGROUP, " & _
		"DECODE(OPS_USD_CLASS,'RGFT','Reg Full-Time','RGPT','Reg Part-time','PT<30','PT Less Than 30','LEAVE','On Leave') USE_CLASS, " & _
		"DECODE(OPS_USD_LOCATION,'MOT','In-House','WFH','Work From Home','WFD','Work From Distance') USE_LOCATION " & _
		"FROM OPS_USER AGT " & _
		"JOIN OPS_USER_DETAIL " & _
		"ON AGT.OPS_USR_ID = OPS_USD_OPS_USR_ID " & _
		"AND TO_DATE(?,'MM/DD/YYYY') BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE " & _
		"JOIN OPS_USER SUP " & _
		"ON SUP.OPS_USR_ID = OPS_USD_SUPERVISOR "
		If PARAMETER_TIMES <> "" Then
			SQLstmt = SQLstmt & "JOIN " & _
			"( " & _
				"SELECT DISTINCT OPS_SCI_OPS_USR_ID " & _
				"FROM " & _
				"( " & _
					"SELECT " & _
					"OPS_SCI_OPS_USR_ID, " & _
					"CONNECT_BY_ROOT(OPS_SCI_START) START_TIME, " & _
					"OPS_SCI_END END_TIME " & _
					"FROM " & _
					"( " & _
						"SELECT " & _
						"OPS_SCI_OPS_USR_ID, " & _
						"OPS_SCI_START, " & _
						"OPS_SCI_END, " & _
						"CASE " & _
							"WHEN LAG(OPS_SCI_START) OVER (PARTITION BY OPS_SCI_OPS_USR_ID ORDER BY OPS_SCI_START) IS NULL " & _
							"OR LAG(OPS_SCI_END) OVER (PARTITION BY OPS_SCI_OPS_USR_ID ORDER BY OPS_SCI_START) <> OPS_SCI_START THEN 1 " & _
						"END START_FLAG " & _
						"FROM " & _
						"( " & _
							"SELECT " & _
							"OPS_SCI_OPS_USR_ID, " & _
							"OPS_SCI_TYPE, " & _
							"OPS_SCI_START, " & _
							"OPS_SCI_END, " & _
							"CASE " & _
								"WHEN " & _
								"( " & _
									"LAG(OPS_SCI_START) OVER (PARTITION BY OPS_SCI_OPS_USR_ID ORDER BY OPS_SCI_START) IS NULL " & _
									"OR LAG(OPS_SCI_END) OVER (PARTITION BY OPS_SCI_OPS_USR_ID ORDER BY OPS_SCI_START) <> OPS_SCI_START " & _
									"OR LEAD(OPS_SCI_END) OVER (PARTITION BY OPS_SCI_OPS_USR_ID ORDER BY OPS_SCI_START) IS NULL " & _
									"OR LEAD(OPS_SCI_START) OVER (PARTITION BY OPS_SCI_OPS_USR_ID ORDER BY OPS_SCI_START) <> OPS_SCI_END " & _
								") " & _
								"AND OPS_SCI_TYPE IN ('LNCH','LNFL') " & _
								"THEN 1 " & _
							"END DELETE_FLAG " & _
							"FROM OPS_SCHEDULE_INFO " & _
							"WHERE TO_DATE(OPS_SCI_START) BETWEEN TO_DATE(?,'MM/DD/YYYY') AND TO_DATE(?,'MM/DD/YYYY') " & _
							"AND OPS_SCI_STATUS = 'APP' " & _
							"AND " & _
							"( " & _
								"OPS_SCI_TYPE IN ('BASE','PICK','ADDT','EXTD','HOLW','MEET','PRES','PROJ','TRAN','FAMP','WFHU','MLTU','NEWH','LNCH','LNFL') " & _
								"OR " & _
								"( " & _
									"OPS_SCI_TYPE = 'OTRG' " & _
									"AND REGEXP_INSTR(UPPER(OPS_SCI_NOTES),'SPLV|HRPP') = 0 " & _
								") " & _
							") " & _
							"AND OPS_SCI_END > OPS_SCI_START " & _
						") " & _
						"WHERE DELETE_FLAG IS NULL " & _
					") " & _
					"WHERE CONNECT_BY_ISLEAF = 1 " & _
					"START WITH START_FLAG = 1 " & _
					"CONNECT BY OPS_SCI_START = PRIOR OPS_SCI_END " & _
					"AND OPS_SCI_OPS_USR_ID = PRIOR OPS_SCI_OPS_USR_ID " & _
				") " & _
				"WHERE "
				SEARCH_PARAMETER_ARRAY(i) = PARAMETER_DATE - 1
				SEARCH_PARAMETER_ARRAY(i+1) = PARAMETER_DATE + 1
				i = i + 2
				USE_ARRAY = Split(PARAMETER_TIMES,",")
				For j = 0 to UBound(USE_ARRAY)
					If Left(USE_ARRAY(j),1) = "S" Then
						If Mid(USE_ARRAY(j),2,2) = "NH" Then
							If j = 0 Then	
								SQLstmt = SQLstmt & "START_TIME BETWEEN CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE) + ((60 * ?)*ROUND(TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS')/(60 * ?)) - TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS'))/86400 AND CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE) + (1/24) + ((60 * ?)*ROUND(TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS')/(60 * ?)) - TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS'))/86400 "
							Else
								SQLstmt = SQLstmt & "OR START_TIME BETWEEN CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE) + ((60 * ?)*ROUND(TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS')/(60 * ?)) - TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS'))/86400 AND CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE) + (1/24) + ((60 * ?)*ROUND(TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS')/(60 * ?)) - TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS'))/86400 "
							End If
							SEARCH_PARAMETER_ARRAY(i) = Mid(USE_ARRAY(j),4)
							SEARCH_PARAMETER_ARRAY(i+1) = Mid(USE_ARRAY(j),4)
							SEARCH_PARAMETER_ARRAY(i+2) = Mid(USE_ARRAY(j),4)
							SEARCH_PARAMETER_ARRAY(i+3) = Mid(USE_ARRAY(j),4)
							i = i + 4
						Else
							If j = 0 Then	
								SQLstmt = SQLstmt & "START_TIME BETWEEN TO_DATE(?,'MMDDYYYYHH24MI') AND TO_DATE(?,'MMDDYYYYHH24MI') "
							Else
								SQLstmt = SQLstmt & "OR START_TIME BETWEEN TO_DATE(?,'MMDDYYYYHH24MI') AND TO_DATE(?,'MMDDYYYYHH24MI') "
							End If
							SEARCH_PARAMETER_ARRAY(i) = Mid(USE_ARRAY(j),2,12)
							SEARCH_PARAMETER_ARRAY(i+1) = Mid(USE_ARRAY(j),14,12)
							i = i + 2
						End If
					Elseif Left(USE_ARRAY(j),1) = "E" Then
						If Mid(USE_ARRAY(j),2,2) = "NH" Then
							If j = 0 Then	
								SQLstmt = SQLstmt & "END_TIME BETWEEN CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE) + ((60 * ?)*ROUND(TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS')/(60 * ?)) - TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS'))/86400 AND CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE) + (1/24) + ((60 * ?)*ROUND(TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS')/(60 * ?)) - TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS'))/86400 "
							Else
								SQLstmt = SQLstmt & "OR END_TIME BETWEEN CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE) + ((60 * ?)*ROUND(TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS')/(60 * ?)) - TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS'))/86400 AND CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE) + (1/24) + ((60 * ?)*ROUND(TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS')/(60 * ?)) - TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS'))/86400 "
							End If
							SEARCH_PARAMETER_ARRAY(i) = Mid(USE_ARRAY(j),4)
							SEARCH_PARAMETER_ARRAY(i+1) = Mid(USE_ARRAY(j),4)
							SEARCH_PARAMETER_ARRAY(i+2) = Mid(USE_ARRAY(j),4)
							SEARCH_PARAMETER_ARRAY(i+3) = Mid(USE_ARRAY(j),4)
							i = i + 4
						Else
							If j = 0 Then	
								SQLstmt = SQLstmt & "END_TIME BETWEEN TO_DATE(?,'MMDDYYYYHH24MI') AND TO_DATE(?,'MMDDYYYYHH24MI') "
							Else
								SQLstmt = SQLstmt & "OR END_TIME BETWEEN TO_DATE(?,'MMDDYYYYHH24MI') AND TO_DATE(?,'MMDDYYYYHH24MI') "
							End If
							SEARCH_PARAMETER_ARRAY(i) = Mid(USE_ARRAY(j),2,12)
							SEARCH_PARAMETER_ARRAY(i+1) = Mid(USE_ARRAY(j),14,12)
							i = i + 2
						End If
					Else
						If j = 0 Then	
							SQLstmt = SQLstmt & "CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE) + ((60 * ?)*ROUND(TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS')/(60 * ?)) - TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS'))/86400 BETWEEN START_TIME AND END_TIME "
						Else
							SQLstmt = SQLstmt & "OR CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE) + ((60 * ?)*ROUND(TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS')/(60 * ?)) - TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'SSSSS'))/86400 BETWEEN START_TIME AND END_TIME "
						End If
						SEARCH_PARAMETER_ARRAY(i) = Mid(USE_ARRAY(j),3)
						SEARCH_PARAMETER_ARRAY(i+1) = Mid(USE_ARRAY(j),3)
						i = i + 2
					End If
				Next
			SQLstmt = SQLstmt & ") ST " & _
			"ON AGT.OPS_USR_ID = ST.OPS_SCI_OPS_USR_ID "
		End If
		If PARAMETER_SHIFT <> "" Then
			SQLstmt = SQLstmt & "JOIN OPS_SCHEDULE_INFO SC " & _
			"ON AGT.OPS_USR_ID = SC.OPS_SCI_OPS_USR_ID " & _
			"AND TO_DATE(SC.OPS_SCI_START) = TO_DATE(?,'MM/DD/YYYY') " & _
			"AND SC.OPS_SCI_STATUS = 'APP' " & _
			"AND SC.OPS_SCI_TYPE IN ("
			SEARCH_PARAMETER_ARRAY(i) = PARAMETER_DATE
			i = i + 1
			USE_ARRAY = Split(PARAMETER_SHIFT,",")
			For j = 0 to UBound(USE_ARRAY)
				If j <> UBound(USE_ARRAY) Then
					SQLstmt = SQLstmt & "?,"
				Else
					SQLstmt = SQLstmt & "?) "
				End If
				SEARCH_PARAMETER_ARRAY(i) = USE_ARRAY(j)
				i = i + 1
			Next
		End If
		SQLstmt = SQLstmt & "LEFT JOIN " & _
		"( " & _
			"SELECT " & _
			"RES_RTE_OPS_USR_ID, " & _
			"MAX(DECODE(RES_RTE_RES_RTG_ID,0,NULL,RES_RTE_RES_RTG_ID)) KEEP (DENSE_RANK LAST ORDER BY RTE_ORDERING) RES_RTE_RES_RTG_ID " & _
			"FROM " & _
			"( " & _
				"SELECT " & _
				"RES_RTE_OPS_USR_ID, " & _
				"RES_RTE_RES_RTG_ID, " & _
				"1 RTE_ORDERING " & _
				"FROM RES_ROUTING " & _
				"WHERE RES_RTE_YEAR = TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE)-(6/24),'YYYY') " & _
				"AND RES_RTE_MONTH = TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE)-(6/24),'MM') " & _
				"UNION ALL " & _
				"SELECT TO_NUMBER(SYS_CDD_NAME), " & _
				"TO_NUMBER(REGEXP_SUBSTR(SYS_CDD_VALUE,'[^;]+',1,2)), " & _
				"2 " & _
				"FROM SYS_CODE_DETAIL " & _
				"WHERE SYS_CDD_SYS_CDM_ID = 508 " & _
				"AND TO_DATE(?,'MM/DD/YYYY') >= TO_DATE(REGEXP_SUBSTR(SYS_CDD_VALUE,'[^;]+',1,1),'MM/DD/YYYY') " & _
				"AND TO_DATE(REGEXP_SUBSTR(SYS_CDD_VALUE,'[^;]+',1,1),'MM/DD/YYYY') >= TO_DATE(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE)) " & _
				"UNION ALL " & _
				"SELECT " & _
				"RES_STA_OPS_USR_ID, " & _
				"RES_STA_RES_RTD_ID, " & _
				"2 " & _
				"FROM RES_STATS_INCENTIVE " & _
				"WHERE RES_STA_DATE = TO_DATE(?,'MM/DD/YYYY') " & _
			") " & _
			"GROUP BY RES_RTE_OPS_USR_ID " & _
		") " & _
		"ON RES_RTE_OPS_USR_ID = OPS_USD_OPS_USR_ID " & _
		"LEFT JOIN " & _
		"( " & _
			"SELECT " & _
			"OPS_ASN_OPS_USR_ID, " & _
			"MAX(NULLIF(OPS_ASN_OPS_ASM_ID,52)) OPS_ASN_OPS_ASM_ID, " & _
			"MAX(DECODE(OPS_ASN_OPS_ASM_ID,52,1,0)) CHAT_ELIGIBLE " & _
			"FROM OPS_ASSIGNMENT " & _
			"WHERE OPS_ASN_OPS_ASM_ID IN (1,2,3,4,5,39,40,41,43,45,50,52) " & _
			"AND TO_DATE(?,'MM/DD/YYYY') BETWEEN OPS_ASN_EFF_DATE AND OPS_ASN_DIS_DATE " & _
			"GROUP BY OPS_ASN_OPS_USR_ID " & _
		") TRN " & _
		"ON TRN.OPS_ASN_OPS_USR_ID = OPS_USD_OPS_USR_ID " & _
		"WHERE OPS_USD_PAY_RATE > 0 "
		SEARCH_PARAMETER_ARRAY(i) = PARAMETER_DATE
		SEARCH_PARAMETER_ARRAY(i+1) = PARAMETER_DATE
		SEARCH_PARAMETER_ARRAY(i+2) = PARAMETER_DATE
		i = i + 3
		If PARAMETER_AGENT <> "" and Instr(PARAMETER_AGENT,"ALL") = 0 Then
			SQLstmt = SQLstmt & "AND AGT.OPS_USR_ID IN ("
			USE_ARRAY = Split(PARAMETER_AGENT,",")
			For j = 0 to UBound(USE_ARRAY)
				If j <> UBound(USE_ARRAY) Then
					SQLstmt = SQLstmt & "?,"
				Else
					SQLstmt = SQLstmt & "?) "
				End If
				SEARCH_PARAMETER_ARRAY(i) = USE_ARRAY(j)
				i = i + 1
			Next
		End If
		If PARAMETER_SUPERVISOR <> "" Then
			SQLstmt = SQLstmt & "AND OPS_USD_SUPERVISOR IN ("
			USE_ARRAY = Split(PARAMETER_SUPERVISOR,",")
			For j = 0 to UBound(USE_ARRAY)
				If j <> UBound(USE_ARRAY) Then
					SQLstmt = SQLstmt & "?,"
				Else
					SQLstmt = SQLstmt & "?) "
				End If
				SEARCH_PARAMETER_ARRAY(i) = USE_ARRAY(j)
				i = i + 1
			Next
		End If
		If PARAMETER_DEPARTMENT <> "" Then
			SQLstmt = SQLstmt & "AND DECODE(OPS_USD_TEAM,'SPT','SPT','SRV','SPT','OSR','SPT',OPS_USD_TYPE) IN ("
			USE_ARRAY = Split(PARAMETER_DEPARTMENT,",")
			For j = 0 to UBound(USE_ARRAY)
				If j <> UBound(USE_ARRAY) Then
					SQLstmt = SQLstmt & "?,"
				Else
					SQLstmt = SQLstmt & "?) "
				End If
				SEARCH_PARAMETER_ARRAY(i) = USE_ARRAY(j)
				i = i + 1
			Next
		Elseif PULSE_SECURITY < 5 Then
			SQLstmt = SQLstmt & "AND OPS_USD_TYPE IN ("
			USE_ARRAY = Split(PULSE_DEPARTMENT,",")
			For j = 0 to UBound(USE_ARRAY)
				If j <> UBound(USE_ARRAY) Then
					SQLstmt = SQLstmt & "?,"
				Else
					SQLstmt = SQLstmt & "?) "
				End If
				SEARCH_PARAMETER_ARRAY(i) = USE_ARRAY(j)
				i = i + 1
			Next
		End If
		If PARAMETER_LOCATION <> "" Then
			SQLstmt = SQLstmt & "AND OPS_USD_LOCATION IN ("
			USE_ARRAY = Split(PARAMETER_LOCATION,",")
			For j = 0 to UBound(USE_ARRAY)
				If j <> UBound(USE_ARRAY) Then
					SQLstmt = SQLstmt & "?,"
				Else
					SQLstmt = SQLstmt & "?) "
				End If
				SEARCH_PARAMETER_ARRAY(i) = USE_ARRAY(j)
				i = i + 1
			Next
		Else
			SQLstmt = SQLstmt & "AND OPS_USD_LOCATION IN ('MOT','WFD','WFH') "
		End If
		If PARAMETER_CLASS <> "" Then
			SQLstmt = SQLstmt & "AND OPS_USD_CLASS IN ("
			USE_ARRAY = Split(PARAMETER_CLASS,",")
			For j = 0 to UBound(USE_ARRAY)
				If j <> UBound(USE_ARRAY) Then
					SQLstmt = SQLstmt & "?,"
				Else
					SQLstmt = SQLstmt & "?) "
				End If
				SEARCH_PARAMETER_ARRAY(i) = USE_ARRAY(j)
				i = i + 1
			Next
		Else
			SQLstmt = SQLstmt & "AND OPS_USD_CLASS IN ('RGFT','RGPT','PT<30','LEAVE') "
		End If
		If PARAMETER_JOB <> "" Then
			SQLstmt = SQLstmt & "AND DECODE(OPS_USD_TYPE,'GRP','AGT',DECODE(OPS_USD_TEAM,'LDA','LED',OPS_USD_JOB)) IN ("
			USE_ARRAY = Split(PARAMETER_JOB,",")
			For j = 0 to UBound(USE_ARRAY)
				If j <> UBound(USE_ARRAY) Then
					SQLstmt = SQLstmt & "?,"
				Else
					SQLstmt = SQLstmt & "?) "
				End If
				SEARCH_PARAMETER_ARRAY(i) = USE_ARRAY(j)
				i = i + 1
			Next
		End If
		If PARAMETER_HIRE <> "" Then
			SQLstmt = SQLstmt & "AND AGT.OPS_USR_HIRE_DATE IN ("
			USE_ARRAY = Split(PARAMETER_HIRE,",")
			For j = 0 to UBound(USE_ARRAY)
				If j <> UBound(USE_ARRAY) Then
					SQLstmt = SQLstmt & "TO_DATE(?,'MM/DD/YYYY'),"
				Else
					SQLstmt = SQLstmt & "TO_DATE(?,'MM/DD/YYYY')) "
				End If
				SEARCH_PARAMETER_ARRAY(i) = USE_ARRAY(j)
				i = i + 1
			Next
		End If
		If PARAMETER_TRAINING <> "" Then
			SQLstmt = SQLstmt & "AND RES_RTE_RES_RTG_ID IN (1,2,3) " & _
			"AND TRN.OPS_ASN_OPS_ASM_ID IN (3,4,5,43) " & _
			"AND TRN.OPS_ASN_OPS_ASM_ID IN ("
			USE_ARRAY = Split(PARAMETER_TRAINING,",")
			For j = 0 to UBound(USE_ARRAY)
				If j <> UBound(USE_ARRAY) Then
					SQLstmt = SQLstmt & "?,"
				Else
					SQLstmt = SQLstmt & "?) "
				End If
				SEARCH_PARAMETER_ARRAY(i) = USE_ARRAY(j)
				i = i + 1
			Next
		End If
		If PARAMETER_CHAT <> "" Then
			SQLstmt = SQLstmt & "AND RES_RTE_RES_RTG_ID = 10 " & _
			"AND TRN.CHAT_ELIGIBLE = 1 "
		End If
		If PARAMETER_ROUTING <> "" Then
			SQLstmt = SQLstmt & "AND RES_RTE_RES_RTG_ID IN ("
			USE_ARRAY = Split(PARAMETER_ROUTING,",")
			For j = 0 to UBound(USE_ARRAY)
				If j <> UBound(USE_ARRAY) Then
					SQLstmt = SQLstmt & "?,"
				Else
					SQLstmt = SQLstmt & "?) "
				End If
				SEARCH_PARAMETER_ARRAY(i) = USE_ARRAY(j)
				i = i + 1
			Next
		End If
	SQLstmt = SQLstmt & ") " & _
	"LEFT JOIN " & _
	"( " & _
		"SELECT DISTINCT TO_NUMBER(TRIM(REGEXP_SUBSTR(RES_DLN_TEXT,'[^-]+',1,1))) NOTE_USR_ID " & _
		"FROM RES_DAILY_STATS_NOTES " & _
		"WHERE " & _
		"( " & _
			"( " & _
				"RES_DLN_DATE = TO_DATE(?,'MM/DD/YYYY') " & _
				"AND " & _
				"( " & _
					"TRIM(RES_DLN_TYPE) IN ('LATE','FLEX','SHFT','STRT','END','NFLX','OFLX','GAP','LMIN','OLAP','ABS','FWP','OTH','OUT','WFH','NA') " & _
					"OR " & _
					"( " & _
						"TRIM(RES_DLN_TYPE) = 'LWAV' " & _
						"AND DECODE(INSTR(RES_DLN_TEXT,'-'),0,NULL,TRIM(SUBSTR(RES_DLN_TEXT,INSTR(RES_DLN_TEXT,'-')+1))) IS NOT NULL " & _
					") " & _
				") " & _
			") " & _
			"OR " & _
			"( " & _
				"RES_DLN_DATE BETWEEN TO_DATE(?,'MM/DD/YYYY') - TO_CHAR(TO_DATE(?,'MM/DD/YYYY'),'D') + 1 AND TO_DATE(?,'MM/DD/YYYY') - TO_CHAR(TO_DATE(?,'MM/DD/YYYY'),'D') + 7 " & _
				"AND TRIM(RES_DLN_TYPE) = 'BWHC' " & _
			") " & _
		") " & _
	") " & _
	"ON USE_AGENT = NOTE_USR_ID " & _
	"LEFT JOIN " & _
	"( " & _
		"SELECT DISTINCT TO_NUMBER(TRIM(REGEXP_SUBSTR(RES_DLN_TEXT,'[^-]+',1,1))) WAIVER_USR_ID " & _
		"FROM RES_DAILY_STATS_NOTES " & _
		"WHERE RES_DLN_DATE = TO_DATE(?,'MM/DD/YYYY') " & _
		"AND TRIM(RES_DLN_TYPE) = 'LWAV' " & _
	") " & _
	"ON USE_AGENT = WAIVER_USR_ID "
	SEARCH_PARAMETER_ARRAY(i) = PARAMETER_DATE
	SEARCH_PARAMETER_ARRAY(i+1) = PARAMETER_DATE
	SEARCH_PARAMETER_ARRAY(i+2) = PARAMETER_DATE
	SEARCH_PARAMETER_ARRAY(i+3) = PARAMETER_DATE
	SEARCH_PARAMETER_ARRAY(i+4) = PARAMETER_DATE
	SEARCH_PARAMETER_ARRAY(i+5) = PARAMETER_DATE
	i = i + 6
	If PARAMETER_WORKGROUP <> "" Then
		SQLstmt = SQLstmt & "WHERE WORKGROUP_VALUE IN ("
		USE_ARRAY = Split(PARAMETER_WORKGROUP,",")
		For j = 0 to UBound(USE_ARRAY)
			If j <> UBound(USE_ARRAY) Then
				SQLstmt = SQLstmt & "?,"
			Else
				SQLstmt = SQLstmt & "?) "
			End If
			SEARCH_PARAMETER_ARRAY(i) = USE_ARRAY(j)
			i = i + 1
		Next
	End If
	SQLstmt = SQLstmt & "ORDER BY AGENT_NAME"
%>