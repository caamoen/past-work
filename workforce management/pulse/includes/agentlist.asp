<!--#include file="pulseheader.asp"-->
<!--#include file="error.asp"-->

<!--
	NOTES:
		SQL statements in search.asp, pending.asp, and error.asp (all executed as RSAGTLIST) contain/use the following shared fields: 	USE_AGENT, AGENT_NAME, SUPERVISOR_NAME, USE_WORKGROUP, USE_CLASS, USE_LOCATION, AGENT_COUNT, NOTE_BOOL, WAIVER_BOOL
-->
<%
	DATATABLES_BOOL = 0
	LUNCH_ERROR_FLAG = 0
	USE_COLSPAN = 6
	If Request.Querystring("REQUEST") <> "" Then
		REQUEST_TYPE = Request.Querystring("REQUEST")
	Else
		REQUEST_TYPE = "SEARCH"
	End If
	
	If REQUEST_TYPE = "SEARCH" Then
		SEARCH_BOOL = 0

		If Request.Querystring("DATE") <> "" Then
			PARAMETER_DATE = CDate(Request.Querystring("DATE"))
		Else
			PARAMETER_DATE = Date
		End If
		If Request.Querystring("AGENT") <> "" Then
			SEARCH_BOOL = 1
			PARAMETER_AGENT = Request.Querystring("AGENT")
		Else
			PARAMETER_AGENT = ""
		End If
		If Request.Querystring("SUPERVISOR") <> "" Then
			SEARCH_BOOL = 1
			PARAMETER_SUPERVISOR = Request.Querystring("SUPERVISOR")
		Else
			PARAMETER_SUPERVISOR = ""
		End If
		If Request.Querystring("DEPARTMENT") <> "" Then
			SEARCH_BOOL = 1
			PARAMETER_DEPARTMENT = Request.Querystring("DEPARTMENT")
		Else
			PARAMETER_DEPARTMENT = ""
		End If
		If Request.Querystring("WORKGROUP") <> "" Then
			SEARCH_BOOL = 1
			PARAMETER_WORKGROUP = Request.Querystring("WORKGROUP")
		Else
			PARAMETER_WORKGROUP = ""
		End If
		If Request.Querystring("CLASS") <> "" Then
			SEARCH_BOOL = 1
			PARAMETER_CLASS = Request.Querystring("CLASS")
		Else
			PARAMETER_CLASS = ""
		End If
		If Request.Querystring("LOCATION") <> "" Then
			SEARCH_BOOL = 1
			PARAMETER_LOCATION = Request.Querystring("LOCATION")
		Else
			PARAMETER_LOCATION = ""
		End If
		If Request.Querystring("JOB") <> "" Then
			SEARCH_BOOL = 1
			PARAMETER_JOB = Request.Querystring("JOB")
		Else
			PARAMETER_JOB = ""
		End If
		If Request.Querystring("SHIFT") <> "" Then
			SEARCH_BOOL = 1
			PARAMETER_SHIFT = Request.Querystring("SHIFT")
		Else
			PARAMETER_SHIFT = ""
		End If
		If Request.Querystring("TIMES") <> "" Then
			SEARCH_BOOL = 1
			PARAMETER_TIMES = Request.Querystring("TIMES")
		Else
			PARAMETER_TIMES = ""
		End If
		If Request.Querystring("HIRE") <> "" Then
			SEARCH_BOOL = 1
			PARAMETER_HIRE = Request.Querystring("HIRE")
		Else
			PARAMETER_HIRE = ""
		End If
		If Request.Querystring("TRAINING") <> "" Then
			SEARCH_BOOL = 1
			PARAMETER_TRAINING = Request.Querystring("TRAINING")
		Else
			PARAMETER_TRAINING = ""
		End If
		If Request.Querystring("ROUTING") <> "" Then
			SEARCH_BOOL = 1
			PARAMETER_ROUTING = Request.Querystring("ROUTING")
		Else
			PARAMETER_ROUTING = ""
		End If
		If Request.Querystring("CHAT") <> "" Then
			SEARCH_BOOL = 1
			PARAMETER_CHAT = Request.Querystring("CHAT")
		Else
			PARAMETER_CHAT = ""
		End If
		If SEARCH_BOOL = 1 Then
			ReDim SEARCH_PARAMETER_ARRAY(50)	'Parameter index
			SEARCH_PARAMETER_ARRAY(0) = PARAMETER_DATE
			i = 1
		%>
			<!--#include file="search.asp"-->
		<%
			cmd.CommandText = SQLstmt
			i = i - 1
			ReDim Preserve SEARCH_PARAMETER_ARRAY(i)
			For i = 0 to UBound(SEARCH_PARAMETER_ARRAY)
				cmd.Parameters(i).value = SEARCH_PARAMETER_ARRAY(i)
			Next
			Erase SEARCH_PARAMETER_ARRAY
			Set RSAGTLIST = cmd.Execute
		End If
	Elseif REQUEST_TYPE = "PENDING" Then		
		If Request.Querystring("DATE") <> "" Then
			PARAMETER_DATE = CDate(Request.Querystring("DATE"))
		Else
			PARAMETER_DATE = Date
		End If
		If Request.Querystring("WORKGROUP") <> "" Then
			PARAMETER_WORKGROUP = Request.Querystring("WORKGROUP")
		Else
			PARAMETER_WORKGROUP = "SLS"
		End If
	%>
		<!--#include file="pending.asp"-->
	<%		
		If PARAMETER_WORKGROUP = "SLS" Then
			USE_COLSPAN = 8
			cmd.CommandText = SQLPendingSLS
			For i = 0 to 20
				cmd.Parameters(i).value = PARAMETER_DATE
			Next 
			Set RSAGTLIST = cmd.Execute
		Elseif PARAMETER_WORKGROUP = "SRV" Then
			USE_COLSPAN = 8
			cmd.CommandText = SQLPendingSRV
			For i = 0 to 18
				cmd.Parameters(i).value = PARAMETER_DATE
			Next 
			Set RSAGTLIST = cmd.Execute
		Elseif PARAMETER_WORKGROUP = "SPT" Then
			USE_COLSPAN = 8
			cmd.CommandText = SQLPendingSPT
			For i = 0 to 19
				cmd.Parameters(i).value = PARAMETER_DATE
			Next 
			Set RSAGTLIST = cmd.Execute
		Else
			cmd.CommandText = SQLPendingOD
			cmd.Parameters(0).value = PARAMETER_DATE
			cmd.Parameters(1).value = PARAMETER_DATE
			cmd.Parameters(2).value = PARAMETER_DATE
			cmd.Parameters(3).value = PARAMETER_DATE
			cmd.Parameters(4).value = PARAMETER_DATE
			cmd.Parameters(5).value = PARAMETER_DATE
			cmd.Parameters(6).value = PARAMETER_DATE
			cmd.Parameters(7).value = PARAMETER_WORKGROUP
			cmd.Parameters(8).value = PARAMETER_DATE
			Set RSAGTLIST = cmd.Execute
		End If
	Elseif REQUEST_TYPE = "ERROR" Then
		USE_COLSPAN = 6

		If Request.Querystring("DATE") <> "" Then
			PARAMETER_DATE = CDate(Request.Querystring("DATE"))
		Else
			PARAMETER_DATE = Date
		End If
		If Request.Querystring("WORKGROUP") <> "" Then
			PARAMETER_WORKGROUP = Request.Querystring("WORKGROUP")
		Else
			PARAMETER_WORKGROUP = "RES"
		End If	
		If PARAMETER_DATE <= Date Then
			cmd.CommandText = SQLErrorDept
			For i = 0 to 43
				If i = 2 or i = 3 or i = 8 or i = 9 or i = 27 or i = 28 Then
					cmd.Parameters(i).value = PARAMETER_WORKGROUP
				Else
					cmd.Parameters(i).value = PARAMETER_DATE
				End If
			Next
		Else
			cmd.CommandText = SQLErrorDeptFuture
			For i = 0 to 30
				If i = 2 or i = 3 or i = 14 or i = 15 Then
					cmd.Parameters(i).value = PARAMETER_WORKGROUP
				Else
					cmd.Parameters(i).value = PARAMETER_DATE
				End If
			Next		
		End If
		Set RSAGTLIST = cmd.Execute
	End If
%>
	<% If IsObject(RSAGTLIST) Then %>
	<%
		SQLstmt = "SELECT * " & _
		"FROM RES_PHONE_STATS " & _
		"WHERE RES_PHN_TYPE = 'LUN' " & _
		"AND RES_PHN_DATE = TO_DATE(?,'MM/DD/YYYY')"
		cmd.CommandText = SQLstmt
		cmd.Parameters(0).value = PARAMETER_DATE
		Set RSUC = cmd.Execute
		If Not RSUC.EOF Then
			LUNCH_ERROR_FLAG = 1
		End If
		Set RSUC = Nothing
	%>
		<form id="PULSE_FORM" data-request="<%=REQUEST_TYPE%>" data-workgroup="<%=PARAMETER_WORKGROUP%>" action="includes/formhandler.asp" method="post">
			<input type="hidden" id="NEWLINE_ID" value="0"/>
			<input type="hidden" id="SCHEDULEID_LIST" name="SCHEDULEID_LIST" value=""/>
			<input type="hidden" name="FORM_DATE" value="<%=PARAMETER_DATE%>"/>
			<div id="PULSE_FORM_DIV" class="table-responsive" style="margin-bottom:1rem;">
				<table id="PULSE_TABLE" class="table table-bordered center" style="margin-bottom:0;">
					<caption class="th-color <% If PARAMETER_DATE = Date Then %>today-color-background<% Else %>past-color-background<% End If %>">
						<% If REQUEST_TYPE = "SEARCH" Then %>
							User Search - <%=FormatDateTime(Now,3)%>
						<% Elseif REQUEST_TYPE = "PENDING" Then %>
							<%=DepartmentName(PARAMETER_WORKGROUP)%> Pending List - <%=FormatDateTime(Now,3)%>
						<% Elseif REQUEST_TYPE = "ERROR" Then %>
							<%=DepartmentName(PARAMETER_WORKGROUP)%> Error List - <%=FormatDateTime(Now,3)%>
						<% End If %>
						<% 
							COUNT_TEXT = ""
							If Not RSAGTLIST.EOF Then
								If RSAGTLIST("AGENT_COUNT") = "1" Then
									COUNT_TEXT = "1 Associate Found"
								Else
									DATATABLES_BOOL = 1
									COUNT_TEXT = RSAGTLIST("AGENT_COUNT") & " Associates Found"
								End If
							End If
						%>
						<div style="float:right"><%=COUNT_TEXT%></div>
					</caption>
					<thead>
						<tr class="th-color <% If PARAMETER_DATE = Date Then %>today-color-background<% Else %>past-color-background<% End If %>">
							<% If REQUEST_TYPE = "PENDING" and (PARAMETER_WORKGROUP = "SLS" or PARAMETER_WORKGROUP = "SRV" or PARAMETER_WORKGROUP = "SPT") Then %>
								<th style="width:10%">Priority</th>
								<th style="width:15%">Associate</th>
								<th class="mobile-hide" style="width:15%">Workgroup</th>
								<% If LUNCH_ERROR_FLAG = 0 Then %>
									<th class="mobile-hide" style="width:10%">PTO Balance</th>
								<% Else %>
									<th class="mobile-hide" style="width:10%">Lunch Minutes</th>
								<% End If %>
								<th class="mobile-hide" style="width:10%">UNP Remaining</th>
								<th style="width:15%">Logins</th>
								<th style="width:15%">Shifts</th>							
								<th style="width:10%">Notes</th>
							<% Else %>
								<th style="width:20%">Associate</th>
								<th class="mobile-hide" style="width:15%">Workgroup</th>
								<% If LUNCH_ERROR_FLAG = 0 Then %>
									<th class="mobile-hide" style="width:15%">PTO Balance</th>
								<% Else %>
									<th class="mobile-hide" style="width:15%">Lunch Minutes</th>
								<% End If %>
								<th style="width:15%">Logins</th>
								<th style="width:20%">Shifts</th>
								<th style="width:15%">Notes</th>
							<% End If %>
						</tr>
					</thead>
					<% If Not RSAGTLIST.EOF Then %>
						<tbody>
						<% Do While Not RSAGTLIST.EOF %>
							<tr id="AGENTROW_<%=RSAGTLIST("USE_AGENT")%>_<%=Right("0" & Month(PARAMETER_DATE),2) & Right("0" & Day(PARAMETER_DATE),2) & Year(PARAMETER_DATE)%>">
								<% If REQUEST_TYPE = "PENDING" and (PARAMETER_WORKGROUP = "SLS" or PARAMETER_WORKGROUP = "SRV" or PARAMETER_WORKGROUP = "SPT") Then %>
									<td data-order="<%=Replace(Replace(Replace(Replace(Replace(RSAGTLIST("PRIORITY"),"T-",""),"st",""),"nd",""),"rd",""),"th","")%>"><%=RSAGTLIST("PRIORITY")%></td>
								<% End If %>
								<td data-order="<%=RSAGTLIST("AGENT_NAME")%>">
									<% 
										SQLstmt = "SELECT * " & _
										"FROM SYS_CODE_DETAIL " & _
										"WHERE SYS_CDD_SYS_CDM_ID = 583 " & _
										"AND SYS_CDD_NAME = ? " & _
										"AND TO_DATE(?,'MM/DD/YYYY') BETWEEN TO_DATE(REGEXP_SUBSTR(SYS_CDD_VALUE,'[^_]+',1,2),'MM/DD/YYYY') + 1 AND TO_DATE('6/27/2020','MM/DD/YYYY')"
										cmd.CommandText = SQLstmt
										cmd.Parameters(0).value = RSAGTLIST("USE_AGENT")
										cmd.Parameters(1).value = PARAMETER_DATE
										Set RSRETURN = cmd.Execute
									%>
									<% If Not RSRETURN.EOF Then %>
										<i title="Returned from SLIP" class="fas fa-notes-medical <% If PARAMETER_DATE = Date Then %>today-color <% Else %>past-color <% End If %>"></i>
									<% End If %>
									<% Set RSRETURN = Nothing %>
									<% If RSAGTLIST("USE_LOCATION") <> "In-House" Then %>
										<i title="<%=RSAGTLIST("USE_LOCATION")%>" class="fas fa-home <% If PARAMETER_DATE = Date Then %>today-color <% Else %>past-color <% End If %>"></i>
									<% End If %>
									<% If RSAGTLIST("USE_CLASS") <> "Reg Full-Time" Then %>
										<i title="<%=RSAGTLIST("USE_CLASS")%>" class="<% If RSAGTLIST("USE_CLASS") = "On Leave" Then %>fas fa-bed <% Else %>far fa-clock <% End If %> <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>"></i>
									<% End If %>
									<% If PULSE_SECURITY >= 2 Then %>
										<a id="SCIUSRNAME_<%=RSAGTLIST("USE_AGENT")%>" title="<%=RSAGTLIST("SUPERVISOR_NAME")%>" style="color:#000;text-decoration:none;" href="/atlas/timecard.asp?user_agent=<%=RSAGTLIST("USE_AGENT")%>" target="_blank">
											<%=RSAGTLIST("AGENT_NAME")%>
										</a>
									<% Else %>
										<%=RSAGTLIST("AGENT_NAME")%>
									<% End If %>
								</td>
								<td class="mobile-hide"><%=RSAGTLIST("USE_WORKGROUP")%></td>
								<td class="mobile-hide">
									<% If LUNCH_ERROR_FLAG = 0 Then %>
										<%
											SQLstmt = "SELECT " & _
											"CASE " & _
												"WHEN " & _
													"TO_DATE(?,'MM/DD/YYYY') + MOD(MOD(TO_DATE('5/18/2019','MM/DD/YYYY') - TO_DATE(?,'MM/DD/YYYY'),14)+14,14) - 14 BETWEEN TO_DATE('3/18/'||TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'YYYY'),'MM/DD/YYYY') AND TO_DATE('3/31/'||TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'YYYY'),'MM/DD/YYYY') " & _
													"AND BALANCE_END_DATE < TO_DATE('4/1/'||TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'YYYY'),'MM/DD/YYYY') " & _
												"THEN '<b>Pre 4/1:</b> ' " & _
												"WHEN " & _
													"TO_DATE(?,'MM/DD/YYYY') + MOD(MOD(TO_DATE('5/18/2019','MM/DD/YYYY') - TO_DATE(?,'MM/DD/YYYY'),14)+14,14) - 14 BETWEEN TO_DATE('3/18/'||TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'YYYY'),'MM/DD/YYYY') AND TO_DATE('3/31/'||TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'YYYY'),'MM/DD/YYYY') " & _
													"AND BALANCE_END_DATE >= TO_DATE('4/1/'||TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'YYYY'),'MM/DD/YYYY') " & _
												"THEN '<b>Post 4/1:</b> ' " & _
												"ELSE NULL " & _
											"END || PTO_BALANCE PTO_BALANCE " & _
											"FROM " & _
											"( " & _
												"SELECT " & _
												"BALANCE_DATE BALANCE_START_DATE, " & _
												"NVL(LEAD(BALANCE_DATE) OVER (ORDER BY BALANCE_DATE) - 1,TO_DATE('12/31/2040','MM/DD/YYYY')) BALANCE_END_DATE, " & _
												"TRUNC(SUM(SUM(PTO_BALANCE)) OVER (ORDER BY BALANCE_DATE RANGE UNBOUNDED PRECEDING),1) PTO_BALANCE " & _
												"FROM " & _
												"( " & _
													"SELECT " & _
													"OPS_ACC_DATE BALANCE_DATE, " & _
													"SUM(TRUNC(OPS_ACC_BALANCE,1)) PTO_BALANCE " & _
													"FROM OPS_ACCRUAL " & _
													"WHERE OPS_ACC_CODE IN ('VACA','PPTV') " & _
													"AND OPS_ACC_OPS_USR_ID = ? " & _
													"GROUP BY OPS_ACC_DATE " & _
													"UNION ALL " & _
													"SELECT " & _
													"TO_DATE(OPS_SCI_START), " & _
													"-1*SUM(ROUND(24*(OPS_SCI_END-OPS_SCI_START),2)) " & _
													"FROM OPS_SCHEDULE_INFO " & _
													"WHERE REGEXP_LIKE(OPS_SCI_TYPE,'^VA|PT$|PP$') " & _
													"AND OPS_SCI_STATUS = 'APP' " & _
													"AND OPS_SCI_OPS_USR_ID = ? " & _
													"AND TO_DATE(OPS_SCI_START) >= (SELECT MAX(OPS_ACC_DATE) FROM OPS_ACCRUAL) " & _
													"GROUP BY TO_DATE(OPS_SCI_START) " & _
													"UNION ALL " & _
													"SELECT " & _
													"OPS_ACC_DATE + 12, " & _
													"OPS_ACC_BALANCE - TRUNC(OPS_ACC_BALANCE,1) " & _
													"FROM OPS_ACCRUAL " & _
													"WHERE OPS_ACC_CODE = 'PPTV' " & _
													"AND OPS_ACC_OPS_USR_ID = ? " & _
													"UNION ALL " & _
													"SELECT " & _
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
														"CONNECT BY ROWNUM < FLOOR((TO_DATE('3/31/'||TO_CHAR(ADD_MONTHS(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),12),'YYYY'),'MM/DD/YYYY') - MAX_DATE)/7) + 2 " & _
													") " & _
													"ON USE_DATE BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE " & _
													"WHERE OPS_USD_OPS_USR_ID = ? " & _
													"UNION ALL " & _
													"SELECT " & _
													"MAX((SELECT ADD_MONTHS(MAX(OPS_ACC_DATE),12) FROM OPS_ACCRUAL WHERE OPS_ACC_CODE = 'SPTO')), " & _
													"ROUND(NVL(SUM(CASE " & _
														"WHEN FLOOR(MONTHS_BETWEEN((SELECT ADD_MONTHS(MAX(OPS_ACC_DATE)-1,12) FROM OPS_ACCRUAL WHERE OPS_ACC_CODE = 'SPTO'),OPS_USR_HIRE_DATE) / 12) >= 25 OR OPS_USR_HIRE_DATE <= TO_DATE('12/31/2000','MM/DD/YYYY') THEN 200/((SELECT ADD_MONTHS(MAX(OPS_ACC_DATE)-1,12) FROM OPS_ACCRUAL WHERE OPS_ACC_CODE = 'SPTO') - (SELECT MAX(OPS_ACC_DATE) FROM OPS_ACCRUAL WHERE OPS_ACC_CODE = 'SPTO') + 1) " & _
														"WHEN FLOOR(MONTHS_BETWEEN((SELECT ADD_MONTHS(MAX(OPS_ACC_DATE)-1,12) FROM OPS_ACCRUAL WHERE OPS_ACC_CODE = 'SPTO'),OPS_USR_HIRE_DATE) / 12) >= 11 THEN 160/((SELECT ADD_MONTHS(MAX(OPS_ACC_DATE)-1,12) FROM OPS_ACCRUAL WHERE OPS_ACC_CODE = 'SPTO') - (SELECT MAX(OPS_ACC_DATE) FROM OPS_ACCRUAL WHERE OPS_ACC_CODE = 'SPTO') + 1) " & _
														"WHEN FLOOR(MONTHS_BETWEEN((SELECT ADD_MONTHS(MAX(OPS_ACC_DATE)-1,12) FROM OPS_ACCRUAL WHERE OPS_ACC_CODE = 'SPTO'),OPS_USR_HIRE_DATE) / 12) >= 5 THEN 120/((SELECT ADD_MONTHS(MAX(OPS_ACC_DATE)-1,12) FROM OPS_ACCRUAL WHERE OPS_ACC_CODE = 'SPTO') - (SELECT MAX(OPS_ACC_DATE) FROM OPS_ACCRUAL WHERE OPS_ACC_CODE = 'SPTO') + 1) " & _
														"ELSE 80/((SELECT ADD_MONTHS(MAX(OPS_ACC_DATE)-1,12) FROM OPS_ACCRUAL WHERE OPS_ACC_CODE = 'SPTO') - (SELECT MAX(OPS_ACC_DATE) FROM OPS_ACCRUAL WHERE OPS_ACC_CODE = 'SPTO') + 1) " & _
													"END * NVL(VACA_SCHEDULED_HOURS,0)/40),0),2) " & _
													"FROM " & _
													"( " & _
														"SELECT " & _
														"ACCRUAL_DATE, " & _
														"OPS_USR_HIRE_DATE, " & _
														"CASE WHEN OPS_USD_CLASS = 'LEAVE' AND LEAVE_COUNT > 30 THEN 0 ELSE OPS_USD_SCH_HOURS END VACA_SCHEDULED_HOURS " & _
														"FROM " & _
														"( " & _
															"SELECT " & _
															"ACCRUAL_DATE, " & _
															"OPS_USR_HIRE_DATE, " & _
															"OPS_USD_CLASS, " & _
															"OPS_USD_SCH_HOURS, " & _
															"COUNT(DECODE(OPS_USD_CLASS,'LEAVE','LEAVE',NULL)) OVER (ORDER BY ACCRUAL_DATE) LEAVE_COUNT " & _
															"FROM " & _
															"( " & _
																"SELECT USE_DATE + ROWNUM - 1 ACCRUAL_DATE " & _
																"FROM " & _
																"( " & _
																	"SELECT MAX(OPS_ACC_DATE) USE_DATE " & _
																	"FROM OPS_ACCRUAL " & _
																	"WHERE OPS_ACC_CODE = 'SPTO' " & _
																") " & _
																"CONNECT BY ROWNUM < ADD_MONTHS(USE_DATE-1,12) - USE_DATE + 2 " & _
															") " & _
															"JOIN " & _
															"( " & _
																"SELECT OPS_USR_HIRE_DATE, OPS_USD_EFF_DATE, OPS_USD_DIS_DATE, OPS_USD_CLASS, OPS_USD_SCH_HOURS " & _
																"FROM OPS_USER_DETAIL " & _
																"JOIN OPS_USER " & _
																"ON OPS_USD_OPS_USR_ID = OPS_USR_ID " & _
																"WHERE OPS_USD_EFF_DATE <= (SELECT ADD_MONTHS(MAX(OPS_ACC_DATE)-1,12) FROM OPS_ACCRUAL WHERE OPS_ACC_CODE = 'SPTO') " & _
																"AND OPS_USD_DIS_DATE >= (SELECT MAX(OPS_ACC_DATE) FROM OPS_ACCRUAL WHERE OPS_ACC_CODE = 'SPTO') " & _
																"AND OPS_USD_OPS_USR_ID = ? " & _
															") " & _
															"ON ACCRUAL_DATE BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE " & _
														") " & _
													") " & _
												") " & _
												"GROUP BY BALANCE_DATE " & _
											") " & _
											"WHERE " & _
											"( " & _
												"TO_DATE(?,'MM/DD/YYYY') + MOD(MOD(TO_DATE('5/18/2019','MM/DD/YYYY') - TO_DATE(?,'MM/DD/YYYY'),14)+14,14) BETWEEN BALANCE_START_DATE AND BALANCE_END_DATE " & _
												"OR " & _
												"( " & _
													"TO_DATE(?,'MM/DD/YYYY') + MOD(MOD(TO_DATE('5/18/2019','MM/DD/YYYY') - TO_DATE(?,'MM/DD/YYYY'),14)+14,14) - 14 BETWEEN TO_DATE('3/18/'||TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'YYYY'),'MM/DD/YYYY') AND TO_DATE('3/31/'||TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'YYYY'),'MM/DD/YYYY') " & _
													"AND TO_DATE('3/31/'||TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'YYYY'),'MM/DD/YYYY') BETWEEN BALANCE_START_DATE AND BALANCE_END_DATE " & _
												") " & _
											") " & _
											"ORDER BY BALANCE_END_DATE"
											cmd.CommandText = SQLstmt
											cmd.Parameters(0).value = PARAMETER_DATE
											cmd.Parameters(1).value = PARAMETER_DATE
											cmd.Parameters(2).value = PARAMETER_DATE
											cmd.Parameters(3).value = PARAMETER_DATE
											cmd.Parameters(4).value = RSAGTLIST("USE_AGENT")
											cmd.Parameters(5).value = RSAGTLIST("USE_AGENT")
											cmd.Parameters(6).value = RSAGTLIST("USE_AGENT")
											cmd.Parameters(7).value = RSAGTLIST("USE_AGENT")
											cmd.Parameters(8).value = RSAGTLIST("USE_AGENT")
											cmd.Parameters(9).value = PARAMETER_DATE
											cmd.Parameters(10).value = PARAMETER_DATE
											cmd.Parameters(11).value = PARAMETER_DATE
											cmd.Parameters(12).value = PARAMETER_DATE
											Set RSPTO = cmd.Execute
										%>
										<% Do While Not RSPTO.EOF %>
											<span style="white-space:nowrap;"><%=RSPTO("PTO_BALANCE")%></span>
											<br/>
											<% RSPTO.MoveNext %>
										<% Loop %>
										<% Set RSPTO = Nothing %>
									<% Else %>
									<% 
										SQLstmt = "SELECT " & _
										"ROUND(RES_PHN_HOURS/60) SWITCH_LUNCH " & _
										"FROM RES_PHONE_STATS " & _
										"WHERE RES_PHN_TYPE = 'LUN' " & _
										"AND RES_PHN_DATE = TO_DATE(?,'MM/DD/YYYY') " & _
										"AND RES_PHN_OPS_USR_ID = ?"
										cmd.CommandText = SQLstmt
										cmd.Parameters(0).value = PARAMETER_DATE
										cmd.Parameters(1).value = RSAGTLIST("USE_AGENT")
										Set RSLUNCH = cmd.Execute
									%>
										<% If Not RSLUNCH.EOF Then %>
											<%=RSLUNCH("SWITCH_LUNCH")%>
										<% End If %>
										<% Set RSLUNCH = Nothing %>
									<% End If %>
								</td>
								<% If REQUEST_TYPE = "PENDING" and (PARAMETER_WORKGROUP = "SLS" or PARAMETER_WORKGROUP = "SRV" or PARAMETER_WORKGROUP = "SPT") Then %>
									<td class="mobile-hide"><%=RSAGTLIST("UNP_REMAINING")%></td>
								<% End If %>
								<%
									SQLstmt = "SELECT " & _
									"TO_DATE(OPS_ATT_ACT_LOGIN) ATT_DATE, " & _
									"TO_CHAR(OPS_ATT_ACT_LOGIN,'HH24:MI') ATT_LOGIN, " & _
									"TO_CHAR(OPS_ATT_ACT_LOGOUT,'HH24:MI') ATT_LOGOUT, " & _
									"CASE " & _
										"WHEN TO_DATE(OPS_ATT_ADJ_LOGIN) <> NVL(LEAD(TO_DATE(OPS_ATT_ADJ_LOGIN)) OVER (ORDER BY OPS_ATT_ACT_LOGIN, OPS_ATT_ACT_LOGOUT),TO_DATE(OPS_ATT_ADJ_LOGIN)) THEN 2 " & _
										"WHEN OPS_ATT_ADJ_LOGOUT <> NVL(LEAD(OPS_ATT_ADJ_LOGIN) OVER (ORDER BY OPS_ATT_ACT_LOGIN, OPS_ATT_ACT_LOGOUT),OPS_ATT_ADJ_LOGOUT) THEN 1 " & _
									"ELSE 0 " & _
									"END GAP_FLAG " & _
									"FROM " & _
									"( " & _
										"SELECT " & _
										"OPS_ATT_ADJ_LOGIN, " & _
										"OPS_ATT_ADJ_LOGOUT, " & _
										"OPS_ATT_ACT_LOGIN, " & _
										"OPS_ATT_ACT_LOGOUT, " & _
										"CONNECT_BY_ROOT(OPS_ATT_ADJ_LOGIN) ROOT_START, " & _
										"MAX(OPS_ATT_ADJ_LOGOUT) OVER (PARTITION BY CONNECT_BY_ROOT(OPS_ATT_ADJ_LOGIN)) ROOT_END " & _
										"FROM " & _
										"( " & _
											"SELECT " & _
											"OPS_ATT_ADJ_LOGIN, " & _
											"OPS_ATT_ADJ_LOGOUT, " & _
											"OPS_ATT_ACT_LOGIN, " & _
											"OPS_ATT_ACT_LOGOUT, " & _
											"CASE " & _
												"WHEN LAG(OPS_ATT_ADJ_LOGIN) OVER (ORDER BY OPS_ATT_ACT_LOGIN, OPS_ATT_ACT_LOGOUT) IS NULL " & _
												"OR LAG(OPS_ATT_ADJ_LOGOUT) OVER (ORDER BY OPS_ATT_ACT_LOGIN, OPS_ATT_ACT_LOGOUT) <> OPS_ATT_ADJ_LOGIN THEN 1 " & _
											"END START_FLAG, " & _
											"ROW_NUMBER() OVER (ORDER BY OPS_ATT_ACT_LOGIN, OPS_ATT_ACT_LOGOUT) TIMEBREAK_ORDER " & _
											"FROM " & _
											"( " & _
												"SELECT " & _
												"TO_DATE(OPS_ATT_DATE || ' ' || OPS_ATT_ADJ_LOGIN,'MM/DD/YYYY HH24:MI') OPS_ATT_ADJ_LOGIN, " & _
												"TO_DATE(OPS_ATT_DATE + CASE WHEN OPS_ATT_ADJ_LOGIN > OPS_ATT_ADJ_LOGOUT THEN 1 ELSE 0 END || ' ' || OPS_ATT_ADJ_LOGOUT,'MM/DD/YYYY HH24:MI') OPS_ATT_ADJ_LOGOUT, " & _
												"TO_DATE(OPS_ATT_DATE || ' ' || OPS_ATT_ACT_LOGIN,'MM/DD/YYYY HH24:MI') OPS_ATT_ACT_LOGIN, " & _
												"TO_DATE(OPS_ATT_DATE + CASE WHEN OPS_ATT_ACT_LOGIN > OPS_ATT_ACT_LOGOUT THEN 1 ELSE 0 END || ' ' || OPS_ATT_ACT_LOGOUT,'MM/DD/YYYY HH24:MI') OPS_ATT_ACT_LOGOUT " & _
												"FROM OPS_ATTENDANCE " & _
												"WHERE OPS_ATT_DATE BETWEEN TO_DATE(?,'MM/DD/YYYY') AND TO_DATE(?,'MM/DD/YYYY') " & _
												"AND OPS_ATT_OPS_USR_ID = ? " & _
												"AND OPS_ATT_TYPE IN ('RES','SC') " & _
											") " & _
										") " & _
										"START WITH START_FLAG = 1 " & _
										"CONNECT BY OPS_ATT_ADJ_LOGIN = PRIOR OPS_ATT_ADJ_LOGOUT " & _
										"AND TIMEBREAK_ORDER = PRIOR TIMEBREAK_ORDER + 1 " & _
									") " & _
									"WHERE ROOT_START < TO_DATE(?,'MM/DD/YYYY') " & _
									"AND ROOT_END > TO_DATE(?,'MM/DD/YYYY')"
									cmd.CommandText = SQLstmt
									cmd.Parameters(0).value = PARAMETER_DATE - 1
									cmd.Parameters(1).value = PARAMETER_DATE + 1
									cmd.Parameters(2).value = RSAGTLIST("USE_AGENT")
									cmd.Parameters(3).value = PARAMETER_DATE + 1
									cmd.Parameters(4).value = PARAMETER_DATE
									Set RSLOGIN = cmd.Execute
								%>
								<td style="vertical-align:top;" <% If Not RSLOGIN.EOF Then %>data-order="<%=RSLOGIN("ATT_LOGIN")%>"<% Else %>data-order=""<% End If %>>
									<% If PARAMETER_DATE = Date Then %>
										<div id="AGENTPHONESTATE_<%=AgentPhone(RSAGTLIST("USE_AGENT"))%>" style="font-weight:bold;font-size:.8rem;"></div>
									<% End If %>
									<% Do While Not RSLOGIN.EOF %>
										<span style="white-space:nowrap;<% If CDate(RSLOGIN("ATT_DATE")) <> PARAMETER_DATE Then %>font-style:italic;font-size:.8rem;<% End If %>"><%=RSLOGIN("ATT_LOGIN")%> - <%=RSLOGIN("ATT_LOGOUT")%></span>
										<br/>
										<% If RSLOGIN("GAP_FLAG") = "1" Then %>
											<div style="line-height:10px;">&nbsp;</div>
										<% End If %>
										<% If RSLOGIN("GAP_FLAG") = "2" Then %>
											<div style="line-height:20px;">&nbsp;</div>
										<% End If %>
										<% RSLOGIN.MoveNext %>
									<% Loop %>
									<% Set RSLOGIN = Nothing %>
								</td>
								<%
									SQLstmt = "SELECT " & _
									"TO_DATE(OPS_SCI_START) SCI_DATE, " & _
									"TO_CHAR(DECODE(OPS_SCI_STATUS,'APP',OPS_SCI_START,GREATEST(NVL(YESTERDAY_END,OPS_SCI_START),OPS_SCI_START)),'HH24:MI') SCI_START, " & _
									"TO_CHAR(DECODE(OPS_SCI_STATUS,'APP',OPS_SCI_END,LEAST(NVL(TOMMOROW_START,OPS_SCI_END),OPS_SCI_END)),'HH24:MI') SCI_END, " & _
									"OPS_SCI_TYPE, " & _
									"OPS_SCI_STATUS, " & _
									"OPS_SCI_NOTES, " & _
									"CASE " & _
										"WHEN OPS_SCI_STATUS <> 'APP' THEN 'PEND' " & _
										"WHEN OPS_SCI_TYPE IN ('PICK','BASE','HOLW','ADDT','EXTD') THEN 'PHONE' " & _
										"WHEN OPS_SCI_TYPE IN ('MEET','PRES','PROJ','TRAN','FAMP','WFHU','MLTU','OTRG','NEWH') THEN 'TRAIN' " & _
										"WHEN OPS_SCI_TYPE IN ('SRPT','SRUN') THEN 'SRED' " & _
										"WHEN OPS_SCI_TYPE IN ('LNCH','LNFL') THEN 'LUNCH' " & _
										"WHEN REGEXP_LIKE(OPS_SCI_TYPE,'^VAC|UN$|PT$|PP$|HOLU|HOLR|SLIP|RESH|RCHG|ROUT|JURY|BRVT') THEN 'VACA' " & _
									"END SCHEDULE_CLASS, " & _
									"CASE " & _
										"WHEN TO_DATE(OPS_SCI_START) <> NVL(LEAD(TO_DATE(OPS_SCI_START)) OVER (ORDER BY TO_DATE(OPS_SCI_START), DECODE(OPS_SCI_STATUS,'APP',1,'SUB',2,'OPT',3), CASE WHEN RES_BUE_ID IS NOT NULL AND OPS_SCI_TYPE IN ('HOLR','HOLU') THEN 1 ELSE 2 END, OPS_SCI_START, OPS_SCI_END, DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'UN$|PP$|PT$'),0,OPS_SCI_TYPE,'AA'||OPS_SCI_TYPE)),TO_DATE(OPS_SCI_START)) THEN 2 " & _
										"WHEN (OPS_SCI_TYPE NOT IN ('HOLR','HOLU') OR RES_BUE_ID IS NULL) AND OPS_SCI_END <> NVL(LEAD(OPS_SCI_START) OVER (ORDER BY TO_DATE(OPS_SCI_START), DECODE(OPS_SCI_STATUS,'APP',1,'SUB',2,'OPT',3), CASE WHEN RES_BUE_ID IS NOT NULL AND OPS_SCI_TYPE IN ('HOLR','HOLU') THEN 1 ELSE 2 END, OPS_SCI_START, OPS_SCI_END, DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'UN$|PP$|PT$'),0,OPS_SCI_TYPE,'AA'||OPS_SCI_TYPE)),OPS_SCI_END) THEN 1 " & _
										"WHEN OPS_SCI_STATUS = 'APP' AND NVL(LEAD(OPS_SCI_STATUS) OVER (ORDER BY TO_DATE(OPS_SCI_START), DECODE(OPS_SCI_STATUS,'APP',1,'SUB',2,'OPT',3), CASE WHEN RES_BUE_ID IS NOT NULL AND OPS_SCI_TYPE IN ('HOLR','HOLU') THEN 1 ELSE 2 END, OPS_SCI_START, OPS_SCI_END, DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'UN$|PP$|PT$'),0,OPS_SCI_TYPE,'AA'||OPS_SCI_TYPE)),'APP') <> 'APP' THEN 1 " & _
										"WHEN OPS_SCI_TYPE IN ('HOLR','HOLU') AND RES_BUE_ID IS NOT NULL AND LEAD(OPS_SCI_TYPE) OVER (ORDER BY TO_DATE(OPS_SCI_START), DECODE(OPS_SCI_STATUS,'APP',1,'SUB',2,'OPT',3), CASE WHEN RES_BUE_ID IS NOT NULL AND OPS_SCI_TYPE IN ('HOLR','HOLU') THEN 1 ELSE 2 END, OPS_SCI_START, OPS_SCI_END, DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'UN$|PP$|PT$'),0,OPS_SCI_TYPE,'AA'||OPS_SCI_TYPE)) NOT IN ('HOLR','HOLU') THEN 1 " & _
										"ELSE 0 " & _
									"END GAP_FLAG " & _
									"FROM " & _
									"( " & _
										"SELECT " & _
										"OPS_SCI_START, " & _
										"OPS_SCI_END, " & _
										"OPS_SCI_TYPE, " & _
										"OPS_SCI_STATUS, " & _
										"OPS_SCI_NOTES, " & _
										"ROOT_START, " & _
										"ROOT_END, " & _
										"MIN(CASE WHEN TO_DATE(ROOT_START) < TO_DATE(?,'MM/DD/YYYY') AND TO_DATE(ROOT_END) > TO_DATE(?,'MM/DD/YYYY') AND OPS_SCI_STATUS = 'APP' THEN ROOT_START END) OVER () MIN_START, " & _
										"MAX(CASE WHEN TO_DATE(ROOT_START) < TO_DATE(?,'MM/DD/YYYY') AND TO_DATE(ROOT_END) > TO_DATE(?,'MM/DD/YYYY') AND OPS_SCI_STATUS = 'APP' THEN ROOT_END END) OVER () MAX_END, " & _
										"MAX(CASE WHEN TO_DATE(ROOT_END) = TO_DATE(?,'MM/DD/YYYY') AND OPS_SCI_STATUS = 'APP' THEN OPS_SCI_END END) OVER () YESTERDAY_END, " & _
										"MIN(CASE WHEN TO_DATE(ROOT_START) = TO_DATE(?,'MM/DD/YYYY') AND OPS_SCI_STATUS = 'APP' THEN OPS_SCI_START END) OVER () TOMMOROW_START " & _
										"FROM " & _
										"( " & _
											"SELECT " & _
											"OPS_SCI_START, " & _
											"OPS_SCI_END, " & _
											"OPS_SCI_TYPE, " & _
											"OPS_SCI_STATUS, " & _
											"OPS_SCI_NOTES, " & _
											"CONNECT_BY_ROOT(OPS_SCI_START) ROOT_START, " & _
											"MAX(OPS_SCI_END) OVER (PARTITION BY CONNECT_BY_ROOT(OPS_SCI_START)) ROOT_END " & _
											"FROM " & _
											"( " & _
												"SELECT " & _
												"OPS_SCI_START, " & _
												"OPS_SCI_END, " & _
												"OPS_SCI_TYPE, " & _
												"OPS_SCI_STATUS, " & _
												"OPS_SCI_NOTES, " & _
												"CASE " & _
													"WHEN LAG(OPS_SCI_START) OVER (ORDER BY OPS_SCI_STATUS, OPS_SCI_START) IS NULL " & _
													"OR LAG(OPS_SCI_END) OVER (ORDER BY OPS_SCI_STATUS, OPS_SCI_START) <> OPS_SCI_START " & _
													"OR LAG(OPS_SCI_STATUS) OVER (ORDER BY OPS_SCI_STATUS, OPS_SCI_START) <> OPS_SCI_STATUS THEN 1 " & _
												"END START_FLAG, " & _
												"ROW_NUMBER() OVER (ORDER BY OPS_SCI_STATUS, OPS_SCI_START, DECODE(OPS_SCI_START,OPS_SCI_END,DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'UN$|PP$|PT$'),0,OPS_SCI_TYPE,'AA'||OPS_SCI_TYPE),'ZZ')) TIMEBREAK_ORDER " & _
												"FROM OPS_SCHEDULE_INFO " & _
												"WHERE TO_DATE(OPS_SCI_START) BETWEEN TO_DATE(?,'MM/DD/YYYY') AND TO_DATE(?,'MM/DD/YYYY') " & _
												"AND OPS_SCI_STATUS IN ('APP','SUB','OPT') " & _
												"AND OPS_SCI_OPS_USR_ID = ? " & _
											") " & _
											"START WITH START_FLAG = 1 " & _
											"CONNECT BY OPS_SCI_START = PRIOR OPS_SCI_END " & _
											"AND OPS_SCI_STATUS = PRIOR OPS_SCI_STATUS " & _
											"AND TIMEBREAK_ORDER = PRIOR TIMEBREAK_ORDER + 1 " & _
										") " & _
									") " & _
									"LEFT JOIN RES_BUDGET_EXCEPTION " & _
									"ON TO_DATE(OPS_SCI_START) = RES_BUE_DATE " & _
									"AND RES_BUE_TYPE = 'NOR' " & _
									"WHERE " & _
									"( " & _
										"OPS_SCI_STATUS = 'APP' " & _
										"AND ROOT_START < TO_DATE(?,'MM/DD/YYYY') " & _
										"AND ROOT_END > TO_DATE(?,'MM/DD/YYYY') " & _
									") " & _
									"OR " & _
									"( " & _
										"OPS_SCI_STATUS <> 'APP' " & _
										"AND " & _
										"( " & _
											"TO_DATE(OPS_SCI_START) = TO_DATE(?,'MM/DD/YYYY') " & _
											"OR " & _
											"( " & _
												"OPS_SCI_START < MAX_END " & _
												"AND OPS_SCI_END > MIN_START " & _
											") " & _
										") " & _
									") " & _
									"ORDER BY SCI_DATE, DECODE(OPS_SCI_STATUS,'APP',1,'SUB',2,'OPT',3), CASE WHEN RES_BUE_ID IS NOT NULL AND OPS_SCI_TYPE IN ('HOLR','HOLU') THEN 1 ELSE 2 END, OPS_SCI_START, OPS_SCI_END, DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'UN$|PP$|PT$'),0,OPS_SCI_TYPE,'AA'||OPS_SCI_TYPE)"
									cmd.CommandText = SQLstmt
									cmd.Parameters(0).value = PARAMETER_DATE + 1
									cmd.Parameters(1).value = PARAMETER_DATE - 1
									cmd.Parameters(2).value = PARAMETER_DATE + 1
									cmd.Parameters(3).value = PARAMETER_DATE - 1
									cmd.Parameters(4).value = PARAMETER_DATE - 1
									cmd.Parameters(5).value = PARAMETER_DATE + 1
									cmd.Parameters(6).value = PARAMETER_DATE - 1
									cmd.Parameters(7).value = PARAMETER_DATE + 1
									cmd.Parameters(8).value = RSAGTLIST("USE_AGENT")
									cmd.Parameters(9).value = PARAMETER_DATE + 1
									cmd.Parameters(10).value = PARAMETER_DATE
									cmd.Parameters(11).value = PARAMETER_DATE
									Set RSSHIFT = cmd.Execute
								%>
								<td style="vertical-align:top;" <% If Not RSSHIFT.EOF Then %> data-order="<%=RSSHIFT("SCI_START")%>" <% Else %>data-order=""<% End If %>>
									<% If Not RSSHIFT.EOF Then %>
										<table style="margin:auto;">
										<% Do While Not RSSHIFT.EOF %>
											<tr class="<%=RSSHIFT("SCHEDULE_CLASS")%>" title="<%=RSSHIFT("OPS_SCI_NOTES")%>" <% If CDate(RSSHIFT("SCI_DATE")) <> PARAMETER_DATE Then %>style="font-style:italic;font-size:.8rem;"<% End If %>>
												<td class="subtable-td-padded-sm"><%=RSSHIFT("OPS_SCI_TYPE")%></td>
												<td class="subtable-td-padded-sm"><%=RSSHIFT("OPS_SCI_STATUS")%></td>
												<td class="subtable-td-padded-sm"><%=RSSHIFT("SCI_START")%></td>
												<td class="subtable-td-padded-sm"><%=RSSHIFT("SCI_END")%></td>
											</tr>
											<% If RSSHIFT("GAP_FLAG") = "1" Then %>
												<tr style="line-height:0px;">
													<td class="subtable-td-padded-sm" colspan="4">&nbsp;</td>
												</tr>
											<% End If %>
											<% If RSSHIFT("GAP_FLAG") = "2" Then %>
												<tr style="line-height:30px;">
													<td class="subtable-td-padded-sm" colspan="4">&nbsp;</td>
												</tr>
											<% End If %>
											<% RSSHIFT.MoveNext %>
										<% Loop %>
										</table>
									<% End If %>
									<% Set RSSHIFT = Nothing %>
									<% If PULSE_SECURITY >= 5 or (PULSE_SECURITY >= 3 and PARAMETER_DATE >= PULSE_PAYPERIOD_START) Then %>
										<i id="EDITBUTTON_<%=RSAGTLIST("USE_AGENT")%>_<%=Right("0" & Month(PARAMETER_DATE),2) & Right("0" & Day(PARAMETER_DATE),2) & Year(PARAMETER_DATE)%>" class="fas fa-angle-down <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %> icon-style"></i>
									<% End If %>
								</td>
								<td style="vertical-align:top;">
									<% FLEX_LENGTH = 0 %>
									<% ERROR_STRING = "" %>
									<% If REQUEST_TYPE = "ERROR" Then %>
										<% ERROR_ARRAY = Split(RSAGTLIST("ERROR_CODE"),";") %>
										<% ACTION_STATUS_ARRAY = Split(RSAGTLIST("ACTION_STATUS"),";") %>
										<% ACTION_DATA_ARRAY = Split(RSAGTLIST("ACTION_DATA"),";") %>
										<% For i = 0 to UBound(ERROR_ARRAY) %>
											<% If ERROR_ARRAY(i) = "STRT" Then %>
												<% FLEX_LENGTH = ACTION_DATA_ARRAY(i) %>
											<% End If %>
											<% If ACTION_STATUS_ARRAY(i) = "1" and (PULSE_SECURITY >= 5 or (PULSE_SECURITY >= 3 and PARAMETER_DATE >= PULSE_PAYPERIOD_START)) Then %>
												<% ERROR_STRING = ERROR_STRING & ERROR_ARRAY(i) & ";" %>
												<button id="ACKBUTTON_<%=RSAGTLIST("USE_AGENT")%>_<%=ERROR_ARRAY(i)%>" type="button" class="btn <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" style="background-color:#fff;">
													<%=ErrorDescription(ERROR_ARRAY(i))%>
												</button>
												<input type="checkbox" id="ACKERROR_<%=RSAGTLIST("USE_AGENT")%>_<%=ERROR_ARRAY(i)%>" name="ACKNOWLEDGE_ERROR" value="<%=RSAGTLIST("USE_AGENT")%>_<%=ERROR_ARRAY(i)%>" style="display:none;" />
											<% Else %>
												<span class="<% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>">
													<%=ErrorDescription(ERROR_ARRAY(i))%>
												</span>
											<% End If %>
											<br/>
											<br/>
										<% Next %>
										<input type="hidden" id="FLEXLENGTH_<%=RSAGTLIST("USE_AGENT")%>" value="<%=FLEX_LENGTH%>"/>
									<% Else %>
									<%
										If PARAMETER_DATE <= Date Then
											cmd.CommandText = SQLErrorAgt
											For i = 0 to 33
												If i = 2 or i = 7 or i = 15 or i = 26 or i = 28 or i = 33 Then
													cmd.Parameters(i).value = RSAGTLIST("USE_AGENT")
												Else
													cmd.Parameters(i).value = PARAMETER_DATE
												End If
											Next
										Else
											cmd.CommandText = SQLErrorAgtFuture
											For i = 0 to 19
												If i = 1 or i = 12 or i = 14 or i = 19 Then
													cmd.Parameters(i).value = RSAGTLIST("USE_AGENT")
												Else
													cmd.Parameters(i).value = PARAMETER_DATE
												End If
											Next
										End If
										Set RSERROR = cmd.Execute
									%>
										<% If Not RSERROR.EOF Then %>
											<% ERROR_ARRAY = Split(RSERROR("ERROR_CODE"),";") %>
											<% ACTION_STATUS_ARRAY = Split(RSERROR("ACTION_STATUS"),";") %>
											<% ACTION_DATA_ARRAY = Split(RSERROR("ACTION_DATA"),";") %>
											<% For i = 0 to UBound(ERROR_ARRAY) %>
												<% If ERROR_ARRAY(i) = "STRT" Then %>
													<% FLEX_LENGTH = ACTION_DATA_ARRAY(i) %>
												<% End If %>
												<% If ACTION_STATUS_ARRAY(i) = "1" and (PULSE_SECURITY >= 5 or (PULSE_SECURITY >= 3 and PARAMETER_DATE >= PULSE_PAYPERIOD_START)) Then %>
													<% ERROR_STRING = ERROR_STRING & ERROR_ARRAY(i) & ";" %>
													<button id="ACKBUTTON_<%=RSAGTLIST("USE_AGENT")%>_<%=ERROR_ARRAY(i)%>" type="button" class="btn <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" style="background-color:#fff;">
														<%=ErrorDescription(ERROR_ARRAY(i))%>
													</button>
													<input type="checkbox" id="ACKERROR_<%=RSAGTLIST("USE_AGENT")%>_<%=ERROR_ARRAY(i)%>" name="ACKNOWLEDGE_ERROR" value="<%=RSAGTLIST("USE_AGENT")%>_<%=ERROR_ARRAY(i)%>" style="display:none;" />
												<% Else %>
													<span class="<% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>">
														<%=ErrorDescription(ERROR_ARRAY(i))%>
													</span>
												<% End If %>
												<br/>
												<br/>
											<% Next %>
											<input type="hidden" id="FLEXLENGTH_<%=RSAGTLIST("USE_AGENT")%>" value="<%=FLEX_LENGTH%>"/>
											<% Set RSERROR = Nothing %>
										<% End If %>
									<% End If %>
									<% If RSAGTLIST("WAIVER_BOOL") = "1" Then %>
										<i title="Lunch Waived" class="fas fa-pizza-slice <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %> icon-style" style="margin-right:10px;"></i>
									<% End If %>
									<span class="searchNotes" data-user="<%=RSAGTLIST("USE_AGENT")%>" data-error="<%=ERROR_STRING%>">
										<i title="View Note(s)" class="<% If RSAGTLIST("NOTE_BOOL") = "1" Then %>fas fa-comment <% Else %>far fa-comment <% End If %> <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %> icon-style"></i>
									</span>
								</td>
							</tr>
							<% If DATATABLES_BOOL = 0 and (PULSE_SECURITY >= 5 or (PULSE_SECURITY >= 3 and PARAMETER_DATE >= PULSE_PAYPERIOD_START)) Then %>
								<tr id="EDITROW_<%=RSAGTLIST("USE_AGENT")%>_<%=Right("0" & Month(PARAMETER_DATE),2) & Right("0" & Day(PARAMETER_DATE),2) & Year(PARAMETER_DATE)%>" style="display:none;" data-flex=<%=FLEX_LENGTH%>>
									<td id="EDITDIV_WRAPPER_<%=RSAGTLIST("USE_AGENT")%>_<%=Right("0" & Month(PARAMETER_DATE),2) & Right("0" & Day(PARAMETER_DATE),2) & Year(PARAMETER_DATE)%>" colspan="<%=USE_COLSPAN%>">
										<div id="EDITDIV_<%=RSAGTLIST("USE_AGENT")%>_<%=Right("0" & Month(PARAMETER_DATE),2) & Right("0" & Day(PARAMETER_DATE),2) & Year(PARAMETER_DATE)%>"></div>
									</td>
								</tr>
							<% End If %>
							<% RSAGTLIST.MoveNext %>
						<% Loop %>
						</tbody>
						<% If PULSE_SECURITY >= 5 or (PULSE_SECURITY >= 3 and PARAMETER_DATE >= PULSE_PAYPERIOD_START) Then %>
							<tfoot>
								<tr>
									<td colspan="<%=USE_COLSPAN%>">
										<input id="PULSE_SUBMIT" type="submit" class="btn th-color <% If PARAMETER_DATE = Date Then %>today-color-background<% Else %>past-color-background<% End If %>" value="Submit Changes"/>
										<div id="OVERLAP_MESSAGE" class="error-color" style="display:none;">
											Fix overlapping schedule entries before submitting.
										</div>
									</td>
								</tr>
							</tfoot>
						<% End If %>
					<% Else %>
						<tr>
							<td colspan="<%=USE_COLSPAN%>">
								No associates found.
							</td>
						</tr>
					<% End If %>				
				</table>
			</div>
		</form>
	<% End If %>
	<script>
		$(document).ready(function() {
			overlapIdList = [];
			scheduleIdList = [];
			<% If DATATABLES_BOOL = 1 Then %>
				var useDataTable = $("#PULSE_TABLE").DataTable({
					"autoWidth": false,
					"paging": false,
					"searching": false,
					"info": false
				});
				<% If PULSE_SECURITY >= 5 or (PULSE_SECURITY >= 3 and PARAMETER_DATE >= PULSE_PAYPERIOD_START) Then %>
					useDataTable.rows().every(function(){
						var idArray = useDataTable.row(this).id().split("_");
						this.child($(
							'<tr id="EDITROW_' + idArray[1] + '_' + idArray[2] + '" style="display:none;">' +
								'<td id="EDITDIV_WRAPPER_' + idArray[1] + '_' + idArray[2] + '" colspan="<%=USE_COLSPAN%>">' +
									'<div id="EDITDIV_' + idArray[1] + '_' + idArray[2] + '"></div>' +
								'</td>' +
							'</tr>'
							)
						).show();
					});
				<% End If %>
			<% End If %>
		});
	</script>
<!--#include file="pulsefunctions.asp"-->
<% Set cmd = Nothing %>
<% Conn.Close %>
<% Set Conn = Nothing %>