<!--#include file="pulseheader.asp"-->
<%
	DATATABLES_BOOL = 0
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
	If Request.Querystring("AGENT") <> "" Then
		PARAMETER_AGENT = Request.Querystring("AGENT")
	Else
		PARAMETER_AGENT = "-1"
	End If
	If Request.Querystring("ERROR") <> "" Then
		PARAMETER_ERROR = Request.Querystring("ERROR")
	Else
		PARAMETER_ERROR = ""
	End If
	
	ReDim NOTE_PARAMETER_ARRAY(18)
	NOTE_PARAMETER_ARRAY(0) = PARAMETER_DATE
	NOTE_PARAMETER_ARRAY(1) = PARAMETER_DATE
	NOTE_PARAMETER_ARRAY(2) = PARAMETER_DATE
	NOTE_PARAMETER_ARRAY(3) = PARAMETER_DATE
	NOTE_PARAMETER_ARRAY(4) = PARAMETER_DATE
	NOTE_PARAMETER_ARRAY(5) = PARAMETER_DATE
	NOTE_PARAMETER_ARRAY(6) = PARAMETER_DATE
	NOTE_PARAMETER_ARRAY(7) = PARAMETER_DATE
	If PARAMETER_AGENT = "-1" Then
		NOTE_PARAMETER_ARRAY(8) = PARAMETER_WORKGROUP
		NOTE_PARAMETER_ARRAY(9) = PARAMETER_WORKGROUP
		i = 10
	Else
		NOTE_PARAMETER_ARRAY(8) = PARAMETER_AGENT
		i = 9
	End If
	
	SQLstmt = "SELECT " & _
	"RES_DLN_ID, " & _
	"OPS_USR_ID, " & _
	"ASSIGNED_USER, " & _
	"ASSIGNED_SUPERVISOR, " & _
	"ASSIGNED_WORKGROUP, " & _
	"ENTERED_USER, " & _
	"RES_DLN_TIME NOTE_DATETIME, " & _
	"TO_CHAR(RES_DLN_TIME,'YYYYMMDDHH24MI') DATETIME_ORDER, " & _
	"NOTE_CODE, " & _
	"NOTE_TEXT, " & _
	"COUNT(*) OVER () NOTE_COUNT " & _
	"FROM " & _
	"( " & _
		"SELECT " & _
		"RES_DLN_ID, " & _
		"ASSIGNED.OPS_USR_ID, " & _
		"ASSIGNED.OPS_USR_NAME ASSIGNED_USER, " & _
		"SUP.OPS_USR_NAME ASSIGNED_SUPERVISOR, " & _
		"CASE " & _
			"WHEN OPS_USD_TYPE = 'RES' OR OPS_USD_TEAM = 'SPT' THEN DECODE(RES_RTE_RES_RTG_ID,1,DECODE(OPS_USD_TEAM,'SLS','Sales Specialty','RES Sales'),4,'SPT VSS',5,'SPT ASR',10,'Elite Service',13,'SPT OSR') " & _
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
		"END ASSIGNED_WORKGROUP, " & _
		"ENTERED.OPS_USR_NAME ENTERED_USER, " & _
		"RES_DLN_TIME, " & _
		"TRIM(RES_DLN_TYPE) NOTE_CODE, " & _
		"DECODE(INSTR(RES_DLN_TEXT,'-'),0,NULL,TRIM(SUBSTR(RES_DLN_TEXT,INSTR(RES_DLN_TEXT,'-')+1))) NOTE_TEXT " & _
		"FROM RES_DAILY_STATS_NOTES " & _
		"JOIN OPS_USER ENTERED " & _
		"ON RES_DLN_OPS_USR_ID = ENTERED.OPS_USR_ID " & _
		"JOIN OPS_USER ASSIGNED " & _
		"ON TRIM(REGEXP_SUBSTR(RES_DLN_TEXT,'[^-]+',1,1)) = ASSIGNED.OPS_USR_ID " & _
		"JOIN OPS_USER_DETAIL " & _
		"ON TRIM(REGEXP_SUBSTR(RES_DLN_TEXT,'[^-]+',1,1)) = OPS_USD_OPS_USR_ID " & _
		"AND TO_DATE(?,'MM/DD/YYYY') BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE " & _
		"JOIN OPS_USER SUP " & _
		"ON OPS_USD_SUPERVISOR = SUP.OPS_USR_ID " & _
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
		"ON TRIM(REGEXP_SUBSTR(RES_DLN_TEXT,'[^-]+',1,1)) = RES_RTE_OPS_USR_ID " & _
		"WHERE " & _
		"( " & _
			"( " & _
				"RES_DLN_DATE = TO_DATE(?,'MM/DD/YYYY') " & _
				"AND TRIM(RES_DLN_TYPE) IN ('LATE','FLEX','SHFT','STRT','END','NFLX','OFLX','GAP','LMIN','LWAV','OLAP','ABS','FWP','OTH','OUT','WFH','NA') " & _
			") " & _
			"OR " & _
			"( " & _
				"RES_DLN_DATE BETWEEN TO_DATE(?,'MM/DD/YYYY') - TO_CHAR(TO_DATE(?,'MM/DD/YYYY'),'D') + 1 AND TO_DATE(?,'MM/DD/YYYY') - TO_CHAR(TO_DATE(?,'MM/DD/YYYY'),'D') + 7 " & _
				"AND TRIM(RES_DLN_TYPE) = 'BWHC' " & _
			") " & _
		") "
		If PARAMETER_AGENT = "-1" Then
			SQLstmt = SQLstmt & "AND " & _
			"( " & _
				"RES_DLN_OPS_USR_TYPE = ? " & _
				"OR 'ALL' = ? " & _
			") " 
		Else
			SQLstmt = SQLstmt & "AND TRIM(REGEXP_SUBSTR(RES_DLN_TEXT,'[^-]+',1,1)) = ? "
		End If
		If PARAMETER_AGENT = "-1" Then
			If PULSE_SECURITY < 5 Then
				SQLstmt = SQLstmt & "AND DECODE(RES_DLN_OPS_USR_TYPE,'SPT','RES',RES_DLN_OPS_USR_TYPE) IN ("
				USE_ARRAY = Split(PULSE_DEPARTMENT,",")
				For j = 0 to UBound(USE_ARRAY)
					If j <> UBound(USE_ARRAY) Then
						SQLstmt = SQLstmt & "?,"
					Else
						SQLstmt = SQLstmt & "?) "
					End If
					NOTE_PARAMETER_ARRAY(i) = USE_ARRAY(j)
					i = i + 1
				Next
			Else
				SQLstmt = SQLstmt & "UNION ALL " & _
				"SELECT RES_DLN_ID, 0, 'General Note', NULL, NULL, OPS_USR_NAME, RES_DLN_TIME, TRIM(RES_DLN_TYPE), DECODE(INSTR(RES_DLN_TEXT,'-'),0,NULL,TRIM(SUBSTR(RES_DLN_TEXT,INSTR(RES_DLN_TEXT,'-')+1))) " & _
				"FROM RES_DAILY_STATS_NOTES " & _
				"JOIN OPS_USER ENTERED " & _
				"ON RES_DLN_OPS_USR_ID = ENTERED.OPS_USR_ID " & _
				"WHERE RES_DLN_DATE = TO_DATE(?,'MM/DD/YYYY') " & _
				"AND 'ALL' = ? " & _
				"AND TRIM(RES_DLN_TYPE) IN ('ABS','FWP','OTH','OUT','WFH','NA') " & _
				"AND TRIM(REGEXP_SUBSTR(RES_DLN_TEXT,'[^-]+',1,1)) = '0' "
				NOTE_PARAMETER_ARRAY(i) = PARAMETER_DATE
				NOTE_PARAMETER_ARRAY(i+1) = PARAMETER_WORKGROUP
				i = i + 2
			End If
		End If
	SQLstmt = SQLstmt & ") " & _
	"ORDER BY DECODE(OPS_USR_ID,0,1,2), ASSIGNED_USER, RES_DLN_TIME"
	cmd.CommandText = SQLstmt
	i = i - 1
	ReDim Preserve NOTE_PARAMETER_ARRAY(i)
	For i = 0 to UBound(NOTE_PARAMETER_ARRAY)
		cmd.Parameters(i).value = NOTE_PARAMETER_ARRAY(i)
	Next
	Erase NOTE_PARAMETER_ARRAY
	Set RSNOTELIST = cmd.Execute
%>
	<form id="NOTE_FORM" <% If PARAMETER_AGENT = "-1" Then %> data-workgroup="<%=PARAMETER_WORKGROUP%>" <% End If %> action="includes/formhandler.asp" method="post">
		<% If PARAMETER_AGENT = "-1" Then %>
			<input type="hidden" id="NEWLINE_ID" value="0"/>
		<% End If %> 
		<input type="hidden" id="NOTEID_LIST" name="NOTEID_LIST" value=""/>
		<input type="hidden" name="FORM_DATE" value="<%=PARAMETER_DATE%>"/>
		<div class="table-responsive">
			<table id="NOTES_TABLE" class="table table-bordered center" style="background-color:#fff;">
				<caption class="th-color <% If PARAMETER_DATE = Date Then %>today-color-background<% Else %>past-color-background<% End If %>">
						<% If PARAMETER_AGENT = "-1" Then %>
							<%=DepartmentName(PARAMETER_WORKGROUP)%> Note List - <%=FormatDateTime(Now,3)%>
						<% Else %>
							<%=AgentName(PARAMETER_AGENT)%>'s Note List - <%=FormatDateTime(Now,3)%>
						<% End If %>
						<% 
							COUNT_TEXT = ""
							If Not RSNOTELIST.EOF Then
								If RSNOTELIST("NOTE_COUNT") = "1" Then
									COUNT_TEXT = "1 Note Found"
								Else
									DATATABLES_BOOL = 1
									COUNT_TEXT = RSNOTELIST("NOTE_COUNT") & " Notes Found"
								End If
							End If
						%>
						<div style="float:right"><%=COUNT_TEXT%></div>
				</caption>
				<thead>
					<tr class="th-color <% If PARAMETER_DATE = Date Then %>today-color-background<% Else %>past-color-background<% End If %>">
							<th style="width:10%">Associate</th>
							<th style="width:10%">Workgroup</th>
							<th style="width:15%">Note Type</th>
							<th class="mobile-hide" style="width:10%">Entered On</th>
							<th class="mobile-hide" style="width:10%">Entered By</th>
							<th style="width:35%">Note</th>							
							<th style="width:10%">Delete?</th>	
					</tr>
				</thead>
				<tbody>
				<% Do While Not RSNOTELIST.EOF %>
					<tr>
						<td <% If RSNOTELIST("OPS_USR_ID") <> "0" Then %> data-order="<%=RSNOTELIST("ASSIGNED_USER")%>" <% Else %> data-order="AAA" <% End If %>>
							<span <% If RSNOTELIST("OPS_USR_ID") <> "0" Then %> class="searchAgent" data-user="<%=RSNOTELIST("OPS_USR_ID")%>" <% Else %>style="font-weight:900;" <% End If %>><%=RSNOTELIST("ASSIGNED_USER")%></span>
						</td>
						<td><%=RSNOTELIST("ASSIGNED_WORKGROUP")%></td>
						<td>
							<%=ErrorDescription(RSNOTELIST("NOTE_CODE"))%>
						</td>
						<td class="mobile-hide" data-order="<%=RSNOTELIST("DATETIME_ORDER")%>"><%=RSNOTELIST("NOTE_DATETIME")%></td>
						<td class="mobile-hide"><%=RSNOTELIST("ENTERED_USER")%></td>
						<td>
							<textarea id="NOTETEXT_<%=RSNOTELIST("RES_DLN_ID")%>" name="NOTETEXT_<%=RSNOTELIST("RES_DLN_ID")%>" class="<% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" style="width:100%;border-radius:5px;" maxlength="247" rows="2"><%=RSNOTELIST("NOTE_TEXT")%></textarea>
						</td>
						<td>
							<button id="NOTEDELBUTTON_<%=RSNOTELIST("RES_DLN_ID")%>" type="button" class="btn <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" style="background-color:#fff;">
								<i class="fas fa-trash"></i>
							</button>
							<input type="checkbox" id="NOTEDELETE_<%=RSNOTELIST("RES_DLN_ID")%>" name="DELETE_ERROR" value="<%=RSNOTELIST("RES_DLN_ID")%>" style="display:none;" />
						</td>
					</tr>
					<% RSNOTELIST.MoveNext %>
				<% Loop %>
				</tbody>
				<tfoot>		
					<tr class="new-entry-color <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" style="display:none;">
						<td colspan="2">
							<% If PARAMETER_AGENT = "-1" Then %>
								<select id="NOTEAGT_0" name="NOTEAGT_0" class="new-entry-color <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>">
								<%
									ReDim NOTE_USER_ARRAY(8)
									NOTE_USER_ARRAY(0) = PARAMETER_DATE
									i = 1
									SQLstmt = "SELECT OPS_USR_ID VALUE, OPS_USR_NAME DESCRIPTION " & _
									"FROM OPS_USER " & _
									"JOIN OPS_USER_DETAIL " & _
									"ON OPS_USR_ID = OPS_USD_OPS_USR_ID " & _
									"AND TO_DATE(?,'MM/DD/YYYY') BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE " & _
									"WHERE OPS_USD_LOCATION IN ('MOT','WFD','WFH') " & _
									"AND OPS_USD_CLASS IN ('RGFT','RGPT','PT<30','LEAVE') " & _
									"AND OPS_USD_TYPE <> 'HRA' " & _
									"AND OPS_USD_PAY_RATE > 0 "
									If PULSE_SECURITY < 5 Then
										SQLstmt = SQLstmt & "AND OPS_USD_TYPE IN ("
										USE_ARRAY = Split(PULSE_DEPARTMENT,",")
										For j = 0 to UBound(USE_ARRAY)
											If j <> UBound(USE_ARRAY) Then
												SQLstmt = SQLstmt & "?,"
											Else
												SQLstmt = SQLstmt & "?) "
											End If
											NOTE_USER_ARRAY(i) = USE_ARRAY(j)
											i = i + 1
										Next
									End If
									SQLstmt = SQLstmt & "ORDER BY DESCRIPTION"
									cmd.CommandText = SQLstmt
									i = i - 1
									ReDim Preserve NOTE_USER_ARRAY(i)
									For i = 0 to UBound(NOTE_USER_ARRAY)
										cmd.Parameters(i).value = NOTE_USER_ARRAY(i)
									Next
									Erase NOTE_USER_ARRAY
									Set RSSELECT = cmd.Execute(SQLstmt)
								%>
									<option selected="selected" value="-1">Not Selected</option>
									<% If PULSE_SECURITY >= 5 Then %>
										<option value="0" style="font-weight:900;">General Note</option>
									<% End If %>
									<% Do While Not RSSELECT.EOF %>
										<option value="<%=RSSELECT("VALUE")%>"><%=RSSELECT("DESCRIPTION")%></option>
										<% RSSELECT.MoveNext %>
									<% Loop %>
									<% Set RSSELECT = Nothing %>
								</select>
							<% Else %>
								<%=AgentName(PARAMETER_AGENT)%>
								<input type="hidden" id="NOTEAGT_0" name="NOTEAGT_0" value="<%=PARAMETER_AGENT%>" />
							<% End If %>
						</td>
						<td>
							<select id="NOTECODE_0" name="NOTECODE_0" class="new-entry-color <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>">
								<option selected="selected" value="-1">Not Selected</option>
								<option value="ABS">Absence</option>
								<% If Instr(PARAMETER_ERROR,"END") <> 0 and (PULSE_SECURITY >= 5 or (PULSE_SECURITY >= 3 and PARAMETER_DATE >= PULSE_PAYPERIOD_START)) Then %>
									<option value="END">End</option>
								<% End If %>
								<% If Instr(PARAMETER_ERROR,"FLEX") <> 0 and (PULSE_SECURITY >= 5 or (PULSE_SECURITY >= 3 and PARAMETER_DATE >= PULSE_PAYPERIOD_START)) Then %>
									<option value="FLEX">Flex</option>
								<% End If %>
								<option value="FWP">Floor Walker Program</option>
								<% If Instr(PARAMETER_ERROR,"GAP") <> 0 and (PULSE_SECURITY >= 5 or (PULSE_SECURITY >= 3 and PARAMETER_DATE >= PULSE_PAYPERIOD_START)) Then %>
									<option value="GAP">Gap</option>
								<% End If %>
								<% If Instr(PARAMETER_ERROR,"LATE") <> 0 and (PULSE_SECURITY >= 5 or (PULSE_SECURITY >= 3 and PARAMETER_DATE >= PULSE_PAYPERIOD_START)) Then %>
									<option value="LATE">Late</option>
								<% End If %>
								<% If Instr(PARAMETER_ERROR,"LMIN") <> 0 and (PULSE_SECURITY >= 5 or (PULSE_SECURITY >= 3 and PARAMETER_DATE >= PULSE_PAYPERIOD_START)) Then %>
									<option value="LMIN">Lunch Minutes</option>
								<% End If %>
								<% If Instr(PARAMETER_ERROR,"LWAV") <> 0 and (PULSE_SECURITY >= 5 or (PULSE_SECURITY >= 3 and PARAMETER_DATE >= PULSE_PAYPERIOD_START)) Then %>
									<option value="LWAV">Lunch Waiver</option>
								<% End If %>
								<% If Instr(PARAMETER_ERROR,"OFLX") <> 0 and (PULSE_SECURITY >= 5 or (PULSE_SECURITY >= 3 and PARAMETER_DATE >= PULSE_PAYPERIOD_START)) Then %>
									<option value="OFLX">Opener Flex</option>
								<% End If %>
								<option value="OTH">Other</option>
								<option value="OUT">Outage</option>
								<% If Instr(PARAMETER_ERROR,"NFLX") <> 0 and (PULSE_SECURITY >= 5 or (PULSE_SECURITY >= 3 and PARAMETER_DATE >= PULSE_PAYPERIOD_START)) Then %>
									<option value="NFLX">NEWH Flex</option>
								<% End If %>
								<% If Instr(PARAMETER_ERROR,"SHFT") <> 0 and (PULSE_SECURITY >= 5 or (PULSE_SECURITY >= 3 and PARAMETER_DATE >= PULSE_PAYPERIOD_START)) Then %>
									<option value="SHFT">No Shift</option>
								<% End If %>
								<% If Instr(PARAMETER_ERROR,"STRT") <> 0 and (PULSE_SECURITY >= 5 or (PULSE_SECURITY >= 3 and PARAMETER_DATE >= PULSE_PAYPERIOD_START)) Then %>
									<option value="STRT">Start</option>
								<% End If %>
								<% If Instr(PARAMETER_ERROR,"BWHC") <> 0 and (PULSE_SECURITY >= 5 or (PULSE_SECURITY >= 3 and PARAMETER_DATE >= PULSE_PAYPERIOD_START)) Then %>
									<option value="BWHC">Weekly Hours</option>
								<% End If %>
								<option value="WFH">WFH Technical Issue</option>
							</select>
						</td>
						<td colspan="4">
							<textarea id="NOTETEXT_0" name="NOTETEXT_0" class="<% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" style="width:100%;border-radius:5px;" maxlength="247" rows="2"></textarea>
						</td>
					</tr>	
					<tr class="<% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>">
						<td colspan="7"><i id="NEWNOTE" class="fas fa-plus-square icon-style-large"></i></td>
					</tr>
					<tr>
						<td colspan="7">
							<input id="NOTE_SUBMIT" type="submit" class="btn th-color <% If PARAMETER_DATE = Date Then %>today-color-background<% Else %>past-color-background<% End If %>" value="Submit Changes"/>
						</td>
					</tr>
				</tfoot>	
			</table>
		</div>
	</form>
	<script>
		$(document).ready(function() {
			noteIdList = [];
			<% If DATATABLES_BOOL = 1 Then %>
				$("#NOTES_TABLE").DataTable({
					"autoWidth": false,
					"paging": false,
					"searching": false,
					"info": false,
					"columnDefs": [
						{"orderable": false, "targets": 6}
					]
				});
			<% End If %>
		});
	</script>
<!--#include file="pulsefunctions.asp"-->
<% Set cmd = Nothing %>
<% Conn.Close %>
<% Set Conn = Nothing %>