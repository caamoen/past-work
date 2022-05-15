<!--#include file="pulseheader.asp"-->
<!--#include file="pending.asp"-->
<%
	If Request.Querystring("DATE") <> "" then
		PARAMETER_DATE = CDate(Request.Querystring("DATE"))
	Else
		PARAMETER_DATE = Date
	End If
%>
<% If InStr(PULSE_DEPARTMENT,"RES") <> 0 or PULSE_SECURITY >= 5 Then %>
	<div class="cddiv">
		<h6>Current Day Requests - <%=Month(PARAMETER_DATE) & "/" & Day(PARAMETER_DATE)%></h6>
		<% 
			cmd.CommandText = SQLPendingSLS
			For i = 0 to 20
				cmd.Parameters(i).value = PARAMETER_DATE
			Next
			Set RESSLSMASTER = cmd.Execute
			If Not RESSLSMASTER.EOF Then
				AGENT_COUNT = RESSLSMASTER("AGENT_COUNT")
			Else 
				AGENT_COUNT = "0"
			End If
		%>
		<table class="cdtable">
			<tr>
				<th style="width:40%">Sales Associates (<%=AGENT_COUNT%>)</th>
				<th style="width:25%">Request</th>
				<th style="width:15%">Avg +/-</th>
				<th style="width:20%">Min +/-</th>
			</tr>
		<%
			Do While Not RESSLSMASTER.EOF
				If Instr(RESSLSMASTER("PRIORITY"),"1st") = 0 Then
					Exit Do
				End If
				SQLstmt = "SELECT ROUND(AVG(OPS_SCN_STAFFED - OPS_SCN_PROJECTION - 1),2) AVG_PLUSMINUS, " & _
				"MIN(OPS_SCN_STAFFED - OPS_SCN_PROJECTION - 1) MIN_PLUSMINUS " & _
				"FROM " & _
				"( " & _
					"SELECT DISTINCT " & _
					"OPS_SCN_INTERVAL, " & _
					"OPS_SCN_PROJECTION, " & _
					"OPS_SCN_STAFFED " & _
					"FROM OPS_SCHEDULE_NEED " & _
					"JOIN OPS_SCHEDULE_INFO " & _
					"ON TO_DATE(OPS_SCN_DATE || ' ' || OPS_SCN_INTERVAL,'MM/DD/YYYY HH24:MI') BETWEEN GREATEST(OPS_SCI_START-(29/1440),CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)-(1/48)) AND OPS_SCI_END - (1/1440) " & _
					"AND OPS_SCI_OPS_USR_ID = ? " & _
					"AND OPS_SCI_STATUS = 'APP' " & _
					"AND OPS_SCI_TYPE IN ('BASE','PICK','ADDT') " & _
					"WHERE TO_DATE(OPS_SCN_DATE || ' ' || OPS_SCN_INTERVAL,'MM/DD/YYYY HH24:MI') BETWEEN GREATEST(TO_DATE(?,'MM/DD/YYYY HH:MI:SS AM')-(29/1440),CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)-(1/48)) AND TO_DATE(?,'MM/DD/YYYY HH:MI:SS AM')-(1/1440) " & _
					"AND OPS_SCN_TYPE = 'RES' " & _
				")"
				cmd.CommandText = SQLstmt
				cmd.Parameters(0).value = RESSLSMASTER("USE_AGENT")
				cmd.Parameters(1).value = RESSLSMASTER("CD_START") 
				cmd.Parameters(2).value = RESSLSMASTER("CD_END")
				Set RESSLSDETAIL = cmd.Execute
		%>
			<% If Not RESSLSDETAIL.EOF Then %>
			<tr>
				<td><span class="searchAgent" data-user="<%=RESSLSMASTER("USE_AGENT")%>"><%=RESSLSMASTER("AGENT_NAME")%></span></td>
				<td><%=FormatDateTime(RESSLSMASTER("CD_START"),4)%> - <%=FormatDateTime(RESSLSMASTER("CD_END"),4)%></td>
				<td><%=RESSLSDETAIL("AVG_PLUSMINUS")%></td>
				<td><%=RESSLSDETAIL("MIN_PLUSMINUS")%></td>
			</tr>		
			<% End If %>
			<% Set RESSLSDETAIL = Nothing %>
			<% RESSLSMASTER.MoveNext %>
		<% Loop %>
		<% Set RESSLSMASTER = Nothing %>
		</table>

		<table class="cdtable">
		<% 
			cmd.CommandText = SQLPendingSRV
			For i = 0 to 18
				cmd.Parameters(i).value = PARAMETER_DATE
			Next
			Set RESSRVMASTER = cmd.Execute
			If Not RESSRVMASTER.EOF Then
				AGENT_COUNT = RESSRVMASTER("AGENT_COUNT")
			Else 
				AGENT_COUNT = "0"
			End If
		%>
			<tr>
				<th style="width:40%">Service Associates (<%=AGENT_COUNT%>)</th>
				<th style="width:25%">&nbsp;</th>
				<th style="width:15%">&nbsp;</th>
				<th style="width:20%">&nbsp;</th>
			</tr>
		<% 
			Do While Not RESSRVMASTER.EOF
				If Instr(RESSRVMASTER("PRIORITY"),"1st") = 0 Then
					Exit Do
				End If
				SQLstmt = "SELECT ROUND(AVG(OPS_SCN_STAFFED - OPS_SCN_PROJECTION - 1),2) AVG_PLUSMINUS, " & _
				"MIN(OPS_SCN_STAFFED - OPS_SCN_PROJECTION - 1) MIN_PLUSMINUS " & _
				"FROM " & _
				"( " & _
					"SELECT DISTINCT " & _
					"OPS_SCN_INTERVAL, " & _
					"OPS_SCN_PROJECTION, " & _
					"OPS_SCN_STAFFED " & _
					"FROM OPS_SCHEDULE_NEED " & _
					"JOIN OPS_SCHEDULE_INFO " & _
					"ON TO_DATE(OPS_SCN_DATE || ' ' || OPS_SCN_INTERVAL,'MM/DD/YYYY HH24:MI') BETWEEN GREATEST(OPS_SCI_START-(29/1440),CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)-(1/48)) AND OPS_SCI_END - (1/1440) " & _
					"AND OPS_SCI_OPS_USR_ID = ? " & _
					"AND OPS_SCI_STATUS = 'APP' " & _
					"AND OPS_SCI_TYPE IN ('BASE','PICK','ADDT') " & _
					"WHERE TO_DATE(OPS_SCN_DATE || ' ' || OPS_SCN_INTERVAL,'MM/DD/YYYY HH24:MI') BETWEEN GREATEST(TO_DATE(?,'MM/DD/YYYY HH:MI:SS AM')-(29/1440),CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)-(1/48)) AND TO_DATE(?,'MM/DD/YYYY HH:MI:SS AM')-(1/1440) " & _
					"AND OPS_SCN_TYPE = 'RES' " & _
				")"
				cmd.CommandText = SQLstmt
				cmd.Parameters(0).value = RESSRVMASTER("USE_AGENT")
				cmd.Parameters(1).value = RESSRVMASTER("CD_START") 
				cmd.Parameters(2).value = RESSRVMASTER("CD_END")
				Set RESSRVDETAIL = cmd.Execute
		%>
			<% If Not RESSRVDETAIL.EOF Then %>
			<tr>
				<td><span class="searchAgent" data-user="<%=RESSRVMASTER("USE_AGENT")%>"><%=RESSRVMASTER("AGENT_NAME")%></span></td>
				<td><%=FormatDateTime(RESSRVMASTER("CD_START"),4)%> - <%=FormatDateTime(RESSRVMASTER("CD_END"),4)%></td>
				<td><%=RESSRVDETAIL("AVG_PLUSMINUS")%></td>
				<td><%=RESSRVDETAIL("MIN_PLUSMINUS")%></td>
			</tr>		
			<% End If %>
			<% Set RESSRVDETAIL = Nothing %>
			<% RESSRVMASTER.MoveNext %>
		<% Loop %>
		<% Set RESSRVMASTER = Nothing %>
		</table>
		
		<table class="cdtable">
		<% 
			cmd.CommandText = SQLPendingSPT
			For i = 0 to 19
				cmd.Parameters(i).value = PARAMETER_DATE
			Next
			Set RESSPTMASTER = cmd.Execute
			If Not RESSPTMASTER.EOF Then
				AGENT_COUNT = RESSPTMASTER("AGENT_COUNT")
			Else 
				AGENT_COUNT = "0"
			End If
		%>
			<tr>
				<th style="width:40%">Support Associates (<%=AGENT_COUNT%>)</th>
				<th style="width:25%">&nbsp;</th>
				<th style="width:15%">&nbsp;</th>
				<th style="width:20%">&nbsp;</th>
			</tr>
		<% 
			Do While Not RESSPTMASTER.EOF
				If Instr(RESSPTMASTER("PRIORITY"),"1st") = 0 Then
					Exit Do
				End If
				SQLstmt = "SELECT ROUND(AVG(OPS_SCN_STAFFED - OPS_SCN_PROJECTION - 1),2) AVG_PLUSMINUS, " & _
				"MIN(OPS_SCN_STAFFED - OPS_SCN_PROJECTION - 1) MIN_PLUSMINUS " & _
				"FROM " & _
				"( " & _
					"SELECT DISTINCT " & _
					"OPS_SCN_INTERVAL, " & _
					"OPS_SCN_PROJECTION, " & _
					"OPS_SCN_STAFFED " & _
					"FROM OPS_SCHEDULE_NEED " & _
					"JOIN OPS_SCHEDULE_INFO " & _
					"ON TO_DATE(OPS_SCN_DATE || ' ' || OPS_SCN_INTERVAL,'MM/DD/YYYY HH24:MI') BETWEEN GREATEST(OPS_SCI_START-(29/1440),CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)-(1/48)) AND OPS_SCI_END - (1/1440) " & _
					"AND OPS_SCI_OPS_USR_ID = ? " & _
					"AND OPS_SCI_STATUS = 'APP' " & _
					"AND OPS_SCI_TYPE IN ('BASE','PICK','ADDT') " & _
					"WHERE TO_DATE(OPS_SCN_DATE || ' ' || OPS_SCN_INTERVAL,'MM/DD/YYYY HH24:MI') BETWEEN GREATEST(TO_DATE(?,'MM/DD/YYYY HH:MI:SS AM')-(29/1440),CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE)-(1/48)) AND TO_DATE(?,'MM/DD/YYYY HH:MI:SS AM')-(1/1440) " & _
					"AND OPS_SCN_TYPE = 'SPT' " & _
				")"
				cmd.CommandText = SQLstmt
				cmd.Parameters(0).value = RESSPTMASTER("USE_AGENT")
				cmd.Parameters(1).value = RESSPTMASTER("CD_START") 
				cmd.Parameters(2).value = RESSPTMASTER("CD_END")
				Set RESSPTDETAIL = cmd.Execute
		%>
			<% If Not RESSPTDETAIL.EOF Then %>
			<tr>
				<td><span class="searchAgent" data-user="<%=RESSPTMASTER("USE_AGENT")%>"><%=RESSPTMASTER("AGENT_NAME")%></span></td>
				<td><%=FormatDateTime(RESSPTMASTER("CD_START"),4)%> - <%=FormatDateTime(RESSPTMASTER("CD_END"),4)%></td>
				<td><%=RESSPTDETAIL("AVG_PLUSMINUS")%></td>
				<td><%=RESSPTDETAIL("MIN_PLUSMINUS")%></td>
			</tr>		
			<% End If %>
			<% Set RESSPTDETAIL = Nothing %>
			<% RESSPTMASTER.MoveNext %>
		<% Loop %>
		<% Set RESSPTMASTER = Nothing %>
		</table>
	</div>
<% Elseif PULSE_DEPARTMENT <> "" Then %>
	<div class="cddiv">
		<h6>Pending Requests - <%=Month(PARAMETER_DATE) & "/" & Day(PARAMETER_DATE)%></h6>
	<%
		i = 0
		
		DEPT_ARRAY = Split(PULSE_DEPARTMENT,",")
		ReDim SUB_PARAMETER_ARRAY(UBound(DEPT_ARRAY)+2)
		
		SQLstmt = "SELECT " & _
		"OPS_USR_ID, " & _
		"OPS_USR_NAME, " & _
		"OPS_SCI_TYPE, " & _
		"ROUND(SUM(24*(OPS_SCI_END-OPS_SCI_START)),2) USE_HOURS " & _
		"FROM OPS_SCHEDULE_INFO " & _
		"JOIN OPS_USER_DETAIL " & _
		"ON OPS_SCI_OPS_USR_ID = OPS_USD_OPS_USR_ID " & _
		"AND TO_DATE(?,'MM/DD/YYYY') BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE " & _
		"JOIN OPS_USER " & _
		"ON OPS_SCI_OPS_USR_ID = OPS_USR_ID " & _
		"WHERE TO_DATE(OPS_SCI_START) = TO_DATE(?,'MM/DD/YYYY') " & _
		"AND OPS_SCI_END > CAST(SYSTIMESTAMP at Time zone 'US/Central' AS DATE) " & _
		"AND OPS_SCI_STATUS = 'SUB' " & _
		"AND OPS_USD_TYPE IN ("
		SUB_PARAMETER_ARRAY(i) = PARAMETER_DATE
		SUB_PARAMETER_ARRAY(i+1) = PARAMETER_DATE
		i = i + 2
		For n = 0 to UBound(DEPT_ARRAY)
			If n <> UBound(DEPT_ARRAY) Then
				SQLstmt = SQLstmt & "?,"
			Else
				SQLstmt = SQLstmt & "?) "
			End If
			SUB_PARAMETER_ARRAY(i) = DEPT_ARRAY(n)
			i = i + 1
		Next
		SQLstmt = SQLstmt & "GROUP BY OPS_USR_ID, OPS_USR_NAME, OPS_SCI_TYPE " & _
		"ORDER BY OPS_USR_NAME"
		cmd.CommandText = SQLstmt
		For n = 0 to UBound(SUB_PARAMETER_ARRAY)
			cmd.Parameters(n).value = SUB_PARAMETER_ARRAY(n)
		Next
		Set RSSUB = cmd.Execute
	%>
		<% If Not RSSUB.EOF Then %>
			<table class="cdtable">
				<tr>
					<th style="width:40%">Associate</th>
					<th style="width:25%">Request Type</th>
					<th style="width:35%">Duration</th>
				</tr>
				<% Do While Not RSSUB.EOF %>
					<tr>
						<td><span class="searchAgent" data-user="<%=RSSUB("OPS_USR_ID")%>"><%=RSSUB("OPS_USR_NAME")%></span></td>
						<td><%=RSSUB("OPS_SCI_TYPE")%></td>
						<td><%=RSSUB("USE_HOURS")%></td>
					</tr>	
					<% RSSUB.MoveNext %>
				<% Loop %>
			</table>
		<% Else %>
			No Requests Found
		<% End If %>
		<% Set RSSUB = Nothing %>
	</div>
<% End If %>
<% Set cmd = Nothing %>
<% Conn.Close %>
<% Set Conn = Nothing %>