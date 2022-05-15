<!--#include file="pulseheader.asp"-->
<% Set FSO = Server.CreateObject("Scripting.FileSystemObject") %>
<% 
	If Request.Querystring("DATE") <> "" Then
		PARAMETER_DATE = CDate(Request.Querystring("DATE"))
	Else
		PARAMETER_DATE = Date
	End If
%>
<% If InStr(PULSE_DEPARTMENT,"RES") <> 0 Then %>
<%
	If Request.Querystring("TIME_SELECTION") <> "" then
		TIME_SELECTION = Request.Querystring("TIME_SELECTION")
	Else
		TIME_SELECTION = "AVAIL"
	End If
	If Request.Querystring("DISPLAY_ARRAY") <> "" then
		DISPLAY_ARRAY = Split(Request.Querystring("DISPLAY_ARRAY"),",")
	Else
		DISPLAY_ARRAY = Array("none","none","none","none")
	End If

	USE_OD = Replace(PULSE_DEPARTMENT,"RES","")
	USE_OD = Replace(USE_OD,",,",",")
	If Right(USE_OD,1) = "," Then
		USE_OD = Left(USE_OD,Len(USE_OD)-1)
	End If
	
	If USE_OD <> "" or PULSE_SECURITY >= 5 Then
		LOOP_COUNT = 3
		If Instr(USE_OD,",") = 0 and PULSE_SECURITY < 5 Then
			OD_NAME = USE_OD
		Else
			OD_NAME = "OD"
		End If 
	Else
		LOOP_COUNT = 2
	End If
%>
	<div class="livediv" style="margin-top:-3px;">
		<h6 style="font-weight:900;margin-bottom:0px;">
			Associates in 
			<select id="AGENT_TIME_SELECTION" class="<% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" style="font-weight:900;">
				<option value="ACDIN"<% If TIME_SELECTION = "ACDIN" Then %> selected="selected" <% End If %>>ACD</option>
				<option value="ACW"<% If TIME_SELECTION = "ACW" Then %> selected="selected" <% End If %>>ACW</option>
				<option value="AVAIL"<% If TIME_SELECTION = "AVAIL" Then %> selected="selected" <% End If %>>AVAIL</option>
				<option value="BREAK"<% If TIME_SELECTION = "BREAK" Then %> selected="selected" <% End If %>>BREAK</option>
				<option value="CALLBACKS"<% If TIME_SELECTION = "CALLBACKS" Then %> selected="selected" <% End If %>>CALLBACKS</option>
				<option value="CHAT QA"<% If TIME_SELECTION = "CHAT QA" Then %> selected="selected" <% End If %>>CHAT</option>
				<option value="DEFAULT"<% If TIME_SELECTION = "DEFAULT" Then %> selected="selected" <% End If %>>DEFAULT</option>
				<option value="LUNCH"<% If TIME_SELECTION = "LUNCH" Then %> selected="selected" <% End If %>>LUNCH</option>
				<option value="MEETING"<% If TIME_SELECTION = "MEETING" Then %> selected="selected" <% End If %>>MEETING</option>
				<option value="OTHER"<% If TIME_SELECTION = "OTHER" Then %> selected="selected" <% End If %>>OTHER</option>
				<option value="OUTBOUND"<% If TIME_SELECTION = "OUTBOUND" Then %> selected="selected" <% End If %>>OUTBOUND</option>
				<option value="PRESENTATION"<% If TIME_SELECTION = "PRESENTATION" Then %> selected="selected" <% End If %>>PRESENTATION</option>
				<option value="PROJECTS"<% If TIME_SELECTION = "PROJECTS" Then %> selected="selected" <% End If %>>PROJECT</option>
				<option value="TECH HELP"<% If TIME_SELECTION = "TECH HELP" Then %> selected="selected" <% End If %>>TECH HELP</option>
				<option value="TRAINING"<% If TIME_SELECTION = "TRAINING" Then %> selected="selected" <% End If %>>TRAINING</option>
			</select>
			<i id="CMSREFRESH" class="fas fa-sync-alt icon-style-small <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" title="Refresh" style="display:none;"></i>
		</h6>
		<input type="hidden" id="SHOWLIVESTATUS_SALES" value="<%=DISPLAY_ARRAY(0)%>" />
		<input type="hidden" id="SHOWLIVESTATUS_SERVICE" value="<%=DISPLAY_ARRAY(1)%>" />
		<input type="hidden" id="SHOWLIVESTATUS_SUPPORT" value="<%=DISPLAY_ARRAY(2)%>" />
		<input type="hidden" id="SHOWLIVESTATUS_OD" value="<%=DISPLAY_ARRAY(3)%>" />

		<% If FSO.FileExists("e:\inetpub\wwwroot\web\rawdata\liveagents.txt") Then %>
			<% Set AgentFile = FSO.getFile("e:\inetpub\wwwroot\web\rawdata\liveagents.txt") %>
			<% If AgentFile.Size > 0 Then %>
				<div style="font-weight:400;font-size:.75rem;margin-top:-1px;">
				<% If DateDiff("s",AgentFile.DateLastModified,Now) < 30 Then %>
					(<%=DateDiff("s",AgentFile.DateLastModified,Now)%>s ago)
				<% Elseif DateDiff("s",AgentFile.DateLastModified,Now) < 60 Then %>
					<span id="CMS_WARNING" style="color:red;">(<%=DateDiff("s",AgentFile.DateLastModified,Now)%>s ago)</span>
				<% Else %>
					<span id="CMS_WARNING" style="color:red;">(<%=Int(DateDiff("s",AgentFile.DateLastModified,Now)/60)%>m ago)</span>
				<% End If %>
				</div>
			<%
				Set AgentFileTextStream = FSO.OpenTextFile("e:\inetpub\wwwroot\web\rawdata\liveagents.txt",1,False,-2)
				AgentFileContents = AgentFileTextStream.ReadAll
				AgentFileLines = Split(AgentFileContents,vbcrlf)
				AgtSQLstmt = ""
				i = 0
				ReDim CMS_PARAMETER_ARRAY(200)
				For n = 2 to UBound(AgentFileLines)-1
					AgentFileItems = Split(Trim(AgentFileLines(n)),";")
					CURRENT_STATE = PulsePhoneStatus(UCase(AgentFileItems(5)))
					If CURRENT_STATE = "" Then
						CURRENT_STATE = PulsePhoneStatus(UCase(AgentFileItems(4)))
					End If
			%>
					<input type="hidden" id="LIVECMSSTATE_<%=AgentFileItems(2)%>" value="<%=CURRENT_STATE%>" />
			<%
					If Instr(TIME_SELECTION,CURRENT_STATE) > 0 Then
						AgtSQLstmt = AgtSQLstmt & "SELECT " & _
						"? PHONE_ID, " & _
						"TO_NUMBER(?) AGENT_SECONDS " & _
						"FROM DUAL UNION ALL "
						CMS_PARAMETER_ARRAY(i) = AgentFileItems(2)
						CMS_PARAMETER_ARRAY(i+1) = AgentFileItems(7)
						i = i + 2
					End If
				Next
				If AgtSQLstmt <> "" Then 
					For j = 0 to LOOP_COUNT
						k = i
						If j = 0 Then
							TEAM_NAME = "Sales"
							USE_DEPT = "RES"
							USE_TEAM = "RES,SLS"
						Elseif j = 1 Then
							TEAM_NAME = "Service"
							USE_DEPT = "RES"
							USE_TEAM = "SES,NEW"					
						Elseif j = 2 Then
							TEAM_NAME = "Support"
							USE_DEPT = "RES"
							USE_TEAM = "OSR,SRV,SPT"
						Else
							TEAM_NAME = OD_NAME
							IF PULSE_SECURITY >= 5 Then
								USE_DEPT = ""
							Else
								USE_DEPT = USE_OD
							End If
							USE_TEAM = ""
						End If
						DEPT_ARRAY = Split(USE_DEPT,",")
						TEAM_ARRAY = Split(USE_TEAM,",")
						ReDim Preserve CMS_PARAMETER_ARRAY(i + UBound(DEPT_ARRAY) + UBound(TEAM_ARRAY) + 1)
						
						SQLstmt = "SELECT OPS_USR_ID, " & _
						"OPS_USR_NAME AGENT_NAME, " & _
						"SUBSTR('0' || FLOOR(AGENT_SECONDS/60),LEAST(LENGTH(FLOOR(AGENT_SECONDS/60)),2)) || ':' || SUBSTR('0' || MOD(AGENT_SECONDS,60),-2) AGENT_TIME, " & _
						"COUNT(*) OVER () TEAM_COUNT " & _
						"FROM " & _
						"( " & _
							Left(AgtSQLstmt,Len(AgtSQLstmt)-10) & _
						") " & _
						"JOIN OPS_USER " & _
						"ON PHONE_ID = OPS_USR_PHN_ID " & _
						"AND TO_DATE(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE)) BETWEEN OPS_USR_EFF_DATE AND OPS_USR_DIS_DATE "
						If USE_DEPT = "" Then
							SQLstmt = SQLstmt & "WHERE OPS_USR_TYPE <> 'RES' "
						Else
							SQLstmt = SQLstmt & "AND OPS_USR_TYPE IN ("
							For n = 0 to UBound(DEPT_ARRAY)
								If n <> UBound(DEPT_ARRAY) Then
									SQLstmt = SQLstmt & "?,"
								Else
									SQLstmt = SQLstmt & "?) "
								End If
								CMS_PARAMETER_ARRAY(k) = DEPT_ARRAY(n)
								k = k + 1
							Next
						End If
						If USE_TEAM <> "" Then
							SQLstmt = SQLstmt & "AND OPS_USR_TEAM IN ("
							For n = 0 to UBound(TEAM_ARRAY)
								If n <> UBound(TEAM_ARRAY) Then
									SQLstmt = SQLstmt & "?,"
								Else
									SQLstmt = SQLstmt & "?) "
								End If
								CMS_PARAMETER_ARRAY(k) = TEAM_ARRAY(n)
								k = k + 1
							Next
						End If
						SQLstmt = SQLstmt & "AND " & _
						"( " & _
							"OPS_USR_JOB IN ('AGT','GSM','GSP','SPC') " & _
							"OR " & _
							"( " & _
								"OPS_USR_TYPE = 'OPS' " & _
								"AND OPS_USR_JOB = 'ANL' " & _
							") " & _
						") " & _
						"ORDER BY AGENT_SECONDS DESC, AGENT_NAME"
						cmd.CommandText = SQLstmt
						For n = 0 to UBound(CMS_PARAMETER_ARRAY)
							cmd.Parameters(n).value = CMS_PARAMETER_ARRAY(n)
						Next
						Set RSLIVEAGENTS = cmd.Execute
					%>	
						<table class="livetable">
							<% If Not RSLIVEAGENTS.EOF Then %>
								<tr>
									<th style="width:70%"><%=TEAM_NAME%> Associates (<%=RSLIVEAGENTS("TEAM_COUNT")%>)</th>
									<th style="width:10%"><i id="SHOWLIVEBUTTON_<% If j = 3 Then %>OD<% Else %><%=UCase(TEAM_NAME)%><% End If %>" class="fas fa-<%=Replace(Replace(DISPLAY_ARRAY(j),"none","plus"),"table-row-group","minus")%> icon-style <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" style="font-size:1em;"></i></th>
									<th style="width:20%"><%=RSLIVEAGENTS("AGENT_TIME")%></th>
								</tr>
							<% Else %>
								<tr>
									<th style="width:70%"><%=TEAM_NAME%> Associates (0)</th>
									<th style="width:10%">&nbsp;</th>
									<th style="width:20%">&nbsp;</th>
								</tr>							
							<% End If %>
							<tbody id="SHOWLIVEBODY_<% If j = 3 Then %>OD<% Else %><%=UCase(TEAM_NAME)%><% End If %>" style="display:<%=DISPLAY_ARRAY(j)%>;">
								<% Do While Not RSLIVEAGENTS.EOF %>
									<tr>
										<td>
											<span class="searchAgent" data-user="<%=RSLIVEAGENTS("OPS_USR_ID")%>"><%=RSLIVEAGENTS("AGENT_NAME")%></span>
										</td>
										<td>&nbsp;</td>
										<td><%=RSLIVEAGENTS("AGENT_TIME")%></td>
									</tr>
									<% RSLIVEAGENTS.MoveNext %>
								<% Loop %>
							</tbody>
							<% Set RSLIVEAGENTS = Nothing %>
						</table>
					<% Next %>
				<% Else %>
					<table class="livetable">
						<tr>
							<th style="width:70%">Sales Associates (0)</th>
							<th style="width:10%">&nbsp;</th>
							<th style="width:20%">&nbsp;</th>
						</tr>							
					</table>
					<table class="livetable">
						<tr>
							<th style="width:70%">Service Associates (0)</th>
							<th style="width:10%">&nbsp;</th>
							<th style="width:20%">&nbsp;</th>
						</tr>							
					</table>
					<table class="livetable">
						<tr>
							<th style="width:70%">Support Associates (0)</th>
							<th style="width:10%">&nbsp;</th>
							<th style="width:20%">&nbsp;</th>
						</tr>							
					</table>
					<% If LOOP_COUNT = 3 Then %>
						<table class="livetable">
							<tr>
								<th style="width:70%"><%=OD_NAME%> Associates (0)</th>
								<th style="width:10%">&nbsp;</th>
								<th style="width:20%">&nbsp;</th>
							</tr>							
						</table>
					<% End If %>
				<% End If %>
				<% AgentFileTextStream.Close %>
				<% Set AgentFileTextStream = Nothing %>
			<% Else %>
				File Not Found
			<% End If %>
			<% Set AgentFile = Nothing %>
		<% Else %>
			File Not Found
		<% End If %>
	</div>
<% Else %>
	<div class="livediv">
		<h6 style="font-weight:900;margin-bottom:0px;">
			Associate State
			<i id="CMSREFRESH" class="fas fa-sync-alt icon-style-small <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" title="Refresh" style="display:none;margin-left:10px;"></i>
		</h6>
		<% If FSO.FileExists("e:\inetpub\wwwroot\web\rawdata\liveagents.txt") Then %>
			<% Set AgentFile = FSO.getFile("e:\inetpub\wwwroot\web\rawdata\liveagents.txt") %>
			<% If AgentFile.Size > 0 Then %>
				<div style="font-weight:400;font-size:.75rem;margin-top:-1px;">
				<% If DateDiff("s",AgentFile.DateLastModified,Now) < 30 Then %>
					(<%=DateDiff("s",AgentFile.DateLastModified,Now)%>s ago)
				<% Elseif DateDiff("s",AgentFile.DateLastModified,Now) < 60 Then %>
					<span id="CMS_WARNING" style="color:red;">(<%=DateDiff("s",AgentFile.DateLastModified,Now)%>s ago)</span>
				<% Else %>
					<span id="CMS_WARNING" style="color:red;">(<%=Int(DateDiff("s",AgentFile.DateLastModified,Now)/60)%>m ago)</span>
				<% End If %>
				</div>
			<%
				Set AgentFileTextStream = FSO.OpenTextFile("e:\inetpub\wwwroot\web\rawdata\liveagents.txt",1,False,-2)
				AgentFileContents = AgentFileTextStream.ReadAll
				AgentFileLines = Split(AgentFileContents,vbcrlf)
				AgtSQLstmt = ""
				i = 0
				ReDim CMS_PARAMETER_ARRAY(1000)
				For n = 2 to UBound(AgentFileLines)-1
					AgentFileItems = Split(Trim(AgentFileLines(n)),";")
					CURRENT_STATE = PulsePhoneStatus(UCase(AgentFileItems(5)))
					If CURRENT_STATE = "" Then
						CURRENT_STATE = PulsePhoneStatus(UCase(AgentFileItems(4)))
					End If
			%>
					<input type="hidden" id="LIVECMSSTATE_<%=AgentFileItems(2)%>" value="<%=CURRENT_STATE%>" />
			<%
					AgtSQLstmt = AgtSQLstmt & "SELECT " & _
					"? PHONE_ID, " & _
					"? AGENT_STATE, " & _
					"TO_NUMBER(?) AGENT_SECONDS " & _
					"FROM DUAL UNION ALL "
					CMS_PARAMETER_ARRAY(i) = AgentFileItems(2)
					CMS_PARAMETER_ARRAY(i+1) = CURRENT_STATE
					CMS_PARAMETER_ARRAY(i+2) = AgentFileItems(7)
					i = i + 3
				Next
				If AgtSQLstmt <> "" Then 			
					SQLstmt = "SELECT OPS_USR_ID, " & _
					"OPS_USR_NAME AGENT_NAME, " & _
					"AGENT_STATE, " & _
					"SUBSTR('0' || FLOOR(AGENT_SECONDS/60),LEAST(LENGTH(FLOOR(AGENT_SECONDS/60)),2)) || ':' || SUBSTR('0' || MOD(AGENT_SECONDS,60),-2) AGENT_TIME, " & _
					"COUNT(*) OVER () TEAM_COUNT " & _
					"FROM " & _
					"( " & _
						Left(AgtSQLstmt,Len(AgtSQLstmt)-10) & _
					") " & _
					"JOIN OPS_USER " & _
					"ON PHONE_ID = OPS_USR_PHN_ID " & _
					"AND TO_DATE(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE)) BETWEEN OPS_USR_EFF_DATE AND OPS_USR_DIS_DATE " & _
					"WHERE DECODE(OPS_USR_TEAM,'SPT','RES',OPS_USR_TYPE) IN ("
					USE_ARRAY = Split(PULSE_DEPARTMENT,",")
					For n = 0 to UBound(USE_ARRAY)
						If n <> UBound(USE_ARRAY) Then
							SQLstmt = SQLstmt & "?,"
						Else
							SQLstmt = SQLstmt & "?) "
						End If
						CMS_PARAMETER_ARRAY(i) = USE_ARRAY(n)
						i = i + 1
					Next
					i = i - 1
					ReDim Preserve CMS_PARAMETER_ARRAY(i)
					SQLstmt = SQLstmt & "AND " & _
					"( " & _
						"OPS_USR_JOB IN ('AGT','GSM','GSP','SPC') " & _
						"OR " & _
						"( " & _
							"OPS_USR_TYPE = 'OPS' " & _
							"AND OPS_USR_JOB = 'ANL' " & _
						") " & _
					") " & _
					"ORDER BY AGENT_NAME"
					cmd.CommandText = SQLstmt
					For n = 0 to UBound(CMS_PARAMETER_ARRAY)
						cmd.Parameters(n).value = CMS_PARAMETER_ARRAY(n)
					Next
					Set RSLIVEAGENTS = cmd.Execute
			%>
					<% If Not RSLIVEAGENTS.EOF Then %>
						<table class="livetable">
							<tr>
								<th style="width:60%">Associates (<%=RSLIVEAGENTS("TEAM_COUNT")%>)</th>
								<th style="width:25%">&nbsp;</th>
								<th style="width:15%">&nbsp;</th>
							</tr>
							<% Do While Not RSLIVEAGENTS.EOF %>
								<tr>
									<td>
										<span class="searchAgent" data-user="<%=RSLIVEAGENTS("OPS_USR_ID")%>"><%=RSLIVEAGENTS("AGENT_NAME")%></span>
									</td>
									<td><%=RSLIVEAGENTS("AGENT_STATE")%></td>
									<td><%=RSLIVEAGENTS("AGENT_TIME")%></td>
								</tr>
								<% RSLIVEAGENTS.MoveNext %>
							<% Loop %>
						</table>
					<% Else %>
						No Associates Found
					<% End If %>
					<% Set RSLIVEAGENTS = Nothing %>
					<% AgentFileTextStream.Close %>
					<% Set AgentFileTextStream = Nothing %>
				<% End If %>
			<% Else %>
				File Not Found
			<% End If %>
			<% Set AgentFile = Nothing %>
		<% Else %>
			File Not Found
		<% End If %>
	</div>
<% End If %>
<div class="livediv">
	<h6 style="font-weight:900;margin-bottom:0px;">Calls Holding</h6>
	<% If FSO.FileExists("e:\inetpub\wwwroot\web\rawdata\livecalls.txt") Then %>
		<% Set CallFile = FSO.getFile("e:\inetpub\wwwroot\web\rawdata\livecalls.txt")%>
		<% If CallFile.Size > 0 Then %>
			<div style="font-weight:400;font-size:.75rem;margin-top:-1px;">
			<% If DateDiff("s",CallFile.DateLastModified,Now) < 30 Then %>
				(<%=DateDiff("s",CallFile.DateLastModified,Now)%>s ago)
			<% Elseif DateDiff("s",CallFile.DateLastModified,Now) < 60 Then %>
				<span style="color:red;">(<%=DateDiff("s",CallFile.DateLastModified,Now)%>s ago)</span>
			<% Else %>
				<span style="color:red;">(<%=Int(DateDiff("s",CallFile.DateLastModified,Now)/60)%>m ago)</span>
			<% End If %>
			</div>
		<%
			Set CallFileTextStream = FSO.OpenTextFile("e:\inetpub\wwwroot\web\rawdata\livecalls.txt",1,False,-2)
			CallFileContents = CallFileTextStream.ReadAll
			CallFileLines = Split(CallFileContents,vbcrlf)
			SQLstmt = ""
			i = 0
			ReDim CMS_PARAMETER_ARRAY(200)
			For n = 1 to UBound(CallFileLines)-1
				CallFileItems = Split(Trim(CallFileLines(n)),";")
				If CDbl(CallFileItems(3)) > 0 and CDbl(CallFileItems(4)) > 0 Then
					SQLstmt = SQLstmt & "SELECT " & _
					"? CALL_NAME, " & _
					"TO_NUMBER(?) CALLS_HOLDING, " & _
					"TO_NUMBER(?) HOLDING_SECONDS " & _
					"FROM DUAL UNION ALL "
					CMS_PARAMETER_ARRAY(i) = CallFileItems(0)
					CMS_PARAMETER_ARRAY(i+1) = CallFileItems(3)
					CMS_PARAMETER_ARRAY(i+2) = CallFileItems(4)
					i = i + 3
				End If
			Next
			If SQLstmt <> "" Then 
				SQLstmt = "SELECT " & _
				"CALL_NAME, " & _
				"CALLS_HOLDING, " & _
				"SUBSTR('0' || FLOOR(HOLDING_SECONDS/60),LEAST(LENGTH(FLOOR(HOLDING_SECONDS/60)),2)) || ':' || SUBSTR('0' || MOD(HOLDING_SECONDS,60),-2) HOLDING_TIME, " & _
				"CASE " & _
					"WHEN HOLDING_SECONDS > 360 THEN '#faa' " & _
					"WHEN HOLDING_SECONDS > 240 THEN '#ffa' " & _
					"ELSE '#fff' " & _
				"END USE_CLASS, " & _
				"SUM(CALLS_HOLDING) OVER () TOTAL_CALLS " & _
				"FROM " & _
				"( " & _
					Left(SQLstmt,Len(SQLstmt)-10) & _
				") " & _
				"ORDER BY HOLDING_SECONDS DESC, CALLS_HOLDING DESC, CALL_NAME"
				cmd.CommandText = SQLstmt
				i = i - 1
				ReDim Preserve CMS_PARAMETER_ARRAY(i)
				For i = 0 to UBound(CMS_PARAMETER_ARRAY)
					cmd.Parameters(i).value = CMS_PARAMETER_ARRAY(i)
				Next
				Erase CMS_PARAMETER_ARRAY
				Set RSLIVECALLS = cmd.Execute
			%>
				<table class="livetable">
					<tr>
						<th style="width:70%">Total Calls</th>
						<th style="width:10%"><%=RSLIVECALLS("TOTAL_CALLS")%></th>
						<th style="width:20%"><%=RSLIVECALLS("HOLDING_TIME")%></th>
					</tr>
					<% Do While Not RSLIVECALLS.EOF %>
						<tr style="background-color:<%=RSLIVECALLS("USE_CLASS")%>;">
							<td><%=RSLIVECALLS("CALL_NAME")%></td>
							<td><%=RSLIVECALLS("CALLS_HOLDING")%></td>
							<td><%=RSLIVECALLS("HOLDING_TIME")%></td>
						</tr>
						<% RSLIVECALLS.MoveNext %>
					<% Loop %>
					<% Set RSLIVECALLS = Nothing %>
				</table>
			<% Else %>
				No Calls Holding
			<% End If %>
			<% CallFileTextStream.Close %>
			<% Set CallFileTextStream = Nothing %>
		<% Else %>
			File Not Found
		<% End If %>
		<% Set CallFile = Nothing %>
	<% Else %>
		File Not Found
	<% End If %>
</div>
<!--#include file="pulsefunctions.asp"-->
<% Set FSO = Nothing  %>
<% Set cmd = Nothing %>
<% Conn.Close %>
<% Set Conn = Nothing %>