<!--#include file="pulseheader.asp"-->
<%
	If Request.Querystring("REQUEST") <> "" Then
		REQUEST_TYPE = Request.Querystring("REQUEST")
	Else
		REQUEST_TYPE = "NORMAL"
	End If
	If Request.Querystring("DATE") <> "" Then
		PARAMETER_DATE = CDate(Request.Querystring("DATE"))
	Else
		PARAMETER_DATE = Date
	End If
	If Request.Querystring("AGENT") <> "" Then
		PARAMETER_AGENT = Request.Querystring("AGENT")
	Else
		PARAMETER_AGENT = "-1"
	End If
	
	If Request.Querystring("PULSEDATE") <> "" Then
		PULSE_DATE = CDate(Request.Querystring("PULSEDATE"))
	Else
		PULSE_DATE = PARAMETER_DATE
	End If

	SQLstmt = "SELECT OPS_USD_TYPE PARAMETER_DEPT " & _
	"FROM OPS_USER_DETAIL " & _
	"WHERE TO_DATE(?,'MM/DD/YYYY') BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE " & _
	"AND OPS_USD_OPS_USR_ID = ?"
	cmd.CommandText = SQLstmt
	cmd.Parameters(0).value = PARAMETER_DATE
	cmd.Parameters(1).value = PARAMETER_AGENT
	Set RSDEPT = cmd.Execute
	If Not RSDEPT.EOF Then
		PARAMETER_DEPT = RSDEPT("PARAMETER_DEPT")
	Else
		PARAMETER_DEPT = "RES"
	End If 
	Set RSDEPT = Nothing
	
	SQLstmt = "SELECT " & _
	"NVL(MAX(DECODE(OPS_PAR_CODE,'INTERVAL_LENGTH',TO_NUMBER(OPS_PAR_VALUE))),6) INTERVAL_LENGTH, " & _
	"NVL(MAX(DECODE(OPS_PAR_CODE,'SLIDER_STEP',TO_NUMBER(OPS_PAR_VALUE))),6) SLIDER_STEP " & _
	"FROM OPS_PARAMETER " & _
	"WHERE OPS_PAR_PARENT_TYPE = 'STF' " & _
	"AND TO_DATE(?,'MM/DD/YYYY') BETWEEN OPS_PAR_EFF_DATE AND OPS_PAR_DIS_DATE " & _
	"AND OPS_PAR_CODE IN ('SLIDER_STEP','INTERVAL_LENGTH')"
	cmd.CommandText = SQLstmt
	cmd.Parameters(0).value = PARAMETER_DATE
	Set RSSLIDER = cmd.Execute
	INTERVAL_LENGTH = RSSLIDER("INTERVAL_LENGTH")
	SLIDER_STEP = RSSLIDER("SLIDER_STEP")
	Set RSSLIDER = Nothing
	
	SQLstmt = "SELECT SYS_CDD_VALUE SCHEDULE_TYPE, " & _
	"CASE " & _
		"WHEN SYS_CDD_VALUE IN ('PICK','BASE','HOLW','ADDT','EXTD') THEN 'PHONE' " & _
		"WHEN SYS_CDD_VALUE IN ('MEET','PRES','PROJ','TRAN','FAMP','WFHU','MLTU','OTRG','NEWH') THEN 'TRAIN' " & _
		"WHEN SYS_CDD_VALUE IN ('SRPT','SRUN') THEN 'SRED' " & _
		"WHEN SYS_CDD_VALUE IN ('LNCH','LNFL') THEN 'LUNCH' " & _
		"WHEN REGEXP_LIKE(SYS_CDD_VALUE,'^VAC|UN$|PT$|PP$|HOLU|HOLR|SLIP|RESH|RCHG|ROUT|JURY|BRVT|FMHL') THEN 'VACA' " & _
	"END SCHEDULE_CLASS " & _
	"FROM SYS_CODE_DETAIL " & _
	"WHERE SYS_CDD_SYS_CDM_ID IN (33,34) " & _
	"ORDER BY SYS_CDD_VALUE"
	Set RSTYPE = Conn.Execute(SQLstmt)
	
	SQLstmt = "SELECT " & _
	"OPS_SCI_ID, " & _
	"OPS_SCI_OPS_USR_ID, " & _
	"TO_DATE(OPS_SCI_START) SCI_DATE, " & _
	"TO_CHAR(DECODE(OPS_SCI_STATUS,'APP',OPS_SCI_START,GREATEST(NVL(YESTERDAY_END,OPS_SCI_START),OPS_SCI_START)),'HH24:MI') SCI_START, " & _
	"CASE " & _
		"WHEN TO_CHAR(OPS_SCI_START,'HH24:MI') <> TO_CHAR(OPS_SCI_END,'HH24:MI') AND TO_CHAR(OPS_SCI_END,'HH24:MI') = '00:00' THEN '24:00' " & _
		"ELSE TO_CHAR(DECODE(OPS_SCI_STATUS,'APP',OPS_SCI_END,LEAST(NVL(TOMMOROW_START,OPS_SCI_END),OPS_SCI_END)),'HH24:MI') " & _
	"END SCI_END, " & _
	"OPS_SCI_TYPE, " & _
	"OPS_SCI_STATUS, " & _
	"OPS_SCI_OPS_USR_TYPE, " & _
	"OPS_SCI_NOTES, " & _
	"CASE " & _
		"WHEN NVL(LAG(TO_DATE(OPS_SCI_START)) OVER (ORDER BY TO_DATE(OPS_SCI_START), DECODE(OPS_SCI_STATUS,'APP',1,'SUB',2,'OPT',3), CASE WHEN RES_BUE_ID IS NOT NULL AND OPS_SCI_TYPE IN ('HOLR','HOLU') THEN 1 ELSE 2 END, TO_CHAR(OPS_SCI_START,'HH24:MI'), TO_CHAR(OPS_SCI_END,'HH24:MI'), DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'UN$|PP$|PT$'),0,OPS_SCI_TYPE,'AA'||OPS_SCI_TYPE)),TO_DATE(OPS_SCI_START)-1) <> TO_DATE(OPS_SCI_START) AND NVL(LEAD(TO_DATE(OPS_SCI_START)) OVER (ORDER BY TO_DATE(OPS_SCI_START), DECODE(OPS_SCI_STATUS,'APP',1,'SUB',2,'OPT',3), CASE WHEN RES_BUE_ID IS NOT NULL AND OPS_SCI_TYPE IN ('HOLR','HOLU') THEN 1 ELSE 2 END, TO_CHAR(OPS_SCI_START,'HH24:MI'), TO_CHAR(OPS_SCI_END,'HH24:MI'), DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'UN$|PP$|PT$'),0,OPS_SCI_TYPE,'AA'||OPS_SCI_TYPE)),TO_DATE(OPS_SCI_START)+1) <> TO_DATE(OPS_SCI_START) THEN 3 " & _
		"WHEN NVL(LAG(TO_DATE(OPS_SCI_START)) OVER (ORDER BY TO_DATE(OPS_SCI_START), DECODE(OPS_SCI_STATUS,'APP',1,'SUB',2,'OPT',3), CASE WHEN RES_BUE_ID IS NOT NULL AND OPS_SCI_TYPE IN ('HOLR','HOLU') THEN 1 ELSE 2 END, TO_CHAR(OPS_SCI_START,'HH24:MI'), TO_CHAR(OPS_SCI_END,'HH24:MI'), DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'UN$|PP$|PT$'),0,OPS_SCI_TYPE,'AA'||OPS_SCI_TYPE)),TO_DATE(OPS_SCI_START)-1) <> TO_DATE(OPS_SCI_START) THEN 1 " & _
		"WHEN NVL(LEAD(TO_DATE(OPS_SCI_START)) OVER (ORDER BY TO_DATE(OPS_SCI_START), DECODE(OPS_SCI_STATUS,'APP',1,'SUB',2,'OPT',3), CASE WHEN RES_BUE_ID IS NOT NULL AND OPS_SCI_TYPE IN ('HOLR','HOLU') THEN 1 ELSE 2 END, TO_CHAR(OPS_SCI_START,'HH24:MI'), TO_CHAR(OPS_SCI_END,'HH24:MI'), DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'UN$|PP$|PT$'),0,OPS_SCI_TYPE,'AA'||OPS_SCI_TYPE)),TO_DATE(OPS_SCI_START)+1) <> TO_DATE(OPS_SCI_START) THEN 2 " & _
		"ELSE 0 " & _
	"END TBODY_FLAG, " & _
	"DECODE(TO_DATE(OPS_SCI_START),TO_DATE(?,'MM/DD/YYYY'),NVL(TO_CHAR(YESTERDAY_END,'SSSSS')/60,0),0) SLIDER_MIN, " & _
	"DECODE(TO_DATE(OPS_SCI_START),TO_DATE(?,'MM/DD/YYYY'),NVL(TO_CHAR(TOMMOROW_START,'SSSSS')/60,1440),1440) SLIDER_MAX, " & _
	"MAX(CASE WHEN 'TRADE' = ? THEN 0 WHEN ROOT_START < TO_DATE(?,'MM/DD/YYYY') AND OPS_SCI_STATUS = 'APP' THEN 1 WHEN ROOT_START = TO_DATE(?,'MM/DD/YYYY') AND OPS_SCI_STATUS = 'APP' THEN 2 ELSE 0 END) OVER () YESTERDAY_AVAILABLE, " & _
	"NVL(TO_CHAR(YESTERDAY_END,'SSSSS')/60,0) YESTERDAY_MIN, " & _
	"MAX(CASE WHEN 'TRADE' = ? THEN 0 WHEN ROOT_END > TO_DATE(?,'MM/DD/YYYY') AND OPS_SCI_STATUS = 'APP' THEN 1 WHEN ROOT_END = TO_DATE(?,'MM/DD/YYYY') AND OPS_SCI_STATUS = 'APP' THEN 2 ELSE 0 END) OVER () TOMORROW_AVAILABLE,  " & _
	"NVL(TO_CHAR(TOMMOROW_START,'SSSSS')/60,1440) TOMORROW_MAX " & _
	"FROM " & _
	"( " & _
		"SELECT " & _
		"OPS_SCI_ID, " & _
		"OPS_SCI_OPS_USR_ID, " & _
		"OPS_SCI_START, " & _
		"OPS_SCI_END, " & _
		"OPS_SCI_TYPE, " & _
		"OPS_SCI_STATUS, " & _
		"OPS_SCI_OPS_USR_TYPE, " & _
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
			"OPS_SCI_ID, " & _
			"OPS_SCI_OPS_USR_ID, " & _
			"OPS_SCI_START, " & _
			"OPS_SCI_END, " & _
			"OPS_SCI_TYPE, " & _
			"OPS_SCI_STATUS, " & _
			"OPS_SCI_OPS_USR_TYPE, " & _
			"OPS_SCI_NOTES, " & _
			"CONNECT_BY_ROOT(OPS_SCI_START) ROOT_START, " & _
			"MAX(OPS_SCI_END) OVER (PARTITION BY CONNECT_BY_ROOT(OPS_SCI_START)) ROOT_END " & _
			"FROM " & _
			"( " & _
				"SELECT " & _
				"OPS_SCI_ID, " & _
				"OPS_SCI_OPS_USR_ID, " & _
				"OPS_SCI_START, " & _
				"OPS_SCI_END, " & _
				"OPS_SCI_TYPE, " & _
				"OPS_SCI_STATUS, " & _
				"OPS_SCI_OPS_USR_TYPE, " & _
				"OPS_SCI_NOTES, " & _
				"CASE " & _
					"WHEN LAG(OPS_SCI_START) OVER (ORDER BY OPS_SCI_STATUS, OPS_SCI_START) IS NULL " & _
					"OR LAG(OPS_SCI_END) OVER (ORDER BY OPS_SCI_STATUS, OPS_SCI_START) <> OPS_SCI_START " & _
					"OR LAG(OPS_SCI_STATUS) OVER (ORDER BY OPS_SCI_STATUS, OPS_SCI_START) <> OPS_SCI_STATUS THEN 1 " & _
				"END START_FLAG, " & _
				"ROW_NUMBER() OVER (ORDER BY OPS_SCI_STATUS, OPS_SCI_START, DECODE(OPS_SCI_START,OPS_SCI_END,DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'UN$|PP$|PT$'),0,OPS_SCI_TYPE,'AA'||OPS_SCI_TYPE),'ZZ')) TIMEBREAK_ORDER " & _
				"FROM OPS_SCHEDULE_INFO " & _
				"WHERE OPS_SCI_STATUS IN ('APP','SUB','OPT') " & _
				"AND " & _
				"( " & _
					"TO_DATE(OPS_SCI_START) = TO_DATE(?,'MM/DD/YYYY') " & _
					"OR " & _
					"( " & _
						"TO_DATE(OPS_SCI_START) IN (TO_DATE(?,'MM/DD/YYYY'),TO_DATE(?,'MM/DD/YYYY')) " & _
						"AND 'TRADE' <> ? " & _
					") " & _
				") " & _
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
	"ORDER BY SCI_DATE, DECODE(OPS_SCI_STATUS,'APP',1,'SUB',2,'OPT',3), CASE WHEN RES_BUE_ID IS NOT NULL AND OPS_SCI_TYPE IN ('HOLR','HOLU') THEN 1 ELSE 2 END, SCI_START, SCI_END, DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'UN$|PP$|PT$'),0,OPS_SCI_TYPE,'AA'||OPS_SCI_TYPE)"
	cmd.CommandText = SQLstmt
	cmd.Parameters(0).value = PARAMETER_DATE - 1
	cmd.Parameters(1).value = PARAMETER_DATE + 1
	cmd.Parameters(2).value = REQUEST_TYPE
	cmd.Parameters(3).value = PARAMETER_DATE
	cmd.Parameters(4).value = PARAMETER_DATE
	cmd.Parameters(5).value = REQUEST_TYPE
	cmd.Parameters(6).value = PARAMETER_DATE + 1
	cmd.Parameters(7).value = PARAMETER_DATE + 1
	cmd.Parameters(8).value = PARAMETER_DATE + 1
	cmd.Parameters(9).value = PARAMETER_DATE - 1
	cmd.Parameters(10).value = PARAMETER_DATE + 1
	cmd.Parameters(11).value = PARAMETER_DATE - 1	
	cmd.Parameters(12).value = PARAMETER_DATE - 1
	cmd.Parameters(13).value = PARAMETER_DATE + 1
	cmd.Parameters(14).value = PARAMETER_DATE
	cmd.Parameters(15).value = PARAMETER_DATE - 1
	cmd.Parameters(16).value = PARAMETER_DATE + 1
	cmd.Parameters(17).value = REQUEST_TYPE	
	cmd.Parameters(18).value = PARAMETER_AGENT
	cmd.Parameters(19).value = PARAMETER_DATE + 1
	cmd.Parameters(20).value = PARAMETER_DATE
	cmd.Parameters(21).value = PARAMETER_DATE
	Set RSSHIFT = cmd.Execute
 
	If Not RSSHIFT.EOF Then
		YESTERDAY_AVAILABLE = RSSHIFT("YESTERDAY_AVAILABLE")
		YESTERDAY_MIN = RSSHIFT("YESTERDAY_MIN")
		TOMORROW_AVAILABLE = RSSHIFT("TOMORROW_AVAILABLE")
		TOMORROW_MAX = RSSHIFT("TOMORROW_MAX")
		If YESTERDAY_AVAILABLE <> "0" or TOMORROW_AVAILABLE <> "0" Then
			EDITSCHEDULE_COLSPAN = 8
		Else
			EDITSCHEDULE_COLSPAN = 7
		End If 
	Else
		YESTERDAY_AVAILABLE = "0"
		YESTERDAY_MIN = "0"
		TOMORROW_AVAILABLE = "0"
		TOMORROW_MAX = "0"
		EDITSCHEDULE_COLSPAN = 7
	End If
%>
	<table id="EDITTABLE_<%=PARAMETER_AGENT%>_<%=Right("0" & Month(PARAMETER_DATE),2) & Right("0" & Day(PARAMETER_DATE),2) & Year(PARAMETER_DATE)%>" style="width:100%;font-size:.75em;">
		<% If REQUEST_TYPE = "TRADE" Then %>
			<caption>
				<%=AgentName(PARAMETER_AGENT)%> - <%=PARAMETER_DATE%>
			</caption>
		<% End If %>
		<thead>
			<tr class="subtable-td-padded-sm <% If PULSE_DATE = Date Then %>today-color<% Else %>past-color<% End If %>">
				<% If EDITSCHEDULE_COLSPAN = 8 Then %>
					<th class="subtable-td-padded-sm" style="width:5%;">Date</th>
				<% End If %>
				<th class="subtable-td-padded-sm" style="width:5%;">&nbsp;</th>
				<th class="subtable-td-padded-sm" style="width:<% If EDITSCHEDULE_COLSPAN = 8 Then %>65<% Else %>70<% End If %>%;"><% If REQUEST_TYPE <> "TRADE" Then %><%=AgentName(PARAMETER_AGENT)%>'s <% End If %>Schedule</th>
				<th class="subtable-td-padded-sm" style="width:5%;">&nbsp;</th>
				<th class="subtable-td-padded-sm" style="width:5%;">Type</th>
				<th class="subtable-td-padded-sm" style="width:5%;">Status</th>
				<th class="subtable-td-padded-sm" style="width:5%;">Dept</th>
				<th class="subtable-td-padded-sm" style="width:5%;">Notes</th>
			</tr>
		</thead>
		<% If Not RSSHIFT.EOF Then %>
			<% If YESTERDAY_AVAILABLE = "2" and (PULSE_SECURITY >= 5 or (PULSE_SECURITY >= 3 and PARAMETER_DATE-1 >= PULSE_PAYPERIOD_START)) Then %>
				<tbody id="SCITBODY_<%=PARAMETER_AGENT%>_<%=Right("0" & Month(PARAMETER_DATE-1),2) & Right("0" & Day(PARAMETER_DATE-1),2) & Year(PARAMETER_DATE-1)%>" class="altdate-entry-color" data-user="<%=PARAMETER_AGENT%>" data-parent-date="<%=Right("0" & Month(PARAMETER_DATE),2) & Right("0" & Day(PARAMETER_DATE),2) & Year(PARAMETER_DATE)%>" data-department="<%=PARAMETER_DEPT%>" data-date-min="<%=YESTERDAY_MIN%>" data-date-max="1440">
				</tbody>			
			<% End If %>
			<% Do While Not RSSHIFT.EOF %>
				<% If RSSHIFT("TBODY_FLAG") = "1" Or RSSHIFT("TBODY_FLAG") = "3" Then %>
					<tbody id="SCITBODY_<%=PARAMETER_AGENT%>_<%=Right("0" & Month(CDate(RSSHIFT("SCI_DATE"))),2) & Right("0" & Day(CDate(RSSHIFT("SCI_DATE"))),2) & Year(CDate(RSSHIFT("SCI_DATE")))%>" <% If CDate(RSSHIFT("SCI_DATE")) <> PARAMETER_DATE Then %>class="altdate-entry-color"<% End If %> data-user="<%=PARAMETER_AGENT%>" data-parent-date="<%=Right("0" & Month(PARAMETER_DATE),2) & Right("0" & Day(PARAMETER_DATE),2) & Year(PARAMETER_DATE)%>" data-department="<%=PARAMETER_DEPT%>" <% If CDate(RSSHIFT("SCI_DATE")) = PARAMETER_DATE-1 Then %>data-date-min="<%=YESTERDAY_MIN%>" data-date-max="1440" <% ElseIf CDate(RSSHIFT("SCI_DATE")) = PARAMETER_DATE Then %>data-date-min="0" data-date-max="1440" <% Else %>data-date-min="0" data-date-max="<%=TOMORROW_MAX%>" <% End If %>>
				<% End If %>
						<tr id="SCHEDULEROW_<%=RSSHIFT("OPS_SCI_ID")%>" class="<% If PULSE_DATE = Date Then %>today-color<% Else %>past-color<% End If %>">
							<input type="hidden" id="SCIDATE_<%=RSSHIFT("OPS_SCI_ID")%>" name="SCIDATE_<%=RSSHIFT("OPS_SCI_ID")%>" value="<%=RSSHIFT("SCI_DATE")%>" />
							<% If EDITSCHEDULE_COLSPAN = 8 Then %>
								<td class="subtable-td-padded-lg">
									<%=Month(CDate(RSSHIFT("SCI_DATE")))%>/<%=Day(CDate(RSSHIFT("SCI_DATE")))%>
								</td>
							<% End If %>
							<td class="subtable-td-padded-lg">
								<input type="hidden" id="SCISTART_<%=RSSHIFT("OPS_SCI_ID")%>" name="SCISTART_<%=RSSHIFT("OPS_SCI_ID")%>" value="<%=RSSHIFT("SCI_START")%>" />
								<div style="display:inline-block;white-space:nowrap;">
									<i id="STARTARROW_LEFT_<%=RSSHIFT("OPS_SCI_ID")%>" class="fas fa-caret-left icon-style-small"></i>
									<span id="STARTTIME_<%=RSSHIFT("OPS_SCI_ID")%>" style="padding:0 1px;"><%=RSSHIFT("SCI_START")%></span>
									<i id="STARTARROW_RIGHT_<%=RSSHIFT("OPS_SCI_ID")%>" class="fas fa-caret-right icon-style-small"></i>
								</div>
							</td>
							<td class="subtable-td-padded-lg">	
								<div id="SLIDER_<%=RSSHIFT("OPS_SCI_ID")%>" style="display:inline-block;width:100%;" data-user="<%=PARAMETER_AGENT%>" data-parent-date="<%=Right("0" & Month(PARAMETER_DATE),2) & Right("0" & Day(PARAMETER_DATE),2) & Year(PARAMETER_DATE)%>" data-slider-disabled="<% If Not (PULSE_SECURITY >= 5 or (PULSE_SECURITY >= 3 and CDate(RSSHIFT("SCI_DATE")) >= PULSE_PAYPERIOD_START)) Then %>true<% Else %>false<% End If %>" data-slider-min="<%=RSSHIFT("SLIDER_MIN")%>" data-slider-max="<%=RSSHIFT("SLIDER_MAX")%>" data-slider-interval="<%=INTERVAL_LENGTH%>" data-slider-step="<%=SLIDER_STEP%>"></div>	
							</td>
							<td class="subtable-td-padded-lg">
								<input type="hidden" id="SCIEND_<%=RSSHIFT("OPS_SCI_ID")%>" name="SCIEND_<%=RSSHIFT("OPS_SCI_ID")%>" value="<%=RSSHIFT("SCI_END")%>" />
								<div style="display:inline-block;white-space:nowrap;">
									<i id="ENDARROW_LEFT_<%=RSSHIFT("OPS_SCI_ID")%>" class="fas fa-caret-left icon-style-small"></i>
									<span id="ENDTIME_<%=RSSHIFT("OPS_SCI_ID")%>" style="padding:0 1px;"><%=RSSHIFT("SCI_END")%></span>
									<i id="ENDARROW_RIGHT_<%=RSSHIFT("OPS_SCI_ID")%>" class="fas fa-caret-right icon-style-small"></i>
								</div>
							</td>
							<td class="subtable-td-padded-lg">
								<input type="hidden" id="SCIUSER_<%=RSSHIFT("OPS_SCI_ID")%>" name="SCIUSER_<%=RSSHIFT("OPS_SCI_ID")%>" value="<%=PARAMETER_AGENT%>" />
								<select id="SCITYPE_<%=RSSHIFT("OPS_SCI_ID")%>" name="SCITYPE_<%=RSSHIFT("OPS_SCI_ID")%>" class="<% If PULSE_DATE = Date Then %> today-color <% Else %> past-color <% End If %> <% If CDate(RSSHIFT("SCI_DATE")) <> PARAMETER_DATE Then %> altdate-entry-color<% End If %>" style="padding-left:8px;" <% If Not (PULSE_SECURITY >= 5 or (PULSE_SECURITY >= 3 and CDate(RSSHIFT("SCI_DATE")) >= PULSE_PAYPERIOD_START)) Then %>disabled="disabled"<% End If %>">
									<% RSTYPE.MoveFirst %>
									<% Do While Not RSTYPE.EOF %>
										<option data-schedule-class="<%=RSTYPE("SCHEDULE_CLASS")%>" <% If RSTYPE("SCHEDULE_TYPE") = RSSHIFT("OPS_SCI_TYPE") Then %>selected="selected"<% End If %> value="<%=RSTYPE("SCHEDULE_TYPE")%>"><%=RSTYPE("SCHEDULE_TYPE")%></option>
										<% RSTYPE.MoveNext %>
									<% Loop %>
								</select>
							</td>
							<td class="subtable-td-padded-lg">
								<select id="SCISTATUS_<%=RSSHIFT("OPS_SCI_ID")%>" name="SCISTATUS_<%=RSSHIFT("OPS_SCI_ID")%>" class="<% If PULSE_DATE = Date Then %> today-color <% Else %> past-color <% End If %> <% If CDate(RSSHIFT("SCI_DATE")) <> PARAMETER_DATE Then %> altdate-entry-color<% End If %>" <% If Not (PULSE_SECURITY >= 5 or (PULSE_SECURITY >= 3 and CDate(RSSHIFT("SCI_DATE")) >= PULSE_PAYPERIOD_START)) Then %>disabled="disabled"<% End If %>>
									<option <% If RSSHIFT("OPS_SCI_STATUS") = "APP" Then %>selected="selected"<% End If %> value="APP">APP</option>
									<option <% If RSSHIFT("OPS_SCI_STATUS") = "SUB" Then %>selected="selected"<% End If %> value="SUB">SUB</option>
									<option <% If RSSHIFT("OPS_SCI_STATUS") = "DNY" Then %>selected="selected"<% End If %> value="DNY">DNY</option>
									<option <% If RSSHIFT("OPS_SCI_STATUS") = "DEL" Then %>selected="selected"<% End If %> value="DEL">DEL</option>
									<option <% If RSSHIFT("OPS_SCI_STATUS") = "OPT" Then %>selected="selected"<% End If %> value="OPT">OPT</option>
								</select>
							</td>
							<td class="subtable-td-padded-lg">
								<select id="SCIUSRTYPE_<%=RSSHIFT("OPS_SCI_ID")%>" name="SCIUSRTYPE_<%=RSSHIFT("OPS_SCI_ID")%>" class="<% If PULSE_DATE = Date Then %> today-color <% Else %> past-color <% End If %> <% If CDate(RSSHIFT("SCI_DATE")) <> PARAMETER_DATE Then %> altdate-entry-color<% End If %>" <% If Not (PULSE_SECURITY >= 5 or (PULSE_SECURITY >= 3 and CDate(RSSHIFT("SCI_DATE")) >= PULSE_PAYPERIOD_START)) Then %>disabled="disabled"<% End If %>>
									<option <% If RSSHIFT("OPS_SCI_OPS_USR_TYPE") = "ACC" Then %>selected="selected"<% End If %> value="ACC">ACC</option>
									<option <% If RSSHIFT("OPS_SCI_OPS_USR_TYPE") = "CRT" Then %>selected="selected"<% End If %> value="CRT">CRT</option>
									<option <% If RSSHIFT("OPS_SCI_OPS_USR_TYPE") = "DOC" Then %>selected="selected"<% End If %> value="DOC">DOC</option>
									<option <% If RSSHIFT("OPS_SCI_OPS_USR_TYPE") = "GRP" Then %>selected="selected"<% End If %> value="GRP">GRP</option>
									<option <% If RSSHIFT("OPS_SCI_OPS_USR_TYPE") = "OPS" Then %>selected="selected"<% End If %> value="OPS">OPS</option>
									<option <% If RSSHIFT("OPS_SCI_OPS_USR_TYPE") = "OSS" Then %>selected="selected"<% End If %> value="OSS">OSS</option>
									<option <% If RSSHIFT("OPS_SCI_OPS_USR_TYPE") = "POP" Then %>selected="selected"<% End If %> value="POP">POP</option>
									<option <% If RSSHIFT("OPS_SCI_OPS_USR_TYPE") = "RES" Then %>selected="selected"<% End If %> value="RES">RES</option>
								</select>
							</td>
							<td class="subtable-td-padded-lg"><i id="NOTESBUTTON_<%=RSSHIFT("OPS_SCI_ID")%>" class="far fa-comment icon-style-small"></i></td>
						</tr>
						<tr id="NOTESROW_<%=RSSHIFT("OPS_SCI_ID")%>" style="display:none;">
							<td class="subtable-td-padded-lg" colspan="<%=EDITSCHEDULE_COLSPAN-6%>">&nbsp;</td>
							<td class="subtable-td-padded-lg">
								<textarea id="SCINOTES_<%=RSSHIFT("OPS_SCI_ID")%>" name="SCINOTES_<%=RSSHIFT("OPS_SCI_ID")%>" class="<% If PULSE_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" style="width:100%;border-radius:5px;" maxlength="255" rows="2" <% If Not (PULSE_SECURITY >= 5 or (PULSE_SECURITY >= 3 and CDate(RSSHIFT("SCI_DATE")) >= PULSE_PAYPERIOD_START)) Then %>disabled="disabled"<% End If %>><%=RSSHIFT("OPS_SCI_NOTES")%></textarea>
							</td>
							<td class="subtable-td-padded-lg" colspan="5">&nbsp;</td>
						</tr>
					<% If RSSHIFT("TBODY_FLAG") = "2" Or RSSHIFT("TBODY_FLAG") = "3" Then %>
						</tbody>
					<% End If %>
				<% RSSHIFT.MoveNext %>
			<% Loop %>
			<% If TOMORROW_AVAILABLE = "2" Then %>
				<tbody id="SCITBODY_<%=PARAMETER_AGENT%>_<%=Right("0" & Month(PARAMETER_DATE+1),2) & Right("0" & Day(PARAMETER_DATE+1),2) & Year(PARAMETER_DATE+1)%>" class="altdate-entry-color" data-user="<%=PARAMETER_AGENT%>" data-parent-date="<%=Right("0" & Month(PARAMETER_DATE),2) & Right("0" & Day(PARAMETER_DATE),2) & Year(PARAMETER_DATE)%>" data-department="<%=PARAMETER_DEPT%>" data-date-min="0" data-date-max="<%=TOMORROW_MAX%>">
				</tbody>			
			<% End If %>
		<% Else %>
			<tbody id="SCITBODY_<%=PARAMETER_AGENT%>_<%=Right("0" & Month(PARAMETER_DATE),2) & Right("0" & Day(PARAMETER_DATE),2) & Year(PARAMETER_DATE)%>" data-user="<%=PARAMETER_AGENT%>" data-parent-date="<%=Right("0" & Month(PARAMETER_DATE),2) & Right("0" & Day(PARAMETER_DATE),2) & Year(PARAMETER_DATE)%>" data-department="<%=PARAMETER_DEPT%>" data-date-min="0" data-date-max="1440">
			</tbody>
		<% End If %>
		<% Set RSSHIFT = Nothing %>
		<tr id="SCHEDULEROW_0" class="new-entry-color <% If PULSE_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" style="display:none;">
			<% If EDITSCHEDULE_COLSPAN = 8 Then %>
				<td class="subtable-td-padded-lg">
					<select id="SCIDATE_0" name="SCIDATE_0" class="new-entry-color <% If PULSE_DATE = Date Then %> today-color<% Else %> past-color<% End If %>">
						<% If YESTERDAY_AVAILABLE <> "0" Then %>
							<option value="<%=Right("0" & Month(PARAMETER_DATE-1),2) & "/" & Right("0" & Day(PARAMETER_DATE-1),2) & "/" & Year(PARAMETER_DATE-1)%>" <% If Not (PULSE_SECURITY >= 5 or (PULSE_SECURITY >= 3 and PARAMETER_DATE-1 >= PULSE_PAYPERIOD_START)) Then %>disabled="disabled"<% End If %>><%=Month(PARAMETER_DATE-1)%>/<%=Day(PARAMETER_DATE-1)%></option>
						<% End If %>
						<option selected="selected" value="<%=Right("0" & Month(PARAMETER_DATE),2) & "/" & Right("0" & Day(PARAMETER_DATE),2) & "/" & Year(PARAMETER_DATE)%>"><%=Month(PARAMETER_DATE)%>/<%=Day(PARAMETER_DATE)%></option>
						<% If TOMORROW_AVAILABLE <> "0" Then %>
							<option value="<%=Right("0" & Month(PARAMETER_DATE+1),2) & "/" & Right("0" & Day(PARAMETER_DATE+1),2) & "/" & Year(PARAMETER_DATE+1)%>"><%=Month(PARAMETER_DATE+1)%>/<%=Day(PARAMETER_DATE+1)%></option>
						<% End If %>
					</select>
				</td>
			<% Else %>
				<input type="hidden" id="SCIDATE_0" name="SCIDATE_0" value="<%=PARAMETER_DATE%>" />
			<% End If %>
			<td class="subtable-td-padded-lg">
				<input type="hidden" id="SCISTART_0" name="SCISTART_0" value="" />
				<div style="display:inline-block;white-space:nowrap;">
					<i id="STARTARROW_LEFT_0" class="fas fa-caret-left icon-style-small"></i>
					<span id="STARTTIME_0" style="padding:0 1px;"></span>
					<i id="STARTARROW_RIGHT_0" class="fas fa-caret-right icon-style-small"></i>
				</div>	
			</td>
			<td class="subtable-td-padded-lg">
				<div id="SLIDER_0" style="display:inline-block;width:100%;" data-user="<%=PARAMETER_AGENT%>" data-parent-date="<%=Right("0" & Month(PARAMETER_DATE),2) & Right("0" & Day(PARAMETER_DATE),2) & Year(PARAMETER_DATE)%>" data-slider-disabled="false" data-slider-min="0" data-slider-max="1440" data-slider-interval="<%=INTERVAL_LENGTH%>" data-slider-step="<%=SLIDER_STEP%>"></div>
			</td>
			<td class="subtable-td-padded-lg">
				<input type="hidden" id="SCIEND_0" name="SCIEND_0" value="" />
				<div style="display:inline-block;white-space:nowrap;">
					<i id="ENDARROW_LEFT_0" class="fas fa-caret-left icon-style-small"></i>
					<span id="ENDTIME_0" style="padding:0 1px;"></span>
					<i id="ENDARROW_RIGHT_0" class="fas fa-caret-right icon-style-small"></i>
				</div>
			</td>
			<td class="subtable-td-padded-lg">
				<input type="hidden" id="SCIUSER_0" name="SCIUSER_0" value="<%=PARAMETER_AGENT%>" />
				<select id="SCITYPE_0" name="SCITYPE_0" class="new-entry-color <% If PULSE_DATE = Date Then %> today-color<% Else %> past-color<% End If %>" style="padding-left:8px;">
					<% RSTYPE.MoveFirst %>
					<% Do While Not RSTYPE.EOF %>
						<option data-schedule-class="<%=RSTYPE("SCHEDULE_CLASS")%>" <% If RSTYPE("SCHEDULE_TYPE") = "EXTD" Then %>selected="selected"<% End If %> value="<%=RSTYPE("SCHEDULE_TYPE")%>"><%=RSTYPE("SCHEDULE_TYPE")%></option>
						<% RSTYPE.MoveNext %>
					<% Loop %>
				</select>
			</td>
			<td class="subtable-td-padded-lg">
				<select id="SCISTATUS_0" name="SCISTATUS_0" class="new-entry-color <% If PULSE_DATE = Date Then %> today-color<% Else %> past-color<% End If %>">
					<option selected="selected" value="APP">APP</option>
					<option value="SUB">SUB</option>
					<option value="DNY">DNY</option>
					<option value="DEL">DEL</option>
					<option value="OPT">OPT</option>
				</select>
			</td>
			<td class="subtable-td-padded-lg">
				<select id="SCIUSRTYPE_0" name="SCIUSRTYPE_0" class="new-entry-color <% If PULSE_DATE = Date Then %> today-color<% Else %> past-color<% End If %>">
					<option value="ACC">ACC</option>
					<option value="CRT">CRT</option>
					<option value="DOC">DOC</option>
					<option value="GRP">GRP</option>
					<option value="OPS">OPS</option>
					<option value="OSS">OSS</option>
					<option value="POP">POP</option>
					<option value="RES">RES</option>
				</select>
			</td>
			<td class="subtable-td-padded-lg"><i id="NOTESBUTTON_0" class="far fa-comment icon-style-small"></i></td>
		</tr>
		<tr id="NOTESROW_0" class="new-entry-color" style="display:none;">
			<td class="subtable-td-padded-lg" colspan="<%=EDITSCHEDULE_COLSPAN-6%>">&nbsp;</td>
			<td class="subtable-td-padded-lg">
				<textarea id="SCINOTES_0" name="SCINOTES_0" class="<% If PULSE_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" style="width:100%;border-radius:5px;" rows="2"></textarea>
			</td>
			<td class="subtable-td-padded-lg" colspan="5">&nbsp;</td>
		</tr>
		<tr class="<% If PULSE_DATE = Date Then %>today-color<% Else %>past-color<% End If %>">
			<td class="subtable-td-padded-lg" colspan="<%=EDITSCHEDULE_COLSPAN-6%>">&nbsp;</td>
			<td class="subtable-td-padded-lg">
				<div style="margin-top:5px;">
					<div style="float:left;">
						<i id="REFRESH_<%=PARAMETER_AGENT%>_<%=Right("0" & Month(PARAMETER_DATE),2) & Right("0" & Day(PARAMETER_DATE),2) & Year(PARAMETER_DATE)%>" class="fas fa-sync-alt icon-style-large" title="Refresh" data-request="<%=REQUEST_TYPE%>"></i>
						<i id="FLEX_<%=PARAMETER_AGENT%>_<%=Right("0" & Month(PARAMETER_DATE),2) & Right("0" & Day(PARAMETER_DATE),2) & Year(PARAMETER_DATE)%>" class="fas fa-running icon-style-large" style="display:none;" title="Flex" data-request="<%=REQUEST_TYPE%>"></i>
					</div>
					<i id="NEWENTRY_<%=PARAMETER_AGENT%>_<%=Right("0" & Month(PARAMETER_DATE),2) & Right("0" & Day(PARAMETER_DATE),2) & Year(PARAMETER_DATE)%>" class="fas fa-plus-square icon-style-large" title="New Entry" data-request="<%=REQUEST_TYPE%>"></i>
					<div style="float:right;">
						<i id="HISTORY_<%=PARAMETER_AGENT%>_<%=Right("0" & Month(PARAMETER_DATE),2) & Right("0" & Day(PARAMETER_DATE),2) & Year(PARAMETER_DATE)%>" class="fas fa-history icon-style-large" title="Schedule History" data-request="<%=REQUEST_TYPE%>"></i>
					</div>
				</div>
			</td>
			<td class="subtable-td-padded-lg" colspan="5">&nbsp;</td>
		</tr>
		<% Set RSTYPE = Nothing %>
	</table>
	<script>
		$(document).ready(function() {
			$("div[id^='SLIDER_'][data-user='<%=PARAMETER_AGENT%>'][data-parent-date='<%=Right("0" & Month(PARAMETER_DATE),2) & Right("0" & Day(PARAMETER_DATE),2) & Year(PARAMETER_DATE)%>']:not([id='SLIDER_0'])").each(function(){
				var idArray = this.id.split("_");
				initializeSlider(idArray[1]);
			})
		});
	</script>
	
<!--#include file="pulsefunctions.asp"-->
<% Set cmd = Nothing %>
<% Conn.Close %>
<% Set Conn = Nothing %>