<!--#include file="pulseheader.asp"-->
<%
	If Request.Querystring("DATE") <> "" then
		PARAMETER_DATE = CDate(Request.Querystring("DATE"))
	Else
		PARAMETER_DATE = Date
	End If
%>
<% If InStr(PULSE_DEPARTMENT,"RES") <> 0 or PULSE_SECURITY >= 5 Then %>
<% 	
	SQLstmt = "SELECT " & _
	"NVL(ROUND(24*SUM(CASE WHEN OPS_SCI_TYPE IN ('CDPT','CDUN') AND OPS_USD_TEAM NOT IN ('SRV','SPT','OSR') THEN OPS_SCI_END-OPS_SCI_START END),2),0) RES_CD, " & _
	"NVL(ROUND(24*SUM(CASE WHEN OPS_SCI_TYPE IN ('SRPT','SRUN') AND OPS_USD_TEAM NOT IN ('SRV','SPT','OSR') THEN OPS_SCI_END-OPS_SCI_START END),2),0) RES_SR, " & _
	"NVL(ROUND(24*SUM(CASE WHEN OPS_SCI_TYPE IN ('SKPP','SKPT','SKUN','WXUN','WXPT','FMUN','FMPP','FMPT') AND OPS_USD_TEAM NOT IN ('SRV','SPT','OSR') THEN OPS_SCI_END-OPS_SCI_START END),2),0) RES_SK, " & _
	"NVL(ROUND(24*SUM(CASE WHEN OPS_SCI_TYPE IN ('VACA','RCHG','ROUT') AND OPS_USD_TEAM NOT IN ('SRV','SPT','OSR') THEN OPS_SCI_END-OPS_SCI_START END),2),0) RES_VACA, " & _
	"NVL(ROUND(24*SUM(CASE WHEN OPS_SCI_TYPE IN ('APPT','APUN','OTPT','OTUN') AND OPS_USD_TEAM NOT IN ('SRV','SPT','OSR') THEN OPS_SCI_END-OPS_SCI_START END),2),0) RES_EX, " & _
	"NVL(ROUND(24*SUM(CASE WHEN OPS_SCI_TYPE = 'ADDT' AND OPS_USD_TEAM NOT IN ('SRV','SPT','OSR') THEN OPS_SCI_END-OPS_SCI_START END),2),0) RES_ADDT, " & _
	"NVL(ROUND(24*SUM(CASE WHEN OPS_SCI_TYPE = 'SLIP' AND OPS_USD_TEAM NOT IN ('SRV','SPT','OSR') THEN OPS_SCI_END-OPS_SCI_START END),2),0) RES_SLIP, " & _
	"NVL(ROUND(24*SUM(CASE WHEN OPS_SCI_TYPE IN ('CDPT','CDUN') AND OPS_USD_TEAM IN ('SRV','SPT','OSR') THEN OPS_SCI_END-OPS_SCI_START END),2),0) SPT_CD, " & _
	"NVL(ROUND(24*SUM(CASE WHEN OPS_SCI_TYPE IN ('SRPT','SRUN') AND OPS_USD_TEAM IN ('SRV','SPT','OSR') THEN OPS_SCI_END-OPS_SCI_START END),2),0) SPT_SR, " & _
	"NVL(ROUND(24*SUM(CASE WHEN OPS_SCI_TYPE IN ('SKPP','SKPT','SKUN','WXUN','WXPT','FMUN','FMPP','FMPT') AND OPS_USD_TEAM IN ('SRV','SPT','OSR') THEN OPS_SCI_END-OPS_SCI_START END),2),0) SPT_SK, " & _
	"NVL(ROUND(24*SUM(CASE WHEN OPS_SCI_TYPE IN ('VACA','RCHG','ROUT') AND OPS_USD_TEAM IN ('SRV','SPT','OSR') THEN OPS_SCI_END-OPS_SCI_START END),2),0) SPT_VACA, " & _
	"NVL(ROUND(24*SUM(CASE WHEN OPS_SCI_TYPE IN ('APPT','APUN','OTPT','OTUN') AND OPS_USD_TEAM IN ('SRV','SPT','OSR') THEN OPS_SCI_END-OPS_SCI_START END),2),0) SPT_EX, " & _
	"NVL(ROUND(24*SUM(CASE WHEN OPS_SCI_TYPE = 'ADDT' AND OPS_USD_TEAM IN ('SRV','SPT','OSR') THEN OPS_SCI_END-OPS_SCI_START END),2),0) SPT_ADDT, " & _
	"NVL(ROUND(24*SUM(CASE WHEN OPS_SCI_TYPE = 'SLIP' AND OPS_USD_TEAM IN ('SRV','SPT','OSR') THEN OPS_SCI_END-OPS_SCI_START END),2),0) SPT_SLIP " & _
	"FROM OPS_SCHEDULE_INFO " & _
	"JOIN OPS_USER_DETAIL " & _
	"ON OPS_USD_OPS_USR_ID = OPS_SCI_OPS_USR_ID " & _
	"AND TO_DATE(OPS_SCI_START) BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE " & _
	"WHERE TO_DATE(OPS_SCI_START) = TO_DATE(?,'MM/DD/YYYY') " & _
	"AND OPS_USD_TYPE = 'RES' " & _
	"AND OPS_SCI_STATUS = 'APP' " & _
	"AND " & _
	"( " & _
		"OPS_SCI_TYPE IN ('CDPT','CDUN','SRPT','SRUN','SKPP','SKPT','SKUN','WXUN','WXPT','FMUN','FMPP','FMPT','VACA','RCHG','ROUT','ADDT','SLIP') " & _
		"OR " & _
		"( " & _
			"OPS_SCI_TYPE IN ('APPT','APUN','OTPT','OTUN') " & _
			"AND TO_DATE(INSERT_DATE) >= TO_DATE(?,'MM/DD/YYYY') " & _
		") " & _
	")"
	cmd.CommandText = SQLstmt
	cmd.Parameters(0).value = PARAMETER_DATE
	cmd.Parameters(1).value = PARAMETER_DATE - 7
	Set RSSTATS = cmd.Execute
%>
	<script>
		google.charts.setOnLoadCallback(drawBasicBar);
		
		function drawBasicBar() {	
			var data = google.visualization.arrayToDataTable([
				["Workgroup", "RES", "SPT"],
				["VACA", <%=RSSTATS("RES_VACA")%>, <%=RSSTATS("SPT_VACA")%>],
				["SR", <%=RSSTATS("RES_SR")%>, <%=RSSTATS("SPT_SR")%>],
				["SK", <%=RSSTATS("RES_SK")%>, <%=RSSTATS("SPT_SK")%>],
				["CD", <%=RSSTATS("RES_CD")%>, <%=RSSTATS("SPT_CD")%>],
				["EX", <%=RSSTATS("RES_EX")%>, <%=RSSTATS("SPT_EX")%>],
				["ADDT", <%=RSSTATS("RES_ADDT")%>, <%=RSSTATS("SPT_ADDT")%>]
				<% If RSSTATS("RES_SLIP") <> "0" or RSSTATS("SPT_SLIP") <> "0" Then %>
					,
					["VLOA", <%=RSSTATS("RES_SLIP")%>, <%=RSSTATS("SPT_SLIP")%>]
				<% End If %>
			]);
			var options = {
				title: "Hours Breakdown - <%=Month(PARAMETER_DATE) & "/" & Day(PARAMETER_DATE)%>",
				fontName: "Noto Sans",
				fontSize: 10,
				width: "100%",
				bar: { 
					groupWidth: "80%" 
				},
				hAxis: {
					ticks: [{v:40,f:"40"},{v:80,f:"80"},{v:120,f:"120"},{v:160,f:"160"},{v:200,f:"200+"}],
					viewWindow: {
						max: 200
					}
				},
				legend: {
					position: "none"
				}
			};
			var barchart = new google.visualization.BarChart(document.getElementById("schedule-stats-div"));
			barchart.draw(data, options);
			
			google.visualization.events.addListener(barchart, "click", barChartClick);
		}
	</script>
	<% Set RSSTATS = Nothing %>
<% Elseif PULSE_DEPARTMENT <> "" Then %>
	<div class="cddiv">
		<h6>Today's Events - <%=Month(PARAMETER_DATE) & "/" & Day(PARAMETER_DATE)%></h6>
		<%
			DEPT_ARRAY = Split(PULSE_DEPARTMENT,",")
			ReDim EVENT_PARAMETER_ARRAY(UBound(DEPT_ARRAY)+2)
			i = 0
			
			SQLstmt = "SELECT " & _
			"OPS_USR_ID, " & _
			"OPS_USR_NAME, " & _
			"DECODE(OPS_SCI_TYPE,'APPT','Appointment','APUN','Appointment','FAMP','FAM Trip','OTPP','Other','OTPT','Other','OTUN','Other','RESH','Time Off Cert','TOUN','Unpaid Time Off','VACA','Vacation','HOLR','Floating Holiday','SLIP','Voluntary LOA','OTRG','COVID Pay Protection') USE_TYPE, " & _
			"ROUND(SUM(24*(OPS_SCI_END-OPS_SCI_START)),2) USE_HOURS " & _
			"FROM OPS_SCHEDULE_INFO " & _
			"JOIN OPS_USER " & _
			"ON OPS_SCI_OPS_USR_ID = OPS_USR_ID " & _
			"JOIN OPS_USER_DETAIL " & _
			"ON OPS_SCI_OPS_USR_ID = OPS_USD_OPS_USR_ID " & _
			"AND TO_DATE(?,'MM/DD/YYYY') BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE " & _
			"LEFT JOIN RES_BUDGET_EXCEPTION " & _
			"ON RES_BUE_DATE = TO_DATE(OPS_SCI_START) " & _
			"AND RES_BUE_TYPE = 'NOR' " & _
			"WHERE TO_DATE(OPS_SCI_START) = TO_DATE(?,'MM/DD/YYYY') " & _
			"AND OPS_SCI_STATUS = 'APP' " & _
			"AND " & _
			"( " & _
				"OPS_SCI_TYPE IN ('APPT','APUN','FAMP','OTPP','OTPT','OTUN','RESH','TOUN','VACA','SLIP') " & _
				"OR " & _
				"( " & _
					"OPS_SCI_TYPE = 'HOLR' " & _
					"AND RES_BUE_ID IS NULL " & _
				") " & _
				"OR " & _
				"( " & _
					"OPS_SCI_TYPE = 'OTRG' " & _
					"AND REGEXP_INSTR(UPPER(OPS_SCI_NOTES),'SPLV|HRPP') > 0 " & _
				") " & _
			") " & _
			"AND OPS_USD_LOCATION IN ('MOT','WFD','WFH') " & _
			"AND OPS_USD_CLASS IN ('RGFT','RGPT','PT<30','LEAVE') " & _
			"AND OPS_USD_TYPE <> 'HRA' " & _
			"AND OPS_USD_PAY_RATE > 0 " & _
			"AND OPS_USD_TYPE IN ("
			EVENT_PARAMETER_ARRAY(i) = PARAMETER_DATE
			EVENT_PARAMETER_ARRAY(i+1) = PARAMETER_DATE
			i = i + 2
			For n = 0 to UBound(DEPT_ARRAY)
				If n <> UBound(DEPT_ARRAY) Then
					SQLstmt = SQLstmt & "?,"
				Else
					SQLstmt = SQLstmt & "?) "
				End If
				EVENT_PARAMETER_ARRAY(i) = DEPT_ARRAY(n)
				i = i + 1
			Next
			SQLstmt = SQLstmt & "GROUP BY OPS_USR_ID, OPS_USR_NAME, DECODE(OPS_SCI_TYPE,'APPT','Appointment','APUN','Appointment','FAMP','FAM Trip','OTPP','Other','OTPT','Other','OTUN','Other','RESH','Time Off Cert','TOUN','Unpaid Time Off','VACA','Vacation','HOLR','Floating Holiday','SLIP','Voluntary LOA','OTRG','COVID Pay Protection') " & _
			"ORDER BY OPS_USR_NAME, USE_TYPE"
			cmd.CommandText = SQLstmt
			For n = 0 to UBound(EVENT_PARAMETER_ARRAY)
				cmd.Parameters(n).value = EVENT_PARAMETER_ARRAY(n)
			Next
			Erase EVENT_PARAMETER_ARRAY
			Set RSEVENT = cmd.Execute
		%>
		<% If Not RSEVENT.EOF Then %>
			<table class="cdtable">
				<tr>
					<th style="width:40%">Associate</th>
					<th style="width:25%">Event Type</th>
					<th style="width:35%">Duration</th>
				</tr>
				<% Do While Not RSEVENT.EOF %>
					<tr>
						<td><span class="searchAgent" data-user="<%=RSEVENT("OPS_USR_ID")%>"><%=RSEVENT("OPS_USR_NAME")%></span></td>
						<td><%=RSEVENT("USE_TYPE")%></td>
						<td><%=RSEVENT("USE_HOURS")%></td>
					</tr>	
					<% RSEVENT.MoveNext %>
				<% Loop %>
			</table>
		<% Else %>
			No Events Found
		<% End If %>
		<% Set RSEVENT = Nothing %>
	<script>
		$(document).ready(function() {
			$("#schedule-stats-div").css("overflow-y", "auto");
		});
	</script>
<% End If %>
<% Set cmd = Nothing %>
<% Conn.Close %>
<% Set Conn = Nothing %>