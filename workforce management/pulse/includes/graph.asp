<!--#include file="pulseheader.asp"-->
<%
	If Request.Querystring("DATE") <> "" then
		PARAMETER_DATE = CDate(Request.Querystring("DATE"))
	Else
		PARAMETER_DATE = Date
	End If
	If Request.Querystring("WORKGROUP") <> "" then
		PARAMETER_WORKGROUP = Replace(Request.Querystring("WORKGROUP"),"ALL","RES;SPT;OSR;SLS;SRV")
	Else
		PARAMETER_WORKGROUP = "RES"
	End If
	WORKGROUP_ARRAY = Split(PARAMETER_WORKGROUP,";")
%>
	<% If PULSE_SECURITY = 6 Then %>
		<form id="CONTROL_FORM" action="includes/formhandler.asp" method="post">
			<input type="hidden" id="NEWLINE_ID" value="0"/>
			<input type="hidden" id="CONTROLID_LIST" name="CONTROLID_LIST" value=""/>
	<% End If %>
<%
	For Each USE_WORKGROUP in WORKGROUP_ARRAY
		If USE_WORKGROUP <> "OSR" Then
			SQLstmt = "SELECT " & _
			"HH24, " & _
			"MI, " & _
			"STAFFED, " & _
			"PROJECTED_NEED, " & _
			"OFFERED, " & _
			"ANSWERED, " & _
			"STAFFING_NEED, " & _
			"DROP_LINE, " & _
			"WEIGHTED_NEED, " & _
			"MAX(GREATEST(NVL(STAFFED,0),NVL(PROJECTED_NEED,0),NVL(OFFERED,0),NVL(ANSWERED,0),NVL(STAFFING_NEED,0),NVL(DROP_LINE,0),NVL(WEIGHTED_NEED,0),5)) OVER () TICK_MAX " & _
			"FROM " & _
			"( " & _
				"SELECT " & _
				"TO_CHAR(TO_DATE(OSN.OPS_SCN_INTERVAL,'HH24:MI'),'HH24') HH24, " & _
				"TO_CHAR(TO_DATE(OSN.OPS_SCN_INTERVAL,'HH24:MI'),'MI') MI, " & _
				"OSN.OPS_SCN_STAFFED STAFFED, " & _
				"GREATEST(CEIL(OSN.OPS_SCN_PROJECTION * NVL(ROUND(DECODE(SUM(NVL2(DLY.RES_DLY_INTERVAL,OSN.OPS_SCN_PRO_CALLS,NULL)) OVER (),0,0,SUM(DLY.RES_DLY_QUEUED) OVER () /SUM(NVL2(DLY.RES_DLY_INTERVAL,OSN.OPS_SCN_PRO_CALLS,NULL)) OVER ())*DECODE(DECODE(SUM(DLY.RES_DLY_ANSWERED) OVER (),0,0,SUM(DLY.RES_DLY_ANSWERED*NVL2(DLY.RES_DLY_INTERVAL,OSN.OPS_SCN_PRO_AHT,NULL)) OVER ()/SUM(DLY.RES_DLY_ANSWERED) OVER ()),0,0,DECODE(SUM(DLY.RES_DLY_ANSWERED) OVER (),0,0,SUM(DLY.RES_DLY_ANSWERED*DLY.RES_DLY_ACTUAL_AHT) OVER ()/SUM(DLY.RES_DLY_ANSWERED) OVER ())/DECODE(SUM(DLY.RES_DLY_ANSWERED) OVER (),0,0,SUM(DLY.RES_DLY_ANSWERED*NVL2(DLY.RES_DLY_INTERVAL,OSN.OPS_SCN_PRO_AHT,NULL)) OVER ()/SUM(DLY.RES_DLY_ANSWERED) OVER ())),4),1)),1) PROJECTED_NEED, " & _
				"CEIL((DLY.RES_DLY_QUEUED*DLY.RES_DLY_ACTUAL_AHT)/1800) OFFERED, " & _
				"CEIL((DLY.RES_DLY_ANSWERED*DLY.RES_DLY_ACTUAL_AHT)/1800) ANSWERED, " & _
				"OSN.OPS_SCN_PROJECTION STAFFING_NEED, " & _
				"OSN.OPS_SCN_PROJECTION + CASE WHEN TO_DATE(?,'MM/DD/YYYY') <= TO_DATE(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE)) THEN NULL ELSE NULLIF(OSN.OPS_SCN_BUFFER,100) END DROP_LINE, " & _
				"GREATEST(ROUND( " & _
				".2*NVL(LAG(OSN.OPS_SCN_PROJECTION) OVER (ORDER BY OSN.OPS_SCN_INTERVAL),OSN.OPS_SCN_PROJECTION) + " & _
				".6*OSN.OPS_SCN_PROJECTION + " & _
				".2*NVL(LEAD(OSN.OPS_SCN_PROJECTION) OVER (ORDER BY OSN.OPS_SCN_INTERVAL),OSN.OPS_SCN_PROJECTION) " & _
				"),DECODE(?,'RES',2,'SPT',2,1)) WEIGHTED_NEED " & _
				"FROM OPS_SCHEDULE_NEED OSN " & _
				"LEFT JOIN RES_DAILY_STATS DLY " & _
				"ON DLY.RES_DLY_DATE = OSN.OPS_SCN_DATE " & _
				"AND DLY.RES_DLY_INTERVAL = OSN.OPS_SCN_INTERVAL " & _
				"AND DLY.RES_DLY_TYPE = DECODE(?,'RES','ALL','SPT','CSD',?) " & _
				"WHERE OSN.OPS_SCN_DATE = TO_DATE(?,'MM/DD/YYYY') " & _
				"AND OSN.OPS_SCN_TYPE = ? " & _
				"GROUP BY OSN.OPS_SCN_INTERVAL, OSN.OPS_SCN_STAFFED, OSN.OPS_SCN_PROJECTION, OSN.OPS_SCN_BUFFER, DLY.RES_DLY_INTERVAL, DLY.RES_DLY_QUEUED, DLY.RES_DLY_ACTUAL_AHT, DLY.RES_DLY_ANSWERED, OSN.OPS_SCN_PRO_CALLS, OSN.OPS_SCN_PRO_AHT " & _
				"ORDER BY OSN.OPS_SCN_INTERVAL " & _
			")"
			cmd.CommandText = SQLstmt
			cmd.Parameters(0).value = PARAMETER_DATE
			cmd.Parameters(1).value = USE_WORKGROUP
			cmd.Parameters(2).value = USE_WORKGROUP
			cmd.Parameters(3).value = USE_WORKGROUP
			cmd.Parameters(4).value = PARAMETER_DATE
			cmd.Parameters(5).value = USE_WORKGROUP
			Set RSGRAPH = cmd.Execute
		Else
			SQLstmt = "SELECT " & _
			"INTERVAL, " & _
			"STAFFED, " & _
			"OFFERED, " & _
			"ANSWERED, " & _
			"STAFFING_NEED, " & _
			"MAX(GREATEST(NVL(STAFFED,0),NVL(OFFERED,0),NVL(ANSWERED,0),NVL(STAFFING_NEED,0),5)) OVER () TICK_MAX " & _
			"FROM " & _
			"( " & _
				"SELECT " & _
				"OSN.OPS_SCN_INTERVAL INTERVAL, " & _
				"OSN.OPS_SCN_STAFFED STAFFED, " & _
				"CEIL((DLY.RES_DLY_QUEUED*DLY.RES_DLY_ACTUAL_AHT)/1800) OFFERED, " & _
				"CEIL((DLY.RES_DLY_ANSWERED*DLY.RES_DLY_ACTUAL_AHT)/1800) ANSWERED, " & _
				"CASE WHEN OSN.OPS_SCN_INTERVAL BETWEEN '06:00' AND '23:30' THEN 0 ELSE GREATEST(OSN.OPS_SCN_PROJECTION,0) END STAFFING_NEED " & _
				"FROM OPS_SCHEDULE_NEED OSN " & _
				"LEFT JOIN " & _
				"( " & _
					"SELECT RES_DLY_DATE, RES_DLY_INTERVAL, SUM(RES_DLY_ANSWERED) RES_DLY_ANSWERED, SUM(RES_DLY_QUEUED) RES_DLY_QUEUED, DECODE(SUM(RES_DLY_ANSWERED),0,0,SUM(RES_DLY_ANSWERED*RES_DLY_ACTUAL_AHT)/SUM(RES_DLY_ANSWERED)) RES_DLY_ACTUAL_AHT " & _
					"FROM RES_DAILY_STATS " & _
					"WHERE RES_DLY_TYPE IN ('ALL','CSD') " & _
					"AND " & _
					"( " & _
						"( " & _
							"RES_DLY_DATE = TO_DATE(?,'MM/DD/YYYY') " & _
							"AND RES_DLY_INTERVAL >= '19:30' " & _
						") " & _
						"OR " & _
						"( " & _
							"RES_DLY_DATE = TO_DATE(?,'MM/DD/YYYY') + 1 " & _
							"AND RES_DLY_INTERVAL <= '07:00' " & _
						") " & _
					") " & _
					"GROUP BY RES_DLY_DATE, RES_DLY_INTERVAL " & _
				")DLY " & _
				"ON DLY.RES_DLY_DATE = OSN.OPS_SCN_DATE " & _
				"AND DLY.RES_DLY_INTERVAL = OSN.OPS_SCN_INTERVAL " & _
				"WHERE " & _
				"( " & _
					"( " & _
						"OSN.OPS_SCN_DATE = TO_DATE(?,'MM/DD/YYYY') " & _
						"AND OSN.OPS_SCN_INTERVAL >= '19:30' " & _
					") " & _
					"OR " & _
					"( " & _
						"OSN.OPS_SCN_DATE = TO_DATE(?,'MM/DD/YYYY') + 1 " & _
						"AND OSN.OPS_SCN_INTERVAL <= '07:00' " & _
					") " & _
				") " & _
				"AND OSN.OPS_SCN_TYPE = 'OSR' " & _
				"ORDER BY OSN.OPS_SCN_DATE, OSN.OPS_SCN_INTERVAL " & _
			")"
			cmd.CommandText = SQLstmt
			cmd.Parameters(0).value = PARAMETER_DATE
			cmd.Parameters(1).value = PARAMETER_DATE
			cmd.Parameters(2).value = PARAMETER_DATE
			cmd.Parameters(3).value = PARAMETER_DATE
			Set RSGRAPH = cmd.Execute
		End If
%>
		<div id="<%=USE_WORKGROUP%>_graph" class="staffing_graph" data-workgroup="<%=Request.Querystring("WORKGROUP")%>" style="text-align:center;"></div>
		<script>
			google.charts.setOnLoadCallback(drawBasicGraph);	
			function drawBasicGraph() {	
				var tickValue = parseInt(<%=RSGRAPH("TICK_MAX")%>);
				var tickStep;
				var tickArray = [];
				if(tickValue <= 10){
					tickStep = 1;
				}
				else if (tickValue <= 20){
					tickStep = 2;
				}
				else if (tickValue <= 40){
					tickStep = 5;
				}
				else if (tickValue <= 120){
					tickStep = 10;
				}
				else{
					tickStep = 20;
				}
				for (var i = 0; i < tickValue + tickStep; i += tickStep) {
					tickArray.push(i);
				}
				
				var data = new google.visualization.DataTable();
				
				<% If USE_WORKGROUP <> "OSR" Then %>
					data.addColumn('timeofday','Interval');
					data.addColumn('number','<%=USE_WORKGROUP%> Staffed');
					<% If PARAMETER_DATE <= Date Then %>
						data.addColumn('number','<%=USE_WORKGROUP%> Projected Need');
						data.addColumn('number','<%=USE_WORKGROUP%> Answered');
						data.addColumn('number','<%=USE_WORKGROUP%> Offered');
					<% End If %>
					data.addColumn('number','<%=USE_WORKGROUP%> Staffing Need');
					<% If PARAMETER_DATE > Date Then %>
						data.addColumn('number','<%=USE_WORKGROUP%> Weighted Need');
					<% End If %>
					<% If RSGRAPH("DROP_LINE") <> "" Then %>
						data.addColumn('number','<%=USE_WORKGROUP%> Drop Line');
					<% End If %>
					
					data.addRows([
					<% Do While Not RSGRAPH.EOF %>
						[
							[<%=RSGRAPH("HH24")%>,<%=RSGRAPH("MI")%>,0],
							<%=RSGRAPH("STAFFED")%>, 
							<% If PARAMETER_DATE <= Date Then %>
								<%=RSGRAPH("PROJECTED_NEED")%>, 
								<%=RSGRAPH("ANSWERED")%>, 
								<%=RSGRAPH("OFFERED")%>, 
							<% End If %>
							<%=RSGRAPH("STAFFING_NEED")%>
							<% If PARAMETER_DATE > Date Then %>
								,
								<%=RSGRAPH("WEIGHTED_NEED")%>
							<% End If %>
							<% If RSGRAPH("DROP_LINE") <> "" Then %>
								,
								<%=RSGRAPH("DROP_LINE")%>
							<% End If %>
						]
						<% RSGRAPH.MoveNext %>
						<% If Not RSGRAPH.EOF Then %>
							,
						<% End If %>
					<% Loop %>
					]);
				<% Else %>
					data.addColumn('string','Interval');
					data.addColumn('number','<%=USE_WORKGROUP%> Staffed');
					<% If PARAMETER_DATE <= Date Then %>
						data.addColumn('number','<%=USE_WORKGROUP%> Answered');
						data.addColumn('number','<%=USE_WORKGROUP%> Offered');
					<% End If %>
					data.addColumn('number','<%=USE_WORKGROUP%> Staffing Need');
					
					data.addRows([
					<% Do While Not RSGRAPH.EOF %>
						[
							"<%=RSGRAPH("INTERVAL")%>",
							<%=RSGRAPH("STAFFED")%>, 
							<% If PARAMETER_DATE <= Date Then %>
								<%=RSGRAPH("ANSWERED")%>, 
								<%=RSGRAPH("OFFERED")%>, 
							<% End If %>
							<%=RSGRAPH("STAFFING_NEED")%>
						]
						<% RSGRAPH.MoveNext %>
						<% If Not RSGRAPH.EOF Then %>
							,
						<% End If %>
					<% Loop %>
					]);		
				<% End If %>
				var options = {
					title: '<%=USE_WORKGROUP%> Graph - <%=PARAMETER_DATE%> (<%=WeekdayName(Weekday(PARAMETER_DATE),True)%>)',
					fontName: 'Noto Sans',
					fontSize: 11,
					width: "100%",
					height: 600,
					curveType: 'function',
					<% If USE_WORKGROUP <> "OSR" Then %>
						<% If PARAMETER_DATE <= Date Then %>
							colors: ['yellow','gray','red','black','blue'],
						<% Else %>
							colors: ['yellow','blue','purple','green'],
						<% End If %>
					<% Else %>
						<% If PARAMETER_DATE <= Date Then %>
							colors: ['yellow','red','black','blue'],
						<% Else %>
							colors: ['yellow','blue'],
						<% End If %>			
					<% End If %>
					lineWidth: 1.75,
					pointSize: 3.5,
					chartArea: {
						width: "85%",
						height: 450,
						backgroundColor: '#DFDFDF'
					},
					hAxis: {
						title:'Interval',
						viewWindowMode: 'maximized',
						<% If USE_WORKGROUP <> "OSR" Then %>
							format: 'HH:mm',
						<% End If %>
						titleTextStyle: {
							fontSize: 16,
							bold: true,
							italic: false
						}
					},
					vAxis: {
						title: '<%=USE_WORKGROUP%> Associates',
						viewWindowMode: 'maximized',
						ticks: tickArray,
						titleTextStyle: {
							fontSize: 16,
							bold: true,
							italic: false
						}
					},
					legend:{
						position:'bottom',
						textStyle: {
							fontSize: '7pt',
						}
					}
				};
				var chart = new google.visualization.LineChart(document.getElementById("<%=USE_WORKGROUP%>_graph"));
				chart.draw(data, options);
				<% If PULSE_SECURITY = 6 Then %>
					$("#CONTROLSWRAPPER_<%=USE_WORKGROUP%>").show();
					overlapIdList = [];
					controlIdList = [];
				<% End If %>
			}
		</script>
		<% If PULSE_SECURITY = 6 Then %>
			<% If USE_WORKGROUP = "RES" or USE_WORKGROUP = "SPT" or USE_WORKGROUP = "OSR" Then %>
				<div id="CONTROLSWRAPPER_<%=USE_WORKGROUP%>" style="width:100%;text-align:center;display:none;" class="<% If PARAMETER_DATE = Date Then %> today-color<% Else %> past-color<% End If %>">
					<div style="font-size:1.5em;font-weight:900;"><%=USE_WORKGROUP%> Controls</div>
					<button id="CONTROLSTRIGGER_<%=USE_WORKGROUP%>_DROPCLOSED" class="btn btn-secondary <% If PARAMETER_DATE = Date Then %> today-color today-color-border white-background<% Else %> past-color past-color-border white-background<% End If %> control-item" type="button" style="margin-top:10px;">Drop Closed</button>
					<button id="CONTROLSTRIGGER_<%=USE_WORKGROUP%>_DROPOPEN" class="btn btn-secondary <% If PARAMETER_DATE = Date Then %> today-color today-color-border white-background<% Else %> past-color past-color-border white-background<% End If %> control-item" type="button" style="margin-top:10px;">Drop Open</button>
					<button id="CONTROLSTRIGGER_<%=USE_WORKGROUP%>_SRUNUNAVAILABLE" class="btn btn-secondary <% If PARAMETER_DATE = Date Then %> today-color today-color-border white-background<% Else %> past-color past-color-border white-background<% End If %> control-item" type="button" style="margin-top:10px;">SRUN Unavailable</button>
					<button id="CONTROLSTRIGGER_<%=USE_WORKGROUP%>_ADDOPEN" class="btn btn-secondary <% If PARAMETER_DATE = Date Then %> today-color today-color-border white-background<% Else %> past-color past-color-border white-background<% End If %> control-item" type="button" style="margin-top:10px;">Add Open</button>
					<button id="CONTROLSTRIGGER_<%=USE_WORKGROUP%>_ADDCLOSED" class="btn btn-secondary <% If PARAMETER_DATE = Date Then %> today-color today-color-border white-background<% Else %> past-color past-color-border white-background<% End If %> control-item" type="button" style="margin-top:10px;">Add Closed</button>
					<button id="CONTROLSTRIGGER_<%=USE_WORKGROUP%>_REDUCELIMIT" class="btn btn-secondary <% If PARAMETER_DATE = Date Then %> today-color today-color-border white-background<% Else %> past-color past-color-border white-background<% End If %> control-item" type="button" style="margin-top:10px;">UNP Limits</button>
					<button id="CONTROLSTRIGGER_<%=USE_WORKGROUP%>_ADDLIMIT" class="btn btn-secondary <% If PARAMETER_DATE = Date Then %> today-color today-color-border white-background<% Else %> past-color past-color-border white-background<% End If %> control-item" type="button" style="margin-top:10px;">OT Limits</button>
					<button id="CONTROLSTRIGGER_<%=USE_WORKGROUP%>_SELFTRADEOFF" class="btn btn-secondary <% If PARAMETER_DATE = Date Then %> today-color today-color-border white-background<% Else %> past-color past-color-border white-background<% End If %> control-item" type="button" style="margin-top:10px;">Self-Trade Off</button>
					<button id="CONTROLSTRIGGER_<%=USE_WORKGROUP%>_PICKCLOSED" class="btn btn-secondary <% If PARAMETER_DATE = Date Then %> today-color today-color-border white-background<% Else %> past-color past-color-border white-background<% End If %> control-item" type="button" style="margin-top:10px;">Pick Closed</button>
					<% If USE_WORKGROUP = "RES" or USE_WORKGROUP = "SPT" Then %>
						<button id="CONTROLSTRIGGER_<%=USE_WORKGROUP%>_NEWHIRE" class="btn btn-secondary <% If PARAMETER_DATE = Date Then %> today-color today-color-border white-background<% Else %> past-color past-color-border white-background<% End If %> control-item" type="button" style="margin-top:10px;">NEWH Wait Period</button>
					<% End If %>
					<div id="CONTROLSDIV_<%=USE_WORKGROUP%>_DROPCLOSED" class="edit-div <% If PARAMETER_DATE = Date Then %> today-color<% Else %> past-color<% End If %>"></div>
					<div id="CONTROLSDIV_<%=USE_WORKGROUP%>_DROPOPEN" class="edit-div <% If PARAMETER_DATE = Date Then %> today-color<% Else %> past-color<% End If %>"></div>
					<div id="CONTROLSDIV_<%=USE_WORKGROUP%>_SRUNUNAVAILABLE" class="edit-div <% If PARAMETER_DATE = Date Then %> today-color<% Else %> past-color<% End If %>"></div>
					<div id="CONTROLSDIV_<%=USE_WORKGROUP%>_ADDOPEN" class="edit-div <% If PARAMETER_DATE = Date Then %> today-color<% Else %> past-color<% End If %>"></div>
					<div id="CONTROLSDIV_<%=USE_WORKGROUP%>_ADDCLOSED" class="edit-div <% If PARAMETER_DATE = Date Then %> today-color<% Else %> past-color<% End If %>"></div>
					<div id="CONTROLSDIV_<%=USE_WORKGROUP%>_REDUCELIMIT" class="edit-div <% If PARAMETER_DATE = Date Then %> today-color<% Else %> past-color<% End If %>"></div>
					<div id="CONTROLSDIV_<%=USE_WORKGROUP%>_ADDLIMIT" class="edit-div <% If PARAMETER_DATE = Date Then %> today-color<% Else %> past-color<% End If %>"></div>
					<div id="CONTROLSDIV_<%=USE_WORKGROUP%>_SELFTRADEOFF" class="edit-div <% If PARAMETER_DATE = Date Then %> today-color<% Else %> past-color<% End If %>"></div>
					<div id="CONTROLSDIV_<%=USE_WORKGROUP%>_PICKCLOSED" class="edit-div <% If PARAMETER_DATE = Date Then %> today-color<% Else %> past-color<% End If %>"></div>
					<% If USE_WORKGROUP = "RES" or USE_WORKGROUP = "SPT" Then %>
						<div id="CONTROLSDIV_<%=USE_WORKGROUP%>_NEWHIRE" class="edit-div <% If PARAMETER_DATE = Date Then %> today-color<% Else %> past-color<% End If %>"></div>
					<% End If %>
					<input id="CONTROL_SUBMIT_<%=USE_WORKGROUP%>" type="submit" class="btn th-color <% If PARAMETER_DATE = Date Then %>today-color-background<% Else %>past-color-background<% End If %>" style="display:none;margin-top:15px;" value="Submit Changes"/>
					<div id="OVERLAP_MESSAGE_<%=USE_WORKGROUP%>" class="error-color" style="display:none;margin-top:15px;">
						Fix overlapping control entries before submitting.
					</div>
				</div>
			<% End If %>
		<% End If %>
	<% Next %>
	<% If PULSE_SECURITY = 6 Then %>
		</form>
	<% End If %>
<% Set cmd = Nothing %>
<% Conn.Close %>
<% Set Conn = Nothing %>