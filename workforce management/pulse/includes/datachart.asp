<!--#include file="pulseheader.asp"-->
<%
	If Request.Querystring("DATE") <> "" then
		PARAMETER_DATE = CDate(Request.Querystring("DATE"))
	Else
		PARAMETER_DATE = Date
	End If
	If Request.Querystring("WORKGROUP") <> "" then
		PARAMETER_WORKGROUP = Replace(Replace(Request.Querystring("WORKGROUP"),"ALLRES","RES;SPT;OSR;SLS;SRV"),"ALLOSS","AIR;PRD;SKD")
	Else
		PARAMETER_WORKGROUP = "RES"
	End If
	WORKGROUP_ARRAY = Split(PARAMETER_WORKGROUP,";")
	
	For Each USE_WORKGROUP in WORKGROUP_ARRAY
		If USE_WORKGROUP = "OSR" Then
			START_INTERVAL = "00:00"
			END_INTERVAL = "05:30"
			ACTUAL_WORKGROUPS = "ALL,CSD,TAS,DHC"
		Elseif USE_WORKGROUP = "SPT" Then
			START_INTERVAL = "06:00"
			END_INTERVAL = "23:30"
			ACTUAL_WORKGROUPS = "CSD,TAS,DHC"
		Elseif USE_WORKGROUP = "RES" Then
			START_INTERVAL = "06:00"
			END_INTERVAL = "23:30"
			ACTUAL_WORKGROUPS = "ALL"
		Elseif USE_WORKGROUP = "SLS" or USE_WORKGROUP = "SRV" Then
			START_INTERVAL = "06:00"
			END_INTERVAL = "23:30"
			ACTUAL_WORKGROUPS = USE_WORKGROUP
		Elseif USE_WORKGROUP = "AIR" or USE_WORKGROUP = "PRD" or USE_WORKGROUP = "SKD" Then
			START_INTERVAL = "07:00"
			END_INTERVAL = "18:30"
			ACTUAL_WORKGROUPS = USE_WORKGROUP
		Elseif USE_WORKGROUP = "CRT" Then
			START_INTERVAL = "08:30"
			END_INTERVAL = "16:30"
			ACTUAL_WORKGROUPS = USE_WORKGROUP
		End If

		i = 0
		ReDim PARAMETER_ARRAY(17)
		
		SQLstmt = "WITH STAFF_DATA AS " & _
		"( " & _
			"SELECT * FROM " & _
			"( " & _
				"SELECT " & _
				"OPS_SCN_INTERVAL USE_INTERVAL, " & _
				"OPS_SCN_STAFFED-OPS_SCN_PROJECTION PLUS_MINUS, " & _
				"OPS_SCN_PRO_CALLS PROJECTED_CALLS, " & _
				"OPS_SCN_PRO_AHT PROJECTED_AHT " & _
				"FROM OPS_SCHEDULE_NEED " & _
				"WHERE OPS_SCN_TYPE = ? " & _
				"AND OPS_SCN_DATE = TO_DATE(?,'MM/DD/YYYY') " & _
				"AND OPS_SCN_INTERVAL BETWEEN ? AND ? " & _
			") " & _
			"NATURAL LEFT JOIN " & _
			"( " & _
				"SELECT RES_DLY_INTERVAL USE_INTERVAL, " & _
				"SUM(RES_DLY_QUEUED) OFFERED, " & _
				"SUM(RES_DLY_ANSWERED) ANSWERED, " & _
				"ROUND(DECODE(SUM(RES_DLY_ANSWERED),0,0,SUM(RES_DLY_ANSWERED*RES_DLY_ACTUAL_AHT)/SUM(RES_DLY_ANSWERED))) ACTUAL_AHT, " & _
				"DECODE(SUM(RES_DLY_QUEUED),0,0,SUM(RES_DLY_QUEUED*RES_DLY_SERVICE_LEVEL)/SUM(RES_DLY_QUEUED)) SERVICE_LEVEL, " & _
				"DECODE(SUM(RES_DLY_ANSWERED),0,0,SUM(RES_DLY_ANSWERED*RES_DLY_AVAILABILITY)/SUM(RES_DLY_ANSWERED)) AVAILABILITY, " & _
				"DECODE(SUM(RES_DLY_ANSWERED),0,0,SUM(RES_DLY_ANSWERED*RES_DLY_ASA)/SUM(RES_DLY_ANSWERED)) ASA " & _
				"FROM RES_DAILY_STATS " & _
				"WHERE RES_DLY_DATE = TO_DATE(?,'MM/DD/YYYY') " & _
				"AND RES_DLY_INTERVAL BETWEEN ? AND ? " & _
				"AND RES_DLY_TYPE IN ("
				PARAMETER_ARRAY(0) = USE_WORKGROUP
				PARAMETER_ARRAY(1) = PARAMETER_DATE
				PARAMETER_ARRAY(2) = START_INTERVAL
				PARAMETER_ARRAY(3) = END_INTERVAL
				PARAMETER_ARRAY(4) = PARAMETER_DATE
				PARAMETER_ARRAY(5) = START_INTERVAL
				PARAMETER_ARRAY(6) = END_INTERVAL
				i = 7
				USE_ARRAY = Split(ACTUAL_WORKGROUPS,",")
				For j = 0 to UBound(USE_ARRAY)
					If j <> UBound(USE_ARRAY) Then
						SQLstmt = SQLstmt & "?,"
					Else
						SQLstmt = SQLstmt & "?) "
					End If
					PARAMETER_ARRAY(i) = USE_ARRAY(j)
					i = i + 1
				Next
				PARAMETER_ARRAY(i) = USE_WORKGROUP
				PARAMETER_ARRAY(i+1) = USE_WORKGROUP				
				PARAMETER_ARRAY(i+2) = PARAMETER_DATE
				PARAMETER_ARRAY(i+3) = PARAMETER_DATE
				PARAMETER_ARRAY(i+4) = PARAMETER_DATE
				PARAMETER_ARRAY(i+5) = PARAMETER_DATE
				i = i + 5
				SQLstmt = SQLstmt & "GROUP BY RES_DLY_INTERVAL " & _
			") " & _
		") " & _
		"SELECT * " & _
		"FROM " & _
		"( " & _
			"SELECT " & _
			"USE_INTERVAL, " & _
			"PLUS_MINUS, " & _
			"CASE WHEN ? IN ('RES','SPT','SLS','SRV') OR PROJECTED_CALLS > 0 THEN PROJECTED_CALLS END PROJECTED_CALLS, " & _
			"OFFERED || NULLIF(' (' || ROUND(100*DECODE(PROJECTED_CALLS,0,NULL,OFFERED/PROJECTED_CALLS)) || '%)',' (%)') OFFERED, " & _
			"OFFERED OFFERED_ORDER, " & _
			"ANSWERED, " & _
			"NULLIF(ROUND(DECODE(OFFERED,0,100,100*ANSWERED/OFFERED)) || '%','%') PCH, " & _
			"CASE WHEN ? IN ('RES','SPT','SLS','SRV') OR PROJECTED_AHT > 0 THEN TO_CHAR(TO_DATE(PROJECTED_AHT,'SSSSS'),'MI:SS') END PROJECTED_AHT, " & _
			"TO_CHAR(TO_DATE(ACTUAL_AHT,'SSSSS'),'MI:SS') ACTUAL_AHT, " & _
			"NULLIF(ROUND(SERVICE_LEVEL,1) || '%','%') SERVICE_LEVEL, " & _
			"NULLIF(ROUND(AVAILABILITY,1) || '%','%') AVAILABILITY, " & _
			"TO_CHAR(TO_DATE(ROUND(ASA),'SSSSS'),'MI:SS') ASA " & _
			"FROM STAFF_DATA " & _
			"ORDER BY USE_INTERVAL " & _
		") " & _
		"UNION ALL " & _
		"SELECT " & _
		"'Totals', " & _
		"NULL, " & _
		"NULLIF(SUM(DECODE(TO_DATE(?,'MM/DD/YYYY'),TO_DATE(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE)),DECODE(OFFERED,NULL,0,PROJECTED_CALLS),PROJECTED_CALLS)),0), " & _
		"SUM(OFFERED) || NULLIF(' (' || ROUND(100*DECODE(SUM(DECODE(OFFERED,NULL,0,PROJECTED_CALLS)),0,0,SUM(DECODE(USE_INTERVAL,'23:00',0,OFFERED))/SUM(DECODE(OFFERED,NULL,0,PROJECTED_CALLS)))) || '%)',' (0%)'), " & _
		"SUM(OFFERED), " & _
		"SUM(ANSWERED), " & _
		"NULLIF(ROUND(DECODE(SUM(OFFERED),0,100,100*SUM(ANSWERED)/SUM(OFFERED))) || '%','%'), " & _
		"NULLIF(TO_CHAR(TO_DATE(ROUND(DECODE(SUM(DECODE(TO_DATE(?,'MM/DD/YYYY'),TO_DATE(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE)),DECODE(OFFERED,NULL,0,PROJECTED_CALLS),PROJECTED_CALLS)),0,0,SUM(PROJECTED_AHT*DECODE(TO_DATE(?,'MM/DD/YYYY'),TO_DATE(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE)),DECODE(OFFERED,NULL,0,PROJECTED_CALLS),PROJECTED_CALLS))/SUM(DECODE(TO_DATE(?,'MM/DD/YYYY'),TO_DATE(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE)),DECODE(OFFERED,NULL,0,PROJECTED_CALLS),PROJECTED_CALLS)))),'SSSSS'),'MI:SS'),'00:00'), " & _
		"TO_CHAR(TO_DATE(ROUND(DECODE(SUM(OFFERED),0,0,SUM(ACTUAL_AHT*OFFERED)/SUM(OFFERED))),'SSSSS'),'MI:SS'), " & _
		"NULLIF(ROUND(DECODE(SUM(ANSWERED),0,0,SUM(OFFERED*SERVICE_LEVEL)/SUM(OFFERED)),1) || '%','%'), " & _
		"NULLIF(ROUND(DECODE(SUM(ANSWERED),0,0,SUM(ANSWERED*AVAILABILITY)/SUM(ANSWERED)),1) || '%','%'), " & _
		"TO_CHAR(TO_DATE(ROUND(DECODE(SUM(ANSWERED),0,0,SUM(ANSWERED*ASA)/SUM(ANSWERED))),'SSSSS'),'MI:SS') " & _
		"FROM STAFF_DATA"
		cmd.CommandText = SQLstmt
		ReDim Preserve PARAMETER_ARRAY(i)
		For i = 0 to UBound(PARAMETER_ARRAY)
			cmd.Parameters(i).value = PARAMETER_ARRAY(i)
		Next
		Set RSCHART = cmd.Execute
	%>
		<div class="table-responsive">
			<table class="data-table table table-bordered table-striped subtable center" data-workgroup="<%=Request.Querystring("WORKGROUP")%>">
				<caption class="th-color <% If PARAMETER_DATE = Date Then %>today-color-background<% Else %>past-color-background<% End If %>">
					<%=DepartmentName(USE_WORKGROUP)%> Data - <%=FormatDateTime(Now,3)%>
				</caption>
				<thead>
					<tr class="th-color <% If PARAMETER_DATE = Date Then %>today-color-background<% Else %>past-color-background<% End If %>">
						<th class="td-padded-lg" style="padding-right:20px !important;">Interval</th>
						<th class="td-padded-lg" style="padding-right:20px !important;">+/-</th>
						<th class="td-padded-lg mobile-hide" style="padding-right:20px !important;">Projected Calls</th>
						<th class="td-padded-lg" style="padding-right:20px !important;">Offered</th>
						<th class="td-padded-lg mobile-hide" style="padding-right:20px !important;">Answered</th>
						<th class="td-padded-lg" style="padding-right:20px !important;">PCH</th>
						<th class="td-padded-lg mobile-hide" style="padding-right:20px !important;">Projected AHT</th>
						<th class="td-padded-lg" style="padding-right:20px !important;">Actual</th>
						<th class="td-padded-lg" style="padding-right:20px !important;">Service Level</th>
						<th class="td-padded-lg" style="padding-right:20px !important;">Avail</th>
						<th class="td-padded-lg" style="padding-right:20px !important;">ASA</th>
					</tr>
				</thead>
				<tbody>
				<% Do While Not RSCHART.EOF %>
					<% If RSCHART("USE_INTERVAL") = "Totals" Then %>
						<% Exit Do %>
					<% End If %>
					<tr>
						<td class="td-padded-sm"><%=RSCHART("USE_INTERVAL")%></td>
						<td class="td-padded-sm"><%=RSCHART("PLUS_MINUS")%></td>
						<td class="td-padded-sm mobile-hide"><%=RSCHART("PROJECTED_CALLS")%></td>
						<td class="td-padded-sm" data-order="<%=RSCHART("OFFERED_ORDER")%>"><%=RSCHART("OFFERED")%></td>
						<td class="td-padded-sm mobile-hide"><%=RSCHART("ANSWERED")%></td>
						<td class="td-padded-sm"><%=RSCHART("PCH")%></td>
						<td class="td-padded-sm mobile-hide"><%=RSCHART("PROJECTED_AHT")%></td>
						<td class="td-padded-sm"><%=RSCHART("ACTUAL_AHT")%></td>
						<td class="td-padded-sm"><%=RSCHART("SERVICE_LEVEL")%></td>
						<td class="td-padded-sm"><%=RSCHART("AVAILABILITY")%></td>
						<td class="td-padded-sm"><%=RSCHART("ASA")%></td>
					</tr>	
					<% RSCHART.MoveNext %>
				<% Loop %>
				</tbody>
				<tfoot>
					<tr style="font-weight:900;">
						<td class="td-padded-sm"><%=RSCHART("USE_INTERVAL")%></td>
						<td class="td-padded-sm"><%=RSCHART("PLUS_MINUS")%></td>
						<td class="td-padded-sm mobile-hide"><%=RSCHART("PROJECTED_CALLS")%></td>
						<td class="td-padded-sm"><%=RSCHART("OFFERED")%></td>
						<td class="td-padded-sm mobile-hide"><%=RSCHART("ANSWERED")%></td>
						<td class="td-padded-sm"><%=RSCHART("PCH")%></td>
						<td class="td-padded-sm mobile-hide"><%=RSCHART("PROJECTED_AHT")%></td>
						<td class="td-padded-sm"><%=RSCHART("ACTUAL_AHT")%></td>
						<td class="td-padded-sm"><%=RSCHART("SERVICE_LEVEL")%></td>
						<td class="td-padded-sm"><%=RSCHART("AVAILABILITY")%></td>
						<td class="td-padded-sm"><%=RSCHART("ASA")%></td>
					</tr>
				</tfoot>
			</table>
		</div>
	<% Next %>
	<script>
		$(document).ready(function() {
			$(".data-table").DataTable({
				"autoWidth": false,
				"paging": false,
				"searching": false,
				"info": false
			});
		});
	</script>
<!--#include file="pulsefunctions.asp"-->
<% Set cmd = Nothing %>
<% Conn.Close %>
<% Set Conn = Nothing %>