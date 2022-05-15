<!--#include file="pulseheader.asp"-->
<%
	If Request.Querystring("DATE") <> "" Then
		PARAMETER_DATE = CDate(Request.Querystring("DATE"))
	Else
		PARAMETER_DATE = Date
	End If
%>
<% If InStr(PULSE_DEPARTMENT,"RES") <> 0 or PULSE_SECURITY >= 5 Then %>
	<% If PARAMETER_DATE <= Date Then %>
	<%
		SQLstmt = "WITH GAUGE_DATA AS " & _
		"( " & _
			"SELECT " & _
			"TO_DATE(?,'MM/DD/YYYY') USE_DATE, " & _
			"ROUND(100*DECODE(PROJECTED_CALLS,0,DECODE(OFFERED,0,1,0),OFFERED/PROJECTED_CALLS),1) VOLUME_VARIANCE, " & _
			"ROUND(100*DECODE(PROJECTED_AHT,0,DECODE(ACTUAL_AHT,0,1,0),ACTUAL_AHT/PROJECTED_AHT),1) AHT_VARIANCE, " & _
			"ROUND(100*(DECODE(PROJECTED_CALLS,0,DECODE(OFFERED,0,1,0),OFFERED/PROJECTED_CALLS) * DECODE(PROJECTED_AHT,0,DECODE(ACTUAL_AHT,0,1,0),ACTUAL_AHT/PROJECTED_AHT)),1) FTE_VARIANCE, " & _
			"ROUND(SERVICE_LEVEL,1) SERVICE_LEVEL, " & _
			"ROUND(100*DECODE(OFFERED,0,1,ANSWERED/OFFERED),1) PCH, " & _
			"ROUND(AVAILABILITY,1) AVAILABILITY, " & _
			"RES_ASA, " & _
			"SPT_ASA, " & _
			"RES_ASA_GOAL, " & _
			"SPT_ASA_GOAL " & _
			"FROM " & _
			"( " & _
				"SELECT " & _
				"SUM(DECODE(RES_DLY_TYPE,'ALL',OPS_SCN_PRO_CALLS,0)) PROJECTED_CALLS, " & _
				"ROUND(DECODE(SUM(DECODE(RES_DLY_TYPE,'ALL',OPS_SCN_PRO_CALLS,0)),0,0,SUM(DECODE(RES_DLY_TYPE,'ALL',OPS_SCN_PRO_CALLS,0)*DECODE(RES_DLY_TYPE,'ALL',OPS_SCN_PRO_AHT,0))/SUM(DECODE(RES_DLY_TYPE,'ALL',OPS_SCN_PRO_CALLS,0)))) PROJECTED_AHT, " & _
				"SUM(DECODE(RES_DLY_TYPE,'ALL',RES_DLY_QUEUED,0)) OFFERED, " & _
				"SUM(DECODE(RES_DLY_TYPE,'ALL',RES_DLY_ANSWERED,0)) ANSWERED, " & _
				"ROUND(DECODE(SUM(DECODE(RES_DLY_TYPE,'ALL',RES_DLY_ANSWERED,0)),0,0,SUM(DECODE(RES_DLY_TYPE,'ALL',RES_DLY_ANSWERED,0)*DECODE(RES_DLY_TYPE,'ALL',RES_DLY_ACTUAL_AHT,0))/SUM(DECODE(RES_DLY_TYPE,'ALL',RES_DLY_ANSWERED,0)))) ACTUAL_AHT, " & _
				"DECODE(SUM(DECODE(RES_DLY_TYPE,'ALL',RES_DLY_QUEUED,0)),0,0,SUM(DECODE(RES_DLY_TYPE,'ALL',RES_DLY_QUEUED,0)*RES_DLY_SERVICE_LEVEL)/SUM(DECODE(RES_DLY_TYPE,'ALL',RES_DLY_QUEUED,0))) SERVICE_LEVEL, " & _
				"DECODE(SUM(DECODE(RES_DLY_TYPE,'ALL',RES_DLY_ANSWERED,0)),0,0,SUM(DECODE(RES_DLY_TYPE,'ALL',RES_DLY_ANSWERED,0)*DECODE(RES_DLY_TYPE,'ALL',RES_DLY_AVAILABILITY,0))/SUM(DECODE(RES_DLY_TYPE,'ALL',RES_DLY_ANSWERED,0))) AVAILABILITY, " & _
				"DECODE(SUM(DECODE(RES_DLY_TYPE,'ALL',RES_DLY_ANSWERED,0)),0,0,SUM(DECODE(RES_DLY_TYPE,'ALL',RES_DLY_ANSWERED,0)*DECODE(RES_DLY_TYPE,'ALL',RES_DLY_ASA,0))/SUM(DECODE(RES_DLY_TYPE,'ALL',RES_DLY_ANSWERED,0))) RES_ASA," & _
				"DECODE(SUM(DECODE(RES_DLY_TYPE,'CSD',RES_DLY_ANSWERED,0)),0,0,SUM(DECODE(RES_DLY_TYPE,'CSD',RES_DLY_ANSWERED,0)*DECODE(RES_DLY_TYPE,'CSD',RES_DLY_ASA,0))/SUM(DECODE(RES_DLY_TYPE,'CSD',RES_DLY_ANSWERED,0))) SPT_ASA " & _
				"FROM RES_DAILY_STATS " & _
				"JOIN OPS_SCHEDULE_NEED " & _
				"ON RES_DLY_DATE = OPS_SCN_DATE " & _
				"AND RES_DLY_INTERVAL = OPS_SCN_INTERVAL " & _
				"AND OPS_SCN_TYPE = 'RES' " & _
				"WHERE RES_DLY_DATE = TO_DATE(?,'MM/DD/YYYY') " & _
				"AND RES_DLY_INTERVAL BETWEEN '06:00' AND '23:30' " & _
				"AND RES_DLY_TYPE IN ('ALL','CSD') " & _
				"HAVING SUM(DECODE(RES_DLY_TYPE,'ALL',OPS_SCN_PRO_CALLS,0)) IS NOT NULL " & _
			") " & _
			"CROSS JOIN " & _
			"( " & _
				"SELECT NVL(MAX(TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,4))),60) RES_ASA_GOAL " & _
				"FROM OPS_PARAMETER " & _
				"WHERE OPS_PAR_PARENT_TYPE = 'STF' " & _
				"AND TO_DATE(?,'MM/DD/YYYY') BETWEEN OPS_PAR_EFF_DATE AND OPS_PAR_DIS_DATE " & _
				"AND OPS_PAR_CODE = 'RESASA' " & _
				"AND REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) = 'G' " & _
			") " & _
			"CROSS JOIN " & _
			"( " & _
				"SELECT NVL(MAX(TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,4))),120) SPT_ASA_GOAL " & _
				"FROM OPS_PARAMETER " & _
				"WHERE OPS_PAR_PARENT_TYPE = 'STF' " & _
				"AND TO_DATE(?,'MM/DD/YYYY') BETWEEN OPS_PAR_EFF_DATE AND OPS_PAR_DIS_DATE " & _
				"AND OPS_PAR_CODE = 'SPTASA' " & _
				"AND REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) = 'G' " & _
			") " & _
		") " & _
		"SELECT " & _
		"'Volume Variance' GAUGE_DESCRIPTION, " & _
		"'VOLUME' GAUGE_ID, " & _
		"ROUND(VOLUME_VARIANCE) GAUGE_VALUE, " & _
		"VOLUME_VARIANCE || '%' GAUGE_TEXT, " & _
		"CASE " & _
			"WHEN REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) = 'G' THEN '#00FF00' " & _
			"WHEN REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) = 'Y' THEN '#FFFF00' " & _
			"ELSE '#FF0000' " & _
		"END GAUGE_COLOR " & _
		"FROM GAUGE_DATA " & _
		"LEFT JOIN OPS_PARAMETER " & _
		"ON OPS_PAR_PARENT_TYPE = 'STF' " & _
		"AND OPS_PAR_CODE = 'VOLUME' " & _
		"AND USE_DATE BETWEEN OPS_PAR_EFF_DATE AND OPS_PAR_DIS_DATE " & _
		"AND INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),TO_CHAR(USE_DATE,'D')) > 0 " & _
		"AND VOLUME_VARIANCE > DECODE(TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3)),0,-1,TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3))) " & _
		"AND VOLUME_VARIANCE <= TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,4)) " & _
		"UNION ALL " & _
		"SELECT " & _
		"'AHT Variance', " & _
		"'AHT', " & _
		"ROUND(AHT_VARIANCE), " & _
		"AHT_VARIANCE || '%', " & _
		"CASE " & _
			"WHEN REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) = 'G' THEN '#00FF00' " & _
			"WHEN REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) = 'Y' THEN '#FFFF00' " & _
			"ELSE '#FF0000' " & _
		"END " & _
		"FROM GAUGE_DATA " & _
		"LEFT JOIN OPS_PARAMETER " & _
		"ON OPS_PAR_PARENT_TYPE = 'STF' " & _
		"AND OPS_PAR_CODE = 'AHT' " & _
		"AND USE_DATE BETWEEN OPS_PAR_EFF_DATE AND OPS_PAR_DIS_DATE " & _
		"AND INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),TO_CHAR(USE_DATE,'D')) > 0 " & _
		"AND AHT_VARIANCE > DECODE(TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3)),0,-1,TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3))) " & _
		"AND AHT_VARIANCE <= TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,4)) " & _
		"UNION ALL " & _
		"SELECT " & _
		"'FTE Variance', " & _
		"'FTE', " & _
		"ROUND(FTE_VARIANCE), " & _
		"FTE_VARIANCE || '%', " & _
		"CASE " & _
			"WHEN REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) = 'G' THEN '#00FF00' " & _
			"WHEN REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) = 'Y' THEN '#FFFF00' " & _
			"ELSE '#FF0000' " & _
		"END " & _
		"FROM GAUGE_DATA " & _
		"LEFT JOIN OPS_PARAMETER " & _
		"ON OPS_PAR_PARENT_TYPE = 'STF' " & _
		"AND OPS_PAR_CODE = 'FTE' " & _
		"AND USE_DATE BETWEEN OPS_PAR_EFF_DATE AND OPS_PAR_DIS_DATE " & _
		"AND INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),TO_CHAR(USE_DATE,'D')) > 0 " & _
		"AND FTE_VARIANCE > DECODE(TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3)),0,-1,TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3))) " & _
		"AND FTE_VARIANCE <= TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,4)) " & _
		"UNION ALL " & _
		"SELECT " & _
		"'Service Level', " & _
		"'SL', " & _
		"ROUND(SERVICE_LEVEL), " & _
		"SERVICE_LEVEL || '%', " & _
		"CASE " & _
			"WHEN REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) = 'G' THEN '#00FF00' " & _
			"WHEN REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) = 'Y' THEN '#FFFF00' " & _
			"ELSE '#FF0000' " & _
		"END " & _
		"FROM GAUGE_DATA " & _
		"LEFT JOIN OPS_PARAMETER " & _
		"ON OPS_PAR_PARENT_TYPE = 'STF' " & _
		"AND OPS_PAR_CODE = 'SL' " & _
		"AND USE_DATE BETWEEN OPS_PAR_EFF_DATE AND OPS_PAR_DIS_DATE " & _
		"AND INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),TO_CHAR(USE_DATE,'D')) > 0 " & _
		"AND SERVICE_LEVEL > DECODE(TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3)),0,-1,TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3))) " & _
		"AND SERVICE_LEVEL <= TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,4)) " & _
		"UNION ALL " & _
		"SELECT " & _
		"'PCH', " & _
		"'PCH', " & _
		"ROUND(PCH), " & _
		"PCH || '%', " & _
		"CASE " & _
			"WHEN REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) = 'G' THEN '#00FF00' " & _
			"WHEN REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) = 'Y' THEN '#FFFF00' " & _
			"ELSE '#FF0000' " & _
		"END " & _
		"FROM GAUGE_DATA " & _
		"LEFT JOIN OPS_PARAMETER " & _
		"ON OPS_PAR_PARENT_TYPE = 'STF' " & _
		"AND OPS_PAR_CODE = 'PCH' " & _
		"AND USE_DATE BETWEEN OPS_PAR_EFF_DATE AND OPS_PAR_DIS_DATE " & _
		"AND INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),TO_CHAR(USE_DATE,'D')) > 0 " & _
		"AND PCH > DECODE(TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3)),0,-1,TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3))) " & _
		"AND PCH <= TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,4)) " & _
		"UNION ALL " & _
		"SELECT " & _
		"'Availability', " & _
		"'AVAILABILITY', " & _
		"ROUND(AVAILABILITY), " & _
		"AVAILABILITY || '%', " & _
		"CASE " & _
			"WHEN REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) = 'G' THEN '#00FF00' " & _
			"WHEN REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) = 'Y' THEN '#FFFF00' " & _
			"ELSE '#FF0000' " & _
		"END " & _
		"FROM GAUGE_DATA " & _
		"LEFT JOIN OPS_PARAMETER " & _
		"ON OPS_PAR_PARENT_TYPE = 'STF' " & _
		"AND OPS_PAR_CODE = 'AVAILABILITY' " & _
		"AND USE_DATE BETWEEN OPS_PAR_EFF_DATE AND OPS_PAR_DIS_DATE " & _
		"AND INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),TO_CHAR(USE_DATE,'D')) > 0 " & _
		"AND AVAILABILITY > DECODE(TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3)),0,-1,TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3))) " & _
		"AND AVAILABILITY <= TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,4)) " & _
		"UNION ALL " & _
		"SELECT " & _
		"'RES ASA', " & _
		"'RESASA', " & _
		"ROUND(100*RES_ASA/RES_ASA_GOAL), " & _
		"TO_CHAR(TO_DATE(ROUND(RES_ASA),'SSSSS'),'MI:SS'), " & _
		"CASE " & _
			"WHEN REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) = 'G' THEN '#00FF00' " & _
			"WHEN REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) = 'Y' THEN '#FFFF00' " & _
			"ELSE '#FF0000' " & _
		"END " & _
		"FROM GAUGE_DATA " & _
		"LEFT JOIN OPS_PARAMETER " & _
		"ON OPS_PAR_PARENT_TYPE = 'STF' " & _
		"AND OPS_PAR_CODE = 'RESASA' " & _
		"AND USE_DATE BETWEEN OPS_PAR_EFF_DATE AND OPS_PAR_DIS_DATE " & _
		"AND INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),TO_CHAR(USE_DATE,'D')) > 0 " & _
		"AND ROUND(RES_ASA,1) > DECODE(TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3)),0,-1,TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3))) " & _
		"AND ROUND(RES_ASA,1) <= TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,4)) " & _
		"UNION ALL " & _
		"SELECT " & _
		"'SPT ASA', " & _
		"'SPTASA', " & _
		"ROUND(100*SPT_ASA/SPT_ASA_GOAL), " & _
		"TO_CHAR(TO_DATE(ROUND(SPT_ASA),'SSSSS'),'MI:SS'), " & _
		"CASE " & _
			"WHEN REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) = 'G' THEN '#00FF00' " & _
			"WHEN REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) = 'Y' THEN '#FFFF00' " & _
			"ELSE '#FF0000' " & _
		"END " & _
		"FROM GAUGE_DATA " & _
		"LEFT JOIN OPS_PARAMETER " & _
		"ON OPS_PAR_PARENT_TYPE = 'STF' " & _
		"AND OPS_PAR_CODE = 'SPTASA' " & _
		"AND USE_DATE BETWEEN OPS_PAR_EFF_DATE AND OPS_PAR_DIS_DATE " & _
		"AND INSTR(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1),TO_CHAR(USE_DATE,'D')) > 0 " & _
		"AND ROUND(SPT_ASA,1) > DECODE(TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3)),0,-1,TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3))) " & _
		"AND ROUND(SPT_ASA,1) <= TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,4))"
		cmd.CommandText = SQLstmt
		cmd.Parameters(0).value = PARAMETER_DATE
		cmd.Parameters(1).value = PARAMETER_DATE
		cmd.Parameters(2).value = PARAMETER_DATE
		cmd.Parameters(3).value = PARAMETER_DATE
		Set RSGAUGE = cmd.Execute
	%>
		<% If Not RSGAUGE.EOF Then %>
			<% Do While Not RSGAUGE.EOF %>
				<div class="circle-stat-container">
					<div <% If PULSE_SECURITY = 6 Then %> id="CIRCLESTAT_<%=RSGAUGE("GAUGE_ID")%>" style="cursor:pointer;" <% End If %> class="circle-stat" data-percent="<%=RSGAUGE("GAUGE_VALUE")%>" data-foregroundColor="<%=RSGAUGE("GAUGE_COLOR")%>" data-fillColor="<%=Replace(RSGAUGE("GAUGE_COLOR"),"00","AA")%>" data-text="<%=RSGAUGE("GAUGE_DESCRIPTION")%>" data-replacePercentageByText="<%=RSGAUGE("GAUGE_TEXT")%>"></div>
				</div>
				<% RSGAUGE.MoveNext %>
			<% Loop %>
			<% If PULSE_SECURITY = 6 Then %>
				<div style="width:100%;text-align:center;margin-top:10px;">
					<form id="CIRCLE_FORM" action="includes/formhandler.asp" method="post">
						<input type="hidden" id="NEWLINE_ID" value="0"/>
						<input type="hidden" id="CIRCLEID_LIST" name="CIRCLEID_LIST" value=""/>
						<div id="CIRCLEDIV_VOLUME" class="edit-div <% If PARAMETER_DATE = Date Then %> today-color<% Else %> past-color<% End If %>"></div>
						<div id="CIRCLEDIV_AHT" class="edit-div <% If PARAMETER_DATE = Date Then %> today-color<% Else %> past-color<% End If %>"></div>
						<div id="CIRCLEDIV_FTE" class="edit-div <% If PARAMETER_DATE = Date Then %> today-color<% Else %> past-color<% End If %>"></div>
						<div id="CIRCLEDIV_SL" class="edit-div <% If PARAMETER_DATE = Date Then %> today-color<% Else %> past-color<% End If %>"></div>
						<div id="CIRCLEDIV_PCH" class="edit-div <% If PARAMETER_DATE = Date Then %> today-color<% Else %> past-color<% End If %>"></div>
						<div id="CIRCLEDIV_AVAILABILITY" class="edit-div <% If PARAMETER_DATE = Date Then %> today-color<% Else %> past-color<% End If %>"></div>
						<div id="CIRCLEDIV_RESASA" class="edit-div <% If PARAMETER_DATE = Date Then %> today-color<% Else %> past-color<% End If %>"></div>
						<div id="CIRCLEDIV_SPTASA" class="edit-div <% If PARAMETER_DATE = Date Then %> today-color<% Else %> past-color<% End If %>"></div>
						<input id="CIRCLE_SUBMIT" type="submit" class="btn th-color <% If PARAMETER_DATE = Date Then %>today-color-background<% Else %>past-color-background<% End If %>" style="display:none;margin-top:15px;" value="Submit Changes"/>
						<div id="OVERLAP_MESSAGE" class="error-color" style="display:none;margin-top:15px;">
							Fix overlapping circle entries before submitting.
						</div>
					</form>
				</div>
			<% End If %>
		<% End If %>
	<% Else %>
	<%
		SQLstmt = "SELECT " & _
		"RES_WINDOW, " & _
		"SPT_WINDOW, " & _
		"RES_PM_WINDOW, " & _
		"SPT_PM_WINDOW, " & _
		"RES_FLAG, " & _
		"CASE " & _
			"WHEN RES_FLAG = 0 THEN '#00FF00' " & _
			"WHEN RES_FLAG = 1 THEN '#FFFF00' " & _
			"ELSE '#FF0000' " & _
		"END RES_FLAG_COLOR, " & _
		"SPT_FLAG, " & _
		"CASE " & _
			"WHEN SPT_FLAG = 0 THEN '#00FF00' " & _
			"WHEN SPT_FLAG = 1 THEN '#FFFF00' " & _
			"ELSE '#FF0000' " & _
		"END SPT_FLAG_COLOR, " & _
		"RES_SCORE, " & _
		"CASE " & _
			"WHEN RES_SCORE > 90 THEN '#00FF00' " & _
			"WHEN RES_SCORE > 80 THEN '#FFFF00' " & _
			"ELSE '#FF0000' " & _
		"END RES_SCORE_COLOR, " & _
		"SPT_SCORE, " & _
		"CASE " & _
			"WHEN SPT_SCORE > 90 THEN '#00FF00' " & _
			"WHEN SPT_SCORE > 80 THEN '#FFFF00' " & _
			"ELSE '#FF0000' " & _
		"END SPT_SCORE_COLOR " & _
		"FROM " & _
		"( " & _
			"SELECT " & _
			"MAX(OPS_SCN_INTERVAL) KEEP (DENSE_RANK FIRST ORDER BY RES_WINDOW DESC NULLS LAST, OPS_SCN_INTERVAL) || ' - ' || TO_CHAR(TO_DATE(MAX(OPS_SCN_INTERVAL) KEEP (DENSE_RANK FIRST ORDER BY RES_WINDOW DESC NULLS LAST, OPS_SCN_INTERVAL),'HH24:MI') + (1/12),'HH24:MI') RES_WINDOW, " & _
			"MAX(OPS_SCN_INTERVAL) KEEP (DENSE_RANK FIRST ORDER BY SPT_WINDOW DESC NULLS LAST, OPS_SCN_INTERVAL) || ' - ' || TO_CHAR(TO_DATE(MAX(OPS_SCN_INTERVAL) KEEP (DENSE_RANK FIRST ORDER BY SPT_WINDOW DESC NULLS LAST, OPS_SCN_INTERVAL),'HH24:MI') + (1/12),'HH24:MI') SPT_WINDOW, " & _
			"MAX(OPS_SCN_INTERVAL) KEEP (DENSE_RANK FIRST ORDER BY RES_PM_WINDOW NULLS LAST, OPS_SCN_INTERVAL) || ' - ' || TO_CHAR(TO_DATE(MAX(OPS_SCN_INTERVAL) KEEP (DENSE_RANK FIRST ORDER BY RES_PM_WINDOW NULLS LAST, OPS_SCN_INTERVAL),'HH24:MI') + (1/12),'HH24:MI') RES_PM_WINDOW, " & _
			"MAX(OPS_SCN_INTERVAL) KEEP (DENSE_RANK FIRST ORDER BY SPT_PM_WINDOW NULLS LAST, OPS_SCN_INTERVAL) || ' - ' || TO_CHAR(TO_DATE(MAX(OPS_SCN_INTERVAL) KEEP (DENSE_RANK FIRST ORDER BY SPT_PM_WINDOW NULLS LAST, OPS_SCN_INTERVAL),'HH24:MI') + (1/12),'HH24:MI') SPT_PM_WINDOW, " & _
			"SUM(RES_FLAG) RES_FLAG, " & _
			"SUM(SPT_FLAG) SPT_FLAG, " & _
			"ROUND(AVG(RES_SCORE)) RES_SCORE, " & _
			"ROUND(AVG(SPT_SCORE)) SPT_SCORE " & _
			"FROM " & _
			"( " & _
				"SELECT " & _
				"OPS_SCN_INTERVAL, " & _
				"CASE " & _
					"WHEN RES_PROJECTED > 100 AND RES_STAFFED/RES_PROJECTED < .9 THEN 1 " & _
					"WHEN RES_PROJECTED > 50 AND RES_STAFFED/RES_PROJECTED < .8 THEN 1 " & _
					"WHEN RES_PROJECTED > 20 AND RES_STAFFED/RES_PROJECTED < .75 THEN 1 " & _
					"WHEN RES_PROJECTED > 10 AND RES_STAFFED/RES_PROJECTED < .7 THEN 1 " & _
					"WHEN RES_STAFFED/RES_PROJECTED < .6 THEN 1 " & _
					"ELSE 0 " & _
				"END RES_FLAG, " & _
				"CASE " & _
					"WHEN SPT_PROJECTED > 20 AND SPT_STAFFED/SPT_PROJECTED < .8 THEN 1 " & _
					"WHEN SPT_PROJECTED > 15 AND SPT_STAFFED/SPT_PROJECTED < .75 THEN 1 " & _
					"WHEN SPT_STAFFED/SPT_PROJECTED < .6 THEN 1 " & _
					"ELSE 0 " & _
				"END SPT_FLAG, " & _
				"GREATEST(DECODE(SIGN(RES_STAFFED - RES_PROJECTED),1,100 - (.5*((RES_STAFFED - RES_PROJECTED)/ RES_PROJECTED))*POWER(RES_STAFFED - RES_PROJECTED,2),100 + (3*((RES_STAFFED - RES_PROJECTED)/ RES_PROJECTED))*POWER(RES_STAFFED - RES_PROJECTED,2)),0) RES_SCORE, " & _
				"GREATEST(DECODE(SIGN(SPT_STAFFED - SPT_PROJECTED),1,100 - (.5*((SPT_STAFFED - SPT_PROJECTED)/ SPT_PROJECTED))*POWER(SPT_STAFFED - SPT_PROJECTED,2),100 + (3*((SPT_STAFFED - SPT_PROJECTED)/ SPT_PROJECTED))*POWER(SPT_STAFFED - SPT_PROJECTED,2)),0) SPT_SCORE, " & _
				"SUM(CASE WHEN OPS_SCN_INTERVAL <= '22:00' THEN RES_PROJECTED ELSE NULL END) OVER (ORDER BY OPS_SCN_INTERVAL ROWS BETWEEN CURRENT ROW AND 3 FOLLOWING) RES_WINDOW, " & _
				"SUM(CASE WHEN OPS_SCN_INTERVAL <= '22:00' THEN SPT_PROJECTED ELSE NULL END) OVER (ORDER BY OPS_SCN_INTERVAL ROWS BETWEEN CURRENT ROW AND 3 FOLLOWING) SPT_WINDOW, " & _
				"AVG(CASE WHEN OPS_SCN_INTERVAL <= '22:00' AND RES_PROJECTED >= 10 THEN RES_STAFFED/RES_PROJECTED ELSE 10 END) OVER (ORDER BY OPS_SCN_INTERVAL ROWS BETWEEN CURRENT ROW AND 3 FOLLOWING) RES_PM_WINDOW, " & _
				"AVG(CASE WHEN OPS_SCN_INTERVAL <= '22:00' AND SPT_PROJECTED >= 5 THEN SPT_STAFFED/SPT_PROJECTED ELSE 10 END) OVER (ORDER BY OPS_SCN_INTERVAL ROWS BETWEEN CURRENT ROW AND 3 FOLLOWING) SPT_PM_WINDOW " & _
				"FROM " & _
				"( " & _
					"SELECT " & _
					"OPS_SCN_INTERVAL, " & _
					"MAX(DECODE(OPS_SCN_TYPE,'RES',OPS_SCN_PROJECTION,NULL)) RES_PROJECTED, " & _
					"MAX(DECODE(OPS_SCN_TYPE,'RES',OPS_SCN_STAFFED,NULL)) RES_STAFFED, " & _
					"MAX(DECODE(OPS_SCN_TYPE,'SPT',OPS_SCN_PROJECTION,NULL)) SPT_PROJECTED, " & _
					"MAX(DECODE(OPS_SCN_TYPE,'SPT',OPS_SCN_STAFFED,NULL)) SPT_STAFFED " & _
					"FROM OPS_SCHEDULE_NEED " & _
					"WHERE OPS_SCN_DATE = TO_DATE(?,'MM/DD/YYYY') " & _
					"AND OPS_SCN_TYPE IN ('RES','SPT') " & _
					"GROUP BY OPS_SCN_INTERVAL " & _
				") " & _
			") " & _
		")"
		cmd.CommandText = SQLstmt
		cmd.Parameters(0).value = PARAMETER_DATE
		Set RSSTAFF = cmd.Execute
	%>
		<% If Not RSSTAFF.EOF Then %>
			<div class="staffdiv">
				<span style="text-align:center;"><h4>RES Staffing Summary</h4></span>
				<table class="stafftable">
					<tr>
						<td>RES Score:</td>
						<td style="font-weight:900;color:<%=RSSTAFF("RES_SCORE_COLOR")%>;"><%=RSSTAFF("RES_SCORE")%></td>
					</tr>
					<tr>
						<td>Busiest projected window:</td>
						<td><%=RSSTAFF("RES_WINDOW")%></td>
					</tr>
					<tr>
						<td>Tightest staffed window:</td>
						<td><%=RSSTAFF("RES_PM_WINDOW")%></td>
					</tr>
					<tr>
						<td>Critical intervals:</td>
						<td style="font-weight:900;color:<%=RSSTAFF("RES_FLAG_COLOR")%>;"><%=RSSTAFF("RES_FLAG")%></td>
					</tr>
				</table>
			</div>
			<div class="staffdiv">
				<span style="text-align:center;"><h4>SPT Staffing Summary</h4></span>
				<table class="stafftable">
					<tr>
						<td>SPT Score:</td>
						<td style="font-weight:900;color:<%=RSSTAFF("SPT_SCORE_COLOR")%>;"><%=RSSTAFF("SPT_SCORE")%></td>
					</tr>
					<tr>
						<td>Busiest projected window:</td>
						<td><%=RSSTAFF("SPT_WINDOW")%></td>
					</tr>
					<tr>
						<td>Tightest staffed window:</td>
						<td><%=RSSTAFF("SPT_PM_WINDOW")%></td>
					</tr>
					<tr>
						<td>Critical intervals:</td>
						<td style="font-weight:900;color:<%=RSSTAFF("SPT_FLAG_COLOR")%>;"><%=RSSTAFF("SPT_FLAG")%></td>
					</tr>
				</table>
			</div>
		<% End If %>
	<% End If %>
	<script>
		$(document).ready(function() {
			$(".circle-stat").circliful({
				textBelow: 1,
				animation: 0,
				fontColor: "#000",
				textColor: "#000"
			});
		});
	</script>
<% Else %>
	<div class="staffdiv">&nbsp;</div>
<% End If %>
<% Set cmd = Nothing %>
<% Conn.Close %>
<% Set Conn = Nothing %>