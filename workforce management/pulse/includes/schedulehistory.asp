<!--#include file="pulseheader.asp"-->
<%	
	DATATABLES_BOOL = 0
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
	
	SQLstmt = "SELECT " & _
	"MODIFY_TIME, " & _
	"TO_CHAR(MODIFY_TIME,'YYYYMMDDHH24MI') MODIFY_ORDER, " & _
	"MODIFY_USER, " & _
	"COUNT(*) OVER () SCIH_COUNT " & _
	"FROM " & _
	"( " & _
		"SELECT " & _
		"MODIFY_TIME, " & _
		"MODIFY_USER, " & _
		"CASE WHEN SCI_LIST = LAG(SCI_LIST) OVER (ORDER BY MODIFY_TIME) THEN 1 ELSE 0 END REMOVE_FLAG " & _
		"FROM " & _
		"( " & _
			"SELECT DISTINCT " & _
			"OPS_SCIH_INSERT_DATE MODIFY_TIME, " & _
			"NVL(DECODE(LOWER(OPS_SCI_OSUSER),'svc_oms','Ops System',OPS_USR_NAME),'N/A') MODIFY_USER, " & _
			"LISTAGG(CASE WHEN OPS_SCI_STATUS IN ('APP','SUB') THEN '(' || OPS_SCI_TYPE || ')(' || OPS_SCI_STATUS || ')' || TO_CHAR(OPS_SCI_START,'HH24:MI') || ' - ' || TO_CHAR(OPS_SCI_END,'HH24:MI') END) WITHIN GROUP (ORDER BY DECODE(OPS_SCI_STATUS,'APP',1,'SUB',2,'OPT',3), CASE WHEN RES_BUE_ID IS NOT NULL AND OPS_SCI_TYPE IN ('HOLR','HOLU') THEN 1 ELSE 2 END, OPS_SCI_START, OPS_SCI_END, DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'UN$|PP$|PT$'),0,OPS_SCI_TYPE,'AA'||OPS_SCI_TYPE)) OVER (PARTITION BY OPS_SCIH_INSERT_DATE) SCI_LIST " & _
			"FROM " & _
			"( " & _
				"SELECT " & _
				"MAX(OPS_SCIH_INSERT_DATE) KEEP (DENSE_RANK LAST ORDER BY OPS_SCIH_INSERT_DATE) OVER (PARTITION BY GROUP_ID) OPS_SCIH_INSERT_DATE, " & _
				"MAX(OPS_SCI_OSUSER) KEEP (DENSE_RANK LAST ORDER BY OPS_SCIH_INSERT_DATE) OVER (PARTITION BY GROUP_ID) OPS_SCI_OSUSER, " & _
				"OPS_SCI_STATUS, " & _
				"OPS_SCI_TYPE, " & _
				"OPS_SCI_START, " & _
				"OPS_SCI_END " & _
				"FROM " & _
				"( " & _
					"SELECT " & _
					"OPS_SCIH_INSERT_DATE, " & _
					"OPS_SCI_OSUSER, " & _
					"OPS_SCI_STATUS, " & _
					"OPS_SCI_TYPE, " & _
					"OPS_SCI_START, " & _
					"OPS_SCI_END, " & _
					"SUM(GROUP_ID) OVER (ORDER BY OPS_SCIH_INSERT_DATE) GROUP_ID " & _
					"FROM " & _
					"( " & _
						"SELECT " & _
						"OPS_SCIH_INSERT_DATE - (5/24) OPS_SCIH_INSERT_DATE, " & _
						"OPS_SCI_OSUSER, " & _
						"OPS_SCI_STATUS, " & _
						"OPS_SCI_TYPE, " & _
						"OPS_SCI_START, " & _
						"OPS_SCI_END, " & _
						"CASE WHEN OPS_SCIH_INSERT_DATE - LAG(OPS_SCIH_INSERT_DATE) OVER (ORDER BY OPS_SCIH_INSERT_DATE, OPS_SCI_ID) <= 1/86400 THEN 0 ELSE 1 END GROUP_ID, " & _
						"CASE " & _
							"WHEN OPS_SCI_STATUS = LAG(OPS_SCI_STATUS) OVER (PARTITION BY OPS_SCI_ID ORDER BY OPS_SCIH_INSERT_DATE) " & _
							"AND OPS_SCI_TYPE = LAG(OPS_SCI_TYPE) OVER (PARTITION BY OPS_SCI_ID ORDER BY OPS_SCIH_INSERT_DATE) " & _
							"AND OPS_SCI_START = LAG(OPS_SCI_START) OVER (PARTITION BY OPS_SCI_ID ORDER BY OPS_SCIH_INSERT_DATE) " & _
							"AND OPS_SCI_END = LAG(OPS_SCI_END) OVER (PARTITION BY OPS_SCI_ID ORDER BY OPS_SCIH_INSERT_DATE) " & _
							"THEN 1 ELSE 0 " & _
						"END REMOVE_FLAG " & _
						"FROM OPS_SCHEDULE_INFO_HIST " & _
						"WHERE OPS_SCI_STATUS IN ('APP','SUB','DEL','DNY') " & _
						"AND NOT " & _
						"( " & _
							"OPS_SCI_STATUS = 'SUB' " & _
							"AND TO_CHAR(OPS_SCI_START,'HH24:MI') = TO_CHAR(OPS_SCI_END,'HH24:MI') " & _
						") " & _
						"AND TO_DATE(OPS_SCI_START) = TO_DATE(?,'MM/DD/YYYY') " & _
						"AND OPS_SCI_OPS_USR_ID = ? " & _
						"AND OPS_SCI_ID IN " & _
						"( " & _
							"SELECT DISTINCT OPS_SCI_ID " & _
							"FROM OPS_SCHEDULE_INFO_HIST " & _
							"WHERE OPS_SCI_STATUS IN ('APP','SUB') " & _
							"AND NOT " & _
							"( " & _
								"OPS_SCI_STATUS = 'SUB' " & _
								"AND TO_CHAR(OPS_SCI_START,'HH24:MI') = TO_CHAR(OPS_SCI_END,'HH24:MI') " & _
							") " & _
							"AND ADD_MONTHS(TO_DATE(OPS_SCIH_INSERT_DATE),48) >= TO_DATE(OPS_SCI_START) " & _
							"AND TO_DATE(OPS_SCI_START) = TO_DATE(?,'MM/DD/YYYY') " & _
							"AND OPS_SCI_OPS_USR_ID = ? " & _
						") " & _
					") " & _
					"WHERE REMOVE_FLAG = 0 " & _
				") " & _
			") " & _
			"LEFT JOIN OPS_USER " & _
			"ON UPPER(REPLACE(OPS_USR_NT_ID,'MLTMTKA\','')) = UPPER(OPS_SCI_OSUSER) " & _
			"AND TO_DATE(OPS_SCIH_INSERT_DATE) BETWEEN OPS_USR_EFF_DATE AND OPS_USR_DIS_DATE " & _
			"LEFT JOIN RES_BUDGET_EXCEPTION " & _
			"ON TO_DATE(OPS_SCI_START) = RES_BUE_DATE " & _
			"AND RES_BUE_TYPE = 'NOR' " & _
		") " & _
	") " & _
	"WHERE REMOVE_FLAG = 0"
	cmd.CommandText = SQLstmt
	cmd.Parameters(0).value = PARAMETER_DATE
	cmd.Parameters(1).value = PARAMETER_AGENT
	cmd.Parameters(2).value = PARAMETER_DATE
	cmd.Parameters(3).value = PARAMETER_AGENT
	Set RSHISTMASTER = cmd.Execute
	
%>
	<div class="table-responsive" style="margin-bottom:1rem;">
		<table id="SCHEDULEHISTORY_TABLE" class="table table-bordered center" style="margin-bottom:0;background-color:#fff;">
			<caption class="th-color <% If PARAMETER_DATE = Date Then %>today-color-background<% Else %>past-color-background<% End If %>">
				<%=AgentName(PARAMETER_AGENT)%>'s Schedule History on <%=PARAMETER_DATE%> - <%=FormatDateTime(Now,3)%>
				<% 
					COUNT_TEXT = ""
					If Not RSHISTMASTER.EOF Then
						If RSHISTMASTER("SCIH_COUNT") = "1" Then
							COUNT_TEXT = "1 Entry Found"
						Else
							DATATABLES_BOOL = 1
							COUNT_TEXT = RSHISTMASTER("SCIH_COUNT") & " Entries Found"
						End If
					End If
				%>
				<div style="float:right"><%=COUNT_TEXT%></div>
			</caption>
			<thead>
				<tr class="th-color <% If PARAMETER_DATE = Date Then %>today-color-background<% Else %>past-color-background<% End If %>">
					<th style="width:35%">Modified Time</th>
					<th style="width:35%">Modified User</th>
					<th style="width:30%">Shifts</th>
				</tr>
			</thead>
			<tbody>
			<% If Not RSHISTMASTER.EOF Then %>
				<% Do While Not RSHISTMASTER.EOF %>
					<tr>
						<td data-order="<%=RSHISTMASTER("MODIFY_ORDER")%>"><%=RSHISTMASTER("MODIFY_TIME")%></td>
						<td><%=RSHISTMASTER("MODIFY_USER")%></td>
						<% 
							SQLstmt = "SELECT " & _
							"OPS_SCI_STATUS, " & _
							"TO_CHAR(OPS_SCI_START,'HH24:MI') SCI_START, " & _
							"TO_CHAR(OPS_SCI_END,'HH24:MI') SCI_END, " & _
							"OPS_SCI_END, " & _
							"OPS_SCI_TYPE, " & _
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
								"WHEN (OPS_SCI_TYPE NOT IN ('HOLR','HOLU') OR RES_BUE_ID IS NULL) AND OPS_SCI_END <> NVL(LEAD(OPS_SCI_START) OVER (ORDER BY DECODE(OPS_SCI_STATUS,'APP',1,'SUB',2), CASE WHEN RES_BUE_ID IS NOT NULL AND OPS_SCI_TYPE IN ('HOLR','HOLU') THEN 1 ELSE 2 END, OPS_SCI_START, OPS_SCI_END, DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'UN$|PP$|PT$'),0,OPS_SCI_TYPE,'AA'||OPS_SCI_TYPE)),OPS_SCI_END) THEN 1 " & _
								"WHEN OPS_SCI_STATUS = 'APP' AND NVL(LEAD(OPS_SCI_STATUS) OVER (ORDER BY DECODE(OPS_SCI_STATUS,'APP',1,'SUB',2), CASE WHEN RES_BUE_ID IS NOT NULL AND OPS_SCI_TYPE IN ('HOLR','HOLU') THEN 1 ELSE 2 END, OPS_SCI_START, OPS_SCI_END, DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'UN$|PP$|PT$'),0,OPS_SCI_TYPE,'AA'||OPS_SCI_TYPE)),'APP') <> 'APP' THEN 1 " & _
								"WHEN OPS_SCI_TYPE IN ('HOLR','HOLU') AND RES_BUE_ID IS NOT NULL AND LEAD(OPS_SCI_TYPE) OVER (ORDER BY DECODE(OPS_SCI_STATUS,'APP',1,'SUB',2), CASE WHEN RES_BUE_ID IS NOT NULL AND OPS_SCI_TYPE IN ('HOLR','HOLU') THEN 1 ELSE 2 END, OPS_SCI_START, OPS_SCI_END, DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'UN$|PP$|PT$'),0,OPS_SCI_TYPE,'AA'||OPS_SCI_TYPE)) NOT IN ('HOLR','HOLU') THEN 1 " & _
								"ELSE 0 " & _
							"END GAP_FLAG " & _
							"FROM " & _
							"( " & _
								"SELECT DISTINCT " & _
								"MAX(OPS_SCI_START) KEEP (DENSE_RANK LAST ORDER BY OPS_SCIH_INSERT_DATE, DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'^VAC|UN$|PT$|PP$|HOLU|HOLR|SLIP|RESH|RCHG|ROUT|JURY|BRVT'),0,OPS_SCI_TYPE,'ZZ' || OPS_SCI_TYPE), DECODE(OPS_SCI_STATUS,'SUB',1,'COM',2,3)) OVER (PARTITION BY OPS_SCI_ID) OPS_SCI_START, " & _
								"MAX(OPS_SCI_END) KEEP (DENSE_RANK LAST ORDER BY OPS_SCIH_INSERT_DATE, DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'^VAC|UN$|PT$|PP$|HOLU|HOLR|SLIP|RESH|RCHG|ROUT|JURY|BRVT'),0,OPS_SCI_TYPE,'ZZ' || OPS_SCI_TYPE), DECODE(OPS_SCI_STATUS,'SUB',1,'COM',2,3)) OVER (PARTITION BY OPS_SCI_ID) OPS_SCI_END, " & _
								"MAX(OPS_SCI_TYPE) KEEP (DENSE_RANK LAST ORDER BY OPS_SCIH_INSERT_DATE, DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'^VAC|UN$|PT$|PP$|HOLU|HOLR|SLIP|RESH|RCHG|ROUT|JURY|BRVT'),0,OPS_SCI_TYPE,'ZZ' || OPS_SCI_TYPE), DECODE(OPS_SCI_STATUS,'SUB',1,'COM',2,3)) OVER (PARTITION BY OPS_SCI_ID) OPS_SCI_TYPE, " & _
								"MAX(OPS_SCI_STATUS) KEEP (DENSE_RANK LAST ORDER BY OPS_SCIH_INSERT_DATE, DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'^VAC|UN$|PT$|PP$|HOLU|HOLR|SLIP|RESH|RCHG|ROUT|JURY|BRVT'),0,OPS_SCI_TYPE,'ZZ' || OPS_SCI_TYPE), DECODE(OPS_SCI_STATUS,'SUB',1,'COM',2,3)) OVER (PARTITION BY OPS_SCI_ID) OPS_SCI_STATUS, " & _
								"MAX(OPS_SCI_NOTES) KEEP (DENSE_RANK LAST ORDER BY OPS_SCIH_INSERT_DATE, DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'^VAC|UN$|PT$|PP$|HOLU|HOLR|SLIP|RESH|RCHG|ROUT|JURY|BRVT'),0,OPS_SCI_TYPE,'ZZ' || OPS_SCI_TYPE), DECODE(OPS_SCI_STATUS,'SUB',1,'COM',2,3)) OVER (PARTITION BY OPS_SCI_ID) OPS_SCI_NOTES " & _
								"FROM OPS_SCHEDULE_INFO_HIST " & _
								"WHERE TO_DATE(OPS_SCI_START) = TO_DATE(?,'MM/DD/YYYY') " & _
								"AND OPS_SCI_OPS_USR_ID = ? " & _
								"AND ADD_MONTHS(TO_DATE(OPS_SCIH_INSERT_DATE),48) >= TO_DATE(OPS_SCI_START) " & _
								"AND OPS_SCIH_INSERT_DATE - (5/24) <= TO_DATE(?,'MM/DD/YYYY HH:MI:SS AM') " & _
							") " & _
							"LEFT JOIN RES_BUDGET_EXCEPTION " & _
							"ON TO_DATE(OPS_SCI_START) = RES_BUE_DATE " & _
							"AND RES_BUE_TYPE = 'NOR' " & _
							"WHERE OPS_SCI_STATUS IN ('APP','SUB') " & _
							"AND NOT " & _
							"( " & _
								"OPS_SCI_STATUS = 'SUB' " & _
								"AND TO_CHAR(OPS_SCI_START,'HH24:MI') = TO_CHAR(OPS_SCI_END,'HH24:MI') " & _
							") " & _
							"ORDER BY DECODE(OPS_SCI_STATUS,'APP',1,'SUB',2), CASE WHEN RES_BUE_ID IS NOT NULL AND OPS_SCI_TYPE IN ('HOLR','HOLU') THEN 1 ELSE 2 END, OPS_SCI_START, OPS_SCI_END, DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'UN$|PP$|PT$'),0,OPS_SCI_TYPE,'AA'||OPS_SCI_TYPE)"
							cmd.CommandText = SQLstmt
							cmd.Parameters(0).value = PARAMETER_DATE
							cmd.Parameters(1).value = PARAMETER_AGENT
							cmd.Parameters(2).value = RSHISTMASTER("MODIFY_TIME")
							Set RSHISTDETAIL = cmd.Execute
						%>
						<td <% If Not RSHISTDETAIL.EOF Then %> data-order="<%=RSHISTDETAIL("SCI_START")%>" <% Else %> data-order="" <% End If %>>
							<% If Not RSHISTDETAIL.EOF Then %>
								<table style="margin:auto;">
								<% Do While Not RSHISTDETAIL.EOF %>
									<tr class="<%=RSHISTDETAIL("SCHEDULE_CLASS")%>" title="<%=RSHISTDETAIL("OPS_SCI_NOTES")%>">
										<td class="subtable-td-padded-sm"><%=RSHISTDETAIL("OPS_SCI_TYPE")%></td>
										<td class="subtable-td-padded-sm"><%=RSHISTDETAIL("OPS_SCI_STATUS")%></td>
										<td class="subtable-td-padded-sm"><%=RSHISTDETAIL("SCI_START")%></td>
										<td class="subtable-td-padded-sm"><%=RSHISTDETAIL("SCI_END")%></td>
									</tr>
									<% If RSHISTDETAIL("GAP_FLAG") = "1" Then %>
										<tr style="line-height:0px;">
											<td class="subtable-td-padded-sm" colspan="4">&nbsp;</td>
										</tr>
									<% End If %>
									<% RSHISTDETAIL.MoveNext %>
								<% Loop %>
								</table>
							<% End If %>
							<% Set RSHISTDETAIL = Nothing %>
						</td>
					</tr>
					<% RSHISTMASTER.MoveNext %>
				<% Loop %>
			<% Else %>
				<tr>
					<td colspan="3">
						No entries found.
					</td>
				</tr>
			<% End If %>
			</tbody>
			<% Set RSHISTMASTER = Nothing %>
		</table>
	</div> 
	<script>
		$(document).ready(function() {
			<% If DATATABLES_BOOL = 1 Then %>
				$("#SCHEDULEHISTORY_TABLE").DataTable({
					"autoWidth": false,
					"paging": false,
					"searching": false,
					"info": false
				});
			<% End If %>
		});
	</script>
<!--#include file="pulsefunctions.asp"-->
<% Set cmd = Nothing %>
<% Conn.Close %>
<% Set Conn = Nothing %>