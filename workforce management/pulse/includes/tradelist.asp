<!--#include file="pulseheader.asp"-->
<!--#include file="tradesql.asp"-->
<%
	If Request.Querystring("DATE") <> "" Then
		PARAMETER_DATE = CDate(Request.Querystring("DATE"))
	Else
		PARAMETER_DATE = Date
	End If
		
	SQLstmt = "SELECT " & _
	"OPS_TRM_ID, " & _
	"OPS_TRM_OPS_USR_ID REQ_USR_ID, " & _
	"RAGT.OPS_USR_NAME REQUEST_AGENT, " & _
	"REQ_DATE REQUEST_DATE, " & _
	"OPS_TRD_ID, " & _
	"OPS_TRC_OPS_USR_ID ACC_USR_ID, " & _
	"AAGT.OPS_USR_NAME ACCEPT_AGENT, " & _
	"ACC_DATE ACCEPT_DATE, " & _
	"REQ_DURATION, " & _
	"DECODE(REQ_DURATION,ACC_DURATION,0,1) DURATION_FLAG, " & _
	"DECODE(REQ_PTO,0,0,1) REQ_PTO_FLAG, " & _
	"DECODE(ACC_PTO,0,0,1) ACC_PTO_FLAG, " & _
	"CASE WHEN REQ_SCI_START = OPS_TRM_START AND REQ_SCI_END = OPS_TRM_END THEN 0 ELSE 1 END TRM_REQUEST_ERROR_FLAG, " & _
	"CASE WHEN REQ_DATE <> ACC_DATE AND (RCHECK_SCI_START BETWEEN ACC_SCI_START AND ACC_SCI_END-(1/1440) OR RCHECK_SCI_END BETWEEN ACC_SCI_START+(1/1440) AND ACC_SCI_END) THEN 1 ELSE 0 END TRM_ACCEPT_ERROR_FLAG, " & _
	"CASE WHEN REQ_DATE <> ACC_DATE AND (ACHECK_SCI_START BETWEEN REQ_SCI_START AND REQ_SCI_END-(1/1440) OR ACHECK_SCI_END BETWEEN REQ_SCI_START+(1/1440) AND REQ_SCI_END) THEN 1 ELSE 0 END TRD_REQUEST_ERROR_FLAG, " & _
	"CASE WHEN ACC_SCI_START BETWEEN OPS_TRD_START AND OPS_TRD_END AND ACC_SCI_END BETWEEN OPS_TRD_START AND OPS_TRD_END THEN 0 ELSE 1 END TRD_ACCEPT_ERROR_FLAG, " & _
	"DECODE(REQ_DATE,ACC_DATE,1,0) SAME_DAY_FLAG, " & _
	"COUNT(*) OVER () TRADE_COUNT " & _
	"FROM OPS_TRADE_MASTER " & _
	"JOIN OPS_TRADE_DETAIL " & _
	"ON OPS_TRM_ID = OPS_TRD_OPS_TRM_ID " & _
	"JOIN OPS_TRADE_COMPLETE " & _
	"ON OPS_TRD_ID = OPS_TRC_OPS_TRD_ID " & _
	"JOIN " & _
	"( " & _
		"SELECT " & _
		"OPS_SCI_OPS_USR_ID, " & _
		"TO_DATE(OPS_SCI_START) REQ_DATE, " & _
		"MIN(OPS_SCI_START) REQ_SCI_START, " & _
		"MAX(OPS_SCI_END) REQ_SCI_END, " & _
		"ROUND(24*SUM(OPS_SCI_END - OPS_SCI_START),2) REQ_DURATION, " & _
		"ROUND(SUM(DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'^VAC|UN$|PT$|PP$|HOLU|SLIP|RESH|RCHG|ROUT|JURY|BRVT'),0,0,24*(OPS_SCI_END-OPS_SCI_START))),2) REQ_PTO " & _
		"FROM OPS_SCHEDULE_INFO " & _
		"WHERE OPS_SCI_STATUS = 'APP' " & _
		"AND OPS_SCI_TYPE NOT IN ('LNCH','LNFL','HOLR') " & _
		"GROUP BY OPS_SCI_OPS_USR_ID, TO_DATE(OPS_SCI_START) " & _
	") REQ " & _
	"ON OPS_TRM_OPS_USR_ID = REQ.OPS_SCI_OPS_USR_ID " & _
	"AND REQ_DATE = TO_DATE(OPS_TRM_START) " & _
	"LEFT JOIN " & _
	"( " & _
		"SELECT " & _
		"OPS_SCI_OPS_USR_ID, " & _
		"TO_DATE(OPS_SCI_START) RCHECK_DATE, " & _
		"MIN(OPS_SCI_START) RCHECK_SCI_START, " & _
		"MAX(OPS_SCI_END) RCHECK_SCI_END " & _
		"FROM OPS_SCHEDULE_INFO " & _
		"WHERE OPS_SCI_STATUS = 'APP' " & _
		"AND OPS_SCI_TYPE NOT IN ('LNCH','LNFL','HOLR') " & _
		"GROUP BY OPS_SCI_OPS_USR_ID, TO_DATE(OPS_SCI_START) " & _
	") RCHECK " & _
	"ON OPS_TRM_OPS_USR_ID = RCHECK.OPS_SCI_OPS_USR_ID " & _
	"AND RCHECK_DATE = TO_DATE(OPS_TRC_START) " & _
	"JOIN " & _
	"( " & _
		"SELECT " & _
		"OPS_SCI_OPS_USR_ID, " & _
		"TO_DATE(OPS_SCI_START) ACC_DATE, " & _
		"MIN(OPS_SCI_START) ACC_SCI_START, " & _
		"MAX(OPS_SCI_END) ACC_SCI_END, " & _
		"ROUND(24*SUM(OPS_SCI_END - OPS_SCI_START),2) ACC_DURATION, " & _
		"ROUND(SUM(DECODE(REGEXP_INSTR(OPS_SCI_TYPE,'^VAC|UN$|PT$|PP$|HOLU|SLIP|RESH|RCHG|ROUT|JURY|BRVT'),0,0,24*(OPS_SCI_END-OPS_SCI_START))),2) ACC_PTO " & _
		"FROM OPS_SCHEDULE_INFO " & _
		"WHERE OPS_SCI_STATUS = 'APP' " & _
		"AND OPS_SCI_TYPE NOT IN ('LNCH','LNFL','HOLR') " & _
		"GROUP BY OPS_SCI_OPS_USR_ID, TO_DATE(OPS_SCI_START) " & _
	") ACC " & _
	"ON OPS_TRC_OPS_USR_ID = ACC.OPS_SCI_OPS_USR_ID " & _
	"AND ACC_DATE = TO_DATE(OPS_TRC_START) " & _
	"LEFT JOIN " & _
	"( " & _
		"SELECT " & _
		"OPS_SCI_OPS_USR_ID, " & _
		"TO_DATE(OPS_SCI_START) ACHECK_DATE, " & _
		"MIN(OPS_SCI_START) ACHECK_SCI_START, " & _
		"MAX(OPS_SCI_END) ACHECK_SCI_END " & _
		"FROM OPS_SCHEDULE_INFO " & _
		"WHERE OPS_SCI_STATUS = 'APP' " & _
		"AND OPS_SCI_TYPE NOT IN ('LNCH','LNFL','HOLR') " & _
		"GROUP BY OPS_SCI_OPS_USR_ID, TO_DATE(OPS_SCI_START) " & _
	") ACHECK " & _
	"ON OPS_TRC_OPS_USR_ID = ACHECK.OPS_SCI_OPS_USR_ID " & _
	"AND ACHECK_DATE = TO_DATE(OPS_TRM_START) " & _
	"JOIN OPS_USER RAGT " & _
	"ON RAGT.OPS_USR_ID = OPS_TRM_OPS_USR_ID " & _
	"JOIN OPS_USER AAGT " & _
	"ON AAGT.OPS_USR_ID = OPS_TRC_OPS_USR_ID " & _
	"WHERE OPS_TRM_STATUS = 'PND' " & _
	"AND OPS_TRD_STATUS = 'PND' " & _
	"AND TO_DATE(OPS_TRM_START) > TO_DATE(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE)) " & _
	"AND TO_DATE(OPS_TRC_START) > TO_DATE(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE)) " & _
	"ORDER BY OPS_TRC_ID"
	Set RSTRADELIST = Conn.Execute(SQLstmt)
	
%>
	<form id="PULSE_FORM" data-request="TRADE" action="includes/formhandler.asp" method="post">
		<input type="hidden" id="NEWLINE_ID" value="0"/>
		<input type="hidden" id="SCHEDULEID_LIST" name="SCHEDULEID_LIST" value=""/>
		<input type="hidden" name="FORM_DATE" value="<%=PARAMETER_DATE%>"/>
		<div id="SCHEDULE_FORM_DIV" class="table-responsive" style="margin-bottom:1rem;">
			<table class="table table-bordered center" style="margin-bottom:0;">
				<caption class="th-color <% If PARAMETER_DATE = Date Then %>today-color-background<% Else %>past-color-background<% End If %>">
						Shift Trade List - <%=FormatDateTime(Now,3)%>
						<% 
							COUNT_TEXT = ""
							If Not RSTRADELIST.EOF Then
								If RSTRADELIST("TRADE_COUNT") = "1" Then
									COUNT_TEXT = "1 Trade Found"
								Else
									COUNT_TEXT = RSTRADELIST("TRADE_COUNT") & " Trades Found"
								End If
							End If
						%>
						<div style="float:right"><%=COUNT_TEXT%></div>
				</caption>
				<% If Not RSTRADELIST.EOF Then %>
					<% Dim TradeArray(1,99) %>
					<% TRADE_COUNTER = -1 %>
					<% Do While Not RSTRADELIST.EOF %>
					<% 
						DUPLICATE_TEXT = ""
						REQUEST_IN_DUP = 0
						REQUEST_OUT_DUP = 0
						ACCEPT_IN_DUP = 0
						ACCEPT_OUT_DUP = 0
						For i = 0 to TRADE_COUNTER
							If TradeArray(0,i)  = CDate(RSTRADELIST("REQUEST_DATE")) and TradeArray(1,i)  = RSTRADELIST("REQUEST_AGENT") Then
								REQUEST_OUT_DUP = 1
								DUPLICATE_TEXT = DUPLICATE_TEXT & RSTRADELIST("REQUEST_AGENT") & " has an overlapping trade for " & CDate(RSTRADELIST("REQUEST_DATE")) & ".<br/>"
							End If
							If TradeArray(0,i)  = CDate(RSTRADELIST("ACCEPT_DATE")) and TradeArray(1,i)  = RSTRADELIST("REQUEST_AGENT") Then
								REQUEST_IN_DUP = 1
								If RSTRADELIST("SAME_DAY_FLAG") = "0" Then
									DUPLICATE_TEXT = DUPLICATE_TEXT & RSTRADELIST("REQUEST_AGENT") & " has an overlapping trade for " & CDate(RSTRADELIST("ACCEPT_DATE")) & ".<br/>"
								End If
							End If
							If TradeArray(0,i)  = CDate(RSTRADELIST("ACCEPT_DATE")) and TradeArray(1,i)  = RSTRADELIST("ACCEPT_AGENT") Then
								ACCEPT_OUT_DUP = 1
								DUPLICATE_TEXT = DUPLICATE_TEXT & RSTRADELIST("ACCEPT_AGENT") & " has an overlapping trade for " & CDate(RSTRADELIST("ACCEPT_DATE")) & ".<br/>"
							End If								
							If TradeArray(0,i)  = CDate(RSTRADELIST("REQUEST_DATE")) and TradeArray(1,i)  = RSTRADELIST("ACCEPT_AGENT") Then
								ACCEPT_IN_DUP = 1
								If RSTRADELIST("SAME_DAY_FLAG") = "0" Then
									DUPLICATE_TEXT = DUPLICATE_TEXT & RSTRADELIST("ACCEPT_AGENT") & " has an overlapping trade for " & CDate(RSTRADELIST("REQUEST_DATE")) & ".<br/>"
								End If
							End If	
						Next
						If REQUEST_OUT_DUP = 0 Then 
							TRADE_COUNTER = TRADE_COUNTER + 1
							TradeArray(0,TRADE_COUNTER) = CDate(RSTRADELIST("REQUEST_DATE"))
							TradeArray(1,TRADE_COUNTER) = RSTRADELIST("REQUEST_AGENT")
						End If
						If REQUEST_IN_DUP = 0 Then 
							TRADE_COUNTER = TRADE_COUNTER + 1
							TradeArray(0,TRADE_COUNTER) = CDate(RSTRADELIST("ACCEPT_DATE"))
							TradeArray(1,TRADE_COUNTER) = RSTRADELIST("REQUEST_AGENT")
						End If
						If ACCEPT_OUT_DUP = 0 Then
							TRADE_COUNTER = TRADE_COUNTER + 1
							TradeArray(0,TRADE_COUNTER) = CDate(RSTRADELIST("ACCEPT_DATE"))
							TradeArray(1,TRADE_COUNTER) = RSTRADELIST("ACCEPT_AGENT")
						End If
						If ACCEPT_IN_DUP = 0 Then
							TRADE_COUNTER = TRADE_COUNTER + 1
							TradeArray(0,TRADE_COUNTER) = CDate(RSTRADELIST("REQUEST_DATE"))
							TradeArray(1,TRADE_COUNTER) = RSTRADELIST("ACCEPT_AGENT")
						End If
					%>		
						<tbody id="TRADETBODY_<%=RSTRADELIST("OPS_TRM_ID")%>">
							<tr>
								<td colspan="5" style="text-align:left;" <% If RSTRADELIST("DURATION_FLAG") = "1" or RSTRADELIST("REQ_PTO_FLAG") = "1" or RSTRADELIST("ACC_PTO_FLAG") = "1" or RSTRADELIST("TRM_REQUEST_ERROR_FLAG") = "1" or RSTRADELIST("TRM_ACCEPT_ERROR_FLAG") = "1" or RSTRADELIST("TRD_REQUEST_ERROR_FLAG") = "1" or RSTRADELIST("TRD_ACCEPT_ERROR_FLAG") = "1" or DUPLICATE_TEXT <> "" Then %>class="error-color-background"<% End If %>>
									<div style="display:inline-block;width:30%;">&nbsp;</div>
									<div style="vertical-align:middle;display:inline-block;width:15%;text-align:center;">
										<div><%=RSTRADELIST("REQUEST_AGENT")%></div>
										<div>(<%=RSTRADELIST("REQUEST_DATE")%>)</div>
									</div>
									<div style="vertical-align:middle;display:inline-block;width:10%;text-align:center;">
										<div style="font-size:1.75em;"><i class="fas fa-exchange-alt"></i></div>
										<div style="margin-top:-10px;font-size:.7em;"><%=RSTRADELIST("REQ_DURATION")%> Hours</div>							
									</div>
									<div style="vertical-align:middle;display:inline-block;width:15%;text-align:center;">
										<div><%=RSTRADELIST("ACCEPT_AGENT")%></div>
										<div>(<%=RSTRADELIST("ACCEPT_DATE")%>)</div>							
									</div>
									<div class="error-color" style="margin-top:15px;text-align:center;">
										<% If RSTRADELIST("DURATION_FLAG") = "1" Then %>
											Trade hours do not match.<br/>
										<% End If %>
										<% If RSTRADELIST("REQ_PTO_FLAG") = "1" Then %>
											<%=RSTRADELIST("REQUEST_AGENT")%> has an invalid schedule code.<br/>
										<% End If %>
										<% If RSTRADELIST("ACC_PTO_FLAG") = "1" Then %>
											<%=RSTRADELIST("ACCEPT_AGENT")%> has an invalid schedule code.<br/>
										<% End If %>									
										<% If RSTRADELIST("TRM_REQUEST_ERROR_FLAG") = "1" Then %>
											<%=RSTRADELIST("REQUEST_AGENT")%> modified their schedule after the trade was accepted.<br/>
										<% End If %>
										<% If RSTRADELIST("TRM_ACCEPT_ERROR_FLAG") = "1" Then %>
											<%=RSTRADELIST("REQUEST_AGENT")%>'s schedule overlaps with <%=RSTRADELIST("ACCEPT_AGENT")%>'s on <%=RSTRADELIST("ACCEPT_DATE")%>.<br/>
										<% End If %>
										<% If RSTRADELIST("TRD_REQUEST_ERROR_FLAG") = "1" Then %>
											<%=RSTRADELIST("ACCEPT_AGENT")%>'s schedule overlaps with <%=RSTRADELIST("REQUEST_AGENT")%>'s on <%=RSTRADELIST("REQUEST_DATE")%>.<br/>
										<% End If %>
										<% If RSTRADELIST("TRD_ACCEPT_ERROR_FLAG") = "1" Then %>
											<%=RSTRADELIST("ACCEPT_AGENT")%> modified their schedule to be outside the valid trade window.<br/>
										<% End If %>	
										<%=DUPLICATE_TEXT%>
									</div>
								</td>
							</tr>
							<tr>
								<td class="th-color <% If PARAMETER_DATE = Date Then %>today-color-background<% Else %>past-color-background<% End If %>">
									Date
								</td>
								<td colspan="2" class="th-color <% If PARAMETER_DATE = Date Then %>today-color-background<% Else %>past-color-background<% End If %>">
									Pre-Trade
								</td>
								<td colspan="2" class="th-color <% If PARAMETER_DATE = Date Then %>today-color-background<% Else %>past-color-background<% End If %>">
									Post-Trade
								</td>
							</tr>
							<tr>
								<td>
									&nbsp;
								</td>
								<td>
									<%=RSTRADELIST("REQUEST_AGENT")%>
								</td>
								<td>
									<%=RSTRADELIST("ACCEPT_AGENT")%>
								</td>
								<td>
									<%=RSTRADELIST("REQUEST_AGENT")%>
								</td>
								<td>
									<%=RSTRADELIST("ACCEPT_AGENT")%>
								</td>
							</tr>
							<tr>
								<td>
									<%=RSTRADELIST("REQUEST_DATE")%>
								</td>
								<td>
								<% 
									cmd.CommandText = PulsePreTradeSQL
									cmd.Parameters(0).value = CDate(RSTRADELIST("REQUEST_DATE"))
									cmd.Parameters(1).value = RSTRADELIST("REQ_USR_ID")
									Set RSSHIFT = cmd.Execute
								%>
									<% If Not RSSHIFT.EOF Then %>
										<table style="margin:auto;">
										<% Do While Not RSSHIFT.EOF %>
											<tr class="<%=RSSHIFT("SCHEDULE_CLASS")%>" title="<%=RSSHIFT("OPS_SCI_NOTES")%>">
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
											<% RSSHIFT.MoveNext %>
										<% Loop %>
										</table>
									<% End If %>
									<% Set RSSHIFT = Nothing %>	
								</td>
								<td>
								<% 
									cmd.CommandText = PulsePreTradeSQL
									cmd.Parameters(0).value = CDate(RSTRADELIST("REQUEST_DATE"))
									cmd.Parameters(1).value = RSTRADELIST("ACC_USR_ID")
									Set RSSHIFT = cmd.Execute
								%>
									<% If Not RSSHIFT.EOF Then %>
										<table style="margin:auto;">
										<% Do While Not RSSHIFT.EOF %>
											<tr class="<%=RSSHIFT("SCHEDULE_CLASS")%>" title="<%=RSSHIFT("OPS_SCI_NOTES")%>">
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
											<% RSSHIFT.MoveNext %>
										<% Loop %>
										</table>
									<% End If %>
									<% Set RSSHIFT = Nothing %>	
								</td>
								<td id="AGENTROW_<%=RSTRADELIST("REQ_USR_ID")%>_<%=Right("0" & Month(CDate(RSTRADELIST("REQUEST_DATE"))),2) & Right("0" & Day(CDate(RSTRADELIST("REQUEST_DATE"))),2) & Year(CDate(RSTRADELIST("REQUEST_DATE")))%>">
									<% If RSTRADELIST("SAME_DAY_FLAG") = "1" Then %>
									<% 
										cmd.CommandText = PulsePostTradeInSQL
										cmd.Parameters(0).value = RSTRADELIST("SAME_DAY_FLAG")
										cmd.Parameters(1).value = CDate(RSTRADELIST("ACCEPT_DATE"))
										cmd.Parameters(2).value = RSTRADELIST("REQ_USR_ID")
										cmd.Parameters(3).value = AgentName(RSTRADELIST("ACC_USR_ID"))
										cmd.Parameters(4).value = CDate(RSTRADELIST("REQUEST_DATE"))
										cmd.Parameters(5).value = CDate(RSTRADELIST("ACCEPT_DATE"))
										cmd.Parameters(6).value = RSTRADELIST("ACC_USR_ID")
										Set RSSHIFT = cmd.Execute
									%>
										<% If Not RSSHIFT.EOF Then %>
											<table <% If DUPLICATE_TEXT = "" Then %> id="TRADETABLE_<%=RSTRADELIST("REQ_USR_ID")%>_<%=Right("0" & Month(CDate(RSTRADELIST("ACCEPT_DATE"))),2) & Right("0" & Day(CDate(RSTRADELIST("ACCEPT_DATE"))),2) & Year(CDate(RSTRADELIST("ACCEPT_DATE")))%>" <% End If %> style="margin:auto;">
											<% Do While Not RSSHIFT.EOF %>
												<tr class="<%=RSSHIFT("SCHEDULE_CLASS")%>" title="<%=RSSHIFT("OPS_SCI_NOTES")%>" <% If RSSHIFT("OPS_SCI_STATUS") = "DEL" Then %> style="display:none;"<% ElseIf RSSHIFT("TRADE_FLAG") = "1" Then %>style="font-style:italic;font-size:.8rem;"<% End If %> data-sciid="<%=RSSHIFT("OPS_SCI_ID")%>" data-scitype="<%=RSSHIFT("OPS_SCI_TYPE")%>" data-scistatus="<%=RSSHIFT("OPS_SCI_STATUS")%>" data-scistart="<%=RSSHIFT("SCI_START")%>" data-sciend="<%=RSSHIFT("SCI_END")%>" data-scinotes="<%=RSSHIFT("OPS_SCI_NOTES")%>">
													<td class="subtable-td-padded-sm"><%=RSSHIFT("OPS_SCI_TYPE")%></td>
													<td class="subtable-td-padded-sm"><%=RSSHIFT("OPS_SCI_STATUS")%></td>
													<td class="subtable-td-padded-sm"><%=RSSHIFT("SCI_START")%></td>
													<td class="subtable-td-padded-sm"><%=Replace(RSSHIFT("SCI_END"),"24:00","00:00")%></td>
												</tr>
												<% If RSSHIFT("GAP_FLAG") = "1" Then %>
													<tr style="line-height:0px;">
														<td class="subtable-td-padded-sm" colspan="4">&nbsp;</td>
													</tr>
												<% End If %>
												<% RSSHIFT.MoveNext %>
											<% Loop %>
											</table>
										<% End If %>
										<% Set RSSHIFT = Nothing %>	
									<% Else %>
									<% 
										cmd.CommandText = PulsePostTradeOutSQL
										cmd.Parameters(0).value = CDate(RSTRADELIST("REQUEST_DATE"))
										cmd.Parameters(1).value = RSTRADELIST("REQ_USR_ID")
										Set RSSHIFT = cmd.Execute
									%>
										<% If Not RSSHIFT.EOF Then %>
											<table <% If DUPLICATE_TEXT = "" Then %> id="TRADETABLE_<%=RSTRADELIST("REQ_USR_ID")%>_<%=Right("0" & Month(CDate(RSTRADELIST("REQUEST_DATE"))),2) & Right("0" & Day(CDate(RSTRADELIST("REQUEST_DATE"))),2) & Year(CDate(RSTRADELIST("REQUEST_DATE")))%>" <% End If %> style="margin:auto;">
											<% Do While Not RSSHIFT.EOF %>
												<tr class="<%=RSSHIFT("SCHEDULE_CLASS")%>" title="<%=RSSHIFT("OPS_SCI_NOTES")%>" <% If RSSHIFT("OPS_SCI_STATUS") = "DEL" Then %> style="display:none;"<% End If %> data-sciid="<%=RSSHIFT("OPS_SCI_ID")%>" data-scitype="<%=RSSHIFT("OPS_SCI_TYPE")%>" data-scistatus="<%=RSSHIFT("OPS_SCI_STATUS")%>" data-scistart="<%=RSSHIFT("SCI_START")%>" data-sciend="<%=RSSHIFT("SCI_END")%>" data-scinotes="<%=RSSHIFT("OPS_SCI_NOTES")%>">
													<td class="subtable-td-padded-sm"><%=RSSHIFT("OPS_SCI_TYPE")%></td>
													<td class="subtable-td-padded-sm"><%=RSSHIFT("OPS_SCI_STATUS")%></td>
													<td class="subtable-td-padded-sm"><%=RSSHIFT("SCI_START")%></td>
													<td class="subtable-td-padded-sm"><%=Replace(RSSHIFT("SCI_END"),"24:00","00:00")%></td>
												</tr>
												<% If RSSHIFT("GAP_FLAG") = "1" Then %>
													<tr style="line-height:0px;">
														<td class="subtable-td-padded-sm" colspan="4">&nbsp;</td>
													</tr>
												<% End If %>
												<% RSSHIFT.MoveNext %>
											<% Loop %>
											</table>
										<% End If %>
										<% Set RSSHIFT = Nothing %>	
									<% End If %>
								</td>
								<td id="AGENTROW_<%=RSTRADELIST("ACC_USR_ID")%>_<%=Right("0" & Month(CDate(RSTRADELIST("REQUEST_DATE"))),2) & Right("0" & Day(CDate(RSTRADELIST("REQUEST_DATE"))),2) & Year(CDate(RSTRADELIST("REQUEST_DATE")))%>">
									<% 
										cmd.CommandText = PulsePostTradeInSQL
										cmd.Parameters(0).value = RSTRADELIST("SAME_DAY_FLAG")
										cmd.Parameters(1).value = CDate(RSTRADELIST("REQUEST_DATE"))
										cmd.Parameters(2).value = RSTRADELIST("ACC_USR_ID")
										cmd.Parameters(3).value = AgentName(RSTRADELIST("REQ_USR_ID"))
										cmd.Parameters(4).value = CDate(RSTRADELIST("ACCEPT_DATE"))
										cmd.Parameters(5).value = CDate(RSTRADELIST("REQUEST_DATE"))
										cmd.Parameters(6).value = RSTRADELIST("REQ_USR_ID")
										Set RSSHIFT = cmd.Execute
									%>
									<% If Not RSSHIFT.EOF Then %>
										<table <% If DUPLICATE_TEXT = "" Then %> id="TRADETABLE_<%=RSTRADELIST("ACC_USR_ID")%>_<%=Right("0" & Month(CDate(RSTRADELIST("REQUEST_DATE"))),2) & Right("0" & Day(CDate(RSTRADELIST("REQUEST_DATE"))),2) & Year(CDate(RSTRADELIST("REQUEST_DATE")))%>" <% End If %> style="margin:auto;">
										<% Do While Not RSSHIFT.EOF %>
											<tr class="<%=RSSHIFT("SCHEDULE_CLASS")%>" title="<%=RSSHIFT("OPS_SCI_NOTES")%>" <% If RSSHIFT("OPS_SCI_STATUS") = "DEL" Then %> style="display:none;"<% ElseIf RSSHIFT("TRADE_FLAG") = "1" Then %>style="font-style:italic;font-size:.8rem;"<% End If %> data-sciid="<%=RSSHIFT("OPS_SCI_ID")%>" data-scitype="<%=RSSHIFT("OPS_SCI_TYPE")%>" data-scistatus="<%=RSSHIFT("OPS_SCI_STATUS")%>" data-scistart="<%=RSSHIFT("SCI_START")%>" data-sciend="<%=RSSHIFT("SCI_END")%>" data-scinotes="<%=RSSHIFT("OPS_SCI_NOTES")%>">
												<td class="subtable-td-padded-sm"><%=RSSHIFT("OPS_SCI_TYPE")%></td>
												<td class="subtable-td-padded-sm"><%=RSSHIFT("OPS_SCI_STATUS")%></td>
												<td class="subtable-td-padded-sm"><%=RSSHIFT("SCI_START")%></td>
												<td class="subtable-td-padded-sm"><%=Replace(RSSHIFT("SCI_END"),"24:00","00:00")%></td>
											</tr>
											<% If RSSHIFT("GAP_FLAG") = "1" Then %>
												<tr style="line-height:0px;">
													<td class="subtable-td-padded-sm" colspan="4">&nbsp;</td>
												</tr>
											<% End If %>
											<% RSSHIFT.MoveNext %>
										<% Loop %>
										</table>
									<% End If %>
									<% Set RSSHIFT = Nothing %>	
								</td>						
							</tr>
							<% If RSTRADELIST("SAME_DAY_FLAG") <> "1" Then %>
								<tr>
									<td>
										<%=RSTRADELIST("ACCEPT_DATE")%>
									</td>
									<td>
										<% 
											cmd.CommandText = PulsePreTradeSQL
											cmd.Parameters(0).value = CDate(RSTRADELIST("ACCEPT_DATE"))
											cmd.Parameters(1).value = RSTRADELIST("REQ_USR_ID")
											Set RSSHIFT = cmd.Execute
										%>
										<% If Not RSSHIFT.EOF Then %>
											<table style="margin:auto;">
											<% Do While Not RSSHIFT.EOF %>
												<tr class="<%=RSSHIFT("SCHEDULE_CLASS")%>" title="<%=RSSHIFT("OPS_SCI_NOTES")%>">
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
												<% RSSHIFT.MoveNext %>
											<% Loop %>
											</table>
										<% End If %>
										<% Set RSSHIFT = Nothing %>	
									</td>
									<td>
										<% 
											cmd.CommandText = PulsePreTradeSQL
											cmd.Parameters(0).value = CDate(RSTRADELIST("ACCEPT_DATE"))
											cmd.Parameters(1).value = RSTRADELIST("ACC_USR_ID")
											Set RSSHIFT = cmd.Execute
										%>
										<% If Not RSSHIFT.EOF Then %>
											<table style="margin:auto;">
											<% Do While Not RSSHIFT.EOF %>
												<tr class="<%=RSSHIFT("SCHEDULE_CLASS")%>" title="<%=RSSHIFT("OPS_SCI_NOTES")%>">
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
												<% RSSHIFT.MoveNext %>
											<% Loop %>
											</table>
										<% End If %>
										<% Set RSSHIFT = Nothing %>	
									</td>
									<td id="AGENTROW_<%=RSTRADELIST("REQ_USR_ID")%>_<%=Right("0" & Month(CDate(RSTRADELIST("ACCEPT_DATE"))),2) & Right("0" & Day(CDate(RSTRADELIST("ACCEPT_DATE"))),2) & Year(CDate(RSTRADELIST("ACCEPT_DATE")))%>">
										<% 
											cmd.CommandText = PulsePostTradeInSQL
											cmd.Parameters(0).value = RSTRADELIST("SAME_DAY_FLAG")
											cmd.Parameters(1).value = CDate(RSTRADELIST("ACCEPT_DATE"))
											cmd.Parameters(2).value = RSTRADELIST("REQ_USR_ID")
											cmd.Parameters(3).value = AgentName(RSTRADELIST("ACC_USR_ID"))
											cmd.Parameters(4).value = CDate(RSTRADELIST("REQUEST_DATE"))
											cmd.Parameters(5).value = CDate(RSTRADELIST("ACCEPT_DATE"))
											cmd.Parameters(6).value = RSTRADELIST("ACC_USR_ID")
											Set RSSHIFT = cmd.Execute
										%>
										<% If Not RSSHIFT.EOF Then %>
											<table <% If DUPLICATE_TEXT = "" Then %> id="TRADETABLE_<%=RSTRADELIST("REQ_USR_ID")%>_<%=Right("0" & Month(CDate(RSTRADELIST("ACCEPT_DATE"))),2) & Right("0" & Day(CDate(RSTRADELIST("ACCEPT_DATE"))),2) & Year(CDate(RSTRADELIST("ACCEPT_DATE")))%>" <% End If %> style="margin:auto;">
											<% Do While Not RSSHIFT.EOF %>
												<tr class="<%=RSSHIFT("SCHEDULE_CLASS")%>" title="<%=RSSHIFT("OPS_SCI_NOTES")%>" <% If RSSHIFT("OPS_SCI_STATUS") = "DEL" Then %> style="display:none;"<% ElseIf RSSHIFT("TRADE_FLAG") = "1" Then %>style="font-style:italic;font-size:.8rem;"<% End If %> data-sciid="<%=RSSHIFT("OPS_SCI_ID")%>" data-scitype="<%=RSSHIFT("OPS_SCI_TYPE")%>" data-scistatus="<%=RSSHIFT("OPS_SCI_STATUS")%>" data-scistart="<%=RSSHIFT("SCI_START")%>" data-sciend="<%=RSSHIFT("SCI_END")%>" data-scinotes="<%=RSSHIFT("OPS_SCI_NOTES")%>">
													<td class="subtable-td-padded-sm"><%=RSSHIFT("OPS_SCI_TYPE")%></td>
													<td class="subtable-td-padded-sm"><%=RSSHIFT("OPS_SCI_STATUS")%></td>
													<td class="subtable-td-padded-sm"><%=RSSHIFT("SCI_START")%></td>
													<td class="subtable-td-padded-sm"><%=Replace(RSSHIFT("SCI_END"),"24:00","00:00")%></td>
												</tr>
												<% If RSSHIFT("GAP_FLAG") = "1" Then %>
													<tr style="line-height:0px;">
														<td class="subtable-td-padded-sm" colspan="4">&nbsp;</td>
													</tr>
												<% End If %>
												<% RSSHIFT.MoveNext %>
											<% Loop %>
											</table>
										<% End If %>
										<% Set RSSHIFT = Nothing %>	
									</td>
									<td id="AGENTROW_<%=RSTRADELIST("ACC_USR_ID")%>_<%=Right("0" & Month(CDate(RSTRADELIST("ACCEPT_DATE"))),2) & Right("0" & Day(CDate(RSTRADELIST("ACCEPT_DATE"))),2) & Year(CDate(RSTRADELIST("ACCEPT_DATE")))%>">
									<% 
										cmd.CommandText = PulsePostTradeOutSQL
										cmd.Parameters(0).value = CDate(RSTRADELIST("ACCEPT_DATE"))
										cmd.Parameters(1).value = RSTRADELIST("ACC_USR_ID")
										Set RSSHIFT = cmd.Execute
									%>
										<% If Not RSSHIFT.EOF Then %>
											<table <% If DUPLICATE_TEXT = "" Then %> id="TRADETABLE_<%=RSTRADELIST("ACC_USR_ID")%>_<%=Right("0" & Month(CDate(RSTRADELIST("ACCEPT_DATE"))),2) & Right("0" & Day(CDate(RSTRADELIST("ACCEPT_DATE"))),2) & Year(CDate(RSTRADELIST("ACCEPT_DATE")))%>" <% End If %> style="margin:auto;">
											<% Do While Not RSSHIFT.EOF %>
												<tr class="<%=RSSHIFT("SCHEDULE_CLASS")%>" title="<%=RSSHIFT("OPS_SCI_NOTES")%>" <% If RSSHIFT("OPS_SCI_STATUS") = "DEL" Then %> style="display:none;"<% End If %> data-sciid="<%=RSSHIFT("OPS_SCI_ID")%>" data-scitype="<%=RSSHIFT("OPS_SCI_TYPE")%>" data-scistatus="<%=RSSHIFT("OPS_SCI_STATUS")%>" data-scistart="<%=RSSHIFT("SCI_START")%>" data-sciend="<%=RSSHIFT("SCI_END")%>" data-scinotes="<%=RSSHIFT("OPS_SCI_NOTES")%>">
													<td class="subtable-td-padded-sm"><%=RSSHIFT("OPS_SCI_TYPE")%></td>
													<td class="subtable-td-padded-sm"><%=RSSHIFT("OPS_SCI_STATUS")%></td>
													<td class="subtable-td-padded-sm"><%=RSSHIFT("SCI_START")%></td>
													<td class="subtable-td-padded-sm"><%=Replace(RSSHIFT("SCI_END"),"24:00","00:00")%></td>
												</tr>
												<% If RSSHIFT("GAP_FLAG") = "1" Then %>
													<tr style="line-height:0px;">
														<td class="subtable-td-padded-sm" colspan="4">&nbsp;</td>
													</tr>
												<% End If %>
												<% RSSHIFT.MoveNext %>
											<% Loop %>
											</table>
										<% End If %>
										<% Set RSSHIFT = Nothing %>	
									</td>						
								</tr>
							<% End If %>
							<tr>
								<td colspan="5">
									<button id="TRADEDELBUTTON_<%=RSTRADELIST("OPS_TRM_ID")%>" type="button" class="trade-button btn <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" style="background-color:#fff;" title="Delete">
										<i class="fas fa-trash"></i>
									</button>
									<input type="checkbox" id="TRADEDELCHECK_<%=RSTRADELIST("OPS_TRM_ID")%>" name="TRADE_ACTION" value="<%=RSTRADELIST("OPS_TRM_ID")%>_<%=RSTRADELIST("OPS_TRD_ID")%>_DEL" style="display:none;" />
									<button id="TRADEDNYBUTTON_<%=RSTRADELIST("OPS_TRM_ID")%>" type="button" class="trade-button btn <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" style="background-color:#fff;" title="Deny">
										<i class="fas fa-ban"></i>
									</button>
									<input type="checkbox" id="TRADEDNYCHECK_<%=RSTRADELIST("OPS_TRM_ID")%>" name="TRADE_ACTION" value="<%=RSTRADELIST("OPS_TRM_ID")%>_<%=RSTRADELIST("OPS_TRD_ID")%>_REQ" style="display:none;" />
									<button id="TRADECOMBUTTON_<%=RSTRADELIST("OPS_TRM_ID")%>" type="button" class="trade-button btn <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" style="background-color:#fff;<% If DUPLICATE_TEXT <> "" Then %>visibility:hidden;<% End If %>" title="Approve">
										<i class="fas fa-thumbs-up"></i>
									</button>
									<input type="checkbox" id="TRADECOMCHECK_<%=RSTRADELIST("OPS_TRM_ID")%>" name="TRADE_ACTION" value="<%=RSTRADELIST("OPS_TRM_ID")%>_<%=RSTRADELIST("OPS_TRD_ID")%>_COM" style="display:none;" />
								</td>
							</tr>
							<tr id="TRADEDENIAL_<%=RSTRADELIST("OPS_TRM_ID")%>" style="display:none;">
								<td class="<% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" colspan="5">
									Reason for denial: <input id="TRADETEXT_<%=RSTRADELIST("OPS_TRM_ID")%>" name="TRADETEXT_<%=RSTRADELIST("OPS_TRM_ID")%>" type="text" class="<% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" style="width:50%;" />
								</td>
							</tr>
							<% If DUPLICATE_TEXT = "" Then %>
								<tr id="EDITROW_<%=RSTRADELIST("REQ_USR_ID")%>_<%=Right("0" & Month(CDate(RSTRADELIST("REQUEST_DATE"))),2) & Right("0" & Day(CDate(RSTRADELIST("REQUEST_DATE"))),2) & Year(CDate(RSTRADELIST("REQUEST_DATE")))%>" style="display:none;" data-trade="<%=RSTRADELIST("OPS_TRM_ID")%>">
									<td id="EDITDIV_WRAPPER_<%=RSTRADELIST("REQ_USR_ID")%>_<%=Right("0" & Month(CDate(RSTRADELIST("REQUEST_DATE"))),2) & Right("0" & Day(CDate(RSTRADELIST("REQUEST_DATE"))),2) & Year(CDate(RSTRADELIST("REQUEST_DATE")))%>" colspan="5">
										<div id="EDITDIV_<%=RSTRADELIST("REQ_USR_ID")%>_<%=Right("0" & Month(CDate(RSTRADELIST("REQUEST_DATE"))),2) & Right("0" & Day(CDate(RSTRADELIST("REQUEST_DATE"))),2) & Year(CDate(RSTRADELIST("REQUEST_DATE")))%>"></div>
									</td>
								</tr>
								<tr id="EDITROW_<%=RSTRADELIST("ACC_USR_ID")%>_<%=Right("0" & Month(CDate(RSTRADELIST("REQUEST_DATE"))),2) & Right("0" & Day(CDate(RSTRADELIST("REQUEST_DATE"))),2) & Year(CDate(RSTRADELIST("REQUEST_DATE")))%>" style="display:none;" data-trade="<%=RSTRADELIST("OPS_TRM_ID")%>">
									<td id="EDITDIV_WRAPPER_<%=RSTRADELIST("ACC_USR_ID")%>_<%=Right("0" & Month(CDate(RSTRADELIST("REQUEST_DATE"))),2) & Right("0" & Day(CDate(RSTRADELIST("REQUEST_DATE"))),2) & Year(CDate(RSTRADELIST("REQUEST_DATE")))%>" colspan="5">
										<div id="EDITDIV_<%=RSTRADELIST("ACC_USR_ID")%>_<%=Right("0" & Month(CDate(RSTRADELIST("REQUEST_DATE"))),2) & Right("0" & Day(CDate(RSTRADELIST("REQUEST_DATE"))),2) & Year(CDate(RSTRADELIST("REQUEST_DATE")))%>"></div>
									</td>
								</tr>
								<% If RSTRADELIST("SAME_DAY_FLAG") <> "1" Then %>
									<tr id="EDITROW_<%=RSTRADELIST("REQ_USR_ID")%>_<%=Right("0" & Month(CDate(RSTRADELIST("ACCEPT_DATE"))),2) & Right("0" & Day(CDate(RSTRADELIST("ACCEPT_DATE"))),2) & Year(CDate(RSTRADELIST("ACCEPT_DATE")))%>" style="display:none;" data-trade="<%=RSTRADELIST("OPS_TRM_ID")%>">
										<td id="EDITDIV_WRAPPER_<%=RSTRADELIST("REQ_USR_ID")%>_<%=Right("0" & Month(CDate(RSTRADELIST("ACCEPT_DATE"))),2) & Right("0" & Day(CDate(RSTRADELIST("ACCEPT_DATE"))),2) & Year(CDate(RSTRADELIST("ACCEPT_DATE")))%>" colspan="5">
											<div id="EDITDIV_<%=RSTRADELIST("REQ_USR_ID")%>_<%=Right("0" & Month(CDate(RSTRADELIST("ACCEPT_DATE"))),2) & Right("0" & Day(CDate(RSTRADELIST("ACCEPT_DATE"))),2) & Year(CDate(RSTRADELIST("ACCEPT_DATE")))%>"></div>
										</td>
									</tr>
									<tr id="EDITROW_<%=RSTRADELIST("ACC_USR_ID")%>_<%=Right("0" & Month(CDate(RSTRADELIST("ACCEPT_DATE"))),2) & Right("0" & Day(CDate(RSTRADELIST("ACCEPT_DATE"))),2) & Year(CDate(RSTRADELIST("ACCEPT_DATE")))%>" style="display:none;" data-trade="<%=RSTRADELIST("OPS_TRM_ID")%>">
										<td id="EDITDIV_WRAPPER_<%=RSTRADELIST("ACC_USR_ID")%>_<%=Right("0" & Month(CDate(RSTRADELIST("ACCEPT_DATE"))),2) & Right("0" & Day(CDate(RSTRADELIST("ACCEPT_DATE"))),2) & Year(CDate(RSTRADELIST("ACCEPT_DATE")))%>" colspan="5">
											<div id="EDITDIV_<%=RSTRADELIST("ACC_USR_ID")%>_<%=Right("0" & Month(CDate(RSTRADELIST("ACCEPT_DATE"))),2) & Right("0" & Day(CDate(RSTRADELIST("ACCEPT_DATE"))),2) & Year(CDate(RSTRADELIST("ACCEPT_DATE")))%>"></div>
										</td>
									</tr>						
								<% End If %>
							<% End If %>
						</tbody>
						<% RSTRADELIST.MoveNext %>
					<% Loop %>
					<% Erase TradeArray %>
					<tr>
						<td colspan="5">
							<input id="PULSE_SUBMIT" type="submit" class="btn th-color <% If PARAMETER_DATE = Date Then %>today-color-background<% Else %>past-color-background<% End If %>" value="Submit Changes"/>
							<div id="OVERLAP_MESSAGE" class="error-color" style="display:none;">
								Fix overlapping schedule entries before submitting.
							</div>
						</td>
					</tr>
				<% Else %>
					<tr>
						<td colspan="5">
							No trades found.
						</td>
					</tr>
				<% End If %>
			</table>
		</div>
	</form>
<script>
	$(document).ready(function() {
		overlapIdList = [];
		scheduleIdList = [];
	});
</script>
<!--#include file="pulsefunctions.asp"-->
<% Set cmd = Nothing %>
<% Conn.Close %>
<% Set Conn = Nothing %>