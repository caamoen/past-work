<!--#include file="pulseheader.asp"-->

<%
	If Request.Querystring("DATE") <> "" Then
		PARAMETER_DATE = CDate(Request.Querystring("DATE"))
	Else
		PARAMETER_DATE = Date
	End If
	If Request.Querystring("AGENT") <> "" Then
		PARAMETER_AGENT = Request.Querystring("AGENT")
	Else
		PARAMETER_AGENT = ""
	End If
	If Request.Querystring("LOCATION") <> "" Then
		PARAMETER_LOCATION = Request.Querystring("LOCATION")
	Else
		PARAMETER_LOCATION = ""
	End If
	If Request.Querystring("DEPARTMENT") <> "" Then
		PARAMETER_DEPARTMENT = Request.Querystring("DEPARTMENT")
	Else
		PARAMETER_DEPARTMENT = ""
	End If
	If Request.Querystring("TEAM") <> "" Then
		PARAMETER_TEAM = Request.Querystring("TEAM")
	Else
		PARAMETER_TEAM = ""
	End If
	If Request.Querystring("JOB") <> "" Then
		PARAMETER_JOB = Request.Querystring("JOB")
	Else
		PARAMETER_JOB = ""
	End If
	SQLstmt = "SELECT " & _
	"PAGE.TYPE_ID, " & _
	"PAGE.ACCESS_ID, " & _
	"PAGE.ACCESS_TYPE, " & _
	"PAGE.ACCESS_NAME, " & _
	"PAGE.ACCESS_ADDRESS, " & _
	"NVL2(SECURITY.ACCESS_ID,1,0) ACCESS_FLAG " & _
	"FROM " & _
	"( " & _
		"SELECT " & _
		"50 TYPE_ID, " & _
		"ACT.SYS_CDD_ID ACCESS_ID, " & _
		"DECODE(ACT.SYS_CDD_SYS_CDM_ID,45,'Department',47,'Page','Application') ACCESS_TYPE, " & _
		"ACT.SYS_CDD_NAME ACCESS_NAME, " & _
		"ACT.SYS_CDD_VALUE ACCESS_ADDRESS " & _
		"FROM SYS_CODE_DETAIL ACT " & _
		"LEFT JOIN SYS_CODE_DETAIL ARC " & _
		"ON ARC.SYS_CDD_SYS_CDM_ID = 497 " & _
		"AND ACT.SYS_CDD_VALUE = ARC.SYS_CDD_VALUE " & _
		"WHERE ACT.SYS_CDD_SYS_CDM_ID IN (45,46,47) " & _
		"AND ARC.SYS_CDD_ID IS NULL " & _
		"UNION ALL " & _
		"SELECT " & _
		"132, " & _
		"OPS_RPM_ID, " & _
		"'Report', " & _
		"OPS_RPM_NAME, " & _
		"OPS_RPM_TYPE || ' ' || OPS_RPM_ID " & _
		"FROM OPS_REPORT_MASTER " & _
		"WHERE OPS_RPM_STAND_ALONE = 'Y' " & _
		"AND OPS_RPM_STATUS = 'ACT' " & _
	")PAGE " & _
	"LEFT JOIN " & _
	"( " & _
		"SELECT " & _
		"SYS_CDD_SYS_CDM_ID TYPE_ID, " & _
		"SYS_CDD_VALUE ACCESS_ID " & _
		"FROM SYS_CODE_DETAIL " & _
		"WHERE SYS_CDD_SYS_CDM_ID IN (50,132) " & _
		"AND SYS_CDD_NAME = ? " & _
	")SECURITY " & _
	"ON PAGE.TYPE_ID = SECURITY.TYPE_ID " & _
	"AND PAGE.ACCESS_ID = SECURITY.ACCESS_ID " & _
	"ORDER BY DECODE(PAGE.ACCESS_TYPE,'Department',1,'Application',2,'Page',3,'Report',4,5), PAGE.ACCESS_ADDRESS, PAGE.ACCESS_NAME"
	cmd.CommandText = SQLstmt
	cmd.Parameters(0).value = PARAMETER_AGENT
	Set RSACCESS = cmd.Execute
%>
	<form id="SECURITY_FORM" action="includes/formhandler.asp" method="post">
		<input type="hidden" name="SECURITY_AGENT" value="<%=PARAMETER_AGENT%>"/>
		<div class="table-responsive" style="margin-bottom:1rem;">
			<table id="SECURITY_TABLE" class="table table-bordered center" style="margin-bottom:0;background-color:#fff;">
				<caption class="th-color <% If PARAMETER_DATE = Date Then %>today-color-background<% Else %>past-color-background<% End If %>">
					<span id="SECURITYACCESSCAPTION_<%=PARAMETER_AGENT%>"></span> - <%=FormatDateTime(Now,3)%>
					<input type="submit" class="btn white-background <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" style="margin-top:-5px;padding:1px 8px;font-size:.75rem;" value="Submit"/>
						<div style="float:right;margin-right:15px;">
							<span id="SECURITY_TEXT" style="margin-right:15px;"></span>							
							<i id="SECURITY_PROFILE" class="fas fa-id-card icon-style white" title="Generate Profile" data-department="<%=PARAMETER_DEPARTMENT%>" data-team="<%=PARAMETER_TEAM%>" data-job="<%=PARAMETER_JOB%>"></i>
							<select id="SECURITY_MATCH" class="th-color" style="font-weight:bold">
								<%
									SQLstmt = "SELECT DISTINCT OPS_USR_ID VALUE, OPS_USR_NAME DESCRIPTION " & _
									"FROM OPS_USER " & _
									"JOIN OPS_USER_DETAIL " & _
									"ON OPS_USR_ID = OPS_USD_OPS_USR_ID " & _
									"AND TO_DATE(?,'MM/DD/YYYY') BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE " & _
									"JOIN SYS_CODE_DETAIL " & _
									"ON SYS_CDD_SYS_CDM_ID = 50 " & _
									"AND OPS_USR_ID = SYS_CDD_NAME " & _
									"WHERE OPS_USR_DIS_DATE >= TO_DATE(?,'MM/DD/YYYY') " & _
									"AND OPS_USR_ID <> 9151 "
									If PULSE_SECURITY <= 5 or PARAMETER_LOCATION = "MOT" or PARAMETER_LOCATION = "WFD" or PARAMETER_LOCATION = "WFH" Then
										SQLstmt = SQLstmt & "AND OPS_USD_LOCATION IN ('MOT','WFD','WFH') "
									Else
										SQLstmt = SQLstmt & "AND OPS_USD_LOCATION NOT IN ('MOT','WFD','WFH') "
									End If
									SQLstmt = SQLstmt & "ORDER BY DESCRIPTION"
									cmd.CommandText = SQLstmt
									cmd.Parameters(0).value = PARAMETER_DATE
									cmd.Parameters(1).value = PARAMETER_DATE
									Set RSSELECT = cmd.Execute(SQLstmt)
								%>
								<% If PARAMETER_LOCATION = "MOT" or PARAMETER_LOCATION = "WFD" or PARAMETER_LOCATION = "WFH" Then %>
									<option value="-1" style="color:black;">Assigned Workgroup</option>
								<% Else %>
									<option value="0" style="color:black;">Not Selected</option>
								<% End If %>
								<% Do While Not RSSELECT.EOF %>
									<option value="<%=RSSELECT("VALUE")%>" style="color:black;"><%=RSSELECT("DESCRIPTION")%></option>
									<% RSSELECT.MoveNext %>
								<% Loop %>
								<% Set RSSELECT = Nothing %>
							</select>
						</div>
				</caption>
				<thead>
					<tr class="th-color <% If PARAMETER_DATE = Date Then %>today-color-background<% Else %>past-color-background<% End If %>">
						<th style="width:25%">Security Type</th>
						<th style="width:35%">Name</th>
						<th style="width:40%">Address</th>
					</tr>
				</thead>
				<tbody>
					<% Do While Not RSACCESS.EOF %>
						<input type="checkbox" id="SECURITYACCESS_<%=RSACCESS("TYPE_ID")%>_<%=RSACCESS("ACCESS_ID")%>" name="SECURITY_ACCESS" value="<%=RSACCESS("TYPE_ID")%>_<%=RSACCESS("ACCESS_ID")%>" <% If RSACCESS("ACCESS_FLAG") = "1" Then %>checked="checked"<% End If %> style="display:none;" />
						<tr id="SECURITYROW_<%=RSACCESS("TYPE_ID")%>_<%=RSACCESS("ACCESS_ID")%>" <% If RSACCESS("ACCESS_FLAG") = "1" Then %>class="new-entry-color"<% End If %> style="cursor:pointer;">
							<td><%=RSACCESS("ACCESS_TYPE")%></td>
							<td><%=RSACCESS("ACCESS_NAME")%></td>
							<td><%=RSACCESS("ACCESS_ADDRESS")%></td>
						</tr>
						<% RSACCESS.MoveNext %>
					<% Loop %>
				</tbody>
				<tfoot>
					<tr>
						<td colspan="3">
							<input type="submit" class="btn th-color <% If PARAMETER_DATE = Date Then %>today-color-background<% Else %>past-color-background<% End If %>" value="Submit Changes"/>
						</td>
					</tr>
				</tfoot>
			</table>
		</div>
	</form>
	<script>
		$(document).ready(function() {
			$("#SECURITY_TABLE").DataTable({
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