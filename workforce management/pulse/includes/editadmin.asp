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
		PARAMETER_AGENT = "-1"
	End If
%>
<table id="EDITADMINDETAILSTABLE_<%=PARAMETER_AGENT%>" style="width:100%;font-size:.75em;">
	<caption class="center <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>">
		<span id="EDITADMINDETAILSCAPTION_<%=PARAMETER_AGENT%>"></span>
	</caption>
	<thead>
		<tr class="subtable-td-padded-sm <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>">
			<th class="subtable-td-padded-sm" style="width:5%;">Dept</th>
			<th class="subtable-td-padded-sm" style="width:5%;">Team</th>
			<th class="subtable-td-padded-sm" style="width:5%;">Job</th>
			<th class="subtable-td-padded-sm" style="width:5%;">Class</th>
			<th class="subtable-td-padded-sm" style="width:10%;">Location</th>
			<th class="subtable-td-padded-sm" style="width:10%;">Hours</th>
			<th class="subtable-td-padded-sm" style="width:10%;">Supervisor</th>
			<th class="subtable-td-padded-sm" style="width:10%;">Phone ID</th>
			<th class="subtable-td-padded-sm" style="width:10%;">Job Code</th>
			<th class="subtable-td-padded-sm" style="width:15%;">Start Date</th>
			<th class="subtable-td-padded-sm" style="width:15%;">End Date</th>
		</tr>
	</thead>
	<% 
		SQLstmt = "SELECT " & _
		"OPS_USD_ID, " & _
		"OPS_USD_TYPE USE_DEPT, " & _
		"OPS_USD_TEAM USE_TEAM, " & _
		"OPS_USD_JOB USE_JOB, " & _
		"OPS_USD_CLASS USE_CLASS, " & _
		"OPS_USD_LOCATION USE_LOCATION, " & _
		"NVL(OPS_USD_SCH_HOURS,0) USE_HOURS, " & _
		"OPS_USD_SUPERVISOR USE_SUPERVISOR, " & _
		"NULLIF(OPS_USD_PHN_ID,0) USE_PHONE, " & _
		"NULLIF(OPS_USD_JOB_CODE,'0') USE_JOB_CODE, " & _
		"NULLIF(OPS_USD_PAY_RATE,0) USE_PAY, " & _
		"OPS_USD_EFF_DATE USE_EFF_DATE, " & _
		"OPS_USD_DIS_DATE USE_DIS_DATE, " & _
		"MIN(OPS_USD_EFF_DATE) OVER () MIN_DATE " & _
		"FROM OPS_USER_DETAIL " & _
		"WHERE OPS_USD_OPS_USR_ID = ? " & _
		"ORDER BY OPS_USD_EFF_DATE, OPS_USD_DIS_DATE"
		cmd.CommandText = SQLstmt
		cmd.Parameters(0).value = PARAMETER_AGENT
		Set RSADMIN = cmd.Execute

		SQLstmt = "SELECT DISTINCT " & _
		"OPS_USD_TYPE ADMIN_DEPT " & _
		"FROM OPS_USER_DETAIL " & _
		"WHERE OPS_USD_TYPE IS NOT NULL " & _
		"AND OPS_USD_OPS_USR_ID IN " & _
		"( " & _
			"SELECT DISTINCT OPS_USD_OPS_USR_ID " & _
			"FROM OPS_USER_DETAIL " & _
			"WHERE OPS_USD_DIS_DATE >= TO_DATE(?,'MM/DD/YYYY') " & _
		") "
		If PULSE_SECURITY <= 5 Then
			SQLstmt = SQLstmt & "AND OPS_USD_LOCATION IN ('MOT','WFD','WFH') "
		End If
		SQLstmt = SQLstmt & "ORDER BY OPS_USD_TYPE"
		cmd.CommandText = SQLstmt
		cmd.Parameters(0).value = PARAMETER_DATE
		Set RSDEPT = cmd.Execute
		
		SQLstmt =  "SELECT * " & _
		"FROM " & _
		"( " & _
			"SELECT DISTINCT " & _
			"OPS_USD_TEAM  ADMIN_TEAM " & _
			"FROM OPS_USER_DETAIL " & _
			"WHERE OPS_USD_TEAM IS NOT NULL " & _
			"AND " & _
			"( " & _
				"NOT REGEXP_LIKE(OPS_USD_TEAM,'[[:digit:]]') " & _
				"OR OPS_USD_TEAM LIKE 'TM_' " & _
			") " & _
			"AND OPS_USD_OPS_USR_ID IN " & _
			"( " & _
				"SELECT DISTINCT OPS_USD_OPS_USR_ID " & _
				"FROM OPS_USER_DETAIL " & _
				"WHERE OPS_USD_DIS_DATE >= TO_DATE(?,'MM/DD/YYYY') " & _
			") "
			If PULSE_SECURITY <= 5 Then
				SQLstmt = SQLstmt & "AND OPS_USD_LOCATION IN ('MOT','WFD','WFH') "
			End If
			SQLstmt = SQLstmt & "UNION " & _
			"SELECT 'NEW' FROM DUAL " & _
		") " & _
		"ORDER BY ADMIN_TEAM"
		cmd.CommandText = SQLstmt
		cmd.Parameters(0).value = PARAMETER_DATE
		Set RSTEAM = cmd.Execute
		
		SQLstmt =  "SELECT DISTINCT " & _
		"OPS_USD_JOB ADMIN_JOB " & _
		"FROM OPS_USER_DETAIL " & _
		"WHERE OPS_USD_JOB IS NOT NULL " & _
		"AND OPS_USD_OPS_USR_ID IN " & _
		"( " & _
			"SELECT DISTINCT OPS_USD_OPS_USR_ID " & _
			"FROM OPS_USER_DETAIL " & _
			"WHERE OPS_USD_DIS_DATE >= TO_DATE(?,'MM/DD/YYYY') " & _
		") "
		If PULSE_SECURITY <= 5 Then
			SQLstmt = SQLstmt & "AND OPS_USD_LOCATION IN ('MOT','WFD','WFH') "
		End If
		SQLstmt = SQLstmt & "ORDER BY OPS_USD_JOB"
		cmd.CommandText = SQLstmt
		cmd.Parameters(0).value = PARAMETER_DATE
		Set RSJOB = cmd.Execute

		SQLstmt =  "SELECT DISTINCT " & _
		"OPS_USD_CLASS ADMIN_CLASS " & _
		"FROM OPS_USER_DETAIL " & _
		"WHERE OPS_USD_CLASS IS NOT NULL " & _
		"AND OPS_USD_OPS_USR_ID IN " & _
		"( " & _
			"SELECT DISTINCT OPS_USD_OPS_USR_ID " & _
			"FROM OPS_USER_DETAIL " & _
			"WHERE OPS_USD_DIS_DATE >= TO_DATE(?,'MM/DD/YYYY') " & _
		") "
		If PULSE_SECURITY <= 5 Then
			SQLstmt = SQLstmt & "AND OPS_USD_LOCATION IN ('MOT','WFD','WFH') "
		End If
		SQLstmt = SQLstmt & "ORDER BY OPS_USD_CLASS"
		cmd.CommandText = SQLstmt
		cmd.Parameters(0).value = PARAMETER_DATE
		Set RSCLASS = cmd.Execute

		SQLstmt =  "SELECT DISTINCT " & _
		"OPS_USD_LOCATION ADMIN_LOCATION " & _
		"FROM OPS_USER_DETAIL " & _
		"WHERE OPS_USD_LOCATION IS NOT NULL " & _
		"AND OPS_USD_OPS_USR_ID IN " & _
		"( " & _
			"SELECT DISTINCT OPS_USD_OPS_USR_ID " & _
			"FROM OPS_USER_DETAIL " & _
			"WHERE OPS_USD_DIS_DATE >= TO_DATE(?,'MM/DD/YYYY') " & _
		") "
		If PULSE_SECURITY <= 5 Then
			SQLstmt = SQLstmt & "AND OPS_USD_LOCATION IN ('MOT','WFD','WFH') "
		End If
		SQLstmt = SQLstmt & "ORDER BY OPS_USD_LOCATION"
		cmd.CommandText = SQLstmt
		cmd.Parameters(0).value = PARAMETER_DATE
		Set RSLOCATION = cmd.Execute
		
		SQLstmt = "SELECT ROWNUM - 1 ADMIN_HOURS " & _
		"FROM DUAL " & _
		"CONNECT BY ROWNUM < 42"
		cmd.CommandText = SQLstmt
		Set RSHOURS = cmd.Execute

		SQLstmt = "SELECT * " & _
		"FROM " & _
		"( " & _
			"SELECT OPS_USR_ID ADMIN_SUPERVISOR, OPS_USR_NAME SUPERVISOR_NAME " & _
			"FROM OPS_USER " & _
			"JOIN OPS_USER_DETAIL " & _
			"ON OPS_USR_ID = OPS_USD_OPS_USR_ID " & _
			"WHERE OPS_USD_JOB IN ('MGR','SUP','ADM') " & _
			"AND OPS_USD_OPS_USR_ID IN " & _
			"( " & _
				"SELECT DISTINCT OPS_USD_OPS_USR_ID " & _
				"FROM OPS_USER_DETAIL " & _
				"WHERE OPS_USD_DIS_DATE >= TO_DATE(?,'MM/DD/YYYY') " & _
			") "
			If PULSE_SECURITY <= 5 Then
				SQLstmt = SQLstmt & "AND OPS_USD_LOCATION IN ('MOT','WFD','WFH') "
			End If
			SQLstmt = SQLstmt & "UNION " & _
			"SELECT DISTINCT " & _
			"OPS_USR_ID, " & _
			"OPS_USR_NAME " & _
			"FROM OPS_USER " & _
			"JOIN OPS_USER_DETAIL " & _
			"ON OPS_USR_ID = OPS_USD_SUPERVISOR " & _
			"WHERE OPS_USD_OPS_USR_ID IN " & _
			"( " & _
				"SELECT DISTINCT OPS_USD_OPS_USR_ID " & _
				"FROM OPS_USER_DETAIL " & _
				"WHERE OPS_USD_DIS_DATE >= TO_DATE(?,'MM/DD/YYYY') " & _
			") "
			If PULSE_SECURITY <= 5 Then
				SQLstmt = SQLstmt & "AND OPS_USD_LOCATION IN ('MOT','WFD','WFH') "
			End If
			SQLstmt = SQLstmt & ") " & _
		"ORDER BY SUPERVISOR_NAME"
		cmd.CommandText = SQLstmt
		cmd.Parameters(0).value = PARAMETER_DATE
		cmd.Parameters(1).value = PARAMETER_DATE
		Set RSSUPERVISOR = cmd.Execute
	%>
	<tbody>
	<% If Not RSADMIN.EOF Then %>
		<% MIN_DATE = RSADMIN("MIN_DATE") %>
		<% Do While Not RSADMIN.EOF %>
			<tr id="ADMINDETAILROW_<%=RSADMIN("OPS_USD_ID")%>" class="<% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" data-user="<%=PARAMETER_AGENT%>">
				<input type="hidden" id="ADMINDETAILUSER_<%=RSADMIN("OPS_USD_ID")%>" name="ADMINDETAILUSER_<%=RSADMIN("OPS_USD_ID")%>" value="<%=PARAMETER_AGENT%>" />
				<td class="subtable-td-padded-lg">
					<select id="ADMINDEPT_<%=RSADMIN("OPS_USD_ID")%>" name="ADMINDEPT_<%=RSADMIN("OPS_USD_ID")%>" class="<% If PARAMETER_DATE = Date Then %> today-color <% Else %> past-color <% End If %>" style="padding-left:8px;">
						<% RSDEPT.MoveFirst %>
						<% Do While Not RSDEPT.EOF %>
							<option <% If RSDEPT("ADMIN_DEPT") = RSADMIN("USE_DEPT") Then %>selected="selected"<% End If %> value="<%=RSDEPT("ADMIN_DEPT")%>"><%=RSDEPT("ADMIN_DEPT")%></option>
							<% RSDEPT.MoveNext %>
						<% Loop %>
					</select>
				</td>
				<td class="subtable-td-padded-lg">
					<select id="ADMINTEAM_<%=RSADMIN("OPS_USD_ID")%>" name="ADMINTEAM_<%=RSADMIN("OPS_USD_ID")%>" class="<% If PARAMETER_DATE = Date Then %> today-color <% Else %> past-color <% End If %>" style="padding-left:8px;">
						<% RSTEAM.MoveFirst %>
						<% Do While Not RSTEAM.EOF %>
							<option <% If RSTEAM("ADMIN_TEAM") = RSADMIN("USE_TEAM") Then %>selected="selected"<% End If %> value="<%=RSTEAM("ADMIN_TEAM")%>"><%=RSTEAM("ADMIN_TEAM")%></option>
							<% RSTEAM.MoveNext %>
						<% Loop %>
					</select>
				</td>
				<td class="subtable-td-padded-lg">
					<select id="ADMINJOB_<%=RSADMIN("OPS_USD_ID")%>" name="ADMINJOB_<%=RSADMIN("OPS_USD_ID")%>" class="<% If PARAMETER_DATE = Date Then %> today-color <% Else %> past-color <% End If %>" style="padding-left:8px;">
						<% RSJOB.MoveFirst %>
						<% Do While Not RSJOB.EOF %>
							<option <% If RSJOB("ADMIN_JOB") = RSADMIN("USE_JOB") Then %>selected="selected"<% End If %> value="<%=RSJOB("ADMIN_JOB")%>" <% If SECURITY_LEVEL <> 6 and RSADMIN("USE_JOB") <> "ADM" and RSJOB("ADMIN_JOB") = "ADM" Then %>disabled="disabled"<% End If %>><%=RSJOB("ADMIN_JOB")%></option>
							<% RSJOB.MoveNext %>
						<% Loop %>
					</select>
				</td>
				<td class="subtable-td-padded-lg">
					<select id="ADMINCLASS_<%=RSADMIN("OPS_USD_ID")%>" name="ADMINCLASS_<%=RSADMIN("OPS_USD_ID")%>" class="<% If PARAMETER_DATE = Date Then %> today-color <% Else %> past-color <% End If %>" style="padding-left:8px;">
						<% RSCLASS.MoveFirst %>
						<% Do While Not RSCLASS.EOF %>
							<option <% If RSCLASS("ADMIN_CLASS") = RSADMIN("USE_CLASS") Then %>selected="selected"<% End If %> value="<%=RSCLASS("ADMIN_CLASS")%>"><%=RSCLASS("ADMIN_CLASS")%></option>
							<% RSCLASS.MoveNext %>
						<% Loop %>
					</select>
				</td>
				<td class="subtable-td-padded-lg">
					<select id="ADMINLOCATION_<%=RSADMIN("OPS_USD_ID")%>" name="ADMINLOCATION_<%=RSADMIN("OPS_USD_ID")%>" class="<% If PARAMETER_DATE = Date Then %> today-color <% Else %> past-color <% End If %>" style="padding-left:8px;">
						<% RSLOCATION.MoveFirst %>
						<% Do While Not RSLOCATION.EOF %>
							<option <% If RSLOCATION("ADMIN_LOCATION") = RSADMIN("USE_LOCATION") Then %>selected="selected"<% End If %> value="<%=RSLOCATION("ADMIN_LOCATION")%>"><%=RSLOCATION("ADMIN_LOCATION")%></option>
							<% RSLOCATION.MoveNext %>
						<% Loop %>
					</select>
				</td>
				<td class="subtable-td-padded-lg">
					<select id="ADMINHOURS_<%=RSADMIN("OPS_USD_ID")%>" name="ADMINHOURS_<%=RSADMIN("OPS_USD_ID")%>" class="<% If PARAMETER_DATE = Date Then %> today-color <% Else %> past-color <% End If %>" style="padding-left:8px;">
						<% RSHOURS.MoveFirst %>
						<% Do While Not RSHOURS.EOF %>
							<option <% If CInt(RSHOURS("ADMIN_HOURS")) = CInt(RSADMIN("USE_HOURS")) Then %>selected="selected"<% End If %> value="<%=RSHOURS("ADMIN_HOURS")%>"><%=RSHOURS("ADMIN_HOURS")%></option>
							<% RSHOURS.MoveNext %>
						<% Loop %>
					</select>
				</td>
				<td class="subtable-td-padded-lg">
					<select id="ADMINSUPERVISOR_<%=RSADMIN("OPS_USD_ID")%>" name="ADMINSUPERVISOR_<%=RSADMIN("OPS_USD_ID")%>" class="<% If PARAMETER_DATE = Date Then %> today-color <% Else %> past-color <% End If %>" style="padding-left:8px;">
						<% RSSUPERVISOR.MoveFirst %>
						<option <% If RSADMIN("USE_SUPERVISOR") = "0" Then %>selected="selected"<% End If %> value="0">N/A</option>
						<% Do While Not RSSUPERVISOR.EOF %>
							<option <% If CInt(RSSUPERVISOR("ADMIN_SUPERVISOR")) = CInt(RSADMIN("USE_SUPERVISOR")) Then %>selected="selected"<% End If %> value="<%=RSSUPERVISOR("ADMIN_SUPERVISOR")%>"><%=RSSUPERVISOR("SUPERVISOR_NAME")%></option>
							<% RSSUPERVISOR.MoveNext %>
						<% Loop %>
					</select>
				</td>
				<td class="subtable-td-padded-lg">
					<input type="text" id="ADMINPHONE_<%=RSADMIN("OPS_USD_ID")%>" name="ADMINPHONE_<%=RSADMIN("OPS_USD_ID")%>" class="center <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" value="<%=RSADMIN("USE_PHONE")%>" maxlength="5" pattern="^8[0-9]{4}$" title="5 digit number, from 80000 - 89999">
				</td>
				<td class="subtable-td-padded-lg">
					<input type="text" id="ADMINJOBCODE_<%=RSADMIN("OPS_USD_ID")%>" name="ADMINJOBCODE_<%=RSADMIN("OPS_USD_ID")%>" class="center <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" value="<%=RSADMIN("USE_JOB_CODE")%>" maxlength="6" pattern="^008[0-9]{3}$" title="6 digit string, beginning with 008">
				</td>
				<input type="hidden" id="ADMINPAY_<%=RSADMIN("OPS_USD_ID")%>" name="ADMINPAY_<%=RSADMIN("OPS_USD_ID")%>" value="<%=RSADMIN("USE_PAY")%>" />
				<td class="subtable-td-padded-lg">
					<input type="text" id="ADMINEFFDATE_<%=RSADMIN("OPS_USD_ID")%>" name="ADMINEFFDATE_<%=RSADMIN("OPS_USD_ID")%>" class="admin-date center <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" value="<%=RSADMIN("USE_EFF_DATE")%>" data-date-min="<%=MIN_DATE%>" data-date-max="12/31/2040">
				</td>
				<td class="subtable-td-padded-lg">
					<input type="text" id="ADMINDISDATE_<%=RSADMIN("OPS_USD_ID")%>" name="ADMINDISDATE_<%=RSADMIN("OPS_USD_ID")%>" class="admin-date center <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" value="<%=RSADMIN("USE_DIS_DATE")%>" data-date-min="<%=MIN_DATE%>" data-date-max="12/31/2040">
				</td>
			</tr>
		<% RSADMIN.MoveNext %>
		<% Loop %>
	<% Else %>
		<% MIN_DATE = PARAMETER_DATE %>
	<% End If %>
	<% Set RSADMIN = Nothing %>
		<tr id="ADMINDETAILROW_0" class="new-entry-color <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" style="display:none;" data-user="<%=PARAMETER_AGENT%>">
			<input type="hidden" id="ADMINDETAILUSER_0" name="ADMINDETAILUSER_0" value="<%=PARAMETER_AGENT%>" />
			<td class="subtable-td-padded-lg">
				<select id="ADMINDEPT_0" name="ADMINDEPT_0" class="<% If PARAMETER_DATE = Date Then %> today-color <% Else %> past-color <% End If %>" style="padding-left:8px;">
					<% RSDEPT.MoveFirst %>
					<% Do While Not RSDEPT.EOF %>
						<option value="<%=RSDEPT("ADMIN_DEPT")%>"><%=RSDEPT("ADMIN_DEPT")%></option>
						<% RSDEPT.MoveNext %>
					<% Loop %>
				</select>
			</td>
			<td class="subtable-td-padded-lg">
				<select id="ADMINTEAM_0" name="ADMINTEAM_0" class="<% If PARAMETER_DATE = Date Then %> today-color <% Else %> past-color <% End If %>" style="padding-left:8px;">
					<% RSTEAM.MoveFirst %>
					<% Do While Not RSTEAM.EOF %>
						<option value="<%=RSTEAM("ADMIN_TEAM")%>"><%=RSTEAM("ADMIN_TEAM")%></option>
						<% RSTEAM.MoveNext %>
					<% Loop %>
				</select>
			</td>
			<td class="subtable-td-padded-lg">
				<select id="ADMINJOB_0" name="ADMINJOB_0" class="<% If PARAMETER_DATE = Date Then %> today-color <% Else %> past-color <% End If %>" style="padding-left:8px;">
					<% RSJOB.MoveFirst %>
					<% Do While Not RSJOB.EOF %>
						<option value="<%=RSJOB("ADMIN_JOB")%>" <% If SECURITY_LEVEL <> 6 and RSJOB("ADMIN_JOB") = "ADM" Then %>disabled="disabled"<% End If %>><%=RSJOB("ADMIN_JOB")%></option>
						<% RSJOB.MoveNext %>
					<% Loop %>
				</select>
			</td>
			<td class="subtable-td-padded-lg">
				<select id="ADMINCLASS_0" name="ADMINCLASS_0" class="<% If PARAMETER_DATE = Date Then %> today-color <% Else %> past-color <% End If %>" style="padding-left:8px;">
					<% RSCLASS.MoveFirst %>
					<% Do While Not RSCLASS.EOF %>
						<option value="<%=RSCLASS("ADMIN_CLASS")%>"><%=RSCLASS("ADMIN_CLASS")%></option>
						<% RSCLASS.MoveNext %>
					<% Loop %>
				</select>
			</td>
			<td class="subtable-td-padded-lg">
				<select id="ADMINLOCATION_0" name="ADMINLOCATION_0" class="<% If PARAMETER_DATE = Date Then %> today-color <% Else %> past-color <% End If %>" style="padding-left:8px;">
					<% RSLOCATION.MoveFirst %>
					<% Do While Not RSLOCATION.EOF %>
						<option value="<%=RSLOCATION("ADMIN_LOCATION")%>"><%=RSLOCATION("ADMIN_LOCATION")%></option>
						<% RSLOCATION.MoveNext %>
					<% Loop %>
				</select>
			</td>
			<td class="subtable-td-padded-lg">
				<select id="ADMINHOURS_0" name="ADMINHOURS_0" class="<% If PARAMETER_DATE = Date Then %> today-color <% Else %> past-color <% End If %>" style="padding-left:8px;">
					<% RSHOURS.MoveFirst %>
					<% Do While Not RSHOURS.EOF %>
						<option value="<%=RSHOURS("ADMIN_HOURS")%>"><%=RSHOURS("ADMIN_HOURS")%></option>
						<% RSHOURS.MoveNext %>
					<% Loop %>
				</select>
			</td>
			<td class="subtable-td-padded-lg">
				<select id="ADMINSUPERVISOR_0" name="ADMINSUPERVISOR_0" class="<% If PARAMETER_DATE = Date Then %> today-color <% Else %> past-color <% End If %>" style="padding-left:8px;">
					<% RSSUPERVISOR.MoveFirst %>
					<option value="0">N/A</option>
					<% Do While Not RSSUPERVISOR.EOF %>
						<option value="<%=RSSUPERVISOR("ADMIN_SUPERVISOR")%>"><%=RSSUPERVISOR("SUPERVISOR_NAME")%></option>
						<% RSSUPERVISOR.MoveNext %>
					<% Loop %>
				</select>
			</td>
			<td class="subtable-td-padded-lg">
				<input type="text" id="ADMINPHONE_0" name="ADMINPHONE_0" class="center <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" value="" maxlength="5" pattern="^8[0-9]{4}$" title="5 digit number, from 80000 - 89999">
			</td>
			<td class="subtable-td-padded-lg">
				<input type="text" id="ADMINJOBCODE_0" name="ADMINJOBCODE_0" class="center <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" value="" maxlength="6" pattern="^008[0-9]{3}$" title="6 digit string, beginning with 008">
			</td>
			<input type="hidden" id="ADMINPAY_0" name="ADMINPAY_0" value="" />
			<td class="subtable-td-padded-lg">
				<input type="text" id="ADMINEFFDATE_0" name="ADMINEFFDATE_0" class="admin-date center <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" value="" data-date-min="<%=MIN_DATE%>" data-date-max="12/31/2040">
			</td>
			<td class="subtable-td-padded-lg">
				<input type="text" id="ADMINDISDATE_0" name="ADMINDISDATE_0" class="admin-date center <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" value="" data-date-min="<%=MIN_DATE%>" data-date-max="12/31/2040">
			</td>
		</tr>
		<% Set RSDEPT = Nothing %>
		<% Set RSTEAM = Nothing %>
		<% Set RSJOB = Nothing %>
		<% Set RSCLASS = Nothing %>
		<% Set RSLOCATION = Nothing %>
		<% Set RSHOURS = Nothing %>
		<% Set RSSUPERVISOR = Nothing %>
		<tr class="<% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>">
			<td class="subtable-td-padded-lg" colspan="11">
				<div style="margin-top:5px;">
					<div style="float:left;margin-left:15px;">
						<i id="ADMINREFRESH_<%=PARAMETER_AGENT%>" class="fas fa-sync-alt icon-style" title="Refresh"></i>
						<i id="ADMINSECURITY_<%=PARAMETER_AGENT%>" class="fas fa-key icon-style" title="Security Access"></i>
						<i id="PULSEHELP_ADMIN" class="fas fa-question icon-style" title="Help"></i>
					</div>
					<i id="NEWADMINDETAILSENTRY_<%=PARAMETER_AGENT%>" class="fas fa-plus-square icon-style" title="New Entry"></i>
				</div>
			</td>
		</tr>
	</tbody>
</table>
<script>
	$(document).ready(function() {
		$(".admin-date:not([id$='_0'])").each(function(){
			$(this).datepicker({
				dateFormat: "m/d/yy",
				showAnim: "slideDown",
				showOtherMonths: true,
				selectOtherMonths: true,
				minDate: new Date($(this).data("date-min")),
				maxDate: new Date($(this).data("date-max")),
				onClose: function(dateText) {
					if(moment(dateText,"M/D/YYYY", true).isValid()){
						$(this).val(moment.min(moment.max(moment(dateText,"MM/DD/YYYY"),moment($(this).data("date-min"),"MM/DD/YYYY")),moment($(this).data("date-max"),"MM/DD/YYYY")).format("l"));
						var idArray = this.id.split("_");
						checkAdminOverlaps("<%=PARAMETER_AGENT%>");
						addAdminList("DETAIL_" + idArray[1]);
						addAdminList("MASTER_" + $("#ADMINDETAILUSER_" + idArray[1]).val());
					}
					else{
						alert("Invalid date. Enter a valid date in MM/DD/YYYY format");
					}
				}
			});
		});
	});
</script>
<!--#include file="pulsefunctions.asp"-->
<% Set cmd = Nothing %>
<% Conn.Close %>
<% Set Conn = Nothing %>