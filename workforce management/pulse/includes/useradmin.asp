<!--#include file="pulseheader.asp"-->

<%
	SEARCH_BOOL = 0
	DATATABLES_BOOL = 0
	If Request.Querystring("DATE") <> "" Then
		PARAMETER_DATE = CDate(Request.Querystring("DATE"))
	Else
		PARAMETER_DATE = Date
	End If
	If Request.Querystring("AGENT") <> "" Then
		SEARCH_BOOL = 1
		PARAMETER_AGENT = Request.Querystring("AGENT")
	Else
		PARAMETER_AGENT = ""
	End If
	If Request.Querystring("SUPERVISOR") <> "" Then
		SEARCH_BOOL = 1
		PARAMETER_SUPERVISOR = Request.Querystring("SUPERVISOR")
	Else
		PARAMETER_SUPERVISOR = ""
	End If
	If Request.Querystring("DEPARTMENT") <> "" Then
		SEARCH_BOOL = 1
		PARAMETER_DEPARTMENT = Request.Querystring("DEPARTMENT")
	Else
		PARAMETER_DEPARTMENT = ""
	End If
	If Request.Querystring("JOB") <> "" Then
		SEARCH_BOOL = 1
		PARAMETER_JOB = Request.Querystring("JOB")
	Else
		PARAMETER_JOB = ""
	End If
	If Request.Querystring("LOCATION") <> "" Then
		SEARCH_BOOL = 1
		PARAMETER_LOCATION = Request.Querystring("LOCATION")
	Else
		PARAMETER_LOCATION = ""
	End If
	If Request.Querystring("HIRE") <> "" Then
		SEARCH_BOOL = 1
		PARAMETER_HIRE = Request.Querystring("HIRE")
	Else
		PARAMETER_HIRE = ""
	End If
	If Request.Querystring("NEWUSER") = "1" Then
		SEARCH_BOOL = 0
		PARAMETER_NEWUSER = 1
		DATATABLES_BOOL = 1
		Set RSAGTLIST = Conn.Execute("SELECT * FROM DUAL WHERE ROWNUM > 1")
	Else
		PARAMETER_NEWUSER = 0
	End If
	If SEARCH_BOOL = 1 Then
		ReDim ADMIN_PARAMETER_ARRAY(25)	'Parameter index
		SQLstmt = "SELECT " & _
		"OPS_USR_ID ADMIN_USER, " & _
		"OPS_USR_NAME ADMIN_NAME, " & _
		"OPS_USR_NT_ID WINDOWS_ID, " & _
		"OPS_USR_ALT_ID_1 PPR, " & _
		"NULLIF(OPS_USR_SUN_ID,'0') NAVIGATOR_ID, " & _
		"NULLIF(OPS_USR_PHN_ID_PC,0) BADGE_ID, " & _
		"NULLIF(OPS_USR_PHN_EXT,0) PHONE_EXTENSION, " & _
		"OPS_USR_ALT_ID_2 TEXT_ADDRESS, " & _
		"OPS_USR_EMAIL_ADDR EMAIL_ADDRESS, " & _
		"COUNT(*) OVER () AGENT_COUNT " & _
		"FROM OPS_USER " & _
		"JOIN OPS_USER_DETAIL " & _
		"ON OPS_USR_ID = OPS_USD_OPS_USR_ID " & _
		"AND TO_DATE(?,'MM/DD/YYYY') BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE "
		If PULSE_SECURITY <= 5 Then 
			SQLstmt = SQLstmt & "WHERE OPS_USD_LOCATION IN ('MOT','WFD','WFH') " & _
			"AND OPS_USR_ID NOT IN (9618,9151) "
		Else
			SQLstmt = SQLstmt & "WHERE OPS_USR_ID NOT IN (9618,9151) "
		End If
		ADMIN_PARAMETER_ARRAY(0) = PARAMETER_DATE
		i = 1
		If PARAMETER_AGENT <> "" and Instr(PARAMETER_AGENT,"ALL") = 0 Then
			SQLstmt = SQLstmt & "AND "
			If PARAMETER_SUPERVISOR <> "" Then
					SQLstmt = SQLstmt & "( "
			End If
			SQLstmt = SQLstmt & "OPS_USD_OPS_USR_ID IN ("
			USE_ARRAY = Split(PARAMETER_AGENT,",")
			For j = 0 to UBound(USE_ARRAY)
				If j <> UBound(USE_ARRAY) Then
					SQLstmt = SQLstmt & "?,"
				Else
					SQLstmt = SQLstmt & "?) "
				End If
				ADMIN_PARAMETER_ARRAY(i) = USE_ARRAY(j)
				i = i + 1
			Next
		End If
		If PARAMETER_SUPERVISOR <> "" and Instr(PARAMETER_AGENT,"ALL") = 0 Then
			If PARAMETER_AGENT <> "" Then
				SQLstmt = SQLstmt & "OR "
			Else
				SQLstmt = SQLstmt & "AND "
			End If
			SQLstmt = SQLstmt & "OPS_USD_SUPERVISOR IN ("
			USE_ARRAY = Split(PARAMETER_SUPERVISOR,",")
			For j = 0 to UBound(USE_ARRAY)
				If j <> UBound(USE_ARRAY) Then
					SQLstmt = SQLstmt & "?,"
				Else
					SQLstmt = SQLstmt & "?) "
				End If
				ADMIN_PARAMETER_ARRAY(i) = USE_ARRAY(j)
				i = i + 1
			Next
			If PARAMETER_AGENT <> "" Then
				SQLstmt = SQLstmt & ") "
			End If
		End If
		If PARAMETER_DEPARTMENT <> "" Then
			SQLstmt = SQLstmt & "AND DECODE(OPS_USD_TEAM,'SPT','SPT','SRV','SPT','OSR','SPT','TRN','TRN',OPS_USD_TYPE) IN ("
			USE_ARRAY = Split(PARAMETER_DEPARTMENT,",")
			For j = 0 to UBound(USE_ARRAY)
				If j <> UBound(USE_ARRAY) Then
					SQLstmt = SQLstmt & "?,"
				Else
					SQLstmt = SQLstmt & "?) "
				End If
				ADMIN_PARAMETER_ARRAY(i) = USE_ARRAY(j)
				i = i + 1
			Next
		End If
		If PARAMETER_JOB <> "" Then
			SQLstmt = SQLstmt & "AND DECODE(OPS_USD_TYPE,'GRP','AGT',DECODE(OPS_USD_TEAM,'LDA','LED',OPS_USD_JOB)) IN ("
			USE_ARRAY = Split(PARAMETER_JOB,",")
			For j = 0 to UBound(USE_ARRAY)
				If j <> UBound(USE_ARRAY) Then
					SQLstmt = SQLstmt & "?,"
				Else
					SQLstmt = SQLstmt & "?) "
				End If
				ADMIN_PARAMETER_ARRAY(i) = USE_ARRAY(j)
				i = i + 1
			Next
		End If
		If PARAMETER_LOCATION <> "" Then
			SQLstmt = SQLstmt & "AND OPS_USD_LOCATION IN ("
			USE_ARRAY = Split(PARAMETER_LOCATION,",")
			For j = 0 to UBound(USE_ARRAY)
				If j <> UBound(USE_ARRAY) Then
					SQLstmt = SQLstmt & "?,"
				Else
					SQLstmt = SQLstmt & "?) "
				End If
				ADMIN_PARAMETER_ARRAY(i) = USE_ARRAY(j)
				i = i + 1
			Next
		End If
		If PARAMETER_HIRE <> "" Then
			SQLstmt = SQLstmt & "AND OPS_USR_HIRE_DATE IN ("
			USE_ARRAY = Split(PARAMETER_HIRE,",")
			For j = 0 to UBound(USE_ARRAY)
				If j <> UBound(USE_ARRAY) Then
					SQLstmt = SQLstmt & "TO_DATE(?,'MM/DD/YYYY'),"
				Else
					SQLstmt = SQLstmt & "TO_DATE(?,'MM/DD/YYYY')) "
				End If
				ADMIN_PARAMETER_ARRAY(i) = USE_ARRAY(j)
				i = i + 1
			Next
		End If
		SQLstmt = SQLstmt & "ORDER BY OPS_USR_NAME"
		cmd.CommandText = SQLstmt
		i = i - 1
		ReDim Preserve ADMIN_PARAMETER_ARRAY(i)
		For i = 0 to UBound(ADMIN_PARAMETER_ARRAY)
			cmd.Parameters(i).value = ADMIN_PARAMETER_ARRAY(i)
		Next
		Erase ADMIN_PARAMETER_ARRAY
		Set RSAGTLIST = cmd.Execute
	End If
%>
	<% If IsObject(RSAGTLIST) Then %>
		<form id="PULSE_FORM" data-request="ADMIN" action="includes/formhandler.asp" method="post">
				<input type="hidden" id="NEWLINE_ID" value="0"/>
				<input type="hidden" id="ADMINID_LIST" name="ADMINID_LIST" value=""/>
				<input type="hidden" name="FORM_DATE" value="<%=PARAMETER_DATE%>"/>
				<div id="PULSE_FORM_DIV" class="table-responsive" style="margin-bottom:1rem;">
					<table id="EDITADMINMASTERTABLE" class="table table-bordered center" style="margin-bottom:0;">
						<caption class="th-color <% If PARAMETER_DATE = Date Then %>today-color-background<% Else %>past-color-background<% End If %>">
							User Admin - <%=FormatDateTime(Now,3)%>
							<% 
								COUNT_TEXT = ""
								If Not RSAGTLIST.EOF Then
									If RSAGTLIST("AGENT_COUNT") = "1" Then
										COUNT_TEXT = "1 Employee Found"
									Else
										DATATABLES_BOOL = 1
										COUNT_TEXT = RSAGTLIST("AGENT_COUNT") & " Employees Found"
									End If
								End If
							%>
							<div style="float:right"><%=COUNT_TEXT%></div>
						</caption>
						<thead>
							<tr class="th-color <% If PARAMETER_DATE = Date Then %>today-color-background<% Else %>past-color-background<% End If %>">
									<th style="width:17.5%">Employee</th>
									<th style="width:17.5%">Windows ID</th>
									<th style="width:10%">PPR</th>
									<th style="width:5%">Navigator ID</th>
									<th style="width:5%">Badge ID</th>
									<th style="width:5%">Phone Ext.</th>
									<th style="width:15%">Text Address</th>							
									<th style="width:20%">Email Address</th>
							</tr>
						</thead>
						<% If Not RSAGTLIST.EOF or PARAMETER_NEWUSER = 1 Then %>
							<tbody>
							<% Do While Not RSAGTLIST.EOF %>
								<tr id="ADMINROW_<%=RSAGTLIST("ADMIN_USER")%>">
									<td>
										<input type="text" id="ADMINUSER_<%=RSAGTLIST("ADMIN_USER")%>" name="ADMINUSER_<%=RSAGTLIST("ADMIN_USER")%>" class="adminfields center <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" style="font-weight:bold;" value="<%=RSAGTLIST("ADMIN_NAME")%>" maxlength="50" pattern="^([a-zA-Z'-]|\s){3,}$" title="50 characters, including - and '">
									</td>
									<td>
										<input type="text" id="ADMINWINDOWS_<%=RSAGTLIST("ADMIN_USER")%>" name="ADMINWINDOWS_<%=RSAGTLIST("ADMIN_USER")%>" class="adminfields center <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" value="<%=RSAGTLIST("WINDOWS_ID")%>" maxlength="33" pattern="^MLTMTKA\\[a-zA-Z'-]*$" title="25 characters, starting with MLTMTKA\">
									</td>
									<td>
										<input type="text" id="ADMINPPR_<%=RSAGTLIST("ADMIN_USER")%>" name="ADMINPPR_<%=RSAGTLIST("ADMIN_USER")%>" class="adminfields center <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" value="<%=RSAGTLIST("PPR")%>" maxlength="9" pattern="^0[0-9]{6}00$" title="9 digits, format as 012345600">
									</td>
									<td>
										<input type="text" id="ADMINNAVIGATOR_<%=RSAGTLIST("ADMIN_USER")%>" name="ADMINNAVIGATOR_<%=RSAGTLIST("ADMIN_USER")%>" class="adminfields center <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" value="<%=RSAGTLIST("NAVIGATOR_ID")%>" maxlength="7" pattern="^[A-Z][0-9A-Z]{2}[0-9]{3,4}$" title="6-7 characters, format as A123456 or ABC123">
									</td>
									<td>
										<input type="text" id="ADMINBADGE_<%=RSAGTLIST("ADMIN_USER")%>" name="ADMINBADGE_<%=RSAGTLIST("ADMIN_USER")%>" class="adminfields center <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" value="<%=RSAGTLIST("BADGE_ID")%>" maxlength="5" pattern="^[1-9][0-9]{0,4}$" title="1-5 digit number">
									</td>
									<td>
										<input type="text" id="ADMINEXT_<%=RSAGTLIST("ADMIN_USER")%>" name="ADMINEXT_<%=RSAGTLIST("ADMIN_USER")%>" class="adminfields center <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" value="<%=RSAGTLIST("PHONE_EXTENSION")%>" maxlength="4" pattern="^[1-9][0-9]{3}$" title="4 digit number">
									</td>
									<td>
										<input type="text" id="ADMINTEXT_<%=RSAGTLIST("ADMIN_USER")%>" name="ADMINTEXT_<%=RSAGTLIST("ADMIN_USER")%>" class="adminfields center <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" value="<%=RSAGTLIST("TEXT_ADDRESS")%>" maxlength="30" pattern="^[0-9]{10}@[a-z0-9.-]+\.[a-z]{3,}$" title="10 digit phone #, plus carrier-specific @">
									</td>
									<td>
										<input type="text" id="ADMINEMAIL_<%=RSAGTLIST("ADMIN_USER")%>" name="ADMINEMAIL_<%=RSAGTLIST("ADMIN_USER")%>" class="adminfields center <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" value="<%=RSAGTLIST("EMAIL_ADDRESS")%>" maxlength="40" pattern="^[a-zA-Z.'-]*@deltavacations.com$" title="email address ending in @deltavacations.com">
									</td>
								</tr>
								<% If DATATABLES_BOOL = 0 Then %>
									<tr id="EDITADMINROW_<%=RSAGTLIST("ADMIN_USER")%>" style="display:none;">
										<td id="EDITADMINDIV_WRAPPER_<%=RSAGTLIST("ADMIN_USER")%>" colspan="8">
											<div id="EDITADMINDIV_<%=RSAGTLIST("ADMIN_USER")%>"></div>
										</td>
									</tr>
								<% End If %>
								<% RSAGTLIST.MoveNext %>
							<% Loop %>
							</tbody>
							<tfoot>
								<% If PARAMETER_NEWUSER = 1 Then %>
									<tr id="ADMINROW_0" class="new-entry-color" style="display:none;">
										<td>
											<input type="text" id="ADMINUSER_0" name="ADMINUSER_0" class="adminfields center <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" style="font-weight:bold;" value="" maxlength="50" pattern="^([a-zA-Z'-]|\s){3,}$" title="50 characters, including - and '">
										</td>
										<td>
											<input type="text" id="ADMINWINDOWS_0" name="ADMINWINDOWS_0" class="adminfields center <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" value="MLTMTKA\" maxlength="33" pattern="^MLTMTKA\\[a-zA-Z'-]*$" title="25 characters, starting with MLTMTKA\">
										</td>
										<td>
											<input type="text" id="ADMINPPR_0" name="ADMINPPR_0" class="adminfields center <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" value="" maxlength="9" pattern="^0[0-9]{6}00$" title="9 digits, format as 012345600">
										</td>
										<td>
											<input type="text" id="ADMINNAVIGATOR_0" name="ADMINNAVIGATOR_0" class="adminfields center <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" value="" maxlength="7" pattern="^[A-Z][0-9A-Z]{2}[0-9]{3,4}$" title="6-7 characters, format as A123456 or ABC123">
										</td>
										<td>
											<input type="text" id="ADMINBADGE_0" name="ADMINBADGE_0" class="adminfields center <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" value="" maxlength="5" pattern="^[1-9][0-9]{0,4}$" title="1-5 digit number">
										</td>
										<td>
											<input type="text" id="ADMINEXT_0" name="ADMINEXT_0" class="adminfields center <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" value="" maxlength="4" pattern="^[1-9][0-9]{3}$" title="4 digit number">
										</td>
										<td>
											<input type="text" id="ADMINTEXT_0" name="ADMINTEXT_0" class="adminfields center <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" value="" maxlength="30" pattern="^[0-9]{10}@[a-z0-9.-]+\.[a-z]{3,}$" title="10 digit phone #, plus carrier-specific @">
										</td>
										<td>
											<input type="text" id="ADMINEMAIL_0" name="ADMINEMAIL_0" class="adminfields center <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>" value="@deltavacations.com" maxlength="40" pattern="^[a-zA-Z.'-]*@deltavacations.com$" title="email address ending in @deltavacations.com">
										</td>
									</tr>
									<tr class="<% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>">
										<td colspan="8"><i id="NEWADMINMASTERENTRY" class="fas fa-plus-square icon-style-large"></i></td>
									</tr>
								<% End If %>
								<tr>
									<td colspan="8">
										<input id="PULSE_SUBMIT" type="submit" class="btn th-color <% If PARAMETER_DATE = Date Then %>today-color-background<% Else %>past-color-background<% End If %>" value="Submit Changes"/>
										<div id="OVERLAP_MESSAGE" class="error-color" style="display:none;">
											Fix overlapping admin entries before submitting.
										</div>
									</td>
								</tr>
							</tfoot>
						<% Else %>
							<tr>
								<td colspan="8">
									No employees found.
								</td>
							</tr>
						<% End If %>			
					</table>
				</div>
			</form>
	<% End If %>
	<script>
		$.fn.dataTable.ext.order["dom-text"] = function(settings, col){
			return this.api().column( col, {order:"index"} ).nodes().map(function(td, i){
				return $("input", td).val();
			});
		}
		$.fn.dataTable.ext.order["dom-text-numeric"] = function(settings, col){
			return this.api().column( col, {order:"index"} ).nodes().map(function(td, i){
				return $("input", td).val() * 1;
			});
		}
		function initAdminDataTable() {
			return $("#EDITADMINMASTERTABLE").DataTable({
				"autoWidth": false,
				"retrieve": true,
				"paging": false,
				"searching": false,
				"info": false,
				"columnDefs": [
					{"orderDataType": "dom-text", type: "string", "targets": [0, 1, 3, 6, 7]},
					{"orderDataType": "dom-text-numeric", "targets": [2, 4, 5]}
				]
				<% If PARAMETER_NEWUSER = 1 Then %>
					,
					"initComplete": function() {
						newAdminMasterEntry();
					}
				<% End If %>
			});
		}
		$(document).ready(function() {
			adminIdList = [];
			overlapIdList = [];
			<% If DATATABLES_BOOL = 1 Then %>
				var useDataTable = initAdminDataTable();
				<% If PARAMETER_NEWUSER = 0 Then %>
					useDataTable.rows().every(function(){
						var idArray = useDataTable.row(this).id().split("_");
						this.child($(
							'<tr id="EDITADMINROW_' + idArray[1] + '" style="display:none;">' + 
								'<td id="EDITADMINDIV_WRAPPER_' + idArray[1] + '" colspan="8">' +
									'<div id="EDITADMINDIV_' + idArray[1] + '"></div>' +
								'</td>' +
							'</tr>'
							)
						).show();
					});
				<% End If %>
			<% End If %>
		});
	</script>
<!--#include file="pulsefunctions.asp"-->
<% Set cmd = Nothing %>
<% Conn.Close %>
<% Set Conn = Nothing %>