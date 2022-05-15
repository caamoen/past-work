<!--#include file="includes/pulseheader.asp"-->
<%
	SQLstmt = "SELECT SEC.SYS_CDD_ID " & _
	"FROM SYS_CODE_DETAIL PAGE " & _
	"JOIN SYS_CODE_DETAIL SEC " & _
	"ON PAGE.SYS_CDD_ID = SEC.SYS_CDD_VALUE " & _
	"JOIN OPS_USER " & _
	"ON OPS_USR_ID = SEC.SYS_CDD_NAME " & _
	"AND TO_DATE(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE)) BETWEEN OPS_USR_EFF_DATE AND OPS_USR_DIS_DATE " & _
	"WHERE PAGE.SYS_CDD_SYS_CDM_ID = 47 " & _
	"AND SEC.SYS_CDD_SYS_CDM_ID = 50 " & _
	"AND OPS_USR_ID = ? " & _
	"AND PAGE.SYS_CDD_VALUE = LOWER(?)"
	cmd.CommandText = SQLstmt
	cmd.Parameters(0).value = PULSE_USR_ID
	cmd.Parameters(1).value = Request.ServerVariables("SCRIPT_NAME")
	Set RSACCESS = cmd.Execute
	If RSACCESS.EOF Then
		Response.Redirect "/index.asp"
	End If
	Set RSACCESS = Nothing 
%>
<!DOCTYPE html>
<html lang="en">
	<head>
		<!-- Required meta tags -->
		<meta charset="utf-8">
		<meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">

		<link rel="stylesheet" href="navbar-fixed-left.min.css?v=1.13">
		<link rel="stylesheet" href="jquery.circliful.css?v=1.1">
		<link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.1.3/css/bootstrap.min.css">
		<link rel="stylesheet" href="chosen.min.css?v=1.02">
		<link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.9.0/css/all.css">
		<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
		<link rel="stylesheet" href="https://fonts.googleapis.com/css?family=Noto+Sans">
		<link rel="stylesheet" href="https://cdn.datatables.net/v/bs4/dt-1.10.21/fh-3.1.7/datatables.min.css"/>
		<link rel="stylesheet" href="pulsecss.css?v=1.42">
		<link rel="stylesheet" href="pacecss.css?v=1.01">
		<link rel="icon" type="image/ico" href="pulseicon.ico">	
		<script src="https://www.gstatic.com/charts/loader.js"></script>
		<script>
			google.charts.load("current", {packages:["corechart"]});
		</script>
		<title>Pulse</title>
	</head>
	<body>
		<input type="hidden" id="PULSE_SECURITY" value="<%=PULSE_SECURITY%>" />
		<nav id="title-bar" class="navbar navbar-dark fixed-top navbar-expand-md p-0 shadow today-color-background">
			<div class="col-auto col-md-2 mr-0 text-nowrap">
				<a id="HOME_ICON" class="navbar-brand" style="margin-right:0;" href="#">
					<i class="fas fa-heartbeat fa"></i>
				</a>
				<div class="navbar-brand" style="margin-right:0;">
					Pulse
					<% If PULSE_SECURITY >= 5 Then %>
						<% If Request.Querystring("MODE") = "ADMIN" Then %>
							<% USE_MODE = "ADMIN" %>
						<% Else %>
							<% USE_MODE = "STAFFING" %>
						<% End If %>
						<sub style="font-size:50%;left:-0.25em;line-height:inherit;">
							<select id="PULSE_MODE" class="today-color-background white">
								<option <% If USE_MODE = "STAFFING" Then %>selected="selected"<% End If %> value="STAFFING">Staffing</option>
								<option <% If USE_MODE = "ADMIN" Then %>selected="selected"<% End If %> value="ADMIN">Admin</option>
							</select>
						</sub>
					<% Else %>
						<input type="hidden" id="PULSE_MODE" value="STAFFING">
					<% End If %>
				</div>
			</div>
			<button class="navbar-toggler" type="button" data-toggle="collapse" data-target="#title-collapse" aria-controls="title-collapse" aria-expanded="false" aria-label="Toggle navigation">
				<span class="navbar-toggler-icon"></span>
			</button>
			<div class="collapse navbar-collapse" id="title-collapse">
				<select class="chosen-select" id="search-bar" multiple style="display:none;"></select>
				<div id="date_div" class="navbar-nav px-3">
					<input id="PARAMETER_DATE" name="PARAMETER_DATE" type="text" value="<%=Date%>" class="white form-control input-sm no-border-input"/>
				</div>
			</div>
		</nav>
		<div class="container-fluid" style="height:100vh;">
			<div class="row">
				<nav class="col-md-2 navbar-fixed-left navbar-expand-md navbar-light bg-light" style="padding-left:5px;z-index:1029">
					<button class="navbar-toggler" type="button" data-toggle="collapse" data-target="#sidebar-collapse" aria-controls="sidebar-collapse" aria-expanded="false" aria-label="Toggle navigation">
						<span class="navbar-toggler-icon"></span>
					</button>
					<div class="collapse navbar-collapse" id="sidebar-collapse">
						<ul class="nav flex-column">
							<% If (InStr(PULSE_DEPARTMENT,"RES") <> 0 or PULSE_SECURITY >= 5) and PULSE_SECURITY >= 1 Then %>
								<div class="btn-group dropright mb-1 staffgroup">
									<button id="GRAPH_BUTTON" type="button" class="btn btn-secondary dropdown-toggle today-color-background" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
										<i class="fas fa-chart-line"></i> Graphs
									</button>
									<div class="dropdown-menu">
										<button id="ALL_GRAPH" class="graph-item dropdown-item" type="button"><span style="font-weight:900;">All Workgroups</span></button>
										<button id="RES_GRAPH" class="graph-item dropdown-item" type="button">Reservations</button>
										<button id="SPT_GRAPH" class="graph-item dropdown-item" type="button">Support Desk</button>
										<button id="OSR_GRAPH" class="graph-item dropdown-item" type="button">Overnight Support</button>
										<button id="SRV_GRAPH" class="graph-item dropdown-item" type="button">Elite Service</button>
										<button id="SLS_GRAPH" class="graph-item dropdown-item" type="button">Sales</button>
									</div>
								</div>
							<% End If %>
							<% If (InStr(PULSE_DEPARTMENT,"RES") <> 0 or InStr(PULSE_DEPARTMENT,"OSS") <> 0 or InStr(PULSE_DEPARTMENT,"CRT") <> 0 or PULSE_SECURITY >= 5) and PULSE_SECURITY >= 1 Then %>
								<div class="btn-group dropright mb-1 staffgroup">
									<button id="DATA_BUTTON" type="button" class="btn btn-secondary dropdown-toggle today-color-background" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
										<i class="fas fa-table"></i> Data
									</button>
									<div class="dropdown-menu">
										<% If (InStr(PULSE_DEPARTMENT,"RES") <> 0 or PULSE_SECURITY >= 5) and PULSE_SECURITY >= 1 Then %>
											<button id="ALLRES_DATA" class="data-item dropdown-item" type="button"><span style="font-weight:900;">All RES & SPT</span></button>
											<button id="RES_DATA" class="data-item dropdown-item" type="button">Reservations</button>
											<button id="SPT_DATA" class="data-item dropdown-item" type="button">Support Desk</button>
											<button id="OSR_DATA" class="data-item dropdown-item" type="button">Overnight Support</button>
											<button id="SRV_DATA" class="data-item dropdown-item" type="button">Elite Service</button>
											<button id="SLS_DATA" class="data-item dropdown-item" type="button">Sales</button>
										<% End If %>
										<% If (InStr(PULSE_DEPARTMENT,"OSS") <> 0 or PULSE_SECURITY >= 5) and PULSE_SECURITY >= 1 Then %>
											<button id="ALLOSS_DATA" class="data-item dropdown-item" type="button"><span style="font-weight:900;">All OSS</span></button>
											<button id="AIR_DATA" class="data-item dropdown-item" type="button">Air Support</button>
											<button id="PRD_DATA" class="data-item dropdown-item" type="button">Product Support</button>
											<button id="SKD_DATA" class="data-item dropdown-item" type="button">Schedule Change</button>
										<% End If %>
										<% If (InStr(PULSE_DEPARTMENT,"CRT") <> 0 or PULSE_SECURITY >= 5) and PULSE_SECURITY >= 1 Then %>
											<button id="CRT_DATA" class="data-item dropdown-item" type="button"><span style="font-weight:900;">Customer Relations</span></button>
										<% End If %>
									</div>
								</div>
							<% End If %>
							<% If (PULSE_DEPARTMENT <> "" or PULSE_SECURITY >= 5) and PULSE_SECURITY >= 1 Then %>
								<div class="btn-group dropright mb-1 staffgroup">
									<% If InStr(PULSE_DEPARTMENT,"RES") = 0 and Instr(PULSE_DEPARTMENT,",") = 0 and PULSE_SECURITY < 5 Then %>
										<button id="<%=PULSE_DEPARTMENT%>_PENDING" class="btn btn-secondary today-color-background pending-item" type="button">
											<i class="fas fa-users" style="margin-right:3px;"></i> Pending
										</button>
									<% Else %>
										<button id="PENDING_BUTTON" type="button" class="btn btn-secondary dropdown-toggle today-color-background" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
											<i class="fas fa-users" style="margin-right:3px;"></i> Pending
										</button>
										<div class="dropdown-menu">
											<% If InStr(PULSE_DEPARTMENT,"RES") <> 0 or PULSE_SECURITY >= 5 Then %>
												<button id="SLS_PENDING" class="pending-item dropdown-item" type="button"><span style="font-weight:900;">Sales</span></button>
												<button id="SRV_PENDING" class="pending-item dropdown-item" type="button"><span style="font-weight:900;">Elite Service</span></button>
												<button id="SPT_PENDING" class="pending-item dropdown-item" type="button"><span style="font-weight:900;">Support Desk</span></button>
											<% End If %>
											<% If InStr(PULSE_DEPARTMENT,"ACC") <> 0 or PULSE_SECURITY >= 5 Then %>
												<button id="ACC_PENDING" class="pending-item dropdown-item" type="button">Accounting</button>
											<% End If %>
											<% If InStr(PULSE_DEPARTMENT,"CRT") <> 0 or PULSE_SECURITY >= 5 Then %>
												<button id="CRT_PENDING" class="pending-item dropdown-item" type="button">Customer Relations</button>
											<% End If %>
											<% If InStr(PULSE_DEPARTMENT,"DOC") <> 0 or PULSE_SECURITY >= 5 Then %>
												<button id="DOC_PENDING" class="pending-item dropdown-item" type="button">Documents</button>
											<% End If %>
											<% If InStr(PULSE_DEPARTMENT,"GRP") <> 0 or PULSE_SECURITY >= 5 Then %>
												<button id="GRP_PENDING" class="pending-item dropdown-item" type="button">Group</button>
											<% End If %>
											<% If InStr(PULSE_DEPARTMENT,"OSS") <> 0 or PULSE_SECURITY >= 5 Then %>
												<button id="OSS_PENDING" class="pending-item dropdown-item" type="button">Operations Support</button>
											<% End If %>
											<% If InStr(PULSE_DEPARTMENT,"POP") <> 0 or PULSE_SECURITY >= 5 Then %>
												<button id="POP_PENDING" class="pending-item dropdown-item" type="button">Product Operations</button>
											<% End If %>
										</div>
									<% End If %>
								</div>
							<% End If %>
							<% If (PULSE_DEPARTMENT <> "" or PULSE_SECURITY >= 5) and PULSE_SECURITY >= 1 Then %>
								<div class="btn-group dropright mb-1 staffgroup">
									<% If InStr(PULSE_DEPARTMENT,"RES") = 0 and Instr(PULSE_DEPARTMENT,",") = 0 and PULSE_SECURITY < 5 Then %>
										<button id="<%=PULSE_DEPARTMENT%>_ERROR" class="btn btn-secondary today-color-background error-item" type="button">
											<i class="fas fa-exclamation-triangle" style="margin-right:1px;"></i> Errors
										</button>
									<% Else %>
										<button id="ERROR_BUTTON" type="button" class="btn btn-secondary dropdown-toggle today-color-background" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
											<i class="fas fa-exclamation-triangle" style="margin-right:1px;"></i> Errors
										</button>
										<div class="dropdown-menu">
											<button id="ALL_ERROR" class="error-item dropdown-item" type="button"><span style="font-weight:900;">All Departments</span></button>
											<% If InStr(PULSE_DEPARTMENT,"ACC") <> 0 or PULSE_SECURITY >= 5 Then %>	
												<button id="ACC_ERROR" class="error-item dropdown-item" type="button">Accounting</button>
											<% End If %>
											<% If InStr(PULSE_DEPARTMENT,"CRT") <> 0 or PULSE_SECURITY >= 5 Then %>
												<button id="CRT_ERROR" class="error-item dropdown-item" type="button">Customer Relations</button>
											<% End If %>
											<% If InStr(PULSE_DEPARTMENT,"DOC") <> 0 or PULSE_SECURITY >= 5 Then %>
												<button id="DOC_ERROR" class="error-item dropdown-item" type="button">Documents</button>
											<% End If %>
											<% If InStr(PULSE_DEPARTMENT,"GRP") <> 0 or PULSE_SECURITY >= 5 Then %>
												<button id="GRP_ERROR" class="error-item dropdown-item" type="button">Group</button>
											<% End If %>
											<% If (InStr(PULSE_DEPARTMENT,"RES") <> 0 and PULSE_SECURITY >= 3) or PULSE_SECURITY >= 5 Then %>
												<button id="NEW_ERROR" class="error-item dropdown-item" type="button">New Hires</button>
											<% End If %>
											<% If InStr(PULSE_DEPARTMENT,"OPS") <> 0 or PULSE_SECURITY >= 5 Then %>	
												<button id="OPS_ERROR" class="error-item dropdown-item" type="button">Operations</button>
											<% End If %>
											<% If InStr(PULSE_DEPARTMENT,"OSS") <> 0 or PULSE_SECURITY >= 5 Then %>
												<button id="OSS_ERROR" class="error-item dropdown-item" type="button">Operations Support</button>
											<% End If %>
											<% If InStr(PULSE_DEPARTMENT,"POP") <> 0 or PULSE_SECURITY >= 5 Then %>
												<button id="POP_ERROR" class="error-item dropdown-item" type="button">Product Operations</button>
											<% End If %>
											<% If InStr(PULSE_DEPARTMENT,"RES") <> 0 or PULSE_SECURITY >= 5 Then %>
												<button id="RES_ERROR" class="error-item dropdown-item" type="button">Reservations</button>
												<button id="SPT_ERROR" class="error-item dropdown-item" type="button">Support Desk</button>
											<% End If %>
										</div>
									<% End If %>
								</div>
							<% End If %>
							<% If (PULSE_DEPARTMENT <> "" or PULSE_SECURITY >= 5) and PULSE_SECURITY >= 1 Then %>
								<div class="btn-group dropright mb-1 staffgroup">
									<% If InStr(PULSE_DEPARTMENT,"RES") = 0 and Instr(PULSE_DEPARTMENT,",") = 0 and PULSE_SECURITY < 5 Then %>
										<button id="<%=PULSE_DEPARTMENT%>_NOTES" class="btn btn-secondary today-color-background note-item" type="button">
											<i class="fas fa-comment"></i> Notes
										</button>
									<% Else %>
										<button id="NOTES_BUTTON" type="button" class="btn btn-secondary dropdown-toggle today-color-background" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
											<i class="fas fa-comment"></i> Notes
										</button>
										<div class="dropdown-menu">
											<button id="ALL_NOTES" class="note-item dropdown-item" type="button"><span style="font-weight:900;">All Departments</span></button>
											<% If InStr(PULSE_DEPARTMENT,"ACC") <> 0 or PULSE_SECURITY >= 5 Then %>	
												<button id="ACC_NOTES" class="note-item dropdown-item" type="button">Accounting</button>
											<% End If %>
											<% If InStr(PULSE_DEPARTMENT,"CRT") <> 0 or PULSE_SECURITY >= 5 Then %>	
												<button id="CRT_NOTES" class="note-item dropdown-item" type="button">Customer Relations</button>
											<% End If %>
											<% If InStr(PULSE_DEPARTMENT,"DOC") <> 0 or PULSE_SECURITY >= 5 Then %>	
												<button id="DOC_NOTES" class="note-item dropdown-item" type="button">Documents</button>
											<% End If %>
											<% If InStr(PULSE_DEPARTMENT,"GRP") <> 0 or PULSE_SECURITY >= 5 Then %>	
												<button id="GRP_NOTES" class="note-item dropdown-item" type="button">Group</button>
											<% End If %>
											<% If InStr(PULSE_DEPARTMENT,"OPS") <> 0 or PULSE_SECURITY >= 5 Then %>	
												<button id="OPS_NOTES" class="note-item dropdown-item" type="button">Operations</button>
											<% End If %>
											<% If InStr(PULSE_DEPARTMENT,"OSS") <> 0 or PULSE_SECURITY >= 5 Then %>	
												<button id="OSS_NOTES" class="note-item dropdown-item" type="button">Operations Support</button>
											<% End If %>
											<% If InStr(PULSE_DEPARTMENT,"POP") <> 0 or PULSE_SECURITY >= 5 Then %>	
												<button id="POP_NOTES" class="note-item dropdown-item" type="button">Product Operations</button>
											<% End If %>
											<% If InStr(PULSE_DEPARTMENT,"RES") <> 0 or PULSE_SECURITY >= 5 Then %>	
												<button id="RES_NOTES" class="note-item dropdown-item" type="button">Reservations</button>
												<button id="SPT_NOTES" class="note-item dropdown-item" type="button">Support Desk</button>
											<% End If %>
										</div>
									<% End If %>
								</div>
							<% End If %>
							<% If PULSE_SECURITY >= 5 Then %>
								<div class="btn-group mb-1 staffgroup" style="margin-top:15px;">
									<button id="TRADES_BUTTON" class="btn btn-secondary today-color-background" type="button">
										<i class="fas fa-exchange-alt"></i> Trades
									</button>
								</div>
								<div class="btn-group mb-1 admingroup">
									<button id="ADMIN_BUTTON" class="btn btn-secondary today-color-background" type="button">
										<i class="fas fa-folder-plus"></i> New User
									</button>
								</div>
							<% End If %>
						</ul>
					</div>
				</nav>
				<div class="col-md-2"></div>
				<main role="main" class="col-md-10 col-lg-10 mt-4 mt-md-5" style="background-color: white;">
					<div id="main-wrap">
						<div id="dashboard-left" class="dashboarddiv-short"></div>
						<div id="dashboard-right" class="dashboarddiv-short">
							<div id="schedule-stats-div"></div>
							<div id="current-day-div"></div>
						</div>
						<div id="dashboarddiv-wide"></div>
					</div>
				</main>
			</div>
		</div>
		<div id="popupdiv" style="min-width:70%;"></div>
		
		<script src="https://code.jquery.com/jquery-3.3.1.min.js"></script>
		<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.min.js"></script>
		<script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.14.3/umd/popper.min.js"></script>
		<script src="https://stackpath.bootstrapcdn.com/bootstrap/4.1.3/js/bootstrap.min.js"></script>
		<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery.form/4.2.2/jquery.form.min.js"></script>
		<script src="https://cdn.datatables.net/v/bs4/dt-1.10.21/fh-3.1.7/datatables.min.js"></script>
		<script src="chosen.jquery.min.js?v=1.1"></script>
		<script src="jquery.bpopup.min.js"></script>
		<script src="moment.min.js"></script>
		<script src="jquery.circliful.min.js"></script>
		<script src="pace.min.js?v=1.1"></script>
		<script src="pulsejs.js?v=2.1"></script>
	</body>
</html>
<%
	SQLstmt = "INSERT INTO OPS_HIT(OPS_HIT_OPS_USR_ID,OPS_HIT_DATE,OPS_HIT_TIME,OPS_HIT_LOAD_TIME,OPS_HIT_PAGE) VALUES (?,TO_DATE(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE)),TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'HH24:MI'),0,LOWER(?))"
	cmd.CommandText = SQLstmt
	cmd.Parameters(0).value = PULSE_USR_ID
	cmd.Parameters(1).value = Request.ServerVariables("SCRIPT_NAME")
	Set RSI = cmd.Execute
	Set RSI = Nothing 
%>
<% Response.CacheControl = "no-cache" %>
<% Response.AddHeader "Pragma", "no-cache" %>
<% Response.Expires = -1 %>
<% Set cmd = Nothing %>
<% Conn.Close %>
<% Set Conn = Nothing %> 
