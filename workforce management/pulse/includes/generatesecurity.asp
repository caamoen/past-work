<!--#include file="pulseheader.asp"-->

<%
	If Request.Querystring("DEPARTMENT") <> "" Then
		PARAMETER_DEPARTMENT = Request.Querystring("DEPARTMENT")
	Else
		PARAMETER_DEPARTMENT = ""
	End If
	If Request.Querystring("TEAM") <> "" Then
		If PARAMETER_DEPARTMENT = "RES" and Request.Querystring("TEAM") = "NEW" Then
			PARAMETER_TEAM = "SES"
		Else
			PARAMETER_TEAM = Request.Querystring("TEAM") 
		End If
	Else
		PARAMETER_TEAM = ""
	End If
	If Request.Querystring("JOB") <> "" Then
		PARAMETER_JOB = Request.Querystring("JOB")
	Else
		PARAMETER_JOB = ""
	End If
	If Request.Querystring("MATCH") <> "" Then
		PARAMETER_MATCH = Request.Querystring("MATCH")
	Else
		PARAMETER_MATCH = "-1"
	End If
	If PARAMETER_MATCH = "-1" Then
		SQLstmt = "SELECT " & _
		"TYPE_ID, " & _
		"ACCESS_ID, " & _
		"DECODE(USER_COUNT,0,'N',CASE WHEN ACCESS_COUNT/USER_COUNT >= .75 THEN 1 ELSE 0 END) ACCESS_FLAG " & _
		"FROM " & _
		"( " & _
			"SELECT " & _
			"TYPE_ID, " & _
			"ACCESS_ID, " & _
			"USER_COUNT, " & _
			"CASE WHEN " & _
				"? = 'DIR' " & _
				"OR " & _
				"( " & _
					"? = 'ADM' " & _
					"AND ? = 'OPS' " & _
				") " & _
				"OR " & _
				"( " & _
					"? = 'MGR' " & _
					"AND ? = SUBSTR(ACCESS_ADDRESS,1,3) " & _
					"AND TYPE_ID = 132 " & _
				") " & _
				"OR " & _
				"( " & _
					"? ='ANL' " & _
					"AND ? = 'OPS' " & _
					"AND TYPE_ID = 132 " & _
				") " & _
				"THEN USER_COUNT " & _
				"ELSE ACCESS_COUNT " & _
			"END ACCESS_COUNT " & _
			"FROM " & _
			"( " & _
				"SELECT " & _
				"PAGE.TYPE_ID, " & _
				"PAGE.ACCESS_ID, " & _
				"PAGE.ACCESS_ADDRESS, " & _
				"MAX(SECURITY.USER_COUNT) OVER () USER_COUNT, " & _
				"COUNT(SECURITY.ACCESS_ID) ACCESS_COUNT " & _
				"FROM " & _
				"( " & _
					"SELECT " & _
					"50 TYPE_ID, " & _
					"ACT.SYS_CDD_ID ACCESS_ID, " & _
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
					"OPS_RPM_TYPE || ' ' || OPS_RPM_ID " & _
					"FROM OPS_REPORT_MASTER " & _
					"WHERE OPS_RPM_STAND_ALONE = 'Y' " & _
					"AND OPS_RPM_STATUS = 'ACT' " & _
				")PAGE " & _
				"LEFT JOIN " & _
				"( " & _
					"SELECT " & _
					"SYS_CDD_SYS_CDM_ID TYPE_ID, " & _
					"SYS_CDD_VALUE ACCESS_ID, " & _
					"OPS_USD_OPS_USR_ID OPS_USR_ID, " & _
					"COUNT(DISTINCT OPS_USD_OPS_USR_ID) OVER () USER_COUNT " & _
					"FROM OPS_USER_DETAIL " & _
					"JOIN SYS_CODE_DETAIL " & _
					"ON SYS_CDD_SYS_CDM_ID IN (50,132) " & _
					"AND OPS_USD_OPS_USR_ID = SYS_CDD_NAME " & _
					"WHERE TO_DATE(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE)) BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE " & _
					"AND OPS_USD_LOCATION IN ('MOT','WFH','WFD') " & _
					"AND OPS_USD_TYPE = ? " & _
					"AND OPS_USD_TEAM = ? " & _
					"AND OPS_USD_JOB = ? " & _
				")SECURITY " & _
				"ON PAGE.TYPE_ID = SECURITY.TYPE_ID " & _
				"AND PAGE.ACCESS_ID = SECURITY.ACCESS_ID " & _
				"GROUP BY PAGE.TYPE_ID, PAGE.ACCESS_ID, PAGE.ACCESS_ADDRESS, SECURITY.USER_COUNT " & _
			") " & _
		")"
		cmd.CommandText = SQLstmt
		cmd.Parameters(0).value = PARAMETER_JOB
		cmd.Parameters(1).value = PARAMETER_JOB
		cmd.Parameters(2).value = PARAMETER_DEPARTMENT
		cmd.Parameters(3).value = PARAMETER_JOB
		cmd.Parameters(4).value = PARAMETER_DEPARTMENT
		cmd.Parameters(5).value = PARAMETER_JOB
		cmd.Parameters(6).value = PARAMETER_DEPARTMENT
		cmd.Parameters(7).value = PARAMETER_DEPARTMENT
		cmd.Parameters(8).value = PARAMETER_TEAM
		cmd.Parameters(9).value = PARAMETER_JOB
	Else
		SQLstmt = "SELECT " & _
		"PAGE.TYPE_ID, " & _
		"PAGE.ACCESS_ID, " & _
		"PAGE.ACCESS_ADDRESS, " & _
		"NVL2(SECURITY.ACCESS_ID,1,0) ACCESS_FLAG " & _
		"FROM " & _
		"( " & _
			"SELECT " & _
			"50 TYPE_ID, " & _
			"ACT.SYS_CDD_ID ACCESS_ID, " & _
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
		"AND PAGE.ACCESS_ID = SECURITY.ACCESS_ID"
		cmd.CommandText = SQLstmt
		cmd.Parameters(0).value = PARAMETER_MATCH
	End If
	Set RSPROFILE = cmd.Execute
%>
	[
	<% Do While Not RSPROFILE.EOF %>
		{
		"typeId": "<%=RSPROFILE("TYPE_ID")%>",
		"accessId": "<%=RSPROFILE("ACCESS_ID")%>",
		"accessFlag": "<%=RSPROFILE("ACCESS_FLAG")%>"
		}
		<% RSPROFILE.MoveNext %>
		<% If Not RSPROFILE.EOF Then %>
			,
		<% End If %>
	<% Loop %>
	]
<!--#include file="pulsefunctions.asp"-->
<% Set cmd = Nothing %>
<% Conn.Close %>
<% Set Conn = Nothing %>