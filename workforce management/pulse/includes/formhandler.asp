<!--#include file="pulseheader.asp"-->
<%
	SQLstmt = "INSERT INTO OPS_HIT(OPS_HIT_OPS_USR_ID, OPS_HIT_DATE, OPS_HIT_TIME, OPS_HIT_PARAMETERS, OPS_HIT_LOAD_TIME, OPS_HIT_PAGE) VALUES (?, TO_DATE(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE)), TO_CHAR(CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),'HH24:MI'), ?, 0, LOWER(?))"
	cmd.CommandText = SQLstmt
	cmd.Parameters(0).value = PULSE_USR_ID
	cmd.Parameters(1).value = Left(PulseURIDecoder(Request.Form),255)
	cmd.Parameters(2).value = Request.ServerVariables("SCRIPT_NAME")
	Set RSI = cmd.Execute
	Set RSI = Nothing

	PARAMETER_DATE = Request.Form("FORM_DATE")
	If Request.Form("TRADE_ACTION") <> "" Then
%>
		<!--#include file="tradesql.asp"-->
<%
		USE_ARRAY = Split(Request.Form("TRADE_ACTION"),",")
		For Each TRADE_ITEM in USE_ARRAY
			TRADE_ARRAY = Split(TRADE_ITEM,"_")
			TRM_ID = Trim(TRADE_ARRAY(0))
			TRD_ID = Trim(TRADE_ARRAY(1))
			TRADE_STATUS = Trim(TRADE_ARRAY(2))
			
			If TRADE_STATUS <> "DEL" Then
				Call TradeEmail(TRM_ID, TRD_ID, TRADE_STATUS)
			End If
			
			If TRADE_STATUS <> "COM" Then
				SQLstmt = "DELETE FROM OPS_TRADE_COMPLETE " & _
				"WHERE OPS_TRC_OPS_TRD_ID = ?"
				cmd.CommandText = SQLstmt
				cmd.Parameters(0).value = TRD_ID
				Set RSD = cmd.Execute
				Set RSD = Nothing
			End If
			
			SQLstmt = "UPDATE OPS_TRADE_MASTER " & _
			"SET OPS_TRM_STATUS = ? " & _
			"WHERE OPS_TRM_ID = ? "
			cmd.CommandText = SQLstmt
			cmd.Parameters(0).value = TRADE_STATUS
			cmd.Parameters(1).value = TRM_ID
			Set RSU = cmd.Execute
			Set RSU = Nothing

			SQLstmt = "UPDATE OPS_TRADE_DETAIL " & _
			"SET OPS_TRD_STATUS = DECODE(?,'COM',DECODE(OPS_TRD_ID,?,'COM','DEL'),?) " & _
			"WHERE OPS_TRD_OPS_TRM_ID = ? "
			cmd.CommandText = SQLstmt
			cmd.Parameters(0).value = TRADE_STATUS
			cmd.Parameters(1).value = TRD_ID
			cmd.Parameters(2).value = TRADE_STATUS
			cmd.Parameters(3).value = TRM_ID
			Set RSU = cmd.Execute
			Set RSU = Nothing
		Next
	End If
	If Request.Form("SCHEDULEID_LIST") <> "" Then		
		NOTIFY_LIST = ""
		ID_ARRAY = Split(Request.Form("SCHEDULEID_LIST"),",")
		For Each USE_ID in ID_ARRAY
			USE_ID = CLng(Trim(USE_ID))
			
			USE_USER = CLng(Request.Form("SCIUSER_" & USE_ID))
			USE_STATUS = Request.Form("SCISTATUS_" & USE_ID)
			If Request.Form("SCISTART_" & USE_ID) = "24:00" Then
				USE_START = CDate(Request.Form("SCIDATE_" & USE_ID)) + 1
			Else
				USE_START = CDate(Request.Form("SCIDATE_" & USE_ID) & " " & Request.Form("SCISTART_" & USE_ID))
			End If
			If Request.Form("SCIEND_" & USE_ID) = "24:00" Then
				USE_END = CDate(Request.Form("SCIDATE_" & USE_ID)) + 1
			Else
				USE_END = CDate(Request.Form("SCIDATE_" & USE_ID) & " " & Request.Form("SCIEND_" & USE_ID))
			End If	
			USE_TYPE = Request.Form("SCITYPE_" & USE_ID)
			USE_USRTYPE = Request.Form("SCIUSRTYPE_" & USE_ID)
			If Trim(Request.Form("SCINOTES_" & USE_ID)) <> "" Then 
				USE_NOTES = Trim(Request.Form("SCINOTES_" & USE_ID))
			Else
				USE_NOTES = "NULL"
			End If
 
			If USE_ID > 0 Then
				If CDate(FormatDateTime(USE_START,2)) = Date and Left(USE_TYPE,2) = "CD" and USE_STATUS = "APP" Then
					SQLstmt = "SELECT " & _
					"OPS_SCI_STATUS, " & _
					"OPS_SCI_TYPE " & _
					"FROM OPS_SCHEDULE_INFO " & _
					"WHERE OPS_SCI_ID = ?"
					cmd.CommandText = SQLstmt
					cmd.Parameters(0).value = USE_ID
					Set RSID = cmd.Execute
					If Not RSID.EOF Then
						If Left(RSID("OPS_SCI_TYPE"),2) <> "CD" or RSID("OPS_SCI_STATUS") <> "APP" Then
							NOTIFY_LIST = NOTIFY_LIST & USE_USER & ","
						End If
					End If
					Set RSID = Nothing
				End If
			
				SQLstmt = "UPDATE OPS_SCHEDULE_INFO " & _
				"SET OPS_SCI_OPS_USR_ID = ?, " & _
				"OPS_SCI_STATUS = ?, " & _
				"OPS_SCI_START = ?, " & _
				"OPS_SCI_END = ?, " & _
				"OPS_SCI_TYPE = ?, " & _
				"OPS_SCI_OPS_USR_TYPE = ?, " & _
				"OPS_SCI_NOTES = NULLIF(?,'NULL') " & _
				"WHERE OPS_SCI_ID = ?"
				cmd.CommandText = SQLstmt
				cmd.Parameters(0).value = USE_USER
				cmd.Parameters(1).value = USE_STATUS
				cmd.Parameters(2).value = USE_START
				cmd.Parameters(3).value = USE_END
				cmd.Parameters(4).value = USE_TYPE
				cmd.Parameters(5).value = USE_USRTYPE
				cmd.Parameters(6).value = USE_NOTES
				cmd.Parameters(7).value = USE_ID
				Set RSU = cmd.Execute
				Set RSU = Nothing			
			ElseIf USE_STATUS <> "DEL" Then
				If CDate(FormatDateTime(USE_START,2)) = Date and Left(USE_TYPE,2) = "CD" and USE_STATUS = "APP" Then
					NOTIFY_LIST = NOTIFY_LIST & USE_USER & ","
				End If
				SQLstmt = "INSERT INTO OPS_SCHEDULE_INFO(OPS_SCI_OPS_USR_ID, OPS_SCI_STATUS, OPS_SCI_START, OPS_SCI_END, OPS_SCI_TYPE, OPS_SCI_OPS_USR_TYPE, OPS_SCI_NOTES, INSERT_DATE, OPS_SCI_INS_USER) VALUES (?,?,?,?,?,?,NULLIF(?,'NULL'),CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),?)"
				cmd.CommandText = SQLstmt
				cmd.Parameters(0).value = USE_USER
				cmd.Parameters(1).value = USE_STATUS
				cmd.Parameters(2).value = USE_START
				cmd.Parameters(3).value = USE_END
				cmd.Parameters(4).value = USE_TYPE
				cmd.Parameters(5).value = USE_USRTYPE
				cmd.Parameters(6).value = USE_NOTES
				cmd.Parameters(7).value = PULSE_USR_ID
				Set RSI = cmd.Execute
				Set RSI = Nothing	
			End If
		Next
		If NOTIFY_LIST <> "" Then
			NOTIFY_LIST = Left(NOTIFY_LIST,Len(NOTIFY_LIST)-1)
			
			Set myMail = CreateObject("CDO.Message")
			myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 1
			myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = "c:\Inetpub\mailroot\pickup"
			myMail.Configuration.Fields.Update
			
			SQLstmt = "SELECT " & _
			"OPS_USR_NAME, " & _
			"OPS_USR_EMAIL_ADDR || ';' || OPS_USR_ALT_ID_2 MAIL_TO " & _
			"FROM OPS_USER " & _
			"WHERE OPS_USR_ID IN (" & NOTIFY_LIST & ") " & _
			"ORDER BY OPS_USR_NAME"
			cmd.CommandText = SQLstmt
			Set RSMAIL = cmd.Execute
			Do While Not RSMAIL.EOF
				myMail.Subject = "New Schedule Adjustment for " & RSMAIL("OPS_USR_NAME")
				myMail.HTMLBody = "Your schedule has been adjusted for " & Date & ". See your Dashboard at https://opscenter.mltvacations.com for your current schedule and please advise the Ops Desk if your lunch requires an adjustment."
				myMail.To = RSMAIL("MAIL_TO")
				myMail.Bcc = "MinotOperationsDesk@deltavacations.com;"
				myMail.From = "noreply@deltavacations.com;"
				myMail.Send
				RSMAIL.MoveNext
			Loop
			Set RSMAIL = Nothing
			Set myMail = Nothing
		End If
	End If
	If Request.Form("ACKNOWLEDGE_ERROR") <> "" Then
		USE_ARRAY = Split(Request.Form("ACKNOWLEDGE_ERROR"),",")
		For Each ERROR_ITEM in USE_ARRAY
			ERROR_ARRAY = Split(ERROR_ITEM,"_")
			ERROR_USER = Trim(ERROR_ARRAY(0))
			ERROR_CODE = Trim(ERROR_ARRAY(1))

			SQLstmt = "INSERT INTO RES_DAILY_STATS_NOTES(RES_DLN_OPS_USR_ID, RES_DLN_DATE, RES_DLN_TIME, RES_DLN_TYPE, RES_DLN_TEXT, RES_DLN_OPS_USR_TYPE) VALUES (?,?,CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),?,?,?)"
			cmd.CommandText = SQLstmt
			cmd.Parameters(0).value = PULSE_USR_ID
			cmd.Parameters(1).value = PARAMETER_DATE
			cmd.Parameters(2).value = ERROR_CODE
			cmd.Parameters(3).value = ERROR_USER
			cmd.Parameters(4).value = AgentDepartment(ERROR_USER)
			Set RSI = cmd.Execute
			Set RSI = Nothing
		Next
	End If
	If Request.Form("DELETE_ERROR") <> "" Then
		DELETE_ARRAY = Split(Request.Form("DELETE_ERROR"),",")
		For Each ERROR_ID in DELETE_ARRAY
			SQLstmt = "DELETE FROM RES_DAILY_STATS_NOTES WHERE RES_DLN_ID = ?"
			cmd.CommandText = SQLstmt
			cmd.Parameters(0).value = Trim(ERROR_ID)
			Set RSD = cmd.Execute
			Set RSD = Nothing			
		Next
	End If
	If Request.Form("NOTEID_LIST") <> "" Then
		ID_ARRAY = Split(Request.Form("NOTEID_LIST"),",")
		For Each USE_ID in ID_ARRAY
			USE_ID = CLng(Trim(USE_ID))
			If Trim(Request.Form("NOTETEXT_" & USE_ID)) <> "" Then 
				USE_NOTES = Trim(Request.Form("NOTETEXT_" & USE_ID))
			Else
				USE_NOTES = "NULL"
			End If 	
			If USE_ID > 0 Then
				SQLstmt = "UPDATE RES_DAILY_STATS_NOTES " & _
				"SET RES_DLN_TEXT = TRIM(REGEXP_SUBSTR(RES_DLN_TEXT,'[^-]+',1,1)) || NULLIF(' - ' || NULLIF(?,'NULL'),' - ') " & _
				"WHERE RES_DLN_ID = ?"
				cmd.CommandText = SQLstmt
				cmd.Parameters(0).value = USE_NOTES
				cmd.Parameters(1).value = USE_ID
				Set RSU = cmd.Execute
				Set RSU = Nothing			
			ElseIf Request.Form("NOTEAGT_" & USE_ID) <> "-1" and Request.Form("NOTECODE_" & USE_ID) <> "-1" Then
				USE_CODE = Request.Form("NOTECODE_" & USE_ID)
				USE_AGENT = Request.Form("NOTEAGT_" & USE_ID)
				
				If USE_NOTES <> "NULL" Then
					USE_NOTES = USE_AGENT & " - " & USE_NOTES
				Else
					USE_NOTES = USE_AGENT
				End If
				SQLstmt = "INSERT INTO RES_DAILY_STATS_NOTES(RES_DLN_OPS_USR_ID, RES_DLN_DATE, RES_DLN_TIME, RES_DLN_TYPE, RES_DLN_TEXT, RES_DLN_OPS_USR_TYPE) VALUES (?,?,CAST(SYSTIMESTAMP AT TIME ZONE 'US/CENTRAL' AS DATE),?,REPLACE(?,'''',CHR(39)),?)"
				cmd.CommandText = SQLstmt
				cmd.Parameters(0).value = PULSE_USR_ID
				cmd.Parameters(1).value = PARAMETER_DATE
				cmd.Parameters(2).value = USE_CODE
				cmd.Parameters(3).value = USE_NOTES
				cmd.Parameters(4).value = AgentDepartment(USE_AGENT)
				Set RSI = cmd.Execute
				Set RSI = Nothing
			End If
		Next
	End If
	If Request.Form("CIRCLEID_LIST") <> "" Then		
		ID_ARRAY = Split(Request.Form("CIRCLEID_LIST"),",")
		For Each USE_ID in ID_ARRAY
			USE_ID = CLng(Trim(USE_ID))
			USE_REQUEST_TYPE = Request.Form("CIRCLEREQUEST_" & USE_ID)
			USE_DOTW = Replace(Request.Form("CIRCLEDOTW_" & USE_ID),", ","")
			USE_COLOR = Request.Form("CIRCLECOLOR_" & USE_ID)
			USE_START = CInt(Request.Form("CIRCLESTART_" & USE_ID))
			USE_END = CInt(Request.Form("CIRCLEEND_" & USE_ID))
			USE_EFF_DATE = CDate(Request.Form("CIRCLEEFFDATE_" & USE_ID))
			USE_DIS_DATE = CDate(Request.Form("CIRCLEDISDATE_" & USE_ID))

			If USE_ID > 0 Then
				If USE_DOTW <> "" and USE_START < USE_END and USE_EFF_DATE <= USE_DIS_DATE Then
					USE_VALUE = USE_DOTW & ";" & USE_COLOR & ";" & USE_START & ";" & USE_END
					SQLstmt = "UPDATE OPS_PARAMETER " & _
					"SET OPS_PAR_VALUE = ?, " & _
					"OPS_PAR_EFF_DATE = ?, " & _
					"OPS_PAR_DIS_DATE = ? " & _
					"WHERE OPS_PAR_ID = ?"
					cmd.CommandText = SQLstmt
					cmd.Parameters(0).value = USE_VALUE
					cmd.Parameters(1).value = USE_EFF_DATE
					cmd.Parameters(2).value = USE_DIS_DATE
					cmd.Parameters(3).value = USE_ID
					Set RSU = cmd.Execute
					Set RSU = Nothing
				Else
					SQLstmt = "DELETE FROM OPS_PARAMETER WHERE OPS_PAR_ID = ?"
					cmd.CommandText = SQLstmt
					cmd.Parameters(0).value = USE_ID
					Set RSD = cmd.Execute
					Set RSD = Nothing
				End If
			Else
				If USE_DOTW <> "" and USE_START < USE_END and USE_EFF_DATE <= USE_DIS_DATE Then
					USE_VALUE = USE_DOTW & ";" & USE_COLOR & ";" & USE_START & ";" & USE_END
					SQLstmt = "INSERT INTO OPS_PARAMETER(OPS_PAR_PARENT_ID, OPS_PAR_PARENT_TYPE, OPS_PAR_CODE, OPS_PAR_VALUE, OPS_PAR_EFF_DATE, OPS_PAR_DIS_DATE) VALUES (0,'STF',?,?,?,?)"
					cmd.CommandText = SQLstmt
					cmd.Parameters(0).value = USE_REQUEST_TYPE
					cmd.Parameters(1).value = USE_VALUE
					cmd.Parameters(2).value = USE_EFF_DATE
					cmd.Parameters(3).value = USE_DIS_DATE
					Set RSI = cmd.Execute
					Set RSI = Nothing
				End If
			End If
		Next
	End If
	If Request.Form("CONTROLID_LIST") <> "" Then		
		ID_ARRAY = Split(Request.Form("CONTROLID_LIST"),",")
		For Each USE_ID in ID_ARRAY
			USE_ID = CLng(Trim(USE_ID))
			USE_WORKGROUP = Request.Form("CONTROLWORKGROUP_" & USE_ID)
			USE_CONTROL_TYPE = Request.Form("CONTROLTYPE_" & USE_ID)
			USE_CONTROL_FIELDS = Split(Request.Form("CONTROLFIELDS_" & USE_ID),",")
			VALID_BOOL = 1
			USE_VALUE = ""
			For Each CONTROL_FIELD in USE_CONTROL_FIELDS
				If CONTROL_FIELD = "DOTW" Then
					USE_VALUE = USE_VALUE & Replace(Request.Form("CONTROLDOTW_" & USE_ID),", ","") & ";"
					If Replace(Request.Form("CONTROLDOTW_" & USE_ID),", ","") = "" Then
						VALID_BOOL = 0
					End If
				Elseif CONTROL_FIELD = "INTERVAL" Then
					USE_VALUE = USE_VALUE & Request.Form("CONTROLSTART_INTERVAL_" & USE_ID) & ";"
					USE_VALUE = USE_VALUE & Request.Form("CONTROLEND_INTERVAL_" & USE_ID) & ";"
					If Request.Form("CONTROLSTART_INTERVAL_" & USE_ID) = Request.Form("CONTROLEND_INTERVAL_" & USE_ID) Then
						VALID_BOOL = 0
					End If
				Elseif CONTROL_FIELD = "SCHEDULED_HOURS" Then
					USE_VALUE = USE_VALUE & Request.Form("CONTROLSTART_HOURS_" & USE_ID) & ";"
					USE_VALUE = USE_VALUE & Request.Form("CONTROLEND_HOURS_" & USE_ID) & ";"
					If Request.Form("CONTROLSTART_HOURS_" & USE_ID) = Request.Form("CONTROLEND_HOURS_" & USE_ID) Then
						VALID_BOOL = 0
					End If
				Elseif CONTROL_FIELD = "EVAL_SCORE" Then
					USE_VALUE = USE_VALUE & Request.Form("CONTROLSTART_SCORE_" & USE_ID) & ";"
					USE_VALUE = USE_VALUE & Request.Form("CONTROLEND_SCORE_" & USE_ID) & ";"
					If Request.Form("CONTROLSTART_SCORE_" & USE_ID) = Request.Form("CONTROLEND_SCORE_" & USE_ID) Then
						VALID_BOOL = 0
					End If
				Elseif CONTROL_FIELD = "REDUCE_HOURS" or CONTROL_FIELD = "ADD_HOURS" or CONTROL_FIELD = "NEW_DAYS" Then
					USE_VALUE = USE_VALUE & Request.Form("CONTROL_VALUE_" & USE_ID) & ";"
				End If
			Next
			USE_VALUE = Left(USE_VALUE,Len(USE_VALUE)-1)
			USE_EFF_DATE = CDate(Request.Form("CONTROLEFFDATE_" & USE_ID))
			USE_DIS_DATE = CDate(Request.Form("CONTROLDISDATE_" & USE_ID))
			If USE_EFF_DATE > USE_DIS_DATE Then
				VALID_BOOL = 0
			End If			
			If USE_ID > 0 Then
				If VALID_BOOL = 1 Then
					SQLstmt = "UPDATE OPS_PARAMETER " & _
					"SET OPS_PAR_VALUE = ?, " & _
					"OPS_PAR_EFF_DATE = ?, " & _
					"OPS_PAR_DIS_DATE = ? " & _
					"WHERE OPS_PAR_ID = ?"
					cmd.CommandText = SQLstmt
					cmd.Parameters(0).value = USE_VALUE
					cmd.Parameters(1).value = USE_EFF_DATE
					cmd.Parameters(2).value = USE_DIS_DATE
					cmd.Parameters(3).value = USE_ID
					Set RSU = cmd.Execute
					Set RSU = Nothing
				Else
					SQLstmt = "DELETE FROM OPS_PARAMETER WHERE OPS_PAR_ID = ?"
					cmd.CommandText = SQLstmt
					cmd.Parameters(0).value = USE_ID
					Set RSD = cmd.Execute
					Set RSD = Nothing
				End If
			Else
				If VALID_BOOL = 1 Then
					SQLstmt = "INSERT INTO OPS_PARAMETER(OPS_PAR_PARENT_ID, OPS_PAR_PARENT_TYPE, OPS_PAR_CODE, OPS_PAR_VALUE, OPS_PAR_EFF_DATE, OPS_PAR_DIS_DATE) VALUES (0,?,?,?,?,?)"
					cmd.CommandText = SQLstmt
					cmd.Parameters(0).value = USE_WORKGROUP
					cmd.Parameters(1).value = USE_CONTROL_TYPE
					cmd.Parameters(2).value = USE_VALUE
					cmd.Parameters(3).value = USE_EFF_DATE
					cmd.Parameters(4).value = USE_DIS_DATE
					Set RSI = cmd.Execute
					Set RSI = Nothing
				End If
			End If
		Next
	End If
	If Request.Form("ADMINID_LIST") <> "" Then		
		ID_ARRAY = Split(Request.Form("ADMINID_LIST"),",")
		ReDim NEWADMIN_ARRAY(1,49)
		NEWADMIN_COUNTER = -1
		For Each ARRAY_ITEM in ID_ARRAY
			ADMIN_ARRAY = Split(ARRAY_ITEM,"_")
			ADMIN_TYPE = Trim(ADMIN_ARRAY(0))
			USE_ID = CLng(Trim(ADMIN_ARRAY(1)))
			If ADMIN_TYPE = "MASTER" Then
				MASTER_ID = USE_ID
				DETAILS_FOUND = -1
				
				USE_DETAILS_ID = 0
				DETAILS_DATE = ""
				ADMIN_USR_EFF_DATE = ""
				ADMIN_USR_DIS_DATE = ""
				
				If Trim(Request.Form("ADMINUSER_" & MASTER_ID)) <> "" Then
					ADMIN_USER = Trim(Request.Form("ADMINUSER_" & MASTER_ID))
				Else
					ADMIN_USER = "NULL"
				End If
				If Trim(Request.Form("ADMINWINDOWS_" & MASTER_ID)) <> "" Then
					ADMIN_WINDOWS = Trim(Request.Form("ADMINWINDOWS_" & MASTER_ID))
				Else
					ADMIN_WINDOWS = "MLTMTKA\"
				End If
				If Trim(Request.Form("ADMINPPR_" & MASTER_ID)) <> "" Then
					ADMIN_PPR = Trim(Request.Form("ADMINPPR_" & MASTER_ID))
				Else
					ADMIN_PPR = "NULL"
				End If 
				If Trim(Request.Form("ADMINNAVIGATOR_" & MASTER_ID)) <> "" Then
					ADMIN_NAVIGATOR = Trim(Request.Form("ADMINNAVIGATOR_" & MASTER_ID))
				Else
					ADMIN_NAVIGATOR = "NULL"
				End If
				If Trim(Request.Form("ADMINBADGE_" & MASTER_ID)) <> "" Then
					ADMIN_BADGE = Trim(Request.Form("ADMINBADGE_" & MASTER_ID))
				Else
					ADMIN_BADGE = "0"
				End If
				If Trim(Request.Form("ADMINEXT_" & MASTER_ID)) <> "" Then
					ADMIN_EXT = Trim(Request.Form("ADMINEXT_" & MASTER_ID))
				Else
					ADMIN_EXT = "0"
				End If
				If Trim(Request.Form("ADMINTEXT_" & MASTER_ID)) <> "" Then
					ADMIN_TEXT = Trim(Request.Form("ADMINTEXT_" & MASTER_ID))
				Else
					ADMIN_TEXT = "NULL"
				End If
				If Trim(Request.Form("ADMINEMAIL_" & MASTER_ID)) <> "" Then
					ADMIN_EMAIL = Trim(Request.Form("ADMINEMAIL_" & MASTER_ID))
				Else
					ADMIN_EMAIL = "@deltavacations.com"
				End If 
				For Each FIELD in Request.Form
					If Left(FIELD,15) = "ADMINDETAILUSER" and FIELD <> "ADMINDETAILUSER_0" and CStr(Request.Form(FIELD)) = CStr(MASTER_ID) Then
						USD_ARRAY = Split(FIELD,"_")
						DETAILS_ID = CLng(USD_ARRAY(1))
						
						If DETAILS_FOUND = -1 and CDate(Request.Form("ADMINDISDATE_" & DETAILS_ID)) >= CDate(Request.Form("ADMINEFFDATE_" & DETAILS_ID)) Then
							DETAILS_FOUND = 0
						End If
						
						If Date >= CDate(Request.Form("ADMINEFFDATE_" & DETAILS_ID)) and Date <= CDate(Request.Form("ADMINDISDATE_" & DETAILS_ID)) Then
							USE_DETAILS_ID = DETAILS_ID
							DETAILS_FOUND = 1
						Elseif Date < CDate(Request.Form("ADMINEFFDATE_" & DETAILS_ID)) and (DETAILS_DATE = "" or CDate(Request.Form("ADMINEFFDATE_" & DETAILS_ID)) < DETAILS_DATE) and CDate(Request.Form("ADMINDISDATE_" & DETAILS_ID)) >= CDate(Request.Form("ADMINEFFDATE_" & DETAILS_ID)) and DETAILS_FOUND = 0 Then
							DETAILS_DATE = CDate(Request.Form("ADMINEFFDATE_" & DETAILS_ID))
							USE_DETAILS_ID = DETAILS_ID
						Elseif Date > CDate(Request.Form("ADMINDISDATE_" & DETAILS_ID)) and (DETAILS_DATE = "" or CDate(Request.Form("ADMINDISDATE_" & DETAILS_ID)) > DETAILS_DATE or DETAILS_DATE > Date) and CDate(Request.Form("ADMINDISDATE_" & DETAILS_ID)) >= CDate(Request.Form("ADMINEFFDATE_" & DETAILS_ID)) and DETAILS_FOUND = 0 Then
							DETAILS_DATE = CDate(Request.Form("ADMINDISDATE_" & DETAILS_ID))
							USE_DETAILS_ID = DETAILS_ID
						End If

						If (ADMIN_USR_EFF_DATE = "" or CDate(Request.Form("ADMINEFFDATE_" & DETAILS_ID)) < ADMIN_USR_EFF_DATE) and CDate(Request.Form("ADMINDISDATE_" & DETAILS_ID)) >= CDate(Request.Form("ADMINEFFDATE_" & DETAILS_ID)) Then
							ADMIN_USR_EFF_DATE = CDate(Request.Form("ADMINEFFDATE_" & DETAILS_ID))
						End If
						If (ADMIN_USR_DIS_DATE = "" or CDate(Request.Form("ADMINDISDATE_" & DETAILS_ID)) > ADMIN_USR_DIS_DATE) and CDate(Request.Form("ADMINDISDATE_" & DETAILS_ID)) >= CDate(Request.Form("ADMINEFFDATE_" & DETAILS_ID)) Then
							ADMIN_USR_DIS_DATE = CDate(Request.Form("ADMINDISDATE_" & DETAILS_ID))
						End If
					End If
				Next
				If DETAILS_FOUND = -1 or USE_DETAILS_ID <> 0 Then
					If USE_DETAILS_ID <> 0 Then
						ADMIN_DEPT = Request.Form("ADMINDEPT_" & USE_DETAILS_ID)
						ADMIN_TEAM = Request.Form("ADMINTEAM_" & USE_DETAILS_ID)
						ADMIN_JOB = Request.Form("ADMINJOB_" & USE_DETAILS_ID)
						ADMIN_CLASS = Request.Form("ADMINCLASS_" & USE_DETAILS_ID)
						ADMIN_SUPERVISOR = Request.Form("ADMINSUPERVISOR_" & USE_DETAILS_ID)
						If Trim(Request.Form("ADMINPHONE_" & USE_DETAILS_ID)) <> "" and Date >= CDate(Request.Form("ADMINEFFDATE_" & USE_DETAILS_ID)) and Date <= CDate(Request.Form("ADMINDISDATE_" & USE_DETAILS_ID)) Then
							ADMIN_PHONE = Trim(Request.Form("ADMINPHONE_" & USE_DETAILS_ID))
						Else
							ADMIN_PHONE = "0"
						End If
					End If
					
					If MASTER_ID > 0 Then
						If DETAILS_FOUND = -1 Then
							SQLstmt = "UPDATE OPS_USER " & _
							"SET OPS_USR_NT_ID = ?, " & _
							"OPS_USR_SUN_ID = NULLIF(?,'NULL'), " & _
							"OPS_USR_PHN_ID_PC = ?, " & _
							"OPS_USR_PHN_EXT = ?, " & _
							"OPS_USR_NAME = REPLACE(?,'''',CHR(39)), " & _
							"OPS_USR_ALT_ID_1 = NULLIF(?,'NULL'), " & _
							"OPS_USR_ALT_ID_2 = NULLIF(?,'NULL'), " & _
							"OPS_USR_EMAIL_ADDR = NULLIF(?,'NULL') " & _
							"WHERE OPS_USR_ID = ?"
							cmd.CommandText = SQLstmt
							cmd.Parameters(0).value = ADMIN_WINDOWS
							cmd.Parameters(1).value = ADMIN_NAVIGATOR
							cmd.Parameters(2).value = ADMIN_BADGE
							cmd.Parameters(3).value = ADMIN_EXT
							cmd.Parameters(4).value = ADMIN_USER
							cmd.Parameters(5).value = ADMIN_PPR
							cmd.Parameters(6).value = ADMIN_TEXT
							cmd.Parameters(7).value = ADMIN_EMAIL
							cmd.Parameters(8).value = MASTER_ID
							Set RSU = cmd.Execute
							Set RSU = Nothing
						Else
							SQLstmt = "UPDATE OPS_USER " & _
							"SET OPS_USR_NT_ID = ?, " & _
							"OPS_USR_SUN_ID = NULLIF(?,'NULL'), " & _
							"OPS_USR_PHN_ID = ?, " & _
							"OPS_USR_PHN_ID_PC = ?, " & _
							"OPS_USR_PHN_EXT = ?, " & _
							"OPS_USR_NAME = REPLACE(?,'''',CHR(39)), " & _
							"OPS_USR_JOB = ?, " & _
							"OPS_USR_TEAM = ?, " & _
							"OPS_USR_SUPERVISOR = ?, " & _
							"OPS_USR_ALT_ID_1 = NULLIF(?,'NULL'), " & _
							"OPS_USR_ALT_ID_2 = NULLIF(?,'NULL'), " & _
							"OPS_USR_TYPE = ?, " & _
							"OPS_USR_CLASS = ?, " & _
							"OPS_USR_EFF_DATE = ?, " & _
							"OPS_USR_DIS_DATE = ?, " & _
							"OPS_USR_HIRE_DATE = ?, " & _
							"OPS_USR_EMAIL_ADDR = NULLIF(?,'NULL') " & _
							"WHERE OPS_USR_ID = ?"
							cmd.CommandText = SQLstmt
							cmd.Parameters(0).value = ADMIN_WINDOWS
							cmd.Parameters(1).value = ADMIN_NAVIGATOR
							cmd.Parameters(2).value = ADMIN_PHONE
							cmd.Parameters(3).value = ADMIN_BADGE
							cmd.Parameters(4).value = ADMIN_EXT
							cmd.Parameters(5).value = ADMIN_USER
							cmd.Parameters(6).value = ADMIN_JOB
							cmd.Parameters(7).value = ADMIN_TEAM
							cmd.Parameters(8).value = ADMIN_SUPERVISOR
							cmd.Parameters(9).value = ADMIN_PPR
							cmd.Parameters(10).value = ADMIN_TEXT
							cmd.Parameters(11).value = ADMIN_DEPT
							cmd.Parameters(12).value = ADMIN_CLASS
							cmd.Parameters(13).value = ADMIN_USR_EFF_DATE
							cmd.Parameters(14).value = ADMIN_USR_DIS_DATE
							cmd.Parameters(15).value = ADMIN_USR_EFF_DATE
							cmd.Parameters(16).value = ADMIN_EMAIL
							cmd.Parameters(17).value = MASTER_ID
							Set RSU = cmd.Execute
							Set RSU = Nothing
						End If
					Elseif DETAILS_FOUND <> -1 Then
						SQLstmt = "INSERT INTO OPS_USER(OPS_USR_NT_ID, OPS_USR_SUN_ID, OPS_USR_PHN_ID, OPS_USR_PHN_ID_PC, OPS_USR_PHN_EXT, OPS_USR_NAME, OPS_USR_JOB, OPS_USR_TEAM, OPS_USR_SUPERVISOR, OPS_USR_ALT_ID_1, OPS_USR_ALT_ID_2, OPS_USR_TYPE, OPS_USR_CLASS, OPS_USR_EFF_DATE, OPS_USR_DIS_DATE, OPS_USR_HIRE_DATE, OPS_USR_EMAIL_ADDR) " & _
						"VALUES(?, NULLIF(?,'NULL'), ?, ?, ?, REPLACE(?,'''',CHR(39)), ?, ?, ?, NULLIF(?,'NULL'), NULLIF(?,'NULL'), ?, ?, ?, ?, ?,NULLIF(?,'NULL'))"
						cmd.CommandText = SQLstmt
						cmd.Parameters(0).value = ADMIN_WINDOWS
						cmd.Parameters(1).value = ADMIN_NAVIGATOR
						cmd.Parameters(2).value = ADMIN_PHONE
						cmd.Parameters(3).value = ADMIN_BADGE
						cmd.Parameters(4).value = ADMIN_EXT
						cmd.Parameters(5).value = ADMIN_USER
						cmd.Parameters(6).value = ADMIN_JOB
						cmd.Parameters(7).value = ADMIN_TEAM
						cmd.Parameters(8).value = ADMIN_SUPERVISOR
						cmd.Parameters(9).value = ADMIN_PPR
						cmd.Parameters(10).value = ADMIN_TEXT
						cmd.Parameters(11).value = ADMIN_DEPT
						cmd.Parameters(12).value = ADMIN_CLASS
						cmd.Parameters(13).value = ADMIN_USR_EFF_DATE
						cmd.Parameters(14).value = ADMIN_USR_DIS_DATE
						cmd.Parameters(15).value = ADMIN_USR_EFF_DATE
						cmd.Parameters(16).value = ADMIN_EMAIL
						Set RSI = cmd.Execute
						Set RSI = Nothing
						
						SQLstmt = "SELECT " & _
						"MAX(OPS_USR_ID) MAX_ID " & _
						"FROM OPS_USER"
						cmd.CommandText = SQLstmt
						Set RSMAX = cmd.Execute
						MAX_ID = CLng(RSMAX("MAX_ID"))
						Set RSMAX = Nothing
						
						SQLstmt = "UPDATE SYS_CODE_DETAIL " & _
						"SET SYS_CDD_NAME = ? " & _
						"WHERE SYS_CDD_SYS_CDM_ID IN (50,132) " & _
						"AND SYS_CDD_NAME = ?"
						cmd.CommandText = SQLstmt
						cmd.Parameters(0).value = MAX_ID
						cmd.Parameters(1).value = MASTER_ID
						Set RSU = cmd.Execute
						Set RSU = Nothing
						
						NEWADMIN_COUNTER = NEWADMIN_COUNTER + 1
						NEWADMIN_ARRAY(0,NEWADMIN_COUNTER) = MASTER_ID
						NEWADMIN_ARRAY(1,NEWADMIN_COUNTER) = MAX_ID
					End If
				End If
			End If
		Next
		If NEWADMIN_COUNTER >= 0 Then
			Redim Preserve NEWADMIN_ARRAY(1,NEWADMIN_COUNTER)
		Else
			Erase NEWADMIN_ARRAY
		End If
		For Each ARRAY_ITEM in ID_ARRAY
			ADMIN_ARRAY = Split(ARRAY_ITEM,"_")
			ADMIN_TYPE = Trim(ADMIN_ARRAY(0))
			USE_ID = CLng(Trim(ADMIN_ARRAY(1)))
			If ADMIN_TYPE = "DETAIL" Then
				DETAILS_ID = USE_ID
				
				ADMIN_USD_EFF_DATE = ""
				ADMIN_USD_DIS_DATE = "" 
				If CDate(Request.Form("ADMINDISDATE_" & DETAILS_ID)) >= CDate(Request.Form("ADMINEFFDATE_" & DETAILS_ID)) Then
					ADMIN_USD_EFF_DATE = CDate(Request.Form("ADMINEFFDATE_" & DETAILS_ID))
					ADMIN_USD_DIS_DATE = CDate(Request.Form("ADMINDISDATE_" & DETAILS_ID))
				End If
				ADMIN_USER = CLng(Request.Form("ADMINDETAILUSER_" & DETAILS_ID))
				ADMIN_DEPT = Request.Form("ADMINDEPT_" & DETAILS_ID)
				ADMIN_TEAM = Request.Form("ADMINTEAM_" & DETAILS_ID)
				ADMIN_JOB = Request.Form("ADMINJOB_" & DETAILS_ID)
				ADMIN_CLASS = Request.Form("ADMINCLASS_" & DETAILS_ID)
				ADMIN_LOCATION = Request.Form("ADMINLOCATION_" & DETAILS_ID)
				ADMIN_HOURS = Request.Form("ADMINHOURS_" & DETAILS_ID)
				ADMIN_SUPERVISOR = Request.Form("ADMINSUPERVISOR_" & DETAILS_ID)
				If Trim(Request.Form("ADMINPHONE_" & DETAILS_ID)) <> "" Then
					ADMIN_PHONE = Trim(Request.Form("ADMINPHONE_" & DETAILS_ID))
				Else
					ADMIN_PHONE = "0"
				End If
				If Trim(Request.Form("ADMINJOBCODE_" & DETAILS_ID)) <> "" and (ADMIN_LOCATION = "MOT" or ADMIN_LOCATION = "WFH" or ADMIN_LOCATION = "WFD") and ADMIN_JOB <> "SUP" and ADMIN_JOB <> "MGR" and ADMIN_JOB <> "DIR" and ADMIN_JOB <> "ADM" and ADMIN_DEPT <> "HRA" Then
					ADMIN_JOBCODE = Trim(Request.Form("ADMINJOBCODE_" & DETAILS_ID))
				Else
					ADMIN_JOBCODE = "NULL"
				End If
				If Trim(Request.Form("ADMINPAY_" & DETAILS_ID)) <> "" and (ADMIN_LOCATION = "MOT" or ADMIN_LOCATION = "WFH" or ADMIN_LOCATION = "WFD") and ADMIN_JOB <> "SUP" and ADMIN_JOB <> "MGR" and ADMIN_JOB <> "DIR" and ADMIN_JOB <> "ADM" and ADMIN_DEPT <> "HRA" Then
					ADMIN_PAY = Round(CDbl(Trim(Request.Form("ADMINPAY_" & DETAILS_ID))),2)
				Else
					ADMIN_PAY = -1
				End If					
				If DETAILS_ID > 0 Then
					If ADMIN_USD_EFF_DATE <> "" Then
						SQLstmt = "UPDATE OPS_USER_DETAIL " & _
						"SET OPS_USD_TYPE = ?, " & _
						"OPS_USD_JOB = ?, " & _
						"OPS_USD_TEAM = ?, " & _
						"OPS_USD_SUPERVISOR = ?, " & _
						"OPS_USD_CLASS = ?, " & _
						"OPS_USD_EFF_DATE = ?, " & _
						"OPS_USD_DIS_DATE = ?, " & _
						"OPS_USD_LOCATION = ?, " & _
						"OPS_USD_PHN_ID = ?, " & _
						"OPS_USD_PAY_RATE = NULLIF(?,-1), " & _
						"OPS_USD_JOB_CODE = NULLIF(?,'NULL'), " & _
						"OPS_USD_SCH_HOURS = ? " & _
						"WHERE OPS_USD_ID = ?"
						cmd.CommandText = SQLstmt
						cmd.Parameters(0).value = ADMIN_DEPT
						cmd.Parameters(1).value = ADMIN_JOB
						cmd.Parameters(2).value = ADMIN_TEAM
						cmd.Parameters(3).value = ADMIN_SUPERVISOR
						cmd.Parameters(4).value = ADMIN_CLASS
						cmd.Parameters(5).value = ADMIN_USD_EFF_DATE
						cmd.Parameters(6).value = ADMIN_USD_DIS_DATE
						cmd.Parameters(7).value = ADMIN_LOCATION
						cmd.Parameters(8).value = ADMIN_PHONE
						cmd.Parameters(9).value = ADMIN_PAY
						cmd.Parameters(10).value = ADMIN_JOBCODE
						cmd.Parameters(11).value = ADMIN_HOURS
						cmd.Parameters(12).value = DETAILS_ID
						Set RSU = cmd.Execute
						Set RSU = Nothing
					Else
						SQLstmt = "DELETE FROM OPS_USER_DETAIL " & _
						"WHERE OPS_USD_ID = ?"
						cmd.CommandText = SQLstmt
						cmd.Parameters(0).value = DETAILS_ID
						Set RSD = cmd.Execute
						Set RSD = Nothing
					End If
				Else
					If ADMIN_USER < 0 Then
						For i = 0 to UBound(NEWADMIN_ARRAY,2)
							If ADMIN_USER = NEWADMIN_ARRAY(0,i) Then
								ADMIN_USER = NEWADMIN_ARRAY(1,i)
								Exit For
							End If
						Next
					End If
					If ADMIN_USER > 0 Then
						SQLstmt = "INSERT INTO OPS_USER_DETAIL(OPS_USD_OPS_USR_ID, OPS_USD_TYPE, OPS_USD_JOB, OPS_USD_TEAM, OPS_USD_SUPERVISOR, OPS_USD_CLASS, OPS_USD_EFF_DATE, OPS_USD_DIS_DATE, OPS_USD_LOCATION, OPS_USD_PHN_ID, OPS_USD_PAY_RATE, OPS_USD_JOB_CODE, OPS_USD_SCH_HOURS) " & _
						"VALUES (?,?,?,?,?,?,?,?,?,?,NULLIF(?,-1),NULLIF(?,'NULL'),?)"
						cmd.CommandText = SQLstmt
						cmd.Parameters(0).value = ADMIN_USER
						cmd.Parameters(1).value = ADMIN_DEPT
						cmd.Parameters(2).value = ADMIN_JOB
						cmd.Parameters(3).value = ADMIN_TEAM
						cmd.Parameters(4).value = ADMIN_SUPERVISOR
						cmd.Parameters(5).value = ADMIN_CLASS
						cmd.Parameters(6).value = ADMIN_USD_EFF_DATE
						cmd.Parameters(7).value = ADMIN_USD_DIS_DATE
						cmd.Parameters(8).value = ADMIN_LOCATION
						cmd.Parameters(9).value = ADMIN_PHONE
						cmd.Parameters(10).value = ADMIN_PAY
						cmd.Parameters(11).value = ADMIN_JOBCODE
						cmd.Parameters(12).value = ADMIN_HOURS
						Set RSI = cmd.Execute
						Set RSI = Nothing
					End If
				End If
			End If
		Next
		SQLstmt = "DELETE FROM SYS_CODE_DETAIL " & _
		"WHERE SYS_CDD_SYS_CDM_ID IN (50,132) " & _
		"AND SYS_CDD_NAME < 0"
		cmd.CommandText = SQLstmt
		Set RSD = cmd.Execute
		Set RSD = Nothing
	End If
	If Request.Form("SECURITY_AGENT") <> "" Then
		SECURITY_AGENT = Request.Form("SECURITY_AGENT")
		ID_ARRAY = Split(Request.Form("SECURITY_ACCESS"),",")
		SQLSecurity = ""
		For Each ARRAY_ITEM in ID_ARRAY
			SECURITY_ARRAY = Split(ARRAY_ITEM,"_")
			USE_TYPE = Trim(SECURITY_ARRAY(0))
			ACCESS_TYPE = Trim(SECURITY_ARRAY(1))
			SQLSecurity = SQLSecurity & "SELECT " & USE_TYPE & " TYPE_ID, " & ACCESS_TYPE & " ACCESS_ID FROM DUAL UNION ALL "
		Next
		If SQLSecurity <> "" Then 
			SQLSecurity = Left(SQLSecurity,Len(SQLSecurity)-11)
		Else
			SQLSecurity = "SELECT -1 TYPE_ID, -1 ACCESS_ID FROM DUAL"
		End If
		
		SQLstmt = "MERGE INTO SYS_CODE_DETAIL ORI " & _
		"USING " & _
		"( " & _
			"SELECT " & _
			"PAGE.TYPE_ID SYS_CDD_SYS_CDM_ID, " & _
			"? SYS_CDD_NAME, " & _
			"PAGE.ACCESS_ID SYS_CDD_VALUE, " & _
			"NVL2(SECURITY.TYPE_ID,1,0) SECURITY_FLAG " & _
			"FROM " & _
			"( " & _
				"SELECT " & _
				"50 TYPE_ID, " & _
				"ACT.SYS_CDD_ID ACCESS_ID " & _
				"FROM SYS_CODE_DETAIL ACT " & _
				"LEFT JOIN SYS_CODE_DETAIL ARC " & _
				"ON ARC.SYS_CDD_SYS_CDM_ID = 497 " & _
				"AND ACT.SYS_CDD_VALUE = ARC.SYS_CDD_VALUE " & _
				"WHERE ACT.SYS_CDD_SYS_CDM_ID IN (45,46,47) " & _
				"AND ARC.SYS_CDD_ID IS NULL " & _
				"UNION ALL " & _
				"SELECT " & _
				"132, " & _
				"OPS_RPM_ID " & _
				"FROM OPS_REPORT_MASTER " & _
				"WHERE OPS_RPM_STAND_ALONE = 'Y' " & _
				"AND OPS_RPM_STATUS = 'ACT' " & _
			") PAGE " & _
			"LEFT JOIN " & _
			"( " & _
				SQLSecurity & _
			") SECURITY " & _
			"ON PAGE.TYPE_ID = SECURITY.TYPE_ID " & _
			"AND PAGE.ACCESS_ID = SECURITY.ACCESS_ID " & _
		") CUR " & _
		"ON " & _
		"( " & _
			"ORI.SYS_CDD_SYS_CDM_ID = CUR.SYS_CDD_SYS_CDM_ID " & _
			"AND ORI.SYS_CDD_NAME = TO_CHAR(CUR.SYS_CDD_NAME) " & _
			"AND ORI.SYS_CDD_VALUE = TO_CHAR(CUR.SYS_CDD_VALUE) " & _
		") " & _
		"WHEN NOT MATCHED THEN INSERT " & _
			"(ORI.SYS_CDD_SYS_CDM_ID, ORI.SYS_CDD_NAME, ORI.SYS_CDD_VALUE) VALUES (CUR.SYS_CDD_SYS_CDM_ID, CUR.SYS_CDD_NAME, CUR.SYS_CDD_VALUE) " & _
			"WHERE CUR.SECURITY_FLAG = 1 " & _
		"WHEN MATCHED THEN UPDATE " & _
			"SET ORI.SYS_CDD_ID = ORI.SYS_CDD_ID " & _
			"DELETE WHERE SECURITY_FLAG = 0"
		cmd.CommandText = SQLstmt
		cmd.Parameters(0).value = SECURITY_AGENT
		Set RSM = cmd.Execute
		Set RSM = Nothing	
	End If
%>
<!--#include file="pulsefunctions.asp"-->
<% Set cmd = Nothing %>
<% Conn.Close %>
<% Set Conn = Nothing %>