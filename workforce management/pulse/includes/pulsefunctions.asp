<%
	Function AgentName(USR_ID)
		Set Functioncmd = Server.CreateObject("ADODB.Command")
		Set Functioncmd.ActiveConnection = Conn
		FunctionSQLstmt = "SELECT OPS_USR_NAME NAME " & _
		"FROM OPS_USER " & _
		"WHERE OPS_USR_ID = ?"
		Functioncmd.CommandText = FunctionSQLstmt
		Functioncmd.Parameters(0).value = USR_ID
		Set RSFUNC = Functioncmd.Execute
		If Not RSFUNC.EOF Then
			AgentName = RSFUNC("NAME")
		Else
			AgentName = "NA"
		End If
		Set RSFUNC = Nothing
		Set Functioncmd = Nothing
	End Function
	
	Function AgentDepartment(USR_ID)
		Set Functioncmd = Server.CreateObject("ADODB.Command")
		Set Functioncmd.ActiveConnection = Conn
		FunctionSQLstmt = "SELECT DECODE(OPS_USD_TEAM,'SPT','SPT','SRV','SPT','OSR','SPT',OPS_USD_TYPE) DEPARTMENT " & _
		"FROM OPS_USER_DETAIL " & _
		"WHERE TO_DATE(?,'MM/DD/YYYY') BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE " & _
		"AND OPS_USD_OPS_USR_ID = ?"
		Functioncmd.CommandText = FunctionSQLstmt
		Functioncmd.Parameters(0).value = PARAMETER_DATE
		Functioncmd.Parameters(1).value = USR_ID
		Set RSFUNC = Functioncmd.Execute
		If Not RSFUNC.EOF Then
			AgentDepartment = RSFUNC("DEPARTMENT")
		Else
			AgentDepartment = "NA"
		End If
		Set RSFUNC = Nothing
		Set Functioncmd = Nothing
	End Function

	Function AgentPhone(USR_ID)
		Set Functioncmd = Server.CreateObject("ADODB.Command")
		Set Functioncmd.ActiveConnection = Conn
		FunctionSQLstmt = "SELECT OPS_USD_PHN_ID PHONE " & _
		"FROM OPS_USER_DETAIL " & _
		"WHERE TO_DATE(?,'MM/DD/YYYY') BETWEEN OPS_USD_EFF_DATE AND OPS_USD_DIS_DATE " & _
		"AND OPS_USD_OPS_USR_ID = ?"
		Functioncmd.CommandText = FunctionSQLstmt
		Functioncmd.Parameters(0).value = PARAMETER_DATE
		Functioncmd.Parameters(1).value = USR_ID
		Set RSFUNC = Functioncmd.Execute
		If Not RSFUNC.EOF Then
			AgentPhone = RSFUNC("PHONE")
		Else
			AgentPhone = "0"
		End If
		Set RSFUNC = Nothing
		Set Functioncmd = Nothing
	End Function
	
	Function ErrorDescription(ERROR_CODE)
		Select Case ERROR_CODE
			Case "LATE"
				ErrorDescription = "Late"
			Case "FLEX"
				ErrorDescription = "Flex"
			Case "SHFT"
				ErrorDescription = "No Shift"
			Case "STRT"
				ErrorDescription = "Start"
			Case "END"
				ErrorDescription = "End"
			Case "NFLX"
				ErrorDescription = "NEWH Flex"
			Case "OFLX"
				ErrorDescription = "Opener Flex"
			Case "GAP"
				ErrorDescription = "Gap"
			Case "LMIN"
				ErrorDescription = "Lunch Minutes"
			Case "LWAV"
				ErrorDescription = "Lunch Waiver"				
			Case "OLAP"
				ErrorDescription = "Overlap"
			Case "IVLD"
				ErrorDescription = "Invalid"
			Case "BWHC"
				ErrorDescription = "Weekly Hours"	
			Case "ABS"
				ErrorDescription = "Absence"
			Case "FWP"
				ErrorDescription = "Floor Walker Program"
			Case "OTH"
				ErrorDescription = "Other"
			Case "OUT"
				ErrorDescription = "Outage"
			Case "WFH"
				ErrorDescription = "WFH Technical Issue"
			Case "NA"
				ErrorDescription = "N/A"
			Case Else
				ErrorDescription = ERROR_CODE
		End Select
	End Function
	
	Function DepartmentName(DEPARTMENT_NAME)
		Select Case DEPARTMENT_NAME
			Case "ACC"
				DepartmentName = "Accounting"
			Case "CRT"
				DepartmentName = "Customer Relations"
			Case "DOC"
				DepartmentName = "Documents"
			Case "GRP"
				DepartmentName = "Group"
			Case "NEW"
				DepartmentName = "New Hires"
			Case "OPS"
				DepartmentName = "Operations"
			Case "OSS"
				DepartmentName = "Operations Support"
			Case "POP"
				DepartmentName = "Product Operations"
			Case "RES"
				DepartmentName = "Reservations"
			Case "SPT"
				DepartmentName = "Support Desk"				
			Case "ALL"
				DepartmentName = "Complete"	
			Case "SLS"
				DepartmentName = "Sales"
			Case "SRV"
				DepartmentName = "Star Elite"
			Case "OSR"
				DepartmentName = "Overnight Support"
			Case "AIR"
				DepartmentName = "Air Support"
			Case "PRD"
				DepartmentName = "Product Support"
			Case "SKD"
				DepartmentName = "Schedule Change"
			Case Else
				DepartmentName = DEPARTMENT_NAME
		End Select
	End Function
	
	Function ShiftColor(SCHEDULE_CLASS)
		Select Case SCHEDULE_CLASS
			Case "PHONE","BASE","PICK","ADDT","EXTD","HOLW"
				ShiftColor = "#1b94d1"
			Case "TRAIN","MEET","PRES","PROJ","TRAN","FAMP","WFHU","MLTU","OTRG","NEWH"
				ShiftColor = "#b2b2b2"
			Case "SRED","SRPT","SRUN"
				ShiftColor = "#e5b219"
			Case "LUNCH","LNCH","LNFL"
				ShiftColor = "#aeedef"
			Case "VACA","CDPT","CDUN","APPT","APUN","TOUN","LLUN","LLPT","OTUN","OTPT","OTPP","SLIP","SKUN","SKPT","SKPP","RESH","RCHG","ROUT","WXPT","WXUN","FMPT","FMPP","FMUN","FMHL","MLUN","MLPT","MLPP","TRPT","TRUN","FLPT","FLUN","BRVT","JURY","HOLU","HOLR"
				ShiftColor = "#ab29b2"
			Case Else
				ShiftColor = "#fff"
		End Select
	End Function
	
	Function PulseURIDecoder(FORM_STRING)
		PulseURIDecoder = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(FORM_STRING,"%21","!"),"%23","#"),"%24","$"),"%26","&"),"%27","''"),"%28","("),"%29",")"),"%2A","*"),"%2B","+"),"%2C",","),"%2F","/"),"%3A",":"),"%3B",";"),"%3D","="),"%3F","?"),"%40","@"),"%5B","["),"%5D","]")
	End Function

	Function PulsePhoneStatus(CMS_PHONE_STATE)
		If Instr(CMS_PHONE_STATE,"AUX") > 0 Then
			PulsePhoneStatus = ""
		ElseIf CMS_PHONE_STATE = "ACDIN" Then
			PulsePhoneStatus = "ACD"
		ElseIf Instr(CMS_PHONE_STATE,"ACW") > 0 Then
			PulsePhoneStatus = "ACW"
		ElseIf CMS_PHONE_STATE = "AVAIL" Then
			PulsePhoneStatus = "AVAIL"
		ElseIf CMS_PHONE_STATE = "BREAK" Then
			PulsePhoneStatus = "BREAK"		
		ElseIf CMS_PHONE_STATE = "CALLBACKS" Then
			PulsePhoneStatus = "CALLBACKS"	
		ElseIf CMS_PHONE_STATE = "CHAT QA" Then
			PulsePhoneStatus = "CHAT"	
		ElseIf CMS_PHONE_STATE = "DEFAULT" Then
			PulsePhoneStatus = "DEFAULT"
		ElseIf CMS_PHONE_STATE = "LUNCH" Then
			PulsePhoneStatus = "LUNCH"	
		ElseIf CMS_PHONE_STATE = "MEETING" Then
			PulsePhoneStatus = "MEETING"
		ElseIf CMS_PHONE_STATE = "OTHER" Then
			PulsePhoneStatus = "OTHER"
		ElseIf CMS_PHONE_STATE = "OUTBOUND" Then
			PulsePhoneStatus = "OUTBOUND"
		ElseIf CMS_PHONE_STATE = "PRESENTATION" Then
			PulsePhoneStatus = "PRESENTATION"
		ElseIf Instr(CMS_PHONE_STATE,"PROJECT") > 0 Then
			PulsePhoneStatus = "PROJECT"
		ElseIf CMS_PHONE_STATE = "RING" Then
			PulsePhoneStatus = "RING"	
		ElseIf CMS_PHONE_STATE = "TECH HELP" Then
			PulsePhoneStatus = "TECH HELP"
		ElseIf CMS_PHONE_STATE = "TRAINING" Then
			PulsePhoneStatus = "TRAINING"	
		Else
			PulsePhoneStatus = ""
		End If
	End Function
	
	Function ControlDecoder(PARAMETER_CONTROL)
		Select Case PARAMETER_CONTROL
			Case "DROPCLOSED"
				ControlDecoder = "DROP_CLOSED"
			Case "DROPOPEN"
				ControlDecoder = "DROP_OPEN"
			Case "SRUNUNAVAILABLE"
				ControlDecoder = "SRUN_UNAVAILABLE"
			Case "ADDOPEN"
				ControlDecoder = "ADD_OPEN"
			Case "ADDCLOSED"
				ControlDecoder = "ADD_CLOSED"
			Case "REDUCELIMIT"
				ControlDecoder = "REDUCE_LIMIT"
			Case "ADDLIMIT"
				ControlDecoder = "ADD_LIMIT"
			Case "SELFTRADEOFF"
				ControlDecoder = "SELFTRADE_OFF"
			Case "PICKCLOSED"
				ControlDecoder = "PICK_CLOSED"
			Case "NEWHIRE"
				ControlDecoder = "NEW_HIRE"
			Case Else
				ControlDecoder = PARAMETER_CONTROL
		End Select
	End Function
	
	Function ControlFields(PARAMETER_CONTROL)
		Select Case PARAMETER_CONTROL
			Case "DROP_CLOSED"
				ControlFields = "DOTW,INTERVAL"
			Case "DROP_OPEN"
				ControlFields = "DOTW,INTERVAL"
			Case "SRUN_UNAVAILABLE"
				ControlFields = "DOTW,INTERVAL"
			Case "ADD_OPEN"
				ControlFields = "DOTW,INTERVAL"
			Case "ADD_CLOSED"
				ControlFields = "DOTW,INTERVAL"
			Case "REDUCE_LIMIT"
				ControlFields = "SCHEDULED_HOURS,EVAL_SCORE,REDUCE_HOURS"
			Case "ADD_LIMIT"
				ControlFields = "ADD_HOURS"
			Case "SELFTRADE_OFF"
				ControlFields = "DOTW,INTERVAL"
			Case "PICK_CLOSED"
				ControlFields = "DOTW,INTERVAL"
			Case "NEW_HIRE"
				ControlFields = "NEW_DAYS"
			Case Else
				ControlFields = ""
		End Select
	End Function
	
	Function ControlTitle(PARAMETER_CONTROL)
		Select Case PARAMETER_CONTROL
			Case "DROP_CLOSED"
				ControlTitle = "Drop Closed (disregards staffing)"
			Case "DROP_OPEN"
				ControlTitle = "Drop Open (disregards staffing)"
			Case "SRUN_UNAVAILABLE"
				ControlTitle = "SRUN Unavailable"
			Case "ADD_OPEN"
				ControlTitle = "Add Open (disregards staffing)"
			Case "ADD_CLOSED"
				ControlTitle = "Add Closed (disregards staffing)"
			Case "REDUCE_LIMIT"
				ControlTitle = "Unpaid Limits"
			Case "ADD_LIMIT"
				ControlTitle = "OT Limits"
			Case "SELFTRADE_OFF"
				ControlTitle = "Self-Trade Closed (disregards staffing)"
			Case "PICK_CLOSED"
				ControlTitle = "Pick Closed (disregards staffing)"
			Case "NEW_HIRE"
				ControlTitle = "New Hire Wait Period"
			Case Else
				ControlTitle = PARAMETER_CONTROL
		End Select
	End Function
	
	Function ControlDateType(PARAMETER_CONTROL, DATE_END)
		If PARAMETER_CONTROL = "REDUCE_LIMIT" or PARAMETER_CONTROL = "ADD_LIMIT" Then
			If DATE_END = "START" Then
				ControlDateType = "sunday"
			Else
				ControlDateType = "saturday"
			End If
		Else
			ControlDateType = "daily"
		End If
	End Function
	
	Function TradeEmail(TRM_ID,TRD_ID,TRADE_STATUS)
		Set Tradecmd = Server.CreateObject("ADODB.Command")
		Set Tradecmd.ActiveConnection = Conn
		
		TradeSQLstmt = "SELECT " & _
		"OPS_TRM_OPS_USR_ID REQ_USR_ID, " & _
		"RAGT.OPS_USR_NAME REQUEST_AGENT, " & _
		"TO_DATE(OPS_TRM_START) REQUEST_DATE, " & _
		"OPS_TRC_OPS_USR_ID ACC_USR_ID, " & _
		"AAGT.OPS_USR_NAME ACCEPT_AGENT, " & _
		"TO_DATE(OPS_TRC_START) ACCEPT_DATE, " & _
		"DECODE(TO_DATE(OPS_TRM_START),TO_DATE(OPS_TRC_START),1,0) SAME_DAY_FLAG, " & _
		"RAGT.OPS_USR_EMAIL_ADDR || ';' || AAGT.OPS_USR_EMAIL_ADDR TO_EMAIL, " & _
		"'MinotOperationsDesk@deltavacations.com;' || RSUP.OPS_USR_EMAIL_ADDR || NULLIF(';' || ASUP.OPS_USR_EMAIL_ADDR,';' || RSUP.OPS_USR_EMAIL_ADDR) CC_EMAIL " & _
		"FROM OPS_TRADE_MASTER " & _
		"JOIN OPS_TRADE_DETAIL " & _
		"ON OPS_TRM_ID = OPS_TRD_OPS_TRM_ID " & _
		"JOIN OPS_TRADE_COMPLETE " & _
		"ON OPS_TRD_ID = OPS_TRC_OPS_TRD_ID " & _
		"JOIN OPS_USER_DETAIL RAGTD " & _
		"ON OPS_TRM_OPS_USR_ID = RAGTD.OPS_USD_OPS_USR_ID " & _
		"AND TO_DATE(OPS_TRM_START) BETWEEN RAGTD.OPS_USD_EFF_DATE AND RAGTD.OPS_USD_DIS_DATE " & _
		"JOIN OPS_USER_DETAIL AAGTD " & _
		"ON OPS_TRC_OPS_USR_ID = AAGTD.OPS_USD_OPS_USR_ID " & _
		"AND TO_DATE(OPS_TRC_START) BETWEEN AAGTD.OPS_USD_EFF_DATE AND AAGTD.OPS_USD_DIS_DATE " & _
		"JOIN OPS_USER RAGT " & _
		"ON OPS_TRM_OPS_USR_ID = RAGT.OPS_USR_ID " & _
		"JOIN OPS_USER AAGT " & _
		"ON OPS_TRC_OPS_USR_ID = AAGT.OPS_USR_ID " & _
		"JOIN OPS_USER RSUP " & _
		"ON RAGTD.OPS_USD_SUPERVISOR = RSUP.OPS_USR_ID " & _
		"JOIN OPS_USER ASUP " & _
		"ON AAGTD.OPS_USD_SUPERVISOR = ASUP.OPS_USR_ID " & _
		"WHERE OPS_TRM_ID = ? " & _
		"AND OPS_TRD_ID = ?"
		Tradecmd.CommandText = TradeSQLstmt
		Tradecmd.Parameters(0).value = TRM_ID
		Tradecmd.Parameters(1).value = TRD_ID
		Set RSTRADELIST = Tradecmd.Execute
		If Not RSTRADELIST.EOF Then
			TRADE_HTML = "<html><head></head><body>" & _
			"<p style=""font-family:Calibri;margin-bottom:10px;"">The following shift trade has been " & Replace(Replace(TRADE_STATUS,"REQ","denied"),"COM","approved") & ".</p>"
			If TRADE_STATUS = "REQ" Then 
				TRADE_HTML = TRADE_HTML & "<p style=""font-family:Calibri;margin-bottom:10px;""><span style=""font-weight:900;"">Reason for denial: </span>" & Request.Form("TRADETEXT_" & TRM_ID) & "</p>"
			End If
			If TRADE_STATUS = "COM" Then 
				TRADE_HTML = TRADE_HTML & "<table style=""width:800px;text-align:center;border-collapse:collapse;"">"
			Else
				TRADE_HTML = TRADE_HTML & "<table style=""width:600px;text-align:center;border-collapse:collapse;"">"
			End If
				TRADE_HTML = TRADE_HTML & "<tr>" & _
					"<td style=""font-family:Calibri;background-color:#395a93;color:#fff;border:1px solid #dee2e6;padding:5px;"">Date</td>" & _
					"<td colspan='2' style=""font-family:Calibri;background-color:#395a93;color:#fff;border:1px solid #dee2e6;padding:5px;"">Pre-Trade</td>"
					If TRADE_STATUS = "COM" Then
						TRADE_HTML = TRADE_HTML & "<td colspan='2' style=""font-family:Calibri;background-color:#395a93;color:#fff;border:1px solid #dee2e6;padding:5px;"">Post-Trade</td>"
					End If
				TRADE_HTML = TRADE_HTML & "</tr>" & _
				"<tr>" & _
					"<td style=""font-family:Calibri;border:1px solid #dee2e6;padding:5px;"">&nbsp;</td>" & _
					"<td style=""font-family:Calibri;border:1px solid #dee2e6;padding:5px;"">" & RSTRADELIST("REQUEST_AGENT") & "</td>" & _
					"<td style=""font-family:Calibri;border:1px solid #dee2e6;padding:5px;"">" & RSTRADELIST("ACCEPT_AGENT") & "</td>"
					If TRADE_STATUS = "COM" Then
						TRADE_HTML = TRADE_HTML & "<td style=""font-family:Calibri;border:1px solid #dee2e6;padding:5px;"">" & RSTRADELIST("REQUEST_AGENT") & "</td>" & _
						"<td style=""font-family:Calibri;border:1px solid #dee2e6;padding:5px;"">" & RSTRADELIST("ACCEPT_AGENT") & "</td>"
					End If
				TRADE_HTML = TRADE_HTML & "</tr>" & _
				"<tr>" & _
					"<td style=""font-family:Calibri;border:1px solid #dee2e6;padding:5px;"">" & RSTRADELIST("REQUEST_DATE") & "</td>" & _
					"<td style=""font-family:Calibri;border:1px solid #dee2e6;padding:5px;"">"
						Tradecmd.CommandText = PulsePreTradeSQL
						Tradecmd.Parameters(0).value = CDate(RSTRADELIST("REQUEST_DATE"))
						Tradecmd.Parameters(1).value = RSTRADELIST("REQ_USR_ID")
						Set RSSHIFT = Tradecmd.Execute
						If Not RSSHIFT.EOF Then
							TRADE_HTML = TRADE_HTML & "<table style=""margin:auto;border-collapse:collapse;"">"
							Do While Not RSSHIFT.EOF
								TRADE_HTML = TRADE_HTML & "<tr>" & _
									"<td style=""font-family:Calibri;background-color:" & ShiftColor(RSSHIFT("SCHEDULE_CLASS")) & ";"">" & RSSHIFT("OPS_SCI_TYPE") & "</td>" & _
									"<td style=""font-family:Calibri;background-color:" & ShiftColor(RSSHIFT("SCHEDULE_CLASS")) & ";"">" & RSSHIFT("OPS_SCI_STATUS") & "</td>" & _
									"<td style=""font-family:Calibri;background-color:" & ShiftColor(RSSHIFT("SCHEDULE_CLASS")) & ";"">" & RSSHIFT("SCI_START") & "</td>" & _
									"<td style=""font-family:Calibri;background-color:" & ShiftColor(RSSHIFT("SCHEDULE_CLASS")) & ";"">" & RSSHIFT("SCI_END") & "</td>" & _
								"</tr>"
								RSSHIFT.MoveNext
							Loop
							TRADE_HTML = TRADE_HTML & "</table>"
						End If
						Set RSSHIFT = Nothing
					TRADE_HTML = TRADE_HTML & "</td>" & _
					"<td style=""font-family:Calibri;border:1px solid #dee2e6;padding:5px;"">"
						Tradecmd.CommandText = PulsePreTradeSQL
						Tradecmd.Parameters(0).value = CDate(RSTRADELIST("REQUEST_DATE"))
						Tradecmd.Parameters(1).value = RSTRADELIST("ACC_USR_ID")
						Set RSSHIFT = Tradecmd.Execute
						If Not RSSHIFT.EOF Then
							TRADE_HTML = TRADE_HTML & "<table style=""margin:auto;border-collapse:collapse;"">"
							Do While Not RSSHIFT.EOF
								TRADE_HTML = TRADE_HTML & "<tr>" & _
									"<td style=""font-family:Calibri;background-color:" & ShiftColor(RSSHIFT("SCHEDULE_CLASS")) & ";"">" & RSSHIFT("OPS_SCI_TYPE") & "</td>" & _
									"<td style=""font-family:Calibri;background-color:" & ShiftColor(RSSHIFT("SCHEDULE_CLASS")) & ";"">" & RSSHIFT("OPS_SCI_STATUS") & "</td>" & _
									"<td style=""font-family:Calibri;background-color:" & ShiftColor(RSSHIFT("SCHEDULE_CLASS")) & ";"">" & RSSHIFT("SCI_START") & "</td>" & _
									"<td style=""font-family:Calibri;background-color:" & ShiftColor(RSSHIFT("SCHEDULE_CLASS")) & ";"">" & RSSHIFT("SCI_END") & "</td>" & _
								"</tr>"
								RSSHIFT.MoveNext
							Loop
							TRADE_HTML = TRADE_HTML & "</table>"
						End If
						Set RSSHIFT = Nothing
					TRADE_HTML = TRADE_HTML & "</td>"
					If TRADE_STATUS = "COM" Then
						TRADE_HTML = TRADE_HTML & "<td style=""font-family:Calibri;border:1px solid #dee2e6;padding:5px;"">" & _
							"<table style=""margin:auto;border-collapse:collapse;"">"
							For Each FIELD in Request.Form
								If Left(FIELD,7) = "SCIUSER" and CStr(Request.Form(FIELD)) = CStr(RSTRADELIST("REQ_USR_ID")) Then
									ID_ARRAY = Split(FIELD,"_")
									SLIDER_ID = ID_ARRAY(1)
									If CDate(Request.Form("SCIDATE_" & SLIDER_ID)) = CDate(RSTRADELIST("REQUEST_DATE")) and Request.Form("SCISTATUS_" & SLIDER_ID) = "APP" Then
										TRADE_HTML = TRADE_HTML & "<tr>" & _
											"<td style=""font-family:Calibri;background-color:" & ShiftColor(Request.Form("SCITYPE_" & SLIDER_ID)) & ";"">" & Request.Form("SCITYPE_" & SLIDER_ID) & "</td>" & _
											"<td style=""font-family:Calibri;background-color:" & ShiftColor(Request.Form("SCITYPE_" & SLIDER_ID)) & ";"">" & Request.Form("SCISTATUS_" & SLIDER_ID) & "</td>" & _
											"<td style=""font-family:Calibri;background-color:" & ShiftColor(Request.Form("SCITYPE_" & SLIDER_ID)) & ";"">" & Request.Form("SCISTART_" & SLIDER_ID) & "</td>" & _
											"<td style=""font-family:Calibri;background-color:" & ShiftColor(Request.Form("SCITYPE_" & SLIDER_ID)) & ";"">" & Replace(Request.Form("SCIEND_" & SLIDER_ID),"24:00","00:00") & "</td>" & _
										"</tr>"
									End If
								End If
							Next
							TRADE_HTML = TRADE_HTML & "</table>" & _
						"</td>" & _
						"<td style=""font-family:Calibri;border:1px solid #dee2e6;padding:5px;"">" & _
							"<table style=""margin:auto;border-collapse:collapse;"">" 
							For Each FIELD in Request.Form
								If Left(FIELD,7) = "SCIUSER" and CStr(Request.Form(FIELD)) = CStr(RSTRADELIST("ACC_USR_ID")) Then
									ID_ARRAY = Split(FIELD,"_")
									SLIDER_ID = ID_ARRAY(1)
									If CDate(Request.Form("SCIDATE_" & SLIDER_ID)) = CDate(RSTRADELIST("REQUEST_DATE")) and Request.Form("SCISTATUS_" & SLIDER_ID) = "APP" Then
										TRADE_HTML = TRADE_HTML & "<tr>" & _
											"<td style=""font-family:Calibri;background-color:" & ShiftColor(Request.Form("SCITYPE_" & SLIDER_ID)) & ";"">" & Request.Form("SCITYPE_" & SLIDER_ID) & "</td>" & _
											"<td style=""font-family:Calibri;background-color:" & ShiftColor(Request.Form("SCITYPE_" & SLIDER_ID)) & ";"">" & Request.Form("SCISTATUS_" & SLIDER_ID) & "</td>" & _
											"<td style=""font-family:Calibri;background-color:" & ShiftColor(Request.Form("SCITYPE_" & SLIDER_ID)) & ";"">" & Request.Form("SCISTART_" & SLIDER_ID) & "</td>" & _
											"<td style=""font-family:Calibri;background-color:" & ShiftColor(Request.Form("SCITYPE_" & SLIDER_ID)) & ";"">" & Replace(Request.Form("SCIEND_" & SLIDER_ID),"24:00","00:00") & "</td>" & _
										"</tr>"
									End If
								End If
							Next
							TRADE_HTML = TRADE_HTML & "</table>" & _
						"</td>"
					End If
				TRADE_HTML = TRADE_HTML & "</tr>"
				If RSTRADELIST("SAME_DAY_FLAG") <> "1" Then
					TRADE_HTML = TRADE_HTML & "<tr>" & _
					"<td style=""font-family:Calibri;border:1px solid #dee2e6;padding:5px;"">" & RSTRADELIST("ACCEPT_DATE") & "</td>" & _
					"<td style=""font-family:Calibri;border:1px solid #dee2e6;padding:5px;"">"
						Tradecmd.CommandText = PulsePreTradeSQL
						Tradecmd.Parameters(0).value = CDate(RSTRADELIST("ACCEPT_DATE"))
						Tradecmd.Parameters(1).value = RSTRADELIST("REQ_USR_ID")
						Set RSSHIFT = Tradecmd.Execute
						If Not RSSHIFT.EOF Then
							TRADE_HTML = TRADE_HTML & "<table style=""margin:auto;border-collapse:collapse;"">"
							Do While Not RSSHIFT.EOF
								TRADE_HTML = TRADE_HTML & "<tr>" & _
									"<td style=""font-family:Calibri;background-color:" & ShiftColor(RSSHIFT("SCHEDULE_CLASS")) & ";"">" & RSSHIFT("OPS_SCI_TYPE") & "</td>" & _
									"<td style=""font-family:Calibri;background-color:" & ShiftColor(RSSHIFT("SCHEDULE_CLASS")) & ";"">" & RSSHIFT("OPS_SCI_STATUS") & "</td>" & _
									"<td style=""font-family:Calibri;background-color:" & ShiftColor(RSSHIFT("SCHEDULE_CLASS")) & ";"">" & RSSHIFT("SCI_START") & "</td>" & _
									"<td style=""font-family:Calibri;background-color:" & ShiftColor(RSSHIFT("SCHEDULE_CLASS")) & ";"">" & RSSHIFT("SCI_END") & "</td>" & _
								"</tr>"
								RSSHIFT.MoveNext
							Loop
							TRADE_HTML = TRADE_HTML & "</table>"
						End If
						Set RSSHIFT = Nothing
					TRADE_HTML = TRADE_HTML & "</td>" & _
					"<td style=""font-family:Calibri;border:1px solid #dee2e6;padding:5px;"">"
						Tradecmd.CommandText = PulsePreTradeSQL
						Tradecmd.Parameters(0).value = CDate(RSTRADELIST("ACCEPT_DATE"))
						Tradecmd.Parameters(1).value = RSTRADELIST("ACC_USR_ID")
						Set RSSHIFT = Tradecmd.Execute
						If Not RSSHIFT.EOF Then
							TRADE_HTML = TRADE_HTML & "<table style=""margin:auto;border-collapse:collapse;"">"
							Do While Not RSSHIFT.EOF
								TRADE_HTML = TRADE_HTML & "<tr>" & _
									"<td style=""font-family:Calibri;background-color:" & ShiftColor(RSSHIFT("SCHEDULE_CLASS")) & ";"">" & RSSHIFT("OPS_SCI_TYPE") & "</td>" & _
									"<td style=""font-family:Calibri;background-color:" & ShiftColor(RSSHIFT("SCHEDULE_CLASS")) & ";"">" & RSSHIFT("OPS_SCI_STATUS") & "</td>" & _
									"<td style=""font-family:Calibri;background-color:" & ShiftColor(RSSHIFT("SCHEDULE_CLASS")) & ";"">" & RSSHIFT("SCI_START") & "</td>" & _
									"<td style=""font-family:Calibri;background-color:" & ShiftColor(RSSHIFT("SCHEDULE_CLASS")) & ";"">" & RSSHIFT("SCI_END") & "</td>" & _
								"</tr>"
								RSSHIFT.MoveNext
							Loop
							TRADE_HTML = TRADE_HTML & "</table>"
						End If
						Set RSSHIFT = Nothing
					TRADE_HTML = TRADE_HTML & "</td>"
					If TRADE_STATUS = "COM" Then
						TRADE_HTML = TRADE_HTML & "<td style=""font-family:Calibri;border:1px solid #dee2e6;padding:5px;"">" & _
							"<table style=""margin:auto;border-collapse:collapse;"">"
							For Each FIELD in Request.Form
								If Left(FIELD,7) = "SCIUSER" and CStr(Request.Form(FIELD)) = CStr(RSTRADELIST("REQ_USR_ID")) Then
									ID_ARRAY = Split(FIELD,"_")
									SLIDER_ID = ID_ARRAY(1)
									If CDate(Request.Form("SCIDATE_" & SLIDER_ID)) = CDate(RSTRADELIST("ACCEPT_DATE")) and Request.Form("SCISTATUS_" & SLIDER_ID) = "APP" Then
										TRADE_HTML = TRADE_HTML & "<tr>" & _
											"<td style=""font-family:Calibri;background-color:" & ShiftColor(Request.Form("SCITYPE_" & SLIDER_ID)) & ";"">" & Request.Form("SCITYPE_" & SLIDER_ID) & "</td>" & _
											"<td style=""font-family:Calibri;background-color:" & ShiftColor(Request.Form("SCITYPE_" & SLIDER_ID)) & ";"">" & Request.Form("SCISTATUS_" & SLIDER_ID) & "</td>" & _
											"<td style=""font-family:Calibri;background-color:" & ShiftColor(Request.Form("SCITYPE_" & SLIDER_ID)) & ";"">" & Request.Form("SCISTART_" & SLIDER_ID) & "</td>" & _
											"<td style=""font-family:Calibri;background-color:" & ShiftColor(Request.Form("SCITYPE_" & SLIDER_ID)) & ";"">" & Replace(Request.Form("SCIEND_" & SLIDER_ID),"24:00","00:00") & "</td>" & _
										"</tr>"
									End If
								End If
							Next
							TRADE_HTML = TRADE_HTML & "</table>" & _
						"</td>" & _
						"<td style=""font-family:Calibri;border:1px solid #dee2e6;padding:5px;"">" & _
							"<table style=""margin:auto;border-collapse:collapse;"">"
							For Each FIELD in Request.Form
								If Left(FIELD,7) = "SCIUSER" and CStr(Request.Form(FIELD)) = CStr(RSTRADELIST("ACC_USR_ID")) Then
									ID_ARRAY = Split(FIELD,"_")
									SLIDER_ID = ID_ARRAY(1)
									If CDate(Request.Form("SCIDATE_" & SLIDER_ID)) = CDate(RSTRADELIST("ACCEPT_DATE")) and Request.Form("SCISTATUS_" & SLIDER_ID) = "APP" Then
										TRADE_HTML = TRADE_HTML & "<tr>" & _
											"<td style=""font-family:Calibri;background-color:" & ShiftColor(Request.Form("SCITYPE_" & SLIDER_ID)) & ";"">" & Request.Form("SCITYPE_" & SLIDER_ID) & "</td>" & _
											"<td style=""font-family:Calibri;background-color:" & ShiftColor(Request.Form("SCITYPE_" & SLIDER_ID)) & ";"">" & Request.Form("SCISTATUS_" & SLIDER_ID) & "</td>" & _
											"<td style=""font-family:Calibri;background-color:" & ShiftColor(Request.Form("SCITYPE_" & SLIDER_ID)) & ";"">" & Request.Form("SCISTART_" & SLIDER_ID) & "</td>" & _
											"<td style=""font-family:Calibri;background-color:" & ShiftColor(Request.Form("SCITYPE_" & SLIDER_ID)) & ";"">" & Replace(Request.Form("SCIEND_" & SLIDER_ID),"24:00","00:00") & "</td>" & _
										"</tr>"
									End If
								End If
							Next
							TRADE_HTML = TRADE_HTML & "</table>" & _
						"</td>"
					End If	
				TRADE_HTML = TRADE_HTML & "</tr>"
			End If
			TRADE_HTML = TRADE_HTML & "</table>" & _
			"<p style=""font-family:Calibri;margin-bottom:10px;"">See your Dashboard at https://opscenter.mltvacations.com for your current schedule and please contact the Ops Desk if you have any questions.</p>" & _
			"</body></html>"
			Set myMail = CreateObject("CDO.Message")
			myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 1
			myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") ="c:\Inetpub\mailroot\pickup"
			myMail.Configuration.Fields.Update
			
			myMail.Subject = "Shift Trade " & Replace(Replace(TRADE_STATUS,"REQ","Denied"),"COM","Approved") & " - " & CDate(RSTRADELIST("REQUEST_DATE"))
			myMail.HTMLBody = TRADE_HTML
			myMail.To = RSTRADELIST("TO_EMAIL")
			myMail.Cc = RSTRADELIST("CC_EMAIL")
			'myMail.Bcc = "cmoen@deltavacations.com;"
			myMail.From = "noreply@deltavacations.com;"
			myMail.Send
			Set myMail = Nothing
		End If
		Set RSTRADELIST = Nothing
		Set Tradecmd = Nothing
	End Function
%>