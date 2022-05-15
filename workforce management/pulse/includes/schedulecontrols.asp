<!--#include file="pulseheader.asp"-->
<%
	If Request.Querystring("WORKGROUP") <> "" Then
		REQUEST_WORKGROUP = Request.Querystring("WORKGROUP")
	Else
		REQUEST_WORKGROUP = "RES"
	End If
	If Request.Querystring("CONTROL") <> "" Then
		REQUEST_CONTROL = Request.Querystring("CONTROL")
	Else
		REQUEST_CONTROL = "NA"
	End If
	If Request.Querystring("DATE") <> "" Then
		PARAMETER_DATE = CDate(Request.Querystring("DATE"))
	Else
		PARAMETER_DATE = Date
	End If
	CONTROL_FIELDS = ControlFields(ControlDecoder(REQUEST_CONTROL))
	CONTROL_FIELD_ARRAY = Split(CONTROL_FIELDS,",")
	SQLstmt = "SELECT * " & _
	"FROM " & _
	"( " & _
		"SELECT " & _
		"OPS_PAR_ID, " & _
		"OPS_PAR_VALUE CONTROL_VALUE, " & _
		"OPS_PAR_EFF_DATE CONTROL_EFF_DATE, " & _
		"OPS_PAR_DIS_DATE CONTROL_DIS_DATE, " & _
		"MIN(CASE WHEN OPS_PAR_DIS_DATE >= TO_DATE(?,'MM/DD/YYYY') THEN OPS_PAR_EFF_DATE END) OVER () MIN_DATE " & _
		"FROM OPS_PARAMETER " & _
		"WHERE OPS_PAR_PARENT_TYPE = ? " & _
		"AND OPS_PAR_CODE = ? " & _
	") " & _
	"WHERE CONTROL_EFF_DATE >= MIN_DATE " & _
	"ORDER BY CONTROL_EFF_DATE, CONTROL_DIS_DATE, OPS_PAR_ID"
	cmd.CommandText = SQLstmt
	cmd.Parameters(0).value = PARAMETER_DATE
	cmd.Parameters(1).value = REQUEST_WORKGROUP
	cmd.Parameters(2).value = ControlDecoder(REQUEST_CONTROL)
	Set RSCONTROL = cmd.Execute
	If Not RSCONTROL.EOF Then
		MIN_DATE = CDate(RSCONTROL("MIN_DATE"))
	Else
		MIN_DATE = PARAMETER_DATE
	End If
%>
	<table id="CONTROLTABLE_<%=REQUEST_WORKGROUP%>_<%=REQUEST_CONTROL%>" class="center" style="width:100%;font-size:.75em;">
		<caption class="center <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>">
			<%=REQUEST_WORKGROUP & " " & ControlTitle(ControlDecoder(REQUEST_CONTROL))%>
		</caption>
		<thead>
			<tr class="subtable-td-padded-sm <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>">
				<% USE_COLSPAN = 2 %>
				<% For Each FIELD in CONTROL_FIELD_ARRAY %>
					<% If FIELD = "DOTW" Then %>
						<th class="subtable-td-padded-sm" style="width:<%=(80/(UBound(CONTROL_FIELD_ARRAY)+1)) - 15%>%;">Day</th>
						<% USE_COLSPAN = USE_COLSPAN + 1 %>
					<% Elseif FIELD = "INTERVAL" Then %>
						<th class="subtable-td-padded-sm" style="width:5%;">&nbsp;</th>
						<th class="subtable-td-padded-sm" style="width:<%=(80/(UBound(CONTROL_FIELD_ARRAY)+1)) + 5%>%;">Interval</th>
						<th class="subtable-td-padded-sm" style="width:5%;">&nbsp;</th>
						<% USE_COLSPAN = USE_COLSPAN + 3 %>
					<% Elseif FIELD = "SCHEDULED_HOURS" Then %>
						<th class="subtable-td-padded-sm" style="width:5%;">&nbsp;</th>
						<th class="subtable-td-padded-sm" style="width:<%=(80/(UBound(CONTROL_FIELD_ARRAY)+1)) - 10%>%;">Scheduled Hours</th>
						<th class="subtable-td-padded-sm" style="width:5%;">&nbsp;</th>
						<% USE_COLSPAN = USE_COLSPAN + 3 %>
					<% Elseif FIELD = "EVAL_SCORE" Then %>
						<th class="subtable-td-padded-sm" style="width:5%;">&nbsp;</th>
						<th class="subtable-td-padded-sm" style="width:<%=(80/(UBound(CONTROL_FIELD_ARRAY)+1)) - 10%>%;">Eval Score</th>
						<th class="subtable-td-padded-sm" style="width:5%;">&nbsp;</th>
						<% USE_COLSPAN = USE_COLSPAN + 3 %>
					<% Elseif FIELD = "REDUCE_HOURS" or FIELD = "ADD_HOURS" or FIELD = "NEW_DAYS" Then %>
						<th class="subtable-td-padded-sm" style="width:<%=80/(UBound(CONTROL_FIELD_ARRAY)+1)%>%;"><%=ControlTitle(ControlDecoder(REQUEST_CONTROL))%></th>
						<% USE_COLSPAN = USE_COLSPAN + 1 %>
					<% End If %>
				<% Next %>
				<th class="subtable-td-padded-sm" style="width:10%;">Eff. Date</th>
				<th class="subtable-td-padded-sm" style="width:10%;">Dis Date</th>
			</tr>
		</thead>
		<tbody>
		<% Do While Not RSCONTROL.EOF %>
			<tr>
				<input type="hidden" id="CONTROLWORKGROUP_<%=RSCONTROL("OPS_PAR_ID")%>" name="CONTROLWORKGROUP_<%=RSCONTROL("OPS_PAR_ID")%>" value="<%=REQUEST_WORKGROUP%>" />
				<input type="hidden" id="CONTROLTYPE_<%=RSCONTROL("OPS_PAR_ID")%>" name="CONTROLTYPE_<%=RSCONTROL("OPS_PAR_ID")%>" value="<%=ControlDecoder(REQUEST_CONTROL)%>" />
				<input type="hidden" id="CONTROLFIELDS_<%=RSCONTROL("OPS_PAR_ID")%>" name="CONTROLFIELDS_<%=RSCONTROL("OPS_PAR_ID")%>" value="<%=CONTROL_FIELDS%>" />
				<% CONTROL_VALUE_ARRAY = Split(RSCONTROL("CONTROL_VALUE"),";") %>
				<% i = 0 %>
				<% For Each FIELD in CONTROL_FIELD_ARRAY %>
					<% If FIELD = "DOTW" Then %>
						<% CONTROL_DOTW = CONTROL_VALUE_ARRAY(i) %>
						<% i = i + 1 %>
						<td class="subtable-td-padded-lg nowrap">
							<input type="checkbox" style="display:none;" id="CONTROLSUN_<%=RSCONTROL("OPS_PAR_ID")%>" name="CONTROLDOTW_<%=RSCONTROL("OPS_PAR_ID")%>" value="1" <% If Instr(CONTROL_DOTW,"1") Then %> checked="checked" <% End If %>/>
							<button id="CONTROLBUTTONSUN_<%=RSCONTROL("OPS_PAR_ID")%>" type="button" class="btn <% If Instr(CONTROL_DOTW,"1") Then %> new-entry-color <% End If %> <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" data-workgroup="<%=REQUEST_WORKGROUP%>" data-control="<%=REQUEST_CONTROL%>" style="background-color:#fff;font-size:.75em;padding:2px;">S</button>
							<input type="checkbox" style="display:none;" id="CONTROLMON_<%=RSCONTROL("OPS_PAR_ID")%>" name="CONTROLDOTW_<%=RSCONTROL("OPS_PAR_ID")%>" value="2" <% If Instr(CONTROL_DOTW,"2") Then %> checked="checked" <% End If %>/>
							<button id="CONTROLBUTTONMON_<%=RSCONTROL("OPS_PAR_ID")%>" type="button" class="btn <% If Instr(CONTROL_DOTW,"2") Then %> new-entry-color <% End If %> <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" data-workgroup="<%=REQUEST_WORKGROUP%>" data-control="<%=REQUEST_CONTROL%>" style="background-color:#fff;font-size:.75em;padding:2px;">M</button>
							<input type="checkbox" style="display:none;" id="CONTROLTUE_<%=RSCONTROL("OPS_PAR_ID")%>" name="CONTROLDOTW_<%=RSCONTROL("OPS_PAR_ID")%>" value="3" <% If Instr(CONTROL_DOTW,"3") Then %> checked="checked" <% End If %>/>
							<button id="CONTROLBUTTONTUE_<%=RSCONTROL("OPS_PAR_ID")%>" type="button" class="btn <% If Instr(CONTROL_DOTW,"3") Then %> new-entry-color <% End If %> <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" data-workgroup="<%=REQUEST_WORKGROUP%>" data-control="<%=REQUEST_CONTROL%>" style="background-color:#fff;font-size:.75em;padding:2px;">T</button>
							<input type="checkbox" style="display:none;" id="CONTROLWED_<%=RSCONTROL("OPS_PAR_ID")%>" name="CONTROLDOTW_<%=RSCONTROL("OPS_PAR_ID")%>" value="4" <% If Instr(CONTROL_DOTW,"4") Then %> checked="checked" <% End If %>/>
							<button id="CONTROLBUTTONWED_<%=RSCONTROL("OPS_PAR_ID")%>" type="button" class="btn <% If Instr(CONTROL_DOTW,"4") Then %> new-entry-color <% End If %> <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" data-workgroup="<%=REQUEST_WORKGROUP%>" data-control="<%=REQUEST_CONTROL%>" style="background-color:#fff;font-size:.75em;padding:2px;">W</button>
							<input type="checkbox" style="display:none;" id="CONTROLTHU_<%=RSCONTROL("OPS_PAR_ID")%>" name="CONTROLDOTW_<%=RSCONTROL("OPS_PAR_ID")%>" value="5" <% If Instr(CONTROL_DOTW,"5") Then %> checked="checked" <% End If %>/>
							<button id="CONTROLBUTTONTHU_<%=RSCONTROL("OPS_PAR_ID")%>" type="button" class="btn <% If Instr(CONTROL_DOTW,"5") Then %> new-entry-color <% End If %> <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" data-workgroup="<%=REQUEST_WORKGROUP%>" data-control="<%=REQUEST_CONTROL%>" style="background-color:#fff;font-size:.75em;padding:2px;">H</button>
							<input type="checkbox" style="display:none;" id="CONTROLFRI_<%=RSCONTROL("OPS_PAR_ID")%>" name="CONTROLDOTW_<%=RSCONTROL("OPS_PAR_ID")%>" value="6" <% If Instr(CONTROL_DOTW,"6") Then %> checked="checked" <% End If %>/>
							<button id="CONTROLBUTTONFRI_<%=RSCONTROL("OPS_PAR_ID")%>" type="button" class="btn <% If Instr(CONTROL_DOTW,"6") Then %> new-entry-color <% End If %> <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" data-workgroup="<%=REQUEST_WORKGROUP%>" data-control="<%=REQUEST_CONTROL%>" style="background-color:#fff;font-size:.75em;padding:2px;">F</button>
							<input type="checkbox" style="display:none;" id="CONTROLSAT_<%=RSCONTROL("OPS_PAR_ID")%>" name="CONTROLDOTW_<%=RSCONTROL("OPS_PAR_ID")%>" value="7" <% If Instr(CONTROL_DOTW,"7") Then %> checked="checked" <% End If %>/>
							<button id="CONTROLBUTTONSAT_<%=RSCONTROL("OPS_PAR_ID")%>" type="button" class="btn <% If Instr(CONTROL_DOTW,"7") Then %> new-entry-color <% End If %> <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" data-workgroup="<%=REQUEST_WORKGROUP%>" data-control="<%=REQUEST_CONTROL%>" style="background-color:#fff;font-size:.75em;padding:2px;">S</button>
						</td>
					<% Elseif FIELD = "INTERVAL" Then %>
						<% CONTROL_START = CONTROL_VALUE_ARRAY(i) %>
						<% CONTROL_END = CONTROL_VALUE_ARRAY(i+1) %>
						<% i = i + 2 %>
						
						<% CONTROL_TYPE = "TIME" %>
						<% CONTROL_RANGE = "true" %>
						<% CONTROL_STEP = "30" %>
						<% CONTROL_INTERVAL = "30" %>
						<% If REQUEST_WORKGROUP = "RES" or REQUEST_WORKGROUP = "SPT" Then %>
							<% CONTROL_MIN = "360" %>
							<% CONTROL_MAX = "1440" %>
						<% Elseif REQUEST_WORKGROUP = "OSR" Then %>
							<% CONTROL_MIN = "0" %>
							<% CONTROL_MAX = "1440" %>
						<% End If %>
						<td class="subtable-td-padded-lg">
							<input type="hidden" id="CONTROLSTART_INTERVAL_<%=RSCONTROL("OPS_PAR_ID")%>" name="CONTROLSTART_INTERVAL_<%=RSCONTROL("OPS_PAR_ID")%>" value="<%=CONTROL_START%>" />
							<div style="display:inline-block;white-space:nowrap;">
								<i id="CONTROLSTARTARROW_INTERVAL_LEFT_<%=RSCONTROL("OPS_PAR_ID")%>" class="fas fa-caret-left icon-style-small"></i>
								<div style="display:inline-block;" id="CONTROLSTARTDISPLAY_INTERVAL_<%=RSCONTROL("OPS_PAR_ID")%>"><%=CONTROL_START%></div>
								<i id="CONTROLSTARTARROW_INTERVAL_RIGHT_<%=RSCONTROL("OPS_PAR_ID")%>" class="fas fa-caret-right icon-style-small" style="margin-right:5px;"></i>
							</div>
						</td>
						<td class="subtable-td-padded-lg">	
							<div id="CONTROLSLIDER_INTERVAL_<%=RSCONTROL("OPS_PAR_ID")%>" style="display:inline-block;width:100%;" data-workgroup="<%=REQUEST_WORKGROUP%>" data-control="<%=REQUEST_CONTROL%>" data-slider-type="<%=CONTROL_TYPE%>" data-slider-min="<%=CONTROL_MIN%>" data-slider-max="<%=CONTROL_MAX%>" data-range="<%=CONTROL_RANGE%>" data-slider-step="<%=CONTROL_STEP%>" data-slider-interval="<%=CONTROL_INTERVAL%>"></div>	
						</td>
						<td class="subtable-td-padded-lg">
							<input type="hidden" id="CONTROLEND_INTERVAL_<%=RSCONTROL("OPS_PAR_ID")%>" name="CONTROLEND_INTERVAL_<%=RSCONTROL("OPS_PAR_ID")%>" value="<%=CONTROL_END%>" />
							<div style="display:inline-block;white-space:nowrap;">
								<i id="CONTROLENDARROW_INTERVAL_LEFT_<%=RSCONTROL("OPS_PAR_ID")%>" class="fas fa-caret-left icon-style-small" style="margin-left:5px;"></i>
								<div style="display:inline-block;" id="CONTROLENDDISPLAY_INTERVAL_<%=RSCONTROL("OPS_PAR_ID")%>"><%=CONTROL_END%></div>
								<i id="CONTROLENDARROW_INTERVAL_RIGHT_<%=RSCONTROL("OPS_PAR_ID")%>" class="fas fa-caret-right icon-style-small"></i>
							</div>
						</td>
					<% Elseif FIELD = "SCHEDULED_HOURS" Then %>
						<% CONTROL_START = CONTROL_VALUE_ARRAY(i) %>
						<% CONTROL_END = CONTROL_VALUE_ARRAY(i+1) %>
						<% i = i + 2 %>
						
						<% CONTROL_TYPE = "NUMBER" %>
						<% CONTROL_RANGE = "true" %>
						<% CONTROL_STEP = "1" %>
						<% CONTROL_INTERVAL = "1" %>
						<% CONTROL_MIN = "0" %>
						<% CONTROL_MAX = "40" %>
						<td class="subtable-td-padded-lg">
							<input type="hidden" id="CONTROLSTART_HOURS_<%=RSCONTROL("OPS_PAR_ID")%>" name="CONTROLSTART_HOURS_<%=RSCONTROL("OPS_PAR_ID")%>" value="<%=CONTROL_START%>" />
							<div style="display:inline-block;white-space:nowrap;">
								<i id="CONTROLSTARTARROW_HOURS_LEFT_<%=RSCONTROL("OPS_PAR_ID")%>" class="fas fa-caret-left icon-style-small"></i>
								<div style="display:inline-block;width:1rem;" id="CONTROLSTARTDISPLAY_HOURS_<%=RSCONTROL("OPS_PAR_ID")%>"><%=CONTROL_START%></div>
								<i id="CONTROLSTARTARROW_HOURS_RIGHT_<%=RSCONTROL("OPS_PAR_ID")%>" class="fas fa-caret-right icon-style-small" style="margin-right:5px;"></i>
							</div>
						</td>
						<td class="subtable-td-padded-lg">	
							<div id="CONTROLSLIDER_HOURS_<%=RSCONTROL("OPS_PAR_ID")%>" style="display:inline-block;width:100%;" data-workgroup="<%=REQUEST_WORKGROUP%>" data-control="<%=REQUEST_CONTROL%>" data-slider-type="<%=CONTROL_TYPE%>" data-slider-min="<%=CONTROL_MIN%>" data-slider-max="<%=CONTROL_MAX%>" data-range="<%=CONTROL_RANGE%>" data-slider-step="<%=CONTROL_STEP%>" data-slider-interval="<%=CONTROL_INTERVAL%>"></div>	
						</td>
						<td class="subtable-td-padded-lg">
							<input type="hidden" id="CONTROLEND_HOURS_<%=RSCONTROL("OPS_PAR_ID")%>" name="CONTROLEND_HOURS_<%=RSCONTROL("OPS_PAR_ID")%>" value="<%=CONTROL_END%>" />
							<div style="display:inline-block;white-space:nowrap;">
								<i id="CONTROLENDARROW_HOURS_LEFT_<%=RSCONTROL("OPS_PAR_ID")%>" class="fas fa-caret-left icon-style-small" style="margin-left:5px;"></i>
								<div style="display:inline-block;width:1rem;" id="CONTROLENDDISPLAY_HOURS_<%=RSCONTROL("OPS_PAR_ID")%>"><%=CONTROL_END%></div>
								<i id="CONTROLENDARROW_HOURS_RIGHT_<%=RSCONTROL("OPS_PAR_ID")%>" class="fas fa-caret-right icon-style-small"></i>
							</div>
						</td>
					<% Elseif FIELD = "EVAL_SCORE" Then %>
						<% CONTROL_START = CONTROL_VALUE_ARRAY(i) %>
						<% CONTROL_END = CONTROL_VALUE_ARRAY(i+1) %>
						<% i = i + 2 %>
						
						<% CONTROL_TYPE = "NUMBER" %>
						<% CONTROL_RANGE = "true" %>
						<% CONTROL_STEP = ".25" %>
						<% CONTROL_INTERVAL = ".25" %>
						<% CONTROL_MIN = "1" %>
						<% CONTROL_MAX = "5" %>
						<td class="subtable-td-padded-lg">
							<input type="hidden" id="CONTROLSTART_SCORE_<%=RSCONTROL("OPS_PAR_ID")%>" name="CONTROLSTART_SCORE_<%=RSCONTROL("OPS_PAR_ID")%>" value="<%=CONTROL_START%>" />
							<div style="display:inline-block;white-space:nowrap;">
								<i id="CONTROLSTARTARROW_SCORE_LEFT_<%=RSCONTROL("OPS_PAR_ID")%>" class="fas fa-caret-left icon-style-small"></i>
								<div style="display:inline-block;width:1.4rem;" id="CONTROLSTARTDISPLAY_SCORE_<%=RSCONTROL("OPS_PAR_ID")%>"><%=CONTROL_START%></div>
								<i id="CONTROLSTARTARROW_SCORE_RIGHT_<%=RSCONTROL("OPS_PAR_ID")%>" class="fas fa-caret-right icon-style-small" style="margin-right:5px;"></i>
							</div>
						</td>
						<td class="subtable-td-padded-lg">	
							<div id="CONTROLSLIDER_SCORE_<%=RSCONTROL("OPS_PAR_ID")%>" style="display:inline-block;width:100%;" data-workgroup="<%=REQUEST_WORKGROUP%>" data-control="<%=REQUEST_CONTROL%>" data-slider-type="<%=CONTROL_TYPE%>" data-slider-min="<%=CONTROL_MIN%>" data-slider-max="<%=CONTROL_MAX%>" data-range="<%=CONTROL_RANGE%>" data-slider-step="<%=CONTROL_STEP%>" data-slider-interval="<%=CONTROL_INTERVAL%>"></div>	
						</td>
						<td class="subtable-td-padded-lg">
							<input type="hidden" id="CONTROLEND_SCORE_<%=RSCONTROL("OPS_PAR_ID")%>" name="CONTROLEND_SCORE_<%=RSCONTROL("OPS_PAR_ID")%>" value="<%=CONTROL_END%>" />
							<div style="display:inline-block;white-space:nowrap;">
								<i id="CONTROLENDARROW_SCORE_LEFT_<%=RSCONTROL("OPS_PAR_ID")%>" class="fas fa-caret-left icon-style-small" style="margin-left:5px;"></i>
								<div style="display:inline-block;width:1.4rem;" id="CONTROLENDDISPLAY_SCORE_<%=RSCONTROL("OPS_PAR_ID")%>"><%=CONTROL_END%></div>
								<i id="CONTROLENDARROW_SCORE_RIGHT_<%=RSCONTROL("OPS_PAR_ID")%>" class="fas fa-caret-right icon-style-small"></i>
							</div>
						</td>
					<% Elseif FIELD = "REDUCE_HOURS" or FIELD = "ADD_HOURS" or FIELD = "NEW_DAYS" Then %>
						<% CONTROL_VALUE = CONTROL_VALUE_ARRAY(i) %>
						<% i = i + 1 %>
						
						<% CONTROL_TYPE = "NUMBER" %>
						<% CONTROL_RANGE = "false" %>
						<% CONTROL_STEP = "1" %>
						<% CONTROL_INTERVAL = "1" %>
						<% CONTROL_MIN = "0" %>
						<% If FIELD = "REDUCE_HOURS" Then %>
							<% CONTROL_MAX = "40" %>
						<% Elseif FIELD = "ADD_HOURS" Then %>
							<% CONTROL_MAX = "50" %>
						<% Elseif FIELD = "NEW_DAYS" Then %>
							<% CONTROL_MAX = "90" %>
						<% End If %>
						<td class="subtable-td-padded-lg">	
							<input type="hidden" id="CONTROL_VALUE_<%=RSCONTROL("OPS_PAR_ID")%>" name="CONTROL_VALUE_<%=RSCONTROL("OPS_PAR_ID")%>" value="<%=CONTROL_VALUE%>" />
							<div id="CONTROLSLIDER_VALUE_<%=RSCONTROL("OPS_PAR_ID")%>" style="display:inline-block;width:100%;" data-workgroup="<%=REQUEST_WORKGROUP%>" data-control="<%=REQUEST_CONTROL%>" data-slider-type="<%=CONTROL_TYPE%>" data-slider-min="<%=CONTROL_MIN%>" data-slider-max="<%=CONTROL_MAX%>" data-range="<%=CONTROL_RANGE%>" data-slider-step="<%=CONTROL_STEP%>" data-slider-interval="<%=CONTROL_INTERVAL%>">
								<div id="CONTROLHANDLE_VALUE_<%=RSCONTROL("OPS_PAR_ID")%>" class="ui-slider-handle custom-handle CONTROL" style="color:#fff;"></div>
							</div>	
						</td>
					<% End If %>				
				<% Next %>
				<td class="subtable-td-padded-lg">
					<input type="text" id="CONTROLEFFDATE_<%=RSCONTROL("OPS_PAR_ID")%>" name="CONTROLEFFDATE_<%=RSCONTROL("OPS_PAR_ID")%>" class="center no-border-input control-date <% If PARAMETER_DATE = Date Then %> today-color <% Else %> past-color <% End If %>" style="max-width:100px;" value="<% If ControlDateType(ControlDecoder(REQUEST_CONTROL),"START") = "sunday" Then %><%=CDate(RSCONTROL("CONTROL_EFF_DATE")) - Weekday(CDate(RSCONTROL("CONTROL_EFF_DATE"))) + 1%><% Else %><%=RSCONTROL("CONTROL_EFF_DATE")%><% End If %>" data-date-min="<% If ControlDateType(ControlDecoder(REQUEST_CONTROL),"START") = "sunday" Then %><%=MIN_DATE - Weekday(MIN_DATE) + 1%><% Else %><%=MIN_DATE%><% End If %>" data-date-max="12/31/2040" data-date-type="<%=ControlDateType(ControlDecoder(REQUEST_CONTROL),"START")%>"/>
				</td>
				<td class="subtable-td-padded-lg">
					<input type="text" id="CONTROLDISDATE_<%=RSCONTROL("OPS_PAR_ID")%>" name="CONTROLDISDATE_<%=RSCONTROL("OPS_PAR_ID")%>" class="center no-border-input control-date <% If PARAMETER_DATE = Date Then %> today-color <% Else %> past-color <% End If %>" style="max-width:100px;" value="<% If ControlDateType(ControlDecoder(REQUEST_CONTROL),"START") = "saturday" and RSCONTROL("CONTROL_DIS_DATE") <> "12/31/2040" Then %><%=CDate(RSCONTROL("CONTROL_DIS_DATE")) - Weekday(CDate(RSCONTROL("CONTROL_DIS_DATE"))) + 7%><% Else %><%=RSCONTROL("CONTROL_DIS_DATE")%><% End If %>" data-date-min="<% If ControlDateType(ControlDecoder(REQUEST_CONTROL),"START") = "saturday" Then %><%=MIN_DATE - Weekday(MIN_DATE) + 7%><% Else %><%=MIN_DATE%><% End If %>" data-date-max="12/31/2040" data-date-type="<%=ControlDateType(ControlDecoder(REQUEST_CONTROL),"END")%>"/>
				</td>
			</tr>
			<% RSCONTROL.MoveNext %>
		<% Loop %>
		</tbody>
		<% Set RSCONTROL = Nothing %>
		<tr class="new-entry-color" style="display:none;">
			<input type="hidden" id="CONTROLWORKGROUP_0" name="CONTROLWORKGROUP_0" value="<%=REQUEST_WORKGROUP%>" />
			<input type="hidden" id="CONTROLTYPE_0" name="CONTROLTYPE_0" value="<%=ControlDecoder(REQUEST_CONTROL)%>" />
			<input type="hidden" id="CONTROLFIELDS_0" name="CONTROLFIELDS_0" value="<%=CONTROL_FIELDS%>" />
			<% For Each FIELD in CONTROL_FIELD_ARRAY %>
				<% If FIELD = "DOTW" Then %>
					<td class="subtable-td-padded-lg nowrap">
						<input type="checkbox" style="display:none;" id="CONTROLSUN_0" name="CONTROLDOTW_0" value="1"/>
						<button id="CONTROLBUTTONSUN_0" type="button" class="btn <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" data-workgroup="<%=REQUEST_WORKGROUP%>" data-control="<%=REQUEST_CONTROL%>" style="background-color:#fff;font-size:.75em;padding:2px;">S</button>
						<input type="checkbox" style="display:none;" id="CONTROLMON_0" name="CONTROLDOTW_0" value="2"/>
						<button id="CONTROLBUTTONMON_0" type="button" class="btn <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" data-workgroup="<%=REQUEST_WORKGROUP%>" data-control="<%=REQUEST_CONTROL%>" style="background-color:#fff;font-size:.75em;padding:2px;">M</button>
						<input type="checkbox" style="display:none;" id="CONTROLTUE_0" name="CONTROLDOTW_0" value="3">
						<button id="CONTROLBUTTONTUE_0" type="button" class="btn <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" data-workgroup="<%=REQUEST_WORKGROUP%>" data-control="<%=REQUEST_CONTROL%>" style="background-color:#fff;font-size:.75em;padding:2px;">T</button>
						<input type="checkbox" style="display:none;" id="CONTROLWED_0" name="CONTROLDOTW_0" value="4"/>
						<button id="CONTROLBUTTONWED_0" type="button" class="btn <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" data-workgroup="<%=REQUEST_WORKGROUP%>" data-control="<%=REQUEST_CONTROL%>" style="background-color:#fff;font-size:.75em;padding:2px;">W</button>
						<input type="checkbox" style="display:none;" id="CONTROLTHU_0" name="CONTROLDOTW_0" value="5"/>
						<button id="CONTROLBUTTONTHU_0" type="button" class="btn <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" data-workgroup="<%=REQUEST_WORKGROUP%>" data-control="<%=REQUEST_CONTROL%>" style="background-color:#fff;font-size:.75em;padding:2px;">H</button>
						<input type="checkbox" style="display:none;" id="CONTROLFRI_0" name="CONTROLDOTW_0" value="6"/>
						<button id="CONTROLBUTTONFRI_0" type="button" class="btn <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" data-workgroup="<%=REQUEST_WORKGROUP%>" data-control="<%=REQUEST_CONTROL%>" style="background-color:#fff;font-size:.75em;padding:2px;">F</button>
						<input type="checkbox" style="display:none;" id="CONTROLSAT_0" name="CONTROLDOTW_0" value="7"/>
						<button id="CONTROLBUTTONSAT_0" type="button" class="btn <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" data-workgroup="<%=REQUEST_WORKGROUP%>" data-control="<%=REQUEST_CONTROL%>" style="background-color:#fff;font-size:.75em;padding:2px;">S</button>
					</td>
				<% Elseif FIELD = "INTERVAL" Then %>
					<% CONTROL_TYPE = "TIME" %>
					<% CONTROL_RANGE = "true" %>
					<% CONTROL_STEP = "30" %>
					<% CONTROL_INTERVAL = "30" %>
					<% If REQUEST_WORKGROUP = "RES" or REQUEST_WORKGROUP = "SPT" Then %>
						<% CONTROL_MIN = "360" %>
						<% CONTROL_MAX = "1440" %>
					<% Elseif REQUEST_WORKGROUP = "OSR" Then %>
						<% CONTROL_MIN = "0" %>
						<% CONTROL_MAX = "1440" %>
					<% End If %>
					<td class="subtable-td-padded-lg">
						<input type="hidden" id="CONTROLSTART_INTERVAL_0" name="CONTROLSTART_INTERVAL_0" value="" />
						<div style="display:inline-block;white-space:nowrap;">
							<i id="CONTROLSTARTARROW_INTERVAL_LEFT_0" class="fas fa-caret-left icon-style-small"></i>
							<div style="display:inline-block;" id="CONTROLSTARTDISPLAY_INTERVAL_0"></div>
							<i id="CONTROLSTARTARROW_INTERVAL_RIGHT_0" class="fas fa-caret-right icon-style-small" style="margin-right:5px;"></i>
						</div>
					</td>
					<td class="subtable-td-padded-lg">	
						<div id="CONTROLSLIDER_INTERVAL_0" style="display:inline-block;width:100%;" data-workgroup="<%=REQUEST_WORKGROUP%>" data-control="<%=REQUEST_CONTROL%>" data-slider-type="<%=CONTROL_TYPE%>" data-slider-min="<%=CONTROL_MIN%>" data-slider-max="<%=CONTROL_MAX%>" data-range="<%=CONTROL_RANGE%>" data-slider-step="<%=CONTROL_STEP%>" data-slider-interval="<%=CONTROL_INTERVAL%>"></div>	
					</td>
					<td class="subtable-td-padded-lg">
						<input type="hidden" id="CONTROLEND_INTERVAL_0" name="CONTROLEND_INTERVAL_0" value="" />
						<div style="display:inline-block;white-space:nowrap;">
							<i id="CONTROLENDARROW_INTERVAL_LEFT_0" class="fas fa-caret-left icon-style-small" style="margin-left:5px;"></i>
							<div style="display:inline-block;" id="CONTROLENDDISPLAY_INTERVAL_0"></div>
							<i id="CONTROLENDARROW_INTERVAL_RIGHT_0" class="fas fa-caret-right icon-style-small"></i>
						</div>
					</td>
				<% Elseif FIELD = "SCHEDULED_HOURS" Then %>					
					<% CONTROL_TYPE = "NUMBER" %>
					<% CONTROL_RANGE = "true" %>
					<% CONTROL_STEP = "1" %>
					<% CONTROL_INTERVAL = "1" %>
					<% CONTROL_MIN = "0" %>
					<% CONTROL_MAX = "40" %>
					<td class="subtable-td-padded-lg">
						<input type="hidden" id="CONTROLSTART_HOURS_0" name="CONTROLSTART_HOURS_0" value="" />
						<div style="display:inline-block;white-space:nowrap;">
							<i id="CONTROLSTARTARROW_HOURS_LEFT_0" class="fas fa-caret-left icon-style-small"></i>
							<div style="display:inline-block;width:1rem;" id="CONTROLSTARTDISPLAY_HOURS_0"></div>
							<i id="CONTROLSTARTARROW_HOURS_RIGHT_0" class="fas fa-caret-right icon-style-small" style="margin-right:5px;"></i>
						</div>
					</td>
					<td class="subtable-td-padded-lg">	
						<div id="CONTROLSLIDER_HOURS_0" style="display:inline-block;width:100%;" data-workgroup="<%=REQUEST_WORKGROUP%>" data-control="<%=REQUEST_CONTROL%>" data-slider-type="<%=CONTROL_TYPE%>" data-slider-min="<%=CONTROL_MIN%>" data-slider-max="<%=CONTROL_MAX%>" data-range="<%=CONTROL_RANGE%>" data-slider-step="<%=CONTROL_STEP%>" data-slider-interval="<%=CONTROL_INTERVAL%>"></div>	
					</td>
					<td class="subtable-td-padded-lg">
						<input type="hidden" id="CONTROLEND_HOURS_0" name="CONTROLEND_HOURS_0" value="" />
						<div style="display:inline-block;white-space:nowrap;">
							<i id="CONTROLENDARROW_HOURS_LEFT_0" class="fas fa-caret-left icon-style-small" style="margin-left:5px;"></i>
							<div style="display:inline-block;width:1rem;" id="CONTROLENDDISPLAY_HOURS_0"></div>
							<i id="CONTROLENDARROW_HOURS_RIGHT_0" class="fas fa-caret-right icon-style-small"></i>
						</div>
					</td>
				<% Elseif FIELD = "EVAL_SCORE" Then %>					
					<% CONTROL_TYPE = "NUMBER" %>
					<% CONTROL_RANGE = "true" %>
					<% CONTROL_STEP = ".25" %>
					<% CONTROL_INTERVAL = ".25" %>
					<% CONTROL_MIN = "1" %>
					<% CONTROL_MAX = "5" %>
					<td class="subtable-td-padded-lg">
						<input type="hidden" id="CONTROLSTART_SCORE_0" name="CONTROLSTART_SCORE_0" value="" />
						<div style="display:inline-block;white-space:nowrap;">
							<i id="CONTROLSTARTARROW_SCORE_LEFT_0" class="fas fa-caret-left icon-style-small"></i>
							<div style="display:inline-block;width:1.4rem;" id="CONTROLSTARTDISPLAY_SCORE_0"></div>
							<i id="CONTROLSTARTARROW_SCORE_RIGHT_0" class="fas fa-caret-right icon-style-small" style="margin-right:5px;"></i>
						</div>
					</td>
					<td class="subtable-td-padded-lg">	
						<div id="CONTROLSLIDER_SCORE_0" style="display:inline-block;width:100%;" data-workgroup="<%=REQUEST_WORKGROUP%>" data-control="<%=REQUEST_CONTROL%>" data-slider-type="<%=CONTROL_TYPE%>" data-slider-min="<%=CONTROL_MIN%>" data-slider-max="<%=CONTROL_MAX%>" data-range="<%=CONTROL_RANGE%>" data-slider-step="<%=CONTROL_STEP%>" data-slider-interval="<%=CONTROL_INTERVAL%>"></div>	
					</td>
					<td class="subtable-td-padded-lg">
						<input type="hidden" id="CONTROLEND_SCORE_0" name="CONTROLEND_SCORE_0" value="" />
						<div style="display:inline-block;white-space:nowrap;">
							<i id="CONTROLENDARROW_SCORE_LEFT_0" class="fas fa-caret-left icon-style-small" style="margin-left:5px;"></i>
							<div style="display:inline-block;width:1.4rem;" id="CONTROLENDDISPLAY_SCORE_0"></div>
							<i id="CONTROLENDARROW_SCORE_RIGHT_0" class="fas fa-caret-right icon-style-small"></i>
						</div>
					</td>
				<% Elseif FIELD = "REDUCE_HOURS" or FIELD = "ADD_HOURS" or FIELD = "NEW_DAYS" Then %>
					<% CONTROL_TYPE = "NUMBER" %>
					<% CONTROL_RANGE = "false" %>
					<% CONTROL_STEP = "1" %>
					<% CONTROL_INTERVAL = "1" %>
					<% CONTROL_MIN = "0" %>
					<% If FIELD = "REDUCE_HOURS" Then %>
						<% CONTROL_MAX = "40" %>
					<% Elseif FIELD = "ADD_HOURS" Then %>
						<% CONTROL_MAX = "50" %>
					<% Elseif FIELD = "NEW_DAYS" Then %>
						<% CONTROL_MAX = "90" %>
					<% End If %>
					<td class="subtable-td-padded-lg">
						<input type="hidden" id="CONTROL_VALUE_0" name="CONTROL_VALUE_0" value="" />
						<div id="CONTROLSLIDER_VALUE_0" style="display:inline-block;width:100%;" data-workgroup="<%=REQUEST_WORKGROUP%>" data-control="<%=REQUEST_CONTROL%>" data-slider-type="<%=CONTROL_TYPE%>" data-slider-min="<%=CONTROL_MIN%>" data-slider-max="<%=CONTROL_MAX%>" data-range="<%=CONTROL_RANGE%>" data-slider-step="<%=CONTROL_STEP%>" data-slider-interval="<%=CONTROL_INTERVAL%>">
							<div id="CONTROLHANDLE_VALUE_0" class="ui-slider-handle custom-handle CONTROL" style="color:#fff;"></div>
						</div>	
					</td>
				<% End If %>				
			<% Next %>
				<td class="subtable-td-padded-lg">
					<input type="text" id="CONTROLEFFDATE_0" name="CONTROLEFFDATE_0" class="center no-border-input control-date <% If PARAMETER_DATE = Date Then %> today-color <% Else %> past-color <% End If %>" style="max-width:100px;" value="<% If ControlDateType(ControlDecoder(REQUEST_CONTROL),"START") = "sunday" Then %><%=PARAMETER_DATE - Weekday(PARAMETER_DATE) + 1%><% Else %><%=PARAMETER_DATE%><% End If %>" data-date-min="<% If ControlDateType(ControlDecoder(REQUEST_CONTROL),"START") = "sunday" Then %><%=MIN_DATE - Weekday(MIN_DATE) + 1%><% Else %><%=MIN_DATE%><% End If %>" data-date-max="12/31/2040" data-date-type="<%=ControlDateType(ControlDecoder(REQUEST_CONTROL),"START")%>"/>
				</td>
				<td class="subtable-td-padded-lg">
					<input type="text" id="CONTROLDISDATE_0" name="CONTROLDISDATE_0" class="center no-border-input control-date <% If PARAMETER_DATE = Date Then %> today-color <% Else %> past-color <% End If %>" style="max-width:100px;" value="12/31/2040" data-date-min="<% If ControlDateType(ControlDecoder(REQUEST_CONTROL),"START") = "saturday" Then %><%=MIN_DATE - Weekday(MIN_DATE) + 7%><% Else %><%=MIN_DATE%><% End If %>" data-date-max="12/31/2040" data-date-type="<%=ControlDateType(ControlDecoder(REQUEST_CONTROL),"END")%>"/>
				</td>
		</tr>
		<tr>
			<td class="subtable-td-padded-lg" colspan="<%=USE_COLSPAN%>">
				<div style="margin-top:5px;">
					<div style="float:left;margin-left:15px;">
						<i id="CONTROLREFRESH_<%=REQUEST_WORKGROUP%>_<%=REQUEST_CONTROL%>" class="fas fa-sync-alt icon-style-large" title="Refresh"></i>
						<i id="PULSEHELP_CONTROL" class="fas fa-question icon-style-large" title="Help"></i>
					</div>
					<i id="NEWCONTROLENTRY_<%=REQUEST_WORKGROUP%>_<%=REQUEST_CONTROL%>" class="fas fa-plus-square icon-style-large" title="New Entry"></i>
				</div>
			</td>
		</tr>
	</table>
	<script>
		$(document).ready(function() {
			$(".control-date:not([id$='_0'])").each(function(){
				$(this).datepicker({
					dateFormat: "m/d/yy",
					showAnim: "slideDown",
					showOtherMonths: true,
					selectOtherMonths: true,
					minDate: new Date($(this).data("date-min")),
					maxDate: new Date($(this).data("date-max")),
					beforeShowDay: function(dateText) {
						if($(this).data("date-type") == "sunday"){
							return [dateText.getDay() == 0,""]
						}
						else if ($(this).data("date-type") == "saturday"){
							return [dateText.getDay() == 6,""]
						}
						else{
							return [1,""]
						}
					},
					onClose: function(dateText) {
						if(moment(dateText,"M/D/YYYY", true).isValid()){
							$(this).val(moment.min(moment.max(moment(dateText,"MM/DD/YYYY"),moment($(this).data("date-min"),"MM/DD/YYYY")),moment($(this).data("date-max"),"MM/DD/YYYY")).format("l"));
							if($(this).data("date-type") == "sunday"){
								$(this).val(moment($(this).val(),"MM/DD/YYYY").day(0).format("l"));
							}
							else if ($(this).data("date-type") == "saturday" && $(this).val() != "12/31/2040"){
								$(this).val(moment($(this).val(),"MM/DD/YYYY").day(6).format("l"));
							}
							var idArray = this.id.split("_");
							checkControlOverlaps("<%=REQUEST_WORKGROUP%>","<%=REQUEST_CONTROL%>");
							addControlList(idArray[1]);
						}
						else{
							alert("Invalid date. Enter a valid date in MM/DD/YYYY format");
						}
					}
				});
			});
			$("div[id^='CONTROLSLIDER_'][data-workgroup='<%=REQUEST_WORKGROUP%>'][data-control='<%=REQUEST_CONTROL%>']:not([id$='_0'])").each(function(){
				var idArray = this.id.split("_");
				initializeControlSlider(idArray[1], idArray[2]);
			})
		});
	</script>
<!--#include file="pulsefunctions.asp"-->
<% Set cmd = Nothing %>
<% Conn.Close %>
<% Set Conn = Nothing %>