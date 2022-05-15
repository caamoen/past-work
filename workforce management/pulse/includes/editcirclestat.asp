<!--#include file="pulseheader.asp"-->
<%
	If Request.Querystring("TYPE") <> "" Then
		REQUEST_TYPE = Request.Querystring("TYPE")
	Else
		REQUEST_TYPE = "VOLUME"
	End If
	If Request.Querystring("DATE") <> "" Then
		PARAMETER_DATE = CDate(Request.Querystring("DATE"))
	Else
		PARAMETER_DATE = Date
	End If
	    
	SLIDER_MIN = 0
	Select Case REQUEST_TYPE
		Case "VOLUME", "AHT"
			SLIDER_MAX = 150
		Case "FTE"
			SLIDER_MAX = 200
		Case "SL", "PCH", "AVAILABILITY"
			SLIDER_MAX = 100
		Case "RESASA", "SPTASA"
			SLIDER_MAX = 300
		Case Else
			SLIDER_MAX = 100
	End Select
	If REQUEST_TYPE = "RESASA" or REQUEST_TYPE = "SPTASA" Then
		SLIDER_TYPE = "TIME"
	Else
		SLIDER_TYPE = "NUMBER"
	End If
	
	SQLstmt = "SELECT * " & _
	"FROM " & _
	"( " & _
		"SELECT " & _
		"OPS_PAR_ID, " & _
		"REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,2) CIRCLE_COLOR, " & _
		"TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3)) CIRCLE_START, " & _
		"TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,4)) CIRCLE_END, " & _
		"DECODE(OPS_PAR_CODE,'RESASA',TO_CHAR(TO_DATE(TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3)),'SSSSS'),'MI:SS'),'SPTASA',TO_CHAR(TO_DATE(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3),'SSSSS'),'MI:SS'),REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,3) || '%') DISPLAY_START, " & _
		"DECODE(OPS_PAR_CODE,'RESASA',TO_CHAR(TO_DATE(TO_NUMBER(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,4)),'SSSSS'),'MI:SS'),'SPTASA',TO_CHAR(TO_DATE(REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,4),'SSSSS'),'MI:SS'),REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,4) || '%') DISPLAY_END, " & _
		"REGEXP_SUBSTR(OPS_PAR_VALUE,'[^;]+',1,1) CIRCLE_DOTW, " & _
		"OPS_PAR_EFF_DATE CIRCLE_EFF_DATE, " & _
		"OPS_PAR_DIS_DATE CIRCLE_DIS_DATE, " & _
		"MIN(CASE WHEN OPS_PAR_DIS_DATE >= TO_DATE(?,'MM/DD/YYYY') THEN OPS_PAR_EFF_DATE END) OVER () MIN_DATE " & _
		"FROM OPS_PARAMETER " & _
		"WHERE OPS_PAR_PARENT_TYPE = 'STF' " & _
		"AND OPS_PAR_CODE = ? " & _
	") " & _
	"WHERE CIRCLE_EFF_DATE >= MIN_DATE " & _
	"ORDER BY CIRCLE_EFF_DATE, CIRCLE_DIS_DATE, CIRCLE_START"
	cmd.CommandText = SQLstmt
	cmd.Parameters(0).value = PARAMETER_DATE
	cmd.Parameters(1).value = REQUEST_TYPE
	Set RSCIRCLE = cmd.Execute
	If Not RSCIRCLE.EOF Then
		MIN_DATE = CDate(RSCIRCLE("MIN_DATE"))
	Else
		MIN_DATE = Date
	End If
%>
	<table id="CIRCLETABLE_<%=REQUEST_TYPE%>" class="center" style="width:100%;font-size:.75em;">
		<caption class="center <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>">
			<%=Replace(Replace(Replace(Replace(Replace(Replace(Replace(REQUEST_TYPE,"VOLUME","Volume Variance"),"AHT","AHT Variance"),"FTE","FTE Variance"),"SL","Service Level"),"AVAILABILITY","Availability"),"RESASA","Reservations ASA"),"SPTASA","Support ASA")%> Parameters
		</caption>
		<thead>
			<tr class="subtable-td-padded-sm <% If PARAMETER_DATE = Date Then %>today-color<% Else %>past-color<% End If %>">
				<th class="subtable-td-padded-sm" style="width:5%;">&nbsp;</th>
				<th class="subtable-td-padded-sm" style="width:50%;">Value</th>
				<th class="subtable-td-padded-sm" style="width:5%;">&nbsp;</th>
				<th class="subtable-td-padded-sm" style="width:5%;">Color</th>
				<th class="subtable-td-padded-sm" style="width:15%;">Day</th>
				<th class="subtable-td-padded-sm" style="width:10%;">Eff. Date</th>
				<th class="subtable-td-padded-sm" style="width:10%;">Dis Date</th>
			</tr>
		</thead>
		<tbody>
		<% Do While Not RSCIRCLE.EOF %>
			<tr>
				<td class="subtable-td-padded-lg">
					<input type="hidden" id="CIRCLEREQUEST_<%=RSCIRCLE("OPS_PAR_ID")%>" name="CIRCLEREQUEST_<%=RSCIRCLE("OPS_PAR_ID")%>" value="<%=REQUEST_TYPE%>" />
					<input type="hidden" id="CIRCLESTART_<%=RSCIRCLE("OPS_PAR_ID")%>" name="CIRCLESTART_<%=RSCIRCLE("OPS_PAR_ID")%>" value="<%=RSCIRCLE("CIRCLE_START")%>" />
					<div style="display:inline-block;white-space:nowrap;">
						<i id="CIRCLESTARTARROW_LEFT_<%=RSCIRCLE("OPS_PAR_ID")%>" class="fas fa-caret-left icon-style-small"></i>
						<span id="CIRCLESTARTDISPLAY_<%=RSCIRCLE("OPS_PAR_ID")%>"><%=RSCIRCLE("DISPLAY_START")%></span>
						<i id="CIRCLESTARTARROW_RIGHT_<%=RSCIRCLE("OPS_PAR_ID")%>" class="fas fa-caret-right icon-style-small"></i>
					</div>
				</td>
				<td class="subtable-td-padded-lg">	
					<div id="CIRCLESLIDER_<%=RSCIRCLE("OPS_PAR_ID")%>" style="display:inline-block;width:100%;" data-request-type="<%=REQUEST_TYPE%>" data-slider-type="<%=SLIDER_TYPE%>" data-slider-min="<%=SLIDER_MIN%>" data-slider-max="<%=SLIDER_MAX%>" data-slider-step="1" data-slider-interval="1"></div>	
				</td>
				<td class="subtable-td-padded-lg">
					<input type="hidden" id="CIRCLEEND_<%=RSCIRCLE("OPS_PAR_ID")%>" name="CIRCLEEND_<%=RSCIRCLE("OPS_PAR_ID")%>" value="<%=RSCIRCLE("CIRCLE_END")%>" />
					<div style="display:inline-block;white-space:nowrap;">
						<i id="CIRCLEENDARROW_LEFT_<%=RSCIRCLE("OPS_PAR_ID")%>" class="fas fa-caret-left icon-style-small"></i>
						<span id="CIRCLEENDDISPLAY_<%=RSCIRCLE("OPS_PAR_ID")%>"><%=RSCIRCLE("DISPLAY_END")%></span>
						<i id="CIRCLEENDARROW_RIGHT_<%=RSCIRCLE("OPS_PAR_ID")%>" class="fas fa-caret-right icon-style-small"></i>
					</div>
				</td>
				<td class="subtable-td-padded-lg">
					<select id="CIRCLECOLOR_<%=RSCIRCLE("OPS_PAR_ID")%>" name="CIRCLECOLOR_<%=RSCIRCLE("OPS_PAR_ID")%>" class="<% If PARAMETER_DATE = Date Then %> today-color <% Else %> past-color <% End If %>" style="padding-left:8px;">
						<option data-circle-class="GREEN" <% If RSCIRCLE("CIRCLE_COLOR") = "G" Then %>selected="selected"<% End If %> value="G">Green</option>
						<option data-circle-class="YELLOW" <% If RSCIRCLE("CIRCLE_COLOR") = "Y" Then %>selected="selected"<% End If %> value="Y">Yellow</option>
					</select>
				</td>
				<td class="subtable-td-padded-lg nowrap">
					<input type="checkbox" style="display:none;" id="CIRCLESUN_<%=RSCIRCLE("OPS_PAR_ID")%>" name="CIRCLEDOTW_<%=RSCIRCLE("OPS_PAR_ID")%>" value="1" <% If Instr(RSCIRCLE("CIRCLE_DOTW"),"1") Then %> checked="checked" <% End If %>/>
					<button id="CIRCLEBUTTONSUN_<%=RSCIRCLE("OPS_PAR_ID")%>" type="button" class="btn <% If Instr(RSCIRCLE("CIRCLE_DOTW"),"1") Then %> new-entry-color <% End If %> <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" style="background-color:#fff;font-size:.75em;padding:2px;">S</button>
					<input type="checkbox" style="display:none;" id="CIRCLEMON_<%=RSCIRCLE("OPS_PAR_ID")%>" name="CIRCLEDOTW_<%=RSCIRCLE("OPS_PAR_ID")%>" value="2" <% If Instr(RSCIRCLE("CIRCLE_DOTW"),"2") Then %> checked="checked" <% End If %>/>
					<button id="CIRCLEBUTTONMON_<%=RSCIRCLE("OPS_PAR_ID")%>" type="button" class="btn <% If Instr(RSCIRCLE("CIRCLE_DOTW"),"2") Then %> new-entry-color <% End If %> <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" style="background-color:#fff;font-size:.75em;padding:2px;">M</button>
					<input type="checkbox" style="display:none;" id="CIRCLETUE_<%=RSCIRCLE("OPS_PAR_ID")%>" name="CIRCLEDOTW_<%=RSCIRCLE("OPS_PAR_ID")%>" value="3" <% If Instr(RSCIRCLE("CIRCLE_DOTW"),"3") Then %> checked="checked" <% End If %>/>
					<button id="CIRCLEBUTTONTUE_<%=RSCIRCLE("OPS_PAR_ID")%>" type="button" class="btn <% If Instr(RSCIRCLE("CIRCLE_DOTW"),"3") Then %> new-entry-color <% End If %> <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" style="background-color:#fff;font-size:.75em;padding:2px;">T</button>
					<input type="checkbox" style="display:none;" id="CIRCLEWED_<%=RSCIRCLE("OPS_PAR_ID")%>" name="CIRCLEDOTW_<%=RSCIRCLE("OPS_PAR_ID")%>" value="4" <% If Instr(RSCIRCLE("CIRCLE_DOTW"),"4") Then %> checked="checked" <% End If %>/>
					<button id="CIRCLEBUTTONWED_<%=RSCIRCLE("OPS_PAR_ID")%>" type="button" class="btn <% If Instr(RSCIRCLE("CIRCLE_DOTW"),"4") Then %> new-entry-color <% End If %> <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" style="background-color:#fff;font-size:.75em;padding:2px;">W</button>
					<input type="checkbox" style="display:none;" id="CIRCLETHU_<%=RSCIRCLE("OPS_PAR_ID")%>" name="CIRCLEDOTW_<%=RSCIRCLE("OPS_PAR_ID")%>" value="5" <% If Instr(RSCIRCLE("CIRCLE_DOTW"),"5") Then %> checked="checked" <% End If %>/>
					<button id="CIRCLEBUTTONTHU_<%=RSCIRCLE("OPS_PAR_ID")%>" type="button" class="btn <% If Instr(RSCIRCLE("CIRCLE_DOTW"),"5") Then %> new-entry-color <% End If %> <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" style="background-color:#fff;font-size:.75em;padding:2px;">H</button>
					<input type="checkbox" style="display:none;" id="CIRCLEFRI_<%=RSCIRCLE("OPS_PAR_ID")%>" name="CIRCLEDOTW_<%=RSCIRCLE("OPS_PAR_ID")%>" value="6" <% If Instr(RSCIRCLE("CIRCLE_DOTW"),"6") Then %> checked="checked" <% End If %>/>
					<button id="CIRCLEBUTTONFRI_<%=RSCIRCLE("OPS_PAR_ID")%>" type="button" class="btn <% If Instr(RSCIRCLE("CIRCLE_DOTW"),"6") Then %> new-entry-color <% End If %> <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" style="background-color:#fff;font-size:.75em;padding:2px;">F</button>
					<input type="checkbox" style="display:none;" id="CIRCLESAT_<%=RSCIRCLE("OPS_PAR_ID")%>" name="CIRCLEDOTW_<%=RSCIRCLE("OPS_PAR_ID")%>" value="7" <% If Instr(RSCIRCLE("CIRCLE_DOTW"),"7") Then %> checked="checked" <% End If %>/>
					<button id="CIRCLEBUTTONSAT_<%=RSCIRCLE("OPS_PAR_ID")%>" type="button" class="btn <% If Instr(RSCIRCLE("CIRCLE_DOTW"),"7") Then %> new-entry-color <% End If %> <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" style="background-color:#fff;font-size:.75em;padding:2px;">S</button>
				</td>
				<td class="subtable-td-padded-lg">
					<input type="text" id="CIRCLEEFFDATE_<%=RSCIRCLE("OPS_PAR_ID")%>" name="CIRCLEEFFDATE_<%=RSCIRCLE("OPS_PAR_ID")%>" class="center no-border-input circle-date <% If PARAMETER_DATE = Date Then %> today-color <% Else %> past-color <% End If %>" style="max-width:100px;" value="<%=RSCIRCLE("CIRCLE_EFF_DATE")%>" data-date-min="<%=MIN_DATE%>" data-date-max="12/31/2040"/>
				</td>
				<td class="subtable-td-padded-lg">
					<input type="text" id="CIRCLEDISDATE_<%=RSCIRCLE("OPS_PAR_ID")%>" name="CIRCLEDISDATE_<%=RSCIRCLE("OPS_PAR_ID")%>" class="center no-border-input circle-date <% If PARAMETER_DATE = Date Then %> today-color <% Else %> past-color <% End If %>" style="max-width:100px;" value="<%=RSCIRCLE("CIRCLE_DIS_DATE")%>" data-date-min="<%=MIN_DATE%>" data-date-max="12/31/2040"/>
				</td>
			</tr>
			<% RSCIRCLE.MoveNext %>
		<% Loop %>
		</tbody>
		<% Set RSCIRCLE = Nothing %>
		<tr class="new-entry-color" style="display:none;">
			<td class="subtable-td-padded-lg">
				<input type="hidden" id="CIRCLEREQUEST_0" name="CIRCLEREQUEST_0" value="<%=REQUEST_TYPE%>" />
				<input type="hidden" id="CIRCLESTART_0" name="CIRCLESTART_0" value="" />
				<div style="display:inline-block;white-space:nowrap;">
					<i id="CIRCLESTARTARROW_LEFT_0" class="fas fa-caret-left icon-style-small"></i>
					<span id="CIRCLESTARTDISPLAY_0"></span>
					<i id="CIRCLESTARTARROW_RIGHT_0" class="fas fa-caret-right icon-style-small"></i>
				</div>
			</td>
			<td class="subtable-td-padded-lg">	
				<div id="CIRCLESLIDER_0" style="display:inline-block;width:100%;" data-request-type="<%=REQUEST_TYPE%>" data-slider-type="<%=SLIDER_TYPE%>" data-slider-min="<%=SLIDER_MIN%>" data-slider-max="<%=SLIDER_MAX%>" data-slider-step="1" data-slider-interval="1"></div>	
			</td>
			<td class="subtable-td-padded-lg">
				<input type="hidden" id="CIRCLEEND_0" name="CIRCLEEND_0" value="" />
				<div style="display:inline-block;white-space:nowrap;">
					<i id="CIRCLEENDARROW_LEFT_0" class="fas fa-caret-left icon-style-small"></i>
					<span id="CIRCLEENDDISPLAY_0"></span>
					<i id="CIRCLEENDARROW_RIGHT_0" class="fas fa-caret-right icon-style-small"></i>
				</div>
			</td>
			<td class="subtable-td-padded-lg">
				<select id="CIRCLECOLOR_0" name="CIRCLECOLOR_0" class="new-entry-color <% If PARAMETER_DATE = Date Then %> today-color <% Else %> past-color <% End If %>" style="padding-left:8px;">
					<option data-circle-class="GREEN" value="G">Green</option>
					<option data-circle-class="YELLOW" value="Y">Yellow</option>
				</select>
			</td>
			<td class="subtable-td-padded-lg nowrap">
				<input type="checkbox" style="display:none;" id="CIRCLESUN_0" name="CIRCLEDOTW_0" value="1"/>
				<button id="CIRCLEBUTTONSUN_0" type="button" class="btn <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" style="background-color:#fff;font-size:.75em;padding:2px;">S</button>
				<input type="checkbox" style="display:none;" id="CIRCLEMON_0" name="CIRCLEDOTW_0" value="2"/>
				<button id="CIRCLEBUTTONMON_0" type="button" class="btn <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" style="background-color:#fff;font-size:.75em;padding:2px;">M</button>
				<input type="checkbox" style="display:none;" id="CIRCLETUE_0" name="CIRCLEDOTW_0" value="3"/>
				<button id="CIRCLEBUTTONTUE_0" type="button" class="btn <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" style="background-color:#fff;font-size:.75em;padding:2px;">T</button>
				<input type="checkbox" style="display:none;" id="CIRCLEWED_0" name="CIRCLEDOTW_0" value="4"/>
				<button id="CIRCLEBUTTONWED_0" type="button" class="btn <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" style="background-color:#fff;font-size:.75em;padding:2px;">W</button>
				<input type="checkbox" style="display:none;" id="CIRCLETHU_0" name="CIRCLEDOTW_0" value="5"/>
				<button id="CIRCLEBUTTONTHU_0" type="button" class="btn <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" style="background-color:#fff;font-size:.75em;padding:2px;">T</button>
				<input type="checkbox" style="display:none;" id="CIRCLEFRI_0" name="CIRCLEDOTW_0" value="6"/>
				<button id="CIRCLEBUTTONFRI_0" type="button" class="btn <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" style="background-color:#fff;font-size:.75em;padding:2px;">F</button>
				<input type="checkbox" style="display:none;" id="CIRCLESAT_0" name="CIRCLEDOTW_0" value="7"/>
				<button id="CIRCLEBUTTONSAT_0" type="button" class="btn <% If PARAMETER_DATE = Date Then %>today-color today-color-border<% Else %>past-color past-color-border<% End If %>" style="background-color:#fff;font-size:.75em;padding:2px;">S</button>
			</td>
			<td class="subtable-td-padded-lg">
				<input type="text" id="CIRCLEEFFDATE_0" name="CIRCLEEFFDATE_0" class="center no-border-input circle-date <% If PARAMETER_DATE = Date Then %> today-color <% Else %> past-color <% End If %>" style="max-width:100px;" value="<%=Date%>" data-date-min="<%=MIN_DATE%>" data-date-max="12/31/2040"/>
			</td>
			<td class="subtable-td-padded-lg">
				<input type="text" id="CIRCLEDISDATE_0" name="CIRCLEDISDATE_0" class="center no-border-input circle-date <% If PARAMETER_DATE = Date Then %> today-color <% Else %> past-color <% End If %>" style="max-width:100px;" value="12/31/2040" data-date-min="<%=MIN_DATE%>" data-date-max="12/31/2040"/>
			</td>
		</tr>
		<tr>
			<td class="subtable-td-padded-lg" colspan="7">
				<div style="margin-top:5px;">
					<div style="float:left;margin-left:15px;">
						<i id="CIRCLEREFRESH_<%=REQUEST_TYPE%>" class="fas fa-sync-alt icon-style-large" title="Refresh"></i>
						<i id="PULSEHELP_CIRCLE" class="fas fa-question icon-style-large" title="Help"></i>
					</div>
					<i id="NEWCIRCLEENTRY_<%=REQUEST_TYPE%>" class="fas fa-plus-square icon-style-large" title="New Entry"></i>
				</div>
			</td>
		</tr>
	</table>
	<script>
		$(document).ready(function() {
			$(".circle-date:not([id$='_0'])").each(function(){
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
							checkCircleOverlaps("<%=REQUEST_TYPE%>");
							addCircleList(idArray[1]);
						}
						else{
							alert("Invalid date. Enter a valid date in MM/DD/YYYY format");
						}
					}
				});
			});
			$("div[id^='CIRCLESLIDER_'][data-request-type='<%=REQUEST_TYPE%>']:not([id='CIRCLESLIDER_0'])").each(function(){
				var idArray = this.id.split("_");
				initializeCircleSlider(idArray[1]);
			})
		});
	</script>
	
<!--#include file="pulsefunctions.asp"-->
<% Set cmd = Nothing %>
<% Conn.Close %>
<% Set Conn = Nothing %>