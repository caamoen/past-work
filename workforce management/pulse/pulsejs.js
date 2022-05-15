	var enterTimestamp;
	var parameterDate = moment().format('L');
	
	var overlapIdList = [];
	var scheduleIdList = [];
	var noteIdList = [];
	var circleIdList = [];
	var controlIdList = [];
	var adminIdList = [];
	
	var liveCMSTimeout;
	var liveCMSRefreshTimeout;
	var currentDayTimeout;
	var scheduleStatsTimeout;
	
	var XHRliveCMS = new XMLHttpRequest();
	var XHRcurrentDay = new XMLHttpRequest();
	var XHRscheduleStats = new XMLHttpRequest();	
	var XHRmain = new XMLHttpRequest();
	var XHRchosenOptions = new XMLHttpRequest();
	
/* General - Start */
	$(document).ready(function() {
		pulseMode($("#PULSE_MODE").val());
		$("#PARAMETER_DATE").datepicker({
			dateFormat: "m/d/yy",
			showAnim: "slideDown",
			showOtherMonths: true,
			selectOtherMonths: true,
			onClose: function(dateText) {
				if(moment(dateText,"M/D/YYYY", true).isValid()){
					var inputDate = moment(dateText,"MM/DD/YYYY").format('L');
					var today = moment().format('L');
					if(today == inputDate){
						$(".past-color-background").removeClass("past-color-background").addClass("today-color-background");
						$(".past-color").removeClass("past-color").addClass("today-color");
						$(".past-color-border").removeClass("past-color-border").addClass("today-color-border");
					}
					else{
						$(".today-color-background").removeClass("today-color-background").addClass("past-color-background");
						$(".today-color").removeClass("today-color").addClass("past-color");
						$(".today-color-border").removeClass("today-color-border").addClass("past-color-border");
					}
					if(parameterDate != inputDate){
						$.when(chosenOptions($("#PULSE_MODE").val())).then(function(){
							if($("#PULSE_MODE").val() == "STAFFING"){
								if($("#PULSE_FORM").length && ($("#PULSE_FORM").data("request") == "PENDING" || $("#PULSE_FORM").data("request") == "ERROR" || $("#search-bar").chosen().val() != "")){
									if ($("#PULSE_FORM").data("request") == "PENDING"){
										pendingList($("#PULSE_FORM").data("workgroup"));
									}
									else if ($("#PULSE_FORM").data("request") == "ERROR"){
										errorList($("#PULSE_FORM").data("workgroup"));
									}
									else{
										$('.chosen-select').trigger('chosen:close');
										getAgentList();
									}
								}
								else if($("#NOTE_FORM").length && $("#NOTE_FORM").data("workgroup") !== undefined){
									noteList($("#NOTE_FORM").data("workgroup"));
								}
								else if($(".staffing_graph").length){
									staffingGraph($(".staffing_graph").data("workgroup"));
								}
								else if($(".data-table").length){
									dataChart($(".data-table").data("workgroup"));
								}
								else{
									dailyStats();
								}
								scheduleStats();
								currentDay();
							}
							else if($("#PULSE_MODE").val() == "ADMIN"){
								if($("#PULSE_FORM").length && $("#search-bar").chosen().val() != ""){
										$('.chosen-select').trigger('chosen:close');
										getAgentList();
								}
							}
						});
						parameterDate = inputDate;
					}
				}
				else{
					alert("Invalid date. Enter a valid date in MM/DD/YYYY format");
				}
			}
		});
		$(window).focus(function() {
			if($("#PULSE_MODE").val() == "STAFFING"){
				liveCMSTimeout = setTimeout(liveCMS,1000*Math.max(Math.min(60-10*$("#PULSE_SECURITY").val(),30),10));
				scheduleStatsTimeout = setTimeout(scheduleStats,1000*Math.max(Math.min(10800-1800*$("#PULSE_SECURITY").val(),3600),1800));
				currentDayTimeout = setTimeout(currentDay,1000*Math.max(Math.min(7560-1440*$("#PULSE_SECURITY").val(),1800),360));
			}
		});
		$(window).blur(function() {
			clearTimeout(liveCMSTimeout);
			clearTimeout(scheduleStatsTimeout);
			clearTimeout(currentDayTimeout);
		});
		$("#main-wrap").on("click","[id^=PULSEHELP_]",function() {
			var idArray = this.id.split("_");
			var bPopup = $("#popupdiv").bPopup({
				speed: 650,
				transition: "slideDown",
				transitionClose: "slideUp",
				content:"ajax",
				contentContainer:"#popupdiv",
				loadUrl: "includes/help.asp?TYPE=" + idArray[1] + "&JUNK="+ new Date().getTime()
			});
		});
	});
/* General - End */

/* Title Bar Buttons - Start */
	function pulseMode(useMode){
		$(".chosen-select").val("").trigger("chosen:updated");
		if(useMode == "STAFFING"){
			chosenOptions("STAFFING");
			$("#dashboard-left").css("display","flex");
			$("#dashboard-right").css("display","block");
			$(".staffgroup").css("display","inline-flex");
			$(".admingroup").css("display","none");
			liveCMS();
			scheduleStats();
			currentDay();
			dailyStats();
		}
		else{
			chosenOptions("ADMIN");
			$("#dashboard-left").css("display","none").empty();
			$("#dashboard-right").css("display","none");
			$("#schedule-stats-div").empty();
			$("#current-day-div").empty();
			$("#dashboarddiv-wide").empty();
			$(".staffgroup").css("display","none");
			$(".admingroup").css("display","inline-flex");
			XHRmain.abort();
			XHRliveCMS.abort();
			clearTimeout(liveCMSTimeout);
			clearTimeout(liveRefreshCMSTimeout);
			liveRefreshCMSTimeout = null;
			XHRscheduleStats.abort();
			clearTimeout(scheduleStatsTimeout);
			XHRcurrentDay.abort();
			clearTimeout(currentDayTimeout);
		}
	}
	$(document).ready(function() {
		$("#title-bar").on("click","#HOME_ICON",function(event) {
			event.preventDefault();
			if($("#search-bar").chosen().val() != ""){
				$(".chosen-select").val("").trigger("chosen:updated");
			}
			else{
				if($(".circle-stat").length || $(".staffdiv").length || $("#PULSE_MODE").val() == "ADMIN"){
					$("#PARAMETER_DATE").val(moment().format('l'));
					parameterDate = moment().format('L');
					$(".past-color-background").removeClass("past-color-background").addClass("today-color-background");
					$(".past-color").removeClass("past-color").addClass("today-color");
					$(".past-color-border").removeClass("past-color-border").addClass("today-color-border");
					chosenOptions($("#PULSE_MODE").val());
				}
				$(".chosen-select").val("").trigger("chosen:updated");
				if($("#PULSE_MODE").val() == "STAFFING"){
					dailyStats();
					scheduleStats();
					currentDay();
				}
				else if($("#PULSE_MODE").val() == "ADMIN"){
					$("#dashboarddiv-wide").empty();
				}
			}
		});
		$("#title-bar").on("change","#PULSE_MODE",function(){return pulseMode($(this).val());});
	});
/* Title Bar Buttons - End */

/* Agent List - Start */
	function chosenOptions(useMode){
		if (useMode === undefined){
			useMode = "STAFFING";
		}
		XHRchosenOptions.abort();
		return XHRchosenOptions = $.ajax({
			url: "includes/chosenoptions.asp?DATE=" + $("#PARAMETER_DATE").val() + "&MODE=" + useMode,
			success: function(result){
				var chosenArray = $("#search-bar").chosen().val();
				$("#search-bar").html(result).val(chosenArray).trigger('chosen:updated');
			},
			cache: false
		});	
	}
	function getAgentList(){
		var searchValue = $("#search-bar").chosen().val();
		var agentValue = [];
		var supervisorValue = [];
		var departmentValue = [];
		var workgroupValue = [];
		var classValue = [];
		var locationValue = [];
		var jobValue = [];
		var shiftValue = [];
		var timesValue = [];
		var hireValue = [];
		var trainingValue = [];
		var routingValue = [];
		var chatValue = "";
		
		searchValue.forEach(function(element) {
			var elementArray = element.split("_");
			switch(elementArray[0]){
				case "AGT":
					agentValue.push(elementArray[1])
					break;
				case "SUP":
					supervisorValue.push(elementArray[1])
					break;
				case "DEPT":
					departmentValue.push(elementArray[1])
					break;
				case "WRK":
					workgroupValue.push(elementArray[1])
					break;
				case "CLASS":
					classValue.push(elementArray[1])
					break;
				case "LOC":
					locationValue.push(elementArray[1])
					break;
				case "JOB":
					jobValue.push(elementArray[1])
					break;
				case "SCH":
					shiftValue.push(elementArray[1])
					break;
				case "TIMES":
					timesValue.push(elementArray[1])
					break;
				case "HIRE":
					hireValue.push(elementArray[1])
					break;
				case "TRN":
					trainingValue.push(elementArray[1])
					break;
				case "ROUT":
					routingValue.push(elementArray[1])
					break;
				case "CHAT":
					chatValue = "1";
					break;
			}
		});
		XHRmain.abort();
		if($("#PULSE_MODE").val() == "STAFFING"){
			XHRmain = $.ajax({
				url: "includes/agentlist.asp?REQUEST=SEARCH&DATE=" + $("#PARAMETER_DATE").val() + "&AGENT=" + agentValue + "&SUPERVISOR=" + supervisorValue + "&DEPARTMENT=" + departmentValue + "&WORKGROUP=" + workgroupValue + "&CLASS=" + classValue + "&LOCATION=" + locationValue + "&JOB=" + jobValue + "&SHIFT=" + shiftValue + "&TIMES=" + timesValue + "&HIRE=" + hireValue + "&TRAINING=" + trainingValue + "&ROUTING=" + routingValue + "&CHAT=" + chatValue,
				success: function(result){
					$("#dashboarddiv-wide").css("display","block").html(result);
					
					$("#PULSE_FORM").ajaxForm({
						beforeSerialize: function() { 
							$("#SCHEDULEID_LIST").val(scheduleIdList.join());
						},
						clearForm: true,
						error: function() {
							$("#dashboarddiv-wide").css("display","block").html("Update Failed");
						},
						success: function() {
							getAgentList();
							currentDay();
						}
					});
				},
				complete: liveStateAgentList,
				error: function(jqXHR, textStatus) {
					if(textStatus != "abort"){
						$("#dashboarddiv-wide").css("display","block").html("Failed to load list. Please try again.");
					}
				},
				cache: false,
				beforeSend: function() {
					$("#dashboarddiv-wide").empty();
				}
			});
		}
		else if ($("#PULSE_MODE").val() == "ADMIN"){
			XHRmain = $.ajax({
				url: "includes/useradmin.asp?DATE=" + $("#PARAMETER_DATE").val() + "&AGENT=" + agentValue + "&SUPERVISOR=" + supervisorValue + "&DEPARTMENT=" + departmentValue + "&JOB=" + jobValue + "&LOCATION=" + locationValue + "&HIRE=" + hireValue + "&NEWUSER=0",
				success: function(result){
					$("#dashboarddiv-wide").css("display","block").html(result);
					
					$("#PULSE_FORM").ajaxForm({
						beforeSerialize: function() { 
							$("#ADMINID_LIST").val(adminIdList.join());
						},
						clearForm: true,
						error: function() {
							$("#dashboarddiv-wide").css("display","block").html("Update Failed");
						},
						success: function() {
							$.when(chosenOptions("ADMIN")).then(function(){
								getAgentList();
							});
						}
					});
				},
				error: function(jqXHR, textStatus) {
					if(textStatus != "abort"){
						$("#dashboarddiv-wide").css("display","block").html("Failed to load list. Please try again.");
					}
				},
				cache: false,
				beforeSend: function() {
					$("#dashboarddiv-wide").empty();
				}
			});
		}
	}
	$(document).ready(function() {
		$(".chosen-select").chosen({
			width: "100%",
			placeholder_text_multiple: "Search",
			search_contains: true,
			display_selected_options: false,
			include_group_label_in_selected: true,
			group_search: true
		});
		$("#title-bar").on("chosen:hiding_dropdown", function(evt) {
			enterTimestamp = evt.timeStamp;
		});
		$("#title-bar").on("keyup",".chosen-container",function(evt) {
			if(evt.which === 13 || evt.which === 119) {
				var d = new Date();
				if(evt.which === 119 || d.getTime() - enterTimestamp >= 60){
					//waits 60 milliseconds after closing dropdown to listen for "Enter"
					getAgentList();
					$('.chosen-select').trigger('chosen:close');
				}
			}
		});
		$("#main-wrap").on("click",".searchAgent",function() {
			$("#search-bar").val("AGT_" + $(this).data("user")).trigger("chosen:updated");
			getAgentList();
		});
	});
/* Agent List - End */

/* Side Menu - Start */
	function staffingGraph(useWorkgroup){
		XHRmain.abort();
		XHRmain = $.ajax({
			url: "includes/graph.asp?DATE=" + $("#PARAMETER_DATE").val() + "&WORKGROUP=" + useWorkgroup,
			success: function(result){
				$("#dashboarddiv-wide").css("display","block").html(result);
				
				$("#CONTROL_FORM").ajaxForm({
					beforeSerialize: function() { 
						$("#CONTROLID_LIST").val(controlIdList.join());
					},
					clearForm: true,
					error: function() {
						$("#dashboarddiv-wide").css("display","block").html("Update Failed");
					},
					success: function() {
						staffingGraph(useWorkgroup);
					}
				});
			},
			error: function(jqXHR, textStatus) {
				if(textStatus != "abort"){
					$("#dashboarddiv-wide").css("display","block").html("Failed to load graph. Please try again.");
				}
			},
			cache: false,
			beforeSend: function() {
				$("#dashboarddiv-wide").empty();
			}
		});
	}
	function pendingList(useWorkgroup){
		XHRmain.abort();
		XHRmain = $.ajax({
			url: "includes/agentlist.asp?REQUEST=PENDING&DATE=" + $("#PARAMETER_DATE").val() + "&WORKGROUP=" + useWorkgroup,
			success: function(result){
				$("#dashboarddiv-wide").css("display","block").html(result);

				$("#PULSE_FORM").ajaxForm({
					beforeSerialize: function() { 
						$("#SCHEDULEID_LIST").val(scheduleIdList.join());
					},
					clearForm: true,
					error: function() {
						$("#dashboarddiv-wide").css("display","block").html("Update Failed");
					},
					success: function() {
						pendingList(useWorkgroup);
						currentDay();
					}
				});
			},
			complete: liveStateAgentList,
			error: function(jqXHR, textStatus) {
				if(textStatus != "abort"){
					$("#dashboarddiv-wide").css("display","block").html("Failed to load list. Please try again.");
				}
			},
			cache: false,
			beforeSend: function() {
				$("#dashboarddiv-wide").empty();
			}
		});
	}
	function dataChart(useWorkgroup){
		XHRmain.abort();
		XHRmain = $.ajax({
			url: "includes/datachart.asp?DATE=" + $("#PARAMETER_DATE").val() + "&WORKGROUP=" + useWorkgroup,
			success: function(result){
				$("#dashboarddiv-wide").css("display","block").html(result);
			},
			error: function(jqXHR, textStatus) {
				if(textStatus != "abort"){
					$("#dashboarddiv-wide").css("display","block").html("Failed to load chart. Please try again.");
				}
			},
			cache: false,
			beforeSend: function() {
				$("#dashboarddiv-wide").empty();
			}
		});
	}
	function errorList(useWorkgroup){
		XHRmain.abort();
		XHRmain = $.ajax({
			url: "includes/agentlist.asp?REQUEST=ERROR&DATE=" + $("#PARAMETER_DATE").val() + "&WORKGROUP=" + useWorkgroup,
			success: function(result){
				$("#dashboarddiv-wide").css("display","block").html(result);

				$("#PULSE_FORM").ajaxForm({
					beforeSerialize: function() { 
						$("#SCHEDULEID_LIST").val(scheduleIdList.join());
					},
					clearForm: true,
					error: function() {
						$("#dashboarddiv-wide").css("display","block").html("Update Failed");
					},
					success: function() {
						errorList(useWorkgroup);
						currentDay();
					}
				});
			},
			complete: liveStateAgentList,
			error: function(jqXHR, textStatus) {
				if(textStatus != "abort"){
					$("#dashboarddiv-wide").css("display","block").html("Failed to load list. Please try again.");
				}
			},
			cache: false,
			beforeSend: function() {
				$("#dashboarddiv-wide").empty();
			}
		});
	}
	function noteList(useWorkgroup){
		XHRmain.abort();
		XHRmain = $.ajax({
			url: "includes/notelist.asp?DATE=" + $("#PARAMETER_DATE").val() + "&WORKGROUP=" + useWorkgroup,
			success: function(result){
				$("#dashboarddiv-wide").css("display","block").html(result);
				
				$("#NOTE_FORM").ajaxForm({
					beforeSerialize: function() { 
						$("#NOTEID_LIST").val(noteIdList.join());
					},
					clearForm: true,
					error: function() {
						$("#dashboarddiv-wide").css("display","block").html("Update Failed");
					},
					success: function() {
						noteList(useWorkgroup);
					}
				});
			},
			error: function(jqXHR, textStatus) {
				if(textStatus != "abort"){
					$("#dashboarddiv-wide").css("display","block").html("Failed to load list. Please try again.");
				}
			},
			cache: false,
			beforeSend: function() {
				$("#dashboarddiv-wide").empty();
			}
		});
	}
	function tradeList(){
		XHRmain.abort();
		XHRmain = $.ajax({
			url: "includes/tradelist.asp?DATE=" + $("#PARAMETER_DATE").val(),
			success: function(result){
				$("#dashboarddiv-wide").css("display","block").html(result);

				$("#PULSE_FORM").ajaxForm({
					beforeSerialize: function() { 
						$("#SCHEDULEID_LIST").val(scheduleIdList.join());
					},
					clearForm: true,
					error: function() {
						$("#dashboarddiv-wide").css("display","block").html("Update Failed");
					},
					success: dailyStats
				});
			},
			error: function(jqXHR, textStatus) {
				if(textStatus != "abort"){
					$("#dashboarddiv-wide").css("display","block").html("Failed to load list. Please try again.");
				}
			},
			cache: false,
			beforeSend: function() {
				$("#dashboarddiv-wide").empty();
			}
		});
	}
	function newAdminUser(){
		XHRmain.abort();
		XHRmain = $.ajax({
			url: "includes/useradmin.asp?DATE=" + $("#PARAMETER_DATE").val() + "&AGENT=&SUPERVISOR=&DEPARTMENT=&JOB=&LOCATION=&HIRE=&NEWUSER=1",
			success: function(result){
				$("#dashboarddiv-wide").css("display","block").html(result);
				
				$("#PULSE_FORM").ajaxForm({
					beforeSerialize: function() { 
						$("#ADMINID_LIST").val(adminIdList.join());
					},
					clearForm: true,
					error: function() {
						$("#dashboarddiv-wide").css("display","block").html("Update Failed");
					},
					success: function() {
						$.when(chosenOptions("ADMIN")).then(function(){
							getAgentList();
						});
					}
				});
			},
			error: function(jqXHR, textStatus) {
				if(textStatus != "abort"){
					$("#dashboarddiv-wide").css("display","block").html("Failed to load list. Please try again.");
				}
			},
			cache: false,
			beforeSend: function() {
				$("#dashboarddiv-wide").empty();
			}
		});	
	}
	$(document).ready(function() {
		$("#sidebar-collapse").on("click",".graph-item",function() {
			var idArray = this.id.split("_");
			staffingGraph(idArray[0]);
		});
		$("#sidebar-collapse").on("click",".pending-item",function() {
			var idArray = this.id.split("_");
			pendingList(idArray[0]);
		});
		$("#sidebar-collapse").on("click",".data-item",function() {
			var idArray = this.id.split("_");
			dataChart(idArray[0]);
		});
		$("#sidebar-collapse").on("click",".error-item",function() {
			var idArray = this.id.split("_");
			errorList(idArray[0]);
		});
		$("#sidebar-collapse").on("click",".note-item",function() {
			var idArray = this.id.split("_");
			noteList(idArray[0]);
		});
		$("#sidebar-collapse").on("click","#TRADES_BUTTON",tradeList);
		$("#sidebar-collapse").on("click","#ADMIN_BUTTON",newAdminUser);
	});
/* Side Menu - End */

/* Live CMS Stats - Start */
	function liveStateAgentList(){
		if($("div[id^='AGENTPHONESTATE_']").length > 0){
			if($("input[id^='LIVECMSSTATE_']").length > 0){
				if($("#CMS_WARNING").length > 0){
					var cssColor = "red";
				}
				else{
					var cssColor = "inherit";
				}
				$("div[id^='AGENTPHONESTATE_']").each(function(){
					var idArray = this.id.split("_");
					if($("#LIVECMSSTATE_" + idArray[1]).length > 0){
						$("#AGENTPHONESTATE_" + idArray[1]).css("color", cssColor).html("(" + $("#LIVECMSSTATE_" + idArray[1]).val() + ")");
					}
					else{
						$("#AGENTPHONESTATE_" + idArray[1]).empty();
					}
				});
			}
			else{
				$("div[id^='AGENTPHONESTATE_']").empty();
			}
		}
	}
	function liveCMS(refreshBool){
		var USE_SELECTION = "";
		var DISPLAY_ARRAY = "";
		if($("#AGENT_TIME_SELECTION").val() === undefined){
			USE_SELECTION = "AVAIL";
			DISPLAY_ARRAY = "none,none,none,none";
		}
		else{
			USE_SELECTION = $("#AGENT_TIME_SELECTION").val();
			DISPLAY_ARRAY = $("#SHOWLIVESTATUS_SALES").val() + "," + $("#SHOWLIVESTATUS_SERVICE").val() + "," + $("#SHOWLIVESTATUS_SUPPORT").val() + "," + $("#SHOWLIVESTATUS_OD").val();
		}
		XHRliveCMS.abort();
		XHRliveCMS = $.ajax({
			url: "includes/livecms.asp?DATE=" + $("#PARAMETER_DATE").val() + "&TIME_SELECTION=" + USE_SELECTION + "&DISPLAY_ARRAY=" + DISPLAY_ARRAY,
			success: function(result){
				$("#dashboard-left").html(result);
			},
			error: function(){
				clearTimeout(liveCMSTimeout);
				liveCMSTimeout = setTimeout(liveCMS,30000);
				clearTimeout(liveRefreshCMSTimeout);
				liveRefreshCMSTimeout = null;
			},
			complete: function(){
				clearTimeout(liveCMSTimeout);
				liveCMSTimeout = setTimeout(liveCMS,1000*Math.max(Math.min(60-10*$("#PULSE_SECURITY").val(),30),10));
				liveStateAgentList();
				if(refreshBool == 1){
					liveCMSRefreshTimeout = setTimeout(function(){
						$("#CMSREFRESH").show();
						liveCMSRefreshTimeout = null;
					}, 1000*Math.max(Math.min(25-5*$("#PULSE_SECURITY").val(),15),5));
				}
				else if(liveCMSRefreshTimeout == null){
					$("#CMSREFRESH").show();
				}
			},
			cache: false
		});	
	}
	$(document).ready(function() {
		$("#main-wrap").on("change","#AGENT_TIME_SELECTION",liveCMS);
		$("#main-wrap").on("click","#CMSREFRESH",function() {
		   liveCMS(1);
		});
		$("#main-wrap").on("click","[id^=SHOWLIVEBUTTON_]",function() {
			var $statusInput = $("#" + this.id.replace("SHOWLIVEBUTTON","SHOWLIVESTATUS"));
			var $liveBody = $("#" + this.id.replace("SHOWLIVEBUTTON","SHOWLIVEBODY"));
			
			if($statusInput.val() == "none"){
				$liveBody.css("display","table-row-group");
				$statusInput.val("table-row-group");
				$(this).removeClass("fa-plus").addClass("fa-minus");
			}
			else{
				$liveBody.css("display","none");
				$statusInput.val("none");
				$(this).removeClass("fa-minus").addClass("fa-plus");			
			}
		});
		$("#main-wrap").on("focus","#AGENT_TIME_SELECTION",function() {
			clearTimeout(liveCMSTimeout);
		});
		$("#main-wrap").on("blur","#AGENT_TIME_SELECTION",function() {
			liveCMSTimeout = setTimeout(liveCMS,1000*Math.max(Math.min(60-10*$("#PULSE_SECURITY").val(),30),10));
		});
	});
/* Live CMS Stats - End */

/* Schedule Stats/Bar Chart - Start */	
	function scheduleStats(){
		XHRscheduleStats.abort();
		XHRscheduleStats = $.ajax({
			url: "includes/schedulestats.asp?DATE=" + $("#PARAMETER_DATE").val(),
			success: function(result){
				$("#schedule-stats-div").html(result);
			},
			error: function(result){
				clearTimeout(scheduleStatsTimeout);
				scheduleStatsTimeout = setTimeout(scheduleStats,30000);
			},
			complete: function(result){
				clearTimeout(scheduleStatsTimeout);
				scheduleStatsTimeout = setTimeout(scheduleStats,1000*Math.max(Math.min(10800-1800*$("#PULSE_SECURITY").val(),3600),1800));
			},
			cache: false
		});	
	}
	function barChartClick(el){
		var useTarget = el["targetID"];
		if (useTarget.substr(0,4) == "bar#"){
			var chosenArray = [];
			switch(useTarget.substr(4,1)){
				case "0":
					chosenArray.push("DEPT_RES");
					break;
				case "1":
					chosenArray.push("DEPT_SPT");
					break;
			}
			switch(useTarget.substr(6,1)){
				case "0":
					chosenArray.push("SCH_VACA", "SCH_ROUT", "SCH_RCHG");
					break;
				case "1":
					chosenArray.push("SCH_SRPT", "SCH_SRUN");
					break;
				case "2":
					chosenArray.push("SCH_SKPP", "SCH_SKPT", "SCH_SKUN", "SCH_FMPP", "SCH_FMPT", "SCH_FMUN", "SCH_FMHL", "SCH_WXUN", "SCH_WXPT");
					break;
				case "3":
					chosenArray.push("SCH_CDPT", "SCH_CDUN");
					break;
				case "4":
					chosenArray.push("SCH_APPT", "SCH_APUN", "SCH_OTPT", "SCH_OTUN");
					break;
				case "5":
					chosenArray.push("SCH_ADDT");
					break;
				case "6":
					chosenArray.push("SCH_SLIP");
					break;
			}
			$("#search-bar").val(chosenArray).trigger("chosen:updated");
			getAgentList();
		}
	}
/* Schedule Stats/Bar Chart - End */	

/* Current Day - Start */	
	function currentDay(){
		XHRcurrentDay.abort();
		XHRcurrentDay = $.ajax({
			url: "includes/currentday.asp?DATE=" + $("#PARAMETER_DATE").val(),
			success: function(result){
				$("#current-day-div").html(result);
			},
			error: function(result){
				clearTimeout(currentDayTimeout);
				currentDayTimeout = setTimeout(currentDay,30000);
			},
			complete: function(result){
				clearTimeout(currentDayTimeout);
				currentDayTimeout = setTimeout(currentDay,1000*Math.max(Math.min(7560-1440*$("#PULSE_SECURITY").val(),1800),360));
			},
			cache: false
		});	
	}
/* Current Day - End */	

/* Staffing/Notes/Trades - Start */
	function editSchedule(userId, useDateString, requestType, refreshBool){
		if(refreshBool == 1){
			$("div[id^='SLIDER_'][data-user='" + userId + "'][data-parent-date='" + useDateString + "']").each(function(){
				var sliderArray = this.id.split("_");
				removeScheduleList(sliderArray[1]);
			});
		}
		return $.ajax({
			url: "includes/editschedule.asp?REQUEST=" + requestType + "&DATE=" + moment(useDateString,"MMDDYYYY").format("L") + "&AGENT=" + userId + "&PULSEDATE=" + $("#PARAMETER_DATE").val(),
			success: function(result){
				$("#EDITDIV_" + userId + "_" + useDateString).html(result);
			},
			cache: false
		});	
	}
	function newScheduleEntry(userId, useDateString, sliderValues) {
		var newlineId = --$("#NEWLINE_ID").get(0).value;
		var cloneRow = $("#EDITTABLE_" + userId + "_" + useDateString + " tr").eq(-3);
		var NotesRow = $("#EDITTABLE_" + userId + "_" + useDateString + " tr").eq(-2);
		$(cloneRow[0].outerHTML.replace(/_0/g, "_" + newlineId)).show().appendTo("#SCITBODY_" + userId + "_" + useDateString);
		$(NotesRow[0].outerHTML.replace(/_0/g, "_" + newlineId)).appendTo("#SCITBODY_" + userId + "_" + useDateString);
		
		newSliderValues(userId, useDateString, newlineId, 0, sliderValues);
		initializeSlider(newlineId);
		addScheduleList(newlineId);
		checkOverlaps(userId, useDateString);
	}
	function newSliderValues(userId, useDateString, sliderId, updateImmediatelyFlag, sliderValues){
		var sliderElements = $("#SCITBODY_" + userId + "_" + useDateString + " div[id^='SLIDER_']:not([id='SLIDER_0']):not([id='SLIDER_" + sliderId + "'])");
		var sliderStep = parseInt($("#SLIDER_" + sliderId).data("slider-step"));
		
		var useDept = $("#SCITBODY_" + userId + "_" + useDateString).data("department");

		if(sliderValues === undefined){
			var useStart = -1;
			if(sliderElements.length > 0){
				for (var i = 0; i < sliderElements.length; i++){
					var idArray = sliderElements[i].id.split("_");
					if($("#SCISTATUS_" + idArray[1]).val() == "APP" && $("#SCITYPE_" + idArray[1]).val() != "HOLR" && $("#SCITYPE_" + idArray[1]).val() != "HOLU"){
						useStart = Math.max($("#" + sliderElements[i].id).slider("values",1), useStart);

						if (i == sliderElements.length - 1){
							useDept = $("#SCIUSRTYPE_" + idArray[1]).val();
						}
					}
				}
			}
			if(useStart == -1){
				if(moment(useDateString,"MMDDYYYY").format("L") == parameterDate){
					var d = new Date();
					useStart = 60 * d.getHours() - (-1)*(sliderStep*Math.round(d.getMinutes()/sliderStep));
				}		
				else{
					useStart = parseInt($("#SCITBODY_" + userId + "_" + useDateString).data("date-min"));
				}
			}
			var useEnd = Math.min(useStart - (-1)*sliderStep, parseInt($("#SCITBODY_" + userId + "_" + useDateString).data("date-max")));
			
			var HH = ("0" + Math.floor(useStart / 60)).substr(-2);
			var MM = ("0" + Math.floor(useStart % 60)).substr(-2);
			$("#STARTTIME_" + sliderId).html(HH + ":" + MM);
			$("#SCISTART_" + sliderId).val(HH + ":" + MM);
			HH = ("0" + Math.floor(useEnd / 60)).substr(-2);
			MM = ("0" + Math.floor(useEnd % 60)).substr(-2);
			$("#ENDTIME_" + sliderId).html(HH + ":" + MM);
			$("#SCIEND_" + sliderId).val(HH + ":" + MM);		
			$("#SCIUSRTYPE_" + sliderId).val(useDept);
		}
		else{
			var useStart = sliderValues.start;
			var useEnd = sliderValues.end;
			var useType = sliderValues.type;
			var useNotes = sliderValues.notes;

			HH = ("0" + Math.floor(useStart / 60)).substr(-2);
			MM = ("0" + Math.floor(useStart % 60)).substr(-2);
			$("#STARTTIME_" + sliderId).html(HH + ":" + MM);
			$("#SCISTART_" + sliderId).val(HH + ":" + MM);
			HH = ("0" + Math.floor(useEnd / 60)).substr(-2);
			MM = ("0" + Math.floor(useEnd % 60)).substr(-2);
			$("#ENDTIME_" + sliderId).html(HH + ":" + MM);
			$("#SCIEND_" + sliderId).val(HH + ":" + MM);
			$("#SCITYPE_" + sliderId).val(useType);	
			$("#SCIUSRTYPE_" + sliderId).val(useDept);			
			$("#SCINOTES_" + sliderId).val(useNotes);
		}
		
		if(updateImmediatelyFlag == 1){
			if($("#SCITBODY_" + userId + "_" + useDateString).hasClass("altdate-entry-color")){
				$("#SCIDATE_" + sliderId).addClass("altdate-entry-color");
				$("#SCITYPE_" + sliderId).addClass("altdate-entry-color");
				$("#SCISTATUS_" + sliderId).addClass("altdate-entry-color");
				$("#SCIUSRTYPE_" + sliderId).addClass("altdate-entry-color");
			}
			else{
				$("#SCIDATE_" + sliderId).removeClass("altdate-entry-color");
				$("#SCITYPE_" + sliderId).removeClass("altdate-entry-color");
				$("#SCISTATUS_" + sliderId).removeClass("altdate-entry-color");
				$("#SCIUSRTYPE_" + sliderId).removeClass("altdate-entry-color");			
			}
			$("#SLIDER_" + sliderId).slider("option","values",[useStart,useEnd]);
			$("#SLIDER_" + sliderId).slider("option","min",parseInt($("#SCITBODY_" + userId + "_" + useDateString).data("date-min")));
			$("#SLIDER_" + sliderId).slider("option","max",parseInt($("#SCITBODY_" + userId + "_" + useDateString).data("date-max")));
		}
	}
	function initializeSlider(sliderId){
		var startArray = $("#SCISTART_" + sliderId).val().split(":");
		var startValue = 60*startArray[0] - (-1)*startArray[1];
		var endArray = $("#SCIEND_" + sliderId).val().split(":");
		var endValue = 60*endArray[0] - (-1)*endArray[1];
		var sliderStep = parseInt($("#SLIDER_" + sliderId).data("slider-step"));
		var sliderMin = $("#SLIDER_" + sliderId).data("slider-min");
		var sliderMax = $("#SLIDER_" + sliderId).data("slider-max");
		var scheduleClass = $("#SCISTATUS_" + sliderId).val() == "APP" ? $("#SCITYPE_" + sliderId).find("option:selected").data("schedule-class") : "PEND";
		var sliderDisabled = $("#SLIDER_" + sliderId).data("slider-disabled");
		
		$("#SLIDER_" + sliderId).slider({
			range: true,
			min: sliderMin,
			max: sliderMax,
			step: sliderStep,
			values: [startValue, endValue],
			disabled: sliderDisabled,
			start: function(event,ui){
				$(this).slider("option","step",sliderStep);
			},
			stop: function(event, ui){
				checkOverlaps($("#SLIDER_" + sliderId).data("user"), $("#SLIDER_" + sliderId).data("parent-date"));
				addScheduleList(sliderId);
			},
			slide: function(event, ui){
				if (ui.handleIndex == 0){
					var HH = ("0" + Math.floor(ui.values[0] / 60)).substr(-2);
					var MM = ("0" + Math.floor(ui.values[0] % 60)).substr(-2);
					$("#STARTTIME_" + sliderId).html(HH + ":" + MM);
					$("#SCISTART_" + sliderId).val(HH + ":" + MM);
				}
				else{
					var HH = ("0" + Math.floor(ui.values[1] / 60)).substr(-2);
					var MM = ("0" + Math.floor(ui.values[1] % 60)).substr(-2);
					$("#ENDTIME_" + sliderId).html(HH + ":" + MM);
					$("#SCIEND_" + sliderId).val(HH + ":" + MM);
				}
				
			}
		});
		$("#SLIDER_" + sliderId).find(".ui-slider-range").addClass(scheduleClass);
	}
	function emptySchedule(userId, useDateString){
		$("div[id^='SLIDER_'][data-user='" + userId + "'][data-parent-date='" + useDateString + "']").each(function(){
			var sliderArray = this.id.split("_");
			removeScheduleList(sliderArray[1]);
		});
		$("#EDITDIV_" + userId + "_" + useDateString).empty();
		removeOverlapList(userId, useDateString);
	}
	function flexSchedule(userId, useDateString, flexBool){
		var flexValue = $("#FLEXLENGTH_" + userId).val();
		if (flexValue != "0"){
			var tbodyCount = 0
			var tbodyElements = $("tbody[id^='SCITBODY_'][data-user='" + userId + "'][data-parent-date='" + useDateString + "']").each(function(){
				if($(this).find("div[id^='SLIDER_']:not([id='SLIDER_0'])").length > 0){
					tbodyCount++;
				}
			});
			if (tbodyCount == 1){
				var sliderElements = $("#SCITBODY_" + userId + "_" + useDateString).find("div[id^='SLIDER_']:not([id='SLIDER_0'])")
				var firstSliderId;
				var lastSliderId;
				for (var i = 0; i < sliderElements.length; i++){
					var currentId = sliderElements[i].id.split("_");
					if($("#SCISTATUS_" + currentId[1]).val() == "APP" && $("#SCITYPE_" + currentId[1]).val() != "HOLR" && $("#SCITYPE_" + currentId[1]).val() != "HOLU"){
						firstSliderId = currentId[1];
						break;
					}
				}
				for (var i = sliderElements.length-1; i > -1; i--){
					var currentId = sliderElements[i].id.split("_");
					if($("#SCISTATUS_" + currentId[1]).val() == "APP" && $("#SCITYPE_" + currentId[1]).val() != "HOLR" && $("#SCITYPE_" + currentId[1]).val() != "HOLU"){
						lastSliderId = currentId[1];
						break;
					}
				}
				if(flexValue != $("#SCISTART_" + firstSliderId).val()){	
					var firstStartArray = $("#SCISTART_" + firstSliderId).val().split(":");
					var firstStartValue = 60*firstStartArray[0] - (-1)*firstStartArray[1];
					var firstEndArray = $("#SCIEND_" + firstSliderId).val().split(":");
					var firstEndValue = 60*firstEndArray[0] - (-1)*firstEndArray[1];
				
					var lastStartArray = $("#SCISTART_" + lastSliderId).val().split(":");
					var lastStartValue = 60*lastStartArray[0] - (-1)*lastStartArray[1];
					var lastEndArray = $("#SCIEND_" + lastSliderId).val().split(":");
					var lastEndValue = 60*lastEndArray[0] - (-1)*lastEndArray[1];
					
					var flexArray = flexValue.split(":");
					var flexLength = 60*flexArray[0] - (-1)*flexArray[1] - firstStartValue;
					if(firstStartValue - (-1)*flexLength >= 0 && firstStartValue - (-1)*flexLength <= firstEndValue && lastEndValue - (-1)*flexLength >= lastStartValue && lastEndValue - (-1)*flexLength <= 1440){
						$("#FLEX_" + userId + "_" + useDateString).show();
						if(flexBool == 1){
							var HH;
							var MM;
							if(firstStartValue - (-1)*flexLength == firstEndValue){
								$("#SCISTATUS_" + firstSliderId).val("DEL");
								$("#SLIDER_" + firstSliderId).find(".ui-slider-range").removeClass("PHONE LUNCH TRAIN VACA SRED").addClass("PEND");
							}
							else{
								$("#SLIDER_" + firstSliderId).slider("values", 0, firstStartValue - (-1)*flexLength);
								HH = ("0" + Math.floor((firstStartValue - (-1)*flexLength) / 60)).substr(-2);
								MM = ("0" + Math.floor((firstStartValue - (-1)*flexLength) % 60)).substr(-2);
								$("#STARTTIME_" + firstSliderId).html(HH + ":" + MM);
								$("#SCISTART_" + firstSliderId).val(HH + ":" + MM);
							}
							if(lastEndValue - (-1)*flexLength == lastStartValue){
								$("#SCISTATUS_" + lastSliderId).val("DEL");
								$("#SLIDER_" + lastSliderId).find(".ui-slider-range").removeClass("PHONE LUNCH TRAIN VACA SRED").addClass("PEND");
							}
							else{
								$("#SLIDER_" + lastSliderId).slider("values", 1, lastEndValue - (-1)*flexLength);
								HH = ("0" + Math.floor((lastEndValue - (-1)*flexLength) / 60)).substr(-2);
								MM = ("0" + Math.floor((lastEndValue - (-1)*flexLength) % 60)).substr(-2);
								$("#ENDTIME_" + lastSliderId).html(HH + ":" + MM);
								$("#SCIEND_" + lastSliderId).val(HH + ":" + MM);
							}							
							addScheduleList(firstSliderId);
							addScheduleList(lastSliderId);
						}
					}
				}
			}
		}
	}
	function checkOverlaps(userId, parentDate){
		var overlapBool = 0;
		var tbodyElements = $("tbody[id^='SCITBODY_'][data-user='" + userId + "'][data-parent-date='" + parentDate + "']");
		for (var k = 0; k < tbodyElements.length; k++){
			if(overlapBool == 1){
				break;
			}
			var sliderElements = $("#" + tbodyElements[k].id).find("div[id^='SLIDER_']:not([id='SLIDER_0'])")
			for (var i = 0; i < sliderElements.length; i++){
				if(overlapBool == 1){
					break;
				}
				var currentId = sliderElements[i].id.split("_");
				if($("#SCISTATUS_" + currentId[1]).val() != "APP" || $("#SCITYPE_" + currentId[1]).val() == "HOLR" || $("#SCITYPE_" + currentId[1]).val() == "HOLU"){
					continue;
				}
				for (var j = i+1; j < sliderElements.length; j++){
					var loopId = sliderElements[j].id.split("_");
					if($("#SCISTATUS_" + loopId[1]).val() != "APP" || $("#SCITYPE_" + loopId[1]).val() == "HOLR" || $("#SCITYPE_" + loopId[1]).val() == "HOLU"){
						continue;
					}
					var currentStart = $("#" + sliderElements[i].id).slider("values",0);
					var currentEnd = $("#" + sliderElements[i].id).slider("values",1);
					var loopStart = $("#" + sliderElements[j].id).slider("values",0);
					var loopEnd = $("#" + sliderElements[j].id).slider("values",1);

					if (currentStart < loopEnd && currentEnd > loopStart){
						overlapBool = 1;
						break;
					}
				}
			}
		}
		if(overlapBool == 1){
			addOverlapList(userId, parentDate);

		}
		else{
			removeOverlapList(userId, parentDate);
		}
	}
	function addOverlapList(overlapUser, overlapDate){
		if(overlapIdList.indexOf(overlapUser + "_" + overlapDate) == -1) {
			overlapIdList.push(overlapUser + "_" + overlapDate);
		}
		$("#AGENTROW_" + overlapUser + "_" + overlapDate).addClass("error-color-background");
		$("#EDITROW_" + overlapUser + "_" + overlapDate).addClass("error-color-background");
		$("#EDITTABLE_" + overlapUser + "_" + overlapDate).find("tr[class~=new-entry-color]").addClass("new-entry-color-error");
		$("#EDITTABLE_" + overlapUser + "_" + overlapDate).find("select").addClass("error-color-background");
		$("tbody").find("[id^='SCITBODY_'][data-user='" + overlapUser + "'][data-parent-date='" + overlapDate + "']").addClass("error-color-background");
		$("#PULSE_FORM_DIV").addClass("error-shadow");
		$("#PULSE_SUBMIT").hide();
		$("#OVERLAP_MESSAGE").show();
	}
	function removeOverlapList(overlapUser, overlapDate){
		if(overlapIdList.indexOf(overlapUser + "_" + overlapDate) != -1) {
			overlapIdList.splice(overlapIdList.indexOf(overlapUser + "_" + overlapDate), 1);
		}
		$("#AGENTROW_" + overlapUser + "_" + overlapDate).removeClass("error-color-background");
		$("#EDITROW_" + overlapUser + "_" + overlapDate).removeClass("error-color-background");
		$("#EDITTABLE_" + overlapUser + "_" + overlapDate).find("tr[class~=new-entry-color]").removeClass("new-entry-color-error");
		$("#EDITTABLE_" + overlapUser + "_" + overlapDate).find("select").removeClass("error-color-background");
		$("tbody").find("[id^='SCITBODY_'][data-user='" + overlapUser + "'][data-parent-date='" + overlapDate + "']").removeClass("error-color-background");
		if(overlapIdList.length == 0){
			$("#PULSE_FORM_DIV").removeClass("error-shadow");
			$("#PULSE_SUBMIT").show();
			$("#OVERLAP_MESSAGE").hide();
		}
	}
	function addScheduleList(sliderId){
		if(scheduleIdList.indexOf(parseInt(sliderId)) == -1) {
			scheduleIdList.push(parseInt(sliderId));
		}
	}
	function removeScheduleList(sliderId){
		if(scheduleIdList.indexOf(parseInt(sliderId)) != -1) {
			scheduleIdList.splice(scheduleIdList.indexOf(parseInt(sliderId)), 1);
		}
	}
	function addNoteList(noteId){
		if(noteIdList.indexOf(parseInt(noteId)) == -1) {
			noteIdList.push(parseInt(noteId));
		}
	}
	$(document).ready(function() {
		$("#main-wrap").on("click","[id^=EDITBUTTON_]",function() {
			var idArray = this.id.split("_");
			if ($(this).hasClass("fa-angle-down")){
				$(this).removeClass("fa-angle-down").addClass("fa-angle-up");
			}
			else{
				$(this).removeClass("fa-angle-up").addClass("fa-angle-down");
			}
			$("#EDITROW_" + idArray[1] + "_" + idArray[2]).toggle();
			if($("#EDITDIV_" + idArray[1] + "_" + idArray[2]).html().length == 0){
				$.when(editSchedule(idArray[1], idArray[2], "NORMAL", 0)).then(function(){
					flexSchedule(idArray[1], idArray[2], 0);
					checkOverlaps(idArray[1], idArray[2]);
				});
			}
		});
		$("#main-wrap").on("change","select[id^=SCIDATE_]",function() {
			var idArray = this.id.split("_");
			var userId = $("#SLIDER_" + idArray[1]).data("user");

			var cutSchedule = $("#SCHEDULEROW_" + idArray[1]).detach();
			var cutNotes = $("#NOTESROW_" + idArray[1]).detach();
			
			$(cutSchedule).appendTo("tbody[id=SCITBODY_" + userId + "_" + $(this).val().replace(/\//g,""));
			$(cutNotes).appendTo("tbody[id=SCITBODY_" + userId + "_" + $(this).val().replace(/\//g,""));
			
			newSliderValues(userId,$(this).val().replace(/\//g,""), idArray[1], 1);
			checkOverlaps($("#SLIDER_" + idArray[1]).data("user"), $("#SLIDER_" + idArray[1]).data("parent-date"));
			addScheduleList(idArray[1]);
		});
		$("#main-wrap").on("change","select[id^=SCITYPE_]",function() {
			var idArray = this.id.split("_");
			var newClass = $("#SCISTATUS_" + idArray[1]).val() == "APP" ? $(this).find("option:selected").data("schedule-class") : "PEND"
			var classString = "PEND PHONE LUNCH TRAIN VACA SRED".replace(newClass,"");
			$("#SLIDER_" + idArray[1]).find(".ui-slider-range").removeClass(classString).addClass(newClass);
			checkOverlaps($("#SLIDER_" + idArray[1]).data("user"), $("#SLIDER_" + idArray[1]).data("parent-date"));
			addScheduleList(idArray[1]);
		});
		$("#main-wrap").on("change","select[id^=SCISTATUS_]",function() {
			var idArray = this.id.split("_");
			var newClass = $(this).val() == "APP" ? $("#SCITYPE_" + idArray[1]).find("option:selected").data("schedule-class") : "PEND"
			var classString = "PEND PHONE LUNCH TRAIN VACA SRED".replace(newClass,"");
			$("#SLIDER_" + idArray[1]).find(".ui-slider-range").removeClass(classString).addClass(newClass);
			checkOverlaps($("#SLIDER_" + idArray[1]).data("user"), $("#SLIDER_" + idArray[1]).data("parent-date"));
			addScheduleList(idArray[1]);
		});
		$("#main-wrap").on("change","textarea[id^=SCINOTES_]",function() {
			var idArray = this.id.split("_");
			addScheduleList(idArray[1]);
		});
		$("#main-wrap").on("click","[id^=STARTARROW_],[id^=ENDARROW_]",function() {
			var idArray = this.id.split("_");
			var sliderDisabled = $("#SLIDER_" + idArray[2]).slider("option","disabled");
			if (sliderDisabled == false){
				if (idArray[0] == "STARTARROW"){
					var handleIndex = 0;
				}
				else{
					var handleIndex = 1;
				}
				var sliderValue = $("#SLIDER_" + idArray[2]).slider("values",handleIndex);
				var intervalLength = parseInt($("#SLIDER_" + idArray[2]).data("slider-interval"));
				$("#SLIDER_" + idArray[2]).slider("option","step",intervalLength);

				if(idArray[1] == "LEFT"){
					if (idArray[0] == "STARTARROW"){
						var sliderMin = $("#SLIDER_" + idArray[2]).slider("option","min");
					}
					else{
						var sliderMin = $("#SLIDER_" + idArray[2]).slider("values",0);
					}
					$("#SLIDER_" + idArray[2]).slider("values",handleIndex,Math.max(sliderValue-intervalLength,sliderMin));
					var HH = ("0" + Math.floor(Math.max(sliderValue-intervalLength,sliderMin) / 60)).substr(-2);
					var MM = ("0" + Math.floor(Math.max(sliderValue-intervalLength,sliderMin) % 60)).substr(-2);
				}
				else{
					if (idArray[0] == "ENDARROW"){
						var sliderMax = $("#SLIDER_" + idArray[2]).slider("option","max");
					}
					else{
						var sliderMax = $("#SLIDER_" + idArray[2]).slider("values",1);
					}
					$("#SLIDER_" + idArray[2]).slider("values",handleIndex,Math.min(sliderValue-(-1)*intervalLength,sliderMax));
					var HH = ("0" + Math.floor(Math.min(sliderValue-(-1)*intervalLength,sliderMax) / 60)).substr(-2);
					var MM = ("0" + Math.floor(Math.min(sliderValue-(-1)*intervalLength,sliderMax) % 60)).substr(-2);
				}
				if (idArray[0] == "STARTARROW"){
					$("#STARTTIME_" + idArray[2]).html(HH + ":" + MM);
					$("#SCISTART_" + idArray[2]).val(HH + ":" + MM);
				}
				else{
					$("#ENDTIME_" + idArray[2]).html(HH + ":" + MM);
					$("#SCIEND_" + idArray[2]).val(HH + ":" + MM);
				}
				checkOverlaps($("#SLIDER_" + idArray[2]).data("user"), $("#SLIDER_" + idArray[2]).data("parent-date"));
				addScheduleList(idArray[2]);
			}
		});
		$("#main-wrap").on("click","[id^=NOTESBUTTON_]",function() {
			var idArray = this.id.split("_");
			$("#NOTESROW_" + idArray[1]).toggle();
		});
		$("#main-wrap").on("click","[id^=NEWENTRY_]",function() {
			var idArray = this.id.split("_");
			newScheduleEntry(idArray[1], idArray[2]);
		});
		$("#main-wrap").on("click","[id^=REFRESH_]",function() {
			var idArray = this.id.split("_");
			editSchedule(idArray[1], idArray[2], $(this).data("request"), 1).then(function(){
				flexSchedule(idArray[1], idArray[2], 0);
				checkOverlaps(idArray[1], idArray[2]);
			});
		});
		$("#main-wrap").on("click","[id^=FLEX_]",function() {
			var idArray = this.id.split("_");
			$.when(editSchedule(idArray[1], idArray[2], $(this).data("request"), 1)).then(function(){
				flexSchedule(idArray[1], idArray[2], 1);
				checkOverlaps(idArray[1], idArray[2]);
			});
		});
		$("#main-wrap").on("click","[id^=HISTORY_]",function() {
			var idArray = this.id.split("_");
			var bPopup = $("#popupdiv").bPopup({
				speed: 650,
				transition: "slideDown",
				transitionClose: "slideUp",
				content:"ajax",
				contentContainer:"#popupdiv",
				loadUrl: "includes/schedulehistory.asp?DATE=" + moment(idArray[2],"MMDDYYYY").format("L") + "&AGENT=" + idArray[1] + "&JUNK="+ new Date().getTime()
			});
		});
		$("#main-wrap").on("click","[id^=ACKBUTTON_]",function() {
			var $errorInput = $("#" + this.id.replace("ACKBUTTON","ACKERROR"));
			
			if ($errorInput.prop("checked") == true){
				$errorInput.prop("checked", false);
				$(this).removeClass("new-entry-color");
			}
			else{
				$errorInput.prop("checked", true);
				$(this).addClass("new-entry-color");			
			}
		});
		$("#main-wrap").on("click",".searchNotes",function() {
			var bPopup = $("#popupdiv").bPopup({
				speed: 650,
				transition: "slideDown",
				transitionClose: "slideUp",
				content:"ajax",
				contentContainer:"#popupdiv",
				loadUrl: "includes/notelist.asp?DATE=" + $("#PARAMETER_DATE").val() + "&AGENT=" + $(this).data("user") + "&ERROR="+ $(this).data("error") + "&JUNK="+ new Date().getTime(),
				loadCallback: function(){

					$("#NOTE_FORM").ajaxForm({
						beforeSerialize: function() { 
							$("#NOTEID_LIST").val(noteIdList.join());
						},
						clearForm: true,
						error: function() {
							$("#popupdiv").html("Update Failed");
						},
						success: function() {
							bPopup.close();
						}
					});
				}
			});
		});	
		$("#main-wrap, #popupdiv").on("click","[id^=NOTEDELBUTTON_]",function() {
			var $errorInput = $("#" + this.id.replace("NOTEDELBUTTON","NOTEDELETE"));
			
			if ($errorInput.prop("checked") == true){
				$errorInput.prop("checked", false);
				$(this).removeClass("new-entry-color");
			}
			else{
				$errorInput.prop("checked", true);
				$(this).addClass("new-entry-color");			
			}
		});
		$("#main-wrap, #popupdiv").on("click","#NEWNOTE",function() {
			var newlineId = --$("#NEWLINE_ID").get(0).value;
			var idArray = this.id.split("_");
			var cloneRow = $("#NOTES_TABLE tr").eq(-3);
			
			$(cloneRow[0].outerHTML.replace(/_0/g, "_" + newlineId)).show().insertBefore(cloneRow);
			
			addNoteList(newlineId);
		});
		$("#main-wrap, #popupdiv").on("change","textarea[id^=NOTETEXT_]",function() {
			var idArray = this.id.split("_");
			addNoteList(idArray[1]);
		});
		$("#main-wrap").on("click",".trade-button",function() {
			var idArray = this.id.split("_");
			var $tradeInput = $("#" + this.id.replace("BUTTON","CHECK"));
			
			if ($tradeInput.prop("checked") == true){
				$tradeInput.prop("checked", false);
				$(this).removeClass("new-entry-color");
			}
			else{
				$tradeInput.prop("checked", true);
				$(this).addClass("new-entry-color");
				$tradeInput.siblings("input:checkbox").prop("checked", false);
				$(this).siblings(".trade-button").removeClass("new-entry-color");
			}
			if (idArray[0] == "TRADECOMBUTTON" && $tradeInput.prop("checked") == true){
				$("#TRADEDENIAL_" + idArray[1]).hide();
				$("#TRADETEXT_" + idArray[1]).val("");
				$("tr[id^='EDITROW_'][data-trade='" + idArray[1] + "']").show();
				$("tr[id^='EDITROW_'][data-trade='" + idArray[1] + "']").each(function(){
					var editArray = this.id.split("_");
					function SwapSchedules(){
						$("#TRADETABLE_" + editArray[1] + "_" + editArray[2] + " tr").each(function(){
							var sciID = $(this).data("sciid");
							if (sciID !== undefined){
								var sciStart = $(this).data("scistart");
								var sciEnd = $(this).data("sciend");
								var sciType = $(this).data("scitype");
								var sciNotes = $(this).data("scinotes");
								
								var startArray = $(this).data("scistart").split(":");
								var startValue = 60*startArray[0] - (-1)*startArray[1];
								
								var endArray = $(this).data("sciend").split(":");
								var endValue = 60*endArray[0] - (-1)*endArray[1];
								
								if(parseInt(sciID) != -1){
									var sciStatus = $(this).data("scistatus");
									
									$("#SCISTART_" + sciID).val(sciStart);
									$("#STARTTIME_" + sciID).html(sciStart);
									$("#SCIEND_" + sciID).val(sciEnd);
									$("#ENDTIME_" + sciID).html(sciEnd);
									$("#SCITYPE_" + sciID).val(sciType);
									$("#SCISTATUS_" + sciID).val(sciStatus);
									
									var newClass = sciStatus == "APP" ? $("#SCITYPE_" + sciID).find("option:selected").data("schedule-class") : "PEND"
									var classString = "PEND PHONE LUNCH TRAIN VACA SRED".replace(newClass,"");
									$("#SLIDER" + sciID).slider("option","values",[startValue, endValue]);
									$("#SLIDER_" + sciID).find(".ui-slider-range").removeClass(classString).addClass(newClass);
									
									$("#SCINOTES_" + sciID).val(sciNotes);
									addScheduleList(sciID);
								}
								else{
									var sliderValues = {
										start: startValue,
										end: endValue,
										type: sciType,
										notes: sciNotes,
									};
									newScheduleEntry(editArray[1], editArray[2], sliderValues);
								}
							}
						});
						checkOverlaps(editArray[1], editArray[2]);
					}
					$.when(editSchedule(editArray[1], editArray[2], "TRADE", 0)).then(SwapSchedules);
				});
				
			}
			else{
				if (idArray[0] == "TRADEDNYBUTTON" && $tradeInput.prop("checked") == true){
					$("#TRADEDENIAL_" + idArray[1]).show();
				}
				else{
					$("#TRADEDENIAL_" + idArray[1]).hide();
					$("#TRADETEXT_" + idArray[1]).val("");
				}
				$("tr[id^='EDITROW_'][data-trade='" + idArray[1] + "']").hide();
				$("tr[id^='EDITROW_'][data-trade='" + idArray[1] + "']").each(function(){
					var editArray = this.id.split("_");
					emptySchedule(editArray[1], editArray[2]);
				});
			}
		});
	});

/* Staffing/Notes/Trades - End */

/* Circle Stats - Start */
	function dailyStats(){
		XHRmain.abort();
		XHRmain = $.ajax({
			url: "includes/dailystats.asp?DATE=" + $("#PARAMETER_DATE").val(),
			success: function(result){
				$("#dashboarddiv-wide").css("display","flex").html(result);

				$("#CIRCLE_FORM").ajaxForm({
					beforeSerialize: function() { 
						$("#CIRCLEID_LIST").val(circleIdList.join());
					},
					clearForm: true,
					error: function() {
						$("#dashboarddiv-wide").css("display","block").html("Update Failed");
					},
					success: function() {
						dailyStats();
					}
				});
			},
			error: function(jqXHR, textStatus) {
				if(textStatus != "abort"){
					$("#dashboarddiv-wide").css("display","block").html("Failed to load stats. Please try again.");
				}
			},
			cache: false,
			beforeSend: function() {
				$("#dashboarddiv-wide").empty();
			}
		});
	}
	function editCircleStat(statType, refreshBool){
		if(refreshBool == 1){
			$("div[id^='CIRCLESLIDER_'][data-request-type='" + statType + "']").each(function(){
				var sliderArray = this.id.split("_");
				removeCircleList(sliderArray[1]);
			});
		}
		return $.ajax({
			url: "includes/editcirclestat.asp?TYPE=" + statType + "&DATE=" + $("#PARAMETER_DATE").val(),
			success: function(result){
				$("#CIRCLEDIV_" + statType).html(result);
			},
			cache: false
		});	
	}
	function newCircleEntry(requestType) {
		var newlineId = --$("#NEWLINE_ID").get(0).value;
		var cloneRow = $("#CIRCLETABLE_" + requestType + " tr").eq(-2);
		$(cloneRow[0].outerHTML.replace(/_0/g, "_" + newlineId)).show().insertBefore(cloneRow);
		newCircleValues(requestType, newlineId);
		addCircleList(newlineId);
		checkCircleOverlaps(requestType);
	}
	function newCircleValues(requestType, sliderId){
		$("#CIRCLEEFFDATE_" + sliderId + ", " + "#CIRCLEDISDATE_" + sliderId).each(function(){
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
						checkCircleOverlaps(requestType);
						addCircleList(sliderId);
					}
					else{
						alert("Invalid date. Enter a valid date in MM/DD/YYYY format");
					}
				}
			});
		});		
		var sliderElements = $("#CIRCLETABLE_" + requestType + " div[id^='CIRCLESLIDER_']:not([id='CIRCLESLIDER_0']):not([id='CIRCLESLIDER_" + sliderId + "'])");
		var sliderStep = parseInt($("#CIRCLESLIDER_" + sliderId).data("slider-step"));
		var sliderType = $("#CIRCLESLIDER_" + sliderId).data("slider-type");
		
		var useStart = -1;
		if(sliderElements.length > 0){
			for (var i = 0; i < sliderElements.length; i++){
				useStart = Math.max($("#" + sliderElements[i].id).slider("values",1), useStart);
			}
		}
		if(useStart == -1){
			useStart = parseInt($("#CIRCLESLIDER_" + sliderId).data("slider-min"));
		}
		var useEnd = Math.min(useStart - (-1)*sliderStep, parseInt($("#CIRCLESLIDER_" + sliderId).data("slider-max")));

		if(sliderType == "NUMBER"){
			var circleStartValue = useStart;
			var circleStartDisplay = useStart + "%";
			var circleEndValue = useEnd;
			var circleEndDisplay = useEnd + "%";
		}
		else{
			var circleStartValue = useStart;
			var MM = ("0" + Math.floor(useStart / 60)).substr(-2);
			var SS = ("0" + Math.floor(useStart % 60)).substr(-2);
			var circleStartDisplay = MM + ":" + SS;
			var circleEndValue = useEnd;
			MM = ("0" + Math.floor(useEnd / 60)).substr(-2);
			SS = ("0" + Math.floor(useEnd % 60)).substr(-2);
			var circleEndDisplay = MM + ":" + SS;
		}
		$("#CIRCLESTART_" + sliderId).val(circleStartValue);
		$("#CIRCLESTARTDISPLAY_" + sliderId).html(circleStartDisplay);
		$("#CIRCLEEND_" + sliderId).val(circleEndValue);
		$("#CIRCLEENDDISPLAY_" + sliderId).html(circleEndDisplay);
		initializeCircleSlider(sliderId);
	}
	function initializeCircleSlider(sliderId){
		var startValue = $("#CIRCLESTART_" + sliderId).val();
		var endValue = $("#CIRCLEEND_" + sliderId).val();
		var sliderStep = parseInt($("#CIRCLESLIDER_" + sliderId).data("slider-step"));
		var sliderMin = $("#CIRCLESLIDER_" + sliderId).data("slider-min");
		var sliderMax = $("#CIRCLESLIDER_" + sliderId).data("slider-max");
		var sliderType = $("#CIRCLESLIDER_" + sliderId).data("slider-type");
		var requestType = $("#CIRCLESLIDER_" + sliderId).data("request-type");
		var circleClass = $("#CIRCLECOLOR_" + sliderId).find("option:selected").data("circle-class");
		
		$("#CIRCLESLIDER_" + sliderId).slider({
			range: true,
			min: sliderMin,
			max: sliderMax,
			step: sliderStep,
			values: [startValue, endValue],
			stop: function(event, ui){
				checkCircleOverlaps(requestType);
				addCircleList(sliderId);
			},
			slide: function(event, ui){
				if (ui.handleIndex == 0){
					$("#CIRCLESTART_" + sliderId).val(ui.values[0]);
					if(sliderType == "NUMBER"){
						$("#CIRCLESTARTDISPLAY_" + sliderId).html(ui.values[0] + "%");
					}
					else{
						var MM = ("0" + Math.floor(ui.values[0] / 60)).substr(-2);
						var SS = ("0" + Math.floor(ui.values[0] % 60)).substr(-2);
						$("#CIRCLESTARTDISPLAY_" + sliderId).html(MM + ":" + SS);
					}
				}
				else{
					$("#CIRCLEEND_" + sliderId).val(ui.values[1]);
					if(sliderType == "NUMBER"){
						$("#CIRCLEENDDISPLAY_" + sliderId).html(ui.values[1] + "%");
					}
					else{
						var MM = ("0" + Math.floor(ui.values[1] / 60)).substr(-2);
						var SS = ("0" + Math.floor(ui.values[1] % 60)).substr(-2);
						$("#CIRCLEENDDISPLAY_" + sliderId).html(MM + ":" + SS);
					}
				}
				
			}
		});
		$("#CIRCLESLIDER_" + sliderId).find(".ui-slider-range").addClass(circleClass);
	}
	function checkCircleOverlaps(requestType){
		var overlapBool = 0;
		var sliderElements = $("#CIRCLETABLE_" + requestType + " div[id^='CIRCLESLIDER_']:not([id='CIRCLESLIDER_0'])");
		for (var i = 0; i < sliderElements.length; i++){
			if(overlapBool == 1){
				break;
			}
			var currentId = sliderElements[i].id.split("_");
			for (var j = i+1; j < sliderElements.length; j++){
				if(overlapBool == 1){
					break;
				}
				var loopId = sliderElements[j].id.split("_");
				var currentStart = $("#" + sliderElements[i].id).slider("values",0);
				var currentEnd = $("#" + sliderElements[i].id).slider("values",1);
				var loopStart = $("#" + sliderElements[j].id).slider("values",0);
				var loopEnd = $("#" + sliderElements[j].id).slider("values",1);

				if (currentStart < loopEnd && currentEnd > loopStart){
					var dotwString = "";
					$("input:checkbox[name=CIRCLEDOTW_" + currentId[1] + "]:checked").each(function(){
						dotwString += $(this).val();
					});
					$("input:checkbox[name=CIRCLEDOTW_" + loopId[1] + "]:checked").each(function(){
						dotwString += $(this).val();
					});
					if(dotwString.split("1").length - 1 > 1 || dotwString.split("2").length - 1 > 1 || dotwString.split("3").length - 1 > 1 || dotwString.split("4").length - 1 > 1 || dotwString.split("5").length - 1 > 1 || dotwString.split("6").length - 1 > 1 || dotwString.split("7").length - 1 > 1){
						var currentEffDate = moment($("#CIRCLEEFFDATE_" + currentId[1]).val(),"MM/DD/YYYY");
						var currentDisDate = moment($("#CIRCLEDISDATE_" + currentId[1]).val(),"MM/DD/YYYY");
						var loopEffDate = moment($("#CIRCLEEFFDATE_" + loopId[1]).val(),"MM/DD/YYYY");
						var loopDisDate = moment($("#CIRCLEDISDATE_" + loopId[1]).val(),"MM/DD/YYYY");
						if (currentEffDate <= loopDisDate && currentDisDate >= loopEffDate){
							overlapBool = 1;
						}
					}
				}
			}
		}
		if(overlapBool == 1){
			addCircleOverlapList(requestType);
		}
		else{
			removeCircleOverlapList(requestType);
		}
	}
	function addCircleOverlapList(requestType){
		if(overlapIdList.indexOf(requestType) == -1) {
			overlapIdList.push(requestType);
		}
		$("#CIRCLEDIV_" + requestType).addClass("error-color-background error-shadow");
		$("#CIRCLEDIV_" + requestType).find("tr[class~=new-entry-color]").addClass("new-entry-color-error");
		$("#CIRCLEDIV_" + requestType).find("select").addClass("error-color-background");
		$("#CIRCLE_SUBMIT").hide();
		$("#OVERLAP_MESSAGE").show();
	}
	function removeCircleOverlapList(requestType){
		if(overlapIdList.indexOf(requestType) != -1) {
			overlapIdList.splice(overlapIdList.indexOf(requestType), 1);
		}
		$("#CIRCLEDIV_" + requestType).removeClass("error-color-background error-shadow");
		$("#CIRCLEDIV_" + requestType).find("tr[class~=new-entry-color]").removeClass("new-entry-color-error");
		$("#CIRCLEDIV_" + requestType).find("select").removeClass("error-color-background");
		if(overlapIdList.length == 0){
			$("#CIRCLE_SUBMIT").show();
			$("#OVERLAP_MESSAGE").hide();
		}
	}
	function addCircleList(sliderId){
		if(circleIdList.indexOf(parseInt(sliderId)) == -1) {
			circleIdList.push(parseInt(sliderId));
		}
	}
	function removeCircleList(sliderId){
		if(circleIdList.indexOf(parseInt(sliderId)) != -1) {
			circleIdList.splice(circleIdList.indexOf(parseInt(sliderId)), 1);
		}
	}
	$(document).ready(function() {
		$("#main-wrap").on("click","[id^=CIRCLESTAT_]",function() {
			var idArray = this.id.split("_");
			$("#CIRCLEDIV_" + idArray[1]).toggle();
			if($(".edit-div").is(":visible")){
				if(overlapIdList.length == 0){
					$("#CIRCLE_SUBMIT").show();
				}
				else{
					$("#OVERLAP_MESSAGE").show();
				}
			}
			else{
				$("#CIRCLE_SUBMIT").hide();
				$("#OVERLAP_MESSAGE").hide();
			}
			if($("#CIRCLEDIV_" + idArray[1]).html().length == 0){
				$.when(editCircleStat(idArray[1], 0)).then(function(){
					checkCircleOverlaps(idArray[1]);
				});
			}
		});
		$("#main-wrap").on("click","[id^=CIRCLESTARTARROW_],[id^=CIRCLEENDARROW_]",function() {
			var idArray = this.id.split("_");
			var sliderType = $("#CIRCLESLIDER_" + idArray[2]).data("slider-type");
			if (idArray[0] == "CIRCLESTARTARROW"){
				var handleIndex = 0;
			}
			else{
				var handleIndex = 1;
			}
			var requestType = $("#CIRCLESLIDER_" + idArray[2]).data("request-type");
			var sliderValue = $("#CIRCLESLIDER_" + idArray[2]).slider("values",handleIndex);
			var intervalLength = parseInt($("#CIRCLESLIDER_" + idArray[2]).data("slider-interval"));

			if(idArray[1] == "LEFT"){
				if (idArray[0] == "CIRCLESTARTARROW"){
					var sliderMin = $("#CIRCLESLIDER_" + idArray[2]).slider("option","min");
				}
				else{
					var sliderMin = $("#CIRCLESLIDER_" + idArray[2]).slider("values",0);
				}
				$("#CIRCLESLIDER_" + idArray[2]).slider("values",handleIndex,Math.max(sliderValue-intervalLength,sliderMin));
				if(sliderType == "NUMBER"){
					var circleValue = Math.max(sliderValue-intervalLength,sliderMin);
					var circleDisplay = Math.max(sliderValue-intervalLength,sliderMin) + "%";
				}
				else{
					var circleValue = Math.max(sliderValue-intervalLength,sliderMin);
					var MM = ("0" + Math.floor(Math.max(sliderValue-intervalLength,sliderMin) / 60)).substr(-2);
					var SS = ("0" + Math.floor(Math.max(sliderValue-intervalLength,sliderMin) % 60)).substr(-2);
					var circleDisplay = MM + ":" + SS;
				}
			}
			else{
				if (idArray[0] == "CIRCLEENDARROW"){
					var sliderMax = $("#CIRCLESLIDER_" + idArray[2]).slider("option","max");
				}
				else{
					var sliderMax = $("#CIRCLESLIDER_" + idArray[2]).slider("values",1);
				}
				$("#CIRCLESLIDER_" + idArray[2]).slider("values",handleIndex,Math.min(sliderValue-(-1)*intervalLength,sliderMax));
				if(sliderType == "NUMBER"){
					var circleValue = Math.min(sliderValue-(-1)*intervalLength,sliderMax);
					var circleDisplay = Math.min(sliderValue-(-1)*intervalLength,sliderMax) + "%";
				}
				else{
					var circleValue = Math.min(sliderValue-(-1)*intervalLength,sliderMax);
					var MM = ("0" + Math.floor(Math.min(sliderValue-(-1)*intervalLength,sliderMax) / 60)).substr(-2);
					var SS = ("0" + Math.floor(Math.min(sliderValue-(-1)*intervalLength,sliderMax) % 60)).substr(-2);
					var circleDisplay = MM + ":" + SS;
				}
			}
			if (idArray[0] == "CIRCLESTARTARROW"){
				$("#CIRCLESTARTDISPLAY_" + idArray[2]).html(circleDisplay);
				$("#CIRCLESTART_" + idArray[2]).val(circleValue);
			}
			else{
				$("#CIRCLEENDDISPLAY_" + idArray[2]).html(circleDisplay);
				$("#CIRCLEEND_" + idArray[2]).val(circleValue);
			}
			checkCircleOverlaps(requestType);
			addCircleList(idArray[2]);
		});
		$("#main-wrap").on("change","select[id^=CIRCLECOLOR_]",function() {
			var idArray = this.id.split("_");
			var newClass = $(this).find("option:selected").data("circle-class")
			var classString = "GREEN YELLOW".replace(newClass,"");
			$("#CIRCLESLIDER_" + idArray[1]).find(".ui-slider-range").removeClass(classString).addClass(newClass);
			addCircleList(idArray[1]);
		});
		$("#main-wrap").on("click","[id^=CIRCLEBUTTON]",function() {
			var idArray = this.id.split("_");
			var $errorInput = $("#" + this.id.replace("CIRCLEBUTTON","CIRCLE"));
			
			if ($errorInput.prop("checked") == true){
				$errorInput.prop("checked", false);
				$(this).removeClass("new-entry-color");
			}
			else{
				$errorInput.prop("checked", true);
				$(this).addClass("new-entry-color");			
			}
			checkCircleOverlaps($("#CIRCLESLIDER_" + idArray[1]).data("request-type"));
			addCircleList(idArray[1]);
		});
		$("#main-wrap").on("click","[id^=NEWCIRCLEENTRY_]",function() {
			var idArray = this.id.split("_");
			newCircleEntry(idArray[1]);
		});
		$("#main-wrap").on("click","[id^=CIRCLEREFRESH_]",function() {
			var idArray = this.id.split("_");
			$.when(editCircleStat(idArray[1], 1)).then(function(){
				checkCircleOverlaps(idArray[1]);
			});
		});
	});
/* Circle Stats - End */

/* Schedule Controls - Start */
	function editScheduleControl(useWorkgroup, useControl, refreshBool){
		if(refreshBool == 1){
			$("div[id^='CONTROLSLIDER_'][data-workgroup='" + useWorkgroup + "'][data-control='" + useControl + "']").each(function(){
				var sliderArray = this.id.split("_");
				removeControlList(sliderArray[1]);
			});
		}
		return $.ajax({
			url: "includes/schedulecontrols.asp?WORKGROUP=" + useWorkgroup + "&CONTROL=" + useControl + "&DATE=" + $("#PARAMETER_DATE").val(),
			success: function(result){
				$("#CONTROLSDIV_" + useWorkgroup + "_" + useControl).html(result);
			},
			cache: false
		});	
	}
	function newControlEntry(useWorkgroup, useControl) {
		var newlineId = --$("#NEWLINE_ID").get(0).value;
		var cloneRow = $("#CONTROLTABLE_" + useWorkgroup + "_" + useControl + " tr").eq(-2);
		$(cloneRow[0].outerHTML.replace(/_0/g, "_" + newlineId)).show().insertBefore(cloneRow);
		newControlValues(useWorkgroup, useControl, newlineId);
		addControlList(newlineId);
		checkControlOverlaps(useWorkgroup, useControl);
	}
	function newControlValues(useWorkgroup, useControl, sliderId){
		$("#CONTROLEFFDATE_" + sliderId + ", " + "#CONTROLDISDATE_" + sliderId).each(function(){
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
						checkControlOverlaps(useWorkgroup, useControl);
						addControlList(sliderId);
					}
					else{
						alert("Invalid date. Enter a valid date in MM/DD/YYYY format");
					}
				}
			});
		});
		var sliderCategory = ["INTERVAL","HOURS","SCORE","VALUE"];
		for (i = 0; i < sliderCategory.length; i++) {
			if ($("#CONTROLTABLE_" + useWorkgroup + "_" + useControl + " div[id='CONTROLSLIDER_" + sliderCategory[i] + "_0']").length > 0){
				var sliderElements = $("#CONTROLTABLE_" + useWorkgroup + "_" + useControl + " div[id^='CONTROLSLIDER_" + sliderCategory[i] + "_']:not([id='CONTROLSLIDER_" + sliderCategory[i] + "_0']):not([id='CONTROLSLIDER_" + sliderCategory[i] + "_" + sliderId + "'])");
				
				var sliderRange = $("#CONTROLSLIDER_" + sliderCategory[i] + "_" + sliderId).data("range");
				var sliderStep = parseFloat($("#CONTROLSLIDER_" + sliderCategory[i] + "_" + sliderId).data("slider-step"));
				var sliderType = $("#CONTROLSLIDER_" + sliderCategory[i] + "_" + sliderId).data("slider-type");
				var sliderMin = parseInt($("#CONTROLSLIDER_" + sliderCategory[i] + "_" + sliderId).data("slider-min"));
				var sliderMax = parseInt($("#CONTROLSLIDER_" + sliderCategory[i] + "_" + sliderId).data("slider-max"));
				if(sliderRange){
					var useStart;
					var useEnd;
					
					if(sliderCategory[i] == "INTERVAL"){
						useStart = sliderMin;
						useEnd = sliderMax;
					}
					else{
						useStart = -1;
						if(sliderElements.length > 0){
							useStart = Math.max($("#" + sliderElements[sliderElements.length-1].id).slider("values",1), useStart);
						}
						if(useStart == -1){
							useStart = sliderMin;
						}
						useEnd = Math.min(useStart - (-1)*sliderStep, sliderMax);
					}
					if(sliderType == "TIME"){
						var MM = ("0" + Math.floor(useStart / 60)).substr(-2);
						var SS = ("0" + Math.floor(useStart % 60)).substr(-2);
						var controlStart = MM + ":" + SS;

						MM = ("0" + Math.floor(useEnd / 60)).substr(-2);
						SS = ("0" + Math.floor(useEnd % 60)).substr(-2);
						var controlEnd = MM + ":" + SS;
					}
					else{
						var controlStart = parseFloat(useStart);
						var controlEnd = parseFloat(useEnd);
					}
					$("#CONTROLSTART_" + sliderCategory[i] + "_" + sliderId).val(controlStart);
					$("#CONTROLSTARTDISPLAY_" + sliderCategory[i] + "_" + sliderId).html(controlStart);
					$("#CONTROLEND_" + sliderCategory[i] + "_" + sliderId).val(controlEnd);
					$("#CONTROLENDDISPLAY_" + sliderCategory[i] + "_" + sliderId).html(controlEnd);
				}
				else{
					var useValue = sliderMin;
					if(sliderType == "TIME"){
						var MM = ("0" + Math.floor(useValue / 60)).substr(-2);
						var SS = ("0" + Math.floor(useValue % 60)).substr(-2);
						var controlValue = MM + ":" + SS;
					}
					else{
						var controlValue = parseFloat(useValue);
					}
					$("#CONTROL_" + sliderCategory[i] + "_" + sliderId).val(controlValue);
				}
				initializeControlSlider(sliderCategory[i], sliderId);
			}
		}
	}
	function initializeControlSlider(sliderCategory, sliderId){
		var sliderRange = $("#CONTROLSLIDER_" + sliderCategory + "_" + sliderId).data("range");
		var sliderType = $("#CONTROLSLIDER_" + sliderCategory + "_" + sliderId).data("slider-type");
		if (sliderRange){
			if(sliderType == "TIME"){
				var startArray = $("#CONTROLSTART_" + sliderCategory + "_" + sliderId).val().split(":");
				var startValue = 60*startArray[0] - (-1)*startArray[1];
				var endArray = $("#CONTROLEND_" + sliderCategory + "_" + sliderId).val().split(":");
				var endValue = 60*endArray[0] - (-1)*endArray[1];
			}
			else{
				var startValue = parseFloat($("#CONTROLSTART_" + sliderCategory + "_" + sliderId).val());
				var endValue = parseFloat($("#CONTROLEND_" + sliderCategory + "_" + sliderId).val());
			}
		}
		else{
			if(sliderType == "TIME"){
				var useArray = $("#CONTROL_" + sliderCategory + "_" + sliderId).val().split(":");
				var useValue = 60*useArray[0] - (-1)*useArray[1];
			}
			else{
				var useValue = parseFloat($("#CONTROL_" + sliderCategory + "_" + sliderId).val());
			}
		}
		var sliderWorkgroup = $("#CONTROLSLIDER_" + sliderCategory + "_" + sliderId).data("workgroup");
		var sliderControl = $("#CONTROLSLIDER_" + sliderCategory + "_" + sliderId).data("control");
		var sliderMin = parseInt($("#CONTROLSLIDER_" + sliderCategory + "_" + sliderId).data("slider-min"));
		var sliderMax = parseInt($("#CONTROLSLIDER_" + sliderCategory + "_" + sliderId).data("slider-max"));
		var sliderStep = parseFloat($("#CONTROLSLIDER_" + sliderCategory + "_" + sliderId).data("slider-step"));
		
		if(sliderRange){
			$("#CONTROLSLIDER_" + sliderCategory + "_" + sliderId).slider({
				range: true,
				values: [startValue, endValue],
				min: sliderMin,
				max: sliderMax,
				step: sliderStep,
				stop: function(event, ui){
					checkControlOverlaps(sliderWorkgroup, sliderControl);
					addControlList(sliderId);
				},
				slide: function(event, ui){
					if (ui.handleIndex == 0){
						if(sliderType == "TIME"){
							var MM = ("0" + Math.floor(ui.values[0] / 60)).substr(-2);
							var SS = ("0" + Math.floor(ui.values[0] % 60)).substr(-2);
							$("#CONTROLSTART_" + sliderCategory + "_" + sliderId).val(MM + ":" + SS);
							$("#CONTROLSTARTDISPLAY_" + sliderCategory + "_" + sliderId).html(MM + ":" + SS);
						}
						else{
							$("#CONTROLSTART_" + sliderCategory + "_" + sliderId).val(ui.values[0]);
							$("#CONTROLSTARTDISPLAY_" + sliderCategory + "_" + sliderId).html(ui.values[0]);
						}
					}
					else{
						if(sliderType == "TIME"){
							var MM = ("0" + Math.floor(ui.values[1] / 60)).substr(-2);
							var SS = ("0" + Math.floor(ui.values[1] % 60)).substr(-2);
							$("#CONTROLEND_" + sliderCategory + "_" + sliderId).val(MM + ":" + SS);
							$("#CONTROLENDDISPLAY_" + sliderCategory + "_" + sliderId).html(MM + ":" + SS);
						}
						else{
							$("#CONTROLEND_" + sliderCategory + "_" + sliderId).val(ui.values[1]);
							$("#CONTROLENDDISPLAY_" + sliderCategory + "_" + sliderId).html(ui.values[1]);
						}
					}
				}
			});
			$("#CONTROLSLIDER_" + sliderCategory + "_" + sliderId).find(".ui-slider-range").addClass("CONTROL");
		}
		else{
			$("#CONTROLSLIDER_" + sliderCategory + "_" + sliderId).slider({
				range: false,
				value: useValue,
				min: sliderMin,
				max: sliderMax,
				step: sliderStep,
				create: function(){
					$("#CONTROLHANDLE_" + sliderCategory + "_" + sliderId).text(useValue);
				},
				stop: function(event, ui){
					addControlList(sliderId);
				},
				slide: function(event, ui){
					$("#CONTROLHANDLE_" + sliderCategory + "_" + sliderId).text(ui.value);
					$("#CONTROL_" + sliderCategory + "_" + sliderId).val(ui.value);
				}
			});		
		}
	}
	function checkControlOverlaps(useWorkgroup, useControl){
		var overlapBool = 0;
		var sliderCategory = ["INTERVAL","HOURS","SCORE"];
		if ($("#CONTROLTABLE_" + useWorkgroup + "_" + useControl + " div[id='CONTROLSLIDER_INTERVAL_0']").length == 0){
			sliderCategory.splice(sliderCategory.indexOf("INTERVAL"), 1);
		}
		if ($("#CONTROLTABLE_" + useWorkgroup + "_" + useControl + " div[id='CONTROLSLIDER_HOURS_0']").length == 0){
			sliderCategory.splice(sliderCategory.indexOf("HOURS"), 1);
		}
		if ($("#CONTROLTABLE_" + useWorkgroup + "_" + useControl + " div[id='CONTROLSLIDER_SCORE_0']").length == 0){
			sliderCategory.splice(sliderCategory.indexOf("SCORE"), 1);
		}
		var dateElements = $("#CONTROLTABLE_" + useWorkgroup + "_" + useControl + " input:text[id^='CONTROLEFFDATE_']:not(input:text[name='CONTROLEFFDATE_0'])");
		for (var i = 0; i < dateElements.length; i++){
			if(overlapBool == 1){
				break;
			}
			var currentId = dateElements[i].id.split("_");
			for (var j = i + 1; j < dateElements.length; j++){
				if(overlapBool == 1){
					break;
				}
				var loopId = dateElements[j].id.split("_");
				var currentEffDate = moment($("#CONTROLEFFDATE_" + currentId[1]).val(),"MM/DD/YYYY");
				var currentDisDate = moment($("#CONTROLDISDATE_" + currentId[1]).val(),"MM/DD/YYYY");
				var loopEffDate = moment($("#CONTROLEFFDATE_" + loopId[1]).val(),"MM/DD/YYYY");
				var loopDisDate = moment($("#CONTROLDISDATE_" + loopId[1]).val(),"MM/DD/YYYY");
				if (currentEffDate <= loopDisDate && currentDisDate >= loopEffDate){
					overlapBool = 1;
					if($("input:checkbox[name='CONTROLDOTW_" + currentId[1] + "']").length > 0){
						var dotwString = "";
						$("input:checkbox[name=CONTROLDOTW_" + currentId[1] + "]:checked").each(function(){
							dotwString += $(this).val();
						});
						$("input:checkbox[name=CONTROLDOTW_" + loopId[1] + "]:checked").each(function(){
							dotwString += $(this).val();
						});
						if(!(dotwString.split("1").length - 1 > 1 || dotwString.split("2").length - 1 > 1 || dotwString.split("3").length - 1 > 1 || dotwString.split("4").length - 1 > 1 || dotwString.split("5").length - 1 > 1 || dotwString.split("6").length - 1 > 1 || dotwString.split("7").length - 1 > 1)){
							overlapBool = 0;
						}
					}
					if(overlapBool == 1){
						for (k = 0; k < sliderCategory.length; k++) {
							if(overlapBool == 0){
								break;
							}
							currentStart = $("#CONTROLSLIDER_" + sliderCategory[k] + "_" + currentId[1]).slider("values",0);
							currentEnd = $("#CONTROLSLIDER_" + sliderCategory[k] + "_" + currentId[1]).slider("values",1);
							loopStart = $("#CONTROLSLIDER_" + sliderCategory[k] + "_" + loopId[1]).slider("values",0);
							loopEnd = $("#CONTROLSLIDER_" + sliderCategory[k] + "_" + loopId[1]).slider("values",1);
							if (!(currentStart < loopEnd && currentEnd > loopStart)){
								overlapBool = 0;
							}
						}
					}
				}
			}
		}
		if(overlapBool == 1){
			addControlOverlapList(useWorkgroup, useControl);
		}
		else{
			removeControlOverlapList(useWorkgroup, useControl);
		}
	}
	function addControlOverlapList(useWorkgroup, useControl){
		if(overlapIdList.indexOf(useWorkgroup + "_" + useControl) == -1) {
			overlapIdList.push(useWorkgroup + "_" + useControl);
		}
		$("#CONTROLSDIV_"+ useWorkgroup + "_" + useControl).addClass("error-color-background error-shadow");
		$("#CONTROLSDIV_"+ useWorkgroup + "_" + useControl).find("tr[class~=new-entry-color]").addClass("new-entry-color-error");
		$("input[id^=CONTROL_SUBMIT_]").filter(":visible").each(function(){
			var idArray = this.id.split("_");
			$("#CONTROL_SUBMIT_" + idArray[2]).hide();
			$("#OVERLAP_MESSAGE_" + idArray[2]).show();
		});
	}
	function removeControlOverlapList(useWorkgroup, useControl){
		if(overlapIdList.indexOf(useWorkgroup + "_" + useControl) != -1) {
			overlapIdList.splice(overlapIdList.indexOf(useWorkgroup + "_" + useControl), 1);
		}
		$("#CONTROLSDIV_"+ useWorkgroup + "_" + useControl).removeClass("error-color-background error-shadow");
		$("#CONTROLSDIV_"+ useWorkgroup + "_" + useControl).find("tr[class~=new-entry-color]").removeClass("new-entry-color-error");
		if(overlapIdList.length == 0){
			$("div[id^=OVERLAP_MESSAGE_]").filter(":visible").each(function(){
				var idArray = this.id.split("_");
				$("#CONTROL_SUBMIT_" + idArray[2]).show();
				$("#OVERLAP_MESSAGE_" + idArray[2]).hide();
			});
		}
	}
	function addControlList(sliderId){
		if(controlIdList.indexOf(parseInt(sliderId)) == -1) {
			controlIdList.push(parseInt(sliderId));
		}
	}
	function removeControlList(sliderId){
		if(controlIdList.indexOf(parseInt(sliderId)) != -1) {
			controlIdList.splice(controlIdList.indexOf(parseInt(sliderId)), 1);
		}
	}
	$(document).ready(function() {
		$("#main-wrap").on("click","[id^=CONTROLSTRIGGER_]",function() {
			var idArray = this.id.split("_");
			$("#CONTROLSDIV_" + idArray[1] + "_" + idArray[2]).toggle();
			if($(".edit-div[id^='CONTROLSDIV_" + idArray[1] + "']").is(":visible")){
				if(overlapIdList.length == 0){
					$("#CONTROL_SUBMIT_" + idArray[1]).show();
				}
				else{
					$("#OVERLAP_MESSAGE_" + idArray[1]).show();
				}
			}
			else{
				$("#CONTROL_SUBMIT_" + idArray[1]).hide();
				$("#OVERLAP_MESSAGE_" + idArray[1]).hide();
			}
			if ($("#CONTROLSDIV_" + idArray[1] + "_" + idArray[2]).is(":visible")){
				if ($("#CONTROLSDIV_" + idArray[1] + "_" + idArray[2]).hasClass("today-color")){
					$(this).removeClass("today-color today-color-border white-background").addClass("today-color-background");
				}
				else{
					$(this).removeClass("past-color past-color-border white-background").addClass("past-color-background");
				}
			}
			else{
				if ($("#CONTROLSDIV_" + idArray[1] + "_" + idArray[2]).hasClass("today-color")){
					$(this).removeClass("today-color-background").addClass("today-color today-color-border white-background");
				}
				else{
					$(this).removeClass("past-color-background").addClass("past-color past-color-border white-background");
				}			
			}
			if($("#CONTROLSDIV_" + idArray[1] + "_" + idArray[2]).html().length == 0){
				$.when(editScheduleControl(idArray[1], idArray[2], 0)).then(function(){
					checkControlOverlaps(idArray[1], idArray[2]);
				});
			}
		});
		$("#main-wrap").on("click","[id^=CONTROLSTARTARROW_],[id^=CONTROLENDARROW_]",function() {
			var idArray = this.id.split("_");
			var sliderType = $("#CONTROLSLIDER_" + idArray[1] + "_" + idArray[3]).data("slider-type");
			var handleIndex = (idArray[0] == "CONTROLSTARTARROW") ? 0 : 1;
			var sliderValue = $("#CONTROLSLIDER_" + idArray[1] + "_" + idArray[3]).slider("values",handleIndex);
			var intervalLength = parseFloat($("#CONTROLSLIDER_" + idArray[1] + "_" + idArray[3]).data("slider-interval"));

			if(idArray[2] == "LEFT"){
				if (idArray[0] == "CONTROLSTARTARROW"){
					var sliderMin = $("#CONTROLSLIDER_" + idArray[1] + "_" + idArray[3]).slider("option","min");
				}
				else{
					var sliderMin = $("#CONTROLSLIDER_" + idArray[1] + "_" + idArray[3]).slider("values",0);
				}
				$("#CONTROLSLIDER_" + idArray[1] + "_" + idArray[3]).slider("values",handleIndex,Math.max(sliderValue-intervalLength,sliderMin));
				if(sliderType == "TIME"){
					var MM = ("0" + Math.floor(Math.max(sliderValue-intervalLength,sliderMin) / 60)).substr(-2);
					var SS = ("0" + Math.floor(Math.max(sliderValue-intervalLength,sliderMin) % 60)).substr(-2);
					var controlValue = MM + ":" + SS;
				}
				else{
					var controlValue = Math.max(sliderValue-intervalLength,sliderMin);
				}
			}
			else{
				if (idArray[0] == "CONTROLENDARROW"){
					var sliderMax = $("#CONTROLSLIDER_" + idArray[1] + "_" + idArray[3]).slider("option","max");
				}
				else{
					var sliderMax = $("#CONTROLSLIDER_" + idArray[1] + "_" + idArray[3]).slider("values",1);
				}
				$("#CONTROLSLIDER_" + idArray[1] + "_" + idArray[3]).slider("values",handleIndex,Math.min(sliderValue-(-1)*intervalLength,sliderMax));
				if(sliderType == "TIME"){
					var MM = ("0" + Math.floor(Math.min(sliderValue-(-1)*intervalLength,sliderMax) / 60)).substr(-2);
					var SS = ("0" + Math.floor(Math.min(sliderValue-(-1)*intervalLength,sliderMax) % 60)).substr(-2);
					var controlValue = MM + ":" + SS;
				}
				else{
					var controlValue = Math.min(sliderValue-(-1)*intervalLength,sliderMax);
				}
			}
			if (idArray[0] == "CONTROLSTARTARROW"){
				$("#CONTROLSTART_" + idArray[1] + "_" + idArray[3]).val(controlValue);
				$("#CONTROLSTARTDISPLAY_" + idArray[1] + "_" + idArray[3]).html(controlValue);
			}
			else{
				$("#CONTROLEND_" + idArray[1] + "_" + idArray[3]).val(controlValue);
				$("#CONTROLENDDISPLAY_" + idArray[1] + "_" + idArray[3]).html(controlValue);
			}
			var sliderWorkgroup = $("#CONTROLSLIDER_" + idArray[1] + "_" + idArray[3]).data("workgroup");
			var sliderControl = $("#CONTROLSLIDER_" + idArray[1] + "_" + idArray[3]).data("control");
			checkControlOverlaps(sliderWorkgroup, sliderControl);
			addControlList(idArray[3]);
		});
		$("#main-wrap").on("click","[id^=CONTROLBUTTON]",function() {
			var idArray = this.id.split("_");
			var $errorInput = $("#" + this.id.replace("CONTROLBUTTON","CONTROL"));
			
			if ($errorInput.prop("checked") == true){
				$errorInput.prop("checked", false);
				$(this).removeClass("new-entry-color");
			}
			else{
				$errorInput.prop("checked", true);
				$(this).addClass("new-entry-color");			
			}
			var buttonWorkgroup = $(this).data("workgroup");
			var buttonControl = $(this).data("control");
			checkControlOverlaps(buttonWorkgroup, buttonControl);
			addControlList(idArray[1]);
		});
		$("#main-wrap").on("click","[id^=NEWCONTROLENTRY_]",function() {
			var idArray = this.id.split("_");
			newControlEntry(idArray[1], idArray[2]);
		});
		$("#main-wrap").on("click","[id^=CONTROLREFRESH_]",function() {
			var idArray = this.id.split("_");
			$.when(editScheduleControl(idArray[1], idArray[2], 1)).then(function(){
				checkControlOverlaps(idArray[1], idArray[2]);
			});
		});
	});
/* Schedule Controls - End */

/* User Admin - Start */
	function editAdmin(userId, refreshBool){
		if(refreshBool == 1){
			$("tr[id^='ADMINDETAILROW_'][data-user='" + userId + "']").each(function(){
				var idArray = this.id.split("_");
				removeAdminList("DETAIL_" + idArray[1]);
			});
		}
		return $.ajax({
			url: "includes/editadmin.asp?DATE=" + $("#PARAMETER_DATE").val() + "&AGENT=" + userId,
			success: function(result){
				$("#EDITADMINDIV_" + userId).html(result);
			},
			complete: function(){
				if($("#ADMINUSER_" + userId).val() != ""){
					$("#EDITADMINDETAILSCAPTION_" + userId).html($("#ADMINUSER_" + userId).val() + "'s User Records");
				}
				else{
					$("#EDITADMINDETAILSCAPTION_" + userId).empty();
				}
			},
			cache: false
		});	
	}
	function newAdminMasterEntry() {
		var newlineId = --$("#NEWLINE_ID").get(0).value;
		var cloneRow = $("#EDITADMINMASTERTABLE tr").eq(-3);
		var useDataTable = initAdminDataTable();
		useDataTable.row.add($(cloneRow[0].outerHTML.replace(/_0/g, "_" + newlineId)).show()).draw();
		useDataTable.row("#ADMINROW_" + newlineId).child($(
			'<tr id="EDITADMINROW_' + newlineId + '" style="display:none;">' + 
				'<td id="EDITADMINDIV_WRAPPER_' + newlineId + '" colspan="8">' +
					'<div id="EDITADMINDIV_' + newlineId + '"></div>' +
				'</td>' +
			'</tr>'
			)
		).show();
		
		addAdminList("MASTER_"+ newlineId);
	}
	function newAdminDetailsEntry(userId) {
		var newlineId = --$("#NEWLINE_ID").get(0).value;
		var cloneRow = $("#EDITADMINDETAILSTABLE_" + userId + " tr").eq(-2);
		$(cloneRow[0].outerHTML.replace(/_0/g, "_" + newlineId)).show().insertBefore(cloneRow);
		
		initializeAdmin(userId, newlineId);
		addAdminList("DETAIL_" + newlineId);
		addAdminList("MASTER_" + $("#ADMINDETAILUSER_" + newlineId).val());
		checkAdminOverlaps(userId);
	}
	function initializeAdmin(userId, adminId){
		$("#ADMINEFFDATE_" + adminId + ", " + "#ADMINDISDATE_" + adminId).each(function(){
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
						checkAdminOverlaps(userId);
						addAdminList("DETAIL_" + adminId);
						addAdminList("MASTER_" + $("#ADMINDETAILUSER_" + adminId).val());
					}
					else{
						alert("Invalid date. Enter a valid date in MM/DD/YYYY format");
					}
				}
			});
		});
		var adminElements = $("tr[id^='ADMINDETAILROW_'][data-user='" + userId + "']:not([id='ADMINDETAILROW_0']):not([id='ADMINDETAILROW_" + adminId + "'])");
		if(adminElements.length > 0){
			var idArray = adminElements[adminElements.length - 1].id.split("_");
			var useDept = $("#ADMINDEPT_" + idArray[1]).val();
			var useTeam = $("#ADMINTEAM_" + idArray[1]).val();
			var useJob = $("#ADMINJOB_" + idArray[1]).val();
			var useClass = $("#ADMINCLASS_" + idArray[1]).val();
			var useLocation = $("#ADMINLOCATION_" + idArray[1]).val();
			var useHours = $("#ADMINHOURS_" + idArray[1]).val();
			var useSupervisor = $("#ADMINSUPERVISOR_" + idArray[1]).val();
			var usePhone = $("#ADMINPHONE_" + idArray[1]).val();
			var useJobCode = $("#ADMINJOBCODE_" + idArray[1]).val();
			var usePay = $("#ADMINPAY_" + idArray[1]).val();
			var useDisDate = moment("12/31/2040","MM/DD/YYYY");
			var useEffDate = moment.min(moment($("#ADMINDISDATE_" + idArray[1]).val(),"MM/DD/YYYY").add(1,"d"), useDisDate);
		}
		else{
			var useDept = "RES";
			var useTeam = "SES";
			var useJob = "AGT";
			var useClass = "RGFT";
			var useLocation = "MOT";
			var useHours = "40";
			var useSupervisor = "9151";
			var usePhone = "";
			var useJobCode = "008816";
			var usePay = "12";
			var useDisDate = moment("12/31/2040","MM/DD/YYYY");
			var useEffDate = moment($("#PARAMETER_DATE").val(),"MM/DD/YYYY");
		}
		$("#ADMINDEPT_" + adminId).val(useDept);
		$("#ADMINTEAM_" + adminId).val(useTeam);
		$("#ADMINJOB_" + adminId).val(useJob);
		$("#ADMINCLASS_" + adminId).val(useClass);
		$("#ADMINLOCATION_" + adminId).val(useLocation);
		$("#ADMINHOURS_" + adminId).val(useHours);
		$("#ADMINSUPERVISOR_" + adminId).val(useSupervisor);	
		$("#ADMINPHONE_" + adminId).val(usePhone);
		$("#ADMINJOBCODE_" + adminId).val(useJobCode);
		$("#ADMINPAY_" + adminId).val(usePay);	
		$("#ADMINEFFDATE_" + adminId).val(useEffDate.format("l"));
		$("#ADMINDISDATE_" + adminId).val(useDisDate.format("l"));
	}
	function checkAdminOverlaps(userId){
		var overlapBool = 0;
		var adminElements = $("tr[id^='ADMINDETAILROW_'][data-user='" + userId + "']:not([id='ADMINDETAILROW_0'])");
		for (var i = 0; i < adminElements.length; i++){
			if(overlapBool == 1){
				break;
			}
			var currentId = adminElements[i].id.split("_");
			for (var j = i + 1; j < adminElements.length; j++){
				if(overlapBool == 1){
					break;
				}
				var loopId = adminElements[j].id.split("_");
				var currentEffDate = moment($("#ADMINEFFDATE_" + currentId[1]).val(),"MM/DD/YYYY");
				var currentDisDate = moment($("#ADMINDISDATE_" + currentId[1]).val(),"MM/DD/YYYY");
				var loopEffDate = moment($("#ADMINEFFDATE_" + loopId[1]).val(),"MM/DD/YYYY");
				var loopDisDate = moment($("#ADMINDISDATE_" + loopId[1]).val(),"MM/DD/YYYY");
				if (currentEffDate <= loopDisDate && currentDisDate >= loopEffDate){
					overlapBool = 1;
				}
			}
		}
		if(overlapBool == 1){
			addAdminOverlapList(userId);
		}
		else{
			removeAdminOverlapList(userId);
		}
	}
	function addAdminOverlapList(overlapUser){
		if(overlapIdList.indexOf(overlapUser) == -1) {
			overlapIdList.push(overlapUser);
		}
		$("#ADMINROW_" + overlapUser).addClass("error-color-background");
		$("#EDITADMINROW_" + overlapUser).addClass("error-color-background");
		$("#EDITADMINDETAILSTABLE_" + overlapUser).find("tr[class~=new-entry-color]").addClass("new-entry-color-error");
		$("#EDITADMINDETAILSTABLE_" + overlapUser).find("input, select").addClass("error-color-background");
		$("#PULSE_FORM_DIV").addClass("error-shadow");
		$("#PULSE_SUBMIT").hide();
		$("#OVERLAP_MESSAGE").show();
	}
	function removeAdminOverlapList(overlapUser){
		if(overlapIdList.indexOf(overlapUser) != -1) {
			overlapIdList.splice(overlapIdList.indexOf(overlapUser), 1);
		}
		$("#ADMINROW_" + overlapUser).removeClass("error-color-background");
		$("#EDITADMINROW_" + overlapUser).removeClass("error-color-background");
		$("#EDITADMINDETAILSTABLE_" + overlapUser).find("tr[class~=new-entry-color]").removeClass("new-entry-color-error");
		$("#EDITADMINDETAILSTABLE_" + overlapUser).find("input, select").removeClass("error-color-background");
		if(overlapIdList.length == 0){
			$("#PULSE_FORM_DIV").removeClass("error-shadow");
			$("#PULSE_SUBMIT").show();
			$("#OVERLAP_MESSAGE").hide();
		}
	}
	function addAdminList(adminId){
		if(adminIdList.indexOf(adminId) == -1) {
			adminIdList.push(adminId);
		}
	}
	function removeAdminList(adminId){
		if(adminIdList.indexOf(adminId) != -1) {
			adminIdList.splice(adminIdList.indexOf(adminId), 1);
		}
	}
	function adminSecurity(userId) {
		var useLocation = "";
		var useDept = "";
		var useTeam = "";
		var useJob = "";
		
		var adminElements = $("tr[id^='ADMINDETAILROW_'][data-user='" + userId + "']:not([id='ADMINDETAILROW_0'])");
		for (var i = adminElements.length - 1; i >= 0; i--){
			var loopId = adminElements[i].id.split("_");
			var loopDisDate = moment($("#ADMINDISDATE_" + loopId[1]).val(),"MM/DD/YYYY").format('L');
			var securityDate = moment().format('L');
			if(securityDate <= loopDisDate){ 
				useLocation = $("#ADMINLOCATION_" + loopId[1]).val();
				useDept = $("#ADMINDEPT_" + loopId[1]).val();
				useTeam = $("#ADMINTEAM_" + loopId[1]).val();
				useJob = $("#ADMINJOB_" + loopId[1]).val();
			}
			else{
				break;
			}
		}
		var bPopup = $("#popupdiv").bPopup({
			speed: 650,
			transition: "slideDown",
			transitionClose: "slideUp",
			content:"ajax",
			contentContainer:"#popupdiv",
			loadUrl: "includes/securityaccess.asp?DATE=" + $("#PARAMETER_DATE").val() + "&AGENT=" + userId + "&LOCATION=" + useLocation + "&DEPARTMENT=" + useDept + "&TEAM=" + useTeam + "&JOB=" + useJob + "&JUNK="+ new Date().getTime(),
			loadCallback: function(){
				if($("#ADMINUSER_" + userId).val() != ""){
					$("#SECURITYACCESSCAPTION_" + userId).html($("#ADMINUSER_" + userId).val() + "'s Access");
				}
				else{
					$("#SECURITYACCESSCAPTION_" + userId).html("New Access");
				}
				$("#SECURITY_FORM").ajaxForm({
					clearForm: true,
					error: function() {
						$("#popupdiv").html("Update Failed");
					},
					success: function() {
						bPopup.close();
					}
				});
			}
		});	
	}
	$(document).ready(function() {
		$("#main-wrap").on("click","[id^=NEWADMINMASTERENTRY]",newAdminMasterEntry);
		$("#main-wrap").on("click","[id^=ADMINUSER_]",function() {
			var idArray = this.id.split("_");
			$("#EDITADMINROW_" + idArray[1]).toggle();
			if($("#EDITADMINDIV_" + idArray[1]).html().length == 0){
				$.when(editAdmin(idArray[1], 0)).then(function(){
					checkAdminOverlaps(idArray[1]);
				});
			}
		});
		$("#main-wrap").on("click","[id^=ADMINREFRESH_]",function() {
			var idArray = this.id.split("_");
			$.when(editAdmin(idArray[1], 1)).then(function(){
				checkAdminOverlaps(idArray[1]);
			});
		});
		$("#main-wrap").on("click","[id^=ADMINSECURITY_]",function() {
			var idArray = this.id.split("_");
			adminSecurity(idArray[1]);
		});
		$("#main-wrap").on("click","[id^=NEWADMINDETAILSENTRY_]",function() {
			var idArray = this.id.split("_");
			newAdminDetailsEntry(idArray[1]);
		});
		$("#main-wrap").on("change","input:text[id^=ADMINUSER_]",function() {
			var idArray = this.id.split("_");
			addAdminList("MASTER_" + idArray[1]);
			if($("#EDITADMINDETAILSCAPTION_" + idArray[1]).length){
				if($("#ADMINUSER_" + idArray[1]).val() != ""){
					$("#EDITADMINDETAILSCAPTION_" + idArray[1]).html($("#ADMINUSER_" + idArray[1]).val() + "'s User Records");
				}
				else{
					$("#EDITADMINDETAILSCAPTION_" + idArray[1]).empty();
				}
			}
		});
		$("#main-wrap").on("change","input:text[id^=ADMINWINDOWS_], input:text[id^=ADMINPPR_], input:text[id^=ADMINNAVIGATOR_], input:text[id^=ADMINBADGE_], input:text[id^=ADMINEXT_], input:text[id^=ADMINTEXT_], input:text[id^=ADMINEMAIL_]",function() {
			var idArray = this.id.split("_");
			addAdminList("MASTER_" + idArray[1]);
		});
		$("#main-wrap").on("change","select[id^=ADMINDEPT_], select[id^=ADMINTEAM_], select[id^=ADMINJOB_], select[id^=ADMINCLASS_], select[id^=ADMINLOCATION_], select[id^=ADMINHOURS_], select[id^=ADMINSUPERVISOR_], input:text[id^=ADMINPHONE_], input:text[id^=ADMINJOBCODE_]",function() {
			var idArray = this.id.split("_");
			addAdminList("DETAIL_" + idArray[1]);
			addAdminList("MASTER_" + $("#ADMINDETAILUSER_" + idArray[1]).val());
		});
		$("#popupdiv").on("click","[id^=SECURITYROW_]",function() {
			$("#SECURITY_TEXT").html("");
			var idArray = this.id.split("_");
			var $accessInput = $("#" + this.id.replace("ROW","ACCESS"));
			
			if ($accessInput.prop("checked") == true){
				$accessInput.prop("checked", false);
				$(this).removeClass("new-entry-color");
			}
			else{
				$accessInput.prop("checked", true);
				$(this).addClass("new-entry-color");			
			}
		});
		$("#popupdiv").on("click","#SECURITY_PROFILE",function() {
			var securityDepartment = $(this).data("department");
			var securityTeam = $(this).data("team");
			var securityJob = $(this).data("job");
			var securityMatch = $("#SECURITY_MATCH").val();
			if(securityMatch != "0"){
				$.ajax({
					dataType: "json",
					url: "includes/generatesecurity.asp?DEPARTMENT=" + securityDepartment + "&TEAM=" + securityTeam + "&JOB=" + securityJob + "&MATCH=" + securityMatch,
					success: function(result){
						var numChanges = 0;
						for(var i = 0; i < result.length; i++){
							var useTypeId = result[i].typeId;
							var useAccessId = result[i].accessId;
							var useAccessFlag = result[i].accessFlag;
							var $accessInput = $("#SECURITYACCESS_" + useTypeId + "_" + useAccessId);
							var $accessRow = $("#SECURITYROW_" + useTypeId + "_" + useAccessId);
							if(useAccessFlag == "1" && $accessInput.prop("checked") == false){
								numChanges++;
								$accessInput.prop("checked", true);
								$accessRow.addClass("new-entry-color");
							}
							else if (useAccessFlag == "0" && $accessInput.prop("checked") == true){
								numChanges++;
								$accessInput.prop("checked", false);
								$accessRow.removeClass("new-entry-color");
							}
						}
						if(numChanges == 1){
							$("#SECURITY_TEXT").html("(1 Change Made)");
						}
						else{
							$("#SECURITY_TEXT").html("(" + numChanges + " Changes Made)");
						}
					},
					error: function(){
						$("#SECURITY_TEXT").html("Profile Failed");
					},
					cache: false
				});
			}
		});
	});
/* User Admin - End */