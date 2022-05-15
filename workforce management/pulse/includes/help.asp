<%	
	If Request.Querystring("TYPE") <> "" Then
		PARAMETER_TYPE = Request.Querystring("TYPE")
	Else
		PARAMETER_TYPE = "SCHEDULE"
	End If
%>
<% If PARAMETER_TYPE = "CONTROL" Then %>
	<div class="white-background" style="border-radius:0.3em;">
		<h3 style="text-align:center;">Schedule Control Notes:</h3>
		<ul style="list-style-position: inside;">
			<li>Eff/Dis dates refer to dates in scheduling applications like Schedule Center, not the current date.</li>
			<li>For weekly parameters (UNP and OT Limits), make sure the eff/dis date are in full weeks (Sunday-Saturday).</li>
			<li>Intervals refer to the start/end times, respectively, so, for example, 06:00 &ndash; 24:00 covers time between 6 AM and midnight.</li>
			<li>For scheduled hours a range of 30 to 32, for example, means: 30 &lt; hours &le; 32.</li>
			<li>Conversely, for eval score a range of 3.25 to 5, for example, means: 3.25 &le; eval score &lt; 5.</li>
			<li>To delete an entry, set the effective date to be greater than the discontinue date.</li>
			<li>When making tweaks, you can modify existing records or set an appropriate discontinue date and create new records. I would suggest the latter, as it helps keep a good record of what occurred.</li>
		</ul>
	</div>
<% Elseif PARAMETER_TYPE = "ADMIN" Then %>
	<div class="white-background" style="border-radius:0.3em;">
		<h3 style="text-align:center;">Admin Notes:</h3>
		<ul style="list-style-position: inside;">
			<li style="font-weight:bold;">Details Notes</li>
			<ul>
				<li>Clicking on the Employee name field opens/closes the details.</li>
				<li>In the details Setting the End Date before the Start Date deletes the entry.</li>
				<li>When adding new users, the start date in the details is whatever date Pulse is set to &mdash; setting Pulse to the hire date helps streamline adding new users.</li>
			</ul>
			<li style="font-weight:bold;">Security Access Notes</li>
			<ul>
				<li>Access to departments, applications, pages, and reports can be found here.</li>
				<li>Click anywhere on the desired row to give/remove access (green background &#61; access granted).</li>
				<li>The icon in the top right corner attempts to generate a security profile based on the department-team-job combination of the selected employee on the selected date in Pulse.</li>
				<li>There are two submit buttons, one on top and one on bottom.</li>
			</ul>
		</ul>
	</div>
<% Elseif PARAMETER_TYPE = "CIRCLE" Then %>
	<div class="white-background" style="border-radius:0.3em;">
		<h3 style="text-align:center;">Circle Stat Notes:</h3>
		<ul style="list-style-position: inside;">
			<li>Only the green and yellow ranges need to be set with this tool &mdash; anything that falls outside of this is coded as red.</li>
			<li>To delete an entry, set the effective date to be greater than the discontinue date.</li>
			<li>When making tweaks, you can modify existing records or set an appropriate discontinue date and create new records. I would suggest the latter, as it helps keep a good record of what occurred.</li>
		</ul>
	</div>
<% End If %>  