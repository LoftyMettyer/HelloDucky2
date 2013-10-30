<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>

<form name="frmOptions" id="frmOptions">
		<table align="center" cellpadding="0" cellspacing="0" width="100%" height="100%">
				<tr>
						<td>
								<table class="invisible" cellspacing="0" cellpadding="2" width="100%" height="100%">
										<tr height="5" valign="top">
												<td width="5"></td>
												<td height="100%" width="100%" align="left"><strong>Options :</strong>
														<br>
														<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																<tr>
																		<td width="5"></td>
																		<td height="1">
																				<%
																						Dim objCalendar As HR.Intranet.Server.CalendarReport
	
																						objCalendar = Session("objCalendar" & Session("CalRepUtilID"))

																						If objCalendar.IncludeBankHolidays_Enabled And objCalendar.IncludeBankHolidays Then
																				%>
																				<input name="chkIncludeBHols" id="chkIncludeBHols" type="checkbox" checked tabindex="0"
																						onclick="refreshInfo();" />
																				<label
																						for="chkIncludeBHols"
																						class="checkbox"
																						tabindex="-1"/>
																				<%
																				ElseIf objCalendar.IncludeBankHolidays_Enabled And objCalendar.IncludeBankHolidays = False Then
																				%>
																				<input name="chkIncludeBHols" id="chkIncludeBHols" type="checkbox" tabindex="0"
																						onclick="refreshInfo();" />
																				<label
																						for="chkIncludeBHols"
																						class="checkbox"
																						tabindex="-1" />
																				<%
																				Else
																				%>
																				<input name="chkIncludeBHols" id="chkIncludeBHols" type="checkbox" disabled="disabled" tabindex="0"
																						onclick="refreshInfo();" />
																				<label
																						for="chkIncludeBHols"
																						class="checkbox checkboxdisabled"
																						tabindex="-1" />
																				<%
																				End If
																				%>
											Include Bank Holidays 
																		
																		</td>
																		<td width="5"></td>
																</tr>
																<tr>
																		<td width="5"></td>
																		<td height="1">
																				<%
																						If objCalendar.IncludeWorkingDaysOnly_Enabled And objCalendar.IncludeWorkingDaysOnly Then
																				%>
																				<input name="chkIncludeWorkingDaysOnly" id="chkIncludeWorkingDaysOnly" type="checkbox" checked tabindex="0"
																						onclick="refreshInfo();" />
																				<label
																						for="chkIncludeWorkingDaysOnly"
																						class="checkbox"
																						tabindex="-1" />
																				<%
																				ElseIf objCalendar.IncludeWorkingDaysOnly_Enabled And objCalendar.IncludeWorkingDaysOnly = False Then
																				%>
																				<input name="chkIncludeWorkingDaysOnly" id="chkIncludeWorkingDaysOnly" type="checkbox" tabindex="0"
																						onclick="refreshInfo();" />
																				<label
																						for="chkIncludeWorkingDaysOnly"
																						class="checkbox"
																						tabindex="-1" />
																				<%
																				Else
																				%>
																				<input name="chkIncludeWorkingDaysOnly" id="chkIncludeWorkingDaysOnly" type="checkbox" disabled="disabled" tabindex="0"
																						onclick="refreshInfo();" />
																				<label
																						for="chkIncludeWorkingDaysOnly"
																						class="checkbox checkboxdisabled"
																						tabindex="-1" />
																				<%
																				End If
																				%>
											Working Days Only 
																		
																		</td>
																		<td width="5"></td>
																</tr>
																<tr>
																		<td width="5"></td>
																		<td height="1">
																				<%
																						If objCalendar.ShowBankHolidays_Enabled And objCalendar.ShowBankHolidays Then
																				%>

																				<input name="chkShadeBHols" id="chkShadeBHols" type="checkbox" checked tabindex="0"
																						onclick="refreshInfo();" />
																				<label
																						for="chkShadeBHols"
																						class="checkbox"
																						tabindex="-1" />
																				<%
																				ElseIf objCalendar.ShowBankHolidays_Enabled And objCalendar.ShowBankHolidays = False Then
																				%>
																				<input name="chkShadeBHols" id="chkShadeBHols" type="checkbox" tabindex="0"
																						onclick="refreshInfo();" />
																				<label
																						for="chkShadeBHols"
																						class="checkbox"
																						tabindex="-1" />
																				<%
																				Else
																				%>
																				<input name="chkShadeBHols" id="chkShadeBHols" type="checkbox" disabled="disabled" tabindex="0"
																						onclick="refreshInfo();" />
																				<label
																						for="chkShadeBHols"
																						class="checkbox checkboxdisabled"
																						tabindex="-1" />
																				<%
																				End If
																				%>
											Show Bank Holidays 
																		
																		</td>
																		<td width="5"></td>
																</tr>
																<tr>
																		<td width="5"></td>
																		<td height="1" nowrap>
																				<%
																						If objCalendar.ShowCaptions Then
																				%>
																				<input name="chkCaptions" id="chkCaptions" type="checkbox" checked tabindex="0"
																						onclick="refreshInfo();" />
																				<%
																				Else
																				%>
																				<input name="chkCaptions" id="chkCaptions" type="checkbox" tabindex="0"
																						onclick="refreshInfo();" />
																				<%
																				End If
																				%>
																				<label
																						for="chkCaptions"
																						class="checkbox"
																						tabindex="-1">
																						Show Calendar Captions 
																				</label>
																		</td>
																		<td width="5"></td>
																</tr>
																<tr>
																		<td width="5"></td>
																		<td height="1">
																				<%
																						If objCalendar.ShowWeekends Then
																				%>
																				<input name="chkShadeWeekends" id="chkShadeWeekends" type="checkbox" checked tabindex="0"
																						onclick="refreshInfo();" />

																				<%
																				Else
																				%>
																				<input name="chkShadeWeekends" id="chkShadeWeekends" type="checkbox" tabindex="0"
																						onclick="refreshInfo();" />

																				<%
																				End If
	
																				objCalendar = Nothing
																				%>
																				<label
																						for="chkShadeWeekends"
																						class="checkbox"
																						tabindex="-1">
																						Show Weekends 
																				</label>
																		</td>
																		<td width="5"></td>
																</tr>
														</table>
												</td>
												<td width="5"></td>
										</tr>
								</table>
						</td>
				</tr>
		</table>

		<input type="hidden" id="txtCalRep_UtilID" name="txtCalRep_UtilID" value='<%Session("CalRepUtilID").ToString()%>'>
</form>
