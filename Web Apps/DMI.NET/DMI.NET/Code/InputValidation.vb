﻿Namespace Code
	Public Class InputValidation

		Public Shared ListOfActions As List(Of String)
		Public Shared ListOfUtilTypes As List(Of String)
		Public Shared ListOfCT_Modes As List(Of String)
		Public Enum WhiteListCollections
			Actions
			UtilTypes
			CT_Modes
		End Enum
		Public Enum StringSanitiseLevel
			None
			HTMLEncode
			FullOWASP
		End Enum
		Public Shared Sub Initialise()
			ListOfActions = New List(Of String)
			With ListOfActions
				.Add("ABSENCEBREAKDOWNALL")
				.Add("ABSENCEBREAKDOWNREC")
				.Add("ADDEXPRCOMPONENT")
				.Add("ADDFROMWAITINGLIST")
				.Add("ADDFROMWAITINGLISTERROR")
				.Add("ADDFROMWAITINGLISTSUCCESS")
				.Add("ALL")
				.Add("ALLRECORDS")
				.Add("BOOKCOURSE")
				.Add("BOOKCOURSEERROR")
				.Add("BOOKCOURSESUCCESS")
				.Add("BRADFORDFACTORALL")
				.Add("BRADFORDFACTORREC")
				.Add("BULKBOOKINGERROR")
				.Add("BULKBOOKINGSUCCESS")
				.Add("CALCULATIONS")
				.Add("CALENDARREPORTS")
				.Add("CALENDARREPORTSREC")
				.Add("CANCEL")
				.Add("CANCELBOOKING")
				.Add("CANCELBOOKING_1")
				.Add("CANCELCOURSE")
				.Add("CANCELCOURSE_1")
				.Add("CANCELCOURSE_2")
				.Add("CLEAR")
				.Add("CLEARFILTER")
				.Add("COPY")
				.Add("CROSSTABS")
				.Add("NINEBOXGRID")
				.Add("CUSTOMREPORTS")
				.Add("DEFAULT")
				.Add("DELETE")
				.Add("EDITEXPRCOMPONENT")
				.Add("EVENTLOG")
				.Add("EXIT")
				.Add("FILTER")
				.Add("FILTERS")
				.Add("FIND")
				.Add("GETBULKBOOKINGSELECTION")
				.Add("GETEXPRESSIONRETURNTYPES")
				.Add("GETPICKLISTSELECTION")
				.Add("INSERTEXPRCOMPONENT")
				.Add("LINK")
				.Add("LINKOLE")
				.Add("LOAD")
				.Add("LOADCALENDARREPORTCOLUMNS")
				.Add("LOADEVENTLOG")
				.Add("LOADEVENTLOGUSERS")
				.Add("LOADEXPRFIELDCOLUMNS")
				.Add("LOADEXPRLOOKUPCOLUMNS")
				.Add("LOADEXPRLOOKUPVALUES")
				.Add("LOADFIND")
				.Add("LOADLOOKUPFIND")
				.Add("LOADTRANSFERCOURSE")
				.Add("LOADBOOKCOURSE")
				.Add("LOADADDFROMWAITINGLIST")
				.Add("LOADTRANSFERBOOKING")
				.Add("LOADREPORTCOLUMNS")
				.Add("LOCATE")
				.Add("LOCATEID")
				.Add("LOGOFF")
				.Add("LOOKUP")
				.Add("MAILMERGE")
				.Add("MOVEFIRST")
				.Add("MOVENEXT")
				.Add("MOVELAST")
				.Add("MOVEPREVIOUS")
				.Add("NEW")
				.Add("PARENT")
				.Add("PICKLIST")
				.Add("QUICKFIND")
				.Add("REFRESHFINDAFTERDELETE")
				.Add("REFRESHFINDAFTERINSERT")
				.Add("RELOAD")
				.Add("SAVE")
				.Add("SAVEERROR")
				.Add("SELECTADDFROMWAITINGLIST_1")
				.Add("SELECTADDFROMWAITINGLIST_2")
				.Add("SELECTADDFROMWAITINGLIST_3")
				.Add("SELECTBOOKCOURSE_1")
				.Add("SELECTBOOKCOURSE_2")
				.Add("SELECTBOOKCOURSE_3")
				.Add("SELECTBULKBOOKINGS")
				.Add("SELECTBULKBOOKINGS_2")
				.Add("SELECTCOMPONENT")
				.Add("SELECTFILTER")
				.Add("SELECTIMAGE")
				.Add("SELECTLINK")
				.Add("SELECTLOOKUP")
				.Add("SELECTOLE")
				.Add("SELECTORDER")
				.Add("SELECTTRANSFERBOOKING_1")
				.Add("SELECTTRANSFERBOOKING_2")
				.Add("SELECTTRANSFERCOURSE")
				.Add("STDRPT_ABSENCECALENDAR")
				.Add("STDREPORT_DATEPROMPT")
				.Add("TRANSFERBOOKING")
				.Add("TRANSFERBOOKINGERROR")
				.Add("TRANSFERBOOKINGSUCCESS")
				.Add("TRANSFERCOURSE")
				.Add("VIEW")
				.Add("WORKFLOW")
				.Add("WORKFLOWOUTOFOFFICE")
				.Add("WORKFLOWPENDINGSTEPS")
			End With

			ListOfUtilTypes = New List(Of String)
			With ListOfUtilTypes
				.Add("UTLBATCHJOB")
				.Add("UTLCROSSTAB")
				.Add("UTLCUSTOMREPORT")
				.Add("UTLDATATRANSFER")
				.Add("UTLEXPORT")
				.Add("UTLGLOBALADD")
				.Add("UTLGLOBALDELETE")
				.Add("UTLGLOBALUPDATE")
				.Add("UTLIMPORT")
				.Add("UTLMAILMERGE")
				.Add("UTLPICKLIST")
				.Add("UTLFILTER")
				.Add("UTLCALCULATION")
				.Add("UTLORDER")
				.Add("UTLMATCHREPORT")
				.Add("UTLABSENCEBREAKDOWN")
				.Add("UTLBRADFORDFACTOR")
				.Add("UTLCALENDARREPORT")
				.Add("UTLLABEL")
				.Add("UTLLABELTYPE")
				.Add("UTLRECORDPROFILE")
				.Add("UTLEMAILADDRESS")
				.Add("UTLEMAILGROUP")
				.Add("UTLSUCCESSION")
				.Add("UTLCAREER")
				.Add("UTLWORKFLOW")
				.Add("UTLWORKFLOWPENDINGSTEPS")
				.Add("UTLORDERDEFINITION")
				.Add("UTLDOCUMENTMAPPING")
				.Add("UTLREPORTPACK")
				.Add("UTLTURNOVER")
				.Add("UTLSTABILITY")
				.Add("UTLSCREEN")
				.Add("UTLTABLE")
				.Add("UTLCOLUMN")
				.Add("UTLNINEBOXGRID")
				.Add("UTLABSENCEBREAKDOWNCONFIGURATION")
        .Add("TALENT")
			End With

			ListOfCT_Modes = New List(Of String)
			With ListOfCT_Modes
				.Add("NONE")
				.Add("BREAKDOWN")
				.Add("LOAD")
				.Add("REFRESH")
				.Add("EMAILGROUP")
				.Add("EMAILGROUPTHENCLOSE")
			End With
		End Sub
	End Class
End Namespace