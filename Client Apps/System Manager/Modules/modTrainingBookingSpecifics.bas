Attribute VB_Name = "modTrainingBookingSpecifics"
Option Explicit

Private mvar_fPreReqsUsed As Boolean
Private mvar_fUnavailabilityUsed As Boolean

Private mvar_fGeneralOK As Boolean
Private mvar_fPreReqsOK As Boolean
Private mvar_fUnavailOK As Boolean
Private mvar_fOverlapOK As Boolean
Private mvar_fOverbookOK As Boolean
Private mvar_fCancelDateOK As Boolean

Private mvar_sGeneralMsg As String
Private mvar_sPreReqsMsg As String
Private mvar_sPreReqsWarningMsg As String
Private mvar_sUnavailMsg As String
Private mvar_sUnavailWarningMsg As String
Private mvar_sCancelDateMsg As String

' Training Course table variables.
Private mvar_lngCourseTableID As Long
Private mvar_sCourseTableName As String
Private mvar_lngCourseTitleID As Long
Private mvar_sCourseTitleName As String
Private mvar_lngCourseStartDateID As Long
Private mvar_sCourseStartDateColumnName As String
Private mvar_lngCourseEndDateID As Long
Private mvar_sCourseEndDateColumnName As String
Private mvar_fCourseIncludeProvisionals As Boolean
'Private mvar_lngCourseNumberBookedID As Long
'Private mvar_sCourseNumberBookedName As String
Private mvar_lngCourseMaxNumberID As Long
Private mvar_sCourseMaxNumberName As String
Private mvar_iCourseOverbookingNotification As Integer
Private mvar_lngCancelCourseID As Long
Private mvar_sCancelCourseColumnName As String

' Pre-requisite table variables.
Private mvar_lngPreReqTableID As Long
Private mvar_sPreReqTableName As String
Private mvar_lngPreReqCourseTitleID As Long
Private mvar_sPreReqCourseTitleName As String
Private mvar_lngPreReqGroupingID As Long
Private mvar_sPreReqGroupingName As String
Private mvar_lngPreReqFailureNotificationID As Long
Private mvar_sPreReqFailureNotificationName As String

' Training Bookings table variables.
Private mvar_lngTrainBookTableID As Long
Private mvar_sTrainBookTableName As String
Private mvar_lngTrainBookStatusID As Long
Private mvar_sTrainBookStatusName As String
Private mvar_iTrainBookOverlapNotification As Integer

' Employee table variables.
Private mvar_lngEmployeeTableID As Long
'Private mvar_lngBulkBookingDefaultViewID As Long

' Unavailability table variables.
Private mvar_lngUnavailTableID As Long
Private mvar_sUnavailTableName As String
Private mvar_lngUnavailFromDateID As Long
Private mvar_sUnavailFromDateName As String
Private mvar_lngUnavailToDateID As Long
Private mvar_sUnavailToDateName As String
Private mvar_lngUnavailFailureNotificationID As Long
Private mvar_sUnavailFailureNotificationName As String

' Constants
Private Const mvar_sPreReqProcedureName = "sp_ASR_TBCheckPreRequisites"
Private Const mvar_sUnavailProcedureName = "sp_ASR_TBCheckUnavailability"
Private Const mvar_sOverbookingProcedureName = "sp_ASR_TBCheckOverbooking"
Private Const mvar_sOverlapProcedureName = "sp_ASR_TBCheckOverlappedBooking"
Private Const mvar_sCourseCancelledCheck = "spASRIntGetCancelCourseDate"

Public Function ConfigureTrainingBookingSpecifics() As Boolean
  ' Configure module specific objects (eg. stored procedures)
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim sErrorMessage As String
  
  fOK = True
  
  mvar_fGeneralOK = True
  mvar_fPreReqsOK = True
  mvar_fUnavailOK = True
  mvar_fOverlapOK = True
  mvar_fOverbookOK = True
  mvar_fCancelDateOK = True
    
  mvar_sGeneralMsg = ""
  mvar_sPreReqsMsg = ""
  mvar_sPreReqsWarningMsg = ""
  mvar_sUnavailMsg = ""
  mvar_sUnavailWarningMsg = ""
  mvar_sCancelDateMsg = ""
  
  ' Read the Course table parameters.
  fOK = ReadCourseRecordParameters
  
  ' Read the Pre-requisite table parameters.
  If fOK Then
    fOK = ReadPreRequisiteParameters
  End If
  
  ' Read the Training Bookings table parameters.
  If fOK Then
    fOK = ReadTrainingBookingParameters
  End If
  
  ' Read the Employee table parameters.
  If fOK Then
    fOK = ReadEmployeeRecordParameters
  End If
  
  ' Read the Unavailability table parameters.
  If fOK Then
    fOK = ReadUnavailabilityParameters
  End If
  
  If fOK Then
    sErrorMessage = ""
    If (Not mvar_fGeneralOK) Or _
      (Not mvar_fOverbookOK) Or _
      (Not mvar_fOverlapOK) Then
      sErrorMessage = "Training Booking specifics not correctly configured." & vbNewLine & _
        "Some functionality will be disabled if you do not change your configuration." & mvar_sGeneralMsg
    Else
      If mvar_fPreReqsUsed Then
        If Not mvar_fPreReqsOK Then
          sErrorMessage = "Training Booking specifics not correctly configured." & vbNewLine & _
            "Pre-requisite checks will be disabled if you do not change your configuration." & mvar_sPreReqsMsg
        ElseIf Len(mvar_sPreReqsWarningMsg) > 0 Then
          sErrorMessage = "Training Booking specifics not correctly configured." & vbNewLine & _
            mvar_sPreReqsWarningMsg & vbNewLine & _
            "Pre-requisite failures will be notified as warnings if you do not change your configuration."
        End If
      End If
    
      If mvar_fUnavailabilityUsed Then
        If Not mvar_fUnavailOK Then
          sErrorMessage = sErrorMessage & _
            IIf(Len(sErrorMessage) > 0, "", "Training Booking specifics not correctly configured.") & vbNewLine & _
            "Unavailability checks will be disabled if you do not change your configuration." & mvar_sUnavailMsg
        ElseIf Len(mvar_sUnavailWarningMsg) > 0 Then
          sErrorMessage = sErrorMessage & _
            IIf(Len(sErrorMessage) > 0, "", "Training Booking specifics not correctly configured.") & vbNewLine & _
            mvar_sUnavailWarningMsg & vbNewLine & _
            "Unavailability failures will be notified as warnings if you do not change your configuration."
        End If
      End If
    End If
    
    If Len(sErrorMessage) > 0 Then
      fOK = (OutputMessage(sErrorMessage & vbNewLine & vbNewLine & "Continue saving changes ?") = vbYes)
    End If
  End If
  
  ' Create the course pre-requisite check stored procedure.
  If fOK Then
'    sSQL = "IF EXISTS" & _
'      " (SELECT Name" & _
'      "   FROM sysobjects" & _
'      "   WHERE id = object_id('" & mvar_sPreReqProcedureName & "')" & _
'      "     AND sysstat & 0xf = 4)" & _
'      " DROP PROCEDURE " & mvar_sPreReqProcedureName
'    gADOCon.Execute sSQL, , adExecuteNoRecords
    DropProcedure mvar_sPreReqProcedureName

    If mvar_fPreReqsUsed And mvar_fPreReqsOK Then
      fOK = CreatePreRequisiteCheckStoredProcedure
    End If
  End If
  
  ' Create the course cancelled date check procedure
  If fOK Then
    DropProcedure mvar_sCourseCancelledCheck

    'If mvar_fPreReqsUsed And mvar_fPreReqsOK Then
      fOK = CreateCourseCancelDateCheckProcedure
    'End If
  End If
  
  ' Create the Unavailability check stored procedure.
  If fOK Then
'    sSQL = "IF EXISTS" & _
'      " (SELECT Name" & _
'      "   FROM sysobjects" & _
'      "   WHERE id = object_id('" & mvar_sUnavailProcedureName & "')" & _
'      "     AND sysstat & 0xf = 4)" & _
'      " DROP PROCEDURE " & mvar_sUnavailProcedureName
'    gADOCon.Execute sSQL, , adExecuteNoRecords
    DropProcedure mvar_sUnavailProcedureName
  
    If mvar_fUnavailabilityUsed And mvar_fUnavailOK Then
      fOK = CreateUnavailabilityCheckStoredProcedure
    End If
  End If
  
  ' Create the Overbooking check stored procedure.
  If fOK Then
'    sSQL = "IF EXISTS" & _
'      " (SELECT Name" & _
'      "   FROM sysobjects" & _
'      "   WHERE id = object_id('" & mvar_sOverbookingProcedureName & "')" & _
'      "     AND sysstat & 0xf = 4)" & _
'      " DROP PROCEDURE " & mvar_sOverbookingProcedureName
'    gADOCon.Execute sSQL, , adExecuteNoRecords
    DropProcedure mvar_sOverbookingProcedureName
  
    If mvar_fOverbookOK Then
      fOK = CreateOverbookingCheckStoredProcedure
    End If
  End If
  
  ' Create the Overlapped Booking check stored procedure.
  If fOK Then
'    sSQL = "IF EXISTS" & _
'      " (SELECT Name" & _
'      "   FROM sysobjects" & _
'      "   WHERE id = object_id('" & mvar_sOverlapProcedureName & "')" & _
'      "     AND sysstat & 0xf = 4)" & _
'      " DROP PROCEDURE " & mvar_sOverlapProcedureName
'    gADOCon.Execute sSQL, , adExecuteNoRecords
    DropProcedure mvar_sOverlapProcedureName
    
    If mvar_fOverlapOK Then
      fOK = CreateOverlappedBookingCheckStoredProcedure
    End If
  End If
  
TidyUpAndExit:
  ConfigureTrainingBookingSpecifics = fOK
  Exit Function
  
ErrorTrap:
  OutputError "Error configuring Training Booking specifics"
  fOK = False
  Resume TidyUpAndExit

End Function


Private Function ReadTrainingBookingParameters() As Boolean
  ' Read the Training Booking parameter values from the database into local variables.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fStatusFound_B As Boolean
  Dim fStatusFound_C As Boolean
  Dim sErrMsg As String
  
  With recModuleSetup
    .Index = "idxModuleParameter"

    ' Get the Training Booking table ID and name.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_TRAINBOOKTABLE
    fOK = Not .NoMatch
    If fOK Then
      'MH20001025 Fault 799 Type Mismatch
      'fOK = Not IsNull(!parametervalue)
      fOK = (IIf(IsNull(!parametervalue), 0, val(!parametervalue)) > 0)
    End If
    If Not fOK Then
      mvar_fPreReqsOK = False
      mvar_fOverlapOK = False
      mvar_fOverbookOK = False
      
      mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  'Training Bookings' table not defined."
    Else
      mvar_lngTrainBookTableID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))

      With recTabEdit
        .Index = "idxTableID"
        .Seek "=", mvar_lngTrainBookTableID
      
        fOK = Not .NoMatch
        If fOK Then
          fOK = Not IsNull(!TableName)
        End If
        If Not fOK Then
          mvar_fPreReqsOK = False
          mvar_fOverlapOK = False
          mvar_fOverbookOK = False
      
          mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  'Training Bookings' table not found."
        Else
          mvar_sTrainBookTableName = !TableName
        End If
      End With
    End If

    If mvar_fPreReqsOK Or mvar_fOverlapOK Or mvar_fOverbookOK Then
      ' Get the Training Booking Status column ID.
      .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_TRAINBOOKSTATUS
      fOK = Not .NoMatch
      If fOK Then
        'MH20001025 Fault 799 Type Mismatch
        'fOK = Not IsNull(!parametervalue)
        fOK = (IIf(IsNull(!parametervalue), 0, val(!parametervalue)) > 0)
      End If
      If Not fOK Then
        mvar_fPreReqsOK = False
        mvar_fOverlapOK = False
        mvar_fOverbookOK = False
        
        mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  '" & mvar_sTrainBookTableName & "' table 'Status' column not defined."
      Else
        mvar_lngTrainBookStatusID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
      
        With recColEdit
          .Index = "idxColumnID"
          .Seek "=", mvar_lngTrainBookStatusID
        
          fOK = Not .NoMatch
          
          If fOK Then
            fOK = Not IsNull(!ColumnName)
          End If
          
          If Not fOK Then
            mvar_fPreReqsOK = False
            mvar_fOverlapOK = False
            mvar_fOverbookOK = False
            
            mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  '" & mvar_sTrainBookTableName & "' table 'Status' column not found."
          Else
            mvar_sTrainBookStatusName = !ColumnName
            
            ' Check that if the Status column is a text column with option/combo entries, all of
            ' the required status values are in the option group/combo.
            If (!columntype <> giCOLUMNTYPE_LOOKUP) And _
              ((!ControlType = giCTRL_COMBOBOX) Or (!ControlType = giCTRL_OPTIONGROUP)) Then
    
              fStatusFound_B = False
              fStatusFound_C = False
              
              recContValEdit.Index = "idxColumnID"
              recContValEdit.Seek ">=", !ColumnID
    
              If Not recContValEdit.NoMatch Then
                Do While Not recContValEdit.EOF
                  If recContValEdit!ColumnID <> !ColumnID Then
                    Exit Do
                  End If
        
                  If Len(Trim(recContValEdit!value)) > 0 Then
                    Select Case UCase(Trim(recContValEdit!value))
                      Case "B"
                        fStatusFound_B = True
                      Case "C"
                        fStatusFound_C = True
                    End Select
                  End If
        
                  recContValEdit.MoveNext
                Loop
              End If
              
              sErrMsg = ""
              If Not fStatusFound_B Then
                sErrMsg = "'B'"
              End If
              If Not fStatusFound_C Then
                sErrMsg = sErrMsg & IIf(Len(sErrMsg) > 0, "& ", "") & "'C'"
              End If
                  
              If Len(sErrMsg) > 0 Then
                mvar_fGeneralOK = False
                mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  '" & mvar_sTrainBookTableName & "' table '" & mvar_sTrainBookStatusName & "' column does not have the required control values : " & sErrMsg
                fOK = False
              End If
            End If
          End If
        End With
      End If
    End If

    ' Get the Training Booking Overlap Notification flag.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_TRAINBOOKOVERLAPNOTIFICATION
    If .NoMatch Then
      mvar_iTrainBookOverlapNotification = 0
    Else
      'MH20001025 Fault 799
      'mvar_iTrainBookOverlapNotification = IIf(IsNull(!parametervalue), 0, !parametervalue)
      mvar_iTrainBookOverlapNotification = val(IIf(IsNull(!parametervalue), 0, !parametervalue))
    End If
  
  End With

  fOK = True
  
TidyUpAndExit:
  ReadTrainingBookingParameters = fOK
  Exit Function
  
ErrorTrap:
  OutputError "Error reading training booking record parameters (Training Booking)"
  fOK = False
  Resume TidyUpAndExit

End Function





Private Function ReadPreRequisiteParameters() As Boolean
  ' Read the Pre-requisite parameter values from the database into local variables.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fNotificationFound_Error As Boolean
  Dim sErrMsg As String
  
  fOK = True
  
  With recModuleSetup
    .Index = "idxModuleParameter"
        
    ' Get the Pre-requisite table ID and name.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_PREREQTABLE
    mvar_fPreReqsUsed = Not .NoMatch
    
    If mvar_fPreReqsUsed Then
      mvar_fPreReqsUsed = Not IsNull(!parametervalue)
    End If
    
    If mvar_fPreReqsUsed Then
      mvar_fPreReqsUsed = (val(!parametervalue) > 0)
    End If
    
    If mvar_fPreReqsUsed Then
      mvar_lngPreReqTableID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
  
      With recTabEdit
        .Index = "idxTableID"
        .Seek "=", mvar_lngPreReqTableID
      
        fOK = Not .NoMatch
        If fOK Then
          fOK = Not IsNull(!TableName)
        End If
        If Not fOK Then
          mvar_fPreReqsOK = False
          mvar_sPreReqsMsg = mvar_sPreReqsMsg & vbNewLine & "  'Course Pre-requisites' table not found."
        Else
          mvar_sPreReqTableName = !TableName
        End If
      End With
  
      If mvar_fPreReqsOK Then
        ' Get the Pre-requisite Course Title column ID.
        .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_PREREQCOURSETITLE
        fOK = Not .NoMatch
        If fOK Then
          'MH20001025 Fault 799 Type Mismatch
          'fOK = Not IsNull(!parametervalue)
          fOK = (IIf(IsNull(!parametervalue), 0, val(!parametervalue)) > 0)
        End If
        If Not fOK Then
          mvar_fPreReqsOK = False
          
          mvar_sPreReqsMsg = mvar_sPreReqsMsg & vbNewLine & "  '" & mvar_sPreReqTableName & "' table 'Course Title' column not defined."
        Else
          mvar_lngPreReqCourseTitleID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
        
          With recColEdit
            .Index = "idxColumnID"
            .Seek "=", mvar_lngPreReqCourseTitleID
          
            fOK = Not .NoMatch
            If fOK Then
              fOK = Not IsNull(!ColumnName)
            End If
            If Not fOK Then
              mvar_fPreReqsOK = False
              
              mvar_sPreReqsMsg = mvar_sPreReqsMsg & vbNewLine & "  '" & mvar_sPreReqTableName & "' table 'Course Title' column not found."
            Else
              mvar_sPreReqCourseTitleName = !ColumnName
            End If
          End With
        End If
      End If
  
      If mvar_fPreReqsOK Then
        ' Get the Pre-requisite Grouping column ID.
        .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_PREREQGROUPING
        fOK = Not .NoMatch
        If fOK Then
          'MH20001025 Fault 799 Type Mismatch
          'fOK = Not IsNull(!parametervalue)
          fOK = (IIf(IsNull(!parametervalue), 0, val(!parametervalue)) > 0)
        End If
        If Not fOK Then
          mvar_fPreReqsOK = False
          
          mvar_sPreReqsMsg = mvar_sPreReqsMsg & vbNewLine & "  '" & mvar_sPreReqTableName & "' table 'Grouping' column not defined."
        Else
          mvar_lngPreReqGroupingID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
        
          With recColEdit
            .Index = "idxColumnID"
            .Seek "=", mvar_lngPreReqGroupingID
          
            fOK = Not .NoMatch
            If fOK Then
              fOK = Not IsNull(!ColumnName)
            End If
            If Not fOK Then
              mvar_fPreReqsOK = False
              
              mvar_sPreReqsMsg = mvar_sPreReqsMsg & vbNewLine & "  '" & mvar_sPreReqTableName & "' table 'Grouping' column not found."
            Else
              mvar_sPreReqGroupingName = !ColumnName
            End If
          End With
        End If
      End If
  
      If mvar_fPreReqsOK Then
        ' Get the Pre-requisite Failure Notification column ID.
        .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_PREREQFAILURE
        fOK = Not .NoMatch
        If fOK Then
          'MH20001025 Fault 799 Type Mismatch
          'fOK = Not IsNull(!parametervalue)
          fOK = (IIf(IsNull(!parametervalue), 0, val(!parametervalue)) > 0)
        End If
        If Not fOK Then
          mvar_fPreReqsOK = False
          
          mvar_sPreReqsMsg = mvar_sPreReqsMsg & vbNewLine & "  '" & mvar_sPreReqTableName & "' table 'Failure Notification' column not defined."
        Else
          mvar_lngPreReqFailureNotificationID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
        
          With recColEdit
            .Index = "idxColumnID"
            .Seek "=", mvar_lngPreReqFailureNotificationID
          
            fOK = Not .NoMatch
            If fOK Then
              fOK = Not IsNull(!ColumnName)
            End If
            If Not fOK Then
              mvar_fPreReqsOK = False
              
              mvar_sPreReqsMsg = mvar_sPreReqsMsg & vbNewLine & "  '" & mvar_sPreReqTableName & "' table 'Failure Notification' column not found."
            Else
              mvar_sPreReqFailureNotificationName = !ColumnName
            
              ' Check that if the Failure Notification column is a text column with option/combo entries, all of
              ' the required values are in the option group/combo.
              If (!columntype <> giCOLUMNTYPE_LOOKUP) And _
                ((!ControlType = giCTRL_COMBOBOX) Or (!ControlType = giCTRL_OPTIONGROUP)) Then
    
                fNotificationFound_Error = False
              
                recContValEdit.Index = "idxColumnID"
                recContValEdit.Seek ">=", !ColumnID
    
                If Not recContValEdit.NoMatch Then
                  Do While Not recContValEdit.EOF
                    If recContValEdit!ColumnID <> !ColumnID Then
                      Exit Do
                    End If
        
                    If Len(Trim(recContValEdit!value)) > 0 Then
                      If UCase(Trim(recContValEdit!value)) = "ERROR" Then
                        fNotificationFound_Error = True
                        Exit Do
                      End If
                    End If
        
                    recContValEdit.MoveNext
                  Loop
                End If
              
                sErrMsg = ""
                If Not fNotificationFound_Error Then
                  sErrMsg = "'Error'"
                End If
                  
                If Len(sErrMsg) > 0 Then
                  mvar_sPreReqsWarningMsg = "'" & mvar_sPreReqTableName & "' table '" & mvar_sPreReqFailureNotificationName & "' column does not have the required control values : " & sErrMsg
                End If
              End If
            End If
          End With
        End If
      End If
    End If
  End With

  fOK = True
  
TidyUpAndExit:
  ReadPreRequisiteParameters = fOK
  Exit Function
  
ErrorTrap:
  OutputError "Error reading pre-requisite record parameters (Training Booking)"
  fOK = False
  Resume TidyUpAndExit

End Function
Private Function ReadUnavailabilityParameters() As Boolean
  ' Read the Pre-requisite parameter values from the database into local variables.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fNotificationFound_Error As Boolean
  Dim sErrMsg As String
  
  fOK = True
  
  With recModuleSetup
    .Index = "idxModuleParameter"
        
    ' Get the Unavailability table ID and name.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_UNAVAILTABLE
    mvar_fUnavailabilityUsed = Not .NoMatch
    
    If mvar_fUnavailabilityUsed Then
      mvar_fUnavailabilityUsed = Not IsNull(!parametervalue)
    End If
    
    If mvar_fUnavailabilityUsed Then
      mvar_fUnavailabilityUsed = (val(!parametervalue) > 0)
    End If
    
    If mvar_fUnavailabilityUsed Then
      mvar_lngUnavailTableID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
  
      With recTabEdit
        .Index = "idxTableID"
        .Seek "=", mvar_lngUnavailTableID
      
        fOK = Not .NoMatch
        If fOK Then
          fOK = Not IsNull(!TableName)
        End If
        If Not fOK Then
          mvar_fUnavailOK = False
          
          mvar_sUnavailMsg = mvar_sUnavailMsg & vbNewLine & "'Unavailability' table not found."
        Else
          mvar_sUnavailTableName = !TableName
        End If
      End With
  
      If mvar_fUnavailOK Then
        ' Get the Unavailability From Date column ID.
        .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_UNAVAILFROMDATE
        fOK = Not .NoMatch
        If fOK Then
          'MH20001025 Fault 799 Type Mismatch
          'fOK = Not IsNull(!parametervalue)
          fOK = (IIf(IsNull(!parametervalue), 0, val(!parametervalue)) > 0)
        End If
        If Not fOK Then
          mvar_fUnavailOK = False
          mvar_sUnavailMsg = mvar_sUnavailMsg & vbNewLine & "'" & mvar_sUnavailTableName & "' table 'From Date' column not defined."
        Else
          mvar_lngUnavailFromDateID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
        
          With recColEdit
            .Index = "idxColumnID"
            .Seek "=", mvar_lngUnavailFromDateID
          
            fOK = Not .NoMatch
            If fOK Then
              fOK = Not IsNull(!ColumnName)
            End If
            If Not fOK Then
              mvar_fUnavailOK = False
              mvar_sUnavailMsg = mvar_sUnavailMsg & vbNewLine & "'" & mvar_sUnavailTableName & "' table 'From Date' column not found."
            Else
              mvar_sUnavailFromDateName = !ColumnName
            End If
          End With
        End If
      End If

      If mvar_fUnavailOK Then
        ' Get the Unavailability To Date column ID.
        .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_UNAVAILTODATE
        fOK = Not .NoMatch
        If fOK Then
          'MH20001025 Fault 799 Type Mismatch
          'fOK = Not IsNull(!parametervalue)
          fOK = (IIf(IsNull(!parametervalue), 0, val(!parametervalue)) > 0)
        End If
        If Not fOK Then
          mvar_fUnavailOK = False
          mvar_sUnavailMsg = mvar_sUnavailMsg & vbNewLine & "'" & mvar_sUnavailTableName & "' table 'To Date' column not defined."
        Else
          mvar_lngUnavailToDateID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
        
          With recColEdit
            .Index = "idxColumnID"
            .Seek "=", mvar_lngUnavailToDateID
          
            fOK = Not .NoMatch
            If fOK Then
              fOK = Not IsNull(!ColumnName)
            End If
            If Not fOK Then
              mvar_fUnavailOK = False
              mvar_sUnavailMsg = mvar_sUnavailMsg & vbNewLine & "'" & mvar_sUnavailTableName & "' table 'To Date' column not found."
            Else
              mvar_sUnavailToDateName = !ColumnName
            End If
          End With
        End If
      End If

      If mvar_fUnavailOK Then
        ' Get the Unavailability Failure Notification column ID.
        .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_UNAVAILFAILURE
        fOK = Not .NoMatch
        If fOK Then
          'MH20001025 Fault 799 Type Mismatch
          'fOK = Not IsNull(!parametervalue)
          fOK = (IIf(IsNull(!parametervalue), 0, val(!parametervalue)) > 0)
        End If
        If Not fOK Then
          mvar_fUnavailOK = False
          mvar_sUnavailMsg = mvar_sUnavailMsg & vbNewLine & "'" & mvar_sUnavailTableName & "' table 'Failure Notification' column not defined."
        Else
          mvar_lngUnavailFailureNotificationID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
        
          With recColEdit
            .Index = "idxColumnID"
            .Seek "=", mvar_lngUnavailFailureNotificationID
          
            fOK = Not .NoMatch
            If fOK Then
              fOK = Not IsNull(!ColumnName)
            End If
            If Not fOK Then
              mvar_fUnavailOK = False
              mvar_sUnavailMsg = mvar_sUnavailMsg & vbNewLine & "'" & mvar_sUnavailTableName & "' table 'Failure Notification' column not found."
            Else
              mvar_sUnavailFailureNotificationName = !ColumnName
            
              ' Check that if the Failure Notification column is a text column with option/combo entries, all of
              ' the required values are in the option group/combo.
              If (!columntype <> giCOLUMNTYPE_LOOKUP) And _
                ((!ControlType = giCTRL_COMBOBOX) Or (!ControlType = giCTRL_OPTIONGROUP)) Then
    
                fNotificationFound_Error = False
              
                recContValEdit.Index = "idxColumnID"
                recContValEdit.Seek ">=", !ColumnID
    
                If Not recContValEdit.NoMatch Then
                  Do While Not recContValEdit.EOF
                    If recContValEdit!ColumnID <> !ColumnID Then
                      Exit Do
                    End If
        
                    If Len(Trim(recContValEdit!value)) > 0 Then
                      If UCase(Trim(recContValEdit!value)) = "ERROR" Then
                        fNotificationFound_Error = True
                        Exit Do
                      End If
                    End If
        
                    recContValEdit.MoveNext
                  Loop
                End If
              
                sErrMsg = ""
                If Not fNotificationFound_Error Then
                  sErrMsg = "'Error'"
                End If
                  
                If Len(sErrMsg) > 0 Then
                  mvar_sUnavailWarningMsg = "'" & mvar_sUnavailTableName & "' table '" & mvar_sUnavailFailureNotificationName & "' column does not have the required control values : " & sErrMsg
                End If
              End If
            End If
          End With
        End If
      End If
    End If
  End With

  fOK = True
  
TidyUpAndExit:
  ReadUnavailabilityParameters = fOK
  Exit Function
  
ErrorTrap:
  OutputError "Error reading unavailability record parameters (Training Booking)"
  fOK = False
  Resume TidyUpAndExit

End Function

Private Function CreatePreRequisiteCheckStoredProcedure() As Boolean
  ' Create the stored procedure for checking if a delegate has satisfied
  ' the pre-requisite criteria for a course.
  On Error GoTo ErrorTrap
  
  Dim fCreatedOK As Boolean
  Dim sProcSQL As String
  
  fCreatedOK = True
  
  If mvar_fPreReqsUsed Then
    ' Construct the stored procedure creation string (if required).
    sProcSQL = "/* ------------------------------------------------ */" & vbNewLine & _
      "/* Training Booking module stored procedure. */" & vbNewLine & _
      "/* Automatically generated by the System manager.   */" & vbNewLine & _
      "/* ------------------------------------------------ */" & vbNewLine & _
      "CREATE PROCEDURE dbo." & mvar_sPreReqProcedureName & " (" & vbNewLine & _
      "    @plngCourseRecordID int," & vbNewLine & _
      "    @plngEmployeeRecordID int," & vbNewLine & _
      "    @piPreReqsMet int OUTPUT" & vbNewLine & _
      ")" & vbNewLine & _
      "AS" & vbNewLine & _
      "BEGIN"
      
    sProcSQL = sProcSQL & vbNewLine & _
      "    /* Return 0 if the given record in the personnel table has satisfied the pre-requisite criteria for the given course record." & vbNewLine & _
      "    Return 1 if the given record in the personnel table has NOT satisfied the pre-requisite criteria for the given course record." & vbNewLine & _
      "    Return 2 if the given record in the personnel table has NOT satisfied the pre-requisite criteria for the given course record but the user can override this failure. */" & vbNewLine & _
      "    DECLARE @dtStartDate datetime," & vbNewLine & _
      "        @sPreReqCourseTitle varchar(MAX)," & vbNewLine & _
      "        @sGrouping varchar(MAX)," & vbNewLine & _
      "        @sFailureNotification varchar(MAX)," & vbNewLine & _
      "        @fPreReqsMet bit," & vbNewLine & _
      "        @fGroupOK bit," & vbNewLine & _
      "        @sCurrentGrouping varchar(MAX)," & vbNewLine & _
      "        @fUseWarning bit," & vbNewLine & _
      "        @fError bit," & vbNewLine & _
      "        @iPreReqCount integer" & vbNewLine & vbNewLine
  
    sProcSQL = sProcSQL & _
      "    SET @fPreReqsMet = 1" & vbNewLine & _
      "    SET @fGroupOK = 1" & vbNewLine & _
      "    SET @sCurrentGrouping = ''" & vbNewLine & _
      "    SET @fUseWarning = 0" & vbNewLine & _
      "    SET @fError = 0" & vbNewLine & vbNewLine
  
    sProcSQL = sProcSQL & _
      "    /* Get the start date of the given course. */" & vbNewLine & _
      "    SELECT @dtStartDate = " & mvar_sCourseStartDateColumnName & vbNewLine & _
      "    FROM " & mvar_sCourseTableName & vbNewLine & _
      "    WHERE id = @plngCourseRecordID" & vbNewLine & vbNewLine
  
    sProcSQL = sProcSQL & _
      "    /* Get the pre-requisites for the given course. */" & vbNewLine & _
      "    DECLARE prerequisites_cursor CURSOR LOCAL FAST_FORWARD FOR" & vbNewLine & _
      "        SELECT LTRIM(RTRIM(UPPER(" & mvar_sPreReqCourseTitleName & ")))," & vbNewLine & _
      "            LTRIM(RTRIM(UPPER(" & mvar_sPreReqGroupingName & ")))," & vbNewLine & _
      "            LTRIM(RTRIM(UPPER(" & mvar_sPreReqFailureNotificationName & ")))" & vbNewLine & _
      "        FROM " & mvar_sPreReqTableName & vbNewLine & _
      "        WHERE id_" & Trim(Str(mvar_lngCourseTableID)) & " = @plngCourseRecordID" & vbNewLine & _
      "        ORDER BY " & mvar_sPreReqGroupingName & vbNewLine & vbNewLine
  
    sProcSQL = sProcSQL & _
      "    OPEN prerequisites_cursor" & vbNewLine & _
      "    FETCH NEXT FROM prerequisites_cursor INTO @sPreReqCourseTitle, @sGrouping, @sFailureNotification" & vbNewLine & _
      "    WHILE (@@fetch_status = 0)" & vbNewLine & _
      "    BEGIN" & vbNewLine & _
      "        IF @sPreReqCourseTitle IS NULL SET @sPreReqCourseTitle = ''" & vbNewLine & _
      "        IF @sGrouping IS NULL SET @sGrouping = ''" & vbNewLine & vbNewLine
  
    sProcSQL = sProcSQL & _
      "        IF (LEN(@sPreReqCourseTitle) > 0) AND (LEN(@sGrouping) > 0)" & vbNewLine & _
      "        BEGIN" & vbNewLine & _
      "            SET @fPreReqsMet = 0" & vbNewLine & vbNewLine
  
    sProcSQL = sProcSQL & _
      "            IF (@sGrouping <> @sCurrentGrouping)" & vbNewLine & _
      "            BEGIN" & vbNewLine & _
      "                /* New pre-requisite group. */" & vbNewLine & _
      "                IF LEN(@sCurrentGrouping) > 0" & vbNewLine & _
      "                BEGIN" & vbNewLine & _
      "                    IF @fGroupOK = 1" & vbNewLine & _
      "                    BEGIN" & vbNewLine & _
      "                        /* New pre-requisite group. */" & vbNewLine & _
      "                        SET @fPreReqsMet = 1" & vbNewLine & _
      "                        BREAK" & vbNewLine & _
      "                    END" & vbNewLine & _
      "                    ELSE" & vbNewLine & _
      "                    BEGIN" & vbNewLine & _
      "                        IF @fError = 0 SET @fUseWarning = 1" & vbNewLine & _
      "                    END" & vbNewLine & _
      "                END" & vbNewLine & vbNewLine & _
      "                SET @fError = 0" & vbNewLine & _
      "                SET @fGroupOK = 1" & vbNewLine & _
      "                SET @sCurrentGrouping = @sGrouping" & vbNewLine & _
      "            END" & vbNewLine & vbNewLine
  
    sProcSQL = sProcSQL & _
      "            SELECT @iPreReqCount = COUNT(" & mvar_sTrainBookTableName & ".id)" & vbNewLine & _
      "            FROM " & mvar_sTrainBookTableName & vbNewLine & _
      "            INNER JOIN " & mvar_sCourseTableName & vbNewLine & _
      "                ON " & mvar_sTrainBookTableName & ".id_" & Trim(Str(mvar_lngCourseTableID)) & " = " & mvar_sCourseTableName & ".id" & vbNewLine & _
      "            WHERE " & mvar_sTrainBookTableName & ".id_" & Trim(Str(mvar_lngEmployeeTableID)) & " = @plngEmployeeRecordID" & vbNewLine & _
      "                AND LEFT(UPPER(" & mvar_sTrainBookTableName & "." & mvar_sTrainBookStatusName & "), 1) = 'B'" & vbNewLine & _
      "                AND " & mvar_sCourseTableName & "." & mvar_sCourseTitleName & " = @sPreReqCourseTitle" & vbNewLine & _
      "                AND " & mvar_sCourseTableName & "." & mvar_sCourseEndDateColumnName & " < @dtStartDate" & vbNewLine & vbNewLine & _
      "            IF @iPreReqCount = 0" & vbNewLine & _
      "            BEGIN" & vbNewLine & _
      "                /* Pre-requisite failure. */" & vbNewLine & _
      "                SET @fGroupOK = 0" & vbNewLine & vbNewLine & _
      "                IF (@sFailureNotification = 'ERROR') SET @fError = 1" & vbNewLine & _
      "            END" & vbNewLine & vbNewLine & _
      "        END" & vbNewLine & vbNewLine & _
      "        FETCH NEXT FROM prerequisites_cursor INTO @sPreReqCourseTitle, @sGrouping, @sFailureNotification" & vbNewLine & _
      "    END" & vbNewLine & _
      "    CLOSE prerequisites_cursor" & vbNewLine & _
      "    DEALLOCATE prerequisites_cursor" & vbNewLine & vbNewLine
  
    sProcSQL = sProcSQL & _
      "    /* Return the pre-requisite criteria satisfaction code. */" & vbNewLine & _
      "    SET @piPreReqsMet = 0" & vbNewLine & _
      "    IF (@fPreReqsMet = 0) AND (@fGroupOK = 0)" & vbNewLine & _
      "    BEGIN" & vbNewLine & _
      "        IF (@fError = 1) AND (@fUseWarning = 0)" & vbNewLine & _
      "        BEGIN" & vbNewLine & _
      "            SET @piPreReqsMet = 1" & vbNewLine & _
      "        END" & vbNewLine & _
      "        ELSE" & vbNewLine & _
      "        BEGIN" & vbNewLine & _
      "            SET @piPreReqsMet = 2" & vbNewLine & _
      "        END" & vbNewLine & _
      "    END" & vbNewLine & _
      "END"
    
    gADOCon.Execute sProcSQL, , adExecuteNoRecords
  End If
  
TidyUpAndExit:
  CreatePreRequisiteCheckStoredProcedure = fCreatedOK
  Exit Function
  
ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating Pre-requisite check stored procedure (Training Booking)"
  Resume TidyUpAndExit

End Function


Private Function CreateCourseCancelDateCheckProcedure() As Boolean
  ' Create the stored procedure for checking if the cancel course date is populated for a course.
  On Error GoTo ErrorTrap
  
  Dim fCreatedOK As Boolean
  Dim sProcSQL As String
  
  fCreatedOK = True
      
  ' Construct the stored procedure creation string (if required).
  sProcSQL = "/* ------------------------------------------------ */" & vbNewLine & _
    "/* Training Booking module stored procedure. */" & vbNewLine & _
    "/* Automatically generated by the System manager.   */" & vbNewLine & _
    "/* ------------------------------------------------ */" & vbNewLine & _
    "CREATE PROCEDURE dbo." & mvar_sCourseCancelledCheck & " (" & vbNewLine & _
    "    @pfError bit OUTPUT," & vbNewLine & _
    "    @piRecID integer," & vbNewLine & _
    "    @pfCancelDate bit OUTPUT" & vbNewLine & _
    ")" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine & _
    "    SET @pfCancelDate = 0" & vbNewLine & vbNewLine

  If mvar_fCancelDateOK Then
    sProcSQL = sProcSQL & _
      "    DECLARE @dtCancelDateColumn  datetime," & vbNewLine & _
      "        @iRecId   integer" & vbNewLine & _
      "" & vbNewLine & _
      "" & vbNewLine & _
      "    SELECT @dtCancelDateColumn = " & mvar_sCourseTableName & "." & mvar_sCancelCourseColumnName & vbNewLine & _
      "     FROM " & mvar_sCourseTableName & vbNewLine & _
      "     WHERE " & mvar_sCourseTableName & ".ID" & " = @piRecID" & vbNewLine & vbNewLine
      
    sProcSQL = sProcSQL & _
      "    IF @dtCancelDateColumn IS NULL" & vbNewLine & _
      "          BEGIN" & vbNewLine & _
      "            SET @pfCancelDate = 0" & vbNewLine & _
      "          END" & vbNewLine & _
      "    ELSE" & vbNewLine & _
      "        BEGIN" & vbNewLine & _
      "            SET @pfCancelDate = 1" & vbNewLine & _
      "        END" & vbNewLine & _
      "SET @pfError=0" & vbNewLine
  End If
    
    sProcSQL = sProcSQL & _
      "END"
    
  gADOCon.Execute sProcSQL, , adExecuteNoRecords
  
TidyUpAndExit:
  CreateCourseCancelDateCheckProcedure = fCreatedOK
  Exit Function
  
ErrorTrap:
  OutputError "Error checking for Cancelled Course Date."
  fCreatedOK = False
  Resume TidyUpAndExit

End Function
Private Function CreateUnavailabilityCheckStoredProcedure() As Boolean
  ' Create the stored procedure for checking if a delegate is unavailable for a course.
  On Error GoTo ErrorTrap
  
  Dim fCreatedOK As Boolean
  Dim sProcSQL As String
  
  fCreatedOK = True
      
  If mvar_fUnavailabilityUsed Then
    ' Construct the stored procedure creation string (if required).
    sProcSQL = "/* ------------------------------------------------ */" & vbNewLine & _
      "/* Training Booking module stored procedure. */" & vbNewLine & _
      "/* Automatically generated by the System manager.   */" & vbNewLine & _
      "/* ------------------------------------------------ */" & vbNewLine & _
      "CREATE PROCEDURE dbo." & mvar_sUnavailProcedureName & " (" & vbNewLine & _
      "  @plngCourseRecordID int," & vbNewLine & _
      "  @plngEmployeeRecordID int," & vbNewLine & _
      "  @piReturnCode int OUTPUT" & vbNewLine & _
      ")" & vbNewLine & _
      "AS" & vbNewLine & _
      "BEGIN"
      
    sProcSQL = sProcSQL & vbNewLine & _
      "  /* Return 0 if the given record in the personnel table IS available for the given course record." & vbNewLine & _
      "  Return 1 if the given record in the personnel table is NOT available for the given course record." & vbNewLine & _
      "  Return 2 if the given record in the personnel table is NOT available for the given course record but the user can override this failure. */" & vbNewLine & _
      "  DECLARE @dtStartDate datetime," & vbNewLine & _
      "    @dtEndDate datetime," & vbNewLine & _
      "    @sNotification varchar(MAX)," & vbNewLine & _
      "    @fUseWarning bit," & vbNewLine & _
      "    @fPassed bit" & vbNewLine & vbNewLine
  
    sProcSQL = sProcSQL & _
      "  /* Get the start and end date of the given course. */" & vbNewLine & _
      "  SELECT @dtStartDate = " & mvar_sCourseStartDateColumnName & "," & vbNewLine & _
      "    @dtEndDate = " & mvar_sCourseEndDateColumnName & vbNewLine & _
      "  FROM " & mvar_sCourseTableName & vbNewLine & _
      "  WHERE id = @plngCourseRecordID" & vbNewLine & vbNewLine & _
      "  SET @fUseWarning = 1" & vbNewLine & _
      "  SET @fPassed = 1" & vbNewLine & vbNewLine

    sProcSQL = sProcSQL & _
      "  /*  Get the given employee's unavailable dates that overlap with the given course. */" & vbNewLine & _
      "  DECLARE unavailability_cursor CURSOR LOCAL FAST_FORWARD FOR" & vbNewLine & _
      "    SELECT LTRIM(RTRIM(UPPER(" & mvar_sUnavailFailureNotificationName & ")))" & vbNewLine & _
      "    FROM " & mvar_sUnavailTableName & vbNewLine & _
      "    WHERE id_" & Trim(Str(mvar_lngEmployeeTableID)) & " = @plngEmployeeRecordID" & vbNewLine & _
      "      AND " & mvar_sUnavailToDateName & " >= @dtStartDate" & vbNewLine & _
      "      AND " & mvar_sUnavailFromDateName & " <= @dtEndDate" & vbNewLine & vbNewLine

    sProcSQL = sProcSQL & _
      "  OPEN unavailability_cursor" & vbNewLine & _
      "  FETCH NEXT FROM unavailability_cursor INTO @sNotification" & vbNewLine & _
      "  WHILE (@@fetch_status = 0)" & vbNewLine & _
      "  BEGIN" & vbNewLine & _
      "    SET @fPassed = 0" & vbNewLine & vbNewLine

    sProcSQL = sProcSQL & _
      "    IF @sNotification = 'ERROR' SET @fUseWarning = 0" & vbNewLine & vbNewLine & _
      "    FETCH NEXT FROM unavailability_cursor INTO @sNotification" & vbNewLine & _
      "  END" & vbNewLine & vbNewLine & _
      "  CLOSE unavailability_cursor" & vbNewLine & _
      "  DEALLOCATE unavailability_cursor" & vbNewLine & vbNewLine

    sProcSQL = sProcSQL & _
      "  IF @fPassed = 0" & vbNewLine & _
      "  BEGIN" & vbNewLine & _
      "    IF @fUseWarning = 1" & vbNewLine & _
      "    BEGIN" & vbNewLine & _
      "      SET @piReturnCode = 2" & vbNewLine & _
      "    END" & vbNewLine & _
      "    ELSE" & vbNewLine & _
      "    BEGIN" & vbNewLine & _
      "      SET @piReturnCode = 1" & vbNewLine & _
      "    END" & vbNewLine & _
      "  END" & vbNewLine & _
      "  ELSE" & vbNewLine & _
      "  BEGIN" & vbNewLine & _
      "    SET @piReturnCode = 0" & vbNewLine & _
      "  END" & vbNewLine & _
      "END"
      
    gADOCon.Execute sProcSQL, , adExecuteNoRecords
  End If
  
TidyUpAndExit:
  CreateUnavailabilityCheckStoredProcedure = fCreatedOK
  Exit Function
  
ErrorTrap:
  OutputError "Error creating Unavailability check stored procedure (Training Booking)"
  fCreatedOK = False
  Resume TidyUpAndExit

End Function


Private Function CreateOverbookingCheckStoredProcedure() As Boolean
  ' Create the stored procedure for checking if another booking
  ' will exceed a course's maximum number of delegates.
  On Error GoTo ErrorTrap
  
  Dim fCreatedOK As Boolean
  Dim sProcSQL As String
  
  fCreatedOK = True
      
  ' Construct the stored procedure creation string (if required).
  sProcSQL = "/* -------------------------------------------------------------------------------- */" & vbNewLine & _
    "/* Training Booking module stored procedure. */" & vbNewLine & _
    "/* Automatically generated by the System Manager.    */" & vbNewLine & _
    "/* -------------------------------------------------------------------------------- */" & vbNewLine & _
    "CREATE PROCEDURE dbo." & mvar_sOverbookingProcedureName & " (" & vbNewLine & _
    "  @plngCourseRecordID int," & vbNewLine & _
    "  @piBookingID int," & vbNewLine & _
    "  @piNewBookings int," & vbNewLine & _
    "  @piReturnCode int OUTPUT" & vbNewLine & _
    ")" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine
    
  sProcSQL = sProcSQL & _
    "  /* Return 0 if the given course will NOT exceed its maximum number of delegates if the given number are added." & vbNewLine & _
    "  Return 1 if the given course WILL exceed its maximum number of delegates if the given number are added." & vbNewLine & _
    "  Return 2 if the given course WILL exceed its maximum number of delegates if the given number are added, but the user can override this failure. */" & vbNewLine & _
    "  DECLARE @iBookedCount int," & vbNewLine & _
    "    @iMaxBookings int" & vbNewLine & vbNewLine

  sProcSQL = sProcSQL & _
    "  /* Get the maximum number of delegates allowed on the given course. */" & vbNewLine & _
    "  SELECT @iMaxBookings = " & mvar_sCourseMaxNumberName & vbNewLine & _
    "  FROM " & mvar_sCourseTableName & vbNewLine & _
    "  WHERE id = @plngCourseRecordID" & vbNewLine & vbNewLine

  sProcSQL = sProcSQL & _
    "  /* Get the number of delegates booked on the given course. */" & vbNewLine & _
    "  SELECT @iBookedCount = COUNT(id)" & vbNewLine & _
    "  FROM " & mvar_sTrainBookTableName & vbNewLine & _
    "  WHERE id_" & Trim(Str(mvar_lngCourseTableID)) & " = @plngCourseRecordID" & vbNewLine & _
    "    AND id <> @piBookingID" & vbNewLine & _
    "    AND (" & mvar_sTrainBookStatusName & " = 'B'" & _
    IIf(mvar_fCourseIncludeProvisionals, " OR " & mvar_sTrainBookStatusName & " = 'P'", "") & ")" & vbNewLine & vbNewLine
    
  sProcSQL = sProcSQL & _
    "  SET @iBookedCount = @iBookedCount + @piNewBookings" & vbNewLine & vbNewLine

  sProcSQL = sProcSQL & _
    "  IF (@iBookedCount > @iMaxBookings) AND (@iMaxBookings > 0)" & vbNewLine & _
    "  BEGIN" & vbNewLine & _
    "    SET @piReturnCode = " & IIf(mvar_iCourseOverbookingNotification = 0, "1", "2") & vbNewLine & _
    "  END" & vbNewLine & _
    "  ELSE" & vbNewLine & _
    "  BEGIN" & vbNewLine & _
    "    SET @piReturnCode = 0" & vbNewLine & _
    "  END" & vbNewLine & _
    "END" & vbNewLine
    
  gADOCon.Execute sProcSQL, , adExecuteNoRecords
  
TidyUpAndExit:
  CreateOverbookingCheckStoredProcedure = fCreatedOK
  Exit Function
  
ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating Overbooking check stored procedure (Training Booking)"
  Resume TidyUpAndExit

End Function



Private Function CreateOverlappedBookingCheckStoredProcedure() As Boolean
  ' Create the stored procedure for checking if another booking
  ' will exceed a course's maximum number of delegates.
  On Error GoTo ErrorTrap
  
  Dim fCreatedOK As Boolean
  Dim sProcSQL As String
  
  fCreatedOK = True
  
  ' Construct the stored procedure creation string (if required).
  sProcSQL = "/* -------------------------------------------------------------------------------- */" & vbNewLine & _
    "/* Training Booking module stored procedure. */" & vbNewLine & _
    "/* Automatically generated by the System manager.   */" & vbNewLine & _
    "/* -------------------------------------------------------------------------------- */" & vbNewLine & _
    "CREATE PROCEDURE dbo." & mvar_sOverlapProcedureName & " (" & vbNewLine & _
    "  @plngCourseRecordID int," & vbNewLine & _
    "  @plngEmployeeRecordID int," & vbNewLine & _
    "  @plngBookingRecordID int," & vbNewLine & _
    "  @piReturnCode int OUTPUT" & vbNewLine & _
    ")" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine

    sProcSQL = sProcSQL & _
        "  /* Return 0 if the given course does NOT overlap with another course that the given delegate is booked on." & vbNewLine & _
        "  Return 1 if the given course DOES overlap with another course that the given delegate is booked on." & vbNewLine & _
        "  Return 2 if the given course does NOT overlap with another course that the given delegate is booked on, but the user can override this failure. */" & vbNewLine & _
        "  DECLARE @iOverlapCount int," & vbNewLine & _
        "    @dtStartDate  datetime," & vbNewLine & _
        "    @dtEndDate  datetime" & vbNewLine & vbNewLine

    sProcSQL = sProcSQL & _
      "  /* Get the dates of the given course. */" & vbNewLine & _
      "  SELECT @dtStartDate = " & mvar_sCourseStartDateColumnName & "," & vbNewLine & _
      "    @dtEndDate = " & mvar_sCourseEndDateColumnName & vbNewLine & _
      "  FROM " & mvar_sCourseTableName & vbNewLine & _
      "  WHERE id = @plngCourseRecordID" & vbNewLine & vbNewLine & _
      "  IF (@dtStartDate IS NULL) AND NOT (@dtEndDate IS NULL) SET @dtStartDate = @dtEndDate" & vbNewLine & _
      "  IF (@dtEndDate IS NULL) AND NOT (@dtStartDate IS NULL) SET @dtEndDate = @dtStartDate" & vbNewLine & vbNewLine


    sProcSQL = sProcSQL & _
      "  IF (@dtStartDate IS NULL) AND (@dtEndDate IS NULL)" & vbNewLine & _
      "  BEGIN" & vbNewLine & _
      "    SET @iOverlapCount = 0" & vbNewLine & _
      "  END" & vbNewLine & _
      "  ELSE" & vbNewLine & _
      "  BEGIN" & vbNewLine & _
      "    SELECT @iOverlapCount = COUNT(" & mvar_sTrainBookTableName & ".ID)" & vbNewLine & _
      "    FROM " & mvar_sTrainBookTableName & vbNewLine & _
      "    INNER JOIN " & mvar_sCourseTableName & vbNewLine & _
      "      ON " & mvar_sTrainBookTableName & ".ID_" & Trim(Str(mvar_lngCourseTableID)) & " = " & mvar_sCourseTableName & ".ID" & vbNewLine & _
      "    WHERE " & mvar_sTrainBookTableName & ".ID_" & Trim(Str(mvar_lngEmployeeTableID)) & " = @plngEmployeeRecordID" & vbNewLine & _
      "      AND " & mvar_sTrainBookTableName & ".ID <> @plngBookingRecordID" & vbNewLine & _
      "      AND (LEFT(UPPER(" & mvar_sTrainBookTableName & "." & mvar_sTrainBookStatusName & "), 1) = 'B'" & vbNewLine & _
      "      OR LEFT(UPPER(" & mvar_sTrainBookTableName & "." & mvar_sTrainBookStatusName & "), 1) = 'P')" & vbNewLine & _
      "      AND NOT(" & mvar_sCourseTableName & "." & mvar_sCourseStartDateColumnName & " IS NULL AND " & mvar_sCourseTableName & "." & mvar_sCourseEndDateColumnName & " IS NULL)" & vbNewLine & _
      "      AND CASE" & vbNewLine & _
      "          WHEN " & mvar_sCourseTableName & "." & mvar_sCourseStartDateColumnName & " IS NULL THEN " & mvar_sCourseTableName & "." & mvar_sCourseEndDateColumnName & vbNewLine & _
      "          ELSE " & mvar_sCourseTableName & "." & mvar_sCourseStartDateColumnName & vbNewLine & _
      "        END <= @dtEndDate" & vbNewLine & _
      "      AND CASE" & vbNewLine & _
      "          WHEN " & mvar_sCourseTableName & "." & mvar_sCourseEndDateColumnName & " IS NULL THEN " & mvar_sCourseTableName & "." & mvar_sCourseStartDateColumnName & vbNewLine & _
      "          ELSE " & mvar_sCourseTableName & "." & mvar_sCourseEndDateColumnName & vbNewLine & _
      "        END >= @dtStartDate" & vbNewLine & _
      "  END" & vbNewLine & vbNewLine

    sProcSQL = sProcSQL & _
      "  IF @iOverlapCount > 0" & vbNewLine & _
      "  BEGIN" & vbNewLine & _
      "    SET @piReturnCode = " & IIf(mvar_iTrainBookOverlapNotification = 0, "1", "2") & vbNewLine & _
      "  END" & vbNewLine & _
      "  ELSE" & vbNewLine & _
      "  BEGIN" & vbNewLine & _
      "    SET @piReturnCode = 0" & vbNewLine & _
      "  END" & vbNewLine & _
      "END"

  gADOCon.Execute sProcSQL, , adExecuteNoRecords
  
TidyUpAndExit:
  CreateOverlappedBookingCheckStoredProcedure = fCreatedOK
  Exit Function
  
ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating Overlapped Booking check stored procedure (Training Booking)"
  Resume TidyUpAndExit

End Function




Private Function ReadCourseRecordParameters() As Boolean
  ' Read the configured Training Booking parameters into member variables.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean

  With recModuleSetup
    .Index = "idxModuleParameter"

    ' Get the Course table ID and name.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSETABLE
    fOK = Not .NoMatch
    If fOK Then
      'MH20001025 Fault 799 Type Mismatch
      'fOK = Not IsNull(!parametervalue)
      fOK = (IIf(IsNull(!parametervalue), 0, val(!parametervalue)) > 0)
    End If
    If Not fOK Then
      mvar_fPreReqsOK = False
      mvar_fUnavailOK = False
      mvar_fOverlapOK = False
      mvar_fOverbookOK = False
      
      mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  'Training Course' table not defined."
    Else
      mvar_lngCourseTableID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))

      With recTabEdit
        .Index = "idxTableID"
        .Seek "=", mvar_lngCourseTableID
      
        fOK = Not .NoMatch
        If fOK Then
          fOK = Not IsNull(!TableName)
        End If
        If Not fOK Then
          mvar_fPreReqsOK = False
          mvar_fUnavailOK = False
          mvar_fOverlapOK = False
          mvar_fOverbookOK = False
          mvar_fCancelDateOK = False
          
          mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  'Training Course' table not found."
        Else
          mvar_sCourseTableName = !TableName
        End If
      End With
    End If
    
    If mvar_fCancelDateOK Then
          ' Get the Course Cancel Date column ID.
          .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSECANCELLATIONDATE
          fOK = Not .NoMatch
          If fOK Then
            fOK = (IIf(IsNull(!parametervalue), 0, val(!parametervalue)) > 0)
          End If
          If Not fOK Then
            mvar_fCancelDateOK = False
            
            mvar_sCancelDateMsg = mvar_sCancelDateMsg & vbNewLine & "  '" & mvar_sCourseTableName & "' table 'Course Cancellation' column not defined."
          Else
            mvar_lngCancelCourseID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
          
            With recColEdit
              .Index = "idxColumnID"
              .Seek "=", mvar_lngCancelCourseID
            
              fOK = Not .NoMatch
              If fOK Then
                fOK = Not IsNull(!ColumnName)
              End If
              If Not fOK Then
                mvar_fCancelDateOK = False
                
                mvar_sCancelDateMsg = mvar_sCancelDateMsg & vbNewLine & "  '" & mvar_sCourseTableName & "' table 'Course Cancellation' column not found."
              Else
                mvar_sCancelCourseColumnName = !ColumnName
              End If
            End With
          End If
    End If
    
    If mvar_fPreReqsOK Then
      ' Get the Course Title column ID.
      .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSETITLE
      fOK = Not .NoMatch
      If fOK Then
        'MH20001025 Fault 799 Type Mismatch
        'fOK = Not IsNull(!parametervalue)
        fOK = (IIf(IsNull(!parametervalue), 0, val(!parametervalue)) > 0)
      End If
      If Not fOK Then
        mvar_fPreReqsOK = False
        
        mvar_sPreReqsMsg = mvar_sPreReqsMsg & vbNewLine & "  '" & mvar_sCourseTableName & "' table 'Course Title' column not defined."
      Else
        mvar_lngCourseTitleID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
      
        With recColEdit
          .Index = "idxColumnID"
          .Seek "=", mvar_lngCourseTitleID
        
          fOK = Not .NoMatch
          If fOK Then
            fOK = Not IsNull(!ColumnName)
          End If
          If Not fOK Then
            mvar_fPreReqsOK = False
            
            mvar_sPreReqsMsg = mvar_sPreReqsMsg & vbNewLine & "  '" & mvar_sCourseTableName & "' table 'Course Title' column not found."
          Else
            mvar_sCourseTitleName = !ColumnName
          End If
        End With
      End If
    End If

    If mvar_fPreReqsOK Or mvar_fUnavailOK Or mvar_fOverlapOK Then
      ' Get the Course Start Date column ID.
      .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSESTARTDATE
      fOK = Not .NoMatch
      If fOK Then
        'MH20001025 Fault 799 Type Mismatch
        'fOK = Not IsNull(!parametervalue)
        fOK = (IIf(IsNull(!parametervalue), 0, val(!parametervalue)) > 0)
      End If
      If Not fOK Then
        mvar_fPreReqsOK = False
        mvar_fUnavailOK = False
        mvar_fOverlapOK = False
        
        mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  '" & mvar_sCourseTableName & "' table 'Start Date' column not defined."
      Else
        mvar_lngCourseStartDateID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
      
        With recColEdit
          .Index = "idxColumnID"
          .Seek "=", mvar_lngCourseStartDateID
        
          fOK = Not .NoMatch
          If fOK Then
            fOK = Not IsNull(!ColumnName)
          End If
          If Not fOK Then
            mvar_fPreReqsOK = False
            mvar_fUnavailOK = False
            mvar_fOverlapOK = False
            
            mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  '" & mvar_sCourseTableName & "' table 'Start Date' column not found."
          Else
            mvar_sCourseStartDateColumnName = !ColumnName
          End If
        End With
      End If
    End If
    
    If mvar_fPreReqsOK Or mvar_fUnavailOK Or mvar_fOverlapOK Then
      ' Get the Course End Date column ID.
      .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSEENDDATE
      fOK = Not .NoMatch
      If fOK Then
        'MH20001025 Fault 799 Type Mismatch
        'fOK = Not IsNull(!parametervalue)
        fOK = (IIf(IsNull(!parametervalue), 0, val(!parametervalue)) > 0)
      End If
      If Not fOK Then
        mvar_fPreReqsOK = False
        mvar_fUnavailOK = False
        mvar_fOverlapOK = False
        
        mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  '" & mvar_sCourseTableName & "' table 'End Date' column not defined."
      Else
        mvar_lngCourseEndDateID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
      
        With recColEdit
          .Index = "idxColumnID"
          .Seek "=", mvar_lngCourseEndDateID
        
          fOK = Not .NoMatch
          If fOK Then
            fOK = Not IsNull(!ColumnName)
          End If
          If Not fOK Then
            mvar_fPreReqsOK = False
            mvar_fUnavailOK = False
            mvar_fOverlapOK = False
            
            mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  '" & mvar_sCourseTableName & "' table 'End Date' column not found."
          Else
            mvar_sCourseEndDateColumnName = !ColumnName
          End If
        End With
      End If
    End If
    
'' JPD - No longer needed
''    If mvar_fOverbookOK Then
''      ' Get the Course Number Booked column ID.
''      .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSENUMBERBOOKED
''      fOK = Not .NoMatch
''      If fOK Then
''        'MH20001025 Fault 799 Type Mismatch
''        'fOK = Not IsNull(!parametervalue)
''        fOK = (IIf(IsNull(!parametervalue), 0, Val(!parametervalue)) > 0)
''      End If
''      If Not fOK Then
''        mvar_fOverbookOK = False
''
''        mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewline & "  '" & mvar_sCourseTableName & "' table 'Number Booked' column not defined."
''      Else
''        mvar_lngCourseNumberBookedID = IIf(IsNull(!parametervalue), 0, Val(!parametervalue))
''
''        With recColEdit
''          .Index = "idxColumnID"
''          .Seek "=", mvar_lngCourseNumberBookedID
''
''          fOK = Not .NoMatch
''          If fOK Then
''            fOK = Not IsNull(!ColumnName)
''          End If
''          If Not fOK Then
''            mvar_fOverbookOK = False
''
''            mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewline & "  '" & mvar_sCourseTableName & "' table 'Number Booked' column not found."
''          Else
''            mvar_sCourseNumberBookedName = !ColumnName
''          End If
''        End With
''      End If
''    End If
    
    If mvar_fOverbookOK Then
      ' Get the Course Max. Number of Delegates column ID.
      .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSEMAXNUMBER
      fOK = Not .NoMatch
      If fOK Then
        'MH20001025 Fault 799 Type Mismatch
        'fOK = Not IsNull(!parametervalue)
        fOK = (IIf(IsNull(!parametervalue), 0, val(!parametervalue)) > 0)
      End If
      If Not fOK Then
        mvar_fOverbookOK = False
            
        mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  '" & mvar_sCourseTableName & "' table 'Max. Number of Delegates' column not defined."
      Else
        mvar_lngCourseMaxNumberID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
      
        With recColEdit
          .Index = "idxColumnID"
          .Seek "=", mvar_lngCourseMaxNumberID
        
          fOK = Not .NoMatch
          If fOK Then
            fOK = Not IsNull(!ColumnName)
          End If
          If Not fOK Then
            mvar_fOverbookOK = False
            
            mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  '" & mvar_sCourseTableName & "' table 'Max. Number of Delegates' column not found."
          Else
            mvar_sCourseMaxNumberName = !ColumnName
          End If
        End With
      End If
    End If
    
    ' Get the Course Include Provisional Bookings flag.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSEINCLUDEPROVISIONALS
    If .NoMatch Then
      mvar_fCourseIncludeProvisionals = False
    Else
      mvar_fCourseIncludeProvisionals = IIf(IsNull(!parametervalue), False, IIf(!parametervalue = "TRUE", True, False))
    End If

    ' Get the Overbooking Notification flag.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSEOVERBOOKINGNOTIFICATION
    If .NoMatch Then
      mvar_iCourseOverbookingNotification = 0
    Else
      'MH20001025 Fault 799
      'mvar_iCourseOverbookingNotification = IIf(IsNull(!parametervalue), 0, !parametervalue)
      mvar_iCourseOverbookingNotification = val(IIf(IsNull(!parametervalue), 0, !parametervalue))
    End If
  End With

  fOK = True
  
TidyUpAndExit:
  ReadCourseRecordParameters = fOK
  Exit Function
  
ErrorTrap:
  OutputError "Error reading course record parameters (Training Booking)"
  fOK = False
  Resume TidyUpAndExit
  
End Function
Private Function ReadEmployeeRecordParameters() As Boolean
  ' Read the Employee Records parameter values from the database into local variables.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  With recModuleSetup
    .Index = "idxModuleParameter"
    
    ' Get the Employee Table ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_EMPLOYEETABLE
    fOK = Not .NoMatch
    If fOK Then
      'MH20001025 Fault 799 Type Mismatch
      fOK = (IIf(IsNull(!parametervalue), 0, val(!parametervalue)) > 0)
    End If
    
    If Not fOK Then
      mvar_fPreReqsOK = False
      mvar_fUnavailOK = False
      mvar_fOverlapOK = False
      mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  'Employee' table not defined."
    Else
      mvar_lngEmployeeTableID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
    End If
    
'    ' Get the Bulk Booking Id
'    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_BULKBOOKINGDEFAULTVIEW
'    fOK = Not .NoMatch
'    If fOK Then
'      fOK = (IIf(IsNull(!parametervalue), 0, Val(!parametervalue)) > 0)
'    End If
'
'    If fOK Then
'      'mvar_lngBulkBookingDefaultViewID = IIf(IsNull(!parametervalue), 0, Val(!parametervalue))
'
'    Else
'
'    End If
  End With

  fOK = True
  
TidyUpAndExit:
  ReadEmployeeRecordParameters = fOK
  Exit Function
  
ErrorTrap:
  OutputError "Error reading employee record parameters (Training Booking)"
  fOK = False
  Resume TidyUpAndExit

End Function
