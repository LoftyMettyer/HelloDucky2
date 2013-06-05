Namespace Things

  Public Class WorkflowElement
    Inherits Things.Base

    '    Public ID As HCMGuid
    '   Public WorkflowID As HCMGuid
    '    Public Type
    Public Caption As String
    Public ConnectionPairID As Integer
    Public LeftCoord As Integer
    Public TopCoord As Integer
    Public DecisionCaptionType As Integer
    Public Identifier As String
    Public TrueFlowIdentifier As String
    Public DataAction As Integer
    Public DataTableID As HCMGuid
    Public DataRecord As Integer
    Public EmailID As HCMGuid
    Public EmailRecord As Integer
    Public WebFormBGColor As Integer
    Public WebFormBGImageID As Integer
    Public WebFormBGImageLocation As Integer
    Public WebFormDefaultFontName As String
    Public WebFormDefaultFontSize As Integer
    Public WebFormDefaultFontBold As Boolean
    Public WebFormDefaultFontItalic As Boolean
    Public WebFormDefaultFontStrikeThru As Boolean
    Public WebFormDefaultFontUnderline As Boolean
    Public WebFormHeight As Integer
    Public WebFormWidth As Integer
    Public RecSelWebFormIdentifier As String
    Public RecSelIdentifier As String
    Public SecondaryDataRecord As Integer
    Public SecondaryRecSelWebFormIdentifier As String
    Public SecondaryRecSelIdentifier As String
    Public EmailSubject As String
    Public TimeoutFrequency As Integer
    Public TimeoutPeriod As Integer
    Public DataRecordTable As HCMGuid
    Public SecondaryDataRecordTable As HCMGuid
    Public TrueFlowType As Integer
    Public TrueFlowExprID As HCMGuid
    Public DescriptionExprID As HCMGuid
    Public WebFormFGColor As Integer
    Public DescHasWorkflowName As Boolean
    Public DescHasElementCaption As Boolean
    Public EmailCCID As HCMGuid
    Public TimeoutExcludeWeekend As Boolean
    Public CompletionMessageType As Integer
    Public CompletionMessage As String
    Public SavedForLaterMessageType As Integer
    Public SavedForLaterMessage As String
    Public FollowOnFormsMessageType As Integer
    Public FollowOnFormsMessage As String
    Public Attachment_Type As String
    Public Attachment_File As String
    Public Attachment_WFElementIdentifier As String
    Public Attachment_WFValueIdentifier As String
    Public Attachment_DBColumnID As HCMGuid
    Public Attachment_DBRecord As HCMGuid
    Public Attachment_DBElement As String
    Public Attachment_DBValue As String

    'Public Overrides Function Commit() As Boolean
    'End Function

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.WorkflowElement
      End Get
    End Property
  End Class

End Namespace
