Imports System.Runtime.InteropServices

Namespace Things

  <ClassInterface(ClassInterfaceType.None), ComVisible(True), Serializable()> _
  Public Class WorkflowElement
    Inherits Things.Base
    Implements iWorkflowElement

    Public Property WorkFlowID As HCMGuid Implements COMInterfaces.iWorkflowElement.WorkflowID
    Public Property Caption As String Implements COMInterfaces.iWorkflowElement.Caption
    Public Property ConnectionPairID As Integer Implements COMInterfaces.iWorkflowElement.ConnectionPairID
    Public Property LeftCoord As Integer Implements COMInterfaces.iWorkflowElement.LeftCoord
    Public Property TopCoord As Integer Implements COMInterfaces.iWorkflowElement.TopCoord
    Public Property DecisionCaptionType As Integer Implements COMInterfaces.iWorkflowElement.DecisionCaptionType
    Public Property Identifier As String Implements COMInterfaces.iWorkflowElement.Identifier
    Public Property TrueFlowIdentifier As String Implements COMInterfaces.iWorkflowElement.TrueFlowIdentifier
    Public Property DataAction As Integer Implements COMInterfaces.iWorkflowElement.DataAction
    Public Property DataTableID As HCMGuid Implements COMInterfaces.iWorkflowElement.DataTableID
    Public Property DataRecord As Integer Implements COMInterfaces.iWorkflowElement.DataRecord
    Public Property EmailID As HCMGuid Implements COMInterfaces.iWorkflowElement.EmailID
    Public Property EmailRecord As Integer Implements COMInterfaces.iWorkflowElement.EmailRecord
    Public Property WebFormBGColor As Integer Implements COMInterfaces.iWorkflowElement.WebFormBGColor
    Public Property WebFormBGImageID As Integer Implements COMInterfaces.iWorkflowElement.WebFormBGImageID
    Public Property WebFormBGImageLocation As Integer Implements COMInterfaces.iWorkflowElement.WebFormBGImageLocation
    Public Property WebFormDefaultFontName As String Implements COMInterfaces.iWorkflowElement.WebFormDefaultFontName
    Public Property WebFormDefaultFontSize As Integer Implements COMInterfaces.iWorkflowElement.WebFormDefaultFontSize
    Public Property WebFormDefaultFontBold As Boolean Implements COMInterfaces.iWorkflowElement.WebFormDefaultFontBold
    Public Property WebFormDefaultFontItalic As Boolean Implements COMInterfaces.iWorkflowElement.WebFormDefaultFontItalic
    Public Property WebFormDefaultFontStrikeThru As Boolean Implements COMInterfaces.iWorkflowElement.WebFormDefaultFontStrikeThru
    Public Property WebFormDefaultFontUnderline As Boolean Implements COMInterfaces.iWorkflowElement.WebFormDefaultFontUnderline
    Public Property WebFormHeight As Integer Implements COMInterfaces.iWorkflowElement.WebFormHeight
    Public Property WebFormWidth As Integer Implements COMInterfaces.iWorkflowElement.WebFormWidth
    Public Property RecSelWebFormIdentifier As String Implements COMInterfaces.iWorkflowElement.RecSelWebFormIdentifier
    Public Property RecSelIdentifier As String Implements COMInterfaces.iWorkflowElement.RecSelIdentifier
    Public Property SecondaryDataRecord As Integer Implements COMInterfaces.iWorkflowElement.SecondaryDataRecord
    Public Property SecondaryRecSelWebFormIdentifier As String Implements COMInterfaces.iWorkflowElement.SecondaryRecSelWebFormIdentifier
    Public Property SecondaryRecSelIdentifier As String Implements COMInterfaces.iWorkflowElement.SecondaryRecSelIdentifier
    Public Property EmailSubject As String Implements COMInterfaces.iWorkflowElement.EmailSubject
    Public Property TimeoutFrequency As Integer Implements COMInterfaces.iWorkflowElement.TimeoutFrequency
    Public Property TimeoutPeriod As Integer Implements COMInterfaces.iWorkflowElement.TimeoutPeriod
    Public Property DataRecordTable As HCMGuid Implements COMInterfaces.iWorkflowElement.DataRecordTable
    Public Property SecondaryDataRecordTable As HCMGuid Implements COMInterfaces.iWorkflowElement.SecondaryDataRecordTable
    Public Property TrueFlowType As Integer Implements COMInterfaces.iWorkflowElement.TrueFlowType
    Public Property TrueFlowExprID As HCMGuid Implements COMInterfaces.iWorkflowElement.TrueFlowExprID
    Public Property DescriptionExprID As HCMGuid Implements COMInterfaces.iWorkflowElement.DescriptionExprID
    Public Property WebFormFGColor As Integer Implements COMInterfaces.iWorkflowElement.WebFormFGColor
    Public Property DescHasWorkflowName As Boolean Implements COMInterfaces.iWorkflowElement.DescHasWorkflowName
    Public Property DescHasElementCaption As Boolean Implements COMInterfaces.iWorkflowElement.DescHasElementCaption
    Public Property EmailCCID As HCMGuid Implements COMInterfaces.iWorkflowElement.EmailCCID
    Public Property TimeoutExcludeWeekend As Boolean Implements COMInterfaces.iWorkflowElement.TimeoutExcludeWeekend
    Public Property CompletionMessageType As Integer Implements COMInterfaces.iWorkflowElement.CompletionMessageType
    Public Property CompletionMessage As String Implements COMInterfaces.iWorkflowElement.CompletionMessage
    Public Property SavedForLaterMessageType As Integer Implements COMInterfaces.iWorkflowElement.SavedForLaterMessageType
    Public Property SavedForLaterMessage As String Implements COMInterfaces.iWorkflowElement.SavedForLaterMessage
    Public Property FollowOnFormsMessageType As Integer Implements COMInterfaces.iWorkflowElement.FollowOnFormsMessageType
    Public Property FollowOnFormsMessage As String Implements COMInterfaces.iWorkflowElement.FollowOnFormsMessage
    Public Property Attachment_Type As String Implements COMInterfaces.iWorkflowElement.Attachment_Type
    Public Property Attachment_File As String Implements COMInterfaces.iWorkflowElement.Attachment_File
    Public Property Attachment_WFElementIdentifier As String Implements COMInterfaces.iWorkflowElement.Attachment_WFElementIdentifier
    Public Property Attachment_WFValueIdentifier As String Implements COMInterfaces.iWorkflowElement.Attachment_WFValueIdentifier
    Public Property Attachment_DBColumnID As HCMGuid Implements COMInterfaces.iWorkflowElement.Attachment_DBColumnID
    Public Property Attachment_DBRecord As HCMGuid Implements COMInterfaces.iWorkflowElement.Attachment_DBRecord
    Public Property Attachment_DBElement As String Implements COMInterfaces.iWorkflowElement.Attachment_DBElement
    Public Property Attachment_DBValue As String Implements COMInterfaces.iWorkflowElement.Attachment_DBValue

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.WorkflowElement
      End Get
    End Property

  End Class

End Namespace
