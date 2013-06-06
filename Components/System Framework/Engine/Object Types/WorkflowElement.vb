Imports System.Runtime.InteropServices

Namespace Things

  <ClassInterface(ClassInterfaceType.None), ComVisible(True), Serializable()>
  Public Class WorkflowElement
    Inherits Base
    Implements IWorkflowElement

    Public Property WorkFlowID As Integer Implements IWorkflowElement.WorkflowID
    Public Property Caption As String Implements IWorkflowElement.Caption
    Public Property ConnectionPairID As Integer Implements IWorkflowElement.ConnectionPairID
    Public Property LeftCoord As Integer Implements IWorkflowElement.LeftCoord
    Public Property TopCoord As Integer Implements IWorkflowElement.TopCoord
    Public Property DecisionCaptionType As Integer Implements IWorkflowElement.DecisionCaptionType
    Public Property Identifier As String Implements IWorkflowElement.Identifier
    Public Property TrueFlowIdentifier As String Implements IWorkflowElement.TrueFlowIdentifier
    Public Property DataAction As Integer Implements IWorkflowElement.DataAction
    Public Property DataTableID As Integer Implements IWorkflowElement.DataTableID
    Public Property DataRecord As Integer Implements IWorkflowElement.DataRecord
    Public Property EmailID As Integer Implements IWorkflowElement.EmailID
    Public Property EmailRecord As Integer Implements IWorkflowElement.EmailRecord
    Public Property WebFormBGColor As Integer Implements IWorkflowElement.WebFormBGColor
    Public Property WebFormBGImageID As Integer Implements IWorkflowElement.WebFormBGImageID
    Public Property WebFormBGImageLocation As Integer Implements IWorkflowElement.WebFormBGImageLocation
    Public Property WebFormDefaultFontName As String Implements IWorkflowElement.WebFormDefaultFontName
    Public Property WebFormDefaultFontSize As Integer Implements IWorkflowElement.WebFormDefaultFontSize
    Public Property WebFormDefaultFontBold As Boolean Implements IWorkflowElement.WebFormDefaultFontBold
    Public Property WebFormDefaultFontItalic As Boolean Implements IWorkflowElement.WebFormDefaultFontItalic
    Public Property WebFormDefaultFontStrikeThru As Boolean Implements IWorkflowElement.WebFormDefaultFontStrikeThru
    Public Property WebFormDefaultFontUnderline As Boolean Implements IWorkflowElement.WebFormDefaultFontUnderline
    Public Property WebFormHeight As Integer Implements IWorkflowElement.WebFormHeight
    Public Property WebFormWidth As Integer Implements IWorkflowElement.WebFormWidth
    Public Property RecSelWebFormIdentifier As String Implements IWorkflowElement.RecSelWebFormIdentifier
    Public Property RecSelIdentifier As String Implements IWorkflowElement.RecSelIdentifier
    Public Property SecondaryDataRecord As Integer Implements IWorkflowElement.SecondaryDataRecord
    Public Property SecondaryRecSelWebFormIdentifier As String Implements IWorkflowElement.SecondaryRecSelWebFormIdentifier
    Public Property SecondaryRecSelIdentifier As String Implements IWorkflowElement.SecondaryRecSelIdentifier
    Public Property EmailSubject As String Implements IWorkflowElement.EmailSubject
    Public Property TimeoutFrequency As Integer Implements IWorkflowElement.TimeoutFrequency
    Public Property TimeoutPeriod As Integer Implements IWorkflowElement.TimeoutPeriod
    Public Property DataRecordTable As Integer Implements IWorkflowElement.DataRecordTable
    Public Property SecondaryDataRecordTable As Integer Implements IWorkflowElement.SecondaryDataRecordTable
    Public Property TrueFlowType As Integer Implements IWorkflowElement.TrueFlowType
    Public Property TrueFlowExprID As Integer Implements IWorkflowElement.TrueFlowExprID
    Public Property DescriptionExprID As Integer Implements IWorkflowElement.DescriptionExprID
    Public Property WebFormFGColor As Integer Implements IWorkflowElement.WebFormFGColor
    Public Property DescHasWorkflowName As Boolean Implements IWorkflowElement.DescHasWorkflowName
    Public Property DescHasElementCaption As Boolean Implements IWorkflowElement.DescHasElementCaption
    Public Property EmailCCID As Integer Implements IWorkflowElement.EmailCCID
    Public Property TimeoutExcludeWeekend As Boolean Implements IWorkflowElement.TimeoutExcludeWeekend
    Public Property CompletionMessageType As Integer Implements IWorkflowElement.CompletionMessageType
    Public Property CompletionMessage As String Implements IWorkflowElement.CompletionMessage
    Public Property SavedForLaterMessageType As Integer Implements IWorkflowElement.SavedForLaterMessageType
    Public Property SavedForLaterMessage As String Implements IWorkflowElement.SavedForLaterMessage
    Public Property FollowOnFormsMessageType As Integer Implements IWorkflowElement.FollowOnFormsMessageType
    Public Property FollowOnFormsMessage As String Implements IWorkflowElement.FollowOnFormsMessage
    Public Property Attachment_Type As String Implements IWorkflowElement.Attachment_Type
    Public Property Attachment_File As String Implements IWorkflowElement.Attachment_File
    Public Property Attachment_WFElementIdentifier As String Implements IWorkflowElement.Attachment_WFElementIdentifier
    Public Property Attachment_WFValueIdentifier As String Implements IWorkflowElement.Attachment_WFValueIdentifier
    Public Property Attachment_DBColumnID As Integer Implements IWorkflowElement.Attachment_DBColumnID
    Public Property Attachment_DBRecord As Integer Implements IWorkflowElement.Attachment_DBRecord
    Public Property Attachment_DBElement As String Implements IWorkflowElement.Attachment_DBElement
    Public Property Attachment_DBValue As String Implements IWorkflowElement.Attachment_DBValue

  End Class

End Namespace
