Imports System.Runtime.InteropServices

Namespace Things

  <ClassInterface(ClassInterfaceType.None), ComVisible(True), Serializable()>
  Public Class WorkflowElement
    Inherits Base
    Implements IWorkflowElement

    Public Property WorkFlowID As Integer Implements COMInterfaces.IWorkflowElement.WorkflowID
    Public Property Caption As String Implements COMInterfaces.IWorkflowElement.Caption
    Public Property ConnectionPairID As Integer Implements COMInterfaces.IWorkflowElement.ConnectionPairID
    Public Property LeftCoord As Integer Implements COMInterfaces.IWorkflowElement.LeftCoord
    Public Property TopCoord As Integer Implements COMInterfaces.IWorkflowElement.TopCoord
    Public Property DecisionCaptionType As Integer Implements COMInterfaces.IWorkflowElement.DecisionCaptionType
    Public Property Identifier As String Implements COMInterfaces.IWorkflowElement.Identifier
    Public Property TrueFlowIdentifier As String Implements COMInterfaces.IWorkflowElement.TrueFlowIdentifier
    Public Property DataAction As Integer Implements COMInterfaces.IWorkflowElement.DataAction
    Public Property DataTableID As Integer Implements COMInterfaces.IWorkflowElement.DataTableID
    Public Property DataRecord As Integer Implements COMInterfaces.IWorkflowElement.DataRecord
    Public Property EmailID As Integer Implements COMInterfaces.IWorkflowElement.EmailID
    Public Property EmailRecord As Integer Implements COMInterfaces.IWorkflowElement.EmailRecord
    Public Property WebFormBGColor As Integer Implements COMInterfaces.IWorkflowElement.WebFormBGColor
    Public Property WebFormBGImageID As Integer Implements COMInterfaces.IWorkflowElement.WebFormBGImageID
    Public Property WebFormBGImageLocation As Integer Implements COMInterfaces.IWorkflowElement.WebFormBGImageLocation
    Public Property WebFormDefaultFontName As String Implements COMInterfaces.IWorkflowElement.WebFormDefaultFontName
    Public Property WebFormDefaultFontSize As Integer Implements COMInterfaces.IWorkflowElement.WebFormDefaultFontSize
    Public Property WebFormDefaultFontBold As Boolean Implements COMInterfaces.IWorkflowElement.WebFormDefaultFontBold
    Public Property WebFormDefaultFontItalic As Boolean Implements COMInterfaces.IWorkflowElement.WebFormDefaultFontItalic
    Public Property WebFormDefaultFontStrikeThru As Boolean Implements COMInterfaces.IWorkflowElement.WebFormDefaultFontStrikeThru
    Public Property WebFormDefaultFontUnderline As Boolean Implements COMInterfaces.IWorkflowElement.WebFormDefaultFontUnderline
    Public Property WebFormHeight As Integer Implements COMInterfaces.IWorkflowElement.WebFormHeight
    Public Property WebFormWidth As Integer Implements COMInterfaces.IWorkflowElement.WebFormWidth
    Public Property RecSelWebFormIdentifier As String Implements COMInterfaces.IWorkflowElement.RecSelWebFormIdentifier
    Public Property RecSelIdentifier As String Implements COMInterfaces.IWorkflowElement.RecSelIdentifier
    Public Property SecondaryDataRecord As Integer Implements COMInterfaces.IWorkflowElement.SecondaryDataRecord
    Public Property SecondaryRecSelWebFormIdentifier As String Implements COMInterfaces.IWorkflowElement.SecondaryRecSelWebFormIdentifier
    Public Property SecondaryRecSelIdentifier As String Implements COMInterfaces.IWorkflowElement.SecondaryRecSelIdentifier
    Public Property EmailSubject As String Implements COMInterfaces.IWorkflowElement.EmailSubject
    Public Property TimeoutFrequency As Integer Implements COMInterfaces.IWorkflowElement.TimeoutFrequency
    Public Property TimeoutPeriod As Integer Implements COMInterfaces.IWorkflowElement.TimeoutPeriod
    Public Property DataRecordTable As Integer Implements COMInterfaces.IWorkflowElement.DataRecordTable
    Public Property SecondaryDataRecordTable As Integer Implements COMInterfaces.IWorkflowElement.SecondaryDataRecordTable
    Public Property TrueFlowType As Integer Implements COMInterfaces.IWorkflowElement.TrueFlowType
    Public Property TrueFlowExprID As Integer Implements COMInterfaces.IWorkflowElement.TrueFlowExprID
    Public Property DescriptionExprID As Integer Implements COMInterfaces.IWorkflowElement.DescriptionExprID
    Public Property WebFormFGColor As Integer Implements COMInterfaces.IWorkflowElement.WebFormFGColor
    Public Property DescHasWorkflowName As Boolean Implements COMInterfaces.IWorkflowElement.DescHasWorkflowName
    Public Property DescHasElementCaption As Boolean Implements COMInterfaces.IWorkflowElement.DescHasElementCaption
    Public Property EmailCCID As Integer Implements COMInterfaces.IWorkflowElement.EmailCCID
    Public Property TimeoutExcludeWeekend As Boolean Implements COMInterfaces.IWorkflowElement.TimeoutExcludeWeekend
    Public Property CompletionMessageType As Integer Implements COMInterfaces.IWorkflowElement.CompletionMessageType
    Public Property CompletionMessage As String Implements COMInterfaces.IWorkflowElement.CompletionMessage
    Public Property SavedForLaterMessageType As Integer Implements COMInterfaces.IWorkflowElement.SavedForLaterMessageType
    Public Property SavedForLaterMessage As String Implements COMInterfaces.IWorkflowElement.SavedForLaterMessage
    Public Property FollowOnFormsMessageType As Integer Implements COMInterfaces.IWorkflowElement.FollowOnFormsMessageType
    Public Property FollowOnFormsMessage As String Implements COMInterfaces.IWorkflowElement.FollowOnFormsMessage
    Public Property Attachment_Type As String Implements COMInterfaces.IWorkflowElement.Attachment_Type
    Public Property Attachment_File As String Implements COMInterfaces.IWorkflowElement.Attachment_File
    Public Property Attachment_WFElementIdentifier As String Implements COMInterfaces.IWorkflowElement.Attachment_WFElementIdentifier
    Public Property Attachment_WFValueIdentifier As String Implements COMInterfaces.IWorkflowElement.Attachment_WFValueIdentifier
    Public Property Attachment_DBColumnID As Integer Implements COMInterfaces.IWorkflowElement.Attachment_DBColumnID
    Public Property Attachment_DBRecord As Integer Implements COMInterfaces.IWorkflowElement.Attachment_DBRecord
    Public Property Attachment_DBElement As String Implements COMInterfaces.IWorkflowElement.Attachment_DBElement
    Public Property Attachment_DBValue As String Implements COMInterfaces.IWorkflowElement.Attachment_DBValue

  End Class

End Namespace
