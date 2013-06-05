Imports System.Runtime.InteropServices

Namespace Things

  <ClassInterface(ClassInterfaceType.None), ComVisible(True), Serializable()> _
  Public Class WorkflowElementItem
    Inherits Things.Base
    Implements IWorkflowElementItem

    Public Shadows Property ID As Integer Implements IWorkflowElementItem.ID
    Public Shadows Property Description As String Implements IWorkflowElementItem.Description
    Public Property ItemType As Integer Implements IWorkflowElementItem.ItemType
    Public Property Caption As String Implements IWorkflowElementItem.Caption
    Public Property DBColumnID As Integer Implements IWorkflowElementItem.DBColumnID
    Public Property DBRecord As Integer Implements IWorkflowElementItem.DBRecord
    Public Property InputReturnType As Integer Implements IWorkflowElementItem.InputReturnType
    Public Property InputSize As Integer Implements IWorkflowElementItem.InputSize
    Public Property InputDecimals As Integer Implements IWorkflowElementItem.InputDecimals
    Public Property InputIdentifier As String Implements IWorkflowElementItem.InputIdentifier
    Public Property InputDefault As String Implements IWorkflowElementItem.InputDefault
    Public Property WFFormIdentifier As String Implements IWorkflowElementItem.WFFormIdentifier
    Public Property WFValueIdentifier As String Implements IWorkflowElementItem.WFValueIdentifier
    Public Property Left As Integer Implements IWorkflowElementItem.Left
    Public Property Top As Integer Implements IWorkflowElementItem.Top
    Public Property Width As Integer Implements IWorkflowElementItem.Width
    Public Property Height As Integer Implements IWorkflowElementItem.Height
    Public Property BackgroundColor As Integer Implements IWorkflowElementItem.BackgroundColor
    Public Property ForegroundColor As Integer Implements IWorkflowElementItem.ForegroundColor
    Public Property FontName As String Implements IWorkflowElementItem.FontName
    Public Property FontSize As Integer Implements IWorkflowElementItem.FontSize
    Public Property FontBold As Boolean Implements IWorkflowElementItem.FontBold
    Public Property FontItalic As Boolean Implements IWorkflowElementItem.FontItalic
    Public Property FontStrikeThru As Boolean Implements IWorkflowElementItem.FontStrikeThru
    Public Property FontUnderline As Boolean Implements IWorkflowElementItem.FontUnderline
    Public Property PictureID As Integer Implements IWorkflowElementItem.PictureID
    Public Property PictureBorder As Integer Implements IWorkflowElementItem.PictureBorder
    Public Property Alignment As Integer Implements IWorkflowElementItem.Alignment
    Public Property ZOrder As Integer Implements IWorkflowElementItem.ZOrder
    Public Property TabIndex As Integer Implements IWorkflowElementItem.TabIndex
    Public Property BackStyle As Integer Implements IWorkflowElementItem.BackStyle
    Public Property BackColorEven As Integer Implements IWorkflowElementItem.BackColorEven
    Public Property BackColorOdd As Integer Implements IWorkflowElementItem.BackColorOdd
    Public Property ColumnHeaders As String Implements IWorkflowElementItem.ColumnHeaders
    Public Property ForeColorEven As Integer Implements IWorkflowElementItem.ForeColorEven
    Public Property ForeColorOdd As Integer Implements IWorkflowElementItem.ForeColorOdd
    Public Property HeaderBackColor As Integer Implements IWorkflowElementItem.HeaderBackColor
    Public Property HeadFontName As String Implements IWorkflowElementItem.HeadFontName
    Public Property HeadFontSize As Integer Implements IWorkflowElementItem.HeadFontSize
    Public Property HeadFontBold As Integer Implements IWorkflowElementItem.HeadFontBold
    Public Property HeadFontItalic As Integer Implements IWorkflowElementItem.HeadFontItalic
    Public Property HeadFontStrikeThru As Integer Implements IWorkflowElementItem.HeadFontStrikeThru
    Public Property HeadFontUnderline As Integer Implements IWorkflowElementItem.HeadFontUnderline
    Public Property Headlines As String Implements IWorkflowElementItem.Headlines
    Public Property TableID As Integer Implements IWorkflowElementItem.TableID
    Public Property ForeColorHighlight As Integer Implements IWorkflowElementItem.ForeColorHighlight
    Public Property BackColorHighlight As Integer Implements IWorkflowElementItem.BackColorHighlight
    Public Property ControlValues As String Implements IWorkflowElementItem.ControlValues
    Public Property LookupTableID As Integer Implements IWorkflowElementItem.LookupTableID
    Public Property LookupColumnID As Integer Implements IWorkflowElementItem.LookupColumnID
    Public Property RecordTableID As Integer Implements IWorkflowElementItem.RecordTableID
    Public Property Orientation As Integer Implements IWorkflowElementItem.Orientation
    Public Property RecordOrderID As Integer Implements IWorkflowElementItem.RecordOrderID
    Public Property RecordFilterID As Integer Implements IWorkflowElementItem.RecordFilterID
    Public Property Behaviour As Integer Implements IWorkflowElementItem.Behaviour
    Public Property Mandatory As Boolean Implements IWorkflowElementItem.Mandatory
    Public Property ExpressionID As Integer Implements IWorkflowElementItem.ExpressionID
    Public Property CaptionType As Integer Implements IWorkflowElementItem.CaptionType
    Public Property DefaultValueType As Integer Implements IWorkflowElementItem.DefaultValueType
    Public Property VerticalOffsetBehaviour As Integer Implements IWorkflowElementItem.VerticalOffsetBehaviour
    Public Property HorizontalOffsetBehaviour As Integer Implements IWorkflowElementItem.HorizontalOffsetBehaviour
    Public Property VerticalOffset As Integer Implements IWorkflowElementItem.VerticalOffset
    Public Property HorizontalOffset As Integer Implements IWorkflowElementItem.HorizontalOffset
    Public Property HeightBehaviour As Integer Implements IWorkflowElementItem.HeightBehaviour
    Public Property WidthBehaviour As Integer Implements IWorkflowElementItem.WidthBehaviour
    Public Property PasswordType As String Implements IWorkflowElementItem.PasswordType
    Public Property FileExtensions As String Implements IWorkflowElementItem.FileExtensions
    Public Property LookupFilterColumnID As Integer Implements IWorkflowElementItem.LookupFilterColumnID
    Public Property LookupFilterOperator As String Implements IWorkflowElementItem.LookupFilterOperator
    Public Property LookupFilterValue As String Implements IWorkflowElementItem.LookupFilterValue

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.WorkflowElementItem
      End Get
    End Property


  End Class
End Namespace
