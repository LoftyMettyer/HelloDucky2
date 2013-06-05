Imports System.Runtime.InteropServices

Namespace Things

  <ClassInterface(ClassInterfaceType.None), ComVisible(True), Serializable()> _
  Public Class WorkflowElementItem
    Inherits Things.Base
    Implements iWorkflowElementItem

    Public Shadows Property ID As Integer Implements iWorkflowElementItem.ID
    Public Shadows Property Description As String Implements iWorkflowElementItem.Description
    Public Property ItemType As Integer Implements iWorkflowElementItem.ItemType
    Public Property Caption As String Implements iWorkflowElementItem.Caption
    Public Property DBColumnID As Integer Implements iWorkflowElementItem.DBColumnID
    Public Property DBRecord As Integer Implements iWorkflowElementItem.DBRecord
    Public Property InputReturnType As Integer Implements iWorkflowElementItem.InputReturnType
    Public Property InputSize As Integer Implements iWorkflowElementItem.InputSize
    Public Property InputDecimals As Integer Implements iWorkflowElementItem.InputDecimals
    Public Property InputIdentifier As String Implements iWorkflowElementItem.InputIdentifier
    Public Property InputDefault As String Implements iWorkflowElementItem.InputDefault
    Public Property WFFormIdentifier As String Implements iWorkflowElementItem.WFFormIdentifier
    Public Property WFValueIdentifier As String Implements iWorkflowElementItem.WFValueIdentifier
    Public Property Left As Integer Implements iWorkflowElementItem.Left
    Public Property Top As Integer Implements iWorkflowElementItem.Top
    Public Property Width As Integer Implements iWorkflowElementItem.Width
    Public Property Height As Integer Implements iWorkflowElementItem.Height
    Public Property BackgroundColor As Integer Implements iWorkflowElementItem.BackgroundColor
    Public Property ForegroundColor As Integer Implements iWorkflowElementItem.ForegroundColor
    Public Property FontName As String Implements iWorkflowElementItem.FontName
    Public Property FontSize As Integer Implements iWorkflowElementItem.FontSize
    Public Property FontBold As Boolean Implements iWorkflowElementItem.FontBold
    Public Property FontItalic As Boolean Implements iWorkflowElementItem.FontItalic
    Public Property FontStrikeThru As Boolean Implements iWorkflowElementItem.FontStrikeThru
    Public Property FontUnderline As Boolean Implements iWorkflowElementItem.FontUnderline
    Public Property PictureID As Integer Implements iWorkflowElementItem.PictureID
    Public Property PictureBorder As Integer Implements iWorkflowElementItem.PictureBorder
    Public Property Alignment As Integer Implements iWorkflowElementItem.Alignment
    Public Property ZOrder As Integer Implements iWorkflowElementItem.ZOrder
    Public Property TabIndex As Integer Implements iWorkflowElementItem.TabIndex
    Public Property BackStyle As Integer Implements iWorkflowElementItem.BackStyle
    Public Property BackColorEven As Integer Implements iWorkflowElementItem.BackColorEven
    Public Property BackColorOdd As Integer Implements iWorkflowElementItem.BackColorOdd
    Public Property ColumnHeaders As String Implements iWorkflowElementItem.ColumnHeaders
    Public Property ForeColorEven As Integer Implements iWorkflowElementItem.ForeColorEven
    Public Property ForeColorOdd As Integer Implements iWorkflowElementItem.ForeColorOdd
    Public Property HeaderBackColor As Integer Implements iWorkflowElementItem.HeaderBackColor
    Public Property HeadFontName As String Implements iWorkflowElementItem.HeadFontName
    Public Property HeadFontSize As Integer Implements iWorkflowElementItem.HeadFontSize
    Public Property HeadFontBold As Integer Implements iWorkflowElementItem.HeadFontBold
    Public Property HeadFontItalic As Integer Implements iWorkflowElementItem.HeadFontItalic
    Public Property HeadFontStrikeThru As Integer Implements iWorkflowElementItem.HeadFontStrikeThru
    Public Property HeadFontUnderline As Integer Implements iWorkflowElementItem.HeadFontUnderline
    Public Property Headlines As String Implements iWorkflowElementItem.Headlines
    Public Property TableID As Integer Implements iWorkflowElementItem.TableID
    Public Property ForeColorHighlight As Integer Implements iWorkflowElementItem.ForeColorHighlight
    Public Property BackColorHighlight As Integer Implements iWorkflowElementItem.BackColorHighlight
    Public Property ControlValues As String Implements iWorkflowElementItem.ControlValues
    Public Property LookupTableID As Integer Implements iWorkflowElementItem.LookupTableID
    Public Property LookupColumnID As Integer Implements iWorkflowElementItem.LookupColumnID
    Public Property RecordTableID As Integer Implements iWorkflowElementItem.RecordTableID
    Public Property Orientation As Integer Implements iWorkflowElementItem.Orientation
    Public Property RecordOrderID As Integer Implements iWorkflowElementItem.RecordOrderID
    Public Property RecordFilterID As Integer Implements iWorkflowElementItem.RecordFilterID
    Public Property Behaviour As Integer Implements iWorkflowElementItem.Behaviour
    Public Property Mandatory As Boolean Implements iWorkflowElementItem.Mandatory
    Public Property ExpressionID As Integer Implements iWorkflowElementItem.ExpressionID
    Public Property CaptionType As Integer Implements iWorkflowElementItem.CaptionType
    Public Property DefaultValueType As Integer Implements iWorkflowElementItem.DefaultValueType
    Public Property VerticalOffsetBehaviour As Integer Implements iWorkflowElementItem.VerticalOffsetBehaviour
    Public Property HorizontalOffsetBehaviour As Integer Implements iWorkflowElementItem.HorizontalOffsetBehaviour
    Public Property VerticalOffset As Integer Implements iWorkflowElementItem.VerticalOffset
    Public Property HorizontalOffset As Integer Implements iWorkflowElementItem.HorizontalOffset
    Public Property HeightBehaviour As Integer Implements iWorkflowElementItem.HeightBehaviour
    Public Property WidthBehaviour As Integer Implements iWorkflowElementItem.WidthBehaviour
    Public Property PasswordType As String Implements iWorkflowElementItem.PasswordType
    Public Property FileExtensions As String Implements iWorkflowElementItem.FileExtensions
    Public Property LookupFilterColumnID As Integer Implements iWorkflowElementItem.LookupFilterColumnID
    Public Property LookupFilterOperator As String Implements iWorkflowElementItem.LookupFilterOperator
    Public Property LookupFilterValue As String Implements iWorkflowElementItem.LookupFilterValue

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.WorkflowElementItem
      End Get
    End Property


  End Class
End Namespace
