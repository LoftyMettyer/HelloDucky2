Namespace Things
  Public Class Column
    Inherits Things.Base

    <System.Xml.Serialization.XmlIgnore()> _
        Public Table As Things.Table

    Public CalcID As HCMGuid
    <System.Xml.Serialization.XmlIgnore()> _
    Public Calculation As Things.Expression

    Public DataType As ScriptDB.ColumnTypes
    Public Size As Integer
    Public Decimals As Integer

    ' Options
    Public Audit As Boolean
    Public Multiline As Boolean
    Public CalculateIfEmpty As Boolean
    Public IsReadOnly As Boolean

    ' Formatting
    Public CaseType As Enums.CaseType
    Public TrimType As Enums.TrimType
    Public Alignment As Enums.AlignType
    Public Mandatory As Boolean
    Public OLEType As ScriptDB.OLEType

    Public DefaultCalcID As HCMGuid
    Public DefaultCalculation As Things.Expression
    Public DefaultValue As String

    Public ReferencedBy As Things.Collection

    Public Sub New()
      ReferencedBy = New Things.Collection
    End Sub

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.Column
      End Get
    End Property

    Public ReadOnly Property DataTypeSize As String
      Get
        Select Case CInt(Me.DataType)
          Case ScriptDB.ColumnTypes.Text
            If Me.Multiline Or Me.Size > 8000 Then
              Return "MAX"
            Else
              Return CStr(Me.Size)
            End If

          Case ScriptDB.ColumnTypes.Numeric
            Return (Me.Size + Me.Decimals).ToString

          Case ScriptDB.ColumnTypes.Logic
            Return "1"

          Case ScriptDB.ColumnTypes.Date
            Return "20"

          Case Else
            Return CStr(Me.Size)

        End Select
      End Get
    End Property

    Public ReadOnly Property SafeReturnType As String
      Get

        Dim sSQLType As String = String.Empty

        Select Case CInt(Me.DataType)
          Case ScriptDB.ColumnTypes.Text, ScriptDB.ColumnTypes.WorkingPattern, ScriptDB.ColumnTypes.Link
            sSQLType = "''"

          Case ScriptDB.ColumnTypes.Integer, ScriptDB.ColumnTypes.Numeric, ScriptDB.ColumnTypes.Logic
            sSQLType = "0"

          Case ScriptDB.ColumnTypes.Date
            sSQLType = "NULL"

          Case Else
            sSQLType = "0"

        End Select

        Return sSQLType

      End Get
    End Property

    ' Declaration syntax for a column type
    Public ReadOnly Property DataTypeSyntax As String
      Get

        Dim sSQLType As String = String.Empty

        Select Case CInt(Me.DataType)
          Case ScriptDB.ColumnTypes.Text
            If Me.Multiline Or Me.Size > 8000 Then
              sSQLType = "varchar(MAX)"
            Else
              sSQLType = String.Format("varchar({0})", Me.Size)
            End If

          Case ScriptDB.ColumnTypes.Integer
            sSQLType = String.Format("integer")

          Case ScriptDB.ColumnTypes.Numeric
            sSQLType = String.Format("numeric({0},{1})", Me.Size, Me.Decimals)

          Case ScriptDB.ColumnTypes.Date
            sSQLType = "datetime"

          Case ScriptDB.ColumnTypes.Logic
            sSQLType = "bit"

          Case ScriptDB.ColumnTypes.WorkingPattern
            sSQLType = "varchar(14)"

          Case ScriptDB.ColumnTypes.Link
            sSQLType = "varchar(255)"

          Case ScriptDB.ColumnTypes.Photograph
            sSQLType = "varchar(255)"

          Case ScriptDB.ColumnTypes.Binary
            sSQLType = "varbinary(MAX)"

        End Select

        Return sSQLType

      End Get

    End Property

    Public ReadOnly Property HasDefaultValue As Boolean
      Get

        If CInt(Me.DefaultCalcID) > 0 Then
          Return True
        Else
          Select Case Me.DataType
            Case ScriptDB.ColumnTypes.Text
              Return Len(Me.DefaultValue) > 0
            Case ScriptDB.ColumnTypes.Numeric, ScriptDB.ColumnTypes.Integer
              Return (CInt(Me.DefaultValue) <> 0)
            Case ScriptDB.ColumnTypes.Logic
              Return CBool(Me.DefaultValue <> True)
            Case Else
              Return False
          End Select
        End If

      End Get
    End Property


    Public ReadOnly Property IsCalculated As Boolean
      Get
        Return (CInt(CalcID) > 0)
      End Get
    End Property

    'Public ReadOnly Property BlankDefintion As String
    '  Get

    '    Dim sBlankDefintion As String = String.Empty

    '    sBlankDefintion = "NULL"

    '    'Select Case Me.DataType
    '    '  Case ScriptDB.ColumnTypes.Date
    '    '    sBlankDefintion = ""
    '    '  Case ScriptDB.ColumnTypes.Logic
    '    '    sBlankDefintion = "0"
    '    '    Case 
    '    'End Select

    '    Return sBlankDefintion

    '  End Get
    'End Property

    '#Region "XML"

    '    Public ReadOnly Property ToXML As String ' Implements Interfaces.iSystemObject.ToXML 'Xml.XmlDocument 
    '      Get

    '        '     Dim sXML As String

    '        Dim sb As New System.Text.StringBuilder
    '        Dim writer As Xml.XmlTextWriter = New Xml.XmlTextWriter(New System.IO.StringWriter(sb))
    '        'Dim returnXML As New Xml.Serialization.XmlSerializer(Me.GetType)

    '        Dim returnXML As New Xml.Serialization.XmlSerializer(Me.GetType)



    '        'dtExport.WriteXml(writer)
    '        'writer.Close()

    '        'sXML = Replace(sb.ToString, "<DocumentElement>", "")
    '        'sXML = Replace(sXML, "</DocumentElement>", "")

    '        'GetXMLFromDataTable = sXML



    '        returnXML.Serialize(writer, Me)
    '        writer.Close()
    '        Return sb.ToString

    '        '(returnXML, Me)



    '        'returnXML.Serialize(objStreamWriter, Me)
    '        'objStreamWriter.Close()
    '        ' Return
    '      End Get
    '    End Property
    '#End Region

    Public Function ApplyFormatting() As String
      Return ApplyFormatting(String.Empty)
    End Function

    Public Function ApplyFormatting(ByRef Prefix As String) As String

      Dim iFormats As Integer = 0
      Dim sFormat As String = Me.Name

      If Not Prefix = String.Empty Then
        sFormat = String.Format("[{0}].[{1}]", Prefix, sFormat)
      End If

      If Me.DataType = ScriptDB.ColumnTypes.Text Then

        ' Case
        Select Case Me.CaseType
          Case Enums.CaseType.Lower
            sFormat = String.Format("LOWER({0})", sFormat)
          Case Enums.CaseType.Proper
            sFormat = String.Format("dbo.udfsys_propercase({0})", sFormat)
          Case Enums.CaseType.Upper
            sFormat = String.Format("UPPER({0})", sFormat)
        End Select

        'Select Case Me.Alignment
        '  Case AlignType.Center
        '  Case AlignType.Left
        '  Case AlignType.Right
        'End Select

        ' Trim type
        Select Case Me.TrimType
          Case Enums.TrimType.Both
            sFormat = String.Format("LTRIM(RTRIM({0}))", sFormat)
          Case Enums.TrimType.Left
            sFormat = String.Format("LTRIM({0})", sFormat)
          Case Enums.TrimType.Right
            sFormat = String.Format("RTRIM({0})", sFormat)
        End Select

      End If

      Return sFormat
    End Function

  End Class
End Namespace

