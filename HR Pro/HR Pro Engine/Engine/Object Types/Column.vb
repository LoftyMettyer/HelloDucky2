Namespace Things
  Public Class Column
    Inherits Things.Base

    <System.Xml.Serialization.XmlIgnore()> _
        Public Table As Things.Table

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

    ' Validation
    Public Mandatory As Boolean


    Public OLEType As ScriptDB.OLEType

    Public CalcID As HCMGuid
    Public DefaultCalcID As HCMGuid

    '    Public DirectInView As Boolean

    'Public Overrides Function Commit() As Boolean
    'End Function

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.Column
      End Get
    End Property

    'Public ReadOnly Property ApplyFormatting As String
    '  Get

    '    Select Case Me.DataType
    '      Case ScriptDB.ColumnTypes.Date
    '        Return String.Format("DATEADD(dd, 0, DATEDIFF(dd, 0, [inserted].[{0}]))", Me.Name)

    '      Case Else
    '        Return Me.Name

    '    End Select

    '  End Get
    'End Property


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
              sSQLType = "[varchar](MAX)"
            Else
              sSQLType = String.Format("[varchar]({0})", Me.Size)
            End If

          Case ScriptDB.ColumnTypes.Integer
            sSQLType = String.Format("[integer]")

          Case ScriptDB.ColumnTypes.Numeric
            sSQLType = String.Format("[numeric]({0},{1})", Me.Size, Me.Decimals)

          Case ScriptDB.ColumnTypes.Date
            sSQLType = "[datetime]"

          Case ScriptDB.ColumnTypes.Logic
            sSQLType = "[bit]"

          Case ScriptDB.ColumnTypes.WorkingPattern
            sSQLType = "[varchar](14)"

          Case ScriptDB.ColumnTypes.Link
            sSQLType = "[varchar](255)"

          Case ScriptDB.ColumnTypes.Photograph
            sSQLType = "[varchar](255)"

          Case ScriptDB.ColumnTypes.Binary
            sSQLType = "[varbinary](MAX)"

        End Select

        Return sSQLType

      End Get

    End Property


    Public ReadOnly Property IsCalculated As Boolean
      Get
        'Return Not (CalcID = 0)

        'botch so we only calculate the age field
        ' 40 = age
        ' 70 = holiday balance
        '
        ' 12307 = 
        ' 14102 = Holiday_Taken_Simple
        ' 10181 = Six Month Proabion Date
        ' 3153 = HR Pro number
        ' 12708 = Fleet Drivers. Vlaue When New
        ' 6025 - Risk Assesments.Last Updated By
        ' 4762 - PR.Continuous Service Years
        ' 5079 - PR.Actual hours
        ' 62 - Annual Rate
        ' 11938 - PR.Professional Membership
        ' 3196 - Is leaver
        ' 154 - NVQ.Unit_1_Recommended_Finish
        '64 - PR.Weekly Rate
        '        Return (CalcID = 70 Or CalcID = 40 Or CalcID = 41 Or CalcID = 12307 Or CalcID = 14102)
        ' 11036 Project_records.Logical_True
        ' 193 - Absence.Duration_Days
        ' 72 - PR.Holiday_Taken
        ' 10944 - PR.Is QAS Role
        ' 4328 - Training_Requests.Last Updated By
        ' 4803 - PR.Payroll_Town
        ' 131 - Loans.Outstanding Balance
        ' 3237 - COurse recs.Is Current Course
        ' 3193 - PR.Is current Employee
        ' 14108 - Adoption. Not_True
        ' 14109 - Adoption. Not_True 2
        ' 14110 - Adoption. Not True 3
        ' 14111 - PR.Site2
        ' 2500 - Training_Booking.CPD_Accredited
        ' 3776 - McFarlanes.Course_Records.Number Booked


        Return (CInt(CalcID) > 0)
        'Return (CalcID = 3776)

      End Get
    End Property

    Public ReadOnly Property BlankDefintion As String
      Get

        Dim sBlankDefintion As String = String.Empty

        sBlankDefintion = "NULL"

        'Select Case Me.DataType
        '  Case ScriptDB.ColumnTypes.Date
        '    sBlankDefintion = ""
        '  Case ScriptDB.ColumnTypes.Logic
        '    sBlankDefintion = "0"
        '    Case 
        'End Select

        Return sBlankDefintion

      End Get
    End Property

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

