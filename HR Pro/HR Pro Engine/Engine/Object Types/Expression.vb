Option Explicit On

Namespace Things

  <Serializable()> _
  Public Class Expression
    Inherits Things.Component

    Public Size As Integer
    Public Decimals As Integer
    'Public DataType As ScriptDB.ColumnTypes
    Public BaseTableID As HCMGuid

    <System.Xml.Serialization.XmlIgnore()> _
    Public BaseTable As Things.Table

    Public UDF As ScriptDB.GeneratedUDF
    Public ExpressionType As ScriptDB.ExpressionType
    'Public GenerateType As ScriptDB.GenerateType = ScriptDB.GenerateType.ComplexUDF
    Public IsDeterministic As Boolean = True

    <System.Xml.Serialization.XmlIgnore()> _
    Public AssociatedColumn As Things.Column

    Private mcolDependencies As New Things.Collection

    Private mcolOrders As Things.Collection
    '    Private mcolFilters As Things.Collection

    ' Private maryAvoidRecursion As ArrayList
    Public Filters As ArrayList
    Public Joins As ArrayList
    Public FromTables As ArrayList
    Private maryWhere As ArrayList
    Private maryDeclarations As ArrayList
    Private maryPrerequisitStatements As ArrayList

    Private mcolLinesOfCode As ScriptDB.LinesOfCode
    'Private mbAddBaseTable As Boolean
    Private mbIsValid As Boolean
    Private mbRequiresRowNumber As Boolean = False
    Private mbCalculatePostAudit As Boolean = False

    'tempry - must solve later
    Private miRecuriveStop As Integer
    Private OnlyReferencesThisTable As Boolean

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.Expression
      End Get
    End Property

    Public ReadOnly Property CalculatePostAudit As Boolean
      Get
        Return mbCalculatePostAudit
      End Get
    End Property

#Region "Generate code"

    '' Any columns that are calculated should be amended to embed their respective calculations
    'Private Function ChangeCalculatedColumnsToExpressions(ByRef objExpression As Things.Component) As Things.Collection

    '  Dim objColumn As Things.Column
    '  Dim objComponent As Things.Component
    '  Dim objReplaceComponent As Things.Component
    '  Dim objNew As New Things.Collection '= objExpression.Clone
    '  Dim bIsProcessed As Boolean = False
    '  Dim objDepends As Things.Base

    '  'objExpression.Objects.Clear()
    '  miRecuriveStop = miRecuriveStop + 1
    '  If miRecuriveStop > 150 Then
    '    Debug.Print(objExpression.Name)
    '    '    Return objExpression.Objects
    '  End If


    '  For Each objComponent In objExpression.Objects
    '    Select Case objComponent.SubType

    '      Case ScriptDB.ComponentTypes.Column

    '        objColumn = Globals.Things.GetObject(Enums.Type.Table, objComponent.TableID).Objects.GetObject(Enums.Type.Column, objComponent.ColumnID)

    '        For Each objDepends In mcolDependencies
    '          If objDepends.Type = Enums.Type.Column Then
    '            If CType(objDepends, Things.Column).ID = objColumn.ID Then
    '              bIsProcessed = True
    '              Exit For
    '            End If
    '          End If
    '        Next

    '        If Not bIsProcessed Then
    '          mcolDependencies.Add(objColumn)
    '        End If

    '        '   Debug.Assert(Not objComponent.ColumnID = 985)


    '        If objColumn.IsCalculated And Not objColumn Is Me.AssociatedColumn And Not bIsProcessed Then
    '          objReplaceComponent = Globals.Things.GetObject(Enums.Type.Table, objComponent.TableID).Objects.GetObject(Enums.Type.Expression, objColumn.CalcID)
    '          objReplaceComponent.BaseExpression = objComponent.BaseExpression
    '          objReplaceComponent.SubType = ScriptDB.ComponentTypes.Expression
    '          objReplaceComponent.Objects = ChangeCalculatedColumnsToExpressions(objReplaceComponent)
    '          objNew.Add(objReplaceComponent)
    '        Else
    '          objNew.Add(objComponent)
    '        End If

    '        '' Add to the dependancy stack
    '        'If Not bIsProcessed Then
    '        '  mcolDependencies.Add(objColumn)
    '        'End If
    '        '     Case ScriptDB.ComponentTypes.Calculation

    '        'For Each objDepends In mcolDependencies
    '        '  If objDepends.Type = Enums.Type.Component Then
    '        '    If CType(objDepends, Things.Component).ID = objComponent.ID Then
    '        '      bIsProcessed = True
    '        '      Exit For
    '        '    End If
    '        '  End If
    '        'Next

    '        'If Not bIsProcessed Then
    '        '  objComponent.Objects = ChangeCalculatedColumnsToExpressions(objComponent)
    '        'End If
    '        'objNew.Add(objComponent)


    '      Case ScriptDB.ComponentTypes.Expression, ScriptDB.ComponentTypes.Function, ScriptDB.ComponentTypes.Calculation

    '        BuildDependancies(objComponent)

    '        'For Each objDepends In mcolDependencies
    '        '  If objDepends.Type = objComponent.Type Then
    '        '    If CType(objDepends, Things.Component).ID = objComponent.ID Then
    '        '      bIsProcessed = True
    '        '      Exit For
    '        '    End If
    '        '  End If
    '        'Next

    '        'If Not bIsProcessed Then
    '        '  mcolDependencies.Add(objComponent)
    '        '  objComponent.Objects = ChangeCalculatedColumnsToExpressions(objComponent)
    '        'Else

    '        'End If
    '        'objComponent.SubType = ScriptDB.ComponentTypes.Expression
    '        '     End If

    '        objNew.Add(objComponent)

    '        'mcolDependencies.Add(objComponent)
    '        '       Else
    '        '     objNew.Add(objComponent)

    '        '    Debug.Print("hhhe")
    '        '    End If

    '      Case Else
    '        objNew.Add(objComponent)

    '    End Select

    '  Next

    '  Return objNew

    'End Function

    Private Sub AddDependancy(ByRef Thing As Things.Base)

      'If Not mcolDependencies.Contains(Thing) Then
      '  mcolDependencies.Add(Thing)
      'End If

    End Sub


    Private Sub BuildDependancies(ByRef objExpression As Things.Component)

      'Dim objColumn As Things.Column
      Dim objComponent As Things.Component
      '      Dim objDependency As Things.Base
      Dim objColumn As Things.Column

      For Each objComponent In objExpression.Objects
        Select Case objComponent.SubType
          Case ScriptDB.ComponentTypes.Column

            objColumn = Globals.Things.GetObject(Enums.Type.Table, objComponent.TableID).Objects.GetObject(Enums.Type.Column, objComponent.ColumnID)
            'If objColumn.IsCalculated Then
            '  objComponent = Globals.Things.GetObject(Enums.Type.Table, objColumn.Table.ID).Objects.GetObject(Enums.Type.Expression, objColumn.CalculationID)
            '  BuildDependancies(objComponent)
            'End If

            If Not mcolDependencies.Contains(objColumn) Then
              mcolDependencies.Add(objColumn)
            End If

          Case ScriptDB.ComponentTypes.Expression, ScriptDB.ComponentTypes.Function, ScriptDB.ComponentTypes.Calculation
            BuildDependancies(objComponent)
        End Select

      Next

    End Sub

    Public Overridable Sub GenerateCode()

      ' Dim objColumn As Things.Column
      Dim objDependency As Things.Base
      Dim sOptions As String = ""
      Dim sCode As String = ""
      Dim aryParameters1 As New ArrayList
      Dim aryParameters2 As New ArrayList
      Dim aryParameters3 As New ArrayList
      Dim iCount As Integer

      sOptions = IIf(Me.Encrypted, "WITH ENCYPTION", "")

      ' Initialise code object
      mcolLinesOfCode = New ScriptDB.LinesOfCode
      mcolLinesOfCode.Clear()
      mcolLinesOfCode.ReturnType = ReturnType
      mcolLinesOfCode.CodeLevel = IIf(Me.ExpressionType = ScriptDB.ExpressionType.ColumnFilter, 2, 1)

      Joins = New ArrayList
      Filters = New ArrayList
      FromTables = New ArrayList
      maryWhere = New ArrayList

      maryDeclarations = New ArrayList
      maryPrerequisitStatements = New ArrayList
      '    maryAvoidRecursion = New ArrayList

      '   maryAvoidRecursion.Clear()
      Joins.Clear()
      maryWhere.Clear()
      maryPrerequisitStatements.Clear()
      maryDeclarations.Clear()


      ' Build the dependencies collection
      mcolDependencies.Clear()
      miRecuriveStop = 0
      BuildDependancies(Me)

      '   BuildDependancies(Me)
      aryParameters1.Clear()
      aryParameters2.Clear()
      aryParameters3.Clear()

      ' Build the executeion code
      If Me.Objects.Count = 0 Then
        mbIsValid = False
      Else
        SQLCode_AddCodeLevel(Me.Objects, mcolLinesOfCode)
      End If

      ' Always add the ID for the record
      aryParameters1.Add("@prm_ID integer")
      aryParameters2.Add("base.ID")
      aryParameters3.Add("@prm_ID")

      If mbRequiresRowNumber Then
        aryParameters1.Add("@rownumber integer")
        aryParameters2.Add("[rownumber]")
        aryParameters3.Add("@rownumber")
      End If

      ' Add other dependancies
      iCount = 0
      For Each objDependency In mcolDependencies
        If objDependency.Type = Enums.Type.Column Then
          If CType(objDependency, Things.Column).Table Is Me.BaseTable Then
            aryParameters1.Add(String.Format("@prm_{0} {1}", objDependency.Name, CType(objDependency, Things.Column).DataTypeSyntax))
            aryParameters2.Add(String.Format("base.[{0}]", objDependency.Name))
            aryParameters3.Add(String.Format("@prm_{0}", objDependency.Name))
          End If
        End If

        If objDependency.Type = Enums.Type.Relation Then

          If Not aryParameters1.Contains(String.Format("@pid_{0} integer", CInt(CType(objDependency, Things.Relation).ParentID))) Then
            aryParameters1.Add(String.Format("@pid_{0} integer", CInt(CType(objDependency, Things.Relation).ParentID)))

            If CType(objDependency, Things.Relation).RelationshipType = ScriptDB.RelationshipType.Parent Then
              aryParameters2.Add(String.Format("base.[ID_{0}]", CInt(CType(objDependency, Things.Relation).ParentID)))
              aryParameters3.Add(String.Format("@prm_{0}]", CInt(CType(objDependency, Things.Relation).ParentID)))
            Else
              aryParameters2.Add("base.[ID]")
              aryParameters3.Add(String.Format("@prm_ID"))
            End If

          End If

        End If

        '  If objDependency.Type = Enums.Type.Setting Then
        '   '          aryParameters1.Add(String.Format("{0}", CInt(CType(objDependency, Things.Setting).Value)))
        '  aryParameters2.Add(String.Format("{0}", CInt(CType(objDependency, Things.Setting).Value)))
        '  End If

        iCount += iCount
      Next


      ' Special parameters for functions
      'If Me.BaseExpression.RequiresRowNumber Then
      '  aryParameters1.Add(String.Format("@base_rownumber integer"))
      '  aryParameters2.Add("ROW_NUMBER() OVER(OVER(ORDER BY base.[ID])")
      'End If


      ' Calling statement
      With UDF

        .Declarations = String.Join(vbNewLine, maryDeclarations.ToArray())
        .Prerequisites = String.Join(vbNewLine, maryPrerequisitStatements.ToArray())
        .JoinCode = String.Format("{0}", String.Join(vbNewLine, Joins.ToArray))
        .FromCode = String.Format("{0}", String.Join(",", FromTables.ToArray))

        .WhereCode = String.Join(vbNewLine, maryWhere.ToArray())
        .WhereCode = IIf(Len(.WhereCode) > 0, "WHERE " + .WhereCode, "")

        Select Case Me.ExpressionType

          ' Wrapper for calculations with associated columns
          Case ScriptDB.ExpressionType.ColumnCalculation
            .Name = String.Format("[{0}].[{1}{2}.{3}]", Me.SchemaName, ScriptDB.Consts.CalculationUDF, Me.AssociatedColumn.Table.Name, Me.AssociatedColumn.Name)


            '  If Me.ReturnType = ScriptDB.ComponentValueTypes.Logic Then
            '    .SelectCode = String.Format("CASE WHEN ({0}) THEN 1 ELSE 0 END", mcolLinesOfCode.Statement)
            '  Else
            .SelectCode = mcolLinesOfCode.Statement
            'End If


            '            SELECT @Result = (CASE WHEN 
            'dbo.udfsys_fieldchangedbetweentwodates((('00000001-00000009')), (('2011-05-01')), (('2011-05-31')), @prm_ID)
            ' THEN 1 ELSE 0 END)



            '            .SelectCode = mcolLinesOfCode.Statement
            .CallingCode = String.Format("{0}({1})", .Name, String.Join(",", aryParameters2.ToArray))
            '        .Where = String.Format("{0}", String.Join(",", Where.ToArray))

            .Code = String.Format("CREATE FUNCTION {0}({1})" & vbNewLine & _
                           "RETURNS {2}" & vbNewLine & _
                           "{3}" & vbNewLine & _
                           "--WITH SCHEMABINDING" & vbNewLine &
                           "AS" & vbNewLine & "BEGIN" & vbNewLine & _
                           "    DECLARE @Result as {2};" & vbNewLine & vbNewLine & _
                           "{4}" & vbNewLine & vbNewLine & _
                           "{5}" & vbNewLine & vbNewLine & _
                           "    -- Execute calculation code" & vbNewLine & _
                           "SELECT @Result = {6}" & vbNewLine & _
                           "                 {7}" & vbNewLine & _
                           "                 {8}" & vbNewLine & _
                           "                 {9}" & vbNewLine & _
                           "    RETURN ISNULL(@Result, {10});" & vbNewLine & _
                           "END" _
                          , .Name, String.Join(", ", aryParameters1.ToArray()) _
                          , Me.AssociatedColumn.DataTypeSyntax, sOptions, .Declarations, .Prerequisites, .SelectCode, .FromCode, .JoinCode, .WhereCode _
                          , Me.AssociatedColumn.SafeReturnType)


            .CodeStub = String.Format("CREATE FUNCTION {0}({1})" & vbNewLine & _
                           "RETURNS {2}" & vbNewLine & _
                           "{3}" & vbNewLine & _
                           "--WITH SCHEMABINDING" & vbNewLine &
                           "AS" & vbNewLine & "BEGIN" & vbNewLine & _
                           "    DECLARE @Result as {2};" & vbNewLine & vbNewLine & _
                           "-- Could not generate this procedure. " & vbNewLine & _
                           "/*" & vbNewLine & _
                           "{4}" & vbNewLine & vbNewLine & _
                           "{5}" & vbNewLine & vbNewLine & _
                           "    -- Execute calculation code" & vbNewLine & _
                           "SELECT @Result = {6}" & vbNewLine & _
                           "                 {7}" & vbNewLine & _
                           "                 {8}" & vbNewLine & _
                           "                 {9}" & vbNewLine & vbNewLine & _
                           "*/" & vbNewLine & _
                           "    RETURN ISNULL(@Result, {10});" & vbNewLine & _
                           "END" _
                          , .Name, String.Join(", ", aryParameters1.ToArray()) _
                          , Me.AssociatedColumn.DataTypeSyntax, sOptions, .Declarations, .Prerequisites, .SelectCode, .FromCode, .JoinCode, .WhereCode _
                          , Me.AssociatedColumn.SafeReturnType)

            ' Wrapper for when this function is used as a filter in an expression
          Case ScriptDB.ExpressionType.ColumnFilter
            mcolLinesOfCode.IsEvaluated = True
            .Name = String.Format("[{0}].[{1}{2}.{3}]", Me.SchemaName, ScriptDB.Consts.CalculationUDF, Me.AssociatedColumn.Table.Name, Me.AssociatedColumn.Name)
            .CallingCode = String.Format("{0}({1})", .Name, String.Join(",", aryParameters2.ToArray))
            '.SelectCode = String.Format("CASE WHEN ({0}) THEN 1 ELSE 0 END", mcolLinesOfCode.Statement)
            .SelectCode = mcolLinesOfCode.Statement

            ' Wrapper for when expression is used as a filter in a view
          Case ScriptDB.ExpressionType.Mask
            .Name = String.Format("[{0}].[{1}{2}]", Me.SchemaName, ScriptDB.Consts.MaskUDF, CInt(Me.BaseExpression.ID))
            .CallingCode = String.Format("{0}({1})", .Name, String.Join(",", aryParameters1.ToArray))
            mcolLinesOfCode.IsEvaluated = True
            .SelectCode = mcolLinesOfCode.Statement

            .Code = String.Format("CREATE FUNCTION {0}(@prm_id integer)" & vbNewLine & _
                           "RETURNS bit" & vbNewLine & _
                           "AS" & vbNewLine & "BEGIN" & vbNewLine & _
                           "    DECLARE @Result AS bit;" & vbNewLine & vbNewLine & _
                           "{4}" & vbNewLine & vbNewLine & _
                           "{5}" & vbNewLine & vbNewLine & _
                           "    -- Execute calculation code" & vbNewLine & _
                           "SELECT @Result = {6}" & vbNewLine & _
                           "                 {7}" & vbNewLine & _
                           "                 {8}" & vbNewLine & _
                           "                 {9}" & vbNewLine & _
                           "    RETURN ISNULL(@Result, 0);" & vbNewLine & _
                           "END" _
                          , .Name, String.Join(", ", aryParameters1.ToArray()) _
                          , "", "", .Declarations, .Prerequisites, .SelectCode, .FromCode, .JoinCode, .WhereCode)

            '                           "SELECT @Result = CASE WHEN ({6}) THEN 1 ELSE 0 END" & vbNewLine & _


            .CodeStub = String.Format("CREATE FUNCTION {0}(@prm_id integer)" & vbNewLine & _
                           "RETURNS bit" & vbNewLine & _
                           "AS" & vbNewLine & "BEGIN" & vbNewLine & _
                           "    DECLARE @Result AS bit;" & vbNewLine & vbNewLine & _
                           "/*{4}" & vbNewLine & vbNewLine & _
                           "{5}" & vbNewLine & vbNewLine & _
                           "    -- Execute calculation code" & vbNewLine & _
                           "SELECT @Result = CASE WHEN ({6}) THEN 1 ELSE 0 END" & vbNewLine & _
                           "                 {7}" & vbNewLine & _
                           "                 {8}" & vbNewLine & _
                           "                 {9}*/" & vbNewLine & _
                           "    RETURN ISNULL(@Result, 1);" & vbNewLine & _
                           "END" _
                          , .Name, String.Join(", ", aryParameters1.ToArray()) _
                          , "", "", .Declarations, .Prerequisites, .SelectCode, .FromCode, .JoinCode, .WhereCode)


          Case ScriptDB.ExpressionType.ReferencedColumn
            .Name = String.Format("[{0}].[{1}{2}.{3}]", Me.SchemaName, ScriptDB.Consts.CalculationUDF, Me.AssociatedColumn.Table.Name, Me.AssociatedColumn.Name)
            .CallingCode = String.Format("{0}({1})", .Name, String.Join(",", aryParameters3.ToArray))
            .SelectCode = mcolLinesOfCode.Statement

          Case ScriptDB.ExpressionType.RecordDescription
            .Name = String.Format("[{0}].[{1}{2}]", Me.SchemaName, ScriptDB.Consts.RecordDescriptionUDF, Me.BaseTable.Name)
            .CallingCode = String.Format("{0}({1})", .Name, String.Join(",", aryParameters2.ToArray))
            .SelectCode = mcolLinesOfCode.Statement

            .Code = String.Format("CREATE FUNCTION {0}({1})" & vbNewLine & _
                           "RETURNS nvarchar(MAX)" & vbNewLine & _
                           "{3}" & vbNewLine & _
                           "--WITH SCHEMABINDING" & vbNewLine &
                           "AS" & vbNewLine & "BEGIN" & vbNewLine & _
                           "    DECLARE @Result AS nvarchar(MAX);" & vbNewLine & vbNewLine & _
                           "{4}" & vbNewLine & vbNewLine & _
                           "{5}" & vbNewLine & vbNewLine & _
                           "    -- Execute calculation code" & vbNewLine & _
                           "SELECT @Result = {6}" & vbNewLine & _
                           "                 {7}" & vbNewLine & _
                           "                 {8}" & vbNewLine & _
                           "                 {9}" & vbNewLine & _
                           "    RETURN ISNULL(@Result, '');" & vbNewLine & _
                           "END" _
                          , .Name, String.Join(", ", aryParameters1.ToArray()) _
                          , "", sOptions, .Declarations, .Prerequisites, .SelectCode, .FromCode, .JoinCode, .WhereCode)

            .CodeStub = String.Format("CREATE FUNCTION {0}({1})" & vbNewLine & _
                           "RETURNS nvarchar(MAX)" & vbNewLine & _
                           "{3}" & vbNewLine & _
                           "--WITH SCHEMABINDING" & vbNewLine &
                           "AS" & vbNewLine & "BEGIN" & vbNewLine & _
                           "    DECLARE @Result AS nvarchar(MAX);" & vbNewLine & vbNewLine & _
                           "-- Could not generate this procedure. " & vbNewLine & _
                           "/*" & vbNewLine & _
                           "{4}" & vbNewLine & vbNewLine & _
                           "{5}" & vbNewLine & vbNewLine & _
                           "    -- Execute calculation code" & vbNewLine & _
                           "SELECT @Result = {6}" & vbNewLine & _
                           "                 {7}" & vbNewLine & _
                           "                 {8}" & vbNewLine & _
                           "                 {9}" & vbNewLine & vbNewLine & _
                           "*/" & vbNewLine & _
                           "    RETURN ISNULL(@Result, '');" & vbNewLine & _
                           "END" _
                          , .Name, String.Join(", ", aryParameters1.ToArray()) _
                          , "", sOptions, .Declarations, .Prerequisites, .SelectCode, .FromCode, .JoinCode, .WhereCode)


            ' Should never be called, but just in case...
          Case Else
            .SelectCode = mcolLinesOfCode.Statement

        End Select




      End With




      ' MONDAY COMMENT OUT - TEST!!!
      'Select Case ReturnType
      '  Case ScriptDB.ComponentValueTypes.Logic
      '    .SelectCode = String.Format("CASE WHEN ({0}) THEN 1 ELSE 0 END", mcolLinesOfCode.Statement) 'working when no if then else
      '    '.SelectCode = String.Format("convert(bit, {0})", mcolLinesOfCode.Statement)
      '    '.SelectCode = mcolLinesOfCode.Statement

      '  Case Else
      '    .SelectCode = mcolLinesOfCode.Statement
      'End Select




      ' .JoinCode = String.Join(vbNewLine, Joins.ToArray())
      '  .FromCode = String.Format(" FROM [{0}].[{1}]", Me.BaseTable.SchemaName, Me.BaseTable.PhysicalName)
      ' .WhereCode = String.Format(" WHERE [{0}].[id] = @pID", Me.BaseTable.PhysicalName)



    End Sub

    Private Sub SQLCode_AddCodeLevel(ByRef [Components] As Things.Collection, ByRef [CodeCluster] As ScriptDB.LinesOfCode)

      'Dim sFilter As String
      Dim objComponent As Things.Component

      Dim guiObjectID As HCMGuid
      '   Dim iValueDataType As ScriptDB.ComponentValueTypes

      Dim iValueSubType As ScriptDB.ColumnTypes = Nothing

      Dim LineOfCode As ScriptDB.CodeElement
      Dim objCalculation As Things.Expression

      'iReturnDataType = drObject.Item("DataType").ToString

      '    mbAddBaseTable = False
      '    sSQLFrom = 

      For Each objComponent In [Components]
        guiObjectID = objComponent.ID

        Select Case objComponent.SubType

          ' A table relationship
          Case ScriptDB.ComponentTypes.Relation
            SQLCode_AddRelation([CodeCluster], objComponent)

            ' Column component
          Case ScriptDB.ComponentTypes.Column
            SQLCode_AddColumn([CodeCluster], objComponent)

            ' Operator component
          Case ScriptDB.ComponentTypes.Operator
            SQLCode_AddOperator(objComponent, [CodeCluster])

            ' Value component
          Case ScriptDB.ComponentTypes.Value, ScriptDB.ComponentTypes.TableValue

            LineOfCode.CodeType = ScriptDB.ComponentTypes.Value

            Select Case objComponent.ValueType
              Case ScriptDB.ComponentValueTypes.Numeric
                LineOfCode.Code = String.Format("{0}", objComponent.ValueNumeric)

              Case ScriptDB.ComponentValueTypes.String
                LineOfCode.Code = String.Format("'{0}'", objComponent.ValueString)

              Case ScriptDB.ComponentValueTypes.Date
                LineOfCode.Code = String.Format("'{0}'", objComponent.ValueDate.ToString("yyyy-MM-dd"))

              Case ScriptDB.ComponentValueTypes.SystemVariable
                LineOfCode.Code = String.Format("{0}", objComponent.ValueString)

              Case Else
                LineOfCode.Code = String.Format("{0}", IIf(objComponent.ValueLogic, 1, 0))

            End Select

            [CodeCluster].Add(LineOfCode)


            ' Function component
          Case ScriptDB.ComponentTypes.Function
            SQLCode_AddFunction(objComponent, [CodeCluster])

            ' An expression or a parameter
          Case ScriptDB.ComponentTypes.Expression
            SQLCode_AddParameter(objComponent, [CodeCluster])

            ' Expression 
          Case ScriptDB.ComponentTypes.Calculation, ScriptDB.ComponentTypes.Expression

            If Not objComponent.BaseExpression.BaseTable.Objects.GetObject(Enums.Type.Expression, objComponent.CalculationID) Is Nothing Then

              ' There has to be a cleaner way, but once again I'm in  a hurry and this DOES work. Amazingly!
              objCalculation = CType(objComponent.BaseExpression.BaseTable.Objects.GetObject(Enums.Type.Expression, objComponent.CalculationID), Things.Expression).Clone
              objCalculation.SetBaseExpression(objComponent.BaseExpression)
              objComponent.BypassValidation = True
              objComponent.Objects = objCalculation.Objects
              objComponent.ReturnType = objCalculation.ReturnType
              SQLCode_AddParameter(objComponent, [CodeCluster])
            Else
              Globals.ErrorLog.Add(ErrorHandler.Section.General, Me.AssociatedColumn.Name, ErrorHandler.Severity.Error, "SQLCode_AddCodeLevel", "can't find expression")

            End If

        End Select

      Next

    End Sub

    Private Sub SQLCode_AddRelation(ByRef [CodeCluster] As ScriptDB.LinesOfCode, ByRef [Component] As Things.Component)

      Dim objTable As Things.Table
      Dim objRelation As Things.Relation
      Dim LineOfCode As ScriptDB.CodeElement

      LineOfCode.CodeType = ScriptDB.ComponentTypes.Relation

      objTable = Globals.Things.GetObject(Enums.Type.Table, [Component].TableID)
      objRelation = AssociatedColumn.Table.GetRelation(objTable.ID)

      If Not mcolDependencies.Contains(objRelation) Then
        mcolDependencies.Add(objRelation)
      End If

      LineOfCode.Code = String.Format("@pid_{0}", CInt([Component].TableID))

      [CodeCluster].Add(LineOfCode)

    End Sub


    Private Sub SQLCode_AddColumn(ByRef [CodeCluster] As ScriptDB.LinesOfCode, ByRef [Component] As Things.Component)

      '      Dim drColumn As System.Data.DataRow
      Dim objThisColumn As Things.Column
      Dim objBaseColumn As Things.Column

      Dim objExpression As Things.Expression
      Dim ChildCodeCluster As ScriptDB.LinesOfCode

      Dim objRelation As Things.Relation
      Dim sRelationCode As String
      Dim sFromCode As String
      Dim sWhereCode As String

      Dim sColumnFilter As String
      Dim sColumnOrder As String
      Dim sColumnJoinCode As String = String.Empty
      Dim iColumnAggregiateType As ScriptDB.AggregiateNumeric

      Dim sPartCode As String
      Dim iPartNumber As Integer
      Dim bIsSummaryColumn As Boolean
      Dim sColumnName As String
      Dim bAddRelation As Boolean

      Dim iBackupType As ScriptDB.ExpressionType

      'Dim drRelations() As DataRow
      'Dim drRelation As System.Data.DataRow

      '     Dim guidTableID As HCMGuid

      'Dim sFilter As String

      Dim LineOfCode As ScriptDB.CodeElement

      IsComplex = False
      LineOfCode.CodeType = ScriptDB.ComponentTypes.Column



      ' there has to be a cleaner way, but for the moment put a dummy objbasecolumn in there so the function does not fail with a blah blah is not set to object error on the .TableID property.
      If Not Component.BaseExpression Is Nothing Then
        objBaseColumn = Component.BaseExpression.AssociatedColumn
      End If

      If objBaseColumn Is Nothing Then
        objBaseColumn = New Things.Column
        objBaseColumn.Table = Me.BaseTable
      End If

      '      objThisColumn = Globals.Things.GetObject(Enums.Type.Table, [Component].TableID).Objects.GetObject(Enums.Type.Column, Component.ColumnID)
      objThisColumn = mcolDependencies.GetObject(Enums.Type.Column, Component.ColumnID)

      '  Debug.Assert(objThisColumn.Name <> "Start_Date")
      'Debug.Assert(objBaseColumn.Table.Name <> "Eye_Tests")

      ' Cannot find 
      If objThisColumn Is Nothing Then
        LineOfCode.Code = ""
        Globals.ErrorLog.Add(ErrorHandler.Section.General, Me.ExpressionType, ErrorHandler.Severity.Error, "SQLCode_AddColumn", "can't find column is dependency stack")
      Else

        ' Is this column referencing the column that this udf is attaching itself to? (i.e. recursion)
        If Component.IsColumnByReference Then
          LineOfCode.Code = String.Format("'{0}-{1}'" _
              , CInt(objThisColumn.Table.ID).ToString.PadLeft(8, "0") _
              , CInt(objThisColumn.ID).ToString.PadLeft(8, "0"))

        ElseIf objThisColumn Is objBaseColumn _
            And Not (Me.ExpressionType = ScriptDB.ExpressionType.ColumnFilter _
            Or Me.ExpressionType = ScriptDB.ExpressionType.Mask _
            Or Me.ExpressionType = ScriptDB.ExpressionType.RecordDescription) Then
          LineOfCode.Code = String.Format("@prm_{0}", objThisColumn.Name)

        ElseIf objThisColumn Is Me.AssociatedColumn _
          And Me.ExpressionType = ScriptDB.ExpressionType.ReferencedColumn Then
          LineOfCode.Code = String.Format("@prm_{0}", objThisColumn.Name)

          ' Does the referenced column have default value on it, then reference the UDF/value of the default rather than the column itself.
        ElseIf (Not objThisColumn.DefaultCalcID = 0 And Me.ExpressionType = ScriptDB.ExpressionType.ColumnDefault) Then
          LineOfCode.Code = String.Format("[dbo].[{0}](@pID)", objThisColumn.Name)

        ElseIf objThisColumn.IsCalculated And objThisColumn.Table Is Me.AssociatedColumn.Table And Not Me.ExpressionType = ScriptDB.ExpressionType.ColumnFilter Then

          If objThisColumn.Calculation Is Nothing Then
            objThisColumn.Calculation = objThisColumn.Table.GetObject(Type.Expression, objThisColumn.CalcID)
          End If

          iBackupType = objThisColumn.Calculation.ExpressionType

          'If objThisColumn.Calculation.ExpressionType = ScriptDB.ExpressionType.ColumnCalculation Then
          '  objThisColumn.Calculation.ExpressionType = ScriptDB.ExpressionType.ReferencedColumn
          'Else
          '  objThisColumn.Calculation.ExpressionType = Component.BaseExpression.ExpressionType
          'End If

          objThisColumn.Calculation.ExpressionType = ScriptDB.ExpressionType.ReferencedColumn
          objThisColumn.Calculation.AssociatedColumn = objThisColumn
          objThisColumn.Calculation.GenerateCode()

          mbRequiresRowNumber = mbRequiresRowNumber Or objThisColumn.Calculation.mbRequiresRowNumber

          If objThisColumn.Calculation.IsComplex Then

            'AddToDependencies(objThisColumn.Calculation.mcolDependencies)

            LineOfCode.Code = objThisColumn.Calculation.UDF.CallingCode
          Else
            AddToDependencies(objThisColumn.Calculation.mcolDependencies)
            LineOfCode.Code = objThisColumn.Calculation.UDF.SelectCode
          End If

          objThisColumn.Calculation.ExpressionType = iBackupType

        Else

          'If is this column on the base table then add directly to the main execute statement,
          ' otherwise add it into child/parent statements array
          If objThisColumn.Table Is objBaseColumn.Table Then

            Select Case Component.BaseExpression.ExpressionType
              Case ScriptDB.ExpressionType.ColumnFilter
                sColumnName = String.Format("[{0}].[{1}]", objThisColumn.Table.Name, objThisColumn.Name)
                'mbAddBaseTable = True

              Case ScriptDB.ExpressionType.Mask
                sColumnName = String.Format("base.[{0}]", objThisColumn.Name)

                ' Needs base table added
                sFromCode = String.Format("FROM [dbo].[{0}] base", objThisColumn.Table.Name)
                If Not FromTables.Contains(sFromCode) Then
                  FromTables.Add(sFromCode)
                End If

                ' Where clause
                sWhereCode = String.Format("base.[ID] = @prm_id")
                If Not maryWhere.Contains(sWhereCode) Then
                  maryWhere.Add(sWhereCode)
                End If

              Case Else
                sColumnName = String.Format("@prm_{0}", objThisColumn.Name)
                'mbAddBaseTable = True

            End Select

            LineOfCode.Code = String.Format("ISNULL({0},{1})", sColumnName, objThisColumn.SafeReturnType)

          Else

            IsComplex = True
            sColumnFilter = String.Empty
            sColumnOrder = String.Empty
            bIsSummaryColumn = False

            ' Is parent or child?
            If Component.IsColumnByReference Then
              objRelation = New Things.Relation
              objRelation.RelationshipType = ScriptDB.RelationshipType.Unknown
            Else
              'objRelation = objBaseColumn.Table.GetRelation(objThisColumn.Table.ID)
              objRelation = Me.BaseTable.GetRelation(objThisColumn.Table.ID)
            End If

            If objRelation.RelationshipType = ScriptDB.RelationshipType.Parent Then
              LineOfCode.Code = String.Format("ISNULL([{0}].[{1}],{2})", objThisColumn.Table.Name, objThisColumn.Name, objThisColumn.SafeReturnType)

              ' Add table join component
              sRelationCode = String.Format("INNER JOIN [dbo].[{0}] ON [{0}].[ID] = [{1}].[ID_{2}]" & vbNewLine _
                , objRelation.Name, objBaseColumn.Table.Name, CInt(objRelation.ParentID))
              If Not Joins.Contains(sRelationCode) Then
                Joins.Add(sRelationCode)
              End If

              ' Needs base table added
              sFromCode = String.Format("FROM [dbo].[{0}]", objBaseColumn.Table.Name)
              If Not FromTables.Contains(sFromCode) Then
                FromTables.Add(sFromCode)
              End If

              ' Where clause
              sWhereCode = String.Format("[dbo].[{0}].ID = @prm_id", objBaseColumn.Table.Name)
              If Not maryWhere.Contains(sWhereCode) Then
                maryWhere.Add(sWhereCode)
              End If

              ' Mark this relation has having to be updated in the parent triggers
              If Not Me.ExpressionType = ScriptDB.ExpressionType.RecordDescription Then
                objRelation = objThisColumn.Table.GetRelation(Me.AssociatedColumn.Table.ID)
                objRelation.DependantOnParent = True
              End If

              ' mbAddBaseTable = True

              ' sSQL = String.Join(vbNewLine, maryDeclarations.ToArray(GetType(String))) & vbNewLine & vbNewLine


            Else

              ' Derive code for any filter on this column in a child table
              If CInt([Component].ColumnFilterID) > 0 Then

                objExpression = New Things.Expression
                ChildCodeCluster = New ScriptDB.LinesOfCode

                objExpression = objThisColumn.Table.Objects.GetObject(Things.Type.Expression, [Component].ColumnFilterID)
                'objExpression.BaseExpression = Me.BaseExpression
                '                objExpression.ExpressionType = ScriptDB.ExpressionType.ColumnFilter
                objExpression.ExpressionType = ScriptDB.ExpressionType.ColumnFilter

                objExpression.AssociatedColumn = objThisColumn

                'objExpression.AssociatedColumn = Me.AssociatedColumn

                'SQLCode_AddCodeLevel(objExpression.Objects, ChildCodeCluster)
                'ChildCodeCluster.CodeLevel = [CodeCluster].CodeLevel + 1
                ' ChildCodeCluster.ReturnType = ScriptDB.ComponentValueTypes.Logic
                ' sColumnFilter = vbNewLine & "                AND " & ChildCodeCluster.Statement

                'sSQL = sSQL & vbNewLine & String.Join(vbNewLine, Joins.ToArray(GetType(String)))
                'objExpression.BaseExpression = Me.BaseExpression
                objExpression.GenerateCode()
                sColumnFilter = vbNewLine & "                AND (" & objExpression.UDF.SelectCode & " = 1)"

                objExpression.Filters.Add(objExpression.UDF.SelectCode)

                maryDeclarations.AddRange(objExpression.maryPrerequisitStatements)


                '     maryPrerequisitStatement()

                '     objExpression.UDF.Prerequisites

                ' Add any pre-requisits
                '    maryPrerequisitStatements.Add()


                ' Add any join statements
                If objExpression.Joins.Count > 0 Then
                  sColumnJoinCode = String.Join(vbNewLine, objExpression.Joins.ToArray())
                End If

              End If

              ' Derive the code for the order on this column in a child table
              If CInt([Component].ColumnOrderID) > 0 Then
                sColumnOrder = SQLCode_AddOrder(objThisColumn.Table, [Component].ColumnOrderID)
              End If


              ' Add calculation for this foreign column to the pre-requisits array
              iPartNumber = maryDeclarations.Count

              iColumnAggregiateType = [Component].ColumnAggregiateType
              Select Case iColumnAggregiateType

                'Case enumAggregiateNumeric.Average
                'Case enumAggregiateNumeric.First
                'Case enumAggregiateNumeric.Last

                Case ScriptDB.AggregiateNumeric.Maximum
                  sPartCode = String.Format("{0}SELECT @part_{1} = MAX([{3}].[{2}])" & vbNewLine _
                              , [CodeCluster].Indentation, iPartNumber _
                              , objThisColumn.Name, objThisColumn.Table.Name)

                Case ScriptDB.AggregiateNumeric.Minimum
                  sPartCode = String.Format("{0}SELECT @part_{1} = MIN([{3}].[{2}])" & vbNewLine _
                              , [CodeCluster].Indentation, iPartNumber _
                              , objThisColumn.Name, objThisColumn.Table.Name)

                  '          Case enumAggregiateNumeric.Specific

                Case ScriptDB.AggregiateNumeric.Total
                  sPartCode = String.Format("{0}SELECT @part_{1} = SUM([{3}].[{2}])" & vbNewLine _
                              , [CodeCluster].Indentation, iPartNumber _
                              , objThisColumn.Name, objThisColumn.Table.Name)
                  bIsSummaryColumn = True

                Case ScriptDB.AggregiateNumeric.Count
                  sPartCode = String.Format("{0}SELECT @part_{1} = COUNT([{3}].[{2}])" & vbNewLine _
                              , [CodeCluster].Indentation, iPartNumber _
                              , objThisColumn.Name, objThisColumn.Table.Name)
                  bIsSummaryColumn = True

                Case Else
                  sPartCode = String.Format("{0}SELECT TOP 1 @part_{1} = [{3}].[{2}]" & vbNewLine _
                              , [CodeCluster].Indentation, iPartNumber _
                              , objThisColumn.Name, objThisColumn.Table.Name)

              End Select

              'AND [_deleteddate] IS NULL{3}" & vbNewLine _
              '              & "{0}WHERE [pid_{2}] = @pID " & vbNewLine _

              ' Add to prereqistits arrays
              If bIsSummaryColumn Then
                maryDeclarations.Add(String.Format("{0}DECLARE @part_{1} numeric(38,8);" _
                    , [CodeCluster].Indentation, iPartNumber))
                sColumnOrder = vbNullString
              Else
                maryDeclarations.Add(String.Format("{0}DECLARE @part_{1} {2};" _
                    , [CodeCluster].Indentation, iPartNumber, objThisColumn.DataTypeSyntax))
              End If


              ' TODO - this needs to reference the relationship to the parent table, how do we deal with get field from db record?!!!!
              If Component.IsColumnByReference Then

                sPartCode = sPartCode & String.Format("{0}FROM [dbo].[{1}]" & vbNewLine _
                    & "{0} " & vbNewLine _
                    & "{0}{2}" & vbNewLine _
                    , [CodeCluster].Indentation _
                    , objThisColumn.Table.Name _
                    , sColumnFilter, sColumnOrder)
              Else
                sPartCode = sPartCode & String.Format("{0}FROM [dbo].[{1}]" & vbNewLine _
                    & "{5}" & vbNewLine _
                    & "{0}WHERE [id_{2}] = @pID_{2} " & vbNewLine _
                    & "{0}{3}" & vbNewLine _
                    & "{0}{4}" & vbNewLine _
                    , [CodeCluster].Indentation _
                    , objThisColumn.Table.Name _
                    , CInt(Me.BaseTable.ID), sColumnFilter, sColumnOrder, sColumnJoinCode)
              End If

              '29/5/11
              ' just changed to  Me.BaseTable from objBaseColumn.Table

              'TODO -  change the relation object to have references to teh poarent and child instead of guids?
              ' Add relation to the dependency stack
              bAddRelation = True
              For Each objDepends As Things.Base In mcolDependencies
                If objDepends.Type = Enums.Type.Relation Then
                  If CType(objDepends, Things.Relation).ParentID = objRelation.ParentID Then
                    bAddRelation = False
                    Exit For
                  End If
                End If
              Next

              If bAddRelation Then
                If Not mcolDependencies.Contains(objRelation) Then
                  mcolDependencies.Add(objRelation)
                End If
              End If

              maryPrerequisitStatements.Add(sPartCode)
              LineOfCode.Code = String.Format("ISNULL(@part_{0},{1})", iPartNumber, objThisColumn.SafeReturnType)

            End If

            ' Add table join component
            'maryJoins.Add(String.Format("INNER JOIN [dbo].[{0}] p{1} ON p{1}.[fk_{2}] = @pID AND p{1}.[_deleteddate] IS NULL" _
            '  , PhysicalName(guidTableID, ObjectPrefix.Table) _
            '  , maryJoins.Count _
            '  , drRelation.Item("parentid").ToString))

            '' Add order
            'If [Component].Item("columnorderid").ToString = ASR.Common.ASRGuid.Empty Then
            '  maryOrders.Add(String.Format("--{0} DESC" _
            '    , PhysicalName([Component].Item("columnorderid").ToString, ObjectPrefix.Table)))
            'End If

            End If

          End If
      End If

        ' Add this column (or reference to it) to the main execute statement
        [CodeCluster].Add(LineOfCode)

        ' Add this column's tableID to the dependency stack
        '  AddTableToDependencies(objThisColumn.Table.ID)
        '   AddColumnToDependencies(objThisColumn)

    End Sub

    Private Sub SQLCode_AddFunction(ByRef [Component] As Things.Component, ByRef [CodeCluster] As ScriptDB.LinesOfCode)

      Dim LineOfCode As ScriptDB.CodeElement

      Dim objCodeLibrary As Things.CodeLibrary
      Dim ChildCodeCluster As ScriptDB.LinesOfCode
      Dim objSetting As Things.Setting
      Dim objIDComponent As Things.Component

      LineOfCode.CodeType = ScriptDB.ComponentTypes.Function
      objCodeLibrary = Globals.Functions.GetObject(Enums.Type.CodeLibrary, Component.FunctionID)
      LineOfCode.Code = objCodeLibrary.Code

      ' Get parameters
      ChildCodeCluster = New ScriptDB.LinesOfCode
      ChildCodeCluster.CodeLevel = [CodeCluster].CodeLevel + 1
      ChildCodeCluster.NestedLevel = CodeCluster.NestedLevel + 1
      ChildCodeCluster.ReturnType = objCodeLibrary.ReturnType

      ' Add module dependancy info for this function
      If objCodeLibrary.HasDependancies Then
        For Each objSetting In objCodeLibrary.Dependancies

          Select Case objSetting.SettingType

            Case SettingType.ModuleSetting

              '     Select Case objSetting.SubType
              'Case Enums.Type.Table

              ' Add it as a relation
              'objTable = Globals.Things.GetObject(Enums.Type.Table, objSetting.Value)
              'objRelation = AssociatedColumn.Table.GetRelation(objTable.ID)
              'AddDependancy(objRelation)

              objIDComponent = New Things.Component
              objIDComponent.SubType = ScriptDB.ComponentTypes.Relation
              objIDComponent.TableID = objSetting.Value
              [Component].Objects.Add(objIDComponent)

              '    Case Enums.Type.Column
              'AddDependancy(Globals.Things.GetObject(Enums.Type.Table, objSetting.Value))

              '  End Select

            Case SettingType.CodeItem
              objIDComponent = New Things.Component
              objIDComponent.SubType = ScriptDB.ComponentTypes.Value
              objIDComponent.ValueString = objSetting.Code
              objIDComponent.ValueType = ScriptDB.ComponentValueTypes.SystemVariable
              [Component].Objects.Add(objIDComponent)

          End Select

        Next
      End If

      'ChildCodeCluster.IsEvaluated = Not objCodeLibrary.BypassValidation
      SQLCode_AddCodeLevel([Component].Objects, ChildCodeCluster)
      LineOfCode.BypassEvaluation = objCodeLibrary.BypassValidation

      'bocth?
      ' If objCodeLibrary.ID = 4 Then
      'LineOfCode.Code = ChildCodeCluster.Statement
      '  Else
      LineOfCode.Code = String.Format(LineOfCode.Code, ChildCodeCluster.ToArray)
      '   End If

      mbRequiresRowNumber = mbRequiresRowNumber Or objCodeLibrary.RowNumberRequired
      mbCalculatePostAudit = mbCalculatePostAudit Or objCodeLibrary.CalculatePostAudit

      ' For functions that return mixed type, make it type safe
      If objCodeLibrary.ReturnType = ScriptDB.ComponentValueTypes.Unknown Then

        Select Case Component.ReturnType
          Case ScriptDB.ComponentValueTypes.Numeric
            LineOfCode.Code = String.Format("convert(float, ({0}))", LineOfCode.Code)

          Case ScriptDB.ComponentValueTypes.Date
            LineOfCode.Code = String.Format("convert(datetime, ({0}))", LineOfCode.Code)

          Case ScriptDB.ComponentValueTypes.Logic
            ' if it doesn't equal "=1 " automatically add it on. (Must tidy this up)
            'If Not Right( LineOfCode.Code,3) = "= 1)" the

            'LineOfCode.Code = String.Format("convert(bit, ({0}))", LineOfCode.Code)
            'CodeCluster.IsEvaluated = Not objCodeLibrary.BypassValidation
            'LineOfCode. = objCodeLibrary.BypassValidation
            'If Not objCodeLibrary.BypassValidation then

          Case ScriptDB.ComponentValueTypes.String
            LineOfCode.Code = String.Format("convert(nvarchar(MAX), ({0}))", LineOfCode.Code)

            'Case ScriptDB.ColumnTypes.Integer
            '  LineOfCode.Code = String.Format("convert(integer, ({0}))", LineOfCode.Code)

        End Select


      End If

      ' Attach the line of code

      [CodeCluster].Add(LineOfCode)

    End Sub

    Private Sub SQLCode_AddParameter(ByRef [Component] As Things.Component, ByRef [CodeCluster] As ScriptDB.LinesOfCode)

      Dim ChildCodeCluster As ScriptDB.LinesOfCode
      Dim LineOfCode As ScriptDB.CodeElement
      Dim objExpression As Things.Expression
      Dim iPartNumber As Integer
      Dim sPartCode As String

      ' Build code for the parameters
      ChildCodeCluster = New ScriptDB.LinesOfCode

      ChildCodeCluster.ReturnType = Component.ReturnType
      ChildCodeCluster.CodeLevel = CodeCluster.CodeLevel + 1
      ChildCodeCluster.NestedLevel = CodeCluster.NestedLevel
      '      ChildCodeCluster.

      ' Nesting is too deep - convert to part number
      If ChildCodeCluster.NestedLevel > 12 Then 'And objCodeLibrary.SplitIntoCase Then

        objExpression = New Things.Expression
        objExpression.ExpressionType = ScriptDB.ExpressionType.Mask
        objExpression.BaseTable = Me.BaseTable
        objExpression.AssociatedColumn = Me.AssociatedColumn
        objExpression.ReturnType = Component.ReturnType
        objExpression.Objects = Component.Objects
        objExpression.GenerateCode()

        iPartNumber = maryDeclarations.Count
        '        maryDeclarations.Add(String.Format("{0}DECLARE @part_{1} {2};", [CodeCluster].Indentation, iPartNumber, ScriptDB.GetSQLColumnDatatype(Component.ReturnType)))
        maryDeclarations.Add(String.Format("{0}DECLARE @part_{1} {2};", [CodeCluster].Indentation, iPartNumber, objExpression.DataTypeSyntax))

        sPartCode = String.Format("{0}SELECT @part_{1} = {2}" & vbNewLine & _
            "{0}{3}" & vbNewLine & _
            "{0}{4}" & vbNewLine & _
            "{0}{5}" & vbNewLine _
            , [CodeCluster].Indentation, iPartNumber _
            , objExpression.UDF.SelectCode, objExpression.UDF.FromCode, objExpression.UDF.JoinCode, objExpression.UDF.WhereCode)
        maryPrerequisitStatements.Add(sPartCode)

        LineOfCode.Code = String.Format("@part_{0}", iPartNumber)
      Else
        SQLCode_AddCodeLevel([Component].Objects, ChildCodeCluster)       

        ' Debug.Assert(Component.IsEvaluated = False)

        ChildCodeCluster.IsEvaluated = Component.IsEvaluated
        ' Debug.Print(Component.SubType)

        ' Debug.Print(Component.IsEvaluated)

        ' [Component].

        'If ChildCodeCluster.CodeLevel = 1 Then
        '  ChildCodeCluster.IsEvaluated = True
        'End If
        'ChildCodeCluster.IsEvaluated = (Component.ReturnType = ScriptDB.ComponentValueTypes.Logic)
        'If Component.ReturnType = ScriptDB.ComponentValueTypes.Logic Then
        '  Debug.Print("hhhd")
        'End If
        '        LineOfCode.Code = String.Format("(CASE WHEN {0} THEN 1 ELSE 0 END)", ChildCodeCluster.Statement)
        '       Else

        sPartCode = ChildCodeCluster.Statement
        '  If ChildCodeCluster.IsCodeFlow Then
        '     LineOfCode.Code = String.Format("(CASE WHEN ({0}=1) THEN 1 ELSE 0 END)", sPartCode)
        '    Else
        LineOfCode.Code = String.Format("{0}", sPartCode)
        '     End If


        'If Not ChildCodeCluster.IsComparison And Component.ReturnType = ScriptDB.ComponentValueTypes.Logic Then
        '  LineOfCode.Code = String.Format("(CASE WHEN ({0}=1) THEN 1 ELSE 0 END)", sPartCode)
        'Else
        '  LineOfCode.Code = String.Format("{0}", sPartCode)
        'End If

        '        End If
        'Debug.Print()

      End If

      [CodeCluster].Add(LineOfCode)

    End Sub

    Private Sub SQLCode_AddOperator(ByVal objComponent As Things.Component, ByRef [CodeCluster] As ScriptDB.LinesOfCode)

      Dim LineOfCode As ScriptDB.CodeElement
      Dim objCodeLibrary As Things.CodeLibrary

      LineOfCode.CodeType = ScriptDB.ComponentTypes.Operator

      ' Get the bits and bobs for this operator
      objCodeLibrary = Globals.Operators.GetObject(Enums.Type.CodeLibrary, objComponent.OperatorID)

      ' Handle 'OR' statements. Force the component builder to wrap the logic clusters into a case statement
      If objCodeLibrary.SplitIntoCase Then
        [CodeCluster].SplitIntoCase()
      Else

        ' We're starting a new section of logic components.
        'If objCodeLibrary.OperatorType = ScriptDB.OperatorSubType.Logic Then
        '  [CodeCluster].StartNewLogicCluster()
        'End If

        LineOfCode.Code = String.Format(" {0} ", objCodeLibrary.Code)
        LineOfCode.OperatorType = objCodeLibrary.OperatorType
        [CodeCluster].Add(LineOfCode)

        If objCodeLibrary.AppendWildcard Then
          [CodeCluster].AppendWildcard()
        End If

        If objCodeLibrary.AfterCode.Length > 0 Then
          LineOfCode.Code = String.Format("{0}", objCodeLibrary.AfterCode)
          [CodeCluster].AddToEnd(LineOfCode)
        End If

      End If

    End Sub

    Public Function SQLCode_AddOrder(ByRef objTable As Things.Table, ByVal [OrderID] As HCMGuid) As String

      Dim objOrderItems As Things.Collection
      Dim objOrderItem As Things.TableOrderItem
      Dim sReturn As String = String.Empty
      Dim aryOrderBy As New ArrayList

      objOrderItems = objTable.GetObject(Enums.Type.TableOrder, OrderID).Objects

      For Each objOrderItem In objOrderItems
        If objOrderItem.ColumnType = "O" Then
          Select Case objOrderItem.Ascending
            Case Enums.Order.Ascending
              aryOrderBy.Add(String.Format("[{0}]{1}", objOrderItem.Column.Name, " ASC"))
            Case Else
              aryOrderBy.Add(String.Format("[{0}]{1}", objOrderItem.Column.Name, " DESC"))
          End Select
        End If
      Next

      If aryOrderBy.Count > 0 Then
        sReturn = "ORDER BY " & String.Join(", ", aryOrderBy.ToArray())
      End If

      Return sReturn

    End Function

#End Region

    ' Declaration syntax for a column type
    Public ReadOnly Property DataTypeSyntax As String
      Get

        Dim sSQLType As String = String.Empty

        Select Case CInt(Me.ReturnType)
          Case ScriptDB.ColumnTypes.Text
            If Me.Size > 8000 Then
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

    Private Sub AddToDependencies(ByRef Dependencies As Things.Collection)

      Dim objDependency As Things.Base
      Dim objColumn As Things.Column
      Dim objRelation As Things.Relation

      For Each objDependency In Dependencies

        If objDependency.Type = Enums.Type.Column Then
          objColumn = CType(objDependency, Things.Column)
          If Not mcolDependencies.Contains(objColumn) Then
            If objColumn.Table Is Me.BaseTable Then
              mcolDependencies.Add(objDependency)
            End If
          End If
        End If

        If objDependency.Type = Enums.Type.Relation Then
          objRelation = CType(objDependency, Things.Relation)
          If Not mcolDependencies.Contains(objRelation) Then
            mcolDependencies.Add(objDependency)
          End If
        End If

      Next

    End Sub

  End Class
End Namespace

