Option Strict On
Option Explicit On

Imports System.Runtime.InteropServices
Imports SystemFramework.Enums
Imports SystemFramework.Enums.Errors
Imports SystemFramework.Structures

<ClassInterface(ClassInterfaceType.None), ComVisible(True), Serializable()>
Public Class Expression
  Inherits Component

  Public Property Size As Integer
  Public Property Decimals As Integer
  Public Property BaseTable As Table

  Public Udf As GeneratedUdf
  Public Property ExpressionType As ExpressionType

  Public Property Dependencies As Collections.ExpressionDependencies

  Public Property Declarations As New ArrayList
  Public Property PreStatements As New ArrayList
  Private Property Joins As ArrayList
  Private Property FromTables As ArrayList
  Private Property Wheres As ArrayList

  Private _linesOfCode As ScriptDB.LinesOfCode
  Private _calculatePostAudit As Boolean

  Private Property CaseCount As Integer
  Private Property RequiresRecordId As Boolean
  Private Property RequiresOvernight As Boolean
  Private Property ReferencesParent As Boolean
  Private Property ReferencesChild As Boolean

  Public Property IsComplex As Boolean
  Public Property IsValid As Boolean = True

  Public Sub New()
    Dependencies = New Collections.ExpressionDependencies
  End Sub

  Public ReadOnly Property CalculatePostAudit As Boolean
    Get
      Return _calculatePostAudit
    End Get
  End Property

#Region "Generate code"

  Private Sub BuildDependancies(ByVal expression As Component)

    Dim bAddThis As Boolean

    For Each component As Component In expression.Components

      bAddThis = False

      Select Case component.SubType
        Case ComponentTypes.Column
          bAddThis = True

        Case ComponentTypes.Function, ComponentTypes.Expression, ComponentTypes.Calculation
          BuildDependancies(component)

        Case ComponentTypes.ConvertedCalculatedColumn
          BuildDependancies(component)
          bAddThis = True

      End Select

      If bAddThis Then

        Dim table As Table = Tables.GetById(component.TableId)
        Dim column As Column = table.Columns.GetById(component.ColumnId)

        Dependencies.Add(column)
        Dependencies.Add(column.Table)

      End If

    Next

  End Sub

  Public Sub GenerateCodeForColumn()

    ' Build the dependencies collection
    Dependencies.Clear()
    BuildDependancies(Me)

    IsComplex = False

    GenerateCode()
  End Sub

  Public Overridable Sub GenerateCode()

    Dim sOptions As String = String.Empty
    Dim aryDependsOn As New ArrayList
    Dim aryComments As New ArrayList
    Dim aryParameters1 As New ArrayList
    Dim aryParameters2 As New ArrayList
    Dim aryParameters3 As New ArrayList

    ' Initialise code object
    _linesOfCode = New ScriptDB.LinesOfCode
    _linesOfCode.Clear()
    _linesOfCode.ReturnType = ReturnType
    _linesOfCode.CodeLevel = If(ExpressionType = ExpressionType.ColumnFilter, 2, 1)

    Joins = New ArrayList
    FromTables = New ArrayList
    Wheres = New ArrayList

    Declarations.Clear()
    PreStatements.Clear()
    Joins.Clear()
    Wheres.Clear()

    ' If calculate only when empty add itself to the dependency stack
    If AssociatedColumn.CalculateIfEmpty Then
      Dependencies.Add(AssociatedColumn)
    End If

    aryParameters1.Clear()
    aryParameters2.Clear()
    aryParameters3.Clear()

    ' Build the execution code
    SQLCode_AddCodeLevel(Components, _linesOfCode)

    ' Add return declaration
    Declarations.Add(String.Format("@Result {0}", DataTypeSyntax))

    ' Add the ID for the record if required
    If RequiresRecordId Or ExpressionType = ExpressionType.ColumnDefault Then
      aryParameters1.Add("@prm_ID integer")
      aryParameters2.Add("base.ID")
      aryParameters3.Add("@prm_ID")
    End If

    ' Some function require the row number of the record as a parameter
    If RequiresOvernight Then
      aryParameters1.Add("@isovernight bit")
      aryParameters2.Add("@isovernight")
      aryParameters3.Add("@isovernight")
    End If

    ' Add other dependancies
    For Each column In Dependencies.Columns
      If column.Table Is BaseTable Then
        aryParameters1.Add(String.Format("@prm_{0} {1}", column.Name, column.DataTypeSyntax))
        aryParameters2.Add(String.Format("base.[{0}]", column.Name))
        aryParameters3.Add(String.Format("@prm_{0}", column.Name))
      End If
    Next

    ' BEGIN....................
    ' clump all the dependency stuff together?

    ' Add child columns
    For Each Dependency In Dependencies.ChildRowDetails
      Declarations.Add(Dependency.Declaration)
      PreStatements.Add(Dependency.Code)
    Next

    For Each objStatement As GeneratedUdf In Dependencies.Statements
      Declarations.Add(objStatement.Declaration)
      PreStatements.Add(objStatement.Code)
    Next

    '.................... END


    For Each relation In Dependencies.Relations

      If Not aryParameters1.Contains(String.Format("@prm_ID_{0} integer", relation.ParentId)) Then
        aryParameters1.Add(String.Format("@prm_ID_{0} integer", relation.ParentId))

        If relation.RelationshipType = RelationshipType.Parent Then
          aryParameters2.Add(String.Format("base.[ID_{0}]", relation.ParentId))
          aryParameters3.Add(String.Format("@prm_ID_{0}", relation.ParentId))
          aryComments.Add(String.Format("Relation :{0}", relation.Name))
        Else
          aryParameters2.Add("base.[ID]")
          aryParameters3.Add(String.Format("@prm_ID"))
          aryComments.Add(String.Format("Relation : {0}", relation.Name))
        End If

      End If
    Next

    '' Add relationship code
    'For Each table In Dependencies.Tables
    '  'aryComments.Add(String.Format("Table : {0}", table.Name))
    '  '        aryDependsOn.Add(String.Format("{0}", table.ID))

    '  'If Relation.RelationshipType = RelationshipType.Parent Then
    '  '  aryParameters2.Add(String.Format("base.[ID_{0}]", Relation.ParentID))
    '  '  aryParameters3.Add(String.Format("@prm_ID_{0}", Relation.ParentID))
    '  '  aryComments.Add(String.Format("Relation :{0}", Relation.Name))
    '  'Else
    '  aryParameters2.Add("base.[ID]")
    '  aryParameters3.Add(String.Format("@prm_ID"))
    '  'aryComments.Add(String.Format("Relation : {0}", Relation.Name))
    '  'End If

    'Next

    ' Calling statement
    With Udf

      If Not IsComplex Then
        .InlineCode = ResultWrapper(_linesOfCode.Statement)
        .InlineCode = .InlineCode.Replace("@prm_", "base.")
        .InlineCode = ScriptDB.Beautify.MakeSingleLine(.InlineCode)
      End If

      Description = ScriptDB.Beautify.MakeSingleLine(Description)

      .BoilerPlate = String.Format("-----------------------------------------------------------------" & vbNewLine & _
            "-- Generated by the Advanced System Framework" & vbNewLine & _
            "-- Column      : {1}.{0}" & vbNewLine & _
            "-- Expression  : {2}" & vbNewLine & _
            "-- Description : {7}" & vbNewLine & _
            "-- Depends on  : {3}" & vbNewLine & _
            "-- Date        : {4}" & vbNewLine & _
            "-- Complexity  : ({5}) {6}" & vbNewLine & _
            "----------------------------------------------------------------" & vbNewLine _
            , AssociatedColumn.Name, AssociatedColumn.Table.Name, BaseExpression.Name _
            , String.Join(", ", aryDependsOn.ToArray()), Now().ToString _
            , Tuning.Rating, Tuning.ExpressionComplexity, Description)
      .Declarations = If(Declarations.Count > 0, "DECLARE " & String.Join("," & vbNewLine & vbTab & vbTab & vbTab, Declarations.ToArray()) & ";" & vbNewLine, "")
      .Prerequisites = If(PreStatements.Count > 0, String.Join(vbNewLine, PreStatements.ToArray()) & vbNewLine & vbNewLine, "")
      .JoinCode = If(Joins.Count > 0, String.Format("{0}", String.Join(vbNewLine, Joins.ToArray)) & vbNewLine, "")
      .FromCode = If(FromTables.Count > 0, String.Format("{0}", String.Join(",", FromTables.ToArray)) & vbNewLine, "")
      .WhereCode = If(Wheres.Count > 0, String.Format("WHERE {0}", String.Join(" AND ", Wheres.ToArray)) & vbNewLine, "")

      ' Code beautify
      .Prerequisites = ScriptDB.Beautify.CleanWhitespace(.Prerequisites)

      Select Case ExpressionType

        Case ExpressionType.ColumnDefault
          .Name = String.Format("[{0}].[{1}{2}.{3}]", SchemaName, ScriptDB.Consts.DefaultValueUdf, AssociatedColumn.Table.Name, AssociatedColumn.Name)
          .SelectCode = ScriptDB.Beautify.CleanWhitespace(_linesOfCode.Statement)
          .CallingCode = String.Format("{0}({1})", .Name, String.Join(",", aryParameters2.ToArray))
          .Code = String.Format("{10}CREATE FUNCTION {0}({1})" & vbNewLine & _
                         "RETURNS {2}" & vbNewLine & _
                         "{3}" & vbNewLine & _
                         "AS" & vbNewLine & "BEGIN" & vbNewLine & _
                         "    {4}" & vbNewLine & vbNewLine & _
                         "    {5}" & vbNewLine & vbNewLine & _
                         "    -- Execute calculation code" & vbNewLine & _
                         "    SELECT @Result = {6}" & vbNewLine & _
                         "                 {7}{8}{9}" & vbNewLine & _
                         "    RETURN {11};" & vbNewLine & _
                         "END" _
                        , .Name, String.Join(", ", aryParameters1.ToArray()) _
                        , AssociatedColumn.DataTypeSyntax, sOptions, .Declarations, .Prerequisites, .SelectCode.Trim, .FromCode, .JoinCode, .WhereCode _
                        , .BoilerPlate, ResultWrapper("@Result"))


          ' Wrapper for calculations with associated columns
        Case ExpressionType.ColumnCalculation
          .Name = String.Format("[{0}].[{1}{2}.{3}]", SchemaName, ScriptDB.Consts.CalculationUdf, AssociatedColumn.Table.Name, AssociatedColumn.Name)
          .SelectCode = _linesOfCode.Statement
          .CallingCode = String.Format("{0}({1})", .Name, String.Join(",", aryParameters2.ToArray))
          .Code = String.Format("{10}CREATE FUNCTION {0}({1})" & vbNewLine & _
                         "RETURNS {2}" & vbNewLine & _
                         "{3}" & vbNewLine & _
                         "AS" & vbNewLine & "BEGIN" & vbNewLine & _
                         "    {4}" & vbNewLine & vbNewLine & _
                         "    {5}" & vbNewLine & vbNewLine & _
                         "    -- Execute calculation code" & vbNewLine & _
                         "    SELECT @Result = {6}" & vbNewLine & _
                         "                 {7}{8}{9}" & vbNewLine & _
                         "    RETURN {11};" & vbNewLine & _
                         "END" _
                        , .Name, String.Join(", ", aryParameters1.ToArray()) _
                        , AssociatedColumn.DataTypeSyntax, sOptions, .Declarations, .Prerequisites, .SelectCode.Trim, .FromCode, .JoinCode, .WhereCode _
                        , .BoilerPlate, ResultWrapper("@Result"))

          ' Wrapper for when this function is used as a filter in an expression
        Case ExpressionType.ColumnFilter
          .Name = String.Format("[{0}].[{1}{2}.{3}]", SchemaName, ScriptDB.Consts.CalculationUdf, AssociatedColumn.Table.Name, AssociatedColumn.Name)
          .CallingCode = String.Format("{0}({1})", .Name, String.Join(",", aryParameters2.ToArray))
          .SelectCode = _linesOfCode.Statement

          ' Wrapper for when expression is used as a filter in a view
        Case ExpressionType.Mask
          .Name = String.Format("[{0}].[{1}{2}]", SchemaName, ScriptDB.Consts.MaskUdf, BaseExpression.Id)
          .CallingCode = String.Format("{0}({1})", .Name, String.Join(",", aryParameters1.ToArray))
          .SelectCode = _linesOfCode.Statement

          .Code = String.Format("CREATE FUNCTION {0}(@prm_ID integer)" & vbNewLine & _
                    "RETURNS bit" & vbNewLine & _
                    "--WITH SCHEMABINDING" & vbNewLine & _
                    "AS" & vbNewLine & "BEGIN" & vbNewLine & vbNewLine & _
                    "{1}" & vbNewLine & vbNewLine & _
                    "{2}" & vbNewLine & vbNewLine & _
                    "    -- Execute calculation code" & vbNewLine & _
                    "    SELECT @Result = {3}" & vbNewLine & _
                    "                 {4}" & vbNewLine & _
                    "                 {5}" & vbNewLine & _
                    "                 {6}" & vbNewLine & _
                    "    RETURN ISNULL(@Result, 0);" & vbNewLine & _
                    "END" _
                    , .Name, .Declarations, .Prerequisites, .SelectCode.Trim, .FromCode, .JoinCode, .WhereCode)

        Case ExpressionType.ReferencedColumn
          .Name = String.Format("[{0}].[{1}{2}.{3}]", SchemaName, ScriptDB.Consts.CalculationUdf, AssociatedColumn.Table.Name, AssociatedColumn.Name)
          .CallingCode = String.Format("{0}({1})", .Name, String.Join(",", aryParameters3.ToArray))
          .SelectCode = _linesOfCode.Statement

        Case ExpressionType.RecordDescription
          .Name = String.Format("[{0}].[{1}{2}]", SchemaName, ScriptDB.Consts.RecordDescriptionUdf, BaseTable.Name)
          .CallingCode = String.Format("{0}({1})", .Name, String.Join(",", aryParameters2.ToArray))
          .SelectCode = _linesOfCode.Statement

          .Code = String.Format("CREATE FUNCTION {0}({1})" & vbNewLine & _
                         "RETURNS nvarchar(MAX)" & vbNewLine & _
                         "{2}" & vbNewLine & _
                         "AS" & vbNewLine & "BEGIN" & vbNewLine & vbNewLine & _
                         "{3}" & vbNewLine & vbNewLine & _
                         "{4}" & vbNewLine & vbNewLine & _
                         "    -- Execute calculation code" & vbNewLine & _
                         "    SELECT @Result = {5}" & vbNewLine & _
                         "                 {6}" & vbNewLine & _
                         "                 {7}" & vbNewLine & _
                         "                 {8}" & vbNewLine & _
                         "    RETURN ISNULL(@Result, '');" & vbNewLine & _
                         "END" _
                        , .Name, String.Join(", ", aryParameters1.ToArray()) _
                        , sOptions, .Declarations, .Prerequisites, .SelectCode.Trim, .FromCode, .JoinCode, .WhereCode)

          ' Should never be called, but just in case...
        Case Else
          .SelectCode = _linesOfCode.Statement

      End Select

    End With

  End Sub

  Private Sub SQLCode_AddCodeLevel(ByVal [addComponents] As IEnumerable(Of Component), ByVal codeCluster As ScriptDB.LinesOfCode)

    Dim objComponent As Component

    Dim lineOfCode As CodeElement
    Dim objCalculation As Expression

    For Each objComponent In [addComponents]

      Select Case objComponent.SubType

        ' A table relationship
        Case ComponentTypes.Relation
          SQLCode_AddRelation([codeCluster], objComponent)

          ' Column component
        Case ComponentTypes.Column
          SQLCode_AddColumn([codeCluster], objComponent)

          ' Operator component
        Case ComponentTypes.Operator
          SQLCode_AddOperator(objComponent, [codeCluster])

          ' Value component
        Case ComponentTypes.Value, ComponentTypes.TableValue
          lineOfCode.CodeType = ComponentTypes.Value

          Select Case objComponent.ValueType
            Case ComponentValueTypes.Numeric
              lineOfCode.Code = String.Format("{0}", objComponent.ValueNumeric)

            Case ComponentValueTypes.String
              lineOfCode.Code = String.Format("'{0}'", objComponent.ValueString.Replace("'", "''"))

            Case ComponentValueTypes.Date
              lineOfCode.Code = String.Format("'{0}'", objComponent.ValueDate.ToString("yyyy-MM-dd"))

            Case ComponentValueTypes.SystemVariable
              lineOfCode.Code = String.Format("{0}", objComponent.ValueString)

            Case Else
              lineOfCode.Code = String.Format("{0}", If(objComponent.ValueLogic, 1, 0))

          End Select

          [codeCluster].Add(lineOfCode)


          ' Function component
        Case ComponentTypes.Function
          SQLCode_AddFunction(objComponent, [codeCluster])

          ' Calculated columns are sucked into this expressions
        Case ComponentTypes.ConvertedCalculatedColumn
          SQLCode_AddParameter(objComponent, [codeCluster], True)

          ' An expression or a parameter
        Case ComponentTypes.Expression
          SQLCode_AddParameter(objComponent, [codeCluster], False)

          ' Calculation 
        Case ComponentTypes.Calculation

          If Not objComponent.BaseExpression.BaseTable.Expressions.GetById(objComponent.CalculationId) Is Nothing Then

            objCalculation = objComponent.BaseExpression.BaseTable.Expressions.GetById(objComponent.CalculationId).Clone

            'objCalculation.StartOfPartNumbers = 0
            objCalculation.BaseExpression = objComponent.BaseExpression
            objComponent.Components = objCalculation.CloneComponents
            objComponent.ReturnType = objCalculation.ReturnType
            SQLCode_AddParameter(objComponent, [codeCluster], False)

          Else
            ErrorLog.Add(Section.General, AssociatedColumn.Name, Severity.Error, _
                "SQLCode_AddCodeLevel", AssociatedColumn.Table.Name & "." & AssociatedColumn.Name & " -- Missing calculation")
            IsValid = False
            IsComplex = True
          End If

          IsComplex = True

        Case ComponentTypes.Filter

          If Not objComponent.BaseExpression.BaseTable.Expressions.GetById(objComponent.FilterId) Is Nothing Then

            objCalculation = objComponent.BaseExpression.BaseTable.Expressions.GetById(objComponent.FilterId).Clone

            'objCalculation.StartOfPartNumbers = 0
            objCalculation.BaseExpression = objComponent.BaseExpression
            objComponent.Components = objCalculation.CloneComponents
            objComponent.ReturnType = ComponentValueTypes.Logic
            SQLCode_AddParameter(objComponent, [codeCluster], False)

          Else
            ErrorLog.Add(Section.General, AssociatedColumn.Name, Severity.Error, _
                "SQLCode_AddCodeLevel", AssociatedColumn.Table.Name & "." & AssociatedColumn.Name & " -- Missing filter")
            IsValid = False
            IsComplex = True
          End If

          IsComplex = True

      End Select

    Next

  End Sub

  Private Sub SQLCode_AddRelation(ByVal codeCluster As ScriptDB.LinesOfCode, ByVal component As Component)

    Dim objTable As Table
    Dim objRelation As Relation
    Dim lineOfCode As CodeElement

    lineOfCode.CodeType = ComponentTypes.Relation

    objTable = Tables.GetById([component].TableId)
    objRelation = AssociatedColumn.Table.GetRelation(objTable.Id)

    Dependencies.Add(objRelation)
    Dependencies.Add(objTable)

    If ExpressionType = ExpressionType.Mask Then
      lineOfCode.Code = "@prm_ID"
    Else
      lineOfCode.Code = String.Format("@prm_ID_{0}", [component].TableId)
    End If

    [codeCluster].Add(lineOfCode)

  End Sub

  Private Sub SQLCode_AddColumn(ByVal codeCluster As ScriptDB.LinesOfCode, ByVal component As Component)

    Dim objThisColumn As Column

    Dim objRelation As Relation
    Dim sRelationCode As String
    Dim sFromCode As String
    Dim sWhereCode As String

    Dim iPartNumber As Integer
    Dim sColumnName As String

    Dim lineOfCode As CodeElement

    lineOfCode.CodeType = ComponentTypes.Column

    objThisColumn = Dependencies.Columns.FirstOrDefault(Function(o) o.Id = component.ColumnId)
    objThisColumn.Tuning.Usage += 1

    ' Is this column referencing the column that this udf is attaching itself to? (i.e. recursion)
    If component.IsColumnByReference Then
      lineOfCode.Code = String.Format("'{0}-{1}'" _
          , objThisColumn.Table.Id.ToString.PadLeft(8, "0"c) _
          , objThisColumn.Id.ToString.PadLeft(8, "0"c))

    ElseIf objThisColumn Is AssociatedColumn _
        And Not (ExpressionType = ExpressionType.ColumnFilter _
        Or ExpressionType = ExpressionType.TriggeredUpdate _
        Or ExpressionType = ExpressionType.Mask _
        Or ExpressionType = ExpressionType.RecordDescription) Then

      If objThisColumn.SafeReturnType = "NULL" Then
        lineOfCode.Code = "@prm_" & objThisColumn.Name
      Else
        lineOfCode.Code = String.Format("ISNULL(@prm_{0},{1})", objThisColumn.Name, objThisColumn.SafeReturnType)
      End If

    ElseIf objThisColumn Is AssociatedColumn _
        And ExpressionType = ExpressionType.ReferencedColumn Then
      lineOfCode.Code = String.Format("@prm_{0}", objThisColumn.Name)

      ' Does the referenced column have default value on it, then reference the UDF/value of the default rather than the column itself.
    ElseIf (Not objThisColumn.DefaultCalculation Is Nothing _
        And ExpressionType = ExpressionType.ColumnDefault _
        And objThisColumn.Table Is AssociatedColumn.Table) Then
      lineOfCode.Code = String.Format("{0}(@prm_ID)", objThisColumn.DefaultCalculation.Udf.Name)

    Else

      'If is this column on the base table then add directly to the main execute statement,
      ' otherwise add it into child/parent statements array
      If objThisColumn.Table Is AssociatedColumn.Table Then


        Select Case component.BaseExpression.ExpressionType
          Case ExpressionType.ColumnFilter, ExpressionType.Mask
            sColumnName = String.Format("base.[{0}]", objThisColumn.Name)

            ' Needs base table added
            sFromCode = String.Format("FROM [dbo].[{0}] base", objThisColumn.Table.Name)
            If Not FromTables.Contains(sFromCode) Then
              FromTables.Add(sFromCode)
            End If

            'HRPRO-2749 PG
            ' Where clause
            sWhereCode = String.Format("base.[ID] = @prm_ID")
            If Not Wheres.Contains(sWhereCode) Then
              Wheres.Add(sWhereCode)
            End If

            IsComplex = True

            Dependencies.Add(objThisColumn)


          Case Else
            sColumnName = String.Format("@prm_{0}", objThisColumn.Name)

        End Select

        If objThisColumn.SafeReturnType = "NULL" Then
          lineOfCode.Code = sColumnName
        Else
          lineOfCode.Code = String.Format("ISNULL({0},{1})", sColumnName, objThisColumn.SafeReturnType)
        End If
      Else

        RequiresRecordId = True
        IsComplex = True

        objRelation = BaseTable.GetRelation(objThisColumn.Table.Id)

        If objRelation.RelationshipType = RelationshipType.Parent Then

          If ExpressionType = ExpressionType.ColumnCalculation Then
            AssociatedColumn.Table.DependsOnChildColumns.AddIfNew(objThisColumn)
          End If

          ' Is the column calculated (should only apply when used getting child column values that also have a filter applied.
          If objThisColumn.IsCalculated And Not objThisColumn.Calculation Is Nothing And ExpressionType = ExpressionType.ColumnFilter Then
            If objThisColumn.Calculation.IsComplex Then
              lineOfCode.Code = objThisColumn.Calculation.Udf.CallingCode.Replace("base.ID", String.Format("@prm_ID_{0}", objThisColumn.Table.Id))
            Else
              lineOfCode.Code = objThisColumn.Calculation.Udf.InlineCode
            End If
            lineOfCode.Code = lineOfCode.Code.Replace("base.", String.Format("{0}.", objThisColumn.Table.Name))
          Else
            If objThisColumn.SafeReturnType = "NULL" Then
              lineOfCode.Code = String.Format("[{0}].[{1}]", objThisColumn.Table.Name, objThisColumn.Name)
            Else
              lineOfCode.Code = String.Format("ISNULL([{0}].[{1}],{2})", objThisColumn.Table.Name, objThisColumn.Name, objThisColumn.SafeReturnType)
            End If
          End If

          ' Add table join component
          sRelationCode = String.Format("LEFT JOIN [dbo].[{0}] ON [{0}].[ID] = base.[ID_{1}]", objRelation.Name, objRelation.ParentId)
          If Not Joins.Contains(sRelationCode) Then
            Joins.Add(sRelationCode)
          End If

          ' Needs base table added
          sFromCode = String.Format("FROM [dbo].[{0}] base", AssociatedColumn.Table.Name)
          If Not FromTables.Contains(sFromCode) Then
            FromTables.Add(sFromCode)
          End If

          ' Where clause
          If ExpressionType = ExpressionType.ColumnFilter And IsComplex Then
            sWhereCode = String.Format("[{0}].[ID] = @prm_ID_{1}", objRelation.Name, objRelation.ParentId)
          Else
            sWhereCode = "base.[ID] = @prm_ID"
          End If

          If Not Wheres.Contains(sWhereCode) Then
            Wheres.Add(sWhereCode)
          End If

          ReferencesParent = True

        Else

          ' Add to dependency stack
          objThisColumn.Table.DependsOnParentColumns.AddIfNew(AssociatedColumn)

          [component].ChildRowDetails.BaseTable = BaseTable
          [component].ChildRowDetails.Order = objThisColumn.Table.TableOrders.GetById([component].ChildRowDetails.OrderId)
          [component].ChildRowDetails.Filter = objThisColumn.Table.Expressions.GetById([component].ChildRowDetails.FilterId)
          [component].ChildRowDetails.Relation = objRelation
          [component].ChildRowDetails.Column = objThisColumn
          iPartNumber = Dependencies.Add([component].ChildRowDetails)

          ' Any columns used in child filters should be added to the udf chain 
          If Not [component].ChildRowDetails.Filter Is Nothing Then
            For Each objColumn In [component].ChildRowDetails.Filter.Dependencies.Columns
              Dependencies.Add(objColumn)
            Next
          End If


          lineOfCode.Code = String.Format("@child_{0}", iPartNumber)
          ReferencesChild = True

        End If
      End If
    End If

    ' Add this column (or reference to it) to the main execute statement
    [codeCluster].Add(lineOfCode)

  End Sub

  Private Sub SQLCode_AddFunction(ByVal component As Component, ByVal codeCluster As ScriptDB.LinesOfCode)

    Dim lineOfCode As CodeElement
    Dim extraCode As CodeElement

    Dim objCodeLibrary As CodeLibrary
    Dim childCodeCluster As ScriptDB.LinesOfCode
    Dim whereCodeCluster As ScriptDB.LinesOfCode
    Dim objSetting As Setting
    Dim objIdComponent As Component
    Dim objTriggeredUpdate As ScriptDB.TriggeredUpdate
    Dim sWhereClause As String = ""
    Dim bAddDefaultDataType As Boolean
    Dim bAddExpressionType As Boolean = False

    lineOfCode.CodeType = ComponentTypes.Function
    objCodeLibrary = Functions.GetById(component.FunctionId)
    lineOfCode.Code = objCodeLibrary.Code
    CaseCount += objCodeLibrary.CaseCount

    ' Get parameters
    childCodeCluster = New ScriptDB.LinesOfCode
    childCodeCluster.CodeLevel = codeCluster.CodeLevel + 1
    childCodeCluster.ReturnType = objCodeLibrary.ReturnType

    ' Add module dependancy info for this function
    If objCodeLibrary.Dependancies.Count > 0 Then
      For Each objSetting In objCodeLibrary.Dependancies

        Select Case objSetting.SettingType

          Case SettingType.ModuleSetting
            objIdComponent = New Component
            objIdComponent.SubType = ComponentTypes.Relation
            objIdComponent.TableId = CInt(objSetting.Value)
            component.Components.Add(objIdComponent)
            IsComplex = True

          Case SettingType.CodeItem
            objIdComponent = New Component
            objIdComponent.SubType = ComponentTypes.Value
            objIdComponent.ValueString = objSetting.Code
            objIdComponent.ValueType = ComponentValueTypes.SystemVariable
            component.Components.Add(objIdComponent)

          Case SettingType.UpdateParameter
            sWhereClause = objSetting.Code

          Case SettingType.DefaultDataType
            bAddDefaultDataType = True

          Case SettingType.ExpressionType
            bAddExpressionType = True

        End Select

      Next
    End If

    ' Does this component need adding to the 'Get Field From Database' stack?
    If objCodeLibrary.IsGetFieldFromDb Then
      GetFieldsFromDb.Add(component)
    End If

    ' Is this expression reliant on the bank holiday table (I'm sure this can be tidyied up)
    If objCodeLibrary.DependsOnBankHoliday And ExpressionType = ExpressionType.ColumnCalculation Then
      objTriggeredUpdate = New ScriptDB.TriggeredUpdate
      objTriggeredUpdate.Column = AssociatedColumn
      objTriggeredUpdate.Id = AssociatedColumn.Id

      ' Get parameters
      whereCodeCluster = New ScriptDB.LinesOfCode
      SQLCode_AddCodeLevel(component.Components, whereCodeCluster)
      objTriggeredUpdate.Where = String.Format(sWhereClause, whereCodeCluster.ToArray)

      If Not ReferencesChild And Not objTriggeredUpdate.Where.Contains("@part_") Then
        objTriggeredUpdate.Where = objTriggeredUpdate.Where.Replace("@prm_", String.Format("[{0}].", BaseTable.PhysicalName))
        OnBankHolidayUpdate.AddIfNew(objTriggeredUpdate)
      End If

    End If

    SQLCode_AddCodeLevel(component.Components, childCodeCluster)

    If bAddDefaultDataType Then
      extraCode = New CodeElement
      extraCode.Code = CInt(component.Components(0).Components(0).ReturnType).ToString
      childCodeCluster.Add(extraCode)
    End If

    If bAddExpressionType Then
      extraCode = New CodeElement

      Select Case component.Components(2).Components(0).ReturnType
        Case ComponentValueTypes.ByRefDate
          extraCode.Code = "datetime"
        Case ComponentValueTypes.ByRefLogic
          extraCode.Code = "bit"
        Case ComponentValueTypes.ByRefNumeric
          extraCode.Code = "numeric"
        Case Else
          extraCode.Code = "string"
      End Select

      childCodeCluster.Add(extraCode)
    End If

    lineOfCode.Code = String.Format(lineOfCode.Code, childCodeCluster.ToArray)
    RequiresOvernight = RequiresOvernight Or objCodeLibrary.OvernightOnly
    _calculatePostAudit = _calculatePostAudit Or objCodeLibrary.CalculatePostAudit
    RequiresRecordId = RequiresRecordId Or objCodeLibrary.RecordIdRequired
    IsTimeDependant = IsTimeDependant Or objCodeLibrary.IsTimeDependant

    Tuning.Rating += objCodeLibrary.Tuning.Rating
    objCodeLibrary.Tuning.Usage += 1

    ' For functions that return mixed type, make it type safe
    If objCodeLibrary.ReturnType = ComponentValueTypes.Unknown And objCodeLibrary.MakeTypeSafe Then

      Select Case component.ReturnType
        Case ComponentValueTypes.Numeric
          lineOfCode.Code = String.Format("convert(numeric(38,8), ({0}))", lineOfCode.Code)

        Case ComponentValueTypes.Date
          lineOfCode.Code = String.Format("convert(datetime, ({0}))", lineOfCode.Code)

        Case ComponentValueTypes.Logic
          lineOfCode.Code = String.Format("{0}", lineOfCode.Code)

        Case ComponentValueTypes.String
          lineOfCode.Code = String.Format("convert(nvarchar(MAX), ({0}))", lineOfCode.Code)

      End Select

    End If

    CaseCount -= objCodeLibrary.CaseCount

    ' Attach the line of code
    codeCluster.Add(lineOfCode)

  End Sub

  Private Sub SQLCode_AddParameter(ByVal component As Component, ByVal codeCluster As ScriptDB.LinesOfCode, ByVal convertedFromColumn As Boolean)

    Dim childCodeCluster As ScriptDB.LinesOfCode
    Dim lineOfCode As CodeElement
    Dim objExpression As Expression
    Dim objColumn As Column

    ' Build code for the parameters
    childCodeCluster = New ScriptDB.LinesOfCode
    childCodeCluster.ReturnType = component.ReturnType
    childCodeCluster.CodeLevel = codeCluster.CodeLevel + 1

    ' Hack to hanld the first clause of an "if... then... else" function. The first parameter can be defined in all manner of ways that we need
    ' to make typesafe (i.e. if its a logic add a '= 1' at the end)
    If component.FunctionId = 4 And component.Parent.Components(0).Id = [component].Id Then
      childCodeCluster.CaseReturnType = CaseReturnType.Condition
    End If

    ' Nesting is too deep - convert to part number
    If CaseCount > 8 Then

      objExpression = New Expression
      objExpression.CaseCount = 0
      objExpression.ExpressionType = ExpressionType
      objExpression.BaseTable = BaseTable
      objExpression.AssociatedColumn = AssociatedColumn
      objExpression.BaseExpression = BaseExpression
      objExpression.ReturnType = component.ReturnType
      objExpression.Components = component.CloneComponents

      objExpression.Dependencies = Dependencies

      objExpression.GenerateCode()

      RequiresRecordId = RequiresRecordId Or objExpression.RequiresRecordId
      RequiresOvernight = RequiresOvernight Or objExpression.RequiresOvernight
      ReferencesParent = ReferencesParent Or objExpression.ReferencesParent
      ReferencesChild = ReferencesChild Or objExpression.ReferencesChild

      ' If first part of an if... then... else process slightly differently
      If childCodeCluster.CaseReturnType = CaseReturnType.Condition Then
        lineOfCode.Code = String.Format("{0} = 1", Dependencies.Add(objExpression))
      Else
        lineOfCode.Code = Dependencies.Add(objExpression)
      End If

      IsComplex = True

    Else

      SQLCode_AddCodeLevel(component.Components, childCodeCluster)
      lineOfCode.Code = String.Format("({0})", childCodeCluster.Statement)
    End If

    ' JIRA-2507 - Hack to handle problems with unique code
    If convertedFromColumn Then
      objColumn = BaseTable.Columns.GetById(component.ColumnId)
      If objColumn.CalculateIfEmpty Then
        lineOfCode.Code = String.Format("ISNULL(NULLIF(@prm_{0},{2}),{1})", objColumn.Name, lineOfCode.Code, objColumn.SafeReturnType)
      End If
    End If

    [codeCluster].Add(lineOfCode)

  End Sub

  Private Sub SQLCode_AddOperator(ByVal objComponent As Component, ByVal codeCluster As ScriptDB.LinesOfCode)

    Dim lineOfCode As CodeElement
    Dim objCodeLibrary As CodeLibrary

    lineOfCode.CodeType = ComponentTypes.Operator

    ' Get the bits and bobs for this operator
    objCodeLibrary = Operators.GetById(objComponent.OperatorId)

    If objCodeLibrary.PreCode.Length > 0 Then
      lineOfCode.Code = objCodeLibrary.PreCode
      codeCluster.InsertBeforePrevious(lineOfCode)
    End If

    lineOfCode.Code = String.Format(" {0} ", objCodeLibrary.Code)
    lineOfCode.OperatorType = objCodeLibrary.OperatorType
    [codeCluster].Add(lineOfCode)

    If objCodeLibrary.AfterCode.Length > 0 Then
      lineOfCode.CodeType = ComponentTypes.Value
      lineOfCode.Code = objCodeLibrary.AfterCode
      [codeCluster].AppendAfterNext(lineOfCode)
    End If

  End Sub

#End Region

  ' Declaration syntax for a column type
  Public ReadOnly Property DataTypeSyntax As String
    Get

      Dim sSqlType As String

      Select Case ReturnType
        Case ComponentValueTypes.Logic
          sSqlType = "bit"

        Case ComponentValueTypes.Numeric
          sSqlType = String.Format("numeric(38,8)")

        Case ComponentValueTypes.Date
          sSqlType = "datetime"

        Case ComponentValueTypes.String
          sSqlType = "varchar(MAX)"

        Case Else
          sSqlType = "varchar(MAX)"

      End Select
      Return sSqlType

    End Get

  End Property

  Private Function ResultWrapper(ByVal statement As String) As String

    Dim sWrapped As String = String.Empty
    Dim sSize As String

    If Options.OverflowSafety Then

      Select Case AssociatedColumn.DataType
        Case ColumnTypes.WorkingPattern
          sWrapped = statement

        Case ColumnTypes.Text, ColumnTypes.Link
          If AssociatedColumn.Multiline Then
            sWrapped = statement
          Else
            sWrapped = String.Format("SUBSTRING(ISNULL({0}, ''), 1, {1})", statement, AssociatedColumn.Size)
          End If

        Case ColumnTypes.Integer, ColumnTypes.Numeric
          If AssociatedColumn.Decimals > 0 Then
            sSize = String.Format("{0}.{1}", New String("9"c, AssociatedColumn.Size - AssociatedColumn.Decimals), New String("9"c, AssociatedColumn.Decimals))
          Else
            sSize = New String("9"c, AssociatedColumn.Size)
          End If
          sWrapped = String.Format("CASE WHEN ISNULL({0}, 0) > {1} OR ISNULL({0}, 0) < -{1} THEN 0 ELSE {0} END", statement, sSize)

        Case ColumnTypes.Date, ColumnTypes.Logic
          sWrapped = statement

      End Select

    Else
      sWrapped = statement
    End If

    Return sWrapped
  End Function

#Region "Cloning"

  Public Overloads Function Clone() As Expression

    Dim objClone As Expression

    ' Clone component properties (shallow clone)
    objClone = CType(MemberwiseClone(), Expression)

    ' Clone the child nodes (deep clone)
    objClone.Components = CloneComponents()

    Return objClone

  End Function

#End Region

End Class

