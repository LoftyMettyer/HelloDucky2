Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions

Namespace Things

  <ClassInterface(ClassInterfaceType.None), ComVisible(True), Serializable()>
  Public Class Expression
    Inherits Component

    Public Property Size As Integer
    Public Property Decimals As Integer
    Public Property BaseTableID As Integer
    Public Property BaseTable As Table

    Public UDF As ScriptDB.GeneratedUDF
    Public Property ExpressionType As ScriptDB.ExpressionType

    Public Property Dependencies As ExpressionDependencies

    Public Property Declarations As New ArrayList
    Public Property PreStatements As New ArrayList
    Public Property Joins As ArrayList
    Public Property FromTables As ArrayList
    Public Property Wheres As ArrayList

    Private _linesOfCode As ScriptDB.LinesOfCode
    Private _calculatePostAudit As Boolean

    Public Property CaseCount As Integer
    Public Property RequiresRecordID As Boolean
    Public Property RequiresOvernight As Boolean
    Public Property ReferencesParent As Boolean
    Public Property ReferencesChild As Boolean

    Public Property IsComplex As Boolean
    Public Property IsValid As Boolean = True

    Public Sub New()
      Dependencies = New ExpressionDependencies
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
          Case ScriptDB.ComponentTypes.Column
            bAddThis = True

          Case ScriptDB.ComponentTypes.Function, ScriptDB.ComponentTypes.Expression, ScriptDB.ComponentTypes.Calculation
            BuildDependancies(component)

          Case ScriptDB.ComponentTypes.ConvertedCalculatedColumn
            BuildDependancies(component)
            bAddThis = True

        End Select

        If bAddThis Then

          Dim table As Table = Globals.Tables.GetById(component.TableID)
          Dim column As Column = table.Columns.GetById(component.ColumnID)

          Dependencies.Add(column)
          Dependencies.Add(column.Table)

        End If

      Next

    End Sub

    Public Sub GenerateCodeForColumn()

      ' Build the dependencies collection
      Dependencies.Clear()
      BuildDependancies(Me)

      Me.IsComplex = False

      GenerateCode()
    End Sub

    Public Overridable Sub GenerateCode()

      Dim sOptions As String = String.Empty
      Dim sCode As String = String.Empty
      Dim aryDependsOn As New ArrayList
      Dim aryComments As New ArrayList
      Dim aryParameters1 As New ArrayList
      Dim aryParameters2 As New ArrayList
      Dim aryParameters3 As New ArrayList

      ' Initialise code object
      _linesOfCode = New ScriptDB.LinesOfCode
      _linesOfCode.Clear()
      _linesOfCode.ReturnType = ReturnType
      _linesOfCode.CodeLevel = If(Me.ExpressionType = ScriptDB.ExpressionType.ColumnFilter, 2, 1)

      Joins = New ArrayList
      FromTables = New ArrayList
      Wheres = New ArrayList

      Declarations.Clear()
      PreStatements.Clear()
      Joins.Clear()
      Wheres.Clear()

      ' If calculate only when empty add itself to the dependency stack
      If Me.AssociatedColumn.CalculateIfEmpty Then
        Dependencies.Add(Me.AssociatedColumn)
      End If

      aryParameters1.Clear()
      aryParameters2.Clear()
      aryParameters3.Clear()

      ' Build the execution code
      SQLCode_AddCodeLevel(Me.Components, _linesOfCode)
      '  Me.CaseCount = CInt(IIf(_linesOfCode.IsLogicBlock, Me.CaseCount + 1, Me.CaseCount))

      ' Add return declaration
      Declarations.Add(String.Format("@Result {0}", Me.DataTypeSyntax))

      ' Add the ID for the record if required
      '      If RequiresRecordID Or Me.IsComplex Or Me.ExpressionType = ScriptDB.ExpressionType.ColumnDefault Then

      If RequiresRecordID Or Me.ExpressionType = ScriptDB.ExpressionType.ColumnDefault Then
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
        If column.Table Is Me.BaseTable Then
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

      For Each objStatement As ScriptDB.GeneratedUDF In Dependencies.Statements
        Declarations.Add(objStatement.Declaration)
        PreStatements.Add(objStatement.Code)
      Next

      '.................... END


      For Each relation In Dependencies.Relations

        If Not aryParameters1.Contains(String.Format("@prm_ID_{0} integer", relation.ParentID)) Then
          aryParameters1.Add(String.Format("@prm_ID_{0} integer", relation.ParentID))

          If relation.RelationshipType = RelationshipType.Parent Then
            aryParameters2.Add(String.Format("base.[ID_{0}]", relation.ParentID))
            aryParameters3.Add(String.Format("@prm_ID_{0}", relation.ParentID))
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
      With UDF

        If Not Me.IsComplex Then
          .InlineCode = ResultWrapper(_linesOfCode.Statement)
          .InlineCode = .InlineCode.Replace("@prm_", "base.")
          .InlineCode = ScriptDB.Beautify.MakeSingleLine(.InlineCode)
        End If

        Me.Description = ScriptDB.Beautify.MakeSingleLine(Me.Description)

        .BoilerPlate = String.Format("-----------------------------------------------------------------" & vbNewLine & _
              "-- Generated by the Advanced System Framework" & vbNewLine & _
              "-- Column      : {1}.{0}" & vbNewLine & _
              "-- Expression  : {2}" & vbNewLine & _
              "-- Description : {7}" & vbNewLine & _
              "-- Depends on  : {3}" & vbNewLine & _
              "-- Date        : {4}" & vbNewLine & _
              "-- Complexity  : ({5}) {6}" & vbNewLine & _
              "----------------------------------------------------------------" & vbNewLine _
              , Me.AssociatedColumn.Name, Me.AssociatedColumn.Table.Name, Me.BaseExpression.Name _
              , String.Join(", ", aryDependsOn.ToArray()), Now().ToString _
              , Me.Tuning.Rating, Me.Tuning.ExpressionComplexity, Me.Description)
        .Declarations = If(Declarations.Count > 0, "DECLARE " & String.Join("," & vbNewLine & vbTab & vbTab & vbTab, Declarations.ToArray()) & ";" & vbNewLine, "")
        .Prerequisites = If(PreStatements.Count > 0, String.Join(vbNewLine, PreStatements.ToArray()) & vbNewLine & vbNewLine, "")
        .JoinCode = If(Joins.Count > 0, String.Format("{0}", String.Join(vbNewLine, Joins.ToArray)) & vbNewLine, "")
        .FromCode = If(FromTables.Count > 0, String.Format("{0}", String.Join(",", FromTables.ToArray)) & vbNewLine, "")
        .WhereCode = If(Wheres.Count > 0, String.Format("WHERE {0}", String.Join(" AND ", Wheres.ToArray)) & vbNewLine, "")

        ' Code beautify
        .Prerequisites = ScriptDB.Beautify.CleanWhitespace(.Prerequisites)

        Select Case Me.ExpressionType

          Case ScriptDB.ExpressionType.ColumnDefault
            .Name = String.Format("[{0}].[{1}{2}.{3}]", Me.SchemaName, ScriptDB.Consts.DefaultValueUDF, Me.AssociatedColumn.Table.Name, Me.AssociatedColumn.Name)
            .SelectCode = ScriptDB.Beautify.CleanWhitespace(_linesOfCode.Statement)
            .CallingCode = String.Format("{0}({1})", .Name, String.Join(",", aryParameters2.ToArray))
            .Code = String.Format("{11}CREATE FUNCTION {0}({1})" & vbNewLine & _
                           "RETURNS {2}" & vbNewLine & _
                           "{3}" & vbNewLine & _
                           "AS" & vbNewLine & "BEGIN" & vbNewLine & _
                           "    {4}" & vbNewLine & vbNewLine & _
                           "    {5}" & vbNewLine & vbNewLine & _
                           "    -- Execute calculation code" & vbNewLine & _
                           "    SELECT @Result = {6}" & vbNewLine & _
                           "                 {7}{8}{9}" & vbNewLine & _
                           "    RETURN {13};" & vbNewLine & _
                           "END" _
                          , .Name, String.Join(", ", aryParameters1.ToArray()) _
                          , Me.AssociatedColumn.DataTypeSyntax, sOptions, .Declarations, .Prerequisites, .SelectCode.Trim, .FromCode, .JoinCode, .WhereCode _
                          , Me.AssociatedColumn.SafeReturnType, .BoilerPlate, .Comments, ResultWrapper("@Result"))


            ' Wrapper for calculations with associated columns
          Case ScriptDB.ExpressionType.ColumnCalculation
            .Name = String.Format("[{0}].[{1}{2}.{3}]", Me.SchemaName, ScriptDB.Consts.CalculationUDF, Me.AssociatedColumn.Table.Name, Me.AssociatedColumn.Name)
            .SelectCode = _linesOfCode.Statement
            .CallingCode = String.Format("{0}({1})", .Name, String.Join(",", aryParameters2.ToArray))
            .Code = String.Format("{11}CREATE FUNCTION {0}({1})" & vbNewLine & _
                           "RETURNS {2}" & vbNewLine & _
                           "{3}" & vbNewLine & _
                           "AS" & vbNewLine & "BEGIN" & vbNewLine & _
                           "    {4}" & vbNewLine & vbNewLine & _
                           "    {5}" & vbNewLine & vbNewLine & _
                           "    -- Execute calculation code" & vbNewLine & _
                           "    SELECT @Result = {6}" & vbNewLine & _
                           "                 {7}{8}{9}" & vbNewLine & _
                           "    RETURN {13};" & vbNewLine & _
                           "END" _
                          , .Name, String.Join(", ", aryParameters1.ToArray()) _
                          , Me.AssociatedColumn.DataTypeSyntax, sOptions, .Declarations, .Prerequisites, .SelectCode.Trim, .FromCode, .JoinCode, .WhereCode _
                          , Me.AssociatedColumn.SafeReturnType, .BoilerPlate, .Comments, ResultWrapper("@Result"))

            ' Wrapper for when this function is used as a filter in an expression
          Case ScriptDB.ExpressionType.ColumnFilter
            .Name = String.Format("[{0}].[{1}{2}.{3}]", Me.SchemaName, ScriptDB.Consts.CalculationUDF, Me.AssociatedColumn.Table.Name, Me.AssociatedColumn.Name)
            .CallingCode = String.Format("{0}({1})", .Name, String.Join(",", aryParameters2.ToArray))
            .SelectCode = _linesOfCode.Statement

            ' Wrapper for when expression is used as a filter in a view
          Case ScriptDB.ExpressionType.Mask
            .Name = String.Format("[{0}].[{1}{2}]", Me.SchemaName, ScriptDB.Consts.MaskUDF, Me.BaseExpression.ID)
            .CallingCode = String.Format("{0}({1})", .Name, String.Join(",", aryParameters1.ToArray))
            .SelectCode = _linesOfCode.Statement

            .Code = String.Format("CREATE FUNCTION {0}(@prm_ID integer)" & vbNewLine & _
                           "RETURNS bit" & vbNewLine & _
                           "--WITH SCHEMABINDING" & vbNewLine & _
                           "AS" & vbNewLine & "BEGIN" & vbNewLine & vbNewLine & _
                           "{4}" & vbNewLine & vbNewLine & _
                           "{5}" & vbNewLine & vbNewLine & _
                           "    -- Execute calculation code" & vbNewLine & _
                           "    SELECT @Result = {6}" & vbNewLine & _
                           "                 {7}" & vbNewLine & _
                           "                 {8}" & vbNewLine & _
                           "                 {9}" & vbNewLine & _
                           "    RETURN ISNULL(@Result, 0);" & vbNewLine & _
                           "END" _
                          , .Name, String.Join(", ", aryParameters1.ToArray()) _
                          , "", "", .Declarations, .Prerequisites, .SelectCode.Trim, .FromCode, .JoinCode, .WhereCode)

          Case ScriptDB.ExpressionType.ReferencedColumn
            .Name = String.Format("[{0}].[{1}{2}.{3}]", Me.SchemaName, ScriptDB.Consts.CalculationUDF, Me.AssociatedColumn.Table.Name, Me.AssociatedColumn.Name)
            .CallingCode = String.Format("{0}({1})", .Name, String.Join(",", aryParameters3.ToArray))
            .SelectCode = _linesOfCode.Statement

          Case ScriptDB.ExpressionType.RecordDescription
            .Name = String.Format("[{0}].[{1}{2}]", Me.SchemaName, ScriptDB.Consts.RecordDescriptionUDF, Me.BaseTable.Name)
            .CallingCode = String.Format("{0}({1})", .Name, String.Join(",", aryParameters2.ToArray))
            .SelectCode = _linesOfCode.Statement

            .Code = String.Format("CREATE FUNCTION {0}({1})" & vbNewLine & _
                           "RETURNS nvarchar(MAX)" & vbNewLine & _
                           "{3}" & vbNewLine & _
                           "AS" & vbNewLine & "BEGIN" & vbNewLine & vbNewLine & _
                           "{4}" & vbNewLine & vbNewLine & _
                           "{5}" & vbNewLine & vbNewLine & _
                           "    -- Execute calculation code" & vbNewLine & _
                           "    SELECT @Result = {6}" & vbNewLine & _
                           "                 {7}" & vbNewLine & _
                           "                 {8}" & vbNewLine & _
                           "                 {9}" & vbNewLine & _
                           "    RETURN ISNULL(@Result, '');" & vbNewLine & _
                           "END" _
                          , .Name, String.Join(", ", aryParameters1.ToArray()) _
                          , "", sOptions, .Declarations, .Prerequisites, .SelectCode.Trim, .FromCode, .JoinCode, .WhereCode)

            ' Should never be called, but just in case...
          Case Else
            .SelectCode = _linesOfCode.Statement

        End Select

      End With

    End Sub

    Private Sub SQLCode_AddCodeLevel(ByVal [Components] As ICollection(Of Component), ByVal [CodeCluster] As ScriptDB.LinesOfCode)

      Dim objComponent As Component

      Dim guiObjectID As Integer

      Dim LineOfCode As ScriptDB.CodeElement
      Dim objCalculation As Expression

      For Each objComponent In [components]
        guiObjectID = objComponent.ID

        Select Case objComponent.SubType

          ' A table relationship
          Case ScriptDB.ComponentTypes.Relation
            SQLCode_AddRelation([CodeCluster], objComponent)
            ' Me.IsComplex = True

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
                LineOfCode.Code = String.Format("'{0}'", objComponent.ValueString.Replace("'", "''"))

              Case ScriptDB.ComponentValueTypes.Date
                LineOfCode.Code = String.Format("'{0}'", objComponent.ValueDate.ToString("yyyy-MM-dd"))

              Case ScriptDB.ComponentValueTypes.SystemVariable
                LineOfCode.Code = String.Format("{0}", objComponent.ValueString)

              Case Else
                LineOfCode.Code = String.Format("{0}", If(objComponent.ValueLogic, 1, 0))

            End Select

            [CodeCluster].Add(LineOfCode)


            ' Function component
          Case ScriptDB.ComponentTypes.Function
            SQLCode_AddFunction(objComponent, [CodeCluster])

            ' Calculated columns are sucked into this expressions
          Case ScriptDB.ComponentTypes.ConvertedCalculatedColumn
            SQLCode_AddParameter(objComponent, [CodeCluster], True)

            ' An expression or a parameter
          Case ScriptDB.ComponentTypes.Expression
            SQLCode_AddParameter(objComponent, [CodeCluster], False)

            ' Calculation 
          Case ScriptDB.ComponentTypes.Calculation

            If Not objComponent.BaseExpression.BaseTable.Expressions.GetById(objComponent.CalculationID) Is Nothing Then

              objCalculation = CType(objComponent.BaseExpression.BaseTable.Expressions.GetById(objComponent.CalculationID).Clone, Expression)

              'objCalculation.StartOfPartNumbers = 0
              objCalculation.BaseExpression = objComponent.BaseExpression
              objComponent.Components = objCalculation.CloneComponents
              objComponent.ReturnType = objCalculation.ReturnType
              SQLCode_AddParameter(objComponent, [CodeCluster], False)

            Else
              Globals.ErrorLog.Add(ErrorHandler.Section.General, Me.AssociatedColumn.Name, ErrorHandler.Severity.Error, _
                  "SQLCode_AddCodeLevel", Me.AssociatedColumn.Table.Name & "." & Me.AssociatedColumn.Name & " -- Missing calculation")
              Me.IsValid = False
              Me.IsComplex = True
            End If

            Me.IsComplex = True

        End Select

      Next

    End Sub

    Private Sub SQLCode_AddRelation(ByVal [CodeCluster] As ScriptDB.LinesOfCode, ByVal [Component] As Component)

      Dim objTable As Table
      Dim objRelation As Relation
      Dim LineOfCode As ScriptDB.CodeElement

      LineOfCode.CodeType = ScriptDB.ComponentTypes.Relation

      objTable = Globals.Tables.GetById([Component].TableID)
      objRelation = AssociatedColumn.Table.GetRelation(objTable.ID)

      Dependencies.Add(objRelation)
      Dependencies.Add(objTable)

      If Me.ExpressionType = ScriptDB.ExpressionType.Mask Then
        LineOfCode.Code = "@prm_ID"
      Else
        LineOfCode.Code = String.Format("@prm_ID_{0}", [Component].TableID)
      End If

      [CodeCluster].Add(LineOfCode)

    End Sub

    Private Sub SQLCode_AddColumn(ByVal [CodeCluster] As ScriptDB.LinesOfCode, ByVal [Component] As Component)

      Dim objThisColumn As Column

      Dim objRelation As Relation
      Dim sRelationCode As String
      Dim sFromCode As String
      Dim sWhereCode As String

      Dim sColumnFilter As String
      Dim sColumnJoinCode As String = String.Empty

      Dim iPartNumber As Integer
      Dim bIsSummaryColumn As Boolean
      Dim sColumnName As String

      Dim LineOfCode As ScriptDB.CodeElement

      LineOfCode.CodeType = ScriptDB.ComponentTypes.Column

      objThisColumn = CType(Dependencies.Columns.FirstOrDefault(Function(o) o.ID = Component.ColumnID), Column)
      objThisColumn.Tuning.Usage += 1

      ' Is this column referencing the column that this udf is attaching itself to? (i.e. recursion)
      If Component.IsColumnByReference Then
        LineOfCode.Code = String.Format("'{0}-{1}'" _
            , objThisColumn.Table.ID.ToString.PadLeft(8, "0"c) _
            , objThisColumn.ID.ToString.PadLeft(8, "0"c))

      ElseIf objThisColumn Is Me.AssociatedColumn _
          And Not (Me.ExpressionType = ScriptDB.ExpressionType.ColumnFilter _
          Or Me.ExpressionType = ScriptDB.ExpressionType.TriggeredUpdate _
          Or Me.ExpressionType = ScriptDB.ExpressionType.Mask _
          Or Me.ExpressionType = ScriptDB.ExpressionType.RecordDescription) Then

        If objThisColumn.SafeReturnType = "NULL" Then
          LineOfCode.Code = "@prm_" & objThisColumn.Name
        Else
          LineOfCode.Code = String.Format("ISNULL(@prm_{0},{1})", objThisColumn.Name, objThisColumn.SafeReturnType)
        End If

      ElseIf objThisColumn Is Me.AssociatedColumn _
          And Me.ExpressionType = ScriptDB.ExpressionType.ReferencedColumn Then
        LineOfCode.Code = String.Format("@prm_{0}", objThisColumn.Name)

        ' Does the referenced column have default value on it, then reference the UDF/value of the default rather than the column itself.
      ElseIf (Not objThisColumn.DefaultCalculation Is Nothing _
          And Me.ExpressionType = ScriptDB.ExpressionType.ColumnDefault _
          And objThisColumn.Table Is Me.AssociatedColumn.Table) Then
        LineOfCode.Code = String.Format("{0}(@prm_ID)", objThisColumn.DefaultCalculation.UDF.Name)

      Else

        'If is this column on the base table then add directly to the main execute statement,
        ' otherwise add it into child/parent statements array
        If objThisColumn.Table Is Me.AssociatedColumn.Table Then

          Select Case Component.BaseExpression.ExpressionType
            Case ScriptDB.ExpressionType.ColumnFilter, ScriptDB.ExpressionType.Mask
              sColumnName = String.Format("base.[{0}]", objThisColumn.Name)

              ' Needs base table added
              sFromCode = String.Format("FROM [dbo].[{0}] base", objThisColumn.Table.Name)
              If Not FromTables.Contains(sFromCode) Then
                FromTables.Add(sFromCode)
              End If

              '' Where clause
              'sWhereCode = String.Format("base.[ID] = @prm_ID")
              'If Not Wheres.Contains(sWhereCode) Then
              '  Wheres.Add(sWhereCode)
              'End If

              Me.IsComplex = True

              Dependencies.Add(objThisColumn)


            Case Else
              sColumnName = String.Format("@prm_{0}", objThisColumn.Name)

          End Select

          If objThisColumn.SafeReturnType = "NULL" Then
            LineOfCode.Code = sColumnName
          Else
            LineOfCode.Code = String.Format("ISNULL({0},{1})", sColumnName, objThisColumn.SafeReturnType)
          End If
        Else

          sColumnFilter = String.Empty
          RequiresRecordID = True
          bIsSummaryColumn = False
          Me.IsComplex = True

          objRelation = Me.BaseTable.GetRelation(objThisColumn.Table.ID)

          If objRelation.RelationshipType = RelationshipType.Parent Then

            If Me.ExpressionType = ScriptDB.ExpressionType.ColumnCalculation Then
              Me.AssociatedColumn.Table.DependsOnChildColumns.AddIfNew(objThisColumn)
            End If

            If objThisColumn.SafeReturnType = "NULL" Then
              LineOfCode.Code = String.Format("[{0}].[{1}]", objThisColumn.Table.Name, objThisColumn.Name)
            Else
              LineOfCode.Code = String.Format("ISNULL([{0}].[{1}],{2})", objThisColumn.Table.Name, objThisColumn.Name, objThisColumn.SafeReturnType)
            End If

            ' Add table join component
            sRelationCode = String.Format("LEFT JOIN [dbo].[{0}] ON [{0}].[ID] = base.[ID_{1}]", objRelation.Name, objRelation.ParentID)
            If Not Joins.Contains(sRelationCode) Then
              Joins.Add(sRelationCode)
            End If

            ' Needs base table added
            sFromCode = String.Format("FROM [dbo].[{0}] base", Me.AssociatedColumn.Table.Name)
            If Not FromTables.Contains(sFromCode) Then
              FromTables.Add(sFromCode)
            End If

            ' Where clause
            If Me.ExpressionType = ScriptDB.ExpressionType.ColumnFilter And Me.IsComplex Then
              sWhereCode = String.Format("[{0}].[ID] = @prm_ID_{1}", objRelation.Name, objRelation.ParentID)
            Else
              sWhereCode = "base.[ID] = @prm_ID"
            End If

            If Not Wheres.Contains(sWhereCode) Then
              Wheres.Add(sWhereCode)
            End If

            Me.ReferencesParent = True

          Else

            ' Add to dependency stack
            objThisColumn.Table.DependsOnParentColumns.AddIfNew(Me.AssociatedColumn)

            [Component].ChildRowDetails.BaseTable = Me.BaseTable
            [Component].ChildRowDetails.Order = objThisColumn.Table.TableOrders.GetById([Component].ChildRowDetails.OrderID)
            [Component].ChildRowDetails.Filter = objThisColumn.Table.Expressions.GetById([Component].ChildRowDetails.FilterID)
            [Component].ChildRowDetails.Relation = objRelation
            [Component].ChildRowDetails.Column = objThisColumn
            iPartNumber = Dependencies.Add([Component].ChildRowDetails)

            ' Any columns used in child filters should be added to the udf chain 
            If Not [Component].ChildRowDetails.Filter Is Nothing Then
              For Each objColumn In [Component].ChildRowDetails.Filter.Dependencies.Columns
                Me.Dependencies.Add(objColumn)
              Next
            End If


            LineOfCode.Code = String.Format("@child_{0}", iPartNumber)
            Me.ReferencesChild = True

            End If
        End If
      End If

      ' Add this column (or reference to it) to the main execute statement
      [CodeCluster].Add(LineOfCode)

    End Sub

    Private Sub SQLCode_AddFunction(ByVal component As Component, ByVal codeCluster As ScriptDB.LinesOfCode)

      Dim LineOfCode As ScriptDB.CodeElement
      Dim ExtraCode As ScriptDB.CodeElement

      Dim objCodeLibrary As CodeLibrary
      Dim ChildCodeCluster As ScriptDB.LinesOfCode
      Dim WhereCodeCluster As ScriptDB.LinesOfCode
      Dim objSetting As Setting
      Dim objIDComponent As Component
      Dim objTriggeredUpdate As ScriptDB.TriggeredUpdate
      Dim sWhereClause As String = ""
      Dim bAddDefaultDataType As Boolean

      LineOfCode.CodeType = ScriptDB.ComponentTypes.Function
      objCodeLibrary = Globals.Functions.GetById(component.FunctionID)
      LineOfCode.Code = objCodeLibrary.Code
      Me.CaseCount += objCodeLibrary.CaseCount

      ' Get parameters
      ChildCodeCluster = New ScriptDB.LinesOfCode
      ChildCodeCluster.CodeLevel = codeCluster.CodeLevel + 1
      ChildCodeCluster.ReturnType = objCodeLibrary.ReturnType

      ' Add module dependancy info for this function
      If objCodeLibrary.Dependancies.Count > 0 Then
        For Each objSetting In objCodeLibrary.Dependancies

          Select Case objSetting.SettingType

            Case SettingType.ModuleSetting
              objIDComponent = New Component
              objIDComponent.SubType = ScriptDB.ComponentTypes.Relation
              objIDComponent.TableID = CInt(objSetting.Value)
              component.Components.Add(objIDComponent)
              Me.IsComplex = True

            Case SettingType.CodeItem
              objIDComponent = New Component
              objIDComponent.SubType = ScriptDB.ComponentTypes.Value
              objIDComponent.ValueString = objSetting.Code
              objIDComponent.ValueType = ScriptDB.ComponentValueTypes.SystemVariable
              component.Components.Add(objIDComponent)

            Case SettingType.UpdateParameter
              sWhereClause = objSetting.Code

            Case SettingType.DefaultDataType
              bAddDefaultDataType = True

          End Select

        Next
      End If

      ' Does this component need adding to the 'Get Field From Database' stack?
      If objCodeLibrary.IsGetFieldFromDB Then
        Globals.GetFieldsFromDB.Add(component)
      End If

      ' Is this expression reliant on the bank holiday table (I'm sure this can be tidyied up)
      If objCodeLibrary.DependsOnBankHoliday And Me.ExpressionType = ScriptDB.ExpressionType.ColumnCalculation Then ' And Not Me.IsComplex Then
        objTriggeredUpdate = New ScriptDB.TriggeredUpdate
        objTriggeredUpdate.Column = Me.AssociatedColumn
        objTriggeredUpdate.ID = Me.AssociatedColumn.ID

        ' Get parameters
        WhereCodeCluster = New ScriptDB.LinesOfCode
        SQLCode_AddCodeLevel(component.Components, WhereCodeCluster)
        objTriggeredUpdate.Where = String.Format(sWhereClause, WhereCodeCluster.ToArray)

        If Not Me.ReferencesChild And Not objTriggeredUpdate.Where.Contains("@part_") Then
          objTriggeredUpdate.Where = objTriggeredUpdate.Where.Replace("@prm_", String.Format("[{0}].", Me.BaseTable.PhysicalName))
          Globals.OnBankHolidayUpdate.AddIfNew(objTriggeredUpdate)
        End If

      End If

      SQLCode_AddCodeLevel(component.Components, ChildCodeCluster)

      If bAddDefaultDataType Then
        ExtraCode = New ScriptDB.CodeElement
        ExtraCode.Code = CInt(component.Components(0).Components(0).ReturnType).ToString
        ChildCodeCluster.Add(ExtraCode)
      End If


      LineOfCode.Code = String.Format(LineOfCode.Code, ChildCodeCluster.ToArray)
      RequiresOvernight = RequiresOvernight Or objCodeLibrary.OvernightOnly
      _calculatePostAudit = _calculatePostAudit Or objCodeLibrary.CalculatePostAudit
      Me.RequiresRecordID = Me.RequiresRecordID Or objCodeLibrary.RecordIDRequired
      Me.IsTimeDependant = Me.IsTimeDependant Or objCodeLibrary.IsTimeDependant

      Me.Tuning.Rating += objCodeLibrary.Tuning.Rating
      objCodeLibrary.Tuning.Usage += 1

      ' For functions that return mixed type, make it type safe
      If objCodeLibrary.ReturnType = ScriptDB.ComponentValueTypes.Unknown And objCodeLibrary.MakeTypeSafe Then

        Select Case component.ReturnType
          Case ScriptDB.ComponentValueTypes.Numeric
            LineOfCode.Code = String.Format("convert(numeric(38,8), ({0}))", LineOfCode.Code)

          Case ScriptDB.ComponentValueTypes.Date
            LineOfCode.Code = String.Format("convert(datetime, ({0}))", LineOfCode.Code)

          Case ScriptDB.ComponentValueTypes.Logic
            LineOfCode.Code = String.Format("{0}", LineOfCode.Code)

          Case ScriptDB.ComponentValueTypes.String
            LineOfCode.Code = String.Format("convert(nvarchar(MAX), ({0}))", LineOfCode.Code)

        End Select

      End If

      Me.CaseCount -= objCodeLibrary.CaseCount

      ' Attach the line of code
      codeCluster.Add(LineOfCode)

    End Sub

    Private Sub SQLCode_AddParameter(ByVal [Component] As Component, ByVal [CodeCluster] As ScriptDB.LinesOfCode, ByVal ConvertedFromColumn As Boolean)

      Dim ChildCodeCluster As ScriptDB.LinesOfCode
      Dim LineOfCode As ScriptDB.CodeElement
      Dim objExpression As Expression
      Dim objColumn As Column

      ' Build code for the parameters
      ChildCodeCluster = New ScriptDB.LinesOfCode
      ChildCodeCluster.ReturnType = Component.ReturnType
      ChildCodeCluster.CodeLevel = CodeCluster.CodeLevel + 1

      ' Hack to hanld the first clause of an "if... then... else" function. The first parameter can be defined in all manner of ways that we need
      ' to make typesafe (i.e. if its a logic add a '= 1' at the end)
      If Component.FunctionID = 4 And Component.Parent.Components(0).ID = [Component].ID Then
        ChildCodeCluster.MakeTypesafe = False
        'Me.CaseCount = Me.CaseCount + 2
      End If

      ' Nesting is too deep - convert to part number
      If Me.CaseCount > 8 Then

        objExpression = New Expression
        objExpression.CaseCount = 0
        objExpression.ExpressionType = Me.ExpressionType
        objExpression.BaseTable = Me.BaseTable
        objExpression.AssociatedColumn = Me.AssociatedColumn
        objExpression.BaseExpression = Me.BaseExpression
        objExpression.ReturnType = Component.ReturnType
        objExpression.Components = Component.CloneComponents

        objExpression.Dependencies = Me.Dependencies

        objExpression.GenerateCode()

        Me.RequiresRecordID = RequiresRecordID Or objExpression.RequiresRecordID
        '   Me.ContainsUniqueCode = ContainsUniqueCode Or objExpression.ContainsUniqueCode
        Me.RequiresOvernight = RequiresOvernight Or objExpression.RequiresOvernight
        Me.ReferencesParent = Me.ReferencesParent Or objExpression.ReferencesParent
        Me.ReferencesChild = Me.ReferencesChild Or objExpression.ReferencesChild

        ' If first part of an if... then... else process slightly differently
        If Not ChildCodeCluster.MakeTypesafe And objExpression.ReturnType = ScriptDB.ComponentValueTypes.Logic Then
          LineOfCode.Code = String.Format("{0} = 1", Dependencies.Add(objExpression))
        Else
          LineOfCode.Code = Dependencies.Add(objExpression)
        End If

        Me.IsComplex = True

      Else

        SQLCode_AddCodeLevel(Component.Components, ChildCodeCluster)
        LineOfCode.Code = String.Format("({0})", ChildCodeCluster.Statement)
      End If

      ' JIRA-2507 - Hack to handle problems with unique code
      If ConvertedFromColumn Then
        objColumn = BaseTable.Columns.GetById(Component.ColumnID)
        If objColumn.CalculateIfEmpty Then
          LineOfCode.Code = String.Format("ISNULL(NULLIF(@prm_{0},{2}),{1})", objColumn.Name, LineOfCode.Code, objColumn.SafeReturnType)
        End If
      End If

      [CodeCluster].Add(LineOfCode)

    End Sub

    Private Sub SQLCode_AddOperator(ByVal objComponent As Component, ByVal [CodeCluster] As ScriptDB.LinesOfCode)

      Dim LineOfCode As ScriptDB.CodeElement
      Dim objCodeLibrary As CodeLibrary

      LineOfCode.CodeType = ScriptDB.ComponentTypes.Operator

      ' Get the bits and bobs for this operator
      objCodeLibrary = Globals.Operators.GetById(objComponent.OperatorID)

      If objCodeLibrary.PreCode.Length > 0 Then
        LineOfCode.Code = objCodeLibrary.PreCode
        CodeCluster.InsertBeforePrevious(LineOfCode)
      End If

      LineOfCode.Code = String.Format(" {0} ", objCodeLibrary.Code)
      LineOfCode.OperatorType = objCodeLibrary.OperatorType
      [CodeCluster].Add(LineOfCode)

      If objCodeLibrary.AfterCode.Length > 0 Then
        LineOfCode.CodeType = ScriptDB.ComponentTypes.Value
        LineOfCode.Code = objCodeLibrary.AfterCode
        [CodeCluster].AppendAfterNext(LineOfCode)
      End If

    End Sub

#End Region

    ' Declaration syntax for a column type
    Public ReadOnly Property DataTypeSyntax As String
      Get

        Dim sSQLType As String = String.Empty

        Select Case Me.ReturnType
          Case ScriptDB.ComponentValueTypes.Logic
            sSQLType = "bit"

          Case ScriptDB.ComponentValueTypes.Numeric
            sSQLType = String.Format("numeric(38,8)")

          Case ScriptDB.ComponentValueTypes.Date
            sSQLType = "datetime"

          Case ScriptDB.ComponentValueTypes.String
            sSQLType = "varchar(MAX)"

          Case Else
            sSQLType = "varchar(MAX)"

        End Select
        Return sSQLType

      End Get

    End Property

    Private Function ResultWrapper(ByVal Statement As String) As String

      Dim sWrapped As String = String.Empty
      Dim sSize As String = String.Empty

      If Globals.Options.OverflowSafety Then

        Select Case Me.AssociatedColumn.DataType
          Case ColumnTypes.WorkingPattern
            sWrapped = Statement

          Case ColumnTypes.Text, ColumnTypes.Link
            If Me.AssociatedColumn.Multiline Then
              sWrapped = Statement
            Else
              sWrapped = String.Format("SUBSTRING(ISNULL({0}, ''), 1, {1})", Statement, Me.AssociatedColumn.Size)
            End If

          Case ColumnTypes.Integer, ColumnTypes.Numeric
            If Me.AssociatedColumn.Decimals > 0 Then
              sSize = String.Format("{0}.{1}", New String("9"c, Me.AssociatedColumn.Size - Me.AssociatedColumn.Decimals), New String("9"c, Me.AssociatedColumn.Decimals))
            Else
              sSize = New String("9"c, Me.AssociatedColumn.Size)
            End If
            sWrapped = String.Format("CASE WHEN ISNULL({0}, 0) > {1} OR ISNULL({0}, 0) < -{1} THEN 0 ELSE {0} END", Statement, sSize)

          Case ColumnTypes.Date, ColumnTypes.Logic
            sWrapped = Statement


        End Select

      Else
        sWrapped = Statement
      End If

      Return sWrapped
    End Function

#Region "Cloning"

    Public Overloads Function Clone() As Expression

      Dim objClone As New Expression

      ' Clone component properties (shallow clone)
      objClone = CType(Me.MemberwiseClone(), Expression)

      ' Clone the child nodes (deep clone)
      objClone.Components = Me.CloneComponents

      Return objClone

    End Function

#End Region

  End Class
End Namespace

