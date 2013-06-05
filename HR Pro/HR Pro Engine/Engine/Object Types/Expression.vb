Option Explicit On

Namespace Things

  <Serializable()> _
  Public Class Expression
    Inherits Things.Component

    Public Size As Integer
    Public Decimals As Integer
    Public BaseTableID As HCMGuid

    <System.Xml.Serialization.XmlIgnore()> _
    Public BaseTable As Things.Table

    <System.Xml.Serialization.XmlIgnore()> _
    Public AssociatedColumn As Things.Column

    Public UDF As ScriptDB.GeneratedUDF
    Public ExpressionType As ScriptDB.ExpressionType

    Public DependsOnColumns As New Things.Collection
    Public Dependencies As New Things.Collection
    Private mcolOrders As Things.Collection

    Public StatementObjects As New ArrayList

    Public Joins As ArrayList
    Public FromTables As ArrayList
    Private Wheres As ArrayList
    Public Declarations As New ArrayList
    Public PreStatements As New ArrayList
    Public ChildColumns As Things.Collection

    Private mcolLinesOfCode As ScriptDB.LinesOfCode

    Public CaseCount As Integer = 0
    Public StartOfPartNumbers As Integer = 0
    Public RequiresRecordID As Boolean = False
    Public RequiresRowNumber As Boolean = False
    Public RequiresOvernight As Boolean = False
    Public ContainsUniqueCode As Boolean = False
    Public ReferencesParent As Boolean = False

    Public IsComplex As Boolean = False
    Public IsValid As Boolean = True

    Private mbCalculatePostAudit As Boolean = False
    Private mbNeedsOriginalValue As Boolean = False
    Private mbCheckTriggerStack As Boolean = False

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

    Private Sub BuildDependancies(ByRef objExpression As Things.Component)

      'Dim objColumn As Things.Column
      Dim objComponent As Things.Component
      '      Dim objDependency As Things.Base
      Dim objColumn As Things.Column

      For Each objComponent In objExpression.Objects
        Select Case objComponent.SubType
          Case ScriptDB.ComponentTypes.Column

            objColumn = Globals.Things.GetObject(Enums.Type.Table, objComponent.TableID).Objects.GetObject(Enums.Type.Column, objComponent.ColumnID)
            If Not Dependencies.Contains(objColumn) Then
              Dependencies.Add(objColumn)
            End If

          Case ScriptDB.ComponentTypes.Expression, ScriptDB.ComponentTypes.Function, ScriptDB.ComponentTypes.Calculation
            BuildDependancies(objComponent)
        End Select

      Next

    End Sub

    Public Overridable Sub GenerateCode()

      Dim objDependency As Things.Base
      Dim sOptions As String = String.Empty
      Dim sCode As String = String.Empty
      Dim sBypassUDFCode As String = String.Empty
      Dim aryDependsOn As New ArrayList
      Dim aryComments As New ArrayList
      Dim aryParameters1 As New ArrayList
      Dim aryParameters2 As New ArrayList
      Dim aryParameters3 As New ArrayList

      ' Initialise code object
      Me.IsComplex = False
      Me.CaseCount = 0
      mcolLinesOfCode = New ScriptDB.LinesOfCode
      mcolLinesOfCode.Clear()
      mcolLinesOfCode.ReturnType = ReturnType
      mcolLinesOfCode.CodeLevel = IIf(Me.ExpressionType = ScriptDB.ExpressionType.ColumnFilter, 2, 1)

      Joins = New ArrayList
      FromTables = New ArrayList
      Wheres = New ArrayList

      Declarations.Clear()
      PreStatements.Clear()

      Joins.Clear()
      Wheres.Clear()
      StatementObjects.Clear()

      ' Build the dependencies collection
      Dependencies.Clear()
      BuildDependancies(Me)

      aryParameters1.Clear()
      aryParameters2.Clear()
      aryParameters3.Clear()

      ' Build the execution code
      SQLCode_AddCodeLevel(Me.Objects, mcolLinesOfCode)

      ' Always add the ID for the record
      If RequiresRecordID Or Me.IsComplex Then
        aryParameters1.Add("@prm_ID integer")
        aryParameters2.Add("base.ID")
        aryParameters3.Add("@prm_ID")
      End If

      ' Some function require the row number of the record as a parameter
      If RequiresRowNumber Then
        aryParameters1.Add("@rownumber integer")
        aryParameters2.Add("[rownumber]")
        aryParameters3.Add("@rownumber")
      End If

      ' Some function require the row number of the record as a parameter
      If RequiresOvernight Then
        aryParameters1.Add("@isovernight bit")
        aryParameters2.Add("@isovernight")
        aryParameters3.Add("@isovernight")
      End If


      ' Is this function going to check to see if any of the dependant tables were called as part of this transaction?
      'If mbCheckTriggerStack Then
      'aryParameters1.Add(String.Format("@originalvalue {0}", Me.AssociatedColumn.DataTypeSyntax))
      'aryParameters2.Add(String.Format("base.[{0}]", Me.AssociatedColumn.Name))
      'aryParameters3.Add("@originalvalue")
      ' End If


      ' Add other dependancies
      For Each objDependency In Dependencies
        If objDependency.Type = Enums.Type.Column Then
          If CType(objDependency, Things.Column).Table Is Me.BaseTable Then
            aryParameters1.Add(String.Format("@prm_{0} {1}", objDependency.Name, CType(objDependency, Things.Column).DataTypeSyntax))
            aryParameters2.Add(String.Format("base.[{0}]", objDependency.Name))
            aryParameters3.Add(String.Format("@prm_{0}", objDependency.Name))
            aryComments.Add(String.Format("Column: {0}", objDependency.Name))
          End If
        End If

        If objDependency.Type = Enums.Type.Relation Then

          If Not aryParameters1.Contains(String.Format("@prm_ID_{0} integer", CInt(CType(objDependency, Things.Relation).ParentID))) Then
            aryParameters1.Add(String.Format("@prm_ID_{0} integer", CInt(CType(objDependency, Things.Relation).ParentID)))

            If CType(objDependency, Things.Relation).RelationshipType = ScriptDB.RelationshipType.Parent Then
              aryParameters2.Add(String.Format("base.[ID_{0}]", CInt(CType(objDependency, Things.Relation).ParentID)))
              aryParameters3.Add(String.Format("@prm_ID_{0}", CInt(CType(objDependency, Things.Relation).ParentID)))
              aryComments.Add(String.Format("Relation :{0}", objDependency.Name))
            Else
              aryParameters2.Add("base.[ID]")
              aryParameters3.Add(String.Format("@prm_ID"))
              aryComments.Add(String.Format("Relation : {0}", objDependency.Name))
            End If

          End If

        End If

        If objDependency.Type = Enums.Type.Table Then
          aryComments.Add(String.Format("Table : {0}", objDependency.Name))
          aryDependsOn.Add(String.Format("{0}", CInt(objDependency.ID)))
        End If

      Next

      ' Do we have caching on this UDF?
      If Me.IsComplex And aryDependsOn.Count > 0 And Not Me.ReferencesParent Then

        ' Flag to force updates through
        If Me.ExpressionType = ScriptDB.ExpressionType.ReferencedColumn Then
          aryParameters1.Add("@forcerefresh bit")
          aryParameters2.Add("@forcerefresh")
          aryParameters3.Add("1")
        Else
          aryParameters1.Add("@forcerefresh bit")
          aryParameters2.Add("@forcerefresh")
          aryParameters3.Add("@forcerefresh")
        End If

        sBypassUDFCode = String.Format("    -- Return the original value if none of the dependent tables are in the trigger stack." & vbNewLine &
            "    IF @forcerefresh = 0 AND NOT EXISTS (SELECT [tablefromid] FROM [dbo].[tbsys_intransactiontrigger] WHERE [tablefromid] IN ({0}) AND [spid] = @@SPID)" & vbNewLine & _
            "        BEGIN" & vbNewLine & _
            "            SELECT @result = [{1}] FROM dbo.[{2}] WHERE [ID] = @prm_ID;" & vbNewLine & _
            "            RETURN @result;" & vbNewLine & _
            "        END" _
            , String.Join(", ", aryDependsOn.ToArray()), Me.AssociatedColumn.Name, Me.AssociatedColumn.Table.PhysicalName)
      End If

      ' Can object be schemabound
      If Me.BaseExpression.IsSchemaBound Then
        sOptions = "--WITH SCHEMABINDING"
      End If

      ' Calling statement
      With UDF

        If Not Me.IsComplex Then
          .InlineCode = ResultWrapper(mcolLinesOfCode.Statement)
          .InlineCode = .InlineCode.Replace("@prm_", "base.")
          .InlineCode = .InlineCode.Replace("@rownumber", "[rownumber]")
          .InlineCode = ScriptDB.Beautify.MakeSingleLine(.InlineCode)
        End If

        Me.Description = ScriptDB.Beautify.MakeSingleLine(Me.Description)

        .BoilerPlate = String.Format("-----------------------------------------------------------------" & vbNewLine & _
              "-- Auto generated by the Advanced .NET Database Scripting Engine" & vbNewLine & _
              "-- Column      : {1}.{0}" & vbNewLine & _
              "-- Expression  : {2}" & vbNewLine & _
              "-- Description : {8}" & vbNewLine & _
              "-- Depends on  : {3}" & vbNewLine & _
              "-- Date        : {4}" & vbNewLine & _
              "-- Complexity  : ({5}) {6}" & vbNewLine & _
              "/*{7}*/" & vbNewLine & _
              "----------------------------------------------------------------" & vbNewLine _
              , Me.AssociatedColumn.Name, Me.AssociatedColumn.Table.Name, Me.BaseExpression.Name _
              , String.Join(", ", aryDependsOn.ToArray()), Now().ToString _
              , Me.Tuning.Rating, Me.Tuning.ExpressionComplexity, .InlineCode, Me.Description)
        .Declarations = IIf(Declarations.Count > 0, "DECLARE " & String.Join("," & vbNewLine, Declarations.ToArray()) & ";" & vbNewLine, "")
        .Prerequisites = IIf(PreStatements.Count > 0, String.Join(vbNewLine, PreStatements.ToArray()) & vbNewLine & vbNewLine, "")
        .JoinCode = IIf(Joins.Count > 0, String.Format("{0}", String.Join(vbNewLine, Joins.ToArray)) & vbNewLine, "")
        .FromCode = IIf(FromTables.Count > 0, String.Format("{0}", String.Join(",", FromTables.ToArray)) & vbNewLine, "")
        .WhereCode = IIf(Wheres.Count > 0, String.Format("WHERE {0}", String.Join(" AND ", Wheres.ToArray)) & vbNewLine, "")

        ' Code beautify
        ScriptDB.Beautify.Cleanwhitespace(.Prerequisites)

        Select Case Me.ExpressionType

          ' Wrapper for calculations with associated columns
          Case ScriptDB.ExpressionType.ColumnCalculation
            .Name = String.Format("[{0}].[{1}{2}.{3}]", Me.SchemaName, ScriptDB.Consts.CalculationUDF, Me.AssociatedColumn.Table.Name, Me.AssociatedColumn.Name)
            .SelectCode = mcolLinesOfCode.Statement
            .CallingCode = String.Format("{0}({1})", .Name, String.Join(",", aryParameters2.ToArray))
            .Code = String.Format("{11}CREATE FUNCTION {0}({1})" & vbNewLine & _
                           "RETURNS {2}" & vbNewLine & _
                           "{3}" & vbNewLine & _
                           "AS" & vbNewLine & "BEGIN" & vbNewLine & _
                           "    DECLARE @Result AS {15};" & vbNewLine & vbNewLine & _
                           "    {13}" & vbNewLine & vbNewLine & _
                           "    {4}{5}" & vbNewLine & vbNewLine & _
                           "    -- Execute calculation code" & vbNewLine & _
                           "    SELECT @Result = {6}" & vbNewLine & _
                           "                 {7}{8}{9}" & vbNewLine & _
                           "    RETURN {14};" & vbNewLine & _
                           "END" _
                          , .Name, String.Join(", ", aryParameters1.ToArray()) _
                          , Me.AssociatedColumn.DataTypeSyntax, sOptions, .Declarations, .Prerequisites, .SelectCode.Trim, .FromCode, .JoinCode, .WhereCode _
                          , Me.AssociatedColumn.SafeReturnType, .BoilerPlate, .Comments, sBypassUDFCode, ResultWrapper("@Result"), ResultDefinition)

            .CodeStub = String.Format("CREATE FUNCTION {0}({1})" & vbNewLine & _
                           "RETURNS {2}" & vbNewLine & _
                           "{3}" & vbNewLine & _
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
            .Name = String.Format("[{0}].[{1}{2}.{3}]", Me.SchemaName, ScriptDB.Consts.CalculationUDF, Me.AssociatedColumn.Table.Name, Me.AssociatedColumn.Name)
            .CallingCode = String.Format("{0}({1})", .Name, String.Join(",", aryParameters2.ToArray))
            '.SelectCode = String.Format("CASE WHEN ({0}) THEN 1 ELSE 0 END", mcolLinesOfCode.Statement)
            .SelectCode = mcolLinesOfCode.Statement

            ' Wrapper for when expression is used as a filter in a view
          Case ScriptDB.ExpressionType.Mask
            .Name = String.Format("[{0}].[{1}{2}]", Me.SchemaName, ScriptDB.Consts.MaskUDF, CInt(Me.BaseExpression.ID))
            .CallingCode = String.Format("{0}({1})", .Name, String.Join(",", aryParameters1.ToArray))
            .SelectCode = mcolLinesOfCode.Statement

            .Code = String.Format("CREATE FUNCTION {0}(@prm_ID integer)" & vbNewLine & _
                           "RETURNS bit" & vbNewLine & _
                           "--WITH SCHEMABINDING" & vbNewLine & _
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

            .CodeStub = String.Format("CREATE FUNCTION {0}(@prm_ID integer)" & vbNewLine & _
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

    End Sub

    Private Sub SQLCode_AddCodeLevel(ByRef [Components] As Things.Collection, ByRef [CodeCluster] As ScriptDB.LinesOfCode)

      Dim objComponent As Things.Component

      Dim guiObjectID As HCMGuid

      Dim iValueSubType As ScriptDB.ColumnTypes = Nothing

      Dim LineOfCode As ScriptDB.CodeElement
      Dim objCalculation As Things.Expression

      For Each objComponent In [Components]
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
                LineOfCode.Code = String.Format("{0}", IIf(objComponent.ValueLogic, 1, 0))

            End Select

            [CodeCluster].Add(LineOfCode)


            ' Function component
          Case ScriptDB.ComponentTypes.Function
            SQLCode_AddFunction(objComponent, [CodeCluster])

            ' An expression or a parameter
          Case ScriptDB.ComponentTypes.Expression
            SQLCode_AddParameter(objComponent, [CodeCluster])
            '            Me.IsComplex = True

            ' Expression 
          Case ScriptDB.ComponentTypes.Calculation, ScriptDB.ComponentTypes.Expression

            If Not objComponent.BaseExpression.BaseTable.Objects.GetObject(Enums.Type.Expression, objComponent.CalculationID) Is Nothing Then

              'objCalculation = Me.AssociatedColumn.Table

              ' There has to be a cleaner way, but once again I'm in  a hurry and this DOES work. Amazingly!
              ' I'm still in a hurry!!! :-(
              objCalculation = CType(objComponent.BaseExpression.BaseTable.Objects.GetObject(Enums.Type.Expression, objComponent.CalculationID), Things.Expression).Clone
              'objCalculation.StartOfPartNumbers = 0
              objCalculation.BaseExpression = objComponent.BaseExpression
              objComponent.Objects = objCalculation.Objects
              objComponent.ReturnType = objCalculation.ReturnType
              SQLCode_AddParameter(objComponent, [CodeCluster])

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

    Private Sub SQLCode_AddRelation(ByRef [CodeCluster] As ScriptDB.LinesOfCode, ByRef [Component] As Things.Component)

      Dim objTable As Things.Table
      Dim objRelation As Things.Relation
      Dim LineOfCode As ScriptDB.CodeElement

      LineOfCode.CodeType = ScriptDB.ComponentTypes.Relation

      objTable = Globals.Things.GetObject(Enums.Type.Table, [Component].TableID)
      objRelation = AssociatedColumn.Table.GetRelation(objTable.ID)

      If Not Dependencies.Contains(objRelation) Then
        Dependencies.Add(objRelation)
      End If

      Dependencies.AddIfNew(objTable)

      LineOfCode.Code = String.Format("@prm_ID_{0}", CInt([Component].TableID))

      [CodeCluster].Add(LineOfCode)

    End Sub

    Private Sub SQLCode_AddColumn(ByRef [CodeCluster] As ScriptDB.LinesOfCode, ByRef [Component] As Things.Component)

      Dim objThisColumn As Things.Column

      Dim objRelation As Things.Relation
      Dim objOrderFilter As Things.TableOrderFilter
      Dim sRelationCode As String
      Dim sFromCode As String
      Dim sWhereCode As String

      Dim sColumnFilter As String
      Dim sColumnJoinCode As String = String.Empty

      Dim iBackupType As ScriptDB.ExpressionType
      Dim sPartCode As String
      Dim iPartNumber As Integer
      Dim bIsSummaryColumn As Boolean
      Dim sColumnName As String

      Dim LineOfCode As ScriptDB.CodeElement

      LineOfCode.CodeType = ScriptDB.ComponentTypes.Column

      objThisColumn = Dependencies.GetObject(Enums.Type.Column, Component.ColumnID)
      objThisColumn.Tuning.Usage += 1

      Dependencies.AddIfNew(objThisColumn.Table)
      Dependencies.AddIfNew(objThisColumn)

      ' Is this column referencing the column that this udf is attaching itself to? (i.e. recursion)
      If Component.IsColumnByReference Then
        LineOfCode.Code = String.Format("'{0}-{1}'" _
            , CInt(objThisColumn.Table.ID).ToString.PadLeft(8, "0") _
            , CInt(objThisColumn.ID).ToString.PadLeft(8, "0"))
        'Me.IsComplex = True

      ElseIf Me.ExpressionType = ScriptDB.ExpressionType.TriggeredUpdate Then
        LineOfCode.Code = String.Format("[{0}].[{1}]", objThisColumn.Table.PhysicalName, objThisColumn.Name)

      ElseIf objThisColumn Is Me.AssociatedColumn _
          And Not (Me.ExpressionType = ScriptDB.ExpressionType.ColumnFilter _
          Or Me.ExpressionType = ScriptDB.ExpressionType.Mask _
          Or Me.ExpressionType = ScriptDB.ExpressionType.RecordDescription) Then
        LineOfCode.Code = String.Format("@prm_{0}", objThisColumn.Name)

      ElseIf objThisColumn Is Me.AssociatedColumn _
        And Me.ExpressionType = ScriptDB.ExpressionType.ReferencedColumn Then
        LineOfCode.Code = String.Format("@prm_{0}", objThisColumn.Name)

        ' Does the referenced column have default value on it, then reference the UDF/value of the default rather than the column itself.
      ElseIf (Not objThisColumn.DefaultCalcID = 0 And Me.ExpressionType = ScriptDB.ExpressionType.ColumnDefault) Then
        LineOfCode.Code = String.Format("[dbo].[{0}](@prm_ID)", objThisColumn.Name)

      ElseIf objThisColumn.IsCalculated And objThisColumn.Table Is Me.AssociatedColumn.Table _
          And Not Me.ExpressionType = ScriptDB.ExpressionType.ColumnFilter And Not Me.ExpressionType = ScriptDB.ExpressionType.Mask Then

        If objThisColumn.Calculation Is Nothing Then
          objThisColumn.Calculation = objThisColumn.Table.GetObject(Type.Expression, objThisColumn.CalcID)
        End If

        iBackupType = objThisColumn.Calculation.ExpressionType
        objThisColumn.Calculation.ExpressionType = ScriptDB.ExpressionType.ReferencedColumn
        objThisColumn.Calculation.AssociatedColumn = objThisColumn

        objThisColumn.Calculation.StartOfPartNumbers = Me.StartOfPartNumbers + Declarations.Count
        objThisColumn.Calculation.GenerateCode()

        objThisColumn.Calculation.ExpressionType = iBackupType
        LineOfCode = AddCalculatedColumn(objThisColumn)
        objThisColumn.Tuning.IncrementSelectAsCalculation()

      Else

        'If is this column on the base table then add directly to the main execute statement,
        ' otherwise add it into child/parent statements array
        If objThisColumn.Table Is Me.AssociatedColumn.Table Then

          Select Case Component.BaseExpression.ExpressionType
            Case ScriptDB.ExpressionType.ColumnFilter
              sColumnName = String.Format("base.[{0}]", objThisColumn.Name)
              Me.IsComplex = True

            Case ScriptDB.ExpressionType.Mask
              sColumnName = String.Format("base.[{0}]", objThisColumn.Name)
              Me.IsComplex = True

              ' Needs base table added
              sFromCode = String.Format("FROM [dbo].[{0}] base", objThisColumn.Table.Name)
              If Not FromTables.Contains(sFromCode) Then
                FromTables.Add(sFromCode)
              End If

              ' Where clause
              sWhereCode = String.Format("base.[ID] = @prm_ID")
              If Not Wheres.Contains(sWhereCode) Then
                Wheres.Add(sWhereCode)
              End If

            Case Else
              sColumnName = String.Format("@prm_{0}", objThisColumn.Name)

          End Select

          LineOfCode.Code = String.Format("ISNULL({0},{1})", sColumnName, objThisColumn.SafeReturnType)

        Else

          sColumnFilter = String.Empty
          RequiresRecordID = True
          bIsSummaryColumn = False
          Me.IsComplex = True

          objRelation = Me.BaseTable.GetRelation(objThisColumn.Table.ID)

          If objRelation.RelationshipType = ScriptDB.RelationshipType.Parent Then

            If Me.ExpressionType = ScriptDB.ExpressionType.ColumnCalculation Then
              Me.AssociatedColumn.Table.DependsOnChildColumns.AddIfNew(objThisColumn)
            End If

            LineOfCode.Code = String.Format("ISNULL([{0}].[{1}],{2})", objThisColumn.Table.Name, objThisColumn.Name, objThisColumn.SafeReturnType)

            ' Add table join component
            sRelationCode = String.Format("LEFT JOIN [dbo].[{0}] ON [{0}].[ID] = base.[ID_{1}]" _
              , objRelation.Name, CInt(objRelation.ParentID))
            If Not Joins.Contains(sRelationCode) Then
              Joins.Add(sRelationCode)
            End If

            ' Needs base table added
            sFromCode = String.Format("FROM [dbo].[{0}] base", Me.AssociatedColumn.Table.Name)
            If Not FromTables.Contains(sFromCode) Then
              FromTables.Add(sFromCode)
            End If

            ' Where clause
            sWhereCode = "base.[ID] = @prm_ID"
            If Not Wheres.Contains(sWhereCode) Then
              Wheres.Add(sWhereCode)
            End If

            Me.ReferencesParent = True

          Else

            ' Add to dependency stack
            objThisColumn.Table.DependsOnParentColumns.AddIfNew(Me.AssociatedColumn)

            ' In a later release this can be tidied up to populate at load time
            [Component].ChildRowDetails.Order = objThisColumn.Table.GetObject(Enums.Type.TableOrder, [Component].ChildRowDetails.OrderID)
            [Component].ChildRowDetails.Filter = objThisColumn.Table.Objects.GetObject(Things.Type.Expression, [Component].ChildRowDetails.FilterID)
            [Component].ChildRowDetails.Relation = objRelation

            '       Debug.Assert(Not Me.AssociatedColumn.Name = "Personnel_Area")

            objOrderFilter = objThisColumn.Table.TableOrderFilter([Component].ChildRowDetails)
            objOrderFilter.IncludedColumns.AddIfNew(objThisColumn)

            ' Add calculation for this foreign column to the pre-requisits array 
            iPartNumber = Declarations.Count + Me.StartOfPartNumbers
            bIsSummaryColumn = ([Component].ChildRowDetails.RowSelection = ScriptDB.ColumnRowSelection.Total Or [Component].ChildRowDetails.RowSelection = ScriptDB.ColumnRowSelection.Count)

            ' Add to prereqistits arrays
            If bIsSummaryColumn Then
              Declarations.Add(String.Format("@part_{1} numeric(38,8)", [CodeCluster].Indentation, iPartNumber))
              sPartCode = String.Format("{0}SELECT @part_{1} = ISNULL(base.[{2}], 0)" & vbNewLine _
                  , [CodeCluster].Indentation, iPartNumber, objThisColumn.Name, objThisColumn.SafeReturnType)

            Else
              Declarations.Add(String.Format("@part_{1} {2}", [CodeCluster].Indentation, iPartNumber, objThisColumn.DataTypeSyntax))
              sPartCode = String.Format("{0}SELECT @part_{1} = ISNULL(base.[{2}],{3})" & vbNewLine _
                  , [CodeCluster].Indentation, iPartNumber, objThisColumn.Name, objThisColumn.SafeReturnType)

            End If

            sPartCode = sPartCode & String.Format("{0} FROM [dbo].[{1}](@prm_ID) base" _
                , [CodeCluster].Indentation, objOrderFilter.Name)

            StatementObjects.Add(objOrderFilter)
            PreStatements.Add(sPartCode)
            If bIsSummaryColumn Then
              LineOfCode.Code = String.Format("@part_{0}", iPartNumber)
            Else
              LineOfCode.Code = String.Format("ISNULL(@part_{0},{1})", iPartNumber, objThisColumn.SafeReturnType)
            End If
          End If
        End If
      End If

      ' Add this column (or reference to it) to the main execute statement
      [CodeCluster].Add(LineOfCode)

    End Sub

    Private Sub SQLCode_AddFunction(ByRef [Component] As Things.Component, ByRef [CodeCluster] As ScriptDB.LinesOfCode)

      Dim LineOfCode As ScriptDB.CodeElement

      Dim objCodeLibrary As Things.CodeLibrary
      Dim ChildCodeCluster As ScriptDB.LinesOfCode
      Dim WhereCodeCluster As ScriptDB.LinesOfCode
      Dim objSetting As Things.Setting
      Dim objIDComponent As Things.Component
      Dim objTriggeredUpdate As ScriptDB.TriggeredUpdate
      Dim sWhereClause As String = ""
      Dim iBackupType As ScriptDB.ExpressionType

      LineOfCode.CodeType = ScriptDB.ComponentTypes.Function
      objCodeLibrary = Globals.Functions.GetObject(Enums.Type.CodeLibrary, Component.FunctionID)
      LineOfCode.Code = objCodeLibrary.Code
      '      CodeCluster.NestedLevel = CodeCluster.NestedLevel + objCodeLibrary.CaseCount
      Me.CaseCount += objCodeLibrary.CaseCount

      ' Get parameters
      ChildCodeCluster = New ScriptDB.LinesOfCode
      ChildCodeCluster.CodeLevel = [CodeCluster].CodeLevel + 1
      ChildCodeCluster.ReturnType = objCodeLibrary.ReturnType

      ' Add module dependancy info for this function
      If objCodeLibrary.HasDependancies Then
        For Each objSetting In objCodeLibrary.Dependancies

          Select Case objSetting.SettingType

            Case SettingType.ModuleSetting
              objIDComponent = New Things.Component
              objIDComponent.SubType = ScriptDB.ComponentTypes.Relation
              objIDComponent.TableID = objSetting.Value
              [Component].Objects.Add(objIDComponent)
              Me.IsComplex = True

            Case SettingType.CodeItem
              objIDComponent = New Things.Component
              objIDComponent.SubType = ScriptDB.ComponentTypes.Value
              objIDComponent.ValueString = objSetting.Code
              objIDComponent.ValueType = ScriptDB.ComponentValueTypes.SystemVariable
              [Component].Objects.Add(objIDComponent)

            Case SettingType.UpdateParameter
              sWhereClause = objSetting.Code

          End Select

        Next
      End If

      ' Does this component need adding to the 'Get Field From Database' stack?
      If objCodeLibrary.IsGetFieldFromDB Then
        Globals.GetFieldsFromDB.Add([Component])
      End If

      ' Is this a unique value that needs evaluating?
      If objCodeLibrary.IsUniqueCode Then
        Globals.UniqueCodes.Add([Component])
      End If

      ' Is this expression reliant on the bank holiday table (I'm sure this can be tidyied up)
      If objCodeLibrary.DependsOnBankHoliday And Me.ExpressionType = ScriptDB.ExpressionType.ColumnCalculation Then
        objTriggeredUpdate = New ScriptDB.TriggeredUpdate
        objTriggeredUpdate.Column = Me.AssociatedColumn
        objTriggeredUpdate.ID = Me.AssociatedColumn.ID

        '' this is the messy bit!
        iBackupType = Me.ExpressionType
        Me.ExpressionType = ScriptDB.ExpressionType.TriggeredUpdate

        ' Get parameters
        WhereCodeCluster = New ScriptDB.LinesOfCode
        SQLCode_AddCodeLevel([Component].Objects, WhereCodeCluster)
        objTriggeredUpdate.Where = String.Format(sWhereClause, WhereCodeCluster.ToArray)

        Globals.OnBankHolidayUpdate.AddIfNew(objTriggeredUpdate)

        Me.ExpressionType = iBackupType

      End If


      SQLCode_AddCodeLevel([Component].Objects, ChildCodeCluster)
      LineOfCode.Code = String.Format(LineOfCode.Code, ChildCodeCluster.ToArray)
      RequiresRowNumber = RequiresRowNumber Or objCodeLibrary.RowNumberRequired
      RequiresOvernight = RequiresOvernight Or objCodeLibrary.OvernightOnly
      mbCalculatePostAudit = mbCalculatePostAudit Or objCodeLibrary.CalculatePostAudit
      Me.RequiresRecordID = RequiresRecordID Or objCodeLibrary.RecordIDRequired
      Me.Tuning.Rating += objCodeLibrary.Tuning.Rating
      objCodeLibrary.Tuning.Usage += 1


      ' For functions that return mixed type, make it type safe
      If objCodeLibrary.ReturnType = ScriptDB.ComponentValueTypes.Unknown And objCodeLibrary.MakeTypeSafe Then

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
            'LineOfCode.Code = 

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
      '      ChildCodeCluster.NestedLevel = CodeCluster.NestedLevel

      '  Debug.Assert(Me.AssociatedColumn.Name <> "Duration")
      ' Nesting is too deep - convert to part number
      If Me.CaseCount > 8 Then

        objExpression = New Things.Expression
        objExpression.ExpressionType = Me.ExpressionType
        objExpression.BaseTable = Me.BaseTable
        objExpression.AssociatedColumn = Me.AssociatedColumn
        objExpression.BaseExpression = Me.BaseExpression
        objExpression.ReturnType = Component.ReturnType
        objExpression.Objects = Component.Objects
        objExpression.StartOfPartNumbers = Declarations.Count + Me.StartOfPartNumbers
        objExpression.GenerateCode()

        Declarations.AddRange(objExpression.Declarations)
        PreStatements.AddRange(objExpression.PreStatements)
        Dependencies.MergeUnique(objExpression.Dependencies)

        iPartNumber = Declarations.Count + Me.StartOfPartNumbers
        Declarations.Add(String.Format("@part_{0} {1}", iPartNumber, objExpression.DataTypeSyntax))

        sPartCode = String.Format("{0}SELECT @part_{1} = {2}" & vbNewLine & _
            "{0}{3}" & vbNewLine & _
            "{0}{4}" & vbNewLine & _
            "{0}{5}" & vbNewLine _
            , [CodeCluster].Indentation, iPartNumber _
            , objExpression.UDF.SelectCode, objExpression.UDF.FromCode, objExpression.UDF.JoinCode, objExpression.UDF.WhereCode)
        PreStatements.Add(sPartCode)

        StatementObjects.Add(objExpression)
        LineOfCode.Code = String.Format("@part_{0}", iPartNumber)
        Me.IsComplex = True

      Else
        SQLCode_AddCodeLevel([Component].Objects, ChildCodeCluster)
        LineOfCode.Code = String.Format("{0}", ChildCodeCluster.Statement)

      End If

      [CodeCluster].Add(LineOfCode)

    End Sub

    Private Sub SQLCode_AddOperator(ByVal objComponent As Things.Component, ByRef [CodeCluster] As ScriptDB.LinesOfCode)

      Dim LineOfCode As ScriptDB.CodeElement
      Dim objCodeLibrary As Things.CodeLibrary

      LineOfCode.CodeType = ScriptDB.ComponentTypes.Operator

      ' Get the bits and bobs for this operator
      objCodeLibrary = Globals.Operators.GetObject(Enums.Type.CodeLibrary, objComponent.OperatorID)

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

    Private Sub AddToDependencies(ByRef Dependencies As Things.Collection)

      Dim objDependency As Things.Base
      Dim objColumn As Things.Column
      Dim objRelation As Things.Relation

      For Each objDependency In Dependencies

        If objDependency.Type = Enums.Type.Column Then
          objColumn = CType(objDependency, Things.Column)
          If Not Dependencies.Contains(objColumn) Then
            If objColumn.Table Is Me.BaseTable Then
              Dependencies.Add(objDependency)
            End If
          End If
        End If

        If objDependency.Type = Enums.Type.Relation Then
          objRelation = CType(objDependency, Things.Relation)
          If Not Dependencies.Contains(objRelation) Then
            Dependencies.Add(objDependency)
          End If
        End If

      Next

    End Sub

    ' Adds a calculated column to the pre-requists stack. This is for efficiency so the UDF is called a minimum of times.
    Private Function AddCalculatedColumn(ByRef ReferencedColumn As Things.Column) As ScriptDB.CodeElement

      Dim sCallingCode As ScriptDB.CodeElement
      Dim sVariableName As String

      'If Not ReferencedColumn.Calculation.IsComplex Then
      '  sCallingCode.Code = ReferencedColumn.Calculation.UDF.SelectCode
      'Else
      If StatementObjects.Contains(ReferencedColumn) Then
        sCallingCode.Code = String.Format("@part_{0}", StatementObjects.IndexOf(ReferencedColumn))
      Else
        sVariableName = StatementObjects.Count + Me.StartOfPartNumbers
        StatementObjects.Add(ReferencedColumn)
        Declarations.Add(String.Format("@part_{0} {1}", sVariableName, ReferencedColumn.DataTypeSyntax))
        PreStatements.Add(String.Format("SELECT @part_{0} = {1}", sVariableName, ReferencedColumn.Calculation.UDF.CallingCode))
        sCallingCode.Code = String.Format("@part_{0}", sVariableName)
      End If

      Me.CaseCount += ReferencedColumn.Calculation.CaseCount
      Me.BaseExpression.IsComplex = True
      Me.Tuning.Rating += ReferencedColumn.Calculation.Tuning.Rating

      '      End If

      Me.RequiresRecordID = RequiresRecordID Or ReferencedColumn.Calculation.RequiresRecordID
      Me.RequiresRowNumber = RequiresRowNumber Or ReferencedColumn.Calculation.RequiresRowNumber
      Me.RequiresOvernight = RequiresOvernight Or ReferencedColumn.Calculation.RequiresOvernight
      Me.ReferencesParent = Me.ReferencesParent Or ReferencedColumn.Calculation.ReferencesParent
      Dependencies.MergeUnique(ReferencedColumn.Calculation.Dependencies)

      Return sCallingCode

    End Function

    'Private Function AddChildColumn(ByRef ChildView As Things.TableOrderFilter, ByRef ReferencedColumn As Things.Column) As ScriptDB.CodeElement

    '  Dim sCallingCode As ScriptDB.CodeElement
    '  Dim sVariableName As String
    '  Dim objChildView As Things.TableOrderFilter

    '  If StatementObjects.Contains(ChildView) Then
    '    sCallingCode.Code = String.Format("@part_{0}", StatementObjects.IndexOf(ReferencedColumn) + 1)

    '    ' append to select statement


    '  Else

    '    StatementObjects.Add(ChildView)
    '    sVariableName = StatementObjects.Count

    '    ' What type/line number are we dealing with?
    '    Select Case ChildView.RowDetails.RowSelection

    '      Case ScriptDB.ColumnRowSelection.First, ScriptDB.ColumnRowSelection.Last, ScriptDB.ColumnRowSelection.Specific
    '        Declarations.Add(String.Format("@part_{1} {2};", sVariableName, ReferencedColumn.DataTypeSyntax))
    '        sCallingCode.Code = String.Format("@part_{0} = base.[{1}]", sVariableName, ReferencedColumn.Name)

    '      Case ScriptDB.ColumnRowSelection.Total
    '        Declarations.Add(String.Format("@part_{0} numeric(38,8);", sVariableName))
    '        sCallingCode.Code = String.Format("@part_{0} = SUM(base.[{1}])", sVariableName, ReferencedColumn.Name)

    '      Case ScriptDB.ColumnRowSelection.Count
    '        Declarations.Add(String.Format("@part_{0} numeric(38,8);", sVariableName))
    '        sCallingCode.Code = String.Format("@part_{0} = COUNT(base.[{1}])", sVariableName, ReferencedColumn.Name)

    '    End Select

    '    PreStatements.Add(String.Format("SELECT @part_{0} = {1}", sVariableName, ReferencedColumn.Calculation.UDF.CallingCode))



    '    sPartCode = sPartCode & String.Format("{0} FROM [dbo].[{1}](@prm_ID) base" _
    '        , [CodeCluster].Indentation, objOrderFilter.Name)


    '    PreStatements.Add(sPartCode)
    '    LineOfCode.Code = String.Format("ISNULL(@part_{0},{1})", iPartNumber, objThisColumn.SafeReturnType)

    '  End If

    '  Return sCallingCode

    'End Function

    Private Function ResultDataType(ByVal ColumnType As ScriptDB.ColumnTypes) As String

      Dim sSQLType As String = String.Empty

      Select Case CInt(ColumnType)
        Case ScriptDB.ColumnTypes.Text
          sSQLType = "varchar(MAX)"

        Case ScriptDB.ColumnTypes.Integer
          sSQLType = "integer"

        Case ScriptDB.ColumnTypes.Numeric
          sSQLType = "numeric(38,8)"

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

    End Function

    Private Function ResultWrapper(ByRef Statement As String) As String

      Dim sWrapped As String = String.Empty
      Dim sSize As String = String.Empty

      If Globals.Options.OverflowSafety Then

        Select Case CInt(Me.AssociatedColumn.DataType)
          Case ScriptDB.ColumnTypes.WorkingPattern
            sWrapped = Statement

          Case ScriptDB.ColumnTypes.Text, ScriptDB.ColumnTypes.Link
            If Me.AssociatedColumn.Multiline Then
              sWrapped = Statement
            Else
              sWrapped = String.Format("CASE WHEN LEN(ISNULL({0}, '')) > {1} THEN '' ELSE {0} END", Statement, Me.AssociatedColumn.Size)
            End If

          Case ScriptDB.ColumnTypes.Integer, ScriptDB.ColumnTypes.Numeric
            If Me.AssociatedColumn.Decimals > 0 Then
              sSize = String.Format("{0}.{1}", New String("9", Me.AssociatedColumn.Size - Me.AssociatedColumn.Decimals), New String("9", Me.AssociatedColumn.Decimals))
            Else
              sSize = New String("9", Me.AssociatedColumn.Size)
            End If
            sWrapped = String.Format("CASE WHEN ISNULL({0}, 0) > {1} THEN 0 ELSE {0} END", Statement, sSize)

          Case ScriptDB.ColumnTypes.Date, ScriptDB.ColumnTypes.Logic
            sWrapped = Statement

        End Select

      Else
        sWrapped = Statement
      End If

      Return sWrapped
    End Function

    Private Function ResultDefinition() As String

      Dim sSQLType As String = String.Empty

      Select Case CInt(Me.AssociatedColumn.DataType)
        Case ScriptDB.ColumnTypes.Text
          sSQLType = "varchar(MAX)"

        Case ScriptDB.ColumnTypes.Integer
          sSQLType = String.Format("integer")

        Case ScriptDB.ColumnTypes.Numeric
          sSQLType = String.Format("numeric(38,8)")

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
    End Function

  End Class
End Namespace

