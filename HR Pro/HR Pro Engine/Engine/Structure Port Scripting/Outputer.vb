'Namespace StructurePort

'  <HideModuleName()>
'  Public Module [Export]

'    Public Dependancies As Things.Collection
'    Private maryStatements As ArrayList

'    Public Sub Initialise()

'      Dependancies = New Things.Collection

'    End Sub

'    ' Returns all the statements as scripts
'    Public Function GetStatements() As String()

'      Dim objObject As Things.Base 'Things.iSystemObject
'      Dim objTable As Things.Table
'      Dim objColumn As Things.Column

'      ' Tack on the table prequisiticts
'      For Each objObject In Dependancies
'        If objObject.Type = Things.Type.Table Then
'          objTable = objObject
'          maryStatements.Insert(0, String.Format("DECLARE @TableID_{0} integer = {0} --{1}", CInt(objTable.ID), objTable.Name))
'        End If

'        If objObject.Type = Things.Type.Column Then
'          objColumn = objObject
'          maryStatements.Insert(0, String.Format("DECLARE @ColumnID_{0} integer = {0} --{1}", CInt(objColumn.ID), objColumn.Name))
'        End If

'      Next

'      Return maryStatements.ToArray(GetType(String))
'    End Function

'    ' Generates the system objects
'    Public Sub CreateStatements(ByRef ProgressInfo As HCMProgressBar)

'      Dim objTable As Things.Table
'      '   Dim objColumn As Things.Column
'      Dim objScreen As Things.Screen
'      '  Dim objWorkflow As Things.Workflow
'      '  Dim objWorkflowElement As Things.WorkflowElement
'      Dim sOutputLine As String

'      maryStatements = New ArrayList

'      ' Tables
'      For Each objTable In Globals.Things

'        '        If objTable.IsSelected Then
'        sOutputLine = String.Format("EXEC spadmin_savedefinition_table (@TableID_{0}, {1}, '{2}', {3}, {4}, {5}, {6}, {7}, {8});" _
'                          , CInt(objTable.ID), objTable.TableType, objTable.Name _
'                          , CInt(objTable.DefaultOrderID), CInt(objTable.RecordDescription.ID), CInt(objTable.DefaultEmailID) _
'                          , CInt(objTable.ManualSummaryColumnBreaks), CInt(objTable.AuditInsert), CInt(objTable.AuditDelete))
'        maryStatements.Add(sOutputLine)

'        '' Columns
'        'For Each objColumn In objTable.Objects(Things.Type.Column)
'        '  sOutputLine = String.Format("EXEC spadmin_savedefinition_column (@ColumnID_{0}, @TableID_{1}, '{2}', {3}, {4}, {5});" _
'        '                  , CInt(objColumn.ID), CInt(objTable.ID), objColumn.Name _
'        '                  , CInt(objColumn.DataType), objColumn.Size, objColumn.Decimals)
'        '  Dependancies.Add(objColumn)
'        '  maryStatements.Add(sOutputLine)
'        'Next

'        ' Expressions

'        ' Screens
'        For Each objScreen In objTable.Objects(Things.Type.Screen)
'          If objScreen.IsSelected Then
'            sOutputLine = String.Format("EXEC spadmin_savedefinition_screen ({0}, @TableID_{1}, '{2}');" _
'                            , CInt(objScreen.ID), CInt(objScreen.Table.ID), objScreen.Name)
'            maryStatements.Add(sOutputLine)
'            Dependancies.Add(objTable)
'          End If
'        Next

'        ' Workflows
'        'For Each objWorkflow In objTable.Objects(Things.Type.Workflow)
'        '  sOutputLine = String.Format("EXEC spadmin_savedefinition_workflow ({0}, '{1}', '{2}', {3}, {4}, @TableID_{5}, '{6}');" _
'        '                      , CInt(objWorkflow.ID), objWorkflow.Name, Replace(objWorkflow.Description, "'", "''") _
'        '                      , objWorkflow.Enabled, objWorkflow.InitiationType, CInt(objWorkflow.BaseTableID), objWorkflow.QueryString)
'        '  maryStatements.Add(sOutputLine)

'        '  For Each objWorkflowElement In objWorkflow.Objects(Things.Type.WorkflowElement)
'        '    'sOutputLine = String.Format("EXEC spadmin_saveworkflowelement ({0}, {1})" & _
'        '    '                  "({0},'{1}','{2}',{3},{4},@TableID_{5},'{6}');" _
'        '    '                  , CInt(objWorkflow.ID), objWorkflow.Name, Replace(objWorkflow.Description, "'", "''") _
'        '    '                  , objWorkflow.Enabled, objWorkflow.InitiationType, CInt(objWorkflow.BaseTableID), objWorkflow.QueryString)



'        '  Next


'        'Next
'        '        End If
'      Next

'    End Sub

'    'Public Sub ExportToXML(ByRef ProgressInfo As HCMProgressBar)

'    '  Dim objTable As Things.Table

'    '  maryStatements = New ArrayList

'    '  ' Tables
'    '  For Each objTable In Globals.Things
'    '    objTable()

'    '  Next

'    'End Sub


'  End Module


'End Namespace
