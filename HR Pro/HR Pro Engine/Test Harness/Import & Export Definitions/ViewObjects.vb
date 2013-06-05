Option Explicit On
'Option Strict On

Imports Infragistics

Public Class ViewObjects

  Public Property Things As HRProEngine.Things.Collection
  Public Property ExportObjects As New HRProEngine.Things.Collection

  Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    'grdObjectSelection.sel

    '  Dim objFile As System.IO.File
    ' Dim objObject As ScriptDB.Things.iSystemObject
    '   Dim objControl As HCMObjectMapping

    ' Read the import file


    ' Bind to things class


    ' Get the user to create the mappings


    ' Generate script


    ' Execute script ?



    'ScriptDB.StructurePort.Initialise()
    'ScriptDB.StructurePort.CreateStatements(objProgress)

    'For Each objObject In ScriptDB.StructurePort.Dependancies
    '  objControl = New HCMObjectMapping
    '  objControl.FromObject = objObject
    '  objControl.Location = New System.Drawing.Point(30, (miMappingsRequired * 25))
    '  objControl.TabIndex = miMappingsRequired
    '  pnlMappings.Controls.Add(objControl)

    '  miMappingsRequired = miMappingsRequired + 1

    '  '      AddMapping(objObject.ID, objObject.Type, objObject.Name)
    'Next

    'objFile.WriteAllLines(txtUpdateScript.Text, ScriptDB.StructurePort.GetStatements)





  End Sub

  Private Sub ViewObjects_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

    Dim bnd As Infragistics.Win.UltraWinGrid.UltraGridBand
    Dim dc As Infragistics.Win.UltraWinGrid.UltraGridColumn


    grdThings.DataSource = Things

    With grdThings.DisplayLayout
      .ViewStyleBand = Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
      .GroupByBox.Hidden = False
      .Override.GroupByRowDescriptionMask = "[type]"  '"[value] ([count] [count,items,item,items])"

      For Each bnd In .Bands
        With bnd

          ' bnd.ColHeadersVisible = False

          '    If .Index = 0 Then
          '      mdtTopLevel.Columns.Add("database", Type.GetType("System.String"), 0)
          '      mdtTopLevel.Columns("database").SetOrdinal(0)
          '      .Columns("database").ValueList = FixedColumnValuelist(gSQL.DatabaseName, "database16")
          '      '.SortedColumns.Add(.Columns("Database"), False, True)
          '    End If

          '    '          ChangeSortOrder(bnd)

          For Each dc In bnd.Columns
            '      Select Case dc.Header.Caption
            '        Case "type"
            '          dc.ValueList = TypeValueList()  '(IIf(.Index = 0, "Tables", "Child Tables"))
            '          'Case "subtype"
            '          '  dc.ValueList = GetValueList("objecttype")  '(IIf(.Index = 0, "Tables", "Child Tables"))
            '        Case "name"
            '        Case Else
            '          dc.Hidden = True
            '      End Select
          Next

        End With
      Next


    'AsrGrid1.DisplayLayout.Bands(0).Indentation = 10
    'AsrGrid1.DisplayLayout.Bands(1).Indentation = 0
    'AsrGrid1.DisplayLayout.Bands(1).IndentationGroupByRow = 0
    End With

    'AsrGrid1.Rows.ExpandAll(True)
    'AsrGrid1.Rows.CollapseAll(True)
    grdThings.Rows(0).Expanded = True


  End Sub

  Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
    Me.Close()
  End Sub

  Private Sub butExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butExport.Click

    Dim objRow As Infragistics.Win.UltraWinGrid.UltraGridRow
    Dim objObject As HRProEngine.Things.Base
    Dim strText As String

    ExportObjects.Clear()

    For Each objRow In grdThings.Selected.Rows
      ExportObjects.Add(objRow.ListObject)
      '      objObject = objRow.ListObject
    Next

    For Each objObject In ExportObjects

      objObject.ToXML("c:\dev\xml.xml")
      'MsgBox(strText)
    Next

  End Sub

End Class