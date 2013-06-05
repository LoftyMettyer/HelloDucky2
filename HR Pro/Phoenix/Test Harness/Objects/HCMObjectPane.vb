Imports Infragistics.Win.UltraWinGrid
Public Class HCMObjectPane

  Public SelectedObjects As Phoenix.Things.Collection

  Public Sub Attach(ByRef Things As Phoenix.Things.Collection)

    ' Some default formatting.
    UltraGrid1.DisplayLayout.MaxBandDepth = 3
    UltraGrid1.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.AutoFree
    UltraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect

    ' Bind the data and structure
    UltraGrid1.DataSource = Things

    FormatDisplay()

  End Sub

  Public Sub ShowColumn(ByVal Key As String, ByVal Show As Boolean)

    For Each objBand As Infragistics.Win.UltraWinGrid.UltraGridBand In UltraGrid1.DisplayLayout.Bands
      If objBand.Columns.Exists(Key) Then
        objBand.Columns(Key).Hidden = Not Show
      End If
    Next

  End Sub

  Private Sub FormatDisplay()

    Dim bDisplayHeader As Boolean

    ' Formatting
    UltraGrid1.Font = New Font("Calibri", 10)

    ' Group by the second band, so that columns/views/expression etc appear in the grid group by.
    'AsrGrid1.DisplayLayout.Override.ExpansionIndicator = UltraWinGrid.ShowExpansionIndicator.CheckOnDisplay
    UltraGrid1.DisplayLayout.Bands(0).SortedColumns.Add("Name", False, False)   ' Default sort order on name alphabetic
    UltraGrid1.DisplayLayout.Bands(1).SortedColumns.Add("Type", False, True)    ' Default sort on child nodes by type
    'AsrGrid1.DisplayLayout.Override. pick up on multselect property

    ' Hide the system columns
    ShowColumn("ID", False)
    ShowColumn("ObjectState", False)
    ShowColumn("Parent", False)
    ShowColumn("Type", False)
    ShowColumn("Description", False)
    ShowColumn("Objects", False)

    ' Apply properties to each of the bands
    For Each objBand As Infragistics.Win.UltraWinGrid.UltraGridBand In UltraGrid1.DisplayLayout.Bands
      objBand.ColHeadersVisible = bDisplayHeader
      objBand.GroupHeadersVisible = False
      objBand.HeaderVisible = False
    Next

  End Sub

  Private Sub UltraGrid1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid1.Click

    Dim objSelectedRow As Infragistics.Win.UltraWinGrid.UltraGridRow
    Dim bToggleSelected As Boolean

    For Each objSelectedRow In UltraGrid1.Selected.Rows
      If objSelectedRow.IsGroupByRow Then
        'If Not objSelectedRow.HasChild Then
        '  Me.SelectedThings.Add(objSelectedRow.ParentRow.ListObject)
        '  RaiseEvent ItemClick(objSelectedRow.ParentRow.ListObject, CType(objSelectedRow.GetChild(ChildRow.First).ListObject, Things.iSystemObject).Type)
        'End If
      Else
        bToggleSelected = CType(objSelectedRow.ListObject, Phoenix.Things.Base).IsSelected
        CType(objSelectedRow.ListObject, Phoenix.Things.Base).IsSelected = True 'Not bToggleSelected

        'Me.SelectedThings.Add(objSelectedRow.ListObject)
        'RaiseEvent ItemClick(objSelectedRow.ListObject, CType(objSelectedRow.ListObject, Things.iSystemObject).Type)
      End If
    Next


  End Sub


  Private Sub UltraGrid1_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid1.InitializeLayout

    e.Layout.Override.HeaderPlacement = HeaderPlacement.FixedOnTop
    e.Layout.Override.ExpansionIndicator = ShowExpansionIndicator.CheckOnDisplay
    e.Layout.GroupByBox.Hidden = True

  End Sub
End Class
