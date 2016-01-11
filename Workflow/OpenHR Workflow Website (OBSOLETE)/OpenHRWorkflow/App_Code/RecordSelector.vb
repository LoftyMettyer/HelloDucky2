Imports System
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web

Imports Microsoft.VisualBasic
Imports Utilities
Imports System.Data

Public Class RecordSelector
    Inherits GridView

    Protected WithEvents customHeader As GridViewRow
    Protected WithEvents customPager As GridViewRow
    Protected WithEvents grdGrid As RecordSelector ' GridView
    Protected WithEvents pagerTable As Table
    Protected WithEvents pagerRow As TableRow
    Protected WithEvents ddl As TextBox ' DropDownList
    Protected WithEvents dataTable As DataTable

    Private Const MAXDROPDOWNROWS As Int16 = 6
    Private m_iVisibleColumnCount As Integer
    Dim iColWidth As Integer = 100  ' default, minimum column width for all grid columns

  Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)

    Dim iIDColumnIndex As Int16
    Dim sDivStyle As String = ""
    Dim iColCount As Integer

    'MyBase.Render(writer)

    If IsLookup Then sDivStyle = ";display:none;"

    ' Custom Header row first.
    'Dim customHeader As GridViewRow = Me.HeaderRow
    customHeader = Me.HeaderRow

    If MyBase.HeaderRow IsNot Nothing Then

      m_iVisibleColumnCount = MyBase.HeaderRow.Cells.Count
      customHeader.ApplyStyle(Me.HeaderStyle)

      Dim sColumnCaption As String = ""

      For iColCount = 0 To Me.HeaderRow.Cells.Count - 1
        sColumnCaption = UCase(customHeader.Cells(iColCount).Text)

        If (Not IsLookup And sColumnCaption = "ID") Or (IsLookup And sColumnCaption.StartsWith("ASRSYS")) Then
          iIDColumnIndex = CShort(iColCount)
        Else
          iIDColumnIndex = -1
        End If

        If (Not IsLookup And (sColumnCaption = "ID" Or (Left(sColumnCaption, 3) = "ID_" And Val(Mid(sColumnCaption, 4)) > 0))) Or _
            (IsLookup And sColumnCaption.StartsWith("ASRSYS")) Then
          m_iVisibleColumnCount = m_iVisibleColumnCount - 1
          customHeader.Cells(iColCount).Style.Add("display", "none")
        Else
          customHeader.Cells(iColCount).Text = Replace(customHeader.Cells(iColCount).Text, "_", " ")

          ' Add each of the gridview's column headers to the header table
          customHeader.Cells(iColCount).ID = Me.ID.ToString & "header" & CStr(iColCount + 1)
          customHeader.Cells(iColCount).Style.Add("width", Unit.Pixel(iColWidth).ToString)
          customHeader.Cells(iColCount).Style.Add("overflow", "hidden")
          customHeader.Cells(iColCount).Style.Add("white-space", "nowrap")
          customHeader.Cells(iColCount).Style.Add("text-overflow", "ellipsis")
          If iColCount <> 0 Then  ' horrible. Should work out how to hide left border if it overlaps the row border.
            customHeader.Cells(iColCount).Style.Add("border-left", "1px solid gray")
          End If
          customHeader.Cells(iColCount).Style.Add("border-bottom", "1px solid gray")
          customHeader.Cells(iColCount).Style.Add("padding-top", "0px")
          customHeader.Cells(iColCount).Style.Add("padding-bottom", "0px")

          Dim ctlTextPanel As New Panel
          With ctlTextPanel
            .ID = "textbox"
            .Style.Add("float", "left")
            .Style.Add("width", "100%")
            .Style.Add("height", "100%")
            .Style.Add("overflow", "hidden")
            .Style.Add("text-overflow", "ellipsis")
            .Style.Add("white-space", "nowrap")
            If MyBase.Rows.Count > 1 Then
              If IsLookup Then
                .Attributes("onclick") = ("$get('txtActiveDDE').value='" & grdGrid.ID.Replace("Grid", "dde") & "';try{setPostbackMode(3);}catch(e){};if(event.ctrlKey){__doPostBack('" & MyBase.UniqueID & "','Sort$" & customHeader.Cells(iColCount).Text & "+');}else{__doPostBack('" & MyBase.UniqueID & "','Sort$") & customHeader.Cells(iColCount).Text & "');}"
              Else
                .Attributes("onclick") = ("try{setPostbackMode(3);}catch(e){};if(event.ctrlKey){__doPostBack('" & MyBase.UniqueID & "','Sort$" & customHeader.Cells(iColCount).Text & "+');}else{__doPostBack('" & MyBase.UniqueID & "','Sort$") & customHeader.Cells(iColCount).Text & "');}"
              End If
            End If
          End With

          ' customHeader.Cells(iColCount).Controls.Add(New LiteralControl(Replace(customHeader.Cells(iColCount).Text, "_", " ")))
          ctlTextPanel.Controls.Add(New LiteralControl(Replace(customHeader.Cells(iColCount).Text, "_", " ")))

          customHeader.Cells(iColCount).Controls.Add(ctlTextPanel)

          Dim ctlSortPanel As New Panel
          With ctlSortPanel
            .ID = "sortpanel"
            .Style.Add("position", "absolute")
            .Style.Add("right", "0px")
            .Style.Add("float", "right")
            ' .Style.Add("width", "16px")
            .Style.Add("height", "100%")
          End With


          Dim ctlSortImage As New Image
          With ctlSortImage
            .ID = "sortimage"

            Select Case ColumnSortDirectionCode(customHeader.Cells(iColCount).Text)
              Case 0
                .Visible = False
              Case 1
                .ImageUrl = "Images/sort-asc.gif"
                .Visible = True
              Case 2
                .ImageUrl = "Images/sort-desc.gif"
                .Visible = True
            End Select

            ' Apply onclick event for sorting to the sort-direction arrow too. Fault HRPRO-1832
            If MyBase.Rows.Count > 1 Then
              If IsLookup Then
                .Attributes("onclick") = ("$get('txtActiveDDE').value='" & grdGrid.ID.Replace("Grid", "dde") & "';try{setPostbackMode(3);}catch(e){};if(event.ctrlKey){__doPostBack('" & MyBase.UniqueID & "','Sort$" & customHeader.Cells(iColCount).Text & "+');}else{__doPostBack('" & MyBase.UniqueID & "','Sort$") & customHeader.Cells(iColCount).Text & "');}"
              Else
                .Attributes("onclick") = ("try{setPostbackMode(3);}catch(e){};if(event.ctrlKey){__doPostBack('" & MyBase.UniqueID & "','Sort$" & customHeader.Cells(iColCount).Text & "+');}else{__doPostBack('" & MyBase.UniqueID & "','Sort$") & customHeader.Cells(iColCount).Text & "');}"
              End If
            End If

          End With

          ctlSortPanel.Controls.Add(ctlSortImage)
          customHeader.Cells(iColCount).Controls.Add(ctlSortPanel)

        End If
      Next
    Else
      ' No records to display.....
      Return
    End If


    ' Div to contain all items
    writer.Write("<div ID='" & Me.ID.ToString.Replace("Grid", "") & "' " & _
        If(IsLookup, "style='", "style='position:absolute;") & _
        " width:" & CalculateWidth() & ";height:" & CalculateHeight() & ";" & _
        If(IsLookup, "", "top:" & Me.Style.Item("TOP").ToString & ";") & _
        If(IsLookup, "", "left:" & Me.Style.Item("LEFT").ToString & ";") & _
        sDivStyle & _
        ";overflow:hidden;border-color:black;border-style:solid;border-width:1px;background-color:" & _
        General.GetHtmlColour(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(customHeader.BackColor.A, customHeader.BackColor.R, customHeader.BackColor.G, customHeader.BackColor.B)).ToString) & _
        "'>")

    'render header row 
    writer.Write("<table ID='" & Me.ID.ToString.Replace("Grid", "") & "Header'" & _
                " style='position:absolute;top:0px;left:0px;width:" & CalculateGridWidth() & ";height:" & CalculateHeaderHeight() & _
                ";table-layout:fixed;border:0'" & _
                " cellspacing='" & Me.CellSpacing.ToString() & "'" & _
                If((MyBase.HeaderRow IsNot Nothing And Me.IsEmpty = False), " class='resizable'", "") & _
                ">")

    If MyBase.HeaderRow IsNot Nothing Then
      customHeader.RenderControl(writer)
      'make invisible default header  row
      Me.HeaderRow.Visible = False
    End If


    writer.Write("</table>")

    ' Create a div to contain the Gridview table, not the pager controls or the header columns.
    ' (This is the one with scrollbars) - too small for lookups
    writer.Write("<div id='" & ClientID.Replace("Grid", "") & "gridcontainer'  style='position:absolute;top:" & CalculateHeaderHeight() & _
                 ";bottom:" & CalculatePagerHeight() & ";left:0px;overflow-x:scroll;overflow-y:scroll;" & _
                 "width:" & CalculateWidth() & ";" & "background-color:" & _
                 General.GetHtmlColour(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Me.BackColor.A, Me.BackColor.R, Me.BackColor.G, Me.BackColor.B)).ToString) & _
                 ";' onscroll=scrollHeader('" & ClientID.Replace("Grid", "gridcontainer") & "')>")


    ' Need to hide the default pager row BEFORE rendering myBase.
    ' Dim customPager As GridViewRow = Me.BottomPagerRow
    customPager = Me.BottomPagerRow

    If Me.BottomPagerRow IsNot Nothing Then
      Me.BottomPagerRow.Visible = False
    End If

    ' now render data rows        
    MyBase.Attributes.CssStyle("overflow") = "auto"
    MyBase.Style("overflow") = "auto"
    MyBase.Style.Add("table-layout", "fixed")
    MyBase.Style.Remove("position")


    MyBase.Attributes.CssStyle("WIDTH") = CalculateGridWidth()

    If IsLookup Then
      MyBase.Style.Remove("top")
      MyBase.Attributes.CssStyle.Remove("LEFT")
      MyBase.Attributes.CssStyle.Remove("TOP")
    End If

    MyBase.Style.Add("border-right", "none")
    MyBase.Style.Add("border-left", "none")
    MyBase.Style.Add("border-top", "solid 1px black")
    MyBase.Style.Add("border-bottom", "solid 1px black")
    AdjustWidthForScrollbar()   ' reduce width if vertical scrollbar

    MyBase.Render(writer)
    writer.Write("</div>")
    ' Render the custom pager row now.
    If customPager IsNot Nothing Then
      writer.Write("<div  border='0' cellspacing='" & Me.CellSpacing.ToString() & "' cellpadding='" & _
                      Me.CellPadding.ToString() & "' style='width:100%;position:absolute;right:0px;bottom:0px;height: " & CalculatePagerHeight() & ";'>")

      Dim iGridwidth As Integer
      iGridwidth = CInt(CalculateGridWidth.Replace("px", "").Replace("%", ""))

      customPager.ApplyStyle(Me.PagerStyle)
      customPager.Visible = True
      customPager.RenderControl(writer)
      writer.Write("</div>")
    End If

    writer.Write("</div>")


  End Sub

  Protected Overrides Sub InitializePager(ByVal row As System.Web.UI.WebControls.GridViewRow, ByVal columnSpan As Integer, ByVal pagedDataSource As System.Web.UI.WebControls.PagedDataSource)
    ' MyBase.InitializePager(row, columnSpan, pagedDataSource)
    InitCustomPager(row, columnSpan, pagedDataSource)
  End Sub



  Private Function CalculateWidth() As String
    Dim strWidth As String = "auto"

    If Not IsLookup Then
      If Not Me.Width.IsEmpty Then
        strWidth = String.Format("{0}{1}", Me.Width.Value, (If((Me.Width.Type = UnitType.Percentage), "%", "px")))
      End If
    Else
      Dim iGridWidth As Integer = m_iVisibleColumnCount * (iColWidth + 2) ' 2 = padding
      iGridWidth = CInt(If(iGridWidth < 0, 1, iGridWidth))
      iGridWidth = CInt(If(iGridWidth < Me.Width.Value, Me.Width.Value, iGridWidth))

      ' do rows exceed height?
      'If MyBase.Rows.Count > MAXDROPDOWNROWS Then
      ' Add scrollbar width
      iGridWidth += 25
      'End If

      If iGridWidth < 250 Then iGridWidth = 250 ' minimum width to ensure paging controls fit.

      strWidth = String.Format("{0}{1}", iGridWidth, (If((Me.Width.Type = UnitType.Percentage), "%", "px")))
    End If

    Return strWidth

  End Function
  Private Function CalculateHeight() As String

    Dim strHeight As String = "auto"
    Dim iHeight As Integer = ControlHeight

    If Not IsLookup Then
      If iHeight > 0 Then
        strHeight = String.Format("{0}{1}", iHeight, "px")
      End If
    Else
      ' Set the size of the grid as per old DropDown setting...
      ' remember that the height for lookups will only be c.21pixels...
      Dim iRowHeight As Integer = iHeight - 6
      iRowHeight = CInt(If(iRowHeight < 21, 21, iRowHeight))
      Dim iDropHeight As Integer = (iRowHeight * CInt(If(MyBase.Rows.Count > MAXDROPDOWNROWS, MAXDROPDOWNROWS, MyBase.Rows.Count))) + 1

      iDropHeight += iRowHeight  ' add row for headers
      iDropHeight += 30   ' Pager height - now it's always displayed.

      strHeight = String.Format("{0}{1}", iDropHeight, "px")
    End If

    Return strHeight

  End Function

  Private Sub AdjustWidthForScrollbar()

    'Dim strHeight As String = "auto"

    ''Adjust available width for the vertical scrollbar.
    'iGapBetweenBorderAndText = (CInt(NullSafeSingle(dr("FontSize")) + 6) \ 4)
    'iEffectiveRowHeight = CInt(NullSafeSingle(dr("FontSize"))) _
    ' + 1 _
    ' + (2 * iGapBetweenBorderAndText)

    'iTempHeight = NullSafeInteger(ctlForm_GridContainer.Height.Value)
    'iTempHeight = CInt(If(iTempHeight < 0, 1, iTempHeight))

    'MyBase.Style.Remove("width")
    'MyBase.Style.Add("width", "183px")

  End Sub

  Private Function CalculateGridWidth() As String
    ' grid width is me.width - vertical scroll bar.
    Dim strWidth As String = "auto"
    Dim iScrollBarWidth As Integer = 0

    If MyBase.Rows.Count > MAXDROPDOWNROWS Then
      iScrollBarWidth = 17
    End If

    If Not IsLookup Then
      If Not Me.Width.IsEmpty Then
        strWidth = String.Format("{0}{1}", Me.Width.Value - iScrollBarWidth, (If((Me.Width.Type = UnitType.Percentage), "%", "px")))
      End If
    Else
      Dim iGridWidth As Integer = m_iVisibleColumnCount * (iColWidth + 2) ' 2 = padding
      iGridWidth = CInt(If(iGridWidth < 0, 1, iGridWidth))
      iGridWidth = CInt(If(iGridWidth < Me.Width.Value, Me.Width.Value, iGridWidth))

      ' do rows exceed height?
      'If MyBase.Rows.Count > MAXDROPDOWNROWS Then
      ' Add scrollbar width
      iGridWidth += 25
      'End If

      If iGridWidth < 250 Then iGridWidth = 250 ' minimum width to ensure paging controls fit.

      strWidth = String.Format("{0}{1}", iGridWidth - iScrollBarWidth, (If((Me.Width.Type = UnitType.Percentage), "%", "px")))
    End If

    Return strWidth


  End Function

  Private Function CalculateHeaderHeight() As String

    Dim strHeaderHeight As String = "auto"
    Dim iHeaderHeight As Integer = 0
    Dim iGridHeight As Integer = ControlHeight

    If NullSafeBoolean(Me.ColumnHeaders) And (NullSafeInteger(Me.HeadLines) > 0) Then
      Dim iGridTopPadding As Integer = CInt(NullSafeSingle(Me.HeadFontSize) / 8)

      iHeaderHeight = CInt(((NullSafeSingle(Me.HeadFontSize) + iGridTopPadding) * NullSafeInteger(Me.HeadLines) * 2) _
       - 2 _
       - (NullSafeSingle(Me.HeadFontSize) * (NullSafeInteger(Me.HeadLines) + 1) * (iGridTopPadding - 1) / 4))

      If iHeaderHeight > NullSafeInteger(iGridHeight) Then
        iHeaderHeight = NullSafeInteger(iGridHeight)
      End If
    End If

    strHeaderHeight = String.Format("{0}{1}", iHeaderHeight, "px")

    Return strHeaderHeight

  End Function

  Private Function CalculatePagerHeight() As String

    'Dim strPagerHeight As String = "auto"
    'Dim iPagerHeight As Integer = 0
    'Dim iGridHeight As Integer = ControlHeight

    'If Me.PageCount > 0 Then
    'Dim iGridTopPadding As Integer = CInt(NullSafeSingle(Me.HeadFontSize) / 8)

    'iPagerHeight = CInt((((NullSafeSingle(Me.HeadFontSize) + iGridTopPadding) * 2) - 2) * 1.5)

    'If iPagerHeight > NullSafeInteger(iGridHeight) Then
    '    iPagerHeight = NullSafeInteger(iGridHeight)
    'End If

    ' NPG20120110 Fault HRPRO-1831
    ' Hide pager bar if control is too narrow to display navigation buttons, or
    ' if it's empty.

    If (Me.Width.Value < 175 And Me.PageCount < 1) Or Me.IsEmpty Then
      Return "0px"
    Else
      Return "24px"
    End If

  End Function

  Public Property IsEmpty() As Boolean
    Get
      If Not ViewState("IsEmpty") Is Nothing Then
        Return DirectCast(ViewState("IsEmpty"), Boolean)
      Else
        Return True
      End If
    End Get
    Set(ByVal value As Boolean)
      ViewState("IsEmpty") = value
    End Set
  End Property

  Public Property ControlHeight() As Integer
    Get
      If Not ViewState("ControlHeight") Is Nothing Then
        Return DirectCast(ViewState("ControlHeight"), Integer)
      Else
        Return 1
      End If
    End Get
    Set(ByVal value As Integer)
      ViewState("ControlHeight") = value
    End Set
  End Property


  Public Property HeadLines() As Integer
    Get
      If Not ViewState("HeadLines") Is Nothing Then
        Return DirectCast(ViewState("HeadLines"), Integer)
      Else
        Return 1
      End If
    End Get
    Set(ByVal value As Integer)
      ViewState("HeadLines") = value
    End Set
  End Property



  Public Property HeadFontSize() As Single
    Get
      If Not ViewState("HeadFontSize") Is Nothing Then
        Return DirectCast(ViewState("HeadFontSize"), Single)
      Else
        Return 8
      End If
    End Get
    Set(ByVal value As Single)
      ViewState("HeadFontSize") = value
    End Set
  End Property

  Public Property ColumnHeaders() As Boolean
    Get
      If Not ViewState("ColumnHeaders") Is Nothing Then
        Return DirectCast(ViewState("ColumnHeaders"), Boolean)
      Else
        Return True
      End If
    End Get
    Set(ByVal value As Boolean)
      ViewState("ColumnHeaders") = value
    End Set
  End Property

  Public Property IsLookup() As Boolean
    Get
      If Not ViewState("IsLookup") Is Nothing Then
        Return DirectCast(ViewState("IsLookup"), Boolean)
      Else
        Return False
      End If
    End Get
    Set(ByVal value As Boolean)
      ViewState("IsLookup") = value
    End Set
  End Property

  Public Property filterSQL() As String
    Get
      If Not ViewState("filterSQL") Is Nothing Then
        Return DirectCast(ViewState("filterSQL"), String)
      Else
        Return ""
      End If
    End Get
    Set(ByVal value As String)
      ViewState("filterSQL") = value
    End Set
  End Property


  Public Property sortBy() As String
    Get
      If Not ViewState("sortBy") Is Nothing Then
        Return DirectCast(ViewState("sortBy"), String)
      Else
        Return ""
      End If
    End Get
    Set(ByVal value As String)
      ViewState("sortBy") = value
    End Set
  End Property

  Private Sub RecordSelector_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.DataBound

    grdGrid = CType(sender, GridView)

    Dim iEffectiveRowHeight As Integer = grdGrid.RowStyle.Height.Value
    Dim iRowCount As Integer = grdGrid.Rows.Count

      grdGrid.Height = iEffectiveRowHeight * iRowCount

  End Sub

  Private Function ColumnSortDirectionCode(ByVal sColName As String) As Integer
    ' me.sortby is a string of all currently set sort orders, 
    ' it'll look something like this:   [Surname] ASC, [Forenames] DESC
    ' This function looks for the passed in column name and returns the following:
    ' 0 = column not found, 1 = ASC, 2 = DESC

    Dim aColList() As String = Split(Me.sortBy, ",")

    For iCount As Integer = 0 To aColList.Length - 1
      If aColList(iCount).Contains("[" & sColName.Replace(" ", "_") & "]") Then

        Select Case Right(aColList(iCount), 4).Trim.ToUpper
          Case "ASC"
            Return 1
          Case "DESC"
            Return 2
          Case Else
            Return 0
        End Select
      End If
    Next

    Return 0

  End Function


  Private Sub RecordSelector_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PageIndexChanged

    ' Default behaviour of the gridview is to highlight the same row again when 
    ' the page is changed. We'll override that.

    grdGrid = CType(sender, GridView)

    Dim iCurrentIndex As Integer

    grdGrid.SelectedIndex = -1
    grdGrid.DataBind()

    ' For lookups, use page and index number (breaks with sorting). For Grids use record ID.
    If IsLookup Then
      For itemIndex As Integer = 0 To grdGrid.Rows.Count - 1
        iCurrentIndex = (grdGrid.PageIndex * grdGrid.PageSize) + itemIndex
        If Me.SelectedID(grdGrid) = iCurrentIndex Then
          grdGrid.SelectedIndex = itemIndex
          grdGrid.DataBind() ' Need to bind (again) to reset selection.
          Exit For
        End If
      Next
    Else
      ' get the grid's id column number.
      Dim iColCount As Integer
      Dim sColCaption As String
      Dim iIDColNum As Integer = -1
      For iColCount = 0 To grdGrid.HeaderRow.Cells.Count - 1
        sColCaption = grdGrid.HeaderRow.Cells(iColCount).Text
        If sColCaption.ToUpper = "ID" Then
          ' Here it is
          iIDColNum = iColCount
          Exit For
        End If
      Next
      If iIDColNum > 0 Then
        For itemIndex As Integer = 0 To grdGrid.Rows.Count - 1
          iCurrentIndex = NullSafeInteger(grdGrid.Rows(itemIndex).Cells(iIDColNum).Text)
          If Me.SelectedID(grdGrid) = iCurrentIndex Then
            grdGrid.SelectedIndex = itemIndex
            grdGrid.DataBind() ' Need to bind (again) to reset selection.
            Exit For
          End If
        Next
      Else
        grdGrid.SelectedIndex = 0
        grdGrid.DataBind() ' Need to bind (again) to reset selection.
      End If
    End If
  End Sub


  Private Sub RecordSelector_PageIndexChanging(ByVal sender As Object, ByVal e As GridViewPageEventArgs) Handles Me.PageIndexChanging

    grdGrid = CType(sender, GridView)

    grdGrid.PageIndex = e.NewPageIndex

    dataTable = TryCast(HttpContext.Current.Session(grdGrid.ID.Replace("Grid", "DATA")), DataTable)

    If IsLookup Then
      ' reapply filter?
      dataTable = SetLookupFilter(dataTable)
    End If

    grdGrid.DataSource = dataTable
    grdGrid.DataBind()

  End Sub


  Private Sub RecordSelector_RowCreated(ByVal sender As Object, ByVal e As GridViewRowEventArgs) Handles Me.RowCreated

    ' custom header only - manually add the rows for sorting.
    If e.Row.RowType = DataControlRowType.Header Then

      Dim tcTableCell As TableCell
      ' Dim d As GridView = sender
      grdGrid = sender

      For Each tcTableCell In e.Row.Cells
        If grdGrid.AllowSorting Then
          Dim lb As System.Web.UI.WebControls.LinkButton
          ' lb = tcTableCell.Controls(0)
          lb = CType(tcTableCell.Controls(0), System.Web.UI.WebControls.LinkButton)
          tcTableCell.Text = lb.Text

          If MyBase.HeaderStyle.Height.Value < 21 Then MyBase.HeaderStyle.Height = Unit.Pixel(NullSafeSingle(Me.HeadFontSize) * 2)
          tcTableCell.ApplyStyle(MyBase.HeaderStyle)
        Else
          tcTableCell.Text &= ""
          If MyBase.HeaderStyle.Height.Value < 21 Then MyBase.HeaderStyle.Height = Unit.Pixel(NullSafeSingle(Me.HeadFontSize) * 2)
          tcTableCell.ApplyStyle(MyBase.HeaderStyle)

        End If
      Next
    ElseIf e.Row.RowType = DataControlRowType.DataRow Then
      ' ID for this row is used to calculate row height at runtime. Can't
      ' specify the row height cos of bug HRPRO-1685.
      ' NB: no need to be unique here, as the ID will automatically be 
      ' prefixed with the control's clientID.
      'PG removed for performance and altered some js where ids where used
      'e.Row.ID = "row" & e.Row.RowIndex.ToString
    End If
  End Sub

  Private Sub RecordSelector_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles Me.RowDataBound
    ' As each row is added to the grid's HTML table, do the following: 
    'check the item being bound is actually a DataRow, if it is, 
    'wire up the required html events and attach the relevant JavaScripts

    grdGrid = CType(sender, RecordSelector)

    Try
      If e.Row.RowType = DataControlRowType.DataRow Then

        ' loop through the columns of this row. Hide ID columns
        For iColCount As Integer = 0 To e.Row.Cells.Count - 1

          Dim sColumnCaption As String = UCase(grdGrid.HeaderRow.Cells(iColCount).Text)

          Dim hidden As Boolean = False

          If (Not IsLookup AndAlso (sColumnCaption = "ID" OrElse (Left(sColumnCaption, 3) = "ID_" AndAlso Val(Mid(sColumnCaption, 4)) > 0))) OrElse (IsLookup AndAlso sColumnCaption.StartsWith("ASRSYS")) Then
            hidden = True
            e.Row.Cells(iColCount).Style.Add("display", "none")
          End If

          ' Format the cells according to DataType
          Dim curSelDataType As String = vbNullString

          If grdGrid.HeaderRow.Cells(iColCount).Text <> vbNullString Then

            curSelDataType = DataBinder.Eval(e.Row.DataItem, grdGrid.HeaderRow.Cells(iColCount).Text).GetType.ToString.ToUpper

            If curSelDataType.Contains("INT") OrElse curSelDataType.Contains("DECIMAL") OrElse curSelDataType.Contains("SINGLE") OrElse curSelDataType.Contains("DOUBLE") Then curSelDataType = "Integer"
            If curSelDataType.Contains("DATETIME") Then curSelDataType = "DateTime"
            If curSelDataType.Contains("BOOLEAN") OrElse curSelDataType.Contains("DBNULL") Then curSelDataType = "Boolean"
          End If
          Try
            Select Case curSelDataType
              Case "DateTime"
                ' Is the cell a date? 
                Dim value As DateTime = DateTime.Parse(e.Row.Cells(iColCount).Text.ToString())
                e.Row.Cells(iColCount).Text = value.ToShortDateString()

                If Not hidden Then e.Row.Cells(iColCount).Style.Add("text-align", "center")
              Case "Integer"
                If Not hidden Then e.Row.Cells(iColCount).Style.Add("text-align", "right")
              Case "Boolean"
                If Not hidden Then e.Row.Cells(iColCount).Style.Add("text-align", "center")
              Case Else   ' String
                ' Careful here: &nbsp is not a real space (it's chr160, not chr20) - might cause probs somewhere down the line....
                e.Row.Cells(iColCount).Text = e.Row.Cells(iColCount).Text.Replace(" ", "&nbsp;")
            End Select
          Catch
            ' um...
          End Try
        Next

        If Not Me.IsEmpty Then
          e.Row.Attributes("onclick") = "SR(this, " & e.Row.RowIndex & ")"
        End If

      ElseIf e.Row.RowType = DataControlRowType.Header Then

        ' Get the lookupfiltervalue column number, if applicable and store to a tag.
        For iColCount As Integer = 0 To e.Row.Cells.Count - 1
          Dim sColumnCaption As String = UCase(e.Row.Cells(iColCount).Text)

          If sColumnCaption.ToUpper = "ASRSYSLOOKUPFILTERVALUE" Then
            grdGrid.Attributes.Remove("LookupFilterColumn")
            grdGrid.Attributes.Add("LookupFilterColumn", iColCount.ToString)
          End If
        Next

      End If
    Catch ex As Exception

    End Try


  End Sub


  Private Function GetHexColor(ByVal aRGBCode As System.Drawing.Color) As String
    Dim strHEX As String

    'Dim A As String = Convert.ToString(aRGBCode.A, 16).PadLeft(2, "0"c).ToUpper
    Dim R As String = Convert.ToString(aRGBCode.R, 16).PadLeft(2, "0"c).ToUpper
    Dim G As String = Convert.ToString(aRGBCode.G, 16).PadLeft(2, "0"c).ToUpper
    Dim B As String = Convert.ToString(aRGBCode.B, 16).PadLeft(2, "0"c).ToUpper

    strHEX = "#" & R & G & B

    Return strHEX

  End Function

  Private Sub RecordSelector_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SelectedIndexChanged

    ' Set the dropdown value to the selected item
    If IsLookup Then

      grdGrid = TryCast(sender, GridView)

      ddl = DirectCast(Parent.FindControl(grdGrid.ID.Replace("Grid", "TextBox")), TextBox)

      ddl.Text = HttpUtility.HtmlDecode(grdGrid.Rows(grdGrid.SelectedIndex).Cells(CInt(ddl.Attributes("LookupColumnIndex"))).Text).ToString().Replace(Chr(160), Chr(32))
    End If
  End Sub

  Private Sub RecordSelector_SelectedIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSelectEventArgs) Handles Me.SelectedIndexChanging

    ' Dim gv As GridView = TryCast(sender, GridView)
    grdGrid = TryCast(sender, RecordSelector)

    ' Get the current page and row index numbers
    Dim iNewSelectedIndex As Integer = (grdGrid.PageIndex * grdGrid.PageSize) + e.NewSelectedIndex

    ' Fault HRPRO-1645, use the ID of the record for RecordSelectors
    If IsLookup Then
      Me.SelectedID(grdGrid) = iNewSelectedIndex
    Else
      ' get the id column number.
      Dim iColCount As Integer
      Dim sColCaption As String
      Dim iIDColNum As Integer = -1
      For iColCount = 0 To grdGrid.HeaderRow.Cells.Count - 1
        sColCaption = grdGrid.HeaderRow.Cells(iColCount).Text
        If sColCaption.ToUpper = "ID" Then
          ' Here it is
          iIDColNum = iColCount
          Exit For
        End If
      Next

      ' Get the ID from the column and store it.
      Me.SelectedID(grdGrid) = If(iIDColNum > 0, NullSafeInteger(grdGrid.Rows(e.NewSelectedIndex).Cells(iIDColNum).Text), -1)

    End If

  End Sub

  Public Property SelectedID(ByVal gv As GridView) As Integer
    Get
      If gv.Attributes("SelectedID") Is Nothing Then
        Return -1
      Else
        Return gv.Attributes("SelectedID")
      End If
    End Get

    Set(ByVal value As Integer)

      gv.Attributes("SelectedID") = value
    End Set
  End Property

  Private Sub RecordSelector_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Sorted

    ' Default behaviour of the gridview is to highlight the same row again when 
    ' the page is changed. We'll override that.

    grdGrid = CType(sender, System.Web.UI.WebControls.GridView)

    Dim iCurrentIndex As Integer

    grdGrid.SelectedIndex = -1
    grdGrid.DataBind()

    If Not IsLookup Then    ' don't handle lookup sorting for now - no ID's you see.
      ' get the grid's id column number.
      Dim iColCount As Integer
      Dim sColCaption As String
      Dim iIDColNum As Integer = -1
      For iColCount = 0 To grdGrid.HeaderRow.Cells.Count - 1
        sColCaption = grdGrid.HeaderRow.Cells(iColCount).Text
        If sColCaption.ToUpper = "ID" Then
          ' Here it is
          iIDColNum = iColCount
          Exit For
        End If
      Next
      If iIDColNum > 0 Then
        For itemIndex As Integer = 0 To grdGrid.Rows.Count - 1
          iCurrentIndex = NullSafeInteger(grdGrid.Rows(itemIndex).Cells(iIDColNum).Text)
          If Me.SelectedID(grdGrid) = iCurrentIndex Then
            grdGrid.SelectedIndex = itemIndex
            grdGrid.DataBind() ' Need to bind (again) to reset selection.
            Exit For
          End If
        Next
      Else
        grdGrid.SelectedIndex = 0
        grdGrid.DataBind() ' Need to bind (again) to reset selection.
      End If
    End If
  End Sub

  Private Sub RecordSelector_Sorting(ByVal sender As Object, ByVal e As GridViewSortEventArgs) Handles Me.Sorting

    grdGrid = CType(sender, GridView)

    GridViewSort(sender, e)
    'this handles the flipping of the sort direction

    Dim SortSQL As String = SortExpressionToSQL(e.SortExpression.Replace(" ", "_").Replace("+", ""), e.SortDirection)

    ' Get the current dataset from the session variable,
    ' Sort it, then store back to session variable.
    dataTable = TryCast(HttpContext.Current.Session(grdGrid.ID.Replace("Grid", "DATA")), DataTable)

    If IsLookup Then
      ' reapply filter?
      dataTable = SetLookupFilter(dataTable)
    End If

    If dataTable IsNot Nothing Then

      Dim dataView As New DataView(dataTable)

      Dim hasBlankRow As Boolean = dataTable.Rows.Count > 0 AndAlso Array.TrueForAll(dataTable.Rows(0).ItemArray, Function(v) IsDBNull(v))
      'remove blank row at top if exists
      If hasBlankRow Then
        dataTable.Rows.RemoveAt(0)
      End If

      If Right(e.SortExpression, 1) = "+" Then
        ' Control Click on column - APPEND to sort order - comma delimited.
        If Me.sortBy.Length > 0 Then
          If Me.sortBy.Contains("[" & e.SortExpression.Replace(" ", "_").Replace("+", "") & "]") Then
            ' this sort order already set, so update.

            Me.sortBy = Me.sortBy.Replace("[" & e.SortExpression.Replace(" ", "_").Replace("+", "") & "]" & _
                                          If(e.SortDirection = WebControls.SortDirection.Ascending, " Desc", " Asc"), "[" & e.SortExpression.Replace(" ", "_").Replace("+", "") & "]" & _
                                          If(e.SortDirection = WebControls.SortDirection.Ascending, " Asc", " Desc"))
          Else
            Me.sortBy &= "," & SortSQL.Replace("+", "")
          End If
        Else
          Me.sortBy = SortSQL.Replace("+", "")
        End If
      Else
        ' Not ctrl-click, so just sort by this column.
        Me.sortBy = SortSQL
      End If

      dataView.Sort = Me.sortBy   ' SortSQL ' Convert.ToString(e.SortExpression).Replace(" ", "_") & " DESC"
      dataTable = dataView.ToTable()

      If hasBlankRow Then
        'add blank row at top if needed
        dataTable.Rows.InsertAt(dataTable.NewRow(), 0)
      End If

      HttpContext.Current.Session(grdGrid.ID.Replace("Grid", "DATA")) = dataTable
      grdGrid.DataSource = dataTable
      grdGrid.DataBind()
    End If

  End Sub

  Public Shared Function SortExpressionToSQL(ByVal SortExpression As String, ByVal sortDir As System.Nullable(Of SortDirection)) As String
    Return ((If((SortExpression Is Nothing), Nothing, "[" & SortExpression & "]")) & " " & sortDir.ToString.Replace("ending", ""))
  End Function

  Public Shared Sub GridViewSort(ByVal sender As Object, ByVal e As GridViewSortEventArgs)
    'this technique leverages the GridView's generic Attributes collection to store the sort column & direction state between runs
    'coupled with the "sender" argument of most event handlers, WebControl.Attributes becomes a handy little piece of ViewState for stuff like this
    Dim grdGrid As GridView = TryCast(sender, GridView)
    Dim previousColumn As String = grdGrid.Attributes("SortColumn")
    grdGrid.Attributes("SortColumn") = e.SortExpression.Replace(" ", "_")
    If previousColumn = e.SortExpression.Replace(" ", "_") Then
      Dim previousDirection As String = grdGrid.Attributes("SortDirection")
      'if we haven't sorted this column yet, default to Ascending (by passing Descending into the flipper)
      e.SortDirection = Flip(If((previousDirection Is Nothing), SortDirection.Descending, DirectCast([Enum].Parse(GetType(SortDirection), previousDirection), SortDirection)))
      grdGrid.Attributes("SortDirection") = e.SortDirection.ToString()
    Else
      grdGrid.Attributes("SortDirection") = SortDirection.Ascending.ToString
    End If
  End Sub

  Public Shared Function Flip(ByVal sortDir As SortDirection) As SortDirection
    Return (If((sortDir = SortDirection.Ascending), SortDirection.Descending, SortDirection.Ascending))
  End Function

  Function SetLookupFilter(ByVal dt As DataTable) As DataTable

    If dt IsNot Nothing Then
      Dim dataView As New DataView(dt)
      dataView.RowFilter = Me.filterSQL    '   "ISNULL([ASRSysLookupFilterValue], '') = 'HERTFORDSHIRE'"

      dt = dataView.ToTable()

      If dt.Rows.Count = 0 Then
        ' create a blank row to display.
        Dim objDataRow As DataRow
        objDataRow = dt.NewRow()
        dt.Rows.InsertAt(objDataRow, 0)
      End If
    End If

    Return dt

  End Function


  Private Sub InitCustomPager(ByVal row As GridViewRow, ByVal columnSpan As Integer, ByVal pagedDataSource As PagedDataSource)

    Dim strPagerFontSize As String = "7"

    Dim pnlPager As Panel = New Panel()
    With pnlPager
      .ID = "pnlPager"
      .Height = CalculatePagerHeight().Replace("px", "")
    End With

    Dim tblPager As Table = New Table()
    With tblPager
      .ID = "tblPager"
      .CellSpacing = 0
      .Style.Add("width", "100%")
      .Style.Add("height", "100%")
      .BorderStyle = BorderStyle.None
      .GridLines = GridLines.Both
    End With

    Dim trPager As TableRow = New TableRow()
    trPager.ID = "trPager"
    trPager.Style.Add("width", "100%")
    trPager.Style.Add("border", "0px")

    Dim ltlPageIndex As Literal = New Literal()
    ltlPageIndex.ID = "ltlPageIndex"
    ltlPageIndex.Text = (Me.PageIndex + 1).ToString()

    Dim ltlPageCount As Literal = New Literal()
    ltlPageCount.ID = "ltlPageCount"
    ltlPageCount.Text = Me.PageCount.ToString()

    Dim tcSearchCell As TableCell = New TableCell()
    With tcSearchCell
      ' .ID = "tcSearchCell"
      .Attributes.Add("id", MyBase.ID.Replace("Grid", "") & "tcSearch") ' add as an attribute to ensure unique ID (other ctlxxx is added at runtime)
      .Style.Add("width", "150px")
      .Style.Add("height", "100%")
      .BorderStyle = WebControls.BorderStyle.None
    End With

    Dim txtSearchBox As TextBox = New TextBox
    With txtSearchBox
      .Style.Add("font-size", "8pt")
      .Style.Add("font-style", "italic")
      .Style.Add("float", "left")
      .Style.Add("top", "2px")
      .Style.Add("left", "3px")
      .Style.Add("width", "150px")
      .Style.Add("height", "15px")
      .Style.Add("border", "solid 1px gray")
      .Text = "filter page..."
      .Attributes.Add("onblur", "if(this.value==""""){this.style.fontStyle=""italic"";this.style.color=""gray"";this.value=""filter page...""}")
      .Attributes.Add("onfocus", "if(this.value=='filter page...'){this.style.color=""black"";this.style.fontStyle=""normal"";this.value=""""}")
      .Attributes.Add("onclick", "event.cancelBubble=true;")
      .Attributes.Add("onkeyup", "filterTable(this, '" & Me.ClientID.ToString & "')")
    End With

    tcSearchCell.Controls.Add(txtSearchBox)

    Dim imgSearchImg As Image = New Image
    With imgSearchImg
      .Style.Add("position", "absolute")
      .Style.Add("top", "4px")
      .Style.Add("left", "136px")
      .Style.Add("height", "15px")
      .Style.Add("width", "15px")
      .ImageUrl = "Images/search.gif"
    End With

    tcSearchCell.Controls.Add(imgSearchImg)
    '  writer.Write("<img src='Images/search.gif' style='position:absolute;top:4px;left:140px;height:15px;width:15px'/>")


    Dim tcPageXofY As TableCell = New TableCell()
    With tcPageXofY
      '.ID = "tcPageXofY"
      .Attributes.Add("id", MyBase.ID.Replace("Grid", "") & "tcPageXofY") ' add as an attribute to ensure unique ID (other ctlxxx is added at runtime)
      .Style.Add("width", "35%")
      .Style.Add("text-align", "right")
      .Style.Add("padding-left", "5px")
      .Style.Add("border", "0px")
      .Height = Unit.Pixel(23)
      .Font.Size = FontUnit.Parse(strPagerFontSize)
      .Font.Name = HeaderStyle.Font.Name
      .Controls.Add(New LiteralControl("Page "))
      .Controls.Add(ltlPageIndex)
      .Controls.Add(New LiteralControl(" of "))
      .Controls.Add(ltlPageCount)
    End With

    Dim ibtnFirst As Button = New Button()
    With ibtnFirst
      .ID = "ibtnFirst"
      .CommandName = "First"
      .UseSubmitBehavior = False
      .ToolTip = "First Page"
      .Width = Unit.Pixel(16)
      .Height = Unit.Pixel(16)
      .Style.Add("cursor", "pointer")
      .Style.Add("border-style", "none")
      .CausesValidation = False
      .Attributes.Add("onclick", If(IsLookup, "try{$get('txtActiveDDE').value='" & grdGrid.ID.Replace("Grid", "dde") & "';setPostbackMode(3);}catch(e){};", "try{setPostbackMode(3);}catch(e){};"))
      AddHandler .Command, AddressOf Me.PagerCommand
    End With

    Dim ibtnPrevious As Button = New Button()
    With ibtnPrevious
      .ID = "ibtnPrevious"
      .CommandName = "Previous"
      .UseSubmitBehavior = False
      .ToolTip = "Previous Page"
      .Width = Unit.Pixel(16)
      .Height = Unit.Pixel(16)
      .Style.Add("cursor", "pointer")
      .Style.Add("border-style", "none")
      .CausesValidation = False
      .Attributes.Add("onclick", If(IsLookup, "try{$get('txtActiveDDE').value='" & grdGrid.ID.Replace("Grid", "dde") & "';setPostbackMode(3);}catch(e){};", "try{setPostbackMode(3);}catch(e){};"))
      AddHandler .Command, AddressOf Me.PagerCommand
    End With

    Dim ibtnNext As Button = New Button()
    With ibtnNext
      .ID = "ibtnNext"
      .CommandName = "Next"
      .UseSubmitBehavior = False
      .ToolTip = "Next Page"
      .Width = Unit.Pixel(16)
      .Height = Unit.Pixel(16)
      .Style.Add("cursor", "pointer")
      .Style.Add("border-style", "none")
      .CausesValidation = False
      .Attributes.Add("onclick", If(IsLookup, "try{$get('txtActiveDDE').value='" & grdGrid.ID.Replace("Grid", "dde") & "';setPostbackMode(3);}catch(e){};", "try{setPostbackMode(3);}catch(e){};"))
      AddHandler .Command, AddressOf Me.PagerCommand
    End With

    Dim ibtnLast As Button = New Button()
    With ibtnLast
      .ID = "ibtnLast"
      .CommandName = "Last"
      .UseSubmitBehavior = False
      .ToolTip = "Last Page"
      .Width = Unit.Pixel(16)
      .Height = Unit.Pixel(16)
      .Style.Add("cursor", "pointer")
      .Style.Add("border-style", "none")
      .CausesValidation = False
      .Attributes.Add("onclick", If(IsLookup, "try{$get('txtActiveDDE').value='" & grdGrid.ID.Replace("Grid", "dde") & "';setPostbackMode(3);}catch(e){};", "try{setPostbackMode(3);}catch(e){};"))
      AddHandler .Command, AddressOf Me.PagerCommand
    End With

    If Me.PageIndex > 0 Then
      ibtnFirst.Style.Add("background-image", "url('Images/page-first.gif')")
      ibtnPrevious.Style.Add("background-image", "url('Images/page-prev.gif')")
      ibtnFirst.Enabled = True
      ibtnPrevious.Enabled = True
    Else
      ibtnFirst.Style.Add("background-image", "url('Images/page-first-disabled.gif')")
      ibtnPrevious.Style.Add("background-image", "url('Images/page-prev-disabled.gif')")
      ibtnFirst.Enabled = False
      ibtnPrevious.Enabled = False
      ibtnFirst.Style.Add("cursor", "default")
      ibtnPrevious.Style.Add("cursor", "default")
    End If

    If Me.PageIndex < Me.PageCount - 1 Then
      ibtnNext.Style.Add("background-image", "url('Images/page-next.gif')")
      ibtnLast.Style.Add("background-image", "url('Images/page-last.gif')")
      ibtnNext.Enabled = True
      ibtnLast.Enabled = True
    Else
      ibtnNext.Style.Add("background-image", "url('Images/page-next-disabled.gif')")
      ibtnLast.Style.Add("background-image", "url('Images/page-last-disabled.gif')")
      ibtnNext.Enabled = False
      ibtnLast.Enabled = False
      ibtnNext.Style.Add("cursor", "default")
      ibtnLast.Style.Add("cursor", "default")
    End If

    Dim tcPagerBtns As TableCell = New TableCell()
    With tcPagerBtns
      '.ID = "tcPagerBtns"
      .Attributes.Add("id", MyBase.ID.Replace("Grid", "") & "tcPagerBtns") ' add as an attribute to ensure unique ID (other ctlxxx is added at runtime)
      .Style.Add("width", "35%")
      .Style.Add("text-align", "center")
      .Style.Add("border", "0px")
      .Controls.Add(ibtnFirst)
      .Controls.Add(ibtnPrevious)
      .Controls.Add(New LiteralControl("&nbsp;&nbsp;"))
      .Controls.Add(ibtnNext)
      .Controls.Add(ibtnLast)
    End With

    Dim ddlPages As DropDownList = New DropDownList()
    With ddlPages
      .ID = "ddlPages"
      ' .CssClass = "paging_gridview_pgr_ddl"
      .Font.Size = FontUnit.Parse(strPagerFontSize)
      .Font.Name = HeaderStyle.Font.Name
      .AutoPostBack = True
      For i As Integer = 1 To Me.PageCount Step +1
        .Items.Add(New ListItem(i.ToString(), i.ToString()))
      Next i
      .SelectedIndex = Me.PageIndex
      .CausesValidation = False
      .Attributes.Add("onclick", "event.cancelBubble=true;")
      .Attributes.Add("onchange", If(IsLookup, "try{$get('txtActiveDDE').value='" & grdGrid.ID.Replace("Grid", "dde") & "';setPostbackMode(3);}catch(e){};", "try{setPostbackMode(3);}catch(e){};"))
      AddHandler .SelectedIndexChanged, AddressOf Me.ddlPages_SelectedIndexChanged
    End With

    Dim tcPagerDDL As TableCell = New TableCell()
    With tcPagerDDL
      ' .ID = "tcPagerDDL"
      .Attributes.Add("id", MyBase.ID.Replace("Grid", "") & "tcPagerDDL") ' add as an attribute to ensure unique ID (other ctlxxx is added at runtime)
      .Font.Size = FontUnit.Parse(strPagerFontSize)
      .Font.Name = HeaderStyle.Font.Name
      .Style.Add("width", "30%")
      .Style.Add("text-align", "right")
      .Style.Add("padding-right", "5px")
      .Style.Add("border", "0px")
      .Controls.Add(New LiteralControl("<table cellpadding=""0"" cellspacing=""0"" frame=""void"" rules=""none""><tr><td style=""padding-right: 5px;border: 0px;"">Page:</td><td style=""border: 0px;"">"))
      .Controls.Add(ddlPages)
      .Controls.Add(New LiteralControl("</td></tr></table>"))
    End With

    ' Hide navigation buttons depending on width of Record Selector
    ' NB Lookup navigation buttons are hidden at runtime using the 'ResizeComboForForm' JS function.
    If Not IsLookup Then
      ' Hide Search box if required
      If MyBase.Width.Value < 420 Then
        'tcSearchCell.Style.Add("display", "none")
        'tcSearchCell.Style.Add("visibility", "hidden")
        tcPageXofY.Style.Add("display", "none")
        tcPageXofY.Style.Add("visibility", "hidden")
        tcPagerBtns.Style.Add("display", "none")
        tcPagerBtns.Style.Add("visibility", "hidden")
      End If

      If MyBase.Width.Value < 175 Then
        ' Hide tcPageXofY AND tcPagerBtns
        'tcPageXofY.Style.Add("display", "none")
        'tcPageXofY.Style.Add("visibility", "hidden")
        'tcPagerBtns.Style.Add("display", "none")
        'tcPagerBtns.Style.Add("visibility", "hidden")
        'trPager.Style.Remove("width")
        'trPager.Style.Add("width", "75px")
        tcSearchCell.Style.Add("display", "none")
        tcSearchCell.Style.Add("visibility", "hidden")
      ElseIf MyBase.Width.Value < 210 Then
        ' Hide just tcPageXofY
        tcPageXofY.Style.Add("display", "none")
        tcPageXofY.Style.Add("visibility", "hidden")
        tcPagerBtns.Style.Add("display", "none")
        tcPagerBtns.Style.Add("visibility", "hidden")
        'trPager.Style.Remove("width")
        'trPager.Style.Add("width", "175px")
      End If

    End If


    'add cells to row
    trPager.Cells.Add(tcSearchCell)
    If MyBase.PageCount > 1 Then
      trPager.Cells.Add(tcPageXofY)
      trPager.Cells.Add(tcPagerBtns)
      trPager.Cells.Add(tcPagerDDL)
    End If

    'add row to table
    tblPager.Rows.Add(trPager)

    'add table to div
    ' NPG20120202 Fault HRPRO-1854, hide the pager row if there are no records.
    If Not Me.IsEmpty Then pnlPager.Controls.Add(tblPager)

    'add div to pager row
    row.Controls.AddAt(0, New TableCell())
    row.Cells(0).ColumnSpan = columnSpan
    row.Cells(0).Controls.Add(pnlPager)
  End Sub

    Protected Sub ddlPages_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim newPageIndex As Integer = CType(sender, DropDownList).SelectedIndex
        OnPageIndexChanging(New GridViewPageEventArgs(newPageIndex))
    End Sub

  Protected Sub PagerCommand(ByVal sender As Object, ByVal e As CommandEventArgs)
    Dim curPageIndex As Integer = Me.PageIndex
    Dim newPageIndex As Integer = 0

    Select Case e.CommandName
      Case "First"
        newPageIndex = 0
      Case "Previous"
        If curPageIndex > 0 Then
          newPageIndex = curPageIndex - 1
        End If
      Case "Next"
        If Not curPageIndex = Me.PageCount Then
          newPageIndex = curPageIndex + 1
        End If
      Case "Last"
        newPageIndex = Me.PageCount
    End Select

    OnPageIndexChanging(New GridViewPageEventArgs(newPageIndex))
  End Sub
End Class


