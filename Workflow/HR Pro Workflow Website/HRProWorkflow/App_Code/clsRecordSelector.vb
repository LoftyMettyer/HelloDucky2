Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.ComponentModel
Imports System.Web

Imports Utilities
Imports System.Data

Public Class RecordSelector
    Inherits GridView

    Protected WithEvents customHeader As GridViewRow
    Protected WithEvents customPager As GridViewRow
    Protected WithEvents grdGrid As RecordSelector ' GridView
    Protected WithEvents pagerTable As Table
    Protected WithEvents pagerRow As TableRow
    Protected WithEvents ddl As DropDownList
    Protected WithEvents dataTable As DataTable

    Private m_iPagerHeight As Integer = 30
    Private Const MAXDROPDOWNROWS As Int16 = 6
    Private m_iVisibleColumnCount As Integer
    Dim iColWidth As Integer = 100  ' default, minimum column width for all grid columns

    Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
        Dim iIDColumnIndex As Int16
        Dim sDivStyle As String = ""
        Dim hdnScrollTop As String = "0"
        Dim objGeneral As New General

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
                    If MyBase.Rows.Count > 1 Then
                        If IsLookup Then
                            customHeader.Cells(iColCount).Attributes("onclick") = ("txtActiveDDE.value='" & grdGrid.ID.Replace("Grid", "dde") & "';try{setPostbackMode(2);}catch(e){};__doPostBack('" & MyBase.UniqueID & "','Sort$") & customHeader.Cells(iColCount).Text & "');"
                        Else
                            customHeader.Cells(iColCount).Attributes("onclick") = ("try{setPostbackMode(2);}catch(e){};if(event.ctrlKey){__doPostBack('" & MyBase.UniqueID & "','Sort$" & customHeader.Cells(iColCount).Text & "+');}else{__doPostBack('" & MyBase.UniqueID & "','Sort$") & customHeader.Cells(iColCount).Text & "');}"
                        End If
                    End If
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

                    End With

                    ctlSortPanel.Controls.Add(ctlSortImage)
                    customHeader.Cells(iColCount).Controls.Add(ctlSortPanel)

                End If
            Next
        Else
            ' No records to display.....
            Return
        End If


        ' Div to contain all items - too small for lookups
        writer.Write("<div ID='" & Me.ID.ToString.Replace("Grid", "") & "' " & _
            IIf(IsLookup, "style='", "style='position:absolute;") & _
            " width:" & CalculateWidth() & ";height:" & CalculateHeight() & ";" & _
            IIf(IsLookup, "", "top:" & Me.Style.Item("TOP").ToString & ";") & _
            IIf(IsLookup, "", "left:" & Me.Style.Item("LEFT").ToString & ";") & _
            sDivStyle & _
            ";overflow:hidden;border-color:black;border-style:solid;border-width:1px;background-color:" & _
            objGeneral.GetHTMLColour(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(customHeader.BackColor.A, customHeader.BackColor.R, customHeader.BackColor.G, customHeader.BackColor.B)).ToString) & _
            "'>")

        'render header row 
        writer.Write("<table ID='" & Me.ID.ToString.Replace("Grid", "") & "Header'" & _
                    " style='position:absolute;top:0px;left:0px;width:" & CalculateGridWidth() & ";height:" & CalculateHeaderHeight() & _
                    ";table-layout:fixed;border:0'" & _
                    " cellspacing='" & Me.CellSpacing.ToString() & "'" & _
                    IIf((MyBase.HeaderRow IsNot Nothing), " class='resizable'", "") & _
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
                     ";bottom:" & CalculatePagerHeight() & ";left:0px;overflow-x:auto;overflow-y:auto;" & _
                     "width:" & CalculateWidth() & ";" & "background-color:#FFFFFF;' onscroll=scrollHeader('" & ClientID.Replace("Grid", "gridcontainer") & "')>")


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
        If customPager IsNot Nothing AndAlso Me.PageCount > 1 Then
            writer.Write("<div  border='0' cellspacing='" & Me.CellSpacing.ToString() & "' cellpadding='" & _
                            Me.CellPadding.ToString() & "' style='position:absolute;right:0px;bottom:0px;height: " & CalculatePagerHeight() & ";'>")
            customPager.ApplyStyle(Me.PagerStyle)
            customPager.Visible = True
            customPager.RenderControl(writer)
            writer.Write("</div>")
        End If

        writer.Write("</div>")


    End Sub

    Private Function CalculateWidth() As [String]
        Dim strWidth As String = "auto"

        If Not IsLookup Then
            If Not Me.Width.IsEmpty Then
                strWidth = [String].Format("{0}{1}", Me.Width.Value, (If((Me.Width.Type = UnitType.Percentage), "%", "px")))
            End If
        Else
            Dim iGridWidth As Integer = m_iVisibleColumnCount * (iColWidth + 2) ' 2 = padding
            iGridWidth = CInt(IIf(iGridWidth < 0, 1, iGridWidth))
            iGridWidth = CInt(IIf(iGridWidth < Me.Width.Value, Me.Width.Value, iGridWidth))

            ' do rows exceed height?
            If MyBase.Rows.Count > MAXDROPDOWNROWS Then
                ' Add scrollbar width
                iGridWidth += 25
            End If

            strWidth = [String].Format("{0}{1}", iGridWidth, (If((Me.Width.Type = UnitType.Percentage), "%", "px")))
        End If

        Return strWidth

    End Function
    Private Function CalculateHeight() As [String]
        Dim strHeight As String = "auto"
        Dim iHeight As Integer = ControlHeight

        If Not IsLookup Then
            If iHeight > 0 Then
                strHeight = [String].Format("{0}{1}", iHeight, "px")
            End If
        Else            
            ' Set the size of the grid as per old DropDown setting...
            ' remember that the height for lookups will only be c.21pixels...
            Dim iRowHeight As Integer = iHeight - 6
            iRowHeight = CInt(IIf(iRowHeight < 21, 21, iRowHeight))
            Dim iDropHeight As Integer = (iRowHeight * CInt(IIf(MyBase.Rows.Count > MAXDROPDOWNROWS, MAXDROPDOWNROWS, MyBase.Rows.Count))) + 1

            iDropHeight = iDropHeight + iRowHeight  ' add row for headers
            iDropHeight = iDropHeight + IIf(MyBase.PageCount > 0, iRowHeight, 0)   ' add row for pager if required.

            strHeight = [String].Format("{0}{1}", iDropHeight, "px")
        End If

        Return strHeight

    End Function

    Private Sub AdjustWidthForScrollbar()

        Dim strHeight As String = "auto"

        ''Adjust available width for the vertical scrollbar.
        'iGapBetweenBorderAndText = (CInt(NullSafeSingle(dr("FontSize")) + 6) \ 4)
        'iEffectiveRowHeight = CInt(NullSafeSingle(dr("FontSize"))) _
        ' + 1 _
        ' + (2 * iGapBetweenBorderAndText)

        'iTempHeight = NullSafeInteger(ctlForm_GridContainer.Height.Value)
        'iTempHeight = CInt(IIf(iTempHeight < 0, 1, iTempHeight))

        'MyBase.Style.Remove("width")
        'MyBase.Style.Add("width", "183px")

    End Sub

    Private Function CalculateGridWidth() As [String]
        ' grid width is me.width - vertical scroll bar.
        Dim strWidth As String = "auto"
        Dim iScrollBarWidth As Integer = 0

        If MyBase.Rows.Count > MAXDROPDOWNROWS Then
            iScrollBarWidth = 17
        End If

        If Not IsLookup Then
            If Not Me.Width.IsEmpty Then
                strWidth = [String].Format("{0}{1}", Me.Width.Value - iScrollBarWidth, (If((Me.Width.Type = UnitType.Percentage), "%", "px")))
            End If
        Else
            Dim iGridWidth As Integer = m_iVisibleColumnCount * (iColWidth + 2) ' 2 = padding
            iGridWidth = CInt(IIf(iGridWidth < 0, 1, iGridWidth))
            iGridWidth = CInt(IIf(iGridWidth < Me.Width.Value, Me.Width.Value, iGridWidth))

            ' do rows exceed height?
            If MyBase.Rows.Count > MAXDROPDOWNROWS Then
                ' Add scrollbar width
                iGridWidth += 25
            End If

            strWidth = [String].Format("{0}{1}", iGridWidth - iScrollBarWidth, (If((Me.Width.Type = UnitType.Percentage), "%", "px")))
        End If

        Return strWidth


    End Function


    Private Function CalculateHeaderHeight() As [String]
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

        strHeaderHeight = [String].Format("{0}{1}", iHeaderHeight, "px")

        Return strHeaderHeight

    End Function

    Private Function CalculatePagerHeight() As [String]
        Dim strPagerHeight As String = "auto"
        Dim iPagerHeight As Integer = 0
        Dim iGridHeight As Integer = ControlHeight

        If Me.PageCount > 1 Then
            Dim iGridTopPadding As Integer = CInt(NullSafeSingle(Me.HeadFontSize) / 8)

            iPagerHeight = CInt((((NullSafeSingle(Me.HeadFontSize) + iGridTopPadding) * 2) - 2) * 1.5)

            If iPagerHeight > NullSafeInteger(iGridHeight) Then
                iPagerHeight = NullSafeInteger(iGridHeight)
            End If

        End If

        strPagerHeight = [String].Format("{0}{1}", iPagerHeight, "px")

        Return strPagerHeight

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

        grdGrid = CType(sender, System.Web.UI.WebControls.GridView)

        Dim iEffectiveRowHeight As Integer = grdGrid.RowStyle.Height.Value
        Dim iRowCount As Integer = grdGrid.Rows.Count

        grdGrid.Height = Unit.Pixel(iEffectiveRowHeight * iRowCount)

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

        grdGrid = CType(sender, System.Web.UI.WebControls.GridView)

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


    Private Sub RecordSelector_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles Me.PageIndexChanging
        ' Dim g As System.Web.UI.WebControls.GridView
        Dim iIDCol As Integer = 0
        grdGrid = CType(sender, System.Web.UI.WebControls.GridView)

        grdGrid.PageIndex = e.NewPageIndex
        ' grdGrid.DataSource = TryCast(HttpContext.Current.Session(grdGrid.ID.Replace("Grid", "DATA")), DataTable)
        dataTable = TryCast(HttpContext.Current.Session(grdGrid.ID.Replace("Grid", "DATA")), DataTable)

        If IsLookup Then
            ' reapply filter?
            dataTable = SetLookupFilter(dataTable)
        End If

        grdGrid.DataSource = dataTable
        grdGrid.DataBind()

    End Sub


    Private Sub RecordSelector_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles Me.RowCreated

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
        End If
    End Sub

    Private Sub RecordSelector_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles Me.RowDataBound
        ' As each row is added to the grid's HTML table, do the following: 
        'check the item being bound is actually a DataRow, if it is, 
        'wire up the required html events and attach the relevant JavaScripts

        Dim iIDCol As Integer = 0
        Dim sRowID As String = "0"
        'Dim grdGrid As System.Web.UI.WebControls.GridView
        Dim mydte As DateTime
        Dim sColumnCaption As String
        Dim objGeneral As New General

        grdGrid = CType(sender, RecordSelector)

        Try

            If e.Row.RowType = DataControlRowType.DataRow Then
                e.Row.Style.Add("overflow", "hidden")
                If Not Me.IsEmpty Then e.Row.Style.Add("cursor", "pointer")

                ' loop through the columns of this row. Hide ID columns
                For iColCount As Integer = 0 To e.Row.Cells.Count - 1
                    sColumnCaption = UCase(grdGrid.HeaderRow.Cells(iColCount).Text)

                    If (Not IsLookup And (sColumnCaption = "ID" Or (Left(sColumnCaption, 3) = "ID_" And Val(Mid(sColumnCaption, 4)) > 0))) Or _
                        (IsLookup And sColumnCaption.StartsWith("ASRSYS")) Then

                        'If sColumnCaption = "ID" Or (Left(sColumnCaption, 3) = "ID_" And Val(Mid(sColumnCaption, 4)) > 0) Then
                        iIDCol = iColCount  ' store ID column number to assign to the javascript click event.
                        e.Row.Cells(iColCount).Style.Add("display", "none")
                        If sColumnCaption = "ID" Then
                            ' Background colour to black.
                            ' Javascript can see this and use it to recognise the ID column. 
                            e.Row.Cells(iColCount).BackColor = Drawing.Color.Black
                        End If
                    End If

                    ' add overflow hidden and nowrap, stops the cells wrapping text or overflowing into adjacent cols.
                    e.Row.Cells(iColCount).Style.Add("overflow", "hidden")
                    e.Row.Cells(iColCount).Style.Add("white-space", "nowrap")

                    ' this sets minimum width, not max.
                    e.Row.Cells(iColCount).Width = Unit.Pixel(iColWidth)

                    ' Format the cells according to DataType
                    Dim curSelDataType As String = vbNullString
                    ' Dim curSelDataType As String = DataBinder.Eval(e.Row.DataItem, grdGrid.HeaderRow.Cells(iColCount).Text).GetType.ToString.ToUpper

                    If grdGrid.HeaderRow.Cells(iColCount).Text <> vbNullString Then

                        curSelDataType = DataBinder.Eval(e.Row.DataItem, grdGrid.HeaderRow.Cells(iColCount).Text).GetType.ToString.ToUpper

                        If curSelDataType.Contains("INT") _
                            Or curSelDataType.Contains("DECIMAL") _
                            Or curSelDataType.Contains("SINGLE") _
                            Or curSelDataType.Contains("DOUBLE") _
                            Then curSelDataType = "Integer"
                        If curSelDataType.Contains("DATETIME") Then curSelDataType = "DateTime"
                        If curSelDataType.Contains("BOOLEAN") Then curSelDataType = "Boolean"
                    End If
                    Try
                        Select Case curSelDataType
                            Case "DateTime"
                                ' Is the cell a date? 
                                mydte = DateTime.Parse(e.Row.Cells(iColCount).Text.ToString())
                                e.Row.Cells(iColCount).Text = mydte.ToShortDateString()
                            Case "Integer"
                                e.Row.Cells(iColCount).Style.Add("text-align", "right")
                            Case "Boolean"
                                e.Row.Cells(iColCount).Style.Add("text-align", "center")
                            Case Else   ' String                                
                                e.Row.Cells(iColCount).Text = e.Row.Cells(iColCount).Text.Replace(" ", "&nbsp;")
                                e.Row.Cells(iColCount).Style.Add("text-align", "left")
                        End Select
                    Catch
                        ' um...
                    End Try
                Next

                ' Add the javascript event to each row for click functionality
                'e.Row.Attributes.Add("onclick", "changeRow('" & grdGrid.ID.ToString & "', '" & e.Row.Cells(iIDCol).Text & "', '" & GetHexColor(grdGrid.SelectedRowStyle.BackColor) & "', '" & iIDCol.ToString & "');oldgridSelectedColor = this.style.backgroundColor;")
                ' Postback selection is really slow on big grids, so using javascript onclick above instead of the select event below.
                'e.Row.Attributes("onclick") = ("SetScrollTopPos('" & grdGrid.ID.ToString & "', document.getElementById('" & grdGrid.ID.Replace("Grid", "gridcontainer") & "').scrollTop);" & _
                '                               "try{setPostbackMode(2);}catch(e){};__doPostBack('" & grdGrid.UniqueID & "','Select$" & e.Row.RowIndex & "');" & _
                '                               "SetScrollTopPos('" & grdGrid.ID.ToString & "', -1);")
                If Not Me.IsEmpty Then
                    e.Row.Attributes("onclick") = ("SetScrollTopPos('" & grdGrid.ID.ToString & "', document.getElementById('" & grdGrid.ID.Replace("Grid", "gridcontainer") & "').scrollTop);" & _
                                                        "try{setPostbackMode(2);}catch(e){};__doPostBack('" & grdGrid.UniqueID & "','Select$" & e.Row.RowIndex & "');")
                End If
            ElseIf e.Row.RowType = DataControlRowType.Pager Then
                ' This enables postback for the grid

                ' Dim pagerTable As Table = DirectCast(e.Row.Cells(0).Controls(0), Table)
                ' Dim pagerRow As TableRow = pagerTable.Rows(0)

                pagerTable = DirectCast(e.Row.Cells(0).Controls(0), Table)
                pagerRow = pagerTable.Rows(0)

                For iColCount As Integer = 0 To pagerTable.Rows(0).Cells.Count - 1
                    If IsLookup Then
                        pagerRow.Cells(iColCount).Attributes.Add("onclick", "try{txtActiveDDE.value='" & grdGrid.ID.Replace("Grid", "dde") & "';setPostbackMode(2);}catch(e){};")
                    Else
                        pagerRow.Cells(iColCount).Attributes.Add("onclick", "try{setPostbackMode(2);}catch(e){};")
                    End If
                Next

            ElseIf e.Row.RowType = DataControlRowType.Header Then

                ' Get the lookupfiltervalue column number, if applicable and store to a tag.
                For iColCount As Integer = 0 To e.Row.Cells.Count - 1
                    sColumnCaption = UCase(e.Row.Cells(iColCount).Text)

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

        Dim A As String = Convert.ToString(aRGBCode.A, 16).PadLeft(2, "0"c).ToUpper
        Dim R As String = Convert.ToString(aRGBCode.R, 16).PadLeft(2, "0"c).ToUpper
        Dim G As String = Convert.ToString(aRGBCode.G, 16).PadLeft(2, "0"c).ToUpper
        Dim B As String = Convert.ToString(aRGBCode.B, 16).PadLeft(2, "0"c).ToUpper

        strHEX = "#" & R & G & B

        Return strHEX

    End Function

    Private Sub RecordSelector_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SelectedIndexChanged

        ' Set the dropdown value to the selected item
        If IsLookup Then
            ' Dim gv As GridView = TryCast(sender, GridView)
            grdGrid = TryCast(sender, GridView)
            ' Dim ddl As DropDownList = DirectCast(Parent.FindControl(gv.ID.Replace("Grid", "TextBox")), DropDownList)
            ddl = DirectCast(Parent.FindControl(grdGrid.ID.Replace("Grid", "TextBox")), DropDownList)

            ddl.Items.Clear()
            ddl.Items.Add(HttpUtility.HtmlDecode(grdGrid.Rows(grdGrid.SelectedIndex).Cells(grdGrid.Attributes("LookupColumnIndex")).Text).ToString())
            ddl.SelectedIndex = 0
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
            Me.SelectedID(grdGrid) = IIf(iIDColNum > 0, NullSafeInteger(grdGrid.Rows(e.NewSelectedIndex).Cells(iIDColNum).Text), -1)

        End If

        'Dim dataTable As DataTable
        'dataTable = DirectCast(grdGrid.DataSource, DataTable)

        ' Don't Databind - it sets all the ID's wrong when selecting...
        ' grdGrid.DataSource = TryCast(HttpContext.Current.Session(grdGrid.ID.Replace("Grid", "DATA")), DataTable)
        ' grdGrid.DataBind()

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

    Private Sub RecordSelector_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs) Handles Me.Sorting
        ' Dim g As System.Web.UI.WebControls.GridView
        Dim iIDCol As Integer = 0
        grdGrid = CType(sender, System.Web.UI.WebControls.GridView)

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
            '            SortSQL = "[Forenames] DESC, [Start_Date] ASC"

            If Right(e.SortExpression, 1) = "+" Then
                ' Control Click on column - APPEND to sort order - comma delimited.
                If Me.sortBy.Length > 0 Then
                    If Me.sortBy.Contains("[" & e.SortExpression.Replace(" ", "_").Replace("+", "") & "]") Then
                        ' this sort order already set, so update.

                        Me.sortBy = Me.sortBy.Replace("[" & e.SortExpression.Replace(" ", "_").Replace("+", "") & "]" & _
                                                      IIf(e.SortDirection = WebControls.SortDirection.Ascending, " Desc", " Asc"), "[" & e.SortExpression.Replace(" ", "_").Replace("+", "") & "]" & _
                                                      IIf(e.SortDirection = WebControls.SortDirection.Ascending, " Asc", " Desc"))
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
            'HttpContext.Current.Session(grdGrid.ID.Replace("Grid", "DATA")) = dataTable
            grdGrid.DataSource = dataView
            grdGrid.DataBind()
        End If

    End Sub

    Public Shared Function SortExpressionToSQL(ByVal SortExpression As String, ByVal sortDir As System.Nullable(Of System.Web.UI.WebControls.SortDirection)) As String
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
            ' HttpContext.Current.Session(btnSender.ID.Replace("refresh", "DATA")) = dataTable

            If dt.Rows.Count = 0 Then
                ' create a blank row to display.
                Dim objDataRow As System.Data.DataRow
                objDataRow = dt.NewRow()
                dt.Rows.InsertAt(objDataRow, 0)
            End If
        End If

        Return dt

    End Function

End Class


