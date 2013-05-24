Option Strict Off
Option Explicit On
Friend Class clsExprExpression
	
	' Expression definition variables.
	Private mlngExpressionID As Integer
	Private msExpressionName As String
	Private mlngBaseTableID As Integer
	Private miReturnType As Short
	Private miExpressionType As Short
	Private mlngParentComponentID As Integer
	Private msOwner As String
	Private msAccess As String
	Private msDescription As String
	Private mlngTimeStamp As Integer
	Private msBaseTableName As String
	Private mbViewInColour As Boolean
	Private mbExpandedNode As Boolean
	
	Public mfDontUpdateTimeStamp As Boolean
	
	' Class handling variables.
	Private mfConstructed As Boolean
	Private mcolComponents As Collection
	Private mobjBadComponent As clsExprComponent
	Private mobjBaseComponent As clsExprComponent
	
	Private msErrorMessage As String
	
	' Array holding the User Defined functions that are needed for this expression
	Private mastrUDFsRequired() As String
	
	Public ReadOnly Property ComponentDescription() As String
		Get
			' Return the expression name.
			ComponentDescription = msExpressionName
			
		End Get
	End Property
	
	
	
	Public Property ExpressionID() As Integer
		Get
			' Return the expression ID.
			ExpressionID = mlngExpressionID
			
		End Get
		Set(ByVal Value As Integer)
			' Set the expression ID.
			If mlngExpressionID <> Value Then
				mlngExpressionID = Value
				mfConstructed = False
			End If
			
		End Set
	End Property
	
	
	Public Property BaseTableID() As Integer
		Get
			' Return the expressions base table ID.
			BaseTableID = mlngBaseTableID
			
		End Get
		Set(ByVal Value As Integer)
			' Set the expression base table property.
			Dim sSQL As String
			Dim rsInfo As ADODB.Recordset
			
			If mlngBaseTableID <> Value Then
				mlngBaseTableID = Value
				
				' Read the parent table name.
				sSQL = "SELECT tableName" & " FROM ASRSysTables" & " WHERE tableID = " & Trim(Str(mlngBaseTableID))
				
				rsInfo = datGeneral.GetRecords(sSQL)
				
				If Not (rsInfo.EOF And rsInfo.BOF) Then
					msBaseTableName = rsInfo.Fields("TableName").Value
				End If
				
				rsInfo.Close()
				'UPGRADE_NOTE: Object rsInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				rsInfo = Nothing
			End If
			
		End Set
	End Property
	
	
	Public Property ReturnType() As Short
		Get
			' Return the expression's return type.
			ReturnType = miReturnType
			
		End Get
		Set(ByVal Value As Short)
			' Set the expression's return type.
			miReturnType = Value
			
		End Set
	End Property
	
	
	Public Property ExpressionType() As Short
		Get
			' Return the expression's parent type property.
			ExpressionType = miExpressionType
			
		End Get
		Set(ByVal Value As Short)
			' Set the expression's type property.
			miExpressionType = Value
			
		End Set
	End Property
	
	
	
	
	Public Property Name() As String
		Get
			' Return the expression name.
			If Not mfConstructed Then
				ConstructExpression()
			End If
			
			Name = msExpressionName
			
		End Get
		Set(ByVal Value As String)
			' Set the expression name.
			msExpressionName = Value
			
		End Set
	End Property
	
	Public Property BaseComponent() As clsExprComponent
		Get
			' Return the component's base component object.
			BaseComponent = mobjBaseComponent
			
		End Get
		Set(ByVal Value As clsExprComponent)
			' Set the component's base component object property.
			mobjBaseComponent = Value
			
		End Set
	End Property
	
	
	
	Public ReadOnly Property ErrorMessage() As String
		Get
			ErrorMessage = msErrorMessage
		End Get
	End Property
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	Public WriteOnly Property ParentComponentID() As Integer
		Set(ByVal Value As Integer)
			' Set the Parent component ID.
			mlngParentComponentID = Value
			
		End Set
	End Property
	
	
	Public ReadOnly Property ComponentType() As Short
		Get
			' Return the component type.
			ComponentType = modExpression.ExpressionComponentTypes.giCOMPONENT_EXPRESSION
			
		End Get
	End Property
	
	
	Public Property Components() As Collection
		Get
			' Return the component collection.
			Components = mcolComponents
			
		End Get
		Set(ByVal Value As Collection)
			' Set the component collection.
			mcolComponents = Value
			
		End Set
	End Property
	
	
	Public Property Owner() As String
		Get
			' Return the expression owner.
			Owner = msOwner
			
		End Get
		Set(ByVal Value As String)
			' Set the expression owner.
			msOwner = Value
			
		End Set
	End Property
	
	Public ReadOnly Property BadComponent() As clsExprComponent
		Get
			' Return the component last caused the expression to fail its validity check.
			BadComponent = mobjBadComponent
			
		End Get
	End Property
	
	
	Public Property Access() As String
		Get
			' Return the access code.
			Access = msAccess
			
		End Get
		Set(ByVal Value As String)
			' Set the access code.
			msAccess = Value
			
		End Set
	End Property
	
	
	
	
	
	
	
	
	Public Property Description() As String
		Get
			' Return the expression's description.
			Description = msDescription
			
		End Get
		Set(ByVal Value As String)
			' Set the expression's descriprion property.
			msDescription = Value
			
		End Set
	End Property
	
	
	Public Property Timestamp() As Integer
		Get
			' Return the expression's timestamp value.
			Timestamp = mlngTimeStamp
			
		End Get
		Set(ByVal Value As Integer)
			' Set the expression's timestamp property.
			mlngTimeStamp = Value
			
		End Set
	End Property
	
	
	
	
	
	
	Public Property BaseTableName() As String
		Get
			' Return the name of the expression's base table.
			BaseTableName = msBaseTableName
			
		End Get
		Set(ByVal Value As String)
			' Set the name of the expression's base table.
			msBaseTableName = Value
			
		End Set
	End Property
	
	
	
	
	
	
	
	
	Public Property ViewInColour() As Boolean
		Get
			
			ViewInColour = mbViewInColour
			
		End Get
		Set(ByVal Value As Boolean)
			
			mbViewInColour = Value
			
		End Set
	End Property
	
	
	Public Property ExpandedNode() As Boolean
		Get
			
			ExpandedNode = mbExpandedNode
			
		End Get
		Set(ByVal Value As Boolean)
			
			mbExpandedNode = Value
			
		End Set
	End Property
	
	Public Sub ResetConstructedFlag(ByRef fValue As Object)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mfConstructed = fValue
		
	End Sub
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		' Create a new collection to hold the expression's components.
		mcolComponents = New Collection
		mfConstructed = False
		mbExpandedNode = False
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		' Disassociate object variables.
		'UPGRADE_NOTE: Object mcolComponents may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mcolComponents = Nothing
		'UPGRADE_NOTE: Object mobjBadComponent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mobjBadComponent = Nothing
		'UPGRADE_NOTE: Object mobjBaseComponent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mobjBaseComponent = Nothing
		
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Public Function DeleteComponent(ByRef pobjComponent As clsExprComponent) As Boolean
		' Remove the given component from the expression.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		Dim iLoop As Short
		Dim iIndex As Short
		
		fOK = True
		iIndex = 0
		
		' Find the given component in the component collection.
		For iLoop = 1 To mcolComponents.Count()
			If pobjComponent Is mcolComponents.Item(iLoop) Then
				iIndex = iLoop
				Exit For
			End If
		Next iLoop
		
		' Delete the current component if it has been found.
		If iIndex > 0 Then
			mcolComponents.Remove(iIndex)
		End If
		
TidyUpAndExit: 
		DeleteComponent = fOK
		Exit Function
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	
	Public Function AddComponent() As clsExprComponent
		' Add a new component to the expression.
		' Returns the new component object.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		Dim objComponent As clsExprComponent
		
		' Instantiate a component object.
		objComponent = New clsExprComponent
		
		' Initialse the new component's properties.
		objComponent.ParentExpression = Me
		
		' Get the new component to handle its own definition.
		fOK = objComponent.NewComponent
		
		If fOK Then
			' If the component definition was confirmed then
			' add the new component to the expression's component
			' collection.
			mcolComponents.Add(objComponent)
		End If
		
TidyUpAndExit: 
		If fOK Then
			AddComponent = objComponent
		Else
			'UPGRADE_NOTE: Object AddComponent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			AddComponent = Nothing
		End If
		' Disassociate object variables.
		'UPGRADE_NOTE: Object objComponent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objComponent = Nothing
		Exit Function
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	
	Public Function SelectExpression(ByRef pfLockTable As Boolean, Optional ByRef plngOptions As Integer = 0) As Boolean
		'  ' Display the expression selection form.
		'  On Error GoTo ErrorTrap
		'
		'  Dim fOK As Boolean
		'  Dim fExit As Boolean
		'  Dim lngOldExpressionID As Long
		'  Dim sSQL As String
		'  Dim rsExpressions As Recordset
		'  Dim frmSelectExpr As frmDefSel
		'  Dim strSQL As String
		'
		'  fOK = (mlngBaseTableID > 0) Or (Not pfLockTable)
		'
		'  If fOK Then
		'    Set frmSelectExpr = New frmDefSel
		'
		'    fExit = False
		'    lngOldExpressionID = mlngExpressionID
		'
		'    With frmSelectExpr
		'
		'        ' Set the selection form properties.
		'        '.LockTable = pfLockTable
		'
		'        'Set .Expression = Me
		'
		'        ' Set the caption if necessary - only really needed because by default the
		'        ' caption is SELECT XXXXX, but when the user chooses CALCULATIONS from the
		'        ' utility menu, they arent really selecting one, so using this, we can change
		'        ' the form caption to Maintain Calculations or something like that !
		'        'If psCaption <> "" Then .Caption = psCaption
		'
		'
		'        strSQL = " type = " & CStr(ExpressionType) & _
		''                 " AND (returnType = " & CStr(ReturnType) & _
		''                      " OR type = " & CStr(giEXPR_RUNTIMECALCULATION) & ")" & _
		''                 " AND parentComponentID = 0"
		'
		'        .SelectedID = mlngExpressionID
		'
		'      ' Loop until an expression is selected, deselected, or the selection
		'      ' form is cancelled.
		'      Do While Not fExit
		'
		'        .TableId = BaseTableID
		'        .TableComboEnabled = Not (pfLockTable)
		'        .TableComboVisible = True
		'
		'        If plngOptions = 0 Then
		'          .Options = edtAdd + edtDelete + edtEdit + edtCopy + edtPrint + edtSelect + edtDeselect + edtProperties
		'        Else
		'          .Options = plngOptions
		'        End If
		'
		'        .EnableRun = False
		'
		'        Select Case ExpressionType
		'        Case giEXPR_RUNTIMEFILTER
		'          .ShowList "FILTERS", strSQL
		'
		'        Case giEXPR_RUNTIMECALCULATION
		'          .ShowList "CALCULATIONS", strSQL
		'
		'        End Select
		'
		'
		'        ' Display the selection form.
		'        .Show vbModal
		'        BaseTableID = .TableId
		'
		'        ' Execute the required operation.
		'        Select Case .Action
		'          ' Add a new expression.
		'          Case edtAdd
		'            NewExpression
		'            .SelectedID = ExpressionID
		'
		'          ' Edit the selected expression.
		'          Case edtEdit
		'            ExpressionID = .SelectedID
		'
		'            If .FromCopy Then
		'              CopyExpression
		'              If ExpressionID > 0 Then
		'                .SelectedID = ExpressionID
		'              End If
		'            Else
		'              EditExpression
		'            End If
		'
		'          ' Copy the selected expression.
		'          'Case edtCopy
		'          '  ExpressionID = .SelectedID
		'          '  CopyExpression
		'
		'          ' Print the selected expression.
		'          Case edtPrint
		'            ExpressionID = .SelectedID
		'            PrintExpression
		'
		'          ' Select the selected expression and return to the calling form.
		'          Case edtSelect
		'            ExpressionID = .SelectedID
		'
		'            ' Check that the selection is valid.
		'            If ValidateSelection Then
		'              fOK = True
		'              fExit = True
		'            End If
		'
		'          ' Deselect all expressions and return to the calling form.
		'          Case edtDeselect
		'            ExpressionID = 0
		'            msExpressionName = ""
		'            fOK = True
		'            fExit = True
		'
		'          ' Exit the selection form and return to the calling form.
		'          Case edtCancel
		'            ' Check if the original expression still exists.
		'            sSQL = "SELECT name" & _
		''              " FROM ASRSysExpressions" & _
		''              " WHERE exprID = " & Trim(Str(lngOldExpressionID))
		'            Set rsExpressions = datGeneral.GetRecords(sSQL)
		'            If rsExpressions.EOF And rsExpressions.BOF Then
		'              ExpressionID = 0
		'              msExpressionName = ""
		'            Else
		'              ExpressionID = lngOldExpressionID
		'              msExpressionName = rsExpressions!Name
		'            End If
		'
		'            rsExpressions.Close
		'            fOK = False
		'            fExit = True
		'        End Select
		'      Loop
		'    End With
		'
		'    Set frmSelectExpr = Nothing
		'  End If
		'
		'TidyUpAndExit:
		'  SelectExpression = fOK
		'  Exit Function
		'
		'ErrorTrap:
		'  fOK = False
		'  Resume TidyUpAndExit
		
	End Function
	
	
	Private Sub CopyExpression()
		'  ' Create a copy of the expression.
		'  On Error GoTo ErrorTrap
		'
		'  Dim fOK As Boolean
		'  Dim sName As String
		'  Dim frmEdit As frmExpression
		'
		'  fOK = False
		'
		'  ' Construct the expression to copy.
		'  Screen.MousePointer = vbHourglass
		'  fOK = ConstructExpression
		'  Screen.MousePointer = vbDefault
		'
		'  If fOK Then
		'    mlngExpressionID = 0
		'
		'    ' Initialise the copied expression's name.
		'    sName = msExpressionName
		'    msExpressionName = "Copy_of_" & Trim(sName)
		'
		'    'MH15062000
		'    msOwner = datGeneral.UserName
		'
		'    ' Display the expression edit form.
		'    Set frmEdit = New frmExpression
		'    Set frmEdit.Expression = Me
		'    frmEdit.Show vbModal
		'
		'    ' If the changes were confirmed then write the changes to the database.
		'    fOK = Not frmEdit.Cancelled
		'
		'    If fOK Then
		'      fOK = WriteExpression_Transaction
		'
		'
		'      'MH20000712
		'      Select Case miExpressionType
		'      Case giEXPR_RUNTIMECALCULATION
		'        Call UtilCreated(utlCalculation, Me.ExpressionID)
		'      Case giEXPR_RUNTIMEFILTER
		'        Call UtilCreated(utlFilter, Me.ExpressionID)
		'      End Select
		'
		'
		'    Else
		'      mfConstructed = False
		'    End If
		'  End If
		'
		'TidyUpAndExit:
		'  ' Commit the data transaction if everything was okay.
		'  Set frmEdit = Nothing
		'  Exit Sub
		'
		'ErrorTrap:
		'  fOK = False
		'  Resume TidyUpAndExit
		'
	End Sub
	
	Private Sub PrintExpression()
		
		'  ' Print the expression.
		'  On Error GoTo ErrorTrap
		'
		'  Dim fOK As Boolean
		'  Dim objComponent As clsExprComponent
		'  Dim objPrintDef As clsPrintDef
		'
		'  ' Construct the expression to print.
		'  Screen.MousePointer = vbHourglass
		'  fOK = ConstructExpression
		'  Screen.MousePointer = vbDefault
		'
		'  If fOK Then
		''    With Printer
		''      .Font = "Tahoma"
		''
		''      ' Print the header.
		''      .CurrentX = giPRINT_XINDENT
		''      .CurrentY = Int(giPRINT_YINDENT / 3)
		''      .FontSize = 8
		''      .FontBold = False
		''      Printer.Print Format(Date, "long date") & ", " & Format(Time, "medium time")
		''
		''      ' Print the title lines.
		''      .CurrentX = giPRINT_XINDENT
		''      .CurrentY = giPRINT_YINDENT
		''      .FontSize = 12
		''      .FontBold = True
		''      Printer.Print ExpressionTypeName(miExpressionType) & " Definition : " & Trim(msExpressionName) & vbNewLine
		''
		''      .FontSize = 10
		''
		''      .CurrentX = giPRINT_XINDENT
		''      Printer.Print "Description : " & Trim(msDescription) & vbNewLine
		''
		''      .CurrentX = giPRINT_XINDENT
		''      Printer.Print "Base Table : " & Trim(msBaseTableName) & vbNewLine
		''
		''      .CurrentX = giPRINT_XINDENT
		''      Printer.Print "Owner : " & Trim(msOwner) & vbNewLine
		''
		''      .CurrentX = giPRINT_XINDENT
		''
		''      'Printer.Print "Access : " & IIf(msAccess = giACCESS_HIDDEN, "Hidden", _
		'''        IIf(msAccess = giAccess_READONLY, "Read only", "Read / Write")) & vbNewLine
		''
		''      Select Case msAccess
		''      Case "RW": Printer.Print "Access : Read / Write" & vbNewLine
		''      Case "RO": Printer.Print "Access : Read only" & vbNewLine
		''      Case "HD": Printer.Print "Access : Hidden" & vbNewLine
		''      End Select
		''
		''      .CurrentX = giPRINT_XINDENT
		''      Printer.Print "Components : " & vbNewLine
		'
		'
		'    Set objPrintDef = New HrPro.clsPrintDef
		'
		'    If objPrintDef.IsOK Then
		'
		'      With objPrintDef
		'        .PrintHeader ExpressionTypeName(miExpressionType) & " Definition : " & Trim(msExpressionName)
		'
		'        .PrintNormal "Description : " & Trim(msDescription)
		'        .PrintNormal
		'
		'        .PrintNormal "Owner : " & Trim(msOwner)
		'        Select Case msAccess
		'        Case "RW": .PrintNormal "Access : Read / Write"
		'        Case "RO": .PrintNormal "Access : Read only"
		'        Case "HD": .PrintNormal "Access : Hidden"
		'        End Select
		'        .PrintNormal
		'
		'        .PrintNormal "Base Table : " & Trim(msBaseTableName)
		'        .PrintNormal
		'
		'        '--------
		'
		'        .PrintTitle "Components"
		'
		'
		'
		'        Printer.FontBold = False
		'
		'        ' Print the components.
		'        For Each objComponent In mcolComponents
		'          fOK = objComponent.PrintComponent(1)
		'          If Not fOK Then
		'            Printer.KillDoc
		'            Exit For
		'          End If
		'        Next
		'        Set objComponent = Nothing
		'
		'
		'
		'        '.EndDoc
		'        .PrintEnd
		'
		'        MsgBox ExpressionTypeName(miExpressionType) & " : " & Trim(msExpressionName) & " printing complete." & vbNewLine & vbNewLine & "(" & Printer.DeviceName & ")", vbInformation, ExpressionTypeName(miExpressionType)
		'
		'      End With
		'    End If
		'
		'
		'  End If
		'
		'TidyUpAndExit:
		'  Set objComponent = Nothing
		'  If Not fOK Then
		'    MsgBox "Unable to print the " & ExpressionTypeName(miExpressionType) & " '" & Name & "'." & vbCr & vbCr & _
		''      Err.Description, vbExclamation + vbOKOnly, App.ProductName
		'  End If
		'  Exit Sub
		'
		'ErrorTrap:
		'  fOK = False
		'  Resume TidyUpAndExit
		'
	End Sub
	
	Public Function PrintComponent(ByRef piLevel As Short) As Boolean
		'  ' Print the component definition to the printer object.
		'  On Error GoTo ErrorTrap
		'
		'  Dim fOK As Boolean
		'  Dim objComponent As clsExprComponent
		'
		'  fOK = True
		'
		'  With Printer
		'    .CurrentX = giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
		'    .CurrentY = .CurrentY + giPRINT_YSPACE
		'    Printer.Print "Parameter : " & ComponentDescription
		'  End With
		'
		'  ' Print the components.
		'  For Each objComponent In mcolComponents
		'    fOK = objComponent.PrintComponent(piLevel + 1)
		'    If Not fOK Then
		'      Printer.KillDoc
		'      Exit For
		'    End If
		'  Next
		'
		'TidyUpAndExit:
		'  Set objComponent = Nothing
		'  PrintComponent = fOK
		'  Exit Function
		'
		'ErrorTrap:
		'  fOK = False
		'  Resume TidyUpAndExit
		'
	End Function
	Public Function CopyComponent() As clsExprExpression
		'  ' Copies the selected component and all of it's children.
		'
		'  On Error GoTo ErrorTrap
		'
		'  Dim iCount As Integer
		'  Dim fOK As Boolean
		'  Dim objCopyComponent As New clsExprExpression
		'
		'    fOK = True
		'    objCopyComponent.ResetConstructedFlag (True)
		'    objCopyComponent.Name = msExpressionName
		'    objCopyComponent.BaseTableID = BaseTableID
		'    objCopyComponent.ExpressionType = ExpressionType
		'    objCopyComponent.ReturnType = ReturnType
		'
		'    ' Copy the children
		'    For iCount = 1 To mcolComponents.Count
		'        objCopyComponent.PasteComponent mcolComponents(iCount), mcolComponents(iCount), True
		'    Next iCount
		'
		'TidyUpAndExit:
		'  If fOK Then
		'    Set CopyComponent = objCopyComponent
		'  Else
		'    Set CopyComponent = Nothing
		'  End If
		'  Set objCopyComponent = Nothing
		'  Exit Function
		'
		'ErrorTrap:
		'  fOK = False
		'  Resume TidyUpAndExit
		
	End Function
	
	Public Function DeleteExpression() As Boolean
		'  ' Delete the expression.
		'  On Error GoTo ErrorTrap
		'
		'  Dim fOK As Boolean
		'  Dim fUsed As Boolean
		'  Dim fInTransaction As Boolean
		'  Dim sSQL As String
		'
		'  fUsed = False
		'  fInTransaction = False
		'
		'  ' Construct the expression from the database definition.
		'  Screen.MousePointer = vbHourglass
		'  fOK = ConstructExpression
		'  Screen.MousePointer = vbDefault
		'
		'  If fOK Then
		'    fInTransaction = True
		'    ' Begin the transaction of data.
		'    gADOCon.BeginTrans
		'
		'    ' Check that the expression can be deleted.
		'    ' ie. is not used anywhere.
		'    fUsed = ExpressionIsUsed
		'
		'    If Not fUsed Then
		'      ' Delete the expression's components.
		'      fOK = DeleteExistingComponents
		'
		'      ' Delete the expression itself.
		'      If fOK Then
		'        sSQL = "DELETE FROM ASRSysExpressions" & _
		''          " WHERE exprID = " & Trim(Str(mlngExpressionID))
		'        gADOCon.Execute sSQL, , adCmdText
		'      End If
		'    End If
		'  End If
		'
		'TidyUpAndExit:
		'  ' Commit the data transaction if everything was okay.
		'  If fOK And (Not fUsed) Then
		'    If fInTransaction Then
		'      gADOCon.CommitTrans
		'    End If
		'  Else
		'    If (Not fUsed) Then
		'      MsgBox "Error deleting the expression.", _
		''        vbOKOnly, App.ProductName
		'    End If
		'    If fInTransaction Then
		'      gADOCon.RollbackTrans
		'    End If
		'  End If
		'  DeleteExpression = fOK
		'  Exit Function
		'
		'ErrorTrap:
		'  fOK = False
		'  Resume TidyUpAndExit
		
	End Function
	
	
	Private Function ExpressionIsUsed() As Boolean
		'  ' Return true if the expression is used somewhere and
		'  ' therefore cannot be deleted.
		'  '
		'  ' Expressions may be used in the following contexts :
		'  '
		'  ' Table definitions - each table can have a record description defined.
		'  ' Column definitions - calculated columns require a calculation to be defined. All columns can have a validation check defined.
		'  ' Views - A filter can be defined for which records can be seen through a view.
		'  ' Expression definitions - expressions can refer to columns in child tables of the expression's base table. A filter for which child records are read can be defined. Expressions can also refer directly to other expressions.
		'  ' Cross-Tabs - A filter can be defined for which records are included in the cross-tab.
		'  ' Custom Reports - Filters can be defined for which records are included in the report, from the report's base table, and any parent or child tables that are also referred to. Custom reports can include calculations.
		'  ' Data Transfer - A filter can be defined for which records are included in the data transfer.
		'  ' Export - Filters can be defined for which records are included in the export, from the export's base table, and any parent or child tables that are also referred to. Exports can include calculations.
		'  ' Global Functions - A filter can be defined for which records are included in the global function. Global functions can include calculations.
		'  ' Mail Merge - A filter can be defined for which records are included in the mail merge. Mail merge definitions can include calculations.
		'  On Error GoTo ErrorTrap
		'
		'  Dim sExprName As String
		'  Dim sExprParentTable As String
		'  Dim sExprType As String
		'  Dim sGlobalFunctionType As String
		'  Dim objComp As clsExprComponent
		'  Dim lngRootExprID As Long
		'  Dim fUsed As Boolean
		'  Dim rsCheck As Recordset
		'  Dim sSQL As String
		'
		'  fUsed = False
		'
		'  ' Check that the expression is not used as a record description for a table.
		'  If Not fUsed Then
		'    sSQL = "SELECT ASRSysTables.tableName" & _
		''      " FROM ASRSysTables" & _
		''      " WHERE ASRSysTables.recordDescExprID = " & Trim(Str(mlngExpressionID))
		'    Set rsCheck = datGeneral.GetRecords(sSQL)
		'    With rsCheck
		'
		'      fUsed = Not (.EOF And .BOF)
		'
		'      If fUsed Then
		'        ' Tell the user why the expression cannot be deleted.
		'        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
		''          "It is used as the record description for the '" & .Fields("tableName") & "' table.", _
		''          vbExclamation + vbOKOnly, App.ProductName
		'      End If
		'
		'      .Close
		'    End With
		'    Set rsCheck = Nothing
		'  End If
		'
		'  If Not fUsed Then
		'    ' Find any columns that use this expression as a Calculation,
		'    ' a Got Focus clause, or a Lost Focus clause.
		'    sSQL = "SELECT ASRSysColumns.columnName, ASRSysTables.tableName" & _
		''      " FROM ASRSysColumns" & _
		''      " INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID" & _
		''      " WHERE (ASRSysColumns.calcExprID = " & Trim(Str(mlngExpressionID)) & _
		''      " OR ASRSysColumns.lostFocusExprID = " & Trim(Str(mlngExpressionID)) & ")"
		'    Set rsCheck = datGeneral.GetRecords(sSQL)
		'    With rsCheck
		'      fUsed = Not (.EOF And .BOF)
		'
		'      If fUsed Then
		'        ' Tell the user why the expression cannot be deleted.
		'        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
		''          "It is used by the '" & !ColumnName & "' column in the '" & !TableName & "' table.", _
		''          vbExclamation + vbOKOnly, App.ProductName
		'      End If
		'
		'      .Close
		'    End With
		'    Set rsCheck = Nothing
		'  End If
		'
		'  If Not fUsed Then
		'    ' Find any views that use this expression as the view filter.
		'    sSQL = "SELECT ASRSysViews.viewName, ASRSysTables.tableName" & _
		''      " FROM ASRSysViews" & _
		''      " INNER JOIN ASRSysTables ON ASRSysViews.viewTableID = ASRSysTables.tableID" & _
		''      " WHERE ASRSysViews.expressionID = " & Trim(Str(mlngExpressionID))
		'    Set rsCheck = datGeneral.GetRecords(sSQL)
		'    With rsCheck
		'      fUsed = Not (.EOF And .BOF)
		'
		'      If fUsed Then
		'        ' Tell the user why the expression cannot be deleted.
		'        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
		''          "It is used as the filter for the '" & !ViewName & "' view, based on the '" & !TableName & "' table.", _
		''          vbExclamation + vbOKOnly, App.ProductName
		'      End If
		'
		'      .Close
		'    End With
		'    Set rsCheck = Nothing
		'  End If
		'
		'  If Not fUsed Then
		'    ' Check that it is not used as a calculation in another expression,
		'    ' or as the filter in another expression.
		'    sSQL = "SELECT componentID" & _
		''      " FROM ASRSysExprComponents" & _
		''      " WHERE (calculationID = " & Trim(Str(mlngExpressionID)) & ")" & _
		''      " OR (fieldSelectionFilter = " & Trim(Str(mlngExpressionID)) & _
		''      "   AND type = " & Trim(Str(giCOMPONENT_FIELD)) & ")"
		'    Set rsCheck = datGeneral.GetRecords(sSQL)
		'    With rsCheck
		'      fUsed = Not (.EOF And .BOF)
		'
		'      If fUsed Then
		'        Set objComp = New clsExprComponent
		'        objComp.ComponentID = !ComponentID
		'        lngRootExprID = objComp.RootExpressionID
		'      End If
		'
		'      .Close
		'    End With
		'    Set rsCheck = Nothing
		'
		'    If fUsed Then
		'      sExprName = "<unknown>"
		'      sExprType = "<unknown>"
		'      sExprParentTable = "<unknown>"
		'
		'      ' Get the expression definition.
		'      sSQL = "SELECT ASRSysExpressions.name, ASRSysTables.tableName, ASRSysExpressions.type" & _
		''        " FROM ASRSysExpressions" & _
		''        " INNER JOIN ASRSysTables ON ASRSysExpressions.TableID = ASRSysTables.tableID" & _
		''        " WHERE ASRSysExpressions.exprID = " & Trim(Str(lngRootExprID))
		'      Set rsCheck = datGeneral.GetRecords(sSQL)
		'      With rsCheck
		'        If Not (.EOF And .BOF) Then
		'          sExprName = !Name
		'          sExprParentTable = !TableName
		'          sExprType = LCase(ExpressionTypeName(!Type))
		'        End If
		'
		'        .Close
		'      End With
		'      Set rsCheck = Nothing
		'
		'      ' Tell the user why the expression cannot be deleted.
		'      MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
		''        "It is used by the " & sExprType & " '" & sExprName & "'," & vbCr & _
		''        "which is owned by the '" & sExprParentTable & "' table.", _
		''        vbExclamation + vbOKOnly, App.ProductName
		'    End If
		'  End If
		'
		'  If Not fUsed Then
		'    ' Check that the expression is not used as a filter for a Cross-Tab.
		'    sSQL = "SELECT name" & _
		''      " FROM ASRSysCrossTab" & _
		''      " WHERE filterID = " & Trim(Str(mlngExpressionID))
		'    Set rsCheck = datGeneral.GetRecords(sSQL)
		'    With rsCheck
		'      fUsed = Not (.EOF And .BOF)
		'
		'      If fUsed Then
		'        ' Tell the user why the expression cannot be deleted.
		'        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
		''          "It is used as the filter for the '" & !Name & "' cross-tab.", _
		''          vbExclamation + vbOKOnly, App.ProductName
		'      End If
		'
		'      .Close
		'    End With
		'    Set rsCheck = Nothing
		'  End If
		'
		'  If Not fUsed Then
		'    ' Check that the expression is not used as a filter for a Custom Report.
		'    sSQL = "SELECT name" & _
		''      " FROM ASRSysCustomReportsName" & _
		''      " WHERE filter = " & Trim(Str(mlngExpressionID)) & _
		''      " OR parent1Filter = " & Trim(Str(mlngExpressionID)) & _
		''      " OR parent2Filter = " & Trim(Str(mlngExpressionID)) & _
		''      " OR childFilter = " & Trim(Str(mlngExpressionID))
		'
		'    Set rsCheck = datGeneral.GetRecords(sSQL)
		'    With rsCheck
		'      fUsed = Not (.EOF And .BOF)
		'
		'      If fUsed Then
		'        ' Tell the user why the expression cannot be deleted.
		'        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
		''          "It is used as a filter in the '" & !Name & "' custom report.", _
		''          vbExclamation + vbOKOnly, App.ProductName
		'      End If
		'
		'      .Close
		'    End With
		'    Set rsCheck = Nothing
		'  End If
		'
		'  If Not fUsed Then
		'    ' Check that the expression is not used as a calculation in a Custom Report.
		'    sSQL = "SELECT ASRSysCustomReportsName.name" & _
		''      " FROM ASRSysCustomReportsDetails" & _
		''      " INNER JOIN ASRSysCustomReportsName ON ASRSysCustomReportsDetails.customReportID = ASRSysCustomReportsName.ID" & _
		''      " WHERE colExprID = " & Trim(Str(mlngExpressionID)) & _
		''      " AND UPPER(type) = 'E'"
		'
		'    Set rsCheck = datGeneral.GetRecords(sSQL)
		'    With rsCheck
		'      fUsed = Not (.EOF And .BOF)
		'
		'      If fUsed Then
		'        ' Tell the user why the expression cannot be deleted.
		'        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
		''          "It is used as a calculation in the '" & !Name & "' custom report.", _
		''          vbExclamation + vbOKOnly, App.ProductName
		'      End If
		'
		'      .Close
		'    End With
		'    Set rsCheck = Nothing
		'  End If
		'
		'  If Not fUsed Then
		'    ' Check that the expression is not used as a filter for a Data Transfer.
		'    sSQL = "SELECT name" & _
		''      " FROM ASRSysDataTransferName" & _
		''      " WHERE filterID = " & Trim(Str(mlngExpressionID))
		'    Set rsCheck = datGeneral.GetRecords(sSQL)
		'    With rsCheck
		'      fUsed = Not (.EOF And .BOF)
		'
		'      If fUsed Then
		'        ' Tell the user why the expression cannot be deleted.
		'        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
		''          "It is used as the filter for the '" & !Name & "' data transfer.", _
		''          vbExclamation + vbOKOnly, App.ProductName
		'      End If
		'
		'      .Close
		'    End With
		'    Set rsCheck = Nothing
		'  End If
		'
		'  If Not fUsed Then
		'    ' Check that the expression is not used as a filter for an Export.
		'    sSQL = "SELECT name" & _
		''      " FROM ASRSysExportName" & _
		''      " WHERE filter = " & Trim(Str(mlngExpressionID)) & _
		''      " OR parent1Filter = " & Trim(Str(mlngExpressionID)) & _
		''      " OR parent2Filter = " & Trim(Str(mlngExpressionID)) & _
		''      " OR childFilter = " & Trim(Str(mlngExpressionID))
		'
		'    Set rsCheck = datGeneral.GetRecords(sSQL)
		'    With rsCheck
		'      fUsed = Not (.EOF And .BOF)
		'
		'      If fUsed Then
		'        ' Tell the user why the expression cannot be deleted.
		'        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
		''          "It is used as a filter in the '" & !Name & "' export.", _
		''          vbExclamation + vbOKOnly, App.ProductName
		'      End If
		'
		'      .Close
		'    End With
		'    Set rsCheck = Nothing
		'  End If
		'
		'  If Not fUsed Then
		'    ' Check that the expression is not used as a calculation in an Export.
		'    sSQL = "SELECT ASRSysExportName.name" & _
		''      " FROM ASRSysExportDetails" & _
		''      " INNER JOIN ASRSysExportName ON ASRSysExportDetails.exportID = ASRSysExportName.ID" & _
		''      " WHERE colExprID = " & Trim(Str(mlngExpressionID)) & _
		''      " AND UPPER(type) = 'E'"
		'
		'    Set rsCheck = datGeneral.GetRecords(sSQL)
		'    With rsCheck
		'      fUsed = Not (.EOF And .BOF)
		'
		'      If fUsed Then
		'        ' Tell the user why the expression cannot be deleted.
		'        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
		''          "It is used as a calculation in the '" & !Name & "' export.", _
		''          vbExclamation + vbOKOnly, App.ProductName
		'      End If
		'
		'      .Close
		'    End With
		'    Set rsCheck = Nothing
		'  End If
		'
		'  If Not fUsed Then
		'    ' Check that the expression is not used as a filter for a Global function.
		'    sSQL = "SELECT name, UPPER(type)" & _
		''      " FROM ASRSysGlobalFunctions" & _
		''      " WHERE filterID = " & Trim(Str(mlngExpressionID))
		'    Set rsCheck = datGeneral.GetRecords(sSQL)
		'    With rsCheck
		'      fUsed = Not (.EOF And .BOF)
		'
		'      If fUsed Then
		'        ' Tell the user why the expression cannot be deleted.
		'        Select Case !Type
		'          Case "U"
		'            sGlobalFunctionType = "update"
		'          Case "A"
		'            sGlobalFunctionType = "add"
		'          Case "D"
		'            sGlobalFunctionType = "delete"
		'          Case Else
		'            sGlobalFunctionType = "function"
		'        End Select
		'        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
		''          "It is used as the filter for the '" & !Name & "' global " & sGlobalFunctionType & ".", _
		''          vbExclamation + vbOKOnly, App.ProductName
		'      End If
		'
		'      .Close
		'    End With
		'    Set rsCheck = Nothing
		'  End If
		'
		'  ' Check that the expression is not used as a calculation for a Global function.
		'  If Not fUsed Then
		'    sSQL = "SELECT ASRSysGlobalFunctions.name, UPPER(ASRSysGlobalFunctions.type)" & _
		''      " FROM ASRSysGlobalItems" & _
		''      " INNER JOIN ASRSysGlobalFunctions ON ASRSysGlobalItems.functionID = ASRSysGlobalFunctions.functionID" & _
		''      " WHERE ASRSysGlobalItems.exprID = " & Trim(Str(mlngExpressionID)) & _
		''      " AND ASRSysGlobalItems.valueType = 4"
		'    Set rsCheck = datGeneral.GetRecords(sSQL)
		'    With rsCheck
		'      fUsed = Not (.EOF And .BOF)
		'
		'      If fUsed Then
		'        ' Tell the user why the expression cannot be deleted.
		'        Select Case !Type
		'          Case "U"
		'            sGlobalFunctionType = "update"
		'          Case "A"
		'            sGlobalFunctionType = "add"
		'          Case "D"
		'            sGlobalFunctionType = "delete"
		'          Case Else
		'            sGlobalFunctionType = "function"
		'        End Select
		'        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
		''          "It is used as a calculation in the '" & !Name & "' global " & sGlobalFunctionType & ".", _
		''          vbExclamation + vbOKOnly, App.ProductName
		'      End If
		'
		'      .Close
		'    End With
		'    Set rsCheck = Nothing
		'  End If
		'
		'  If Not fUsed Then
		'    ' Check that the expression is not used as a filter for a Mail Merge.
		'    sSQL = "SELECT name" & _
		''      " FROM ASRSysMailMergeName" & _
		''      " WHERE filterID = " & Trim(Str(mlngExpressionID))
		'    Set rsCheck = datGeneral.GetRecords(sSQL)
		'    With rsCheck
		'      fUsed = Not (.EOF And .BOF)
		'
		'      If fUsed Then
		'        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
		''          "It is used as the filter for the '" & !Name & "' mail merge.", _
		''          vbExclamation + vbOKOnly, App.ProductName
		'      End If
		'
		'      .Close
		'    End With
		'    Set rsCheck = Nothing
		'  End If
		'
		'  ' Check that the expression is not used as a calculation for a Mail Merge.
		'  If Not fUsed Then
		'    sSQL = "SELECT ASRSysMailMergeName.name" & _
		''      " FROM ASRSysMailMergeColumns" & _
		''      " INNER JOIN ASRSysMailMergeName ON ASRSysMailMergeColumns.mailMergeID = ASRSysMailMergeName.mailMergeID" & _
		''      " WHERE ASRSysMailMergeColumns.columnID = " & Trim(Str(mlngExpressionID)) & _
		''      " AND UPPER(ASRSysMailMergeColumns.type) = 'E'"
		'    Set rsCheck = datGeneral.GetRecords(sSQL)
		'    With rsCheck
		'      fUsed = Not (.EOF And .BOF)
		'
		'      If fUsed Then
		'        ' Tell the user why the expression cannot be deleted.
		'        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
		''          "It is used as a calculation in the '" & !Name & "' mail merge.", _
		''          vbExclamation + vbOKOnly, App.ProductName
		'      End If
		'
		'      .Close
		'    End With
		'    Set rsCheck = Nothing
		'  End If
		'
		'TidyUpAndExit:
		'  ' Disassociate object variables.
		'  Set rsCheck = Nothing
		'  Set objComp = Nothing
		'
		'  ExpressionIsUsed = fUsed
		'  Exit Function
		'
		'ErrorTrap:
		'  MsgBox "Error checking if the expression is used.", _
		''    vbOKOnly + vbExclamation, App.ProductName
		'  fUsed = True
		'  Resume TidyUpAndExit
		'
	End Function
	
	
	
	
	
	Public Function ValidityMessage(ByRef piInvalidityCode As Short) As Object
		' Return the text nmessage that describes the given
		' expression invalidity code.
		
		Select Case piInvalidityCode
			
			Case modExpression.ExprValidationCodes.giEXPRVALIDATION_NOERRORS
				'UPGRADE_WARNING: Couldn't resolve default property of object ValidityMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ValidityMessage = "No errors."
				
			Case modExpression.ExprValidationCodes.giEXPRVALIDATION_MISSINGOPERAND
				'UPGRADE_WARNING: Couldn't resolve default property of object ValidityMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ValidityMessage = "Missing operand."
				
			Case modExpression.ExprValidationCodes.giEXPRVALIDATION_SYNTAXERROR
				'UPGRADE_WARNING: Couldn't resolve default property of object ValidityMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ValidityMessage = "Syntax error."
				
			Case modExpression.ExprValidationCodes.giEXPRVALIDATION_EXPRTYPEMISMATCH
				'UPGRADE_WARNING: Couldn't resolve default property of object ValidityMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ValidityMessage = "Return type mismatch."
				
			Case modExpression.ExprValidationCodes.giEXPRVALIDATION_UNKNOWNERROR
				'UPGRADE_WARNING: Couldn't resolve default property of object ValidityMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ValidityMessage = "Unknown error."
				
			Case modExpression.ExprValidationCodes.giEXPRVALIDATION_OPERANDTYPEMISMATCH
				'UPGRADE_WARNING: Couldn't resolve default property of object ValidityMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ValidityMessage = "Operand type mismatch."
				
			Case modExpression.ExprValidationCodes.giEXPRVALIDATION_PARAMETERTYPEMISMATCH
				'UPGRADE_WARNING: Couldn't resolve default property of object ValidityMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ValidityMessage = "Parameter type mismatch."
				
			Case modExpression.ExprValidationCodes.giEXPRVALIDATION_NOCOMPONENTS
				'UPGRADE_WARNING: Couldn't resolve default property of object ValidityMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ValidityMessage = "The " & LCase(ExpressionTypeName(miExpressionType)) & " must have at least one component."
				
			Case modExpression.ExprValidationCodes.giEXPRVALIDATION_PARAMETERSYNTAXERROR
				'UPGRADE_WARNING: Couldn't resolve default property of object ValidityMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ValidityMessage = "Function parameter syntax error."
				
			Case modExpression.ExprValidationCodes.giEXPRVALIDATION_PARAMETERNOCOMPONENTS
				'UPGRADE_WARNING: Couldn't resolve default property of object ValidityMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ValidityMessage = "The function parameter expression must have at least one component."
				
			Case modExpression.ExprValidationCodes.giEXPRVALIDATION_FILTEREVALUATION
				' JPD20020419 Fault 3687
				'ValidityMessage = "Logic components must be compared with other logic components."
				'UPGRADE_WARNING: Couldn't resolve default property of object ValidityMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ValidityMessage = "Error creating SQL runtime code."
				
				' JPD20020419 Fault 3687
			Case modExpression.ExprValidationCodes.giEXPRVALIDATION_SQLERROR
				'UPGRADE_WARNING: Couldn't resolve default property of object ValidityMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ValidityMessage = "The complexity of the " & LCase(ExpressionTypeName(miExpressionType)) & " has caused the following SQL error : " & vbNewLine & vbNewLine & "'" & msErrorMessage & "'" & vbNewLine & vbNewLine & "Try simplifying the " & LCase(ExpressionTypeName(miExpressionType)) & "."
				
				' JPD20020419 Fault 3687
			Case modExpression.ExprValidationCodes.giEXPRVALIDATION_ASSOCSQLERROR
				'UPGRADE_WARNING: Couldn't resolve default property of object ValidityMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ValidityMessage = "The complexity of this " & LCase(ExpressionTypeName(miExpressionType)) & " would cause an expression that uses this " & LCase(ExpressionTypeName(miExpressionType)) & " to suffer from the following SQL error : " & vbNewLine & vbNewLine & "'" & msErrorMessage & "'" & vbNewLine & vbNewLine & "Try simplifying this " & LCase(ExpressionTypeName(miExpressionType)) & "."
				
				'JPD 20040507 Fault 8600
			Case modExpression.ExprValidationCodes.giEXPRVALIDATION_CYCLIC
				'UPGRADE_WARNING: Couldn't resolve default property of object ValidityMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ValidityMessage = "Invalid definition due to cyclic reference."
				
			Case Else
				'UPGRADE_WARNING: Couldn't resolve default property of object ValidityMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ValidityMessage = "The function parameter expression must have at least one component."
				
		End Select
		
	End Function
	
	Public Function ContainsExpression(ByRef plngExprID As Integer) As Boolean
		' Retrun TRUE if the current expression (or any of its sub expressions)
		' contains the given expression. This ensures no cyclic expressions get created.
		'JPD 20040507 Fault 8600
		On Error GoTo ErrorTrap
		
		Dim iLoop1 As Short
		
		ContainsExpression = False
		
		For iLoop1 = 1 To mcolComponents.Count()
			If ContainsExpression Then
				Exit For
			End If
			
			With mcolComponents.Item(iLoop1)
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().ContainsExpression. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ContainsExpression = .ContainsExpression(plngExprID)
			End With
		Next iLoop1
		
TidyUpAndExit: 
		Exit Function
		
ErrorTrap: 
		ContainsExpression = True
		Resume TidyUpAndExit
		
	End Function
	
	Public Sub NewExpression()
		'  ' Handle the definition of a new expression.
		'  On Error GoTo ErrorTrap
		'
		'  Dim fOK As Boolean
		'  Dim frmEdit As frmExpression
		'
		'  fOK = True
		'
		'  ' Initialize the properties for a new expression.
		'  InitialiseExpression
		'
		'  ' Display the expression definition form.
		'  Set frmEdit = New frmExpression
		'  With frmEdit
		'    Set .Expression = Me
		'    .Show vbModal
		'
		'    fOK = Not .Cancelled
		'  End With
		'
		'  If fOK Then
		'    ' Write the new expression to the database.
		'    fOK = WriteExpression_Transaction
		'
		'
		'    'MH20000712
		'    Select Case miExpressionType
		'    Case giEXPR_RUNTIMECALCULATION
		'      Call UtilCreated(utlCalculation, Me.ExpressionID)
		'    Case giEXPR_RUNTIMEFILTER
		'      Call UtilCreated(utlFilter, Me.ExpressionID)
		'    End Select
		'
		'
		'    ' If the write failed then re-initialize the
		'    ' properties for a new expression.
		'    If Not fOK Then
		'      InitialiseExpression
		'    End If
		'  End If
		'
		'TidyUpAndExit:
		'  ' Disassociate object variables.
		'  Set frmEdit = Nothing
		'  Exit Sub
		'
		'ErrorTrap:
		'  fOK = False
		'  Resume TidyUpAndExit
		'
	End Sub
	
	Public Function WriteExpression_Transaction() As Boolean
		' Transaction wrapper for the 'WriteExpression' function.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		
		' Begin the transaction of data.
    'gADOCon.BeginTrans()
		
		fOK = WriteExpression
		
TidyUpAndExit: 
		' Commit the data transaction if everything was okay.
    'If fOK Then
    '	gADOCon.CommitTrans()
    'Else
    '	gADOCon.RollbackTrans()
    'End If
		WriteExpression_Transaction = fOK
		Exit Function
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	
	Public Function WriteExpression() As Boolean
		'  ' Write the expression definition to the database.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		Dim sSQL As String
		Dim objComponent As clsExprComponent
		
		fOK = True
		
		If mlngExpressionID = 0 Then
			
			'MH20010712 Need keep manual record of allocated IDs incase users
			'in SYS MGR have created expressions but not yet saved changes
			'ExpressionID = UniqueColumnValue("ASRSysExpressions", "exprID")
			'ExpressionID = GetUniqueID("Expressions", "ASRSysExpressions", "exprID")
			
			'JPD20010911 Setting ExpressionID clears the mfConstructed flag, which we don't want.
			' So just set the mlngExpressionID variable. NB. The mfConstructed flag is only reset
			' when the code is stepped through, not when run without breakpoints. So no real
			' runtime error, but it just didn't make logical sense.
			mlngExpressionID = GetUniqueID("Expressions", "ASRSysExpressions", "exprID")
			
			' Add a record for the new expression.
			fOK = (mlngExpressionID > 0)
			
			If fOK Then
				sSQL = "INSERT INTO ASRSysExpressions" & " (exprID, name, TableID, returnType, returnSize, returnDecimals, " & " type, parentComponentID, Username, access, description, ViewInColour, ExpandedNode)" & " VALUES(" & Trim(Str(mlngExpressionID)) & ", " & "'" & Replace(Trim(msExpressionName), "'", "''") & "', " & Trim(Str(mlngBaseTableID)) & ", " & IIf(miExpressionType = modExpression.ExpressionTypes.giEXPR_RUNTIMECALCULATION, Trim(Str(modExpression.ExpressionValueTypes.giEXPRVALUE_UNDEFINED)), Trim(Str(miReturnType))) & ", " & "0,0, " & Trim(Str(miExpressionType)) & ", " & Trim(Str(mlngParentComponentID)) & ", " & "'" & Replace(Trim(msOwner), "'", "''") & "', " & "'" & Trim(msAccess) & "', " & "'" & Replace(Trim(msDescription), "'", "'") & "', " & IIf(mbViewInColour, "1, ", "0, ") & IIf(mbExpandedNode, "1", "0") & ")"
				gADOCon.Execute(sSQL,  , ADODB.CommandTypeEnum.adCmdText)
				
			End If
		Else
			sSQL = "UPDATE ASRSysExpressions" & " SET name = '" & Replace(Trim(msExpressionName), "'", "''") & "'," & " TableID = " & Trim(Str(mlngBaseTableID)) & "," & " returnType = " & IIf(miExpressionType = modExpression.ExpressionTypes.giEXPR_RUNTIMECALCULATION, Trim(Str(modExpression.ExpressionValueTypes.giEXPRVALUE_UNDEFINED)), Trim(Str(miReturnType))) & "," & " returnSize = 0," & " returnDecimals = 0," & " type = " & Trim(Str(miExpressionType)) & "," & " parentComponentID = " & Trim(Str(mlngParentComponentID)) & "," & " Username = '" & Replace(Trim(msOwner), "'", "''") & "'," & " access = '" & Trim(msAccess) & "'," & " description = '" & Replace(Trim(msDescription), "'", "''") & "', " & " ViewInColour = " & IIf(mbViewInColour, "1", "0") & " WHERE exprID = " & Trim(Str(mlngExpressionID))
			
			'" owner = '" & Trim(msOwner) & "'," & _
			'
			gADOCon.Execute(sSQL,  , ADODB.CommandTypeEnum.adCmdText)
		End If
		
		If fOK Then
			' Delete the expression's existing components from the database.
			fOK = DeleteExistingComponents
			
			If fOK Then
				' Add any components for this expression.
				For	Each objComponent In mcolComponents
					objComponent.ParentExpression = Me
					fOK = objComponent.WriteComponent
					
					If Not fOK Then
						Exit For
					End If
				Next objComponent
			End If
		End If
		
TidyUpAndExit: 
		'UPGRADE_NOTE: Object objComponent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objComponent = Nothing
		WriteExpression = fOK
		Exit Function
		
ErrorTrap: 
		'NO MSGBOX ON THE SERVER ! - MsgBox "Error saving the expression.", _
		'vbOKOnly + vbExclamation, App.ProductName
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	
	Public Function RuntimeCode(ByRef psRuntimeCode As String, ByRef palngSourceTables As Object, ByRef pfApplyPermissions As Boolean, ByRef pfValidating As Boolean, ByRef pavPromptedValues As Object, Optional ByRef plngFixedExprID As Integer = 0, Optional ByRef psFixedSQLCode As String = "") As Boolean
		' Return the SQL code that defines the expression.
		' Used when creating the 'where clause' for view definitions.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		Dim iLoop1 As Short
		Dim iLoop2 As Short
		Dim iLoop3 As Short
		Dim iParameter1Index As Short
		Dim iParameter2Index As Short
		Dim iMinOperatorPrecedence As Short
		Dim iMaxOperatorPrecedence As Short
		Dim sCode As String
		Dim sComponentCode As String
		Dim vParameter1 As Object
		Dim vParameter2 As Object
    Dim avValues(,) As Object
		
		fOK = True
		sCode = ""
		
		iMinOperatorPrecedence = -1
		iMaxOperatorPrecedence = -1
		
		' Create an array of the components in the expression.
		' Column 1 = operator id.
		' Column 2 = component where clause code.
		ReDim avValues(2, mcolComponents.Count())
		For iLoop1 = 1 To mcolComponents.Count()
			With mcolComponents.Item(iLoop1)
				' If the current component is an operator then read the operator id into the array.
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop1).ComponentType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If .ComponentType = modExpression.ExpressionComponentTypes.giCOMPONENT_OPERATOR Then
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object avValues(1, iLoop1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					avValues(1, iLoop1) = .Component.OperatorID
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop1).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					iMinOperatorPrecedence = IIf(iMinOperatorPrecedence > .Component.Precedence Or iMinOperatorPrecedence = -1, .Component.Precedence, iMinOperatorPrecedence)
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop1).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					iMaxOperatorPrecedence = IIf(iMaxOperatorPrecedence < .Component.Precedence Or iMaxOperatorPrecedence = -1, .Component.Precedence, iMaxOperatorPrecedence)
				End If
				
				' JPD20020419 Fault 3687
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().RuntimeCode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				fOK = .RuntimeCode(sComponentCode, palngSourceTables, pfApplyPermissions, pfValidating, pavPromptedValues, plngFixedExprID, psFixedSQLCode)
				
				If fOK Then
					'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					avValues(2, iLoop1) = sComponentCode
				End If
			End With
			
			If Not fOK Then
				Exit For
			End If
		Next iLoop1
		
		If fOK Then
			' Loop throught the expression's components checking that they are valid.
			' Evaluate operators in the correct order.
			For iLoop1 = iMinOperatorPrecedence To iMaxOperatorPrecedence
				For iLoop2 = 1 To mcolComponents.Count()
					With mcolComponents.Item(iLoop2)
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).ComponentType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If .ComponentType = modExpression.ExpressionComponentTypes.giCOMPONENT_OPERATOR Then
							'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If .Component.Precedence = iLoop1 Then
								' Check that the operator has the correct parameter types.
								' Read the value that follows the current operator.
								iParameter1Index = 0
								iParameter2Index = 0
								
								' Read the index of the first parameter.
								For iLoop3 = iLoop2 + 1 To UBound(avValues, 2)
									'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If avValues(2, iLoop3) <> vbNullString Then
										iParameter1Index = iLoop3
										Exit For
									End If
								Next iLoop3
								
								' If a parameter has been found then read its value.
								' Otherwise the expression is invalid.
								If iParameter1Index > 0 Then
									'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iParameter1Index). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object vParameter1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									vParameter1 = avValues(2, iParameter1Index)
								End If
								
								' Read a second parameter if required.
								'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If (.Component.OperandCount = 2) Then
									'UPGRADE_WARNING: Couldn't resolve default property of object vParameter1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object vParameter2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									vParameter2 = vParameter1
									iParameter2Index = iParameter1Index
									iParameter1Index = 0
									
									' Read the index of the parameter's value if there is one.
									For iLoop3 = iLoop2 - 1 To 1 Step -1
										'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										If avValues(2, iLoop3) <> vbNullString Then
											iParameter1Index = iLoop3
											Exit For
										End If
									Next iLoop3
									
									' If a parameter has been found then read its value.
									' Otherwise the expression is invalid.
									If iParameter1Index > 0 Then
										'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iParameter1Index). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object vParameter1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										vParameter1 = avValues(2, iParameter1Index)
										
										' JPD20020415 Fault 3662 - Need to cast values as float for division operators
										'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										If .Component.CastAsFloat Then
											'UPGRADE_WARNING: Couldn't resolve default property of object vParameter1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											vParameter1 = "Cast(" & vParameter1 & " As Float)"
										End If
										
									End If
									
									' Update the array to reflect the constructed SQL code.
									'UPGRADE_WARNING: Couldn't resolve default property of object avValues(1, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									avValues(1, iLoop2) = vbNullString
									'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If .Component.SQLType = "O" Then
										'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										If .Component.CheckDivideByZero Then
											' JPD20020415 Fault 3638
											'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											Select Case .Component.OperatorID
												Case 16 'Modulus
													'UPGRADE_WARNING: Couldn't resolve default property of object vParameter2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													'UPGRADE_WARNING: Couldn't resolve default property of object vParameter1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													avValues(2, iLoop2) = "(case when " & vParameter2 & " = 0 then 0 else (" & vbNewLine & vParameter1 & " - (CAST((" & vParameter1 & " / " & vParameter2 & ") AS INT) * " & vParameter2 & ")" & vbNewLine & ") end)"
												Case Else
													'UPGRADE_WARNING: Couldn't resolve default property of object vParameter2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													'UPGRADE_WARNING: Couldn't resolve default property of object avValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													'UPGRADE_WARNING: Couldn't resolve default property of object vParameter1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													avValues(2, iLoop2) = "(case when " & vParameter2 & " = 0 then 0 else (" & vbNewLine & vParameter1 & " " & vbNewLine & avValues(2, iLoop2) & " " & vbNewLine & vParameter2 & vbNewLine & ") end)"
											End Select
										Else
											'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											Select Case .Component.OperatorID
												Case 5 'And
													'UPGRADE_WARNING: Couldn't resolve default property of object vParameter2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													'UPGRADE_WARNING: Couldn't resolve default property of object vParameter1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													avValues(2, iLoop2) = "(CASE WHEN (" & vParameter1 & " = 1) AND (" & vParameter2 & " = 1) THEN 1 ELSE 0 END)"
												Case 6 'Or
													'UPGRADE_WARNING: Couldn't resolve default property of object vParameter2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													'UPGRADE_WARNING: Couldn't resolve default property of object vParameter1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													avValues(2, iLoop2) = "(CASE WHEN (" & vParameter1 & " = 1) OR (" & vParameter2 & " = 1) THEN 1 ELSE 0 END)"
												Case Else
													'UPGRADE_WARNING: Couldn't resolve default property of object vParameter2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													'UPGRADE_WARNING: Couldn't resolve default property of object avValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													'UPGRADE_WARNING: Couldn't resolve default property of object vParameter1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													avValues(2, iLoop2) = "(" & vbNewLine & vParameter1 & " " & vbNewLine & avValues(2, iLoop2) & " " & vbNewLine & vParameter2 & vbNewLine & ")"
											End Select
										End If
									Else
										'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object vParameter2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object vParameter1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										avValues(2, iLoop2) = IIf(Len(.Component.SQLFixedParam1) > 0, "(", "") & avValues(2, iLoop2) & vbNewLine & "(" & vbNewLine & vParameter1 & vbNewLine & ", " & vbNewLine & vParameter2 & vbNewLine & ")" & IIf(Len(.Component.SQLFixedParam1) > 0, " " & .Component.SQLFixedParam1 & ")", "")
									End If
									'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iParameter1Index). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									avValues(2, iParameter1Index) = vbNullString
									'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iParameter2Index). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									avValues(2, iParameter2Index) = vbNullString
								Else
									' Update the array to reflect the constructed SQL code.
									'UPGRADE_WARNING: Couldn't resolve default property of object avValues(1, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									avValues(1, iLoop2) = vbNullString
									'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If .Component.SQLType = "O" Then
										'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										Select Case .Component.OperatorID
											Case 13 'Not
												'UPGRADE_WARNING: Couldn't resolve default property of object vParameter1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												avValues(2, iLoop2) = "(CASE WHEN " & vParameter1 & " = 1 THEN 0 ELSE 1 END)"
											Case Else
												'UPGRADE_WARNING: Couldn't resolve default property of object vParameter1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												'UPGRADE_WARNING: Couldn't resolve default property of object avValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												avValues(2, iLoop2) = "(" & vbNewLine & avValues(2, iLoop2) & " " & vbNewLine & vParameter1 & vbNewLine & ")"
										End Select
									Else
										'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object vParameter1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										avValues(2, iLoop2) = IIf(Len(.Component.SQLFixedParam1) > 0, "(", "") & avValues(2, iLoop2) & vbNewLine & "(" & vbNewLine & vParameter1 & vbNewLine & ")" & IIf(Len(.Component.SQLFixedParam1) > 0, " " & .Component.SQLFixedParam1 & ")", "")
									End If
									'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iParameter1Index). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									avValues(2, iParameter1Index) = vbNullString
								End If
								
								If (miExpressionType = modExpression.ExpressionTypes.giEXPR_RUNTIMECALCULATION) Or (miExpressionType = modExpression.ExpressionTypes.giEXPR_RUNTIMEFILTER) Or (miExpressionType = modExpression.ExpressionTypes.giEXPR_LINKFILTER) Then
									'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If (.Component.ReturnType = modExpression.ExpressionValueTypes.giEXPRVALUE_LOGIC) And ((.Component.OperatorID <> 5) And (.Component.OperatorID <> 6) And (.Component.OperatorID <> 13)) Then
										'UPGRADE_WARNING: Couldn't resolve default property of object avValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										avValues(2, iLoop2) = "(CASE WHEN (" & avValues(2, iLoop2) & ") THEN 1 ELSE 0 END)"
									End If
								End If
							End If
						End If
					End With
				Next iLoop2
			Next iLoop1
			
			For iLoop1 = 1 To UBound(avValues, 2)
				'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If avValues(2, iLoop1) <> vbNullString Then
					'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sCode = avValues(2, iLoop1)
					Exit For
				End If
			Next iLoop1
			
		End If
		
TidyUpAndExit: 
		If fOK Then
			psRuntimeCode = sCode
		Else
			psRuntimeCode = ""
		End If
		RuntimeCode = fOK
		
		Exit Function
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	
	
	
	Public Function RuntimeFilterCode(ByRef psFilterCode As String, ByRef pfApplyPermissions As Boolean, Optional ByRef pfValidating As Boolean = False, Optional ByRef pavPromptedValues As Object = Nothing, Optional ByRef plngFixedExprID As Integer = 0, Optional ByRef psFixedSQLCode As String = "") As Boolean
		' Return TRUE if the filter code was created okay.
		' Return the runtime filter SQL code in the parameter 'psFilterCode'.
		' Apply permissions to the filter code only if the 'pfApplyPermissions' parameter is TRUE.
		' The filter code is to be used to validate the expression if the 'pfValidating' parameter is TRUE.
		' This is used to suppress prompting the user for promted values, when we are only validating the expression.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		Dim iLoop1 As Short
		Dim iLoop2 As Short
		Dim iNextIndex As Short
		Dim sSQL As String
		Dim sWhereCode As String
		Dim sBaseTableSource As String
		Dim sRuntimeFilterSQL As String
    Dim alngSourceTables(,) As Integer
    Dim avRelatedTables(,) As Object
		Dim rsInfo As ADODB.Recordset
		Dim objTableView As CTablePrivilege
		
		' Check if the 'validating' parameter is set.
		' If not, set it to FALSE.
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(pfValidating) Then
			pfValidating = False
		End If
		
		' Construct the expression from the database definition.
		fOK = ConstructExpression
		
		If fOK Then
			sBaseTableSource = msBaseTableName
			If pfApplyPermissions Then
				' Get the 'realSource' of the table.
				objTableView = gcoTablePrivileges.Item(msBaseTableName)
				If objTableView.TableType = Declarations.TableTypes.tabChild Then
					sBaseTableSource = objTableView.RealSource
				End If
				'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				objTableView = Nothing
			End If
			
			sRuntimeFilterSQL = "SELECT DISTINCT " & sBaseTableSource & ".id" & vbNewLine & "FROM " & sBaseTableSource & " " & vbNewLine
			
			' Create an array of the IDs of the tables/view referred to in the expression.
			' This is used for joining all of the tables/views used.
			' Column 1 = 0 if this row is for a table, 1 if it is for a view.
			' Column 2 = table/view ID.
			ReDim alngSourceTables(2, 0)
			
			' Get the filter code.
			' JPD20020419 Fault 3687
			fOK = RuntimeCode(sWhereCode, alngSourceTables, pfApplyPermissions, pfValidating, pavPromptedValues, plngFixedExprID, psFixedSQLCode)
		End If
		
		If fOK Then
			' Create an array of the tables related to the expression base table.
			' Used when Joining any other tables/view used.
			' Column 1 = 'parent' if the expression's base table is the parent of the other table
			'            'child' if the expression's base table is the child of the other table
			' Column 2 = ID of the other table
			ReDim avRelatedTables(2, 0)
			sSQL = "SELECT 'parent' AS relationship, childID AS tableID" & " FROM ASRSysRelations" & " WHERE parentID = " & Trim(Str(mlngBaseTableID)) & " UNION" & " SELECT 'child' AS relationship, parentID AS tableID" & " FROM ASRSysRelations" & " WHERE childID = " & Trim(Str(mlngBaseTableID))
			rsInfo = datGeneral.GetRecords(sSQL)
			With rsInfo
				Do While Not .EOF
					iNextIndex = UBound(avRelatedTables, 2) + 1
					ReDim Preserve avRelatedTables(2, iNextIndex)
					'UPGRADE_WARNING: Couldn't resolve default property of object avRelatedTables(1, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					avRelatedTables(1, iNextIndex) = .Fields("relationship").Value
					'UPGRADE_WARNING: Couldn't resolve default property of object avRelatedTables(2, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					avRelatedTables(2, iNextIndex) = .Fields("TableID").Value
					.MoveNext()
				Loop 
				.Close()
			End With
			'UPGRADE_NOTE: Object rsInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsInfo = Nothing
			
			' Join any other tables/views used.
			For iLoop1 = 1 To UBound(alngSourceTables, 2)
				If alngSourceTables(1, iLoop1) = 0 Then
					objTableView = gcoTablePrivileges.FindTableID(alngSourceTables(2, iLoop1))
				Else
					objTableView = gcoTablePrivileges.FindViewID(alngSourceTables(2, iLoop1))
				End If
				
				If objTableView.TableID = mlngBaseTableID Then
					' Join a view on the base table.
					If Not pfApplyPermissions Then
						sRuntimeFilterSQL = sRuntimeFilterSQL & "LEFT OUTER JOIN " & objTableView.TableName & " ON " & sBaseTableSource & ".id = " & objTableView.TableName & ".id" & vbNewLine
					Else
						sRuntimeFilterSQL = sRuntimeFilterSQL & "LEFT OUTER JOIN " & objTableView.RealSource & " ON " & sBaseTableSource & ".id = " & objTableView.RealSource & ".id" & vbNewLine
					End If
				Else
					' Join a table/view on a parent/child related to the base table.
					For iLoop2 = 1 To UBound(avRelatedTables, 2)
						'UPGRADE_WARNING: Couldn't resolve default property of object avRelatedTables(2, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If avRelatedTables(2, iLoop2) = objTableView.TableID Then
							
							If Not pfApplyPermissions Then
								'UPGRADE_WARNING: Couldn't resolve default property of object avRelatedTables(1, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If avRelatedTables(1, iLoop2) = "parent" Then
									sRuntimeFilterSQL = sRuntimeFilterSQL & "LEFT OUTER JOIN " & objTableView.TableName & " ON " & sBaseTableSource & ".id = " & objTableView.TableName & ".id_" & Trim(Str(mlngBaseTableID)) & " " & vbNewLine
								Else
									'UPGRADE_WARNING: Couldn't resolve default property of object avRelatedTables(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									sRuntimeFilterSQL = sRuntimeFilterSQL & "LEFT OUTER JOIN " & objTableView.TableName & " ON " & sBaseTableSource & ".id_" & Trim(Str(avRelatedTables(2, iLoop2))) & " = " & objTableView.TableName & ".id " & vbNewLine
								End If
							Else
								'UPGRADE_WARNING: Couldn't resolve default property of object avRelatedTables(1, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If avRelatedTables(1, iLoop2) = "parent" Then
									sRuntimeFilterSQL = sRuntimeFilterSQL & "LEFT OUTER JOIN " & objTableView.RealSource & " ON " & sBaseTableSource & ".id = " & objTableView.RealSource & ".id_" & Trim(Str(mlngBaseTableID)) & " " & vbNewLine
								Else
									'UPGRADE_WARNING: Couldn't resolve default property of object avRelatedTables(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									sRuntimeFilterSQL = sRuntimeFilterSQL & "LEFT OUTER JOIN " & objTableView.RealSource & " ON " & sBaseTableSource & ".id_" & Trim(Str(avRelatedTables(2, iLoop2))) & " = " & objTableView.RealSource & ".id " & vbNewLine
								End If
							End If
							
							Exit For
						End If
					Next iLoop2
				End If
				
				'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				objTableView = Nothing
			Next iLoop1
			
			' Add the filter 'where' clause code.
			If Len(sWhereCode) > 0 Then
				sWhereCode = sWhereCode & " = 1"
				
				sRuntimeFilterSQL = sRuntimeFilterSQL & "WHERE " & vbNewLine & sWhereCode & vbNewLine
			End If
		End If
		
TidyUpAndExit: 
		If fOK Then
			psFilterCode = sRuntimeFilterSQL
		Else
			psFilterCode = ""
		End If
		RuntimeFilterCode = fOK
		
		Exit Function
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	
  Friend Function RuntimeCalculationCode(ByRef palngSourceTables As Object, ByRef psCalcCode As String, ByRef pfApplyPermissions As Boolean _
                                         , Optional ByRef pfValidating As Boolean = False, Optional ByRef pavPromptedValues As Object = Nothing _
                                         , Optional ByRef plngFixedExprID As Integer = 0, Optional ByRef psFixedSQLCode As String = "") As Boolean
    ' Return TRUE if the Calculation code was created okay.
    ' Return the runtime Calculation SQL code in the parameter 'psCalcCode'.
    ' Apply permissions to the Calculation code only if the 'pfApplyPermissions' parameter is TRUE.
    On Error GoTo ErrorTrap

    Dim fOK As Boolean
    Dim sRuntimeSQL As String
    Dim avDummyPrompts(,) As Object

    ' Check if the 'validating' parameter is set.
    ' If not, set it to FALSE.
    'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
    If IsNothing(pfValidating) Then
      pfValidating = False
    End If

    ' Construct the expression from the database definition.
    fOK = ConstructExpression()

    If fOK Then
      ' Get the Calculation code.
      'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
      If IsNothing(pavPromptedValues) Then
        ReDim avDummyPrompts(1, 0)
        fOK = RuntimeCode(sRuntimeSQL, palngSourceTables, pfApplyPermissions, pfValidating, avDummyPrompts, plngFixedExprID, psFixedSQLCode)
      Else
        fOK = RuntimeCode(sRuntimeSQL, palngSourceTables, pfApplyPermissions, pfValidating, pavPromptedValues, plngFixedExprID, psFixedSQLCode)
      End If
    End If

    If fOK Then
      If pfApplyPermissions Then
        fOK = (ValidateExpression(True) = modExpression.ExprValidationCodes.giEXPRVALIDATION_NOERRORS)
      End If
    End If

    If fOK And (miReturnType = modExpression.ExpressionValueTypes.giEXPRVALUE_LOGIC) Then
      sRuntimeSQL = "convert(bit, " & sRuntimeSQL & ")"
    End If

TidyUpAndExit:
    If fOK Then
      psCalcCode = sRuntimeSQL
    Else
      psCalcCode = ""
    End If
    RuntimeCalculationCode = fOK

    Exit Function

ErrorTrap:
    fOK = False
    Resume TidyUpAndExit

  End Function
	
  Friend Function DeleteExistingComponents() As Boolean
    ' Delete the expression's components and sub-expression's
    ' (ie. function parameter expressions) from the database.
    On Error GoTo ErrorTrap

    Dim fOK As Boolean
    Dim sSQL As String
    Dim sDeletedExpressionIDs As String
    Dim rsSubExpressions As ADODB.Recordset
    Dim objExpr As clsExprExpression

    fOK = True
    sDeletedExpressionIDs = ""

    ' Get the expression's function components from the database.
    sSQL = "SELECT ASRSysExpressions.exprID" & " FROM ASRSysExpressions" & " INNER JOIN ASRSysExprComponents" & "   ON ASRSysExpressions.parentComponentID = ASRSysExprComponents.componentID" & " AND ASRSysExprComponents.exprID = " & Trim(Str(mlngExpressionID))
    rsSubExpressions = datGeneral.GetRecordsInTransaction(sSQL)
    With rsSubExpressions
      Do While (Not .EOF) And fOK
        ' Instantiate each function parameter expression.
        ' Instruct the function parameter expression to delete its components.
        objExpr = New clsExprExpression
        objExpr.ExpressionID = .Fields("ExprID").Value
        fOK = objExpr.DeleteExistingComponents
        'UPGRADE_NOTE: Object objExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objExpr = Nothing

        ' Add the ID of the sub-expression to the string of sub-expressions to be deleted.
        sDeletedExpressionIDs = sDeletedExpressionIDs & IIf(Len(sDeletedExpressionIDs) > 0, ", ", "") & Trim(Str(.Fields("ExprID").Value))

        .MoveNext()
      Loop

      .Close()
    End With
    'UPGRADE_NOTE: Object rsSubExpressions may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsSubExpressions = Nothing

    If Len(sDeletedExpressionIDs) > 0 Then
      ' Delete all existing sub-expressions for this expression from the database.
      sSQL = "DELETE FROM ASRSysExpressions" & " WHERE exprID IN (" & sDeletedExpressionIDs & ")"
      gADOCon.Execute(sSQL, , ADODB.CommandTypeEnum.adCmdText)
    End If

    ' Delete all existing components for this expression from the database.
    sSQL = "DELETE FROM ASRSysExprComponents" & " WHERE exprID = " & Trim(Str(mlngExpressionID))
    gADOCon.Execute(sSQL, , ADODB.CommandTypeEnum.adCmdText)

TidyUpAndExit:
    'UPGRADE_NOTE: Object objExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    objExpr = Nothing

    'UPGRADE_NOTE: Object rsSubExpressions may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsSubExpressions = Nothing
    DeleteExistingComponents = fOK
    Exit Function

ErrorTrap:
    fOK = False
    Resume TidyUpAndExit

  End Function
	
  Public Function ValidateExpression(ByRef pfTopLevel As Boolean) As Short
    ' Validate the expression. Return a code defining the validity of the expression.
    ' NB. This function is also good for evaluating the return type of an expression
    ' which has definite return type (eg. function sub-expressions, runtime calcs, etc).
    On Error GoTo ErrorTrap

    Dim iLoop1 As Short
    Dim iLoop2 As Short
    Dim iLoop3 As Short
    Dim iParam1Type As Short
    Dim iParam2Type As Short
    Dim iParameter1Index As Short
    Dim iParameter2Index As Short
    Dim iParam1ReturnType As Short
    Dim iParam2ReturnType As Short
    Dim iOperatorReturnType As modExpression.ExpressionValueTypes
    Dim iBadLogicColumnIndex As Short
    Dim iMinOperatorPrecedence As Short
    Dim iMaxOperatorPrecedence As Short
    Dim iValidationCode As modExpression.ExprValidationCodes
    Dim iEvaluatedReturnType As modExpression.ExpressionValueTypes
    Dim aiDummyValues(,) As Short
    Dim avDummyPrompts(,) As Object
    Dim iTempReturnType As Short

    ReDim avDummyPrompts(1, 0)

    ' Initialise variables.
    iBadLogicColumnIndex = 0
    iMinOperatorPrecedence = -1
    iMaxOperatorPrecedence = -1
    iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_NOERRORS
    'UPGRADE_NOTE: Object mobjBadComponent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mobjBadComponent = Nothing

    ' If there are no expression components then report the error.
    If mcolComponents.Count() = 0 Then
      iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_NOCOMPONENTS
    End If

    ' Create an array of the component return types and operator ids.
    ' Index 1 = operator id.
    ' Index 2 = return type.
    '
    ' Eg. the expression
    ' 'abc'
    ' Concatenated with
    ' Function 'uppercase'
    '   <parameter>
    '      Field 'personnel.surname'
    '
    ' will be represented in the array as
    ' null,  giEXPRVALUE_CHARACTER
    '   17,  giEXPRVALUE_CHARACTER
    ' null,  giEXPRVALUE_CHARACTER
    '
    ' The operators are then evaluated to leave the array as :
    ' null,  null
    ' null,  giEXPRVALUE_CHARACTER
    ' null,  null
    '
    ' The one remaining value in the second column, after all operators have been evaluated
    ' is the return type of the expression.
    ReDim aiDummyValues(2, mcolComponents.Count())

    For iLoop1 = 1 To mcolComponents.Count()
      ' Stop validating the expression if we already know its invalid.
      If iValidationCode <> modExpression.ExprValidationCodes.giEXPRVALIDATION_NOERRORS Then
        Exit For
      End If

      With mcolComponents.Item(iLoop1)
        ' If the current component is an operator then read the operator id into the array.
        'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop1).ComponentType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If .ComponentType = modExpression.ExpressionComponentTypes.giCOMPONENT_OPERATOR Then
          'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          aiDummyValues(1, iLoop1) = .Component.OperatorID

          ' Remember the min and max operator precedence levels for later.
          'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop1).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          iMinOperatorPrecedence = IIf((iMinOperatorPrecedence > .Component.Precedence) Or (iMinOperatorPrecedence = -1), .Component.Precedence, iMinOperatorPrecedence)
          'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop1).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          iMaxOperatorPrecedence = IIf((iMaxOperatorPrecedence < .Component.Precedence) Or (iMaxOperatorPrecedence = -1), .Component.Precedence, iMaxOperatorPrecedence)
        Else
          aiDummyValues(1, iLoop1) = -1
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop1).ComponentType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If .ComponentType = modExpression.ExpressionComponentTypes.giCOMPONENT_FUNCTION Then
          ' Validate the function.
          ' NB. This also determines the function's return type if not already known.
          'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          iValidationCode = .Component.ValidateFunction
          If iValidationCode <> modExpression.ExprValidationCodes.giEXPRVALIDATION_NOERRORS Then
            'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop1).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If .Component.BadComponent Is Nothing Then
              mobjBadComponent = mcolComponents.Item(iLoop1)
            Else
              'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              mobjBadComponent = .Component.BadComponent
            End If
            Exit For
          End If
        End If

        ' Read the component return type into the array.
        'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().ReturnType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aiDummyValues(2, iLoop1) = .ReturnType
      End With
    Next iLoop1

    ' Loop throught the expression's components checking that they are valid.
    ' Evaluate operators in the correct order of precedence.
    For iLoop1 = iMinOperatorPrecedence To iMaxOperatorPrecedence
      ' Stop validating the expression if we already know it's invalid.
      If iValidationCode <> modExpression.ExprValidationCodes.giEXPRVALIDATION_NOERRORS Then
        Exit For
      End If

      For iLoop2 = 1 To mcolComponents.Count()
        ' Stop validating the expression if we already know it's invalid.
        If iValidationCode <> modExpression.ExprValidationCodes.giEXPRVALIDATION_NOERRORS Then
          Exit For
        End If

        With mcolComponents.Item(iLoop2)
          'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).ComponentType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          If .ComponentType = modExpression.ExpressionComponentTypes.giCOMPONENT_OPERATOR Then
            'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If .Component.Precedence = iLoop1 Then
              ' Check that the operator has the correct parameter types.
              ' Read the dummy value that follows the current operator.
              iParameter1Index = 0
              iParameter2Index = 0
              For iLoop3 = iLoop2 + 1 To UBound(aiDummyValues, 2)
                ' If an operator follows the operator then the expression is invalid.
                If aiDummyValues(1, iLoop3) > 0 Then
                  mobjBadComponent = mcolComponents.Item(iLoop2)
                  iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_MISSINGOPERAND
                  Exit For
                End If

                ' Read the index of the parameter.
                If aiDummyValues(2, iLoop3) > -1 Then
                  iParameter1Index = iLoop3
                  Exit For
                End If
              Next iLoop3

              ' If a parameter has been found then read its dummy value.
              ' Otherwise the expression is invalid.
              If iParameter1Index = 0 Then
                mobjBadComponent = mcolComponents.Item(iLoop2)
                iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_MISSINGOPERAND
              End If

              ' Read a second parameter if required.
              'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              If (.Component.OperandCount = 2) And (iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_NOERRORS) Then

                iParameter2Index = iParameter1Index

                ' Read the dummy value that precedes the current operator.
                iParameter1Index = 0
                For iLoop3 = iLoop2 - 1 To 1 Step -1
                  ' If an operator follows the operator then the expression is invalid.
                  If aiDummyValues(1, iLoop3) > 0 Then
                    mobjBadComponent = mcolComponents.Item(iLoop2)
                    iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_MISSINGOPERAND
                    Exit For
                  End If

                  ' Read the index of the parameter.
                  If aiDummyValues(2, iLoop3) > -1 Then
                    iParameter1Index = iLoop3
                    Exit For
                  End If
                Next iLoop3

                ' If a parameter has been found then read its dummy value.
                ' Otherwise the expression is invalid.
                If iParameter1Index = 0 Then
                  mobjBadComponent = mcolComponents.Item(iLoop2)
                  iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_MISSINGOPERAND
                End If
              End If

              ' Validate the operator by evaluating it with the dummy parmameters.
              ' NB. This also determines the operator's return type if not already known.
              ' Only try to evaluate the dummy operation if we still think
              ' it is valid.
              If iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_NOERRORS Then
                'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If Not ValidateOperatorParameters(.Component.OperatorID, iOperatorReturnType, aiDummyValues(2, iParameter1Index), IIf(.Component.OperandCount = 2, aiDummyValues(2, iParameter2Index), modExpression.ExpressionValueTypes.giEXPRVALUE_UNDEFINED)) Then

                  iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_OPERANDTYPEMISMATCH
                  mobjBadComponent = mcolComponents.Item(iLoop2)
                Else
                  ' Check that operators with logic parameters are valid.
                  If (iBadLogicColumnIndex = 0) And (iParameter2Index > 0) Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().ComponentType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    iParam1Type = mcolComponents.Item(iParameter1Index).ComponentType
                    'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().ReturnType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    iParam1ReturnType = mcolComponents.Item(iParameter1Index).ReturnType
                    'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().ComponentType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    iParam2Type = mcolComponents.Item(iParameter2Index).ComponentType
                    'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().ReturnType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    iParam2ReturnType = mcolComponents.Item(iParameter2Index).ReturnType

                    If ((iParam1Type = modExpression.ExpressionComponentTypes.giCOMPONENT_FIELD) And (iParam1ReturnType = modExpression.ExpressionValueTypes.giEXPRVALUE_LOGIC)) And (((iParam2Type <> modExpression.ExpressionComponentTypes.giCOMPONENT_FIELD) And (iParam2Type <> modExpression.ExpressionComponentTypes.giCOMPONENT_VALUE)) Or (iParam2ReturnType <> modExpression.ExpressionValueTypes.giEXPRVALUE_LOGIC)) Then

                      iBadLogicColumnIndex = iParameter1Index
                    End If

                    If ((iParam2Type = modExpression.ExpressionComponentTypes.giCOMPONENT_FIELD) And (iParam2ReturnType = modExpression.ExpressionValueTypes.giEXPRVALUE_LOGIC)) And (((iParam1Type <> modExpression.ExpressionComponentTypes.giCOMPONENT_FIELD) And (iParam1Type <> modExpression.ExpressionComponentTypes.giCOMPONENT_VALUE)) Or (iParam1ReturnType <> modExpression.ExpressionValueTypes.giEXPRVALUE_LOGIC)) Then

                      iBadLogicColumnIndex = iParameter2Index
                    End If
                  End If

                  ' Update the array to reflect the evaluated operation.
                  aiDummyValues(1, iLoop2) = -1
                  aiDummyValues(2, iParameter1Index) = -1
                  aiDummyValues(2, iParameter2Index) = -1
                End If
              End If
            End If
          End If
        End With
      Next iLoop2
    Next iLoop1

    ' Check the expression has valid syntax (ie. if the components have evaluated to a single value).
    ' Get the evaluated return type while we're at it.
    If iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_NOERRORS Then
      iEvaluatedReturnType = modExpression.ExpressionValueTypes.giEXPRVALUE_UNDEFINED

      For iLoop1 = 1 To UBound(aiDummyValues, 2)
        If aiDummyValues(2, iLoop1) > -1 Then
          ' If the expression has more than one component after evaluating
          ' all of the operators then the expression is invalid.
          If iEvaluatedReturnType <> modExpression.ExpressionValueTypes.giEXPRVALUE_UNDEFINED Then
            iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_SYNTAXERROR
            Exit For
          End If

          iEvaluatedReturnType = aiDummyValues(2, iLoop1)
        End If
      Next iLoop1
    End If

    ' Set the expression's return type if it is not already set.
    If iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_NOERRORS Then
      If (miReturnType = modExpression.ExpressionValueTypes.giEXPRVALUE_UNDEFINED) Or (miReturnType = modExpression.ExpressionValueTypes.giEXPRVALUE_BYREF_UNDEFINED) Then
        miReturnType = iEvaluatedReturnType
      End If
    End If

    ' Check the evaluated return type matches the pre-set return type.
    If (iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_NOERRORS) And (iEvaluatedReturnType <> miReturnType) Then
      iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_EXPRTYPEMISMATCH
    End If

    ' JPD20020419 Fault 3687
    ' Run the filter's SQL code to see if it is valid.
    If (iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_NOERRORS) And pfTopLevel Then
      iTempReturnType = miReturnType
      iValidationCode = ValidateSQLCode()
      miReturnType = iTempReturnType
    End If

TidyUpAndExit:
    ValidateExpression = iValidationCode
    Exit Function

ErrorTrap:
    iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_UNKNOWNERROR
    Resume TidyUpAndExit

  End Function

  Private Function ValidateSQLCode(Optional ByRef plngFixedExprID As Integer = 0, Optional ByRef psFixedSQLCode As String = "") As modExpression.ExprValidationCodes
    ' Validate the expression's SQL code. This picks up on errors such as too many nested levels of the CASE statement.
    On Error GoTo ErrorTrap

    Dim lngCalcViews(,) As Object
    Dim intCount As Short
    Dim sSource As String
    Dim sSPCode As String
    Dim strJoinCode As String
    Dim iValidationCode As modExpression.ExprValidationCodes
    Dim sSQLCode As String
    Dim lngOriginalExprID As Integer
    Dim sOriginalSQLCode As String
    Dim alngSourceTables(,) As Object
    Dim sProcName As String
    Dim avDummyPrompts(,) As Object
    Dim intStart As Short
    Dim intFound As Short

    ReDim avDummyPrompts(1, 0)

    iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_NOERRORS

    If ((Not ExprDeleted((Me.ExpressionID))) Or (mlngExpressionID = 0)) And ((miExpressionType = modExpression.ExpressionTypes.giEXPR_VIEWFILTER) Or (miExpressionType = modExpression.ExpressionTypes.giEXPR_RUNTIMEFILTER) Or (miExpressionType = modExpression.ExpressionTypes.giEXPR_RUNTIMECALCULATION)) Then

      mfConstructed = True

      If ((miExpressionType = modExpression.ExpressionTypes.giEXPR_VIEWFILTER) Or (miExpressionType = modExpression.ExpressionTypes.giEXPR_RUNTIMEFILTER)) Then
        If RuntimeFilterCode(sSQLCode, False, True, avDummyPrompts, plngFixedExprID, psFixedSQLCode) Then

          On Error GoTo SQLCodeErrorTrap

          sProcName = datGeneral.UniqueSQLObjectName("tmpsp_ASRExprTest", 4)

          ' Create the test stored procedure to see if the filter expression is valid.
          sSPCode = " CREATE PROCEDURE " & sProcName & " AS " & sSQLCode
          gADOCon.Execute(sSPCode, , ADODB.CommandTypeEnum.adCmdText)

          datGeneral.DropUniqueSQLObject(sProcName, 4)

          On Error GoTo ErrorTrap
        Else
          iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_FILTEREVALUATION
        End If
      End If

      If (miExpressionType = modExpression.ExpressionTypes.giEXPR_RUNTIMECALCULATION) Then
        ReDim lngCalcViews(2, 0)
        If RuntimeCalculationCode(lngCalcViews, sSQLCode, False, True, avDummyPrompts, plngFixedExprID, psFixedSQLCode) Then
          ' Add the required views to the JOIN code.
          strJoinCode = vbNullString
          For intCount = 1 To UBound(lngCalcViews, 2)
            ' JPD20020513 Fault 3871 - Join parent tables as well as views.
            If lngCalcViews(1, intCount) = 1 Then
              sSource = gcoTablePrivileges.FindViewID(lngCalcViews(2, intCount)).RealSource
            Else
              sSource = gcoTablePrivileges.FindTableID(lngCalcViews(2, intCount)).RealSource
            End If

            strJoinCode = strJoinCode & vbNewLine & " LEFT OUTER JOIN " & sSource & " ON " & msBaseTableName & ".ID = " & sSource & ".ID"
          Next

          sSQLCode = "SELECT " & sSQLCode & " FROM " & msBaseTableName & strJoinCode

          On Error GoTo SQLCodeErrorTrap

          sProcName = datGeneral.UniqueSQLObjectName("tmpsp_ASRExprTest", 4)

          ' Create the test stored procedure to see if the filter expression is valid.
          sSPCode = " CREATE PROCEDURE " & sProcName & " AS " & sSQLCode
          gADOCon.Execute(sSPCode, , ADODB.CommandTypeEnum.adCmdText)

          ' Drop the test stored procedure.
          datGeneral.DropUniqueSQLObject(sProcName, 4)

          On Error GoTo ErrorTrap
        Else
          iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_FILTEREVALUATION
        End If
      End If

      If iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_NOERRORS Then
        ' Need to check if all calcs/filters that use this filter are still okay.
        'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
        If (IsNothing(plngFixedExprID) And IsNothing(psFixedSQLCode)) Or ((plngFixedExprID = 0) And (psFixedSQLCode = "")) Then
          lngOriginalExprID = mlngExpressionID

          ' Create an array of the IDs of the tables/view referred to in the expression.
          ' This is used for joining all of the tables/views used.
          ' Column 1 = 0 if this row is for a table, 1 if it is for a view.
          ' Column 2 = table/view ID.
          ReDim alngSourceTables(2, 0)

          RuntimeCode(sSQLCode, alngSourceTables, False, True, avDummyPrompts, plngFixedExprID, psFixedSQLCode)
          sOriginalSQLCode = sSQLCode
        Else
          lngOriginalExprID = plngFixedExprID
          sOriginalSQLCode = psFixedSQLCode
        End If

        iValidationCode = ValidateAssociatedExpressionsSQLCode(lngOriginalExprID, sOriginalSQLCode)
      End If
    End If

TidyUpAndExit:
    ValidateSQLCode = iValidationCode
    Exit Function

SQLCodeErrorTrap:
    'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
    If (IsNothing(plngFixedExprID) And IsNothing(psFixedSQLCode)) Or ((plngFixedExprID = 0) And (psFixedSQLCode = "")) Then
      iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_SQLERROR
    Else
      iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_ASSOCSQLERROR
    End If
    msErrorMessage = Err.Description

    Do
      intStart = intFound
      intFound = InStr(intStart + 1, msErrorMessage, "]")
    Loop While intFound > 0

    If intStart > 0 And intStart < Len(Trim(msErrorMessage)) Then
      msErrorMessage = Trim(Mid(msErrorMessage, intStart + 1))
    End If

    'UPGRADE_NOTE: Object mobjBadComponent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mobjBadComponent = Nothing
    Resume TidyUpAndExit

ErrorTrap:
    iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_UNKNOWNERROR
    Resume TidyUpAndExit

  End Function

	Private Function ValidateAssociatedExpressionsSQLCode(ByRef plngFixedExpressionID As Integer, ByRef psFixedSQLCode As String) As modExpression.ExprValidationCodes
		' Validate the SQL code for any expressions that use this expression.
		' This picks up on errors such as too many nested levels of the CASE statement.
		Dim iValidationCode As modExpression.ExprValidationCodes
		Dim sSQL As String
		Dim rsTemp As ADODB.Recordset
		Dim objComp As clsExprComponent
		Dim objExpr As clsExprExpression
		
		iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_NOERRORS
		
		' Do nothing if this is a new expression
		If mlngExpressionID = 0 Then
			ValidateAssociatedExpressionsSQLCode = iValidationCode
			Exit Function
		End If
		
		sSQL = "SELECT componentID" & " FROM ASRSysExprComponents" & " WHERE calculationID = " & mlngExpressionID & " OR filterID = " & mlngExpressionID & " OR (fieldSelectionFilter = " & mlngExpressionID & " AND type = " & CStr(modExpression.ExpressionComponentTypes.giCOMPONENT_FIELD) & ")"
		rsTemp = datGeneral.GetRecords(sSQL)
		With rsTemp
			Do While (Not .EOF) And (iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_NOERRORS)
				objComp = New clsExprComponent
				objComp.ComponentID = .Fields("ComponentID").Value
				
				objExpr = New clsExprExpression
				objExpr.ExpressionID = objComp.RootExpressionID
				objExpr.ConstructExpression()
				iValidationCode = objExpr.ValidateSQLCode(plngFixedExpressionID, psFixedSQLCode)
				If iValidationCode <> modExpression.ExprValidationCodes.giEXPRVALIDATION_NOERRORS Then
					msErrorMessage = objExpr.ErrorMessage
				End If
				'UPGRADE_NOTE: Object objExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				objExpr = Nothing
				
				'UPGRADE_NOTE: Object objComp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				objComp = Nothing
				
				.MoveNext()
			Loop 
			.Close()
		End With
		'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsTemp = Nothing
		
		ValidateAssociatedExpressionsSQLCode = iValidationCode
		
	End Function
	
	
	
	
	
	Public Function ConstructExpression() As Boolean
		' Read the expression definition from the database and
		' construct the hierarchy of component class objects.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		Dim sSQL As String
		Dim objComponent As clsExprComponent
		Dim rsExpression As ADODB.Recordset
		Dim rsComponents As ADODB.Recordset
		
		fOK = True
		
		'JPD 20031110 Fault 7544
		ReadPersonnelParameters()
		
		' Do nothing if the expression is already constructed.
		If mfConstructed Then
			If miExpressionType = modExpression.ExpressionTypes.giEXPR_RUNTIMECALCULATION Then
				miReturnType = modExpression.ExpressionValueTypes.giEXPRVALUE_UNDEFINED
			End If
			
			If mlngExpressionID > 0 Then
				' Get the expression timestamp.
				sSQL = sSQL & "SELECT CONVERT(integer, ASRSysExpressions.timestamp) AS intTimestamp" & " FROM ASRSysExpressions" & " WHERE exprID = " & Trim(Str(mlngExpressionID))
				rsExpression = datGeneral.GetRecords(sSQL)
				With rsExpression
					fOK = Not (.EOF And .BOF)
					If fOK Then
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						If Not mfDontUpdateTimeStamp Then mlngTimeStamp = IIf(IsDbNull(.Fields("intTimestamp").Value), 0, .Fields("intTimestamp").Value)
					End If
					.Close()
				End With
				'UPGRADE_NOTE: Object rsExpression may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				rsExpression = Nothing
			End If
		Else
			'TM20030604 Fault - create different SQL code depending on the Expression Type.
			'    ' Get the expression definition.
			'    sSQL = sSQL & "SELECT ASRSysExpressions.name," & _
			''      " ASRSysExpressions.TableID," & _
			''      " ASRSysExpressions.returnType," & _
			''      " ASRSysExpressions.type," & _
			''      " ASRSysExpressions.parentComponentID," & _
			''      " ASRSysExpressions.Username," & _
			''      " ASRSysExpressions.access," & _
			''      " ASRSysExpressions.description," & _
			''      " ASRSysExpressions.ViewInColour," & _
			''      " CONVERT(integer, ASRSysExpressions.timestamp) AS intTimestamp," & _
			''      " ASRSysTables.tableName" & _
			''      " FROM ASRSysExpressions" & _
			''      " INNER JOIN ASRSysTables ON ASRSysExpressions.TableID = ASRSysTables.tableID" & _
			''      " WHERE exprID = " & Trim(Str(mlngExpressionID))
			'
			'      '" ASRSysExpressions.owner," &
			
			' Get the expression definition.
			If miExpressionType = modExpression.ExpressionTypes.giEXPR_UTILRUNTIMEFILTER Then
				' Utility runtime filters are not tied to a base table.
				sSQL = sSQL & "SELECT ASRSysExpressions.name," & " 0 AS tableID," & " ASRSysExpressions.returnType," & " ASRSysExpressions.type," & " ASRSysExpressions.parentComponentID," & " ASRSysExpressions.Username," & " ASRSysExpressions.access," & " ASRSysExpressions.description," & " ASRSysExpressions.ViewInColour," & " CONVERT(integer, ASRSysExpressions.timestamp) AS intTimestamp," & " '' AS tableName" & " FROM ASRSysExpressions" & " WHERE exprID = " & Trim(Str(mlngExpressionID))
				
			ElseIf miExpressionType = modExpression.ExpressionTypes.giEXPR_RECORDINDEPENDANTCALC Then 
				sSQL = sSQL & "SELECT ASRSysExpressions.name," & " ASRSysExpressions.TableID," & " ASRSysExpressions.returnType," & " ASRSysExpressions.type," & " ASRSysExpressions.parentComponentID," & " ASRSysExpressions.Username," & " ASRSysExpressions.access," & " ASRSysExpressions.description," & " ASRSysExpressions.ViewInColour," & " CONVERT(integer, ASRSysExpressions.timestamp) AS intTimestamp," & " ASRSysTables.tableName" & " FROM ASRSysExpressions" & " LEFT OUTER JOIN ASRSysTables ON ASRSysExpressions.TableID = ASRSysTables.tableID" & " WHERE exprID = " & Trim(Str(mlngExpressionID))
				
			Else
				'TM29102003 Fault 7421 - 'LEFT OUTER' JOIN instead of 'INNER'.
				'      sSQL = sSQL & "SELECT ASRSysExpressions.name," & _
				''        " ASRSysExpressions.TableID," & _
				''        " ASRSysExpressions.returnType," & _
				''        " ASRSysExpressions.type," & _
				''        " ASRSysExpressions.parentComponentID," & _
				''        " ASRSysExpressions.Username," & _
				''        " ASRSysExpressions.access," & _
				''        " ASRSysExpressions.description," & _
				''        " ASRSysExpressions.ViewInColour," & _
				''        " CONVERT(integer, ASRSysExpressions.timestamp) AS intTimestamp," & _
				''        " ASRSysTables.tableName" & _
				''        " FROM ASRSysExpressions" & _
				''        " INNER JOIN ASRSysTables ON ASRSysExpressions.TableID = ASRSysTables.tableID" & _
				''        " WHERE exprID = " & Trim(Str(mlngExpressionID))
				sSQL = sSQL & "SELECT ASRSysExpressions.name," & " ASRSysExpressions.TableID," & " ASRSysExpressions.returnType," & " ASRSysExpressions.type," & " ASRSysExpressions.parentComponentID," & " ASRSysExpressions.Username," & " ASRSysExpressions.access," & " ASRSysExpressions.description," & " ASRSysExpressions.ViewInColour," & " CONVERT(integer, ASRSysExpressions.timestamp) AS intTimestamp," & " ASRSysTables.tableName" & " FROM ASRSysExpressions" & " LEFT OUTER JOIN ASRSysTables ON ASRSysExpressions.TableID = ASRSysTables.tableID" & " WHERE exprID = " & Trim(Str(mlngExpressionID))
			End If
			
			rsExpression = datGeneral.GetRecords(sSQL)
			With rsExpression
				fOK = Not (.EOF And .BOF)
				If fOK Then
					' Read the expression's properties.
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					msExpressionName = IIf(IsDbNull(.Fields("Name").Value), "", .Fields("Name").Value)
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					mlngBaseTableID = IIf(IsDbNull(.Fields("TableID").Value), 0, .Fields("TableID").Value)
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					miReturnType = IIf(IsDbNull(.Fields("ReturnType").Value), modExpression.ExpressionValueTypes.giEXPRVALUE_UNDEFINED, .Fields("ReturnType").Value)
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					miExpressionType = IIf(IsDbNull(.Fields("Type").Value), modExpression.ExpressionTypes.giEXPR_UNKNOWNTYPE, .Fields("Type").Value)
					
					If miExpressionType = modExpression.ExpressionTypes.giEXPR_RUNTIMECALCULATION Then
						miReturnType = modExpression.ExpressionValueTypes.giEXPRVALUE_UNDEFINED
					End If
					
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					mlngParentComponentID = IIf(IsDbNull(.Fields("ParentComponentID").Value), 0, .Fields("ParentComponentID").Value)
					'msOwner = IIf(IsNull(!Owner), gsUserName, !Owner)
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					msOwner = IIf(IsDbNull(.Fields("Username").Value), gsUsername, .Fields("Username").Value)
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					msAccess = IIf(IsDbNull(.Fields("Access").Value), "RW", .Fields("Access").Value)
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					msDescription = IIf(IsDbNull(.Fields("Description").Value), "", .Fields("Description").Value)
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					mlngTimeStamp = IIf(IsDbNull(.Fields("intTimestamp").Value), 0, .Fields("intTimestamp").Value)
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					msBaseTableName = IIf(IsDbNull(.Fields("TableName").Value), "", .Fields("TableName").Value)
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					mbViewInColour = IIf(IsDbNull(.Fields("ViewInColour").Value), False, .Fields("ViewInColour").Value)
					
				Else
					' Initialise the expression.
					InitialiseExpression()
				End If
				
				.Close()
			End With
			'UPGRADE_NOTE: Object rsExpression may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsExpression = Nothing
			
			If fOK Then
				' Clear the expressions collection of components.
				ClearComponents()
				
				' Get the expression definition.
				sSQL = "SELECT *" & " FROM ASRSysExprComponents" & " WHERE exprID = " & Trim(Str(mlngExpressionID)) & " ORDER BY componentID"
				rsComponents = datGeneral.GetRecords(sSQL)
				
				Do While (Not rsComponents.EOF) And fOK
					' Instantiate a new component object.
					objComponent = New clsExprComponent
					
					With objComponent
						' Initialise the new component's properties.
						.ParentExpression = Me
						.ComponentID = rsComponents.Fields("ComponentID").Value
						
						' Instruct the new component to read it's own definition from the database.
						fOK = .ConstructComponent(rsComponents)
					End With
					
					If fOK Then
						' If the component definition was read correctly then
						' add the new component to the expression's component collection.
						mcolComponents.Add(objComponent)
					End If
					
					' Disassociate object variables.
					'UPGRADE_NOTE: Object objComponent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					objComponent = Nothing
					
					rsComponents.MoveNext()
				Loop 
				
				rsComponents.Close()
			End If
		End If
		
TidyUpAndExit: 
		mfConstructed = fOK
		'UPGRADE_NOTE: Object rsExpression may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsExpression = Nothing
		'UPGRADE_NOTE: Object rsComponents may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsComponents = Nothing
		'UPGRADE_NOTE: Object objComponent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objComponent = Nothing
		ConstructExpression = fOK
		Exit Function
		
ErrorTrap: 
		fOK = False
		'NO MSGBOX ON THE SERVER ! - MsgBox "Error constructing the expression.", _
		'vbOKOnly + vbExclamation, App.ProductName
		Err.Number = False
		Resume TidyUpAndExit
		
	End Function
	
	
	Private Sub InitialiseExpression()
		' Initialize the properties for a new expression,
		' and clear the expression's component collection.
		ExpressionID = 0
		
		msExpressionName = ""
		mlngParentComponentID = 0
		msOwner = gsUsername
		msAccess = "RW"
		msDescription = ""
		mlngTimeStamp = 0
		
		mfConstructed = True
		
		' Clear any existing components from
		' the expression's component collection.
		ClearComponents()
		
	End Sub
	
	
	
	
	Public Sub ClearComponents()
		' Clear the expression's component collection.
		On Error GoTo ErrorTrap
		
		' Remove all components from the collection.
		Do While mcolComponents.Count() > 0
			mcolComponents.Remove(1)
		Loop 
		'UPGRADE_NOTE: Object mcolComponents may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mcolComponents = Nothing
		
		' Re-instantiate the collection.
		mcolComponents = New Collection
		
		Exit Sub
		
ErrorTrap: 
		Err.Number = False
		
	End Sub
	
	Public Function Initialise(ByRef plngBaseTableID As Integer, ByRef plngExpressionID As Integer, ByRef piType As Short, ByRef piReturnType As Short) As Boolean
		' Initialise the expression object.
		' Return TRUE if everything was initialised okay.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		
		fOK = True
		
		BaseTableID = plngBaseTableID
		ExpressionID = plngExpressionID
		miExpressionType = piType
		miReturnType = piReturnType
		
TidyUpAndExit: 
		Initialise = fOK
		Exit Function
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	
	
	
	
	
	Public Function ValidateSelection() As Boolean
		' Validate the expression section.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		Dim sSQL As String
		Dim objExpr As clsExprExpression
		Dim rsCheck As ADODB.Recordset
		
		fOK = True
		
		' Construct the expression to print.
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		fOK = ConstructExpression
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		
TidyUpAndExit: 
		ValidateSelection = fOK
		Exit Function
		
ErrorTrap: 
		fOK = False
		'NO MSGBOX ON THE SERVER ! - MsgBox Err.Description, vbExclamation + vbOKOnly, App.ProductName
		Err.Number = False
		Resume TidyUpAndExit
		
	End Function
	
	' Creates a UDF for this expression if its required
	Public Function UDFCode(ByRef psRuntimeCode() As String, ByRef palngSourceTables As Object, ByRef pfApplyPermissions As Boolean, ByRef pfValidating As Boolean, Optional ByRef plngFixedExprID As Integer = 0, Optional ByRef psFixedSQLCode As String = "") As Boolean
		
		Dim iLoop1 As Short
		Dim fOK As Boolean
		
		fOK = True
		
		For iLoop1 = 1 To mcolComponents.Count()
			With mcolComponents.Item(iLoop1)
				
				' Add the created UDFs to the total list
				'JPD 20040227 Fault 8146
				'fOK = .UDFCode(psRuntimeCode(), palngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents().UDFCode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				fOK = .UDFCode(psRuntimeCode, palngSourceTables, pfApplyPermissions, True, plngFixedExprID, psFixedSQLCode)
				
			End With
			
			If Not fOK Then
				Exit For
			End If
			
		Next iLoop1
		
		UDFCode = fOK
		
	End Function
	
	
	Public Function UDFCalculationCode(ByRef palngSourceTables As Object, ByRef psCalcCode() As String, ByRef pfApplyPermissions As Boolean, Optional ByRef pfValidating As Boolean = False, Optional ByRef plngFixedExprID As Integer = 0, Optional ByRef psFixedSQLCode As String = "") As Boolean
		' Return TRUE if the Calculation code was created okay.
		' Return the runtime Calculation SQL code in the parameter 'psCalcCode'.
		' Apply permissions to the Calculation code only if the 'pfApplyPermissions' parameter is TRUE.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		Dim sRuntimeSQL As String
		
		' Check if the 'validating' parameter is set.
		' If not, set it to FALSE.
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(pfValidating) Then
			pfValidating = False
		End If
		
		' Construct the expression from the database definition.
		fOK = ConstructExpression
		
		If fOK Then
			' Get the Calculation code.
			' JPD20020419 Fault 3687
			fOK = UDFCode(psCalcCode, palngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
		End If
		
		If fOK Then
			If pfApplyPermissions Then
				fOK = (ValidateExpression(True) = modExpression.ExprValidationCodes.giEXPRVALIDATION_NOERRORS)
			End If
		End If
		
TidyUpAndExit: 
		If Not fOK Then
			psCalcCode(UBound(psCalcCode)) = ""
		End If
		UDFCalculationCode = fOK
		
		Exit Function
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	
	Public Function UDFFilterCode(ByRef pastrFilterCode() As String, ByRef pfApplyPermissions As Boolean, Optional ByRef pfValidating As Boolean = False, Optional ByRef plngFixedExprID As Integer = 0, Optional ByRef psFixedSQLCode As String = "") As Boolean
		' Return TRUE if the filter code was created okay.
		' Return the runtime filter SQL code in the parameter 'pastrFilterCode'.
		' Apply permissions to the filter code only if the 'pfApplyPermissions' parameter is TRUE.
		' The filter code is to be used to validate the expression if the 'pfValidating' parameter is TRUE.
		' This is used to suppress prompting the user for promted values, when we are only validating the expression.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		Dim iLoop1 As Short
		Dim iLoop2 As Short
		Dim iNextIndex As Short
		Dim sSQL As String
		Dim sWhereCode As String
		Dim sBaseTableSource As String
		Dim sRuntimeFilterSQL As String
    Dim alngSourceTables(,) As Integer
		Dim avRelatedTables() As Object
		Dim rsInfo As ADODB.Recordset
		Dim objTableView As CTablePrivilege
		
		' Check if the 'validating' parameter is set.
		' If not, set it to FALSE.
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(pfValidating) Then
			pfValidating = False
		End If
		
		' Construct the expression from the database definition.
		fOK = ConstructExpression
		
		If fOK Then
			sBaseTableSource = msBaseTableName
			If pfApplyPermissions Then
				' Get the 'realSource' of the table.
				objTableView = gcoTablePrivileges.Item(msBaseTableName)
				If objTableView.TableType = Declarations.TableTypes.tabChild Then
					sBaseTableSource = objTableView.RealSource
				End If
				'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				objTableView = Nothing
			End If
			
			' Create an array of the IDs of the tables/view referred to in the expression.
			' This is used for joining all of the tables/views used.
			' Column 1 = 0 if this row is for a table, 1 if it is for a view.
			' Column 2 = table/view ID.
			ReDim alngSourceTables(2, 0)
			
			' Get the filter code.
			fOK = UDFCode(pastrFilterCode, alngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
		End If
		
TidyUpAndExit: 
		If Not fOK Then
			pastrFilterCode(UBound(pastrFilterCode)) = ""
		End If
		UDFFilterCode = fOK
		
		Exit Function
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	
	Public Sub UDFFunctions(ByRef pbCreate As Boolean)
		
		Dim iCount As Short
		Dim strDropCode As String
		Dim strFunctionName As String
		Dim sUDFCode As String
		Dim clsData As clsDataAccess
		Dim iStart As Short
		Dim iEnd As Short
		Dim strFunctionNumber As String
		
		Const FUNCTIONPREFIX As String = "udf_ASRSys_"
		
		On Error GoTo ExecuteSQL_ERROR
		
		If gbEnableUDFFunctions Then
			
			clsData = New clsDataAccess
			
			For iCount = 1 To UBound(mastrUDFsRequired)
				
				'JPD 20060110 Fault 10509
				'strFunctionName = Mid(mastrUDFsRequired(iCount), 17, 15)
				iStart = InStr(mastrUDFsRequired(iCount), FUNCTIONPREFIX) + Len(FUNCTIONPREFIX)
				iEnd = InStr(1, Mid(mastrUDFsRequired(iCount), 1, 1000), "(@Pers")
				strFunctionNumber = Mid(mastrUDFsRequired(iCount), iStart, iEnd - iStart)
				strFunctionName = FUNCTIONPREFIX & strFunctionNumber
				
				'Drop existing function (could exist if the expression is used more than once in a report)
				strDropCode = "IF EXISTS" & " (SELECT *" & "   FROM sysobjects" & "   WHERE id = object_id('[" & Replace(gsUsername, "'", "''") & "]." & strFunctionName & "')" & "     AND sysstat & 0xf = 0)" & " DROP FUNCTION [" & gsUsername & "]." & strFunctionName
				gADOCon.Execute(strDropCode)
				
				' Create the new function
				If pbCreate Then
					sUDFCode = mastrUDFsRequired(iCount)
					gADOCon.Execute(sUDFCode)
				End If
				
			Next iCount
		End If
		
		Exit Sub
		
ExecuteSQL_ERROR: 
		
		msErrorMessage = "Error whilst creating user defined functions." & vbNewLine & Err.Description
		
	End Sub
End Class