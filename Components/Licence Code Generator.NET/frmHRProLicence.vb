Option Strict Off
Option Explicit On
Friend Class frmHRProLicence
	Inherits System.Windows.Forms.Form
	
    'Private WithEvents gSysTray As clsSysTray
	Private mstrAllowedInputCharacters As String
	
	Private Sub PopulateModules(ByRef lstTemp As System.Windows.Forms.ListBox)
		
		Dim lngBit As Integer
		
		lngBit = 1
		With lstTemp
			.Items.Add(New VB6.ListBoxItem("Personnel  ", lngBit)) : lngBit = lngBit * 2
			.Items.Add(New VB6.ListBoxItem("Recruitment", lngBit)) : lngBit = lngBit * 2
			.Items.Add(New VB6.ListBoxItem("Absence    ", lngBit)) : lngBit = lngBit * 2
			.Items.Add(New VB6.ListBoxItem("Training   ", lngBit)) : lngBit = lngBit * 2
			.Items.Add(New VB6.ListBoxItem("Intranet   ", lngBit)) : lngBit = lngBit * 2
			.Items.Add(New VB6.ListBoxItem("AFD        ", lngBit)) : lngBit = lngBit * 2
			.Items.Add(New VB6.ListBoxItem("Full SysMgr", lngBit)) : lngBit = lngBit * 2
			.Items.Add(New VB6.ListBoxItem("CMG        ", lngBit)) : lngBit = lngBit * 2
			.Items.Add(New VB6.ListBoxItem("Quick Address", lngBit)) : lngBit = lngBit * 2
			.Items.Add(New VB6.ListBoxItem("Payroll (Shared Table)", lngBit)) : lngBit = lngBit * 2
			.Items.Add(New VB6.ListBoxItem("Workflow", lngBit)) : lngBit = lngBit * 2
			.Items.Add(New VB6.ListBoxItem("V1 Integration", lngBit)) : lngBit = lngBit * 2
			.Items.Add(New VB6.ListBoxItem("Mobile Interface", lngBit)) : lngBit = lngBit * 2
			.Items.Add(New VB6.ListBoxItem("Fusion Integration", lngBit)) : lngBit = lngBit * 2
			.Items.Add(New VB6.ListBoxItem("XML Exports", lngBit)) : lngBit = lngBit * 2
			.Items.Add(New VB6.ListBoxItem("3rd Party Tables", lngBit)) : lngBit = lngBit * 2
			.Items.Add(New VB6.ListBoxItem("9-Box Grid Reports", lngBit)) : lngBit = lngBit * 2
			.Items.Add(New VB6.ListBoxItem("Editable Grids", lngBit)) : lngBit = lngBit * 2
			.Items.Add(New VB6.ListBoxItem("Power Customisation Pack", lngBit)) : lngBit = lngBit * 2
		End With
		
	End Sub
	
	'UPGRADE_WARNING: Event cboType.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboType.SelectedIndexChanged

        Select Case cboType.SelectedIndex
            Case 0
                txtDatUsers.Enabled = True
                txtIntUsers.Enabled = True
                txtSSIUsers.Enabled = True
                txtHeadcount.Enabled = False

                txtDatUsers.BackColor = Color.White
                txtIntUsers.BackColor = Color.White
                txtSSIUsers.BackColor = Color.White
                txtHeadcount.BackColor = Me.BackColor
            Case 1
                txtDatUsers.Enabled = False
                txtIntUsers.Enabled = False
                txtSSIUsers.Enabled = False
                txtHeadcount.Enabled = True

                txtDatUsers.BackColor = Me.BackColor
                txtIntUsers.BackColor = Me.BackColor
                txtSSIUsers.BackColor = Me.BackColor
                txtHeadcount.BackColor = Color.White
            Case 2
                txtDatUsers.Enabled = False
                txtIntUsers.Enabled = False
                txtSSIUsers.Enabled = False
                txtHeadcount.Enabled = True

                txtDatUsers.BackColor = Me.BackColor
                txtIntUsers.BackColor = Me.BackColor
                txtSSIUsers.BackColor = Me.BackColor
                txtHeadcount.BackColor = Color.White
            Case 3
                txtDatUsers.Enabled = True
                txtIntUsers.Enabled = True
                txtSSIUsers.Enabled = True
                txtHeadcount.Enabled = True

                txtDatUsers.BackColor = Color.White
                txtIntUsers.BackColor = Color.White
                txtSSIUsers.BackColor = Color.White
                txtHeadcount.BackColor = Color.White
            Case 4
                txtDatUsers.Enabled = True
                txtIntUsers.Enabled = True
                txtSSIUsers.Enabled = True
                txtHeadcount.Enabled = True

                txtDatUsers.BackColor = Color.White
                txtIntUsers.BackColor = Color.White
                txtSSIUsers.BackColor = Color.White
                txtHeadcount.BackColor = Color.White
        End Select

    End Sub
	
	Private Sub cmdClipboard_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClipboard.Click
		My.Computer.Clipboard.Clear()
		My.Computer.Clipboard.SetText(Me.LicenceKey)
	End Sub
	
	Private Sub cmdRead_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRead.Click
		
		Dim objLicence As New clsLicence
		Dim lngCustNo As Integer
		Dim lngUsers As Integer
		Dim lngModules As Integer
		Dim lngCount As Integer
		
		With objLicence
			.ValidateCreationDate = False
			.LicenceKey = Me.LicenceKey
			
			If Not .IsValid Then
				MsgBox("Invalid Key")
				Exit Sub
			End If
			
            If (.CustomerNo < 1000 Or .CustomerNo > 9999) Then 'And vbCompiled Then
                MsgBox("Invalid Licence Key", MsgBoxStyle.Exclamation)
            Else
                txtCustomerNo.Text = CStr(.CustomerNo)
                txtDatUsers.Text = CStr(.DATUsers)
                txtIntUsers.Text = CStr(.DMIMUsers)
                txtSSIUsers.Text = CStr(.SSIUsers)
                txtHeadcount.Text = CStr(.Headcount)

                If IsDate(.ExpiryDate) And Year(.ExpiryDate) > 1900 Then
                    'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
                    txtExpiryDate.Text = .ExpiryDate
                Else
                    'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
                    txtExpiryDate.Text = ""
                End If
                cboType.SelectedIndex = .LicenceType

                For lngCount = 0 To lstModules.Items.Count - 1
                    lstModules.SetItemChecked(lngCount, (.Modules And VB6.GetItemData(lstModules, lngCount)))
                Next

            End If
		End With
		
		'UPGRADE_NOTE: Object objLicence may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objLicence = Nothing
		
	End Sub
	
	Private Sub cmdSuppClipboard_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSuppClipboard.Click
		My.Computer.Clipboard.Clear()
		My.Computer.Clipboard.SetText(txtSupportOutput(0).Text & "-" & txtSupportOutput(1).Text & "-" & txtSupportOutput(2).Text & "-" & txtSupportOutput(3).Text & "-" & txtSupportOutput(4).Text)
	End Sub
	
	Private Sub frmHRProLicence_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        txtExpiryDate.Format = DateTimePickerFormat.Custom
        txtExpiryDate.CustomFormat = " "

		SSTab1.SelectedIndex = 0
		Frame2.BackColor = Me.BackColor
		Frame3.BackColor = Me.BackColor
		Frame4.BackColor = Me.BackColor
		PopulateModules(lstModules)
		'PopulateModules lstModules
		
		mstrAllowedInputCharacters = GenerateAlphaString
		
		'Only show the read licence tab if in development!
		'On Local Error Resume Next
		'Err.Clear
		'Debug.Print 1 / 0
		'Me.SSTab1.TabVisible(1) = (Err.Number > 0)
		
		cboType.SelectedIndex = 0
		
        'gSysTray = New clsSysTray
        'gSysTray.SourceWindow = Me
        'gSysTray.ChangeIcon(Me.Icon)
		
	End Sub
	
    'Private Sub frmHRProLicence_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
    '	Dim Cancel As Boolean = eventArgs.Cancel
    '	Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
    '	gSysTray.RemoveFromSysTray()
    '	eventArgs.Cancel = Cancel
    'End Sub
	
    ''UPGRADE_WARNING: Event frmHRProLicence.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    'Private Sub frmHRProLicence_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
    '	If Me.WindowState = System.Windows.Forms.FormWindowState.Minimized Then
    '		gSysTray.MinToSysTray()
    '	End If
    'End Sub

    'Private Sub gSysTray_LButtonDblClk() Handles gSysTray.LButtonDblClk
    '	If Me.WindowState = System.Windows.Forms.FormWindowState.Minimized Then
    '		gSysTray.RemoveFromSysTray()
    '		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
    '	End If
    'End Sub
	
	Private Sub SSTab1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SSTab1.SelectedIndexChanged
		Static PreviousTab As Short = SSTab1.SelectedIndex()
		
		fraCustomerDetails(0).Enabled = (SSTab1.SelectedIndex = 0)
		fraLicenceGenerate.Enabled = (SSTab1.SelectedIndex = 0)
		
		'fraLicenceRead.Enabled = (SSTab1.Tab = 1)
		'fraCustomerDetails(1).Enabled = (SSTab1.Tab = 1)
		
		PreviousTab = SSTab1.SelectedIndex()
	End Sub
	
	
	Private Sub cmdGenerate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdGenerate.Click
		
		Dim objLicence As clsLicenceWrite2
		Dim lngCount As Integer
		Dim lngModules As Integer
		
		'Validate customer number...
		With txtCustomerNo
			If Len(.Text) <> 4 Or Val(.Text) < 1000 Then
				MsgBox("Invalid Customer Number", MsgBoxStyle.Exclamation)
				.Focus()
				Exit Sub
			End If
		End With
		
		
		'Check with modules have been selected...
		With lstModules
			lngModules = 0
			For lngCount = 0 To .Items.Count - 1
				If .GetItemChecked(lngCount) Then
					lngModules = lngModules + VB6.GetItemData(lstModules, lngCount)
				End If
			Next 
			
			If lngModules = 0 Then
				MsgBox("No Modules selected", MsgBoxStyle.Exclamation)
				.Focus()
				Exit Sub
			End If
		End With
		
		
		objLicence = New clsLicenceWrite2
		
		With objLicence
			
			.CustomerNo = Val(txtCustomerNo.Text)
			.DATUsers = Val(txtDatUsers.Text)
			.DMIMUsers = Val(txtIntUsers.Text)
			.SSIUsers = Val(txtSSIUsers.Text)
			.Headcount = Val(txtHeadcount.Text)
			
            If IsDate(txtExpiryDate.Text) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object txtExpiryDate.DateValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.ExpiryDate = txtExpiryDate.Value.Date.ToOADate
            End If
			
			.LicenceType = cboType.SelectedIndex
			
			.Modules = lngModules
			
			'UPGRADE_WARNING: Couldn't resolve default property of object objLicence.LicenceKey2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Me.LicenceKey = .LicenceKey2
			
		End With
		
		'UPGRADE_NOTE: Object objLicence may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objLicence = Nothing
		
	End Sub
	
	
	Private Function GenerateAlphaString() As String
		
		Dim strOutput As String
		Dim lngCount As Integer
		Dim lngLoop As Integer
		
		'Only allow these characters...
		strOutput = vbNullString
		
		For lngCount = Asc("A") + lngLoop To Asc("Z")
			strOutput = strOutput & Chr(lngCount)
		Next 
		
		For lngCount = Asc("0") + lngLoop To Asc("9")
			strOutput = strOutput & Chr(lngCount)
		Next 
		
		GenerateAlphaString = strOutput
		
	End Function
	
	Private Sub txtCustomerNo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustomerNo.Enter
		With txtCustomerNo
			.SelectionStart = 0
			.SelectionLength = Len(.Text)
		End With
	End Sub
	
	Private Sub txtDatUsers_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDatUsers.Enter
		With txtDatUsers
			.SelectionStart = 0
			.SelectionLength = Len(.Text)
		End With
	End Sub
	
	Private Sub txtIntUsers_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtIntUsers.Enter
		With txtIntUsers
			.SelectionStart = 0
			.SelectionLength = Len(.Text)
		End With
	End Sub
	
	Private Sub txtLicence_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtLicence.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtLicence.GetIndex(eventSender)
		KeyAscii = Asc(UCase(Chr(KeyAscii)))
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtSSIUsers_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSSIUsers.Enter
		With txtSSIUsers
			.SelectionStart = 0
			.SelectionLength = Len(.Text)
		End With
	End Sub
	
	'UPGRADE_WARNING: Event txtLicence.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtLicence_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLicence.TextChanged
		Dim Index As Short = txtLicence.GetIndex(eventSender)
		
		If Len(txtLicence(Index).Text) >= 3 And txtLicence(Index).SelectionStart = 4 Then
			If Index < txtLicence.UBound Then
				txtLicence(Index + 1).Focus()
			End If
		End If
		
	End Sub
	
	Private Sub txtLicence_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLicence.Enter
		Dim Index As Short = txtLicence.GetIndex(eventSender)
		With txtLicence(Index)
			.SelectionStart = 0
			.SelectionLength = Len(.Text)
		End With
	End Sub
	
	Private Sub txtLicence_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtLicence.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = txtLicence.GetIndex(eventSender)
		
		If KeyCode = System.Windows.Forms.Keys.V And (Shift And VB6.ShiftConstants.CtrlMask) Then
			If My.Computer.Clipboard.GetText Like "??????-??????-??????-??????-??????-??????" Then
				LicenceKey = My.Computer.Clipboard.GetText
				KeyCode = 0
				Shift = 0
			End If
		End If
		
	End Sub
	
	'UPGRADE_WARNING: Event txtSupportInput.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtSupportInput_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSupportInput.TextChanged
		Dim Index As Short = txtSupportInput.GetIndex(eventSender)
		
		If Len(txtSupportInput(Index).Text) >= 4 And txtSupportInput(Index).SelectionStart = 4 Then
			If Index < txtSupportInput.UBound Then
				txtSupportInput(Index + 1).Focus()
			End If
		End If
		
	End Sub
	
	Private Sub txtSupportInput_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSupportInput.Enter
		Dim Index As Short = txtSupportInput.GetIndex(eventSender)
		With txtSupportInput(Index)
			.SelectionStart = 0
			.SelectionLength = Len(.Text)
		End With
	End Sub
	
	Private Sub txtSupportInput_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtSupportInput.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = txtSupportInput.GetIndex(eventSender)
		
		'Check if a user is trying to paste in a whole licence key
		'If they are, then separate it into each text box.
		If KeyCode = System.Windows.Forms.Keys.V And (Shift And VB6.ShiftConstants.CtrlMask) Then
			If My.Computer.Clipboard.GetText Like "????-????-????-????-????" Then
				txtSupportInput(0).Text = Mid(My.Computer.Clipboard.GetText, 1, 4)
				txtSupportInput(1).Text = Mid(My.Computer.Clipboard.GetText, 6, 4)
				txtSupportInput(2).Text = Mid(My.Computer.Clipboard.GetText, 11, 4)
				txtSupportInput(3).Text = Mid(My.Computer.Clipboard.GetText, 16, 4)
				txtSupportInput(4).Text = Mid(My.Computer.Clipboard.GetText, 21, 4)
				KeyCode = 0
				Shift = 0
			End If
		End If
		
	End Sub
	
	Private Sub txtSupportInput_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSupportInput.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtSupportInput.GetIndex(eventSender)
		
		Dim strChar As String
		
		'Allow control characters...
		If KeyAscii > 31 Then
			
			strChar = UCase(Chr(KeyAscii))
			If InStr(mstrAllowedInputCharacters, strChar) > 0 Then
				KeyAscii = Asc(strChar)
			Else
				KeyAscii = 0
			End If
			
		End If
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	
	Public Property LicenceKey() As String
		Get
			
			Dim lngCount As Integer
			
			LicenceKey = vbNullString
			For lngCount = txtLicence.LBound To txtLicence.UBound
				LicenceKey = LicenceKey & IIf(LicenceKey <> vbNullString, "-", "") & txtLicence(lngCount).Text
			Next 
			
		End Get
		Set(ByVal Value As String)
			
			Dim lngCount As Integer
			
			If Value Like "??????-??????-??????-??????-??????-??????" Then
				txtLicence(0).Text = Mid(Value, 1, 6)
				txtLicence(1).Text = Mid(Value, 8, 6)
				txtLicence(2).Text = Mid(Value, 15, 6)
				txtLicence(3).Text = Mid(Value, 22, 6)
				txtLicence(4).Text = Mid(Value, 29, 6)
				txtLicence(5).Text = Mid(Value, 36, 6)
			End If
			
		End Set
	End Property
	
	'UPGRADE_WARNING: Event txtSupportOutput.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtSupportOutput_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSupportOutput.TextChanged
		Dim Index As Short = txtSupportOutput.GetIndex(eventSender)
		
		If Len(txtSupportOutput(Index).Text) >= 4 And txtSupportOutput(Index).SelectionStart = 4 Then
			If Index < txtSupportOutput.UBound Then
				txtSupportOutput(Index + 1).Focus()
			End If
		End If
		
	End Sub
	
    'Private Function vbCompiled() As Boolean

    '	On Error Resume Next
    '	Err.Clear()
    '	Debug.Print(1 / 0)
    '	vbCompiled = (Err.Number = 0)

    'End Function
	'
	'Public Function ConvertStringToNumber2(strInput As String) As Long
	'
	'  Dim lngRandomDigit As Long
	'  Dim strAlphaString As String
	'  Dim lngOutput As Long
	'  Dim lngFactor As Double
	'  Dim lngCount As Long
	'
	'  On Error GoTo exitf
	'
	'  lngRandomDigit = Asc(Mid(strInput, 1, 1)) - 64
	'  strAlphaString = GenerateAlphaString(lngRandomDigit)
	'
	'  lngOutput = (InStr(strAlphaString, Mid(strInput, Len(strInput), 1)) - 1)
	'
	'  lngFactor = 32
	'  For lngCount = Len(strInput) - 1 To 2 Step -1
	'    lngOutput = lngOutput + _
	''      ((InStr(strAlphaString, Mid(strInput, lngCount, 1)) - 1) * lngFactor)
	'    lngFactor = lngFactor * 32
	'  Next
	'
	'  ConvertStringToNumber2 = lngOutput
	'
	'exitf:
	'
	'End Function

    Private Sub txtExpiryDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtExpiryDate.ValueChanged

        txtExpiryDate.Format = DateTimePickerFormat.Long
        ' txtExpiryDate.CustomFormat = " "


    End Sub
End Class