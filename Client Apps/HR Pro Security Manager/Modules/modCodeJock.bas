Attribute VB_Name = "modCodeJock"

Public Sub LoadSkin( _
      ByRef frm As Form, _
      ByRef objSkin As Object, _
      Optional sStylePath As String, _
      Optional sStyleIni As String)

'  ************************************************************
'      Here are the styles available.
'
'        Office2007.cjstyles:
'        -NORMALAQUA.INI
'        -NORMALBLUE.INI
'
'        Vista.cjstyles:
'        -NORMALBLACK.INI
'        -NORMALBLUE.INI
'        -NORMALSILVER.INI
'
'        WinXP.Luna.cjstyles:
'        -EXTRALARGEBLUE.INI
'        -EXTRALARGEHOMESTEAD.INI
'        -EXTRALARGEMETALLIC.INI
'        -LARGEBLUE.INI
'        -LARGEHOMESTEAD.INI
'        -LARGEMETALLIC.INI
'        -NORMALBLUE.INI
'        -NORMALHOMESTEAD.INI
'        -NORMALMETALLIC.INI
'
'        WinXP.Royale.cjstyles:
'        -EXTRALARGEFONTSROYALE.INI
'        -LARGEFONTSROYALE.INI
'        -NORMALROYALE.INI
'  ************************************************************

  With objSkin
    'Some dlls don't like to be hooked.
    .ExcludeModule "msado15.dll"
    .ExcludeModule "oledb32.dll"
    .ExcludeModule "msadce.dll"
    .ExcludeModule "msadcer.dll"
    .ExcludeModule "ws2_32.dll"
    .ExcludeModule "ws2help.dll"
    .ExcludeModule "netapi32.dll"

    'Loads the skin
    If Trim(sStylePath) = vbNullString Then
      .LoadSkin CodeJockStylePath, CodeJockStyleIni
    Else
      .LoadSkin sStylePath, sStyleIni
    End If
  
    ' Handle some of those "interesting" 3rd party controls
    .AddWindowClass "MSMaskWndClass", "Edit"
  
    'Applies the currently loaded skin to the specified window
    .ApplyWindow frm.hWnd
  End With
End Sub

Public Property Get CodeJockStylePath() As String
'  CodeJockStylePath = App.Path + "\Styles\COA.cjstyles"
  CodeJockStylePath = App.Path & "\" & App.EXEName & ".exe"
End Property

Public Property Get CodeJockStyleIni() As String
  CodeJockStyleIni = "NORMALSILVER.INI"
End Property

