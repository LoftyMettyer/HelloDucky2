VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOutlookFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
  
Private mlngFolderID As Long
Private mlngTableID As Long
Private mstrName As String
Private mlngFolderType As Long
Private mstrFixedPath As String
Private mlngCalcExprID As Long

Private mfNew As Boolean
Private mfChanged As Boolean
Private mfDeleted As Boolean


Public Property Get FolderID() As Long
  FolderID = mlngFolderID
End Property

Public Property Let FolderID(ByVal lngNewValue As Long)
  mlngFolderID = lngNewValue
End Property

Public Property Get TableID() As Long
  TableID = mlngTableID
End Property

Public Property Let TableID(ByVal lngNewValue As Long)
  mlngTableID = lngNewValue
End Property

Public Property Get Name() As String
  Name = mstrName
End Property

Public Property Let Name(ByVal strNewValue As String)
  mstrName = strNewValue
End Property

Public Property Get FolderType() As Long
  FolderType = mlngFolderType
End Property

Public Property Let FolderType(ByVal lngNewValue As Long)
  mlngFolderType = lngNewValue
End Property

Public Property Get FixedPath() As String
  FixedPath = mstrFixedPath
End Property

Public Property Let FixedPath(ByVal strNewValue As String)
  mstrFixedPath = strNewValue
End Property

Public Property Get CalcExprID() As Long
  CalcExprID = mlngCalcExprID
End Property

Public Property Let CalcExprID(ByVal lngNewValue As Long)
  mlngCalcExprID = lngNewValue
End Property


Public Function SelectOutlookFolder() As Boolean

  Dim frmSelectOutlookFolder As frmSelection
  Dim lngOldFolderID As Long
  Dim lngTableID As Long
  Dim fExit As Boolean
  Dim strSQL As String

  SelectOutlookFolder = False

  If TableID > 0 Then

    fExit = False
    lngOldFolderID = mlngFolderID
    lngTableID = Me.TableID
    Set frmSelectOutlookFolder = New frmSelection

    Do While Not fExit
      ' Set the selection form type.
      ' Set the selection form properties.
      Me.TableID = lngTableID
      Set frmSelectOutlookFolder.OutlookFolder = Me
      frmSelectOutlookFolder.Show vbModal

      Select Case frmSelectOutlookFolder.Action
        Case edtAdd
          NewOutlookFolder

        Case edtEdit
          FolderID = frmSelectOutlookFolder.SelectedID
          EditOutlookFolder False

        Case edtCopy
          FolderID = frmSelectOutlookFolder.SelectedID
          EditOutlookFolder True

        Case edtDelete
          'TM20020528 Fault 3924
          If MsgBox("Delete this outlook folder reference, are you sure?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
            FolderID = frmSelectOutlookFolder.SelectedID
            DeleteOutlookFolder
          End If

        Case edtPrint
          FolderID = frmSelectOutlookFolder.SelectedID
          PrintFolderDefinition

        Case edtSelect
          FolderID = frmSelectOutlookFolder.SelectedID
          SelectOutlookFolder = True
          fExit = True

        ' Deselect all expressions and return to the calling form.
        Case edtDeselect
          FolderID = -1
          fExit = True

        Case edtCancel
          ' Check if the original OutlookFolder still exists.
          With recOutlookFolders
            .Index = "idxFolderID"
            .Seek "=", lngOldFolderID

            If .NoMatch Then
              FolderID = 0
            Else
              If !Deleted Then
                FolderID = 0
              End If
            End If
          End With
          
          fExit = True
      
      End Select
    Loop

    Set frmSelectOutlookFolder = Nothing
  End If

End Function


Private Function NewOutlookFolder()

  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim frmEdit As frmOutlookFolder
  
  mfNew = True
  mfChanged = False
  mfDeleted = False
  mlngFolderID = 0
  ConstructFolder

  Set frmEdit = New frmOutlookFolder
  frmEdit.Initialise Me, False
  frmEdit.Show vbModal
  fOK = Not frmEdit.Cancelled

  daoWS.BeginTrans

  If fOK Then

    With frmEdit.OutlookFolder
'      .FolderID
      .TableID = mlngTableID
      .Name = mstrName
      .FolderType = mlngFolderType
      .FixedPath = mstrFixedPath
      .CalcExprID = mlngCalcExprID
      .IsNew = True

      fOK = WriteFolder

    End With
  End If
  UnLoad frmEdit
  Set frmEdit = Nothing


TidyUpAndExit:
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
  Else
    daoWS.Rollback
  End If
  NewOutlookFolder = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function


Private Function EditOutlookFolder(blnCopy As Boolean)

  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim frmEdit As frmOutlookFolder
  
  ' Begin the transaction of data to the local database.
  daoWS.BeginTrans
  
  fOK = ConstructFolder
  
  If fOK Then
    Set frmEdit = New frmOutlookFolder
    
    frmEdit.Initialise Me, blnCopy
    frmEdit.Show vbModal
    
    fOK = Not frmEdit.Cancelled
      
    If fOK Then
      
      With frmEdit.OutlookFolder
        .TableID = mlngTableID
        .Name = mstrName
        .FolderType = mlngFolderType
        .FixedPath = mstrFixedPath
        .CalcExprID = mlngCalcExprID
        .IsNew = blnCopy
        mfChanged = True
      End With

      fOK = WriteFolder
    End If

    UnLoad frmEdit
    Set frmEdit = Nothing
  End If
  
TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
  Else
    daoWS.Rollback
  End If
  EditOutlookFolder = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function


Public Function WriteFolder() As Boolean

  On Error GoTo ErrorTrap

  Dim lngNewID As Long
  Dim fOK As Boolean
  Dim lngCount As Long

  fOK = True

  With recOutlookFolders

    If mlngFolderID = 0 Then
      mlngFolderID = Database.UniqueColumnValue("tmpOutlookFolders", "FolderID")
      .AddNew
    Else
      FindFolder
      .Edit
    End If

    !FolderID = mlngFolderID
    !TableID = mlngTableID
    !Name = mstrName
    !FolderType = mlngFolderType
    !FixedPath = mstrFixedPath
    !ExprID = mlngCalcExprID

    !New = mfNew
    !Changed = mfChanged
    !Deleted = mfDeleted

    .Update

  End With


TidyUpAndExit:
  WriteFolder = fOK
  Exit Function

ErrorTrap:
  If ASRDEVELOPMENT Then
    MsgBox Err.Description, vbCritical, "ASRDEVELOPMENT"
    Stop
  End If
  
  fOK = False
  Err = False
  Resume TidyUpAndExit

End Function


Private Function ConstructFolder()

  If FindFolder Then
    ConstructFolder = ReadFolder
  Else
    ConstructFolder = False
    mlngFolderID = 0
    'mlngTableID = 0
    mstrName = vbNullString
    mlngFolderType = 0
    mstrFixedPath = vbNullString
    mlngCalcExprID = 0
  End If

End Function


Private Function FindFolder() As Boolean
  On Error GoTo ErrorTrap
  
  With recOutlookFolders
    .Index = "idxFolderID"
    .Seek "=", FolderID
    FindFolder = (Not .NoMatch)
  End With
  
  Exit Function
  
ErrorTrap:
  FindFolder = False

End Function


Private Function ReadFolder() As Boolean
  ' Read the Email Folder definition from the Email Folder recordset.
  On Error GoTo ErrorTrap

  Dim fReadOk As Boolean

  fReadOk = False

  With recOutlookFolders

    mlngFolderID = !FolderID
    'mlngTableID = !TableID
    mstrName = !Name
    mlngFolderType = !FolderType
    mstrFixedPath = IIf(IsNull(!FixedPath), vbNullString, !FixedPath)
    mlngCalcExprID = !ExprID
    mfNew = !New
    mfChanged = !Changed
    mfDeleted = !Deleted

  End With
  
  
  fReadOk = True
  
TidyUpAndExit:
  ReadFolder = fReadOk
  Exit Function

ErrorTrap:
  fReadOk = False
  Err = False
  Resume TidyUpAndExit

End Function


Private Function DeleteOutlookFolder() As Boolean

  Dim mfrmUse As frmUsage

  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  
  fOK = True
  
  If fOK Then
    Set mfrmUse = New frmUsage
    mfrmUse.ResetList
    If FolderIsUsed(mfrmUse) Then
      Screen.MousePointer = vbDefault
      mfrmUse.ShowMessage Me.Name & " Folder", "The Folder cannot be deleted as the Folder is used by the following:", UsageCheckObject.OutlookFolder
      fOK = False
    End If
    UnLoad mfrmUse
    Set mfrmUse = Nothing
  End If

  If fOK Then

    fOK = ConstructFolder
    
    If fOK Then
      If mfDeleted = False Then
        mfDeleted = True
        fOK = WriteFolder
      End If
    End If
  
  End If
  
TidyUpAndExit:
  If fOK Then
    mfDeleted = True
  End If
  DeleteOutlookFolder = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Property Get Deleted() As Boolean
  Deleted = mfDeleted
End Property

Public Property Let Deleted(ByVal blnNewValue As Boolean)
  mfDeleted = blnNewValue
End Property


Private Function FolderIsUsed(mfrmUse As frmUsage) As Boolean

  Dim rsOutlook As DAO.Recordset
  Dim strSQL As String
  Dim fUsed As Boolean

  fUsed = False

  'MH20040324 Outlook Calendar Link folder
  strSQL = "SELECT DISTINCT tmpOutlookLinks.Title, tmpOutlookLinks.Deleted, tmpTables.TableName " & _
         "FROM tmpOutlookLinksDestinations, tmpOutlookLinks, tmpOutlookFolders, tmpTables " & _
         "WHERE tmpOutlookLinksDestinations.FolderID = " & CStr(Me.FolderID) & _
         " AND tmpOutlookLinksDestinations.LinkID = tmpOutlookLinks.LinkID " & _
         " AND tmpOutlookLinks.TableID = tmpTables.TableID " & _
         " AND tmpOutlookLinksDestinations.FolderID = tmpOutlookFolders.FolderID " & _
         " AND (tmpOutlookLinks.TableID = tmpOutlookFolders.TableID OR tmpOutlookFolders.TableID = 0)" & _
         " AND tmpOutlookLinks.Deleted = False"
  Set rsOutlook = daoDb.OpenRecordset(strSQL)
  With rsOutlook
    If Not (.EOF And .BOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList "Outlook Calendar Link : " & !Title & " <" & !TableName & ">"
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsOutlook = Nothing
  
  
  FolderIsUsed = fUsed

End Function


Public Sub PrintFolderDefinition()

  ' Print the expression.
  On Error GoTo ErrorTrap
  
  Dim sSQL As String
  Dim fOK As Boolean
  Dim sTableName As String
  Dim objComponent As CExprComponent
  Dim rsInfo As DAO.Recordset
  Dim objPrinter As clsPrintDef
  Dim lngBottom As Long

  glngPageNum = 0

  fOK = ConstructFolder

  ' Load the printer object
  Set objPrinter = New clsPrintDef
  With objPrinter
    If .IsOK Then
      If .PrintStart(True) Then
        .PrintHeader "Outlook Folder Definition : " & mstrName
  
        If mlngFolderType = 0 Then
          .PrintNormal "Type : Fixed"
          .PrintNormal "Folder : " & mstrFixedPath
        Else
          .PrintNormal "Type : Calculated"
          .PrintNormal "Table Name : " & GetTableName(mlngTableID)
          .PrintNormal "Calculation Name : " & GetExpressionName(mlngCalcExprID)
        End If
        
        
        lngBottom = CalculateBottomOfPage
        Printer.CurrentY = lngBottom + 1
        CheckEndOfPage2 lngBottom, False
        Printer.EndDoc
        
        Dim objDefPrinter As cSetDfltPrinter
        
        Set objDefPrinter = New cSetDfltPrinter
        Do
          objDefPrinter.SetPrinterAsDefault gstrDefaultPrinterName
        Loop While Printer.DeviceName <> gstrDefaultPrinterName
        Set objDefPrinter = Nothing
        
        Set objComponent = Nothing
      End If
    End If
  End With
  Set objPrinter = Nothing
  
TidyUpAndExit:
  If Not fOK Then
    MsgBox "Unable to print the outlook folder '" & Name & "'." & vbCrLf & vbCrLf & _
      Err.Description, vbExclamation + vbOKOnly, App.ProductName
  End If
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub

Public Property Get IsNew() As Boolean
  IsNew = mfNew
End Property

Public Property Let IsNew(ByVal blnNewValue As Boolean)
  mfNew = blnNewValue
End Property


Public Function CloneOutlookFolder(pavCloneRegister As Variant) As clsOutlookFolder
  
  ' Clone the current Outlook.
  ' Used when copying table definitions.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim lngNewColumnID As Long
  Dim objClonedOutlook As clsOutlookFolder
  Dim lngNewExprID As Long
  
  ' Read the Outlook definition.
  fOK = ConstructFolder

  If fOK Then
    ' Instantiate the clone Outlook.
    Set objClonedOutlook = New clsOutlookFolder
    ' Clone the current Outlook's properties.
    With objClonedOutlook
      .IsNew = True
      .Deleted = False
      '.Changed = False
      .FolderID = 0
      .Name = Me.Name
      .FolderType = Me.FolderType
      .FixedPath = Me.FixedPath

      .CalcExprID = 0
      If Me.CalcExprID > 0 Then
        For iIndex = 1 To UBound(pavCloneRegister, 2)
          If pavCloneRegister(1, iIndex) = "EXPRESSION" And _
            pavCloneRegister(2, iIndex) = Me.CalcExprID Then
            .CalcExprID = pavCloneRegister(3, iIndex)
            Exit For
          End If
        Next iIndex
      End If
    
    End With
  End If
  
TidyUpAndExit:
  If fOK Then
    Set CloneOutlookFolder = objClonedOutlook
  Else
    Set CloneOutlookFolder = Nothing
  End If
  ' Disassociate object variables.
  Set objClonedOutlook = Nothing
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function
