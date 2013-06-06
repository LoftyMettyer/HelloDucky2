VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'ODBC constants
Private Const SQL_DROP As Long = 1
Private Const SQL_NTS As Long = -3
Private Const SQL_NULL_DATA As Long = -1

'SQLGetInfo constants
Public Enum SQLInfo
  SQL_SERVER_NAME = 13
  SQL_DATABASE_NAME = 16
  SQL_DBMS_NAME = 17
  SQL_DBMS_VER = 18
  SQL_KEYWORDS = 89
  SQL_MAX_COLUMN_NAME_LEN = 30
  SQL_MAX_COLUMNS_IN_TABLE = 101
  SQL_MAX_TABLE_NAME_LEN = 35
End Enum

'ODBC datatype constants
Private Const SQL_SIGNED_OFFSET As Long = -20
Private Const SQL_CHAR As Long = 1
Private Const SQL_INTEGER As Long = 4
Private Const SQL_SMALLINT As Long = 5
Private Const SQL_C_CHAR As Long = SQL_CHAR
Private Const SQL_C_LONG As Long = SQL_INTEGER
Private Const SQL_C_SHORT As Long = SQL_SMALLINT
Private Const SQL_C_SSHORT As Long = SQL_C_SHORT + SQL_SIGNED_OFFSET
Private Const SQL_C_SLONG As Long = SQL_C_LONG + SQL_SIGNED_OFFSET

'ODBC return value constants
Private Const SQL_SUCCESS As Long = 0
Private Const SQL_SUCCESS_WITH_INFO As Long = 1

Private Const STR_LEN = 254

'ODBC API functions
Private Declare Function SQLAllocStmt Lib "odbc32.dll" (ByVal hDbc&, phstmt&) As Integer
Private Declare Function SQLFreeStmt Lib "odbc32.dll" (ByVal hstmt&, ByVal fOption%) As Integer
Private Declare Function SQLGetData Lib "odbc32.dll" (ByVal hstmt&, ByVal icol%, ByVal fCType%, ByVal RGBValue As String, ByVal cbValueMax&, pcbValue&) As Integer
Private Declare Function SQLColumns Lib "odbc32.dll" (ByVal hstmt&, ByVal szTblQualifier As String, ByVal cbTblQualifier%, ByVal szTblOwner As String, ByVal cbTblOwner%, ByVal szTblName As String, ByVal cbTblName%, ByVal szColName As String, ByVal cbColName%) As Integer
Private Declare Function SQLBindCol Lib "odbc32.dll" (ByVal hstmt&, ByVal icol%, ByVal fCType%, RGBValue As Any, ByVal cbValueMax&, pcbValue&) As Integer
Private Declare Function SQLFetch Lib "odbc32.dll" (ByVal hstmt&) As Integer
Private Declare Function SQLGetInfo Lib "odbc32.dll" (ByVal hDbc&, ByVal fInfoType%, ByRef rgbInfoValue As Any, ByVal cbInfoMax%, cbInfoOut%) As Integer
Private Declare Function SQLGetInfoString Lib "odbc32.dll" Alias "SQLGetInfo" (ByVal hDbc&, ByVal fInfoType%, ByVal rgbInfoValue As String, ByVal cbInfoMax%, cbInfoOut%) As Integer
Private Declare Function SQLGetTypeInfo Lib "odbc32.dll" (ByVal hstmt&, ByVal fSqlType%) As Integer
Private Declare Function SQLNativeSql Lib "odbc32.dll" (ByVal hDbc&, ByVal szSqlStrIn$, ByVal cbSqlStrIn&, ByVal szSqlStr$, ByVal cbSqlStrMax&, pcbSqlStr&) As Integer

Function FormatError(ByVal ErrorMsg As String) As String
'  On Error Resume Next
'
'  Dim intStart As Integer, intFound As Integer
'
'  intFound = 0
'  Do
'    intStart = intFound
'    intFound = InStr(intStart + 1, ErrorMsg, "]")
'  Loop While intFound > 0
'
'  If intStart > 0 And intStart < Len(Trim(ErrorMsg)) Then
'    FormatError = Left(ErrorMsg, intStart) & _
'      vbCr & vbCr & Trim(Mid(ErrorMsg, intStart + 1))
'  Else
'    FormatError = ErrorMsg
'  End If
  FormatError = Trim(Mid(ErrorMsg, InStrRev(ErrorMsg, "]") + 1))
  
End Function

Function GetColDefault(ByVal hDbc As Long, ByVal TableName As String, ByVal ColumnName As String) As String
  Dim lngStmtHnd As Long
  Dim intReturn As Integer
  Dim lngDefaultLen As Long
  Dim strColDefault As String
    
  'Get ODBC statement handle
  intReturn = SQLAllocStmt(hDbc, lngStmtHnd)
  If Success(intReturn) Then
    'Get info for the required column
    intReturn = SQLColumns(lngStmtHnd, vbNullString, 0, _
      vbNullString, 0, TableName, Len(TableName), _
      ColumnName, Len(ColumnName))
      
    If Success(intReturn) Then
      'Get the result set
      intReturn = SQLFetch(lngStmtHnd)
      If Success(intReturn) Then
        'Initialise string for column default
        strColDefault = String(STR_LEN + 1, 0)
        
        'Get the column default value
        intReturn = SQLGetData(lngStmtHnd, 13, SQL_C_CHAR, strColDefault, STR_LEN, lngDefaultLen)
        If Success(intReturn) Then
          If lngDefaultLen > 0 Then
            If Left(strColDefault, 1) = "'" Then
              GetColDefault = Mid(strColDefault, 2, lngDefaultLen - 2)
            Else
              GetColDefault = Left(strColDefault, lngDefaultLen)
            End If
          Else
            GetColDefault = vbNullString
          End If
        End If
      End If
    End If
    
    'Free the ODBC statement handle
    Call SQLFreeStmt(lngStmtHnd, SQL_DROP)
  End If
  
End Function

Function GetInfo(ByVal hDbc As Long, ByVal InfoType As SQLInfo) As Integer
  Dim intInfo As Integer
  Dim intReturn As Integer
  
  intInfo = 0
  intReturn = SQLGetInfo(hDbc, InfoType, intInfo, 2, SQL_NULL_DATA)
  If Success(intReturn) Then
    GetInfo = intInfo
  Else
    GetInfo = 0
  End If
  
End Function

Function GetInfoString(ByVal hDbc As Long, ByVal InfoType As SQLInfo, ByVal InfoLen As Integer) As String
  Dim strInfo As String
  Dim intReturn As Integer
  
  strInfo = String(InfoLen, 0)
  intReturn = SQLGetInfoString(hDbc, InfoType, strInfo, InfoLen, InfoLen)
  If Success(intReturn) Then
    GetInfoString = IIf(InfoLen > 0, Left(strInfo, InfoLen), vbNullString)
  Else
    GetInfoString = vbNullString
  End If
  
End Function

Function Success(ByVal ReturnValue As Integer) As Boolean
  Success = (ReturnValue = SQL_SUCCESS Or ReturnValue = SQL_SUCCESS_WITH_INFO)
End Function

Function NativeSQL(ByVal hDbc As Long, ByVal SQLString As String)
  Dim intReturn As Integer
  Dim strNative As String
  Dim lngNativeSize As Long
  
  lngNativeSize = 1024
  strNative = String(lngNativeSize + 1, 0)
  
  intReturn = SQLNativeSql(hDbc, SQLString, Len(SQLString), strNative, lngNativeSize, lngNativeSize)
  If Success(intReturn) And lngNativeSize > 0 Then
    NativeSQL = Left(strNative, lngNativeSize)
  Else
    NativeSQL = vbNullString
  End If
  
End Function

'Function GetColCreateString(ByVal hDbc As Long, ByVal psColumnName As String, ByVal DataType As Long, ByVal Size As Integer, ByVal Decimals As Integer) As String
'  Dim lngStmtHnd As Long
'  Dim intReturn As Integer
'  Dim strType As String
'  Dim lngTypeLen As Long
'  Dim strCreate As String
'  Dim lngCreateLen As Long
'  Dim strColCreate As String
'
'  strColCreate = vbNullString
'
'  'Get ODBC statement handle
'  intReturn = SQLAllocStmt(hDbc, lngStmtHnd)
'  If Success(intReturn) Then
'    'Get info for the required data type
'    intReturn = SQLGetTypeInfo(lngStmtHnd, DataType)
'
'    If Success(intReturn) Then
'      'Get the result set
'      intReturn = SQLFetch(lngStmtHnd)
'      If Success(intReturn) Then
'        'Initialise string for column default
'        strType = String(STR_LEN + 1, 0)
'        strCreate = String(STR_LEN + 1, 0)
'
'        'Get the column default value
'        Call SQLGetData(lngStmtHnd, 1, SQL_C_CHAR, strType, STR_LEN, lngTypeLen)
'        Call SQLGetData(lngStmtHnd, 6, SQL_C_CHAR, strCreate, STR_LEN, lngCreateLen)
'        If lngTypeLen > 0 Then
'          strType = UCase(Left(strType, lngTypeLen))
'          'strCreate = IIf(lngCreateLen > 0, UCase(Left(strCreate, lngCreateLen)), vbNullString)
'          If lngCreateLen > 0 Then
'            strCreate = UCase(Left(strCreate, lngCreateLen))
'          Else
'            strCreate = vbNullString
'          End If
'
'          strColCreate = psColumnName & Space(1) & strType
'          If InStr(1, strCreate, "LENGTH") > 0 Or InStr(1, strCreate, "PRECISION") > 0 Then
'            strColCreate = strColCreate & "(" & Size
'            If InStr(1, strCreate, "SCALE") > 0 Then
'              strColCreate = strColCreate & "," & Decimals
'            End If
'            strColCreate = strColCreate & ")"
'          End If
'        End If
'      End If
'    End If
'
'    'Free the ODBC statement handle
'    Call SQLFreeStmt(lngStmtHnd, SQL_DROP)
'  End If
'
'  GetColCreateString = strColCreate
'
'End Function

