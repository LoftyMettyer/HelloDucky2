﻿Option Explicit On
Option Strict On

Imports HR.Intranet.Server.BaseClasses
Imports System.Data.SqlClient
Imports HR.Intranet.Server.Structures

Public Class Database
	Inherits BaseForDMI

	Public Function GetEmailAddress(lngRecordID As Integer, lngEmailAddrCalc As Integer) As String

		' Check if the user can create New instances of the given category.

		Try

			Dim prmResult = New SqlParameter("Result", SqlDbType.VarChar)
			prmResult.Direction = ParameterDirection.Output

			Dim prmEmail = New SqlParameter("EmailID", SqlDbType.Int)
			prmEmail.Value = lngEmailAddrCalc

			Dim prmRecordID = New SqlParameter("RecordID", SqlDbType.Int)
			prmRecordID.Value = lngRecordID

			DB.ExecuteSP("spASRSysEmailAddr", prmResult, prmEmail, prmRecordID)

			Return prmResult.Value.ToString()

		Catch ex As Exception
			Return ""

		End Try

	End Function

	' Return 0 or an error code
	Public Sub CheckLogin(ByRef Login As LoginInfo, ApplicationVersion As String)

		Try

			Dim prmSuccessFlag = New SqlParameter("piSuccessFlag", SqlDbType.Int)
			prmSuccessFlag.Direction = ParameterDirection.Output

			Dim prmErrorMessage = New SqlParameter("psErrorMessage", SqlDbType.VarChar, 255)
			prmErrorMessage.Direction = ParameterDirection.Output

			' Yes, I know - this parameter is spelt incorrectly (Not my fault). To rectify mean regenerating the stored proc!
			Dim prmMinPasswordLength = New SqlParameter("piMinPassordLength", SqlDbType.Int)
			prmMinPasswordLength.Direction = ParameterDirection.Output

			Dim prmIntranetVersion = New SqlParameter("psIntranetAppVersion", SqlDbType.VarChar, 50)
			prmIntranetVersion.Value = ApplicationVersion

			Dim prmPasswordLength = New SqlParameter("piPasswordLength", SqlDbType.Int)
			prmPasswordLength.Value = Len(Login.Password)

			Dim prmUserType = New SqlParameter("piUserType", SqlDbType.Int)
			prmUserType.Direction = ParameterDirection.Output

			Dim prmUserGroup = New SqlParameter("psUserGroup", SqlDbType.VarChar, 250)
			prmUserGroup.Direction = ParameterDirection.Output
			prmUserGroup.Value = ""

			Dim prmSelfServiceUserType = New SqlParameter("iSelfServiceUserType", SqlDbType.Int)
			prmSelfServiceUserType.Direction = ParameterDirection.Output

			DB.ExecuteSP("sp_ASRIntCheckLogin", prmSuccessFlag, prmErrorMessage, prmMinPasswordLength, prmIntranetVersion, prmPasswordLength, prmUserType, prmUserGroup, prmSelfServiceUserType)

			Login.UserType = CInt(prmUserType.Value)
			Login.SelfServiceUserType = CInt(prmSelfServiceUserType.Value)
			Login.UserGroup = prmUserGroup.Value.ToString()
			Login.LoginFailReason = prmErrorMessage.Value.ToString()

		Catch ex As Exception
			Throw

		End Try


	End Sub

	Public Sub LogOut()

		Try

			Dim prmLoggingIn As New SqlParameter("blnLoggingIn", SqlDbType.Bit)
			prmLoggingIn.Value = False

			Dim prmUser As New SqlParameter("strUsername", SqlDbType.VarChar, 1000)
			prmUser.Value = Replace(_login.Username, "'", "''")

			DB.ExecuteSP("sp_ASRIntAuditAccess", prmLoggingIn, prmUser)

		Catch ex As Exception

		End Try

	End Sub

	Public Sub SaveUserSetting(strSection As String, strKey As String, varSetting As Object)

		Try

			DB.ExecuteSP("sp_ASRIntSaveSetting" _
					, New SqlParameter("psSection", SqlDbType.VarChar, 255) With {.Value = strSection} _
					, New SqlParameter("psKey", SqlDbType.VarChar, 255) With {.Value = strKey} _
					, New SqlParameter("pfUserSetting", SqlDbType.Bit) With {.Value = True} _
					, New SqlParameter("psValue", SqlDbType.VarChar, -1) With {.Value = varSetting})

		Catch ex As Exception
			Throw

		End Try

	End Sub

	Public Function GetUserSetting(strSection As String, strKey As String, varDefault As Object) As Object

		Dim prmResult = New SqlParameter("psResult", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

		Try

			DB.ExecuteSP("sp_ASRIntGetSetting" _
					, New SqlParameter("psSection", SqlDbType.VarChar, -1) With {.Value = strSection} _
					, New SqlParameter("psKey", SqlDbType.VarChar, -1) With {.Value = strKey} _
					, New SqlParameter("psDefault", SqlDbType.VarChar, -1) With {.Value = varDefault} _
					, New SqlParameter("pfUserSetting", SqlDbType.Bit) With {.Value = True} _
					, prmResult)

		Catch ex As Exception
			Return varDefault

		End Try

		Return prmResult.value

	End Function

End Class
