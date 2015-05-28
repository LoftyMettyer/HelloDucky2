Option Explicit On
Option Strict On

Imports System.Collections.Generic
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.BaseClasses
Imports System.Data.SqlClient
Imports HR.Intranet.Server.Metadata

Public Class Database
	Inherits BaseForDMI

	Friend UserSettings As ICollection(Of UserSetting)

	Public Function GetEmailAddress(lngRecordID As Integer, lngEmailAddrCalc As Integer) As String

		' Check if the user can create New instances of the given category.

		Try

			Dim prmResult = New SqlParameter("hResult", SqlDbType.VarChar, 8000) With {.Direction = ParameterDirection.Output}

			DB.ExecuteSP("spASRSysEmailAddr", prmResult _
										, New SqlParameter("@EmailID", SqlDbType.Int) With {.Value = lngEmailAddrCalc} _
										, New SqlParameter("@recordID", SqlDbType.Int) With {.Value = lngRecordID})

			Return prmResult.Value.ToString()

		Catch ex As Exception
            Return ""

        End Try

    End Function


    Public Sub SaveUserSetting(strSection As String, strKey As String, varSetting As Object)

        Dim objSetting As New UserSetting
        UserSettings = SessionInfo.UserSettings

        Try

            DB.ExecuteSP("sp_ASRIntSaveSetting" _
                    , New SqlParameter("psSection", SqlDbType.VarChar, 255) With {.Value = strSection} _
                    , New SqlParameter("psKey", SqlDbType.VarChar, 255) With {.Value = strKey} _
                    , New SqlParameter("pfUserSetting", SqlDbType.Bit) With {.Value = True} _
                    , New SqlParameter("psValue", SqlDbType.VarChar, -1) With {.Value = varSetting})

            ' Update UserSettings collection as this is what's actually used post-login.
            If UserSettings.GetUserSetting(strSection, strKey) Is Nothing Then
                objSetting.Section = strSection
                objSetting.Key = strKey
                objSetting.Value = varSetting
                UserSettings.Add(objSetting)
            Else
                objSetting = UserSettings.GetUserSetting(strSection, strKey)
                objSetting.Value = varSetting
            End If

        Catch ex As Exception
            Throw

        End Try

    End Sub

	''' <summary>
	''' Saves the system settings
	''' </summary>
	''' <param name="strSection">The Section</param>
	''' <param name="strKey">The Key</param>
	''' <param name="varSetting">The Value to store</param>
	Public Sub SaveSystemSetting(strSection As String, strKey As String, varSetting As Object)

		Try

			DB.ExecuteSP("spsys_setsystemsetting" _
					, New SqlParameter("section", SqlDbType.VarChar, 255) With {.Value = LCase(strSection)} _
					, New SqlParameter("settingkey", SqlDbType.VarChar, 255) With {.Value = LCase(strKey)} _
					, New SqlParameter("settingvalue", SqlDbType.VarChar, -1) With {.Value = varSetting})

		Catch ex As Exception
			Throw

		End Try

	End Sub

	Public Function GetUserSetting(strSection As String, strKey As String, varDefault As Object) As Object

		Dim prmResult = New SqlParameter("psResult", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

		Try

			DB.ExecuteSP("spASRIntGetSetting" _
					, New SqlParameter("psSection", SqlDbType.VarChar, -1) With {.Value = strSection} _
					, New SqlParameter("psKey", SqlDbType.VarChar, -1) With {.Value = strKey} _
					, New SqlParameter("psDefault", SqlDbType.VarChar, -1) With {.Value = varDefault} _
					, New SqlParameter("pfUserSetting", SqlDbType.Bit) With {.Value = True} _
					, prmResult)

		Catch ex As Exception
			Return varDefault

		End Try

		Return prmResult.Value

	End Function

	Public Function GetSystemSetting(strSection As String, strKey As String, varDefault As Object) As Object

		Dim prmResult = New SqlParameter("psResult", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

		Try

			DB.ExecuteSP("spASRIntGetSetting" _
					, New SqlParameter("psSection", SqlDbType.VarChar, -1) With {.Value = strSection} _
					, New SqlParameter("psKey", SqlDbType.VarChar, -1) With {.Value = strKey} _
					, New SqlParameter("psDefault", SqlDbType.VarChar, -1) With {.Value = varDefault} _
					, New SqlParameter("pfUserSetting", SqlDbType.Bit) With {.Value = False} _
					, prmResult)

		Catch ex As Exception
			Return varDefault

		End Try

		Return prmResult.Value

	End Function

	Public Function GetRecordTimestamp(RecordID As Integer, RealSource As String) As Integer

		Dim prmTimestamp As New SqlParameter("piTimestamp", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

		Try
			DB.ExecuteSP("spASRIntGetTimestamp" _
				, prmTimestamp _
				, New SqlParameter("piRecordID", SqlDbType.Int) With {.Value = RecordID} _
				, New SqlParameter("psRealsource", SqlDbType.VarChar, 255) With {.Value = RealSource})

			Return CInt(prmTimestamp.Value)

		Catch ex As Exception
			Return 0

		End Try

	End Function

	Public Function GetTableOrders(TableID As Integer, ViewID As Integer) As DataTable

		Return DB.GetDataTable("sp_ASRIntGetTableOrders", CommandType.StoredProcedure _
				, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = TableID} _
				, New SqlParameter("piViewID", SqlDbType.Int) With {.Value = ViewID})

	End Function

	Public Function GetUtilityUsage(UtilType As UtilityType, ID As Integer) As DataTable

		Return DB.GetDataTable("sp_ASRIntDefUsage", CommandType.StoredProcedure _
				, New SqlParameter("intType", SqlDbType.Int) With {.Value = UtilType} _
				, New SqlParameter("intID", SqlDbType.Int) With {.Value = ID})

	End Function

End Class
