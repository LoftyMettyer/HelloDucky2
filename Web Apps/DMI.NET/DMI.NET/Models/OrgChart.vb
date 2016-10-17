Option Explicit On
Option Strict Off

Imports System.IO
Imports System.Data.SqlClient
Imports HR.Intranet.Server

Namespace Models

	Public Class OrgChart

		Public Property EmployeeID() As Integer
		Public Property EmployeeForenames() As String
		Public Property EmployeeSurname() As String
		Public Property EmployeeStaffNo() As String
		Public Property LineManagerStaffNo() As String
		Public Property EmployeeJobTitle() As String
		Public Property HierarchyLevel() As Integer
		Public Property PhotoPath() As String
		Public Property AbsenceTypeClass() As String

		Public Function LoadModel() As List(Of OrgChart)

			Dim iLoggedInUser As Integer = CInt(CleanNumeric(HttpContext.Current.Session("LoggedInUserRecordID")))

			Dim objSession As SessionInfo = CType(HttpContext.Current.Session("SessionContext"), SessionInfo)
			Dim objDataAccess As New clsDataAccess(objSession.LoginInfo)

			Dim orgCharts = New List(Of OrgChart)
			Dim sErrorDescription = ""

			' User record now been identified
			If iLoggedInUser = 0 Then
				HttpContext.Current.Session("ErrorText") = "Current user not identified."
				Return orgCharts
			End If

			Try
				Dim rstHierarchyRecords = objDataAccess.GetDataTable("spASRIntOrgChart", CommandType.StoredProcedure _
							, New SqlParameter("RootID", SqlDbType.Int) With {.Value = iLoggedInUser})

				Dim additionalClasses As String

				If rstHierarchyRecords.Rows.Count = 0 Then
					' No records returned
					HttpContext.Current.Session("ErrorText") = "No matching records found."
				Else

					For Each objRow As DataRow In rstHierarchyRecords.Rows

						additionalClasses = " ui-corner-all"

						' highlight the current user's node
						If CInt(objRow(0)) = iLoggedInUser Then
							additionalClasses &= " ui-state-highlight"
						Else
							additionalClasses &= " ui-state-default"
						End If

						' resize Photos to 48x48px
						Dim photoSource As String = ""

						If Not IsDBNull(objRow(7)) And rstHierarchyRecords.Columns(7).DataType.Name.ToLower = "byte[]" Then

							Dim oleType As Short = Val(Encoding.UTF8.GetString(objRow(7), 8, 2))
							If oleType = 2 Then	'Embeded
								Dim abtImage = CType(objRow(7), Byte())
								Dim binaryData As Byte() = New Byte(abtImage.Length - 400) {}
								Try
									Buffer.BlockCopy(abtImage, 400, binaryData, 0, abtImage.Length - 400)
									'Create an image based on the embeded (Base64) image and resize it to 48x48

									Dim ms As New MemoryStream(binaryData)
									Dim img As Drawing.Image = Drawing.Image.FromStream(ms, True)

									img = img.GetThumbnailImage(48, 48, Nothing, IntPtr.Zero)
									photoSource = "data:image/jpeg;base64," & ImageToBase64String(img)
								Catch exp As ArgumentNullException
									photoSource = "../Content/images/anonymous.png"
								End Try
							ElseIf oleType = 3 Then 'Link
                        photoSource = "../Content/images/anonymous.png"
                     End If
						Else 'No picture is defined for user, use anonymous one
							photoSource = "../Content/images/anonymous.png"
						End If

						Dim sAbsenceReasonClass As String = ""
						If Not IsDBNull(objRow(9)) Then sAbsenceReasonClass = "REASON#" & objRow(9) & "#"

						orgCharts.Add(New OrgChart() With {
								.EmployeeID = CInt(objRow(0)),
								.EmployeeForenames = HttpUtility.HtmlEncode(objRow(1).ToString()),
								.EmployeeSurname = HttpUtility.HtmlEncode(objRow(2).ToString()),
								.EmployeeStaffNo = HttpUtility.HtmlEncode(objRow(3).ToString()),
								.LineManagerStaffNo = HttpUtility.HtmlEncode(objRow(4).ToString()),
								.EmployeeJobTitle = HttpUtility.HtmlEncode(objRow(5).ToString()),
								.HierarchyLevel = CInt(objRow(6)),
								.PhotoPath = photoSource,
								.AbsenceTypeClass = HttpUtility.HtmlEncode(objRow(8).ToString()) & HttpUtility.HtmlEncode(additionalClasses) & " " &
																 HttpUtility.HtmlEncode(sAbsenceReasonClass) & " " &
																 HttpUtility.HtmlEncode(objRow(10).ToString()) & " "})
					Next
				End If

			Catch ex As SqlException

				Select Case ex.Number
					Case 2812
						sErrorDescription = "The required setup for your organisation chart has not been completed."
					Case 217
						sErrorDescription = "There is a circular reference in your reporting structure."
					Case Else
						sErrorDescription = ex.Message

				End Select
				HttpContext.Current.Session("ErrorText") = sErrorDescription

			Catch ex As Exception
				HttpContext.Current.Session("ErrorText") = ex.Message

			End Try

			Return orgCharts

		End Function


	End Class

End Namespace