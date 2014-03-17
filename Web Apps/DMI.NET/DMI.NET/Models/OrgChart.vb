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

			Dim iTopLevelRecID As Integer = CInt(CleanNumeric(HttpContext.Current.Session("TopLevelRecID")))

			Dim objSession As SessionInfo = CType(HttpContext.Current.Session("SessionContext"), SessionInfo)
			Dim objDataAccess As New clsDataAccess(objSession.LoginInfo)

			If iTopLevelRecID = 0 Then
				If HttpContext.Current.Session("LoggedInUserRecordID") > 0 Then
					iTopLevelRecID = HttpContext.Current.Session("LoggedInUserRecordID")
				End If
			End If

			Dim orgCharts = New List(Of OrgChart)
			Dim sErrorDescription = ""

			Try
				Dim rstHierarchyRecords = objDataAccess.GetDataTable("spASRIntOrgChart", CommandType.StoredProcedure _
							, New SqlParameter("RootID", SqlDbType.Int) With {.Value = iTopLevelRecID})



				Dim additionalClasses As String

				If rstHierarchyRecords.Rows.Count = 0 Then
					' No records returned
					sErrorDescription = "Error generating Organisation Chart. No matching records found."
				Else

					For Each objRow As DataRow In rstHierarchyRecords.Rows

						additionalClasses = " ui-corner-all"

						' highlight the current user's node
						If CInt(objRow(0)) = iTopLevelRecID Then
							additionalClasses &= " ui-state-active"
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
							ElseIf oleType = 3 Then	'Link
								Dim unc As String = Trim(Encoding.UTF8.GetString(objRow(7), 290, 60))
								Dim fileName As String = Trim(Path.GetFileName(Encoding.UTF8.GetString(objRow(7), 10, 70))).Replace("\", "/")
								Dim fullPath As String = Trim(Encoding.UTF8.GetString(objRow(7), 80, 210)).Replace("\", "/")
								photoSource = "file:///" & unc & "/" & fullPath & "/" & fileName
							End If
						Else 'No picture is defined for user, use anonymous one
							photoSource = "../Content/images/anonymous.png"
						End If


						orgCharts.Add(New OrgChart() With {
								.EmployeeID = CInt(objRow(0)),
								.EmployeeForenames = objRow(1).ToString(),
								.EmployeeSurname = objRow(2).ToString(),
								.EmployeeStaffNo = objRow(3).ToString(),
								.LineManagerStaffNo = objRow(4).ToString(),
								.EmployeeJobTitle = objRow(5).ToString(),
								.HierarchyLevel = CInt(objRow(6)),
								.PhotoPath = photoSource,
								.AbsenceTypeClass = objRow(8).ToString() & additionalClasses & " " &
																 objRow(9).ToString() & " " &
																 objRow(10).ToString() & " "})

					Next
				End If


			Catch ex As SqlException

				Select Case ex.Number
					Case 217
						sErrorDescription = "There is a circular reference in your reporting structure."
					Case Else
						sErrorDescription = "Error generating Organisation Chart." & vbCrLf & ex.Message

				End Select
				HttpContext.Current.Session("ErrorText") = sErrorDescription

			Catch ex As Exception
				sErrorDescription = "Error generating Organisation Chart." & vbCrLf & ex.Message
				HttpContext.Current.Session("ErrorText") = sErrorDescription

			End Try


			Return orgCharts

		End Function


	End Class

End Namespace