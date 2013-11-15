Imports System.IO

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

			Const adStateOpen = 1

			Dim cmdThousandFindColumns = CreateObject("ADODB.Command")
			cmdThousandFindColumns.CommandText = "spASRIntOrgChart"
			cmdThousandFindColumns.CommandType = 4 ' Stored Procedure
			cmdThousandFindColumns.ActiveConnection = HttpContext.Current.Session("databaseConnection")
			cmdThousandFindColumns.CommandTimeout = 180

			Dim prmRootID = cmdThousandFindColumns.CreateParameter("RootID", 3, 1)
			cmdThousandFindColumns.Parameters.Append(prmRootID)
			prmRootID.value = CleanNumeric(HttpContext.Current.Session("TopLevelRecID"))

			Err.Clear()

			Dim rstHierarchyRecords As ADODB.Recordset
			Dim sErrorDescription = ""

			Try
				rstHierarchyRecords = cmdThousandFindColumns.Execute
			Catch ex As Exception
				rstHierarchyRecords = Nothing
				sErrorDescription = "Error generating Organisation Chart." & vbCrLf & ex.Message
			End Try
			
			Dim orgCharts = New List(Of OrgChart)

			If Len(sErrorDescription) = 0 Then
				If rstHierarchyRecords.state = adStateOpen Then
					Dim additionalClasses As String

					Do While Not rstHierarchyRecords.EOF
						additionalClasses = " ui-corner-all"

						' highlight the current user's node
						If CType(rstHierarchyRecords.fields(0).value, String) = CType(HttpContext.Current.Session("TopLevelRecID"), String) Then
							additionalClasses &= " ui-state-active"
						Else
							additionalClasses &= " ui-state-default"
						End If

						' resize Photos to 48x48px
						Dim photoSource As String = ""

						If Not IsDBNull(rstHierarchyRecords.fields(7).value) Then
							Dim oleType As Short = Val(Encoding.UTF8.GetString(rstHierarchyRecords.fields(7).value, 8, 2))
							If oleType = 2 Then	'Embeded
								Dim abtImage = CType(rstHierarchyRecords.fields(7).value, Byte())
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
								Dim unc As String = Trim(Encoding.UTF8.GetString(rstHierarchyRecords.fields(7).value, 290, 60))
								Dim fileName As String = Trim(Path.GetFileName(Encoding.UTF8.GetString(rstHierarchyRecords.fields(7).value, 10, 70))).Replace("\", "/")
								Dim fullPath As String = Trim(Encoding.UTF8.GetString(rstHierarchyRecords.fields(7).value, 80, 210)).Replace("\", "/")
								photoSource = "file:///" & unc & "/" & fullPath & "/" & fileName
							End If
						Else 'No picture is defined for user, use anonymous one
							photoSource = "../Content/images/anonymous.png"
						End If


						orgCharts.Add(New OrgChart() With {
							.EmployeeID = rstHierarchyRecords.fields(0).value,
							.EmployeeForenames = rstHierarchyRecords.fields(1).value,
							.EmployeeSurname = rstHierarchyRecords.fields(2).value,
							.EmployeeStaffNo = rstHierarchyRecords.fields(3).value,
							.LineManagerStaffNo = rstHierarchyRecords.fields(4).value,
							.EmployeeJobTitle = rstHierarchyRecords.fields(5).value,
							.HierarchyLevel = rstHierarchyRecords.fields(6).value,
							.PhotoPath = photoSource,
							.AbsenceTypeClass = rstHierarchyRecords.fields(8).value & additionalClasses & " " &
															 rstHierarchyRecords.fields(9).value & " " &
															 rstHierarchyRecords.fields(10).value & " "})

						rstHierarchyRecords.moveNext()
					Loop

				End If
			End If

			Return orgCharts

		End Function


	End Class

End Namespace