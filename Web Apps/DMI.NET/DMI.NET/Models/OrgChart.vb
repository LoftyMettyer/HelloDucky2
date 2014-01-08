Imports System.IO
Imports System.Data.SqlClient
Imports ADODB
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

			Const adStateOpen = 1


			Dim iTopLevelRecID As Integer = CleanNumeric(HttpContext.Current.Session("TopLevelRecID"))

			If iTopLevelRecID = 0 Then
				Dim iSingleRecordViewID As Integer = CleanNumeric(HttpContext.Current.Session("SingleRecordViewID"))

				Dim prmRecordID = New SqlParameter("piRecordID", SqlDbType.Int)
				prmRecordID.Direction = ParameterDirection.Output

				Dim prmRecordCount = New SqlParameter("piRecordCount", SqlDbType.Int)
				prmRecordCount.Direction = ParameterDirection.Output

				Err.Clear()
				clsDataAccess.GetDataSet("spASRIntGetSelfServiceRecordID", prmRecordID, prmRecordCount, New SqlParameter("piViewID", iSingleRecordViewID))

				If Err.Number = 0 And prmRecordCount.Value = 1 Then
					' Only one record.
					iTopLevelRecID = CLng(prmRecordID.Value)
				End If
			End If


			Dim cmdThousandFindColumns = CreateObject("ADODB.Command")
			cmdThousandFindColumns.CommandText = "spASRIntOrgChart"
			cmdThousandFindColumns.CommandType = 4 ' Stored Procedure
			cmdThousandFindColumns.ActiveConnection = HttpContext.Current.Session("databaseConnection")
			cmdThousandFindColumns.CommandTimeout = 180

			Dim prmRootID = cmdThousandFindColumns.CreateParameter("RootID", 3, 1)
			cmdThousandFindColumns.Parameters.Append(prmRootID)
			prmRootID.value = iTopLevelRecID

			Err.Clear()
			HttpContext.Current.Session("ErrorText") = ""

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
				If rstHierarchyRecords.State = adStateOpen Then
					Dim additionalClasses As String

					If rstHierarchyRecords.EOF And rstHierarchyRecords.BOF Then
						' No records returned
						sErrorDescription = "Error generating Organisation Chart. No matching records found."
					Else
						Do While Not rstHierarchyRecords.EOF
							additionalClasses = " ui-corner-all"

							' highlight the current user's node
							If CType(rstHierarchyRecords.Fields(0).Value, String) = CType(iTopLevelRecID, String) Then
								additionalClasses &= " ui-state-active"
							Else
								additionalClasses &= " ui-state-default"
							End If

							' resize Photos to 48x48px
							Dim photoSource As String = ""

							If Not IsDBNull(rstHierarchyRecords.Fields(7).Value) And rstHierarchyRecords.Fields(7).Type = DataTypeEnum.adVarBinary Then
								Dim oleType As Short = Val(Encoding.UTF8.GetString(rstHierarchyRecords.Fields(7).Value, 8, 2))
								If oleType = 2 Then	'Embeded
									Dim abtImage = CType(rstHierarchyRecords.Fields(7).Value, Byte())
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
									Dim unc As String = Trim(Encoding.UTF8.GetString(rstHierarchyRecords.Fields(7).Value, 290, 60))
									Dim fileName As String = Trim(Path.GetFileName(Encoding.UTF8.GetString(rstHierarchyRecords.Fields(7).Value, 10, 70))).Replace("\", "/")
									Dim fullPath As String = Trim(Encoding.UTF8.GetString(rstHierarchyRecords.Fields(7).Value, 80, 210)).Replace("\", "/")
									photoSource = "file:///" & unc & "/" & fullPath & "/" & fileName
								End If
							Else 'No picture is defined for user, use anonymous one
								photoSource = "../Content/images/anonymous.png"
							End If


							orgCharts.Add(New OrgChart() With {
									.EmployeeID = rstHierarchyRecords.Fields(0).Value,
									.EmployeeForenames = rstHierarchyRecords.Fields(1).Value,
									.EmployeeSurname = rstHierarchyRecords.Fields(2).Value,
									.EmployeeStaffNo = rstHierarchyRecords.Fields(3).Value,
									.LineManagerStaffNo = rstHierarchyRecords.Fields(4).Value,
									.EmployeeJobTitle = rstHierarchyRecords.Fields(5).Value,
									.HierarchyLevel = rstHierarchyRecords.Fields(6).Value,
								.PhotoPath = photoSource,
									.AbsenceTypeClass = rstHierarchyRecords.Fields(8).Value & additionalClasses & " " &
																	 rstHierarchyRecords.Fields(9).Value & " " &
																	 rstHierarchyRecords.Fields(10).Value & " "})

							rstHierarchyRecords.MoveNext()
						Loop
					End If

				End If
			End If

			If sErrorDescription.Length > 0 Then
				HttpContext.Current.Session("ErrorText") = sErrorDescription
			End If

			Return orgCharts

		End Function


	End Class

End Namespace