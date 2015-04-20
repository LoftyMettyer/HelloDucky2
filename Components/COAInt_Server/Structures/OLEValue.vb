Option Strict On
Option Explicit On

Imports System.Text
Imports System.IO

Namespace Structures
	Public Class OLEValue
		Public ColumnID As Integer
		Public Value As Byte()

		Private _miOLEType As Short
		Private _mstrFileName As String
		Private _mstrPath As String
		Private _mstrUnc As String
		Public DocumentSize As String
		Public FileCreateDate As String
		Public FileModifyDate As String

		Public ReadOnly Property FileName As String
			Get
				If _miOLEType = 2 Then
					Return _mstrFileName & "::EMBEDDED_OLE_DOCUMENT::"
				Else
					Return _mstrFileName & "::LINKED_OLE_DOCUMENT::"
				End If
			End Get
		End Property

		Public Sub ExtractProperties()

			Try

				If Not Value Is Nothing Then
					'_msOLEVersionType = Encoding.UTF8.GetString(objBytes, 0, 8)
					_miOLEType = CShort(Encoding.UTF8.GetString(Value, 8, 2))
					'_mstrDisplayFileName = Trim(Path.GetFileName(Encoding.UTF8.GetString(objBytes, 10, 70)))
					_mstrFileName = Trim(Path.GetFileName(Encoding.UTF8.GetString(Value, 10, 70)))
					_mstrPath = Trim(Encoding.UTF8.GetString(Value, 80, 210))
					_mstrUnc = Trim(Encoding.UTF8.GetString(Value, 290, 60))
					DocumentSize = Trim(Encoding.UTF8.GetString(Value, 350, 10))
					FileCreateDate = Trim(Encoding.UTF8.GetString(Value, 360, 20))
					FileModifyDate = Trim(Encoding.UTF8.GetString(Value, 380, 20))
				End If

			Catch ex As Exception
				Throw

			End Try

		End Sub


		Public Function ConvertPhotoToBase64() As String

			Try

				If _miOLEType = 2 Then

					Dim abtImage = Value
					Dim binaryData As Byte() = New Byte(abtImage.Length - 400) {}
					Try
						Buffer.BlockCopy(abtImage, 400, binaryData, 0, abtImage.Length - 400)
						Return Convert.ToBase64String(binaryData, 0, binaryData.Length)

					Catch exp As ArgumentNullException
						Console.WriteLine("Binary data array is null.")

					End Try
				Else
					If _mstrPath.Length > 0 AndAlso _mstrPath.Substring(0, 2) = "\\" Then
						' Return _mstrPath & "\" & _mstrFileName & "::LINKED_OLE_DOCUMENT::" & vbTab & DocumentSize & vbTab & FileCreateDate & vbTab & FileModifyDate
						Return _mstrPath & "\" & _mstrFileName
					Else
						' Return _mstrUnc & _mstrPath & "\" & _mstrFileName & "::LINKED_OLE_DOCUMENT::" & vbTab & DocumentSize & vbTab & FileCreateDate & vbTab & FileModifyDate
						Return _mstrUnc & _mstrPath & "\" & _mstrFileName
					End If

				End If

			Catch ex As Exception
				Throw

			End Try

			Return ""

		End Function



	End Class
End Namespace
