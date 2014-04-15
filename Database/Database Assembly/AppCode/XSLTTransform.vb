Imports System.IO
Imports System.Xml
Imports System.Xml.Xsl
Imports System.Data.SqlTypes
Imports Microsoft.SqlServer.Server

Public Class XSLTTransform

	<SqlFunction(Name:="udfASRNetApplyXsltTransform")> _
	Public Shared Function XMLTransform(inputDataXML As SqlXml, inputTransformXML As SqlXml) As SqlXml

		Try

			Dim memoryXml As MemoryStream = New System.IO.MemoryStream()
			Dim xslt As New XslCompiledTransform()
			Dim output As XmlReader = Nothing

			xslt.Load(inputTransformXML.CreateReader())

			' Output the newly constructed XML
			Dim outputWriter As New XmlTextWriter(memoryXml, System.Text.Encoding.[Default])
			xslt.Transform(inputDataXML.CreateReader(), Nothing, outputWriter, Nothing)
			memoryXml.Seek(0, System.IO.SeekOrigin.Begin)
			output = New XmlTextReader(memoryXml)

			Return New SqlXml(output)

		Catch ex As Exception
			Return New SqlXml()

		End Try


	End Function



End Class