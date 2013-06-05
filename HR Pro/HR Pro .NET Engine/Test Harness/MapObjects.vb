Public Class MapObjects

  Private mobjPhoenix As Phoenix.HCM

  Private objProgress As Phoenix.HCMProgressBar
  Private Initialised As Boolean = False
  Private miMappingsRequired As Integer = 0

#Region "Progress Bar Handling"

  Private Sub UpdateProgress1(ByVal Value As Long)
    ProgressBar1.Value = Value
  End Sub
  Private Sub UpdateProgress2(ByVal Value As Long)
    ProgressBar2.Value = Value
  End Sub

#End Region

  Private Sub InitialiseStuff()

    objProgress = New Phoenix.HCMProgressBar
    AddHandler objProgress.Update1, AddressOf UpdateProgress1
    AddHandler objProgress.Update2, AddressOf UpdateProgress2

    If Not Initialised Then
      mobjPhoenix.Initialise()
      Phoenix.CommitDB.Open()

      CurrentPhase.Text = "Populating Objects..."
      Phoenix.Things.PopulateSystemThings()
      '   Phoenix.Things.PopulateThings(objProgress)
      Phoenix.Things.PopulateModuleSettings(objProgress)

      Initialised = True
    End If

  End Sub

  Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butGetObjectSelection.Click

    InitialiseStuff()
    grdObjectSelection.Attach(mobjPhoenix.ReturnThings)

  End Sub

  Public Shared Function ToXML(ByVal objTable As Phoenix.Things.Table) As String

    Dim xmlDoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()
    Dim xmlSerializer As New System.Xml.Serialization.XmlSerializer(objTable.[GetType]())
    Using xmlStream As System.IO.MemoryStream = New System.IO.MemoryStream()
      xmlSerializer.Serialize(xmlStream, objTable)
      xmlStream.Position = 0
      xmlDoc.Load(xmlStream)
      Return xmlDoc.InnerXml
    End Using

  End Function

  Public Function ToXML2(ByVal objTable As Phoenix.Things.Table) As String

    Dim objPropertyVal As Object

    Dim xmlDoc As System.Xml.XmlDocument = New System.Xml.XmlDocument
    Dim typObj As System.Type = objTable.GetType

    Dim declaration As System.Xml.XmlNode = xmlDoc.CreateNode(System.Xml.XmlNodeType.XmlDeclaration, Nothing, Nothing)
    xmlDoc.AppendChild(declaration)

    Dim xmlRoot As System.Xml.XmlElement = xmlDoc.CreateElement(typObj.Name)
    xmlDoc.AppendChild(xmlRoot)

    Dim piObjs As System.Reflection.PropertyInfo() = typObj.GetProperties()
    For Each piObj As System.Reflection.PropertyInfo In piObjs

      If piObj.GetIndexParameters.Length = 0 Then
        objPropertyVal = piObj.GetValue(objTable, Nothing)
        '      Else
        '       objPropertyVal = piObj.GetValue(

        '   piObj.


        If Not objPropertyVal Is Nothing Then

          Dim xmlSubElement As System.Xml.XmlElement = xmlDoc.CreateElement(piObj.Name)
          xmlSubElement.InnerText = objPropertyVal.ToString
          xmlRoot.AppendChild(xmlSubElement)
        End If

      End If

    Next

    ToXML2 = xmlDoc.InnerXml

  End Function


  Private Sub butGetMappings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butGetMappings.Click

    Dim strXML As String = vbNullString

    '  Dim objTable As ScriptDB.Things.Table

    ' Export selected objects as an XML file
    'strXML = ToXML(ScriptDB.HCM.Things(0))
    '    strXML = CType(ScriptDB.HCM.Things(0), ScriptDB.Things.Table).ToXML

    'For Each objTable In ScriptDB.HCM.Things
    '  strXML = strXML + objTable.ToXML
    'Next

    'strXML = ScriptDB.HCM.Things.se

    'strXML = ScriptDB.HCM.Things.ToXML

    mobjPhoenix.ReturnThings.ToXML(txtUpdateScript.Text)


    '    ScriptDB.HCM.Things(0).
    '   .SerializeTo(txtUpdateScript.Text)

    ''Dim objWriter As New System.IO.StreamWriter(txtUpdateScript.Text)
    ''objWriter.Write(strXML)
    ''objWriter.Close()




    '  strXML = ToXML2(ScriptDB.HCM.Things(0))


    'grdObjectSelection.sel


    'Dim objObject As ScriptDB.Things.iSystemObject
    'Dim objControl As HCMObjectMapping

    'ScriptDB.StructurePort.Initialise()
    'ScriptDB.StructurePort.CreateStatements(objProgress)

    'For Each objObject In ScriptDB.StructurePort.Dependancies
    '  objControl = New HCMObjectMapping
    '  objControl.FromObject = objObject
    '  objControl.Location = New System.Drawing.Point(30, (miMappingsRequired * 25))
    '  objControl.TabIndex = miMappingsRequired
    '  pnlMappings.Controls.Add(objControl)

    '  miMappingsRequired = miMappingsRequired + 1

    '  '      AddMapping(objObject.ID, objObject.Type, objObject.Name)
    'Next

    'objFile.WriteAllLines(txtUpdateScript.Text, ScriptDB.StructurePort.GetStatements)



  End Sub

End Class