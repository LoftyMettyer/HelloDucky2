Public Class HCMObjectMapping

  Public Property MapToID As HRProEngine.HCMGuid

  Public WriteOnly Property FromObject As HRProEngine.Things.Base ' ScriptDB.Things.iSystemObject
    Set(ByVal value As HRProEngine.Things.Base) ' ScriptDB.Things.iSystemObject)

      lblType.Text = value.Type
      cboFrom.Text = value.Name
      cboFrom.ValueMember = value.ID

      Select Case value.Type
        Case HRProEngine.Things.Type.Table
          chkNewObject.Checked = False
        Case HRProEngine.Things.Type.Column
          chkNewObject.Checked = True

      End Select

    End Set
  End Property


End Class
