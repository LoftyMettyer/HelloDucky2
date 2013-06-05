Public Class HCMObjectMapping

  Public Property MapToID As Phoenix.HCMGuid

  Public WriteOnly Property FromObject As Phoenix.Things.Base ' ScriptDB.Things.iSystemObject
    Set(ByVal value As Phoenix.Things.Base) ' ScriptDB.Things.iSystemObject)

      lblType.Text = value.Type
      cboFrom.Text = value.Name
      cboFrom.ValueMember = value.ID

      Select Case value.Type
        Case Phoenix.Things.Type.Table
          chkNewObject.Checked = False
        Case Phoenix.Things.Type.Column
          chkNewObject.Checked = True

      End Select

    End Set
  End Property


End Class
