Namespace Things
  Public Class TableOrderFilter
    Inherits Things.Base

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.TableOrderFilter
      End Get
    End Property

    Public Relation As Things.Relation
    Public Order As Things.TableOrder
    Public Filter As Things.Expression
    Public ComponentNumber As Long
    Public UDF As ScriptDB.GeneratedUDF

    Public ReadOnly Property IncludedColumns As Things.Collection
      Get
        Return Me.Objects
      End Get
    End Property

    Public Overrides Property Name As String
      Get
        Dim sName As String

        sName = String.Format("udftab_{0}", Parent.Name)

        If Not Order Is Nothing Then
          sName = sName + "_" + Order.Name
        End If

        If Not Filter Is Nothing Then
          sName = sName + "_" + Filter.Name
        End If

        If Not Relation Is Nothing Then
          sName = String.Format("{0}_{1}", sName, CInt(Relation.ParentID))
        End If

        Return sName
      End Get
      Set(ByVal value As String)
        MyBase.Name = value
      End Set
    End Property

    Public Sub GenerateCode()

      Dim objOrderItem As Things.TableOrderItem
      Dim objColumn As Things.Column
      Dim aryOrderBy As New ArrayList
      Dim aryColumnList As New ArrayList
      Dim aryReturnDefintion As New ArrayList
      Dim aryParameters As New ArrayList
      Dim aryWheres As New ArrayList
      Dim aryJoins As New ArrayList

      ' Build the where clause
      If Not Filter Is Nothing Then
        Filter.AssociatedColumn = Me.Parent.Objects(Enums.Type.Column)(0)
        Filter.ExpressionType = ScriptDB.ExpressionType.ColumnFilter
        Filter.GenerateCode()
        aryWheres.Add(String.Format("({0} = 1)", Filter.UDF.SelectCode))
        aryJoins.Add(Filter.UDF.JoinCode)
      End If

      ' Build the order by clause
      If Not Order Is Nothing Then
        For Each objOrderItem In Order.Objects
          If objOrderItem.ColumnType = "O" Then
            Select Case objOrderItem.Ascending
              Case Enums.Order.Ascending
                aryOrderBy.Add(String.Format("[{0}]{1}", objOrderItem.Column.Name, " ASC"))
              Case Else
                aryOrderBy.Add(String.Format("[{0}]{1}", objOrderItem.Column.Name, " DESC"))
            End Select
          End If
        Next
      End If

      ' Add foreign key
      If Not Relation Is Nothing Then
        aryParameters.Add(String.Format("@prmID_{0} integer", CInt(Relation.ParentID)))
        aryWheres.Add(String.Format("[ID_{0}] = @prmID_{0}", CInt(Relation.ParentID)))
      End If

      ' Add the included columns
      For Each objColumn In IncludedColumns
        aryColumnList.Add(String.Format("base.[{0}]", objColumn.Name))
        aryReturnDefintion.Add(String.Format("[{0}] {1}", objColumn.Name, objColumn.DataTypeSyntax))
      Next

      With UDF
        .WhereCode = IIf(aryWheres.Count > 0, "WHERE ", "") & String.Join(" AND ", aryWheres.ToArray())
        .OrderCode = IIf(aryOrderBy.Count > 0, "ORDER BY ", "") & String.Join(", ", aryOrderBy.ToArray())
        .Code = String.Format("CREATE FUNCTION dbo.[{0}]({1})" & vbNewLine & _
                       "RETURNS @results TABLE({2})" & vbNewLine & _
                       "--WITH SCHEMABINDING" & vbNewLine & _
                       "AS" & vbNewLine & "BEGIN" & vbNewLine & _
                       "INSERT @Results SELECT {3}" & vbNewLine & _
                       "        FROM dbo.[{4}] base" & vbNewLine & _
                       "        {5}" & vbNewLine & _
                       "        {6}" & vbNewLine & _
                       "        {7}" & vbNewLine & _
                       "    RETURN;" & vbNewLine & _
                       "END" _
                      , Me.Name, String.Join(", ", aryParameters.ToArray()) _
                      , String.Join(", ", aryReturnDefintion.ToArray()), String.Join(", ", aryColumnList.ToArray()) _
                      , Me.Parent.Name, String.Join(vbNewLine, aryJoins.ToArray()), .WhereCode, .OrderCode)

      End With




    End Sub

  End Class
End Namespace