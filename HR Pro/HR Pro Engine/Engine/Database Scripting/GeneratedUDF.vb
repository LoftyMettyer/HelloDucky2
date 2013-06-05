Namespace ScriptDB

  <Serializable()>
  Public Structure GeneratedUDF

    Public Name As String
    Public Code As String
    Public CallingCode As String
    Public InlineCode As String
    Public Type As Type

    Public CodeStub As String

    Public IsSchemaBound As Boolean

    Public JoinCode As String
    Public FromCode As String
    Public SelectCode As String
    Public WhereCode As String
    Public OrderCode As String
    Public Declarations As String
    Public Prerequisites As String

    Public BoilerPlate As String
    Public Comments As String

    Public Declaration As String
    Public PartNumber As Integer

    Public ReadOnly Property BaseName() As String
      Get
        ' Debug.Assert(Name.StartsWith("[dbo].["))
        Return Name.Substring(7, Name.Length - 8)
      End Get
    End Property
  
    Public Function SqlCreate() As String
      Return Code
    End Function

    Public Function SqlAlter() As String
      Return Code.Replace("CREATE FUNCTION", "ALTER FUNCTION")
    End Function

    Public Function SqlCreateStub() As String
      Return CodeStub
    End Function

    Public Function SqlAlterStub() As String
      Return CodeStub.Replace("CREATE FUNCTION", "ALTER FUNCTION")
    End Function

  End Structure
End Namespace
