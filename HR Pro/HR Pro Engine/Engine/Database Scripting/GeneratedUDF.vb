﻿Namespace ScriptDB

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

  End Structure
End Namespace
