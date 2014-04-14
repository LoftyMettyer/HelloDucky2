
Imports System
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.DirectoryServices
Imports System.Runtime.InteropServices
Imports Microsoft.SqlServer.Server

Partial Public Class DomainInfo
  <Microsoft.SqlServer.Server.SqlFunction(Name:="udfASRNetGetDomains", DataAccess:=DataAccessKind.None)> _
    Public Shared Function GetDomains() As SqlString

    Dim tmpDomains As String = String.Empty

    Try
      ' AE20080418 Fault 13106
      '// Search for objectCategory type "Domain"
      'Dim srch As DirectorySearcher = New DirectorySearcher("objectCategory=domain")
      'Dim coll As SearchResultCollection = srch.FindAll()

      'For Each rs As SearchResult In coll
      '  Dim resultPropColl As ResultPropertyCollection = rs.Properties
      '  For Each domainName As Object In resultPropColl("name")
      '    tmpDomains &= domainName.ToString() & ";"
      '  Next
      'Next

      ' AE20080205 Fault 13143
      'Dim forest As ActiveDirectory.Forest = ActiveDirectory.Forest.GetCurrentForest()

      ''// Enumerate over each returned domain.
      'For Each domain As ActiveDirectory.Domain In forest.Domains
      '  Dim ldapString As String = "LDAP://" & domain.Name
      '  Dim entry As DirectoryEntry = New DirectoryEntry(ldapString)

      '  Try
      '    Dim searcher As DirectorySearcher = New DirectorySearcher(entry)
      '    Dim result As SearchResult = searcher.FindOne()
      '    Dim domainName As String = CStr(result.Properties("name")(0)) & ";"

      '    tmpDomains &= domainName & ";"
      '  Catch ex As Exception

      '  End Try
      'Next

      Dim currentForest As ActiveDirectory.Forest = ActiveDirectory.Forest.GetCurrentForest()
      Dim gc As ActiveDirectory.GlobalCatalog = currentForest.FindGlobalCatalog()
      Dim ds As DirectorySearcher = gc.GetDirectorySearcher()
      'AE20080910 Fault #13370
      'ds.Filter = "objectCategory=domain"
      'AE20080917 Fault #13372
      'ds.Filter = "(&(objectCategory=trustedDomain)(name=*))"
      'ds.Filter = "(&(|(objectCategory=trustedDomain)(objectCategory=domain))(name=*))"
      ds.Filter = "(&(objectCategory=trustedDomain)(name=*))"
      ds.PropertiesToLoad.Add("name")
      ds.Sort.Direction = SortDirection.Ascending
      ds.Sort.PropertyName = "name"
      ds.SearchScope = SearchScope.Subtree

      Dim domainList As New ArrayList
      Try
        Dim myDomain As String = GetFQDNFromDomainName(Environment.UserDomainName)
        domainList.Add(myDomain)
        tmpDomains = myDomain & ";"

        Dim results As SearchResultCollection = ds.FindAll

        For Each result As SearchResult In results
          Dim domainName As String = result.Properties("name")(0).ToString()

          If (Not domainName.Equals(String.Empty)) AndAlso (Not domainList.Contains(domainName)) Then
            domainList.Add(domainName)
            tmpDomains &= domainName & ";"
          End If

        Next
        results = Nothing

      Catch ex As Exception
        Throw ex
      Finally
        domainList.Clear()
        domainList = Nothing
        ds.Dispose()
        ds = Nothing
        gc.Dispose()
        gc = Nothing
        currentForest.Dispose()
        currentForest = Nothing
      End Try

    Catch ex As Exception
      Throw ex
    End Try

    Return New SqlString(tmpDomains)

  End Function

  Private Shared Function GetDomainNameFromFQDN(ByVal fqdn As String) As String

    Dim tmpDomains As String = String.Empty

    Dim de As New DirectoryEntry("LDAP://" & fqdn)
    Dim ds As New DirectorySearcher(de, "objectCategory=domain")

    Try
      Dim result As SearchResult = ds.FindOne
      tmpDomains &= result.Properties("name")(0).ToString()
      result = Nothing
    Catch
      Return tmpDomains
    End Try

    Return tmpDomains
  End Function

  <Microsoft.SqlServer.Server.SqlFunction(Name:="udfASRNetGetDomainNameFromFQDN", DataAccess:=DataAccessKind.None)> _
  Public Shared Function GetNetBiosDomainNameFromFQDN(ByVal fqdn As String) As String

    Dim domainName As String = String.Empty
    Try
      ' get the root namespace
      Dim rootDSE As DirectoryEntry = New System.DirectoryServices.DirectoryEntry("LDAP://" & fqdn & "/RootDSE")
      ' get the name of the domain we're currently in
      Dim domainDN As String = CType(rootDSE.Properties("DefaultNamingContext")(0), String)
      Dim parts As DirectoryEntry = New DirectoryEntry("LDAP://CN=Partitions,CN=Configuration," & domainDN)
      For Each part As DirectoryEntry In parts.Children
        ' search the AD Configuration container for our domain name
        If (CType(part.Properties("nCName")(0), String) = domainDN) Then
          ' Properties are case sensitive!
          domainName = CType(part.Properties("nETBIOSName")(0), String).ToLower
        End If
      Next

    Catch
      Return domainName
    End Try

    Return domainName
  End Function

  Private Shared Function GetFQDNFromDomainName(ByVal dn As String) As String

    Dim ldapPath As String = String.Empty

    Try
      Dim objContext As ActiveDirectory.DirectoryContext = _
        New ActiveDirectory.DirectoryContext(ActiveDirectory.DirectoryContextType.Domain, dn)

      Dim objDomain As ActiveDirectory.Domain = ActiveDirectory.Domain.GetDomain(objContext)
      ldapPath = objDomain.Name
    Catch e As DirectoryServicesCOMException
      Return String.Empty
    End Try

    Return ldapPath

  End Function

   <Microsoft.SqlServer.Server.SqlFunction(Name:="udfASRNetGetDomainLogins", DataAccess:=DataAccessKind.None)> _
   Public Shared Function GetDomainLogins(ByVal domainPath As String) As <SqlFacet(MaxSize:=-1)> SqlString

      'AE20090126 Fault #13577
      'Dim domainName As String = GetDomainNameFromFQDN(domainPath)
      Dim domainName As String = GetNetBiosDomainNameFromFQDN(domainPath)
      If (domainName.Equals(String.Empty)) Then
         domainName = GetDomainNameFromFQDN(domainPath)

         If (domainName.Equals(String.Empty)) Then
            Return New SqlString(String.Empty)
         End If
      End If

      Dim tmpUsers As String = String.Empty
      Dim entry As DirectoryEntry = New DirectoryEntry("LDAP://" & domainPath)

      '// Search for objectCategory type "person"
      'AE20071203 Fault #12669
      'Dim srch As DirectorySearcher = New DirectorySearcher("(&(objectCategory=user)(showInAddressBook=*))")
      Dim srch As DirectorySearcher = New DirectorySearcher(entry, "(&(objectCategory=person)(objectClass=user))")

      Try
         srch.PageSize = 1000
         srch.PropertiesToLoad.Add("sAMAccountName")
         srch.Sort.Direction = SortDirection.Ascending
         srch.Sort.PropertyName = "sAMAccountName"
         srch.SearchScope = SearchScope.Subtree

         Dim coll As SearchResultCollection = srch.FindAll()

         'AE20071203 Fault #12816
         'Try
         '// Enumerate over each returned person.
         For Each rs As SearchResult In coll
            Dim resultPropColl As ResultPropertyCollection = rs.Properties
            For Each userName As Object In resultPropColl("sAMAccountName")
               tmpUsers &= String.Concat(domainName, "\", userName.ToString(), ";")
            Next
         Next

      Catch ex As DirectoryServicesCOMException
         Throw ex
      Catch ex As Exception
         Throw ex
      Finally
         srch.Dispose()
         srch = Nothing
         entry.Dispose()
         entry = Nothing
      End Try

      Return New SqlString(tmpUsers)
   End Function

   <Microsoft.SqlServer.Server.SqlProcedure(Name:="spASRGetWindowsUsersFromAssembly")> _
   Public Shared Sub GetWindowsUsers(ByVal domainPath As String)

      '// Get the specified domain entry
      Dim entry As DirectoryEntry = New DirectoryEntry("LDAP://" & domainPath)

      '// Search for objectCategory type "person"
      Dim srch As DirectorySearcher = New DirectorySearcher(entry, "(&(objectCategory=person)(objectClass=user))")
      srch.PropertiesToLoad.Add("sAMAccountName")
      srch.PageSize = 1000
      srch.Sort.PropertyName = "sAMAccountName"
      srch.Sort.Direction = SortDirection.Ascending

      Try
         ' Create the record and specify the metadata for the columns.
         Dim record As New SqlDataRecord(New SqlMetaData("User", SqlDbType.NVarChar, 100))

         ' Mark the begining of the result-set.
         SqlContext.Pipe.SendResultsStart(record)

         Dim resultcol As SearchResultCollection = srch.FindAll()

         ' Enumerate over each returned user.
         For Each rs As SearchResult In resultcol

            ' Set row value for the user.
            Dim name As String = rs.Properties("sAMAccountName")(0).ToString()
            record.SetString(0, name)

            ' Send the row back to the client.
            SqlContext.Pipe.SendResultsRow(record)
         Next

      Catch ex As DirectoryServicesCOMException
         Throw ex
      Catch ex As Exception
         Throw ex
      Finally
         ' Mark the end of the result-set.
         SqlContext.Pipe.SendResultsEnd()
         srch.Dispose()
         entry.Dispose()
      End Try

   End Sub

  <Microsoft.SqlServer.Server.SqlProcedure(Name:="spASRGetWindowsGroupsFromAssembly")> _
  Public Shared Sub GetWindowsGroups(ByVal domainPath As String)

    'AE20090126 Fault #13577
    'Dim domainName As String = GetDomainNameFromFQDN(domainPath)
    Dim domainName As String = GetNetBiosDomainNameFromFQDN(domainPath)
    If (domainName.Equals(String.Empty)) Then
      domainName = GetDomainNameFromFQDN(domainPath)

      If (domainName.Equals(String.Empty)) Then
        Throw New Exception("Domain " & domainPath & " not operational")
        Return
      End If
    End If

    '// Get the specified domain entry
    Dim entry As DirectoryEntry = New DirectoryEntry("LDAP://" & domainPath)

    '// Search for objectCategory type "group"
    Dim srch As DirectorySearcher = New DirectorySearcher(entry, "(objectCategory=group)")
    srch.PageSize = 1000
    srch.PropertiesToLoad.Add("sAMAccountName")
    srch.PropertiesToLoad.Add("description")
    srch.Sort.Direction = SortDirection.Ascending
    srch.Sort.PropertyName = "sAMAccountName"

    Try
      ' Create the record and specify the metadata for the columns.
      Dim record As New SqlDataRecord( _
        New SqlMetaData("Group", SqlDbType.NVarChar, 255), _
        New SqlMetaData("Comment", SqlDbType.NVarChar, 255))

      ' Mark the begining of the result-set.
      SqlContext.Pipe.SendResultsStart(record)

      Dim resultcol As SearchResultCollection = srch.FindAll()

         ' Enumerate over each returned group.
         For Each rs As SearchResult In resultcol

            ' Set row value for the group.
            Dim name As String = rs.Properties("sAMAccountName")(0).ToString()
            record.SetString(0, domainName & "\" & name)

            Dim description As String = String.Empty
            If rs.Properties.Contains("description") Then
               description = rs.Properties("description")(0).ToString()
            End If
            record.SetString(1, description)

            ' Send the row back to the client.
            SqlContext.Pipe.SendResultsRow(record)
         Next
    Catch ex As DirectoryServicesCOMException
      Throw ex
    Catch ex As Exception
      Throw ex
    Finally
      ' Mark the end of the result-set.
      SqlContext.Pipe.SendResultsEnd()
      srch.Dispose()
      srch = Nothing
      entry.Dispose()
      entry = Nothing
    End Try

  End Sub

  <Microsoft.SqlServer.Server.SqlProcedure(Name:="spASRGetDomainPolicyFromAssembly")> _
    Public Shared Sub GetDomainPolicy(<Out()> ByRef lockoutDuration As Long, <Out()> ByRef lockoutThreshold As Long, _
                        <Out()> ByRef lockoutObservationWindow As Long, <Out()> ByRef maxPwdAge As Long, <Out()> ByRef minPwdAge As Long, _
                        <Out()> ByRef minPwdLength As Long, <Out()> ByRef pwdHistoryLength As Long, <Out()> ByRef pwdProperties As Long)

    Try
      Dim rootDSE As New DirectoryEntry("LDAP://RootDSE")
      Dim group As New DirectoryEntry("LDAP://" & rootDSE.Properties("defaultNamingContext")(0).ToString())

      'AE20071203 Fault #12816
      'Try
      Dim dp As New DomainPolicy(group)

      lockoutDuration = CType(dp.LockoutDuration.TotalSeconds, Long)
      lockoutThreshold = dp.LockoutThreshold
      lockoutObservationWindow = CType(dp.LockoutObservationWindow.TotalSeconds, Long)
      maxPwdAge = CType(dp.MaxPasswordAge.TotalDays, Long)
      minPwdAge = CType(dp.MinPasswordAge.TotalDays, Long)
      minPwdLength = dp.MinPasswordLength
      pwdHistoryLength = dp.PasswordHistoryLength
      pwdProperties = dp.PasswordProperties

      dp.Dispose()
      dp = Nothing
      group.Dispose()
      group = Nothing
      rootDSE.Dispose()
      rootDSE = Nothing

    Catch ex As Exception

    End Try

    Return

  End Sub

  <Microsoft.SqlServer.Server.SqlFunction(Name:="udfASRNetUserIsMemberOfGroup", DataAccess:=DataAccessKind.None)> _
  Public Shared Function UserIsMemberOfGroup(ByVal domain As String, ByVal username As String, ByVal group As String) As SqlBoolean
    Dim isMember As Boolean = False

    Dim result As SearchResult
    Dim userDN As String = String.Empty
    Dim groupDN As String = String.Empty

    Dim domainEntryString As String = "LDAP://" & domain
    Dim domainEntry As DirectoryEntry = New DirectoryEntry(domainEntryString)
    Dim srch As New DirectorySearcher(domainEntry)

    Try
      Try
        srch.Filter = "sAMAccountName=" & username
        result = srch.FindOne()

        userDN = result.Path

      Catch ex As Exception
        Throw New ApplicationException("Couldn't get distinguishedName for user.", ex)
      End Try

      Dim user As DirectoryEntry = New DirectoryEntry(userDN)

      If IsNothing(user) Then
        Throw New ArgumentNullException("user")
      End If


      ' Get the SID for the Group
      Dim objectSid As Byte() = GetGroupSID(domainEntry, group)

      Try
        If Not IsNothing(objectSid) Then
          ' now retrieve the tokenGroups property on the user
          Dim props() As String = {"tokenGroups"}
          user.RefreshCache(props)

          ' cycle through the tokenGroups and see if one matches the SID of the group
          '
          For Each entry As Byte() In user.Properties("tokenGroups")
            If CompareByteArrays(entry, objectSid) Then
              isMember = True
              Exit For
            End If
          Next

        End If

      Catch ex As Exception
        Throw New ApplicationException("An error occurred while checking group membership.", ex)
      Finally
        objectSid = Nothing
        user.Dispose()
        user = Nothing
      End Try

    Catch ex As Exception
      Return New SqlBoolean(isMember)
    Finally
      srch.Dispose()
      srch = Nothing
      domainEntry.Dispose()
      domainEntry = Nothing
      result = Nothing
    End Try

    Return New SqlBoolean(isMember)
  End Function

  <Microsoft.SqlServer.Server.SqlProcedure(Name:="spASRGroupsUserIsMemberOf")> _
  Public Shared Sub GroupsUserIsMemberOf(ByVal user As String)

    '// First get the username and domain 
    Dim domain As String = String.Empty
    Dim username As String = String.Empty

    Try
      If user.IndexOf("\") > 0 Then
        domain = user.Substring(0, user.IndexOf("\"))
        username = user.Substring(user.IndexOf("\") + 1)
      Else
        Dim domainNameEntry As DirectoryEntry = New DirectoryEntry("LDAP://")
        Dim domainNameSrch As DirectorySearcher = New DirectorySearcher("objectCategory=Domain")
        Dim domainNameResult As SearchResult = domainNameSrch.FindOne()
        domain = CType(domainNameResult.Properties("name")(0), String)
        username = user
      End If
    Catch ex As Exception
      Throw New ApplicationException("Couldn't get domain for user.", ex)
    End Try

    If domain = String.Empty OrElse username = String.Empty Then
      Throw New ArgumentNullException("user")
    End If

    '// Now lets get the userDN and groupDN
    Dim result As SearchResult
    Dim userDN As String = String.Empty
    Dim groupDN As String = String.Empty

    Dim domainEntryString As String = "LDAP://" & domain
    Dim domainEntry As DirectoryEntry = New DirectoryEntry(domainEntryString)
    Dim srch As New DirectorySearcher(domainEntry)

    Try
      Try
        srch.Filter = "sAMAccountName=" & username
        result = srch.FindOne()
        userDN = result.Path
      Catch ex As Exception
        Throw New ApplicationException("Couldn't get distinguishedName for user.", ex)
      End Try

      Dim userDE As DirectoryEntry = New DirectoryEntry(userDN)

      If IsNothing(userDE) Then
        Throw New ArgumentNullException("user")
      End If

      Try
        ' Create the record and specify the metadata for the columns.
        Dim record As New SqlDataRecord( _
          New SqlMetaData("User", SqlDbType.NVarChar, 256), _
          New SqlMetaData("Group", SqlDbType.NVarChar, 256))

        ' Mark the begining of the result-set.
        SqlContext.Pipe.SendResultsStart(record)

        ' now retrieve the tokenGroups property on the user
        Dim props() As String = {"tokenGroups"}
        userDE.RefreshCache(props)

        ' cycle through the tokenGroups and get the group name
        For Each objectSid As Byte() In userDE.Properties("tokenGroups")
          Dim searchSid As String = ConvertByteToStringSid(objectSid)

          If searchSid <> String.Empty Then
            ' Set row value for the user.
            record.SetString(0, domain & "\" & username)
            record.SetString(1, domain & "\" & GetGroupFromSID(domainEntry, searchSid))

            ' Send the row back to the client.
            SqlContext.Pipe.SendResultsRow(record)

          End If
        Next

      Catch ex As Exception
        Throw New ApplicationException("An error occurred while checking group membership.", ex)
      Finally
        SqlContext.Pipe.SendResultsEnd()
        userDE.Dispose()
        userDE = Nothing
      End Try

    Catch ex As Exception
      Return
    Finally
      srch.Dispose()
      srch = Nothing
      domainEntry.Dispose()
      domainEntry = Nothing
      result = Nothing
    End Try

    Return
  End Sub

  Private Shared Function GetGroupFromSID(ByVal domainEntry As DirectoryEntry, ByVal objectSID As String) As String
    Dim filterLDAP As String = String.Format("(&(objectCategory=group)(objectSid={0}))", objectSID)

    Try
      Dim srch As New DirectorySearcher(domainEntry, filterLDAP)

      Dim result As SearchResult = srch.FindOne()

      If Not result Is Nothing Then
        Dim group As String = CType(result.Properties("name")(0), String)

        result = Nothing
        srch.Dispose()

        Return group
      Else
        srch.Dispose()

        Return String.Empty
      End If
    Catch ex As Exception
      Throw New Exception("Error getting the group from the object SID.", ex)
    End Try

  End Function

  Private Shared Function GetGroupSID(ByVal domainEntry As DirectoryEntry, ByVal groupName As String) As Byte()
    Dim filterLDAP As String = String.Format("(sAMAccountName={0})", groupName)

    Try
      Dim srch As New DirectorySearcher(domainEntry, filterLDAP)
      srch.PropertiesToLoad.Add("objectSid")

      Dim result As SearchResult = srch.FindOne()

      If Not result Is Nothing Then
        Dim objectSID As Byte() = CType(result.Properties("objectSid")(0), Byte())

        result = Nothing
        srch.Dispose()

        Return objectSID
      Else
        srch.Dispose()

        Return Nothing
      End If
    Catch ex As Exception
      Throw New Exception("Error getting the object SID.", ex)
    End Try
  End Function

  Private Shared Function CompareByteArrays(ByVal data1 As Byte(), ByVal data2 As Byte()) As Boolean

    Try
      ' If both are null, they're equal
      If IsNothing(data1) AndAlso IsNothing(data2) Then
        Return True
      End If

      ' If either but not both are null, they're not equal
      If IsNothing(data1) OrElse IsNothing(data2) Then
        Return False
      End If

      If data1.Length <> data2.Length Then
        Return False
      End If

      For index As Integer = 0 To data1.Length - 1
        If data1(index) <> data2(index) Then
          Return False
        End If
      Next

      Return True
    Catch ex As Exception
      Throw New ApplicationException("An error occurred while comparing byte arrays.", ex)
    End Try

  End Function

  Private Shared Function ConvertByteToStringSid(ByVal sidBytes As Byte()) As String
    Dim strSid As StringBuilder = New StringBuilder()
    strSid.Append("S-")

    Try
      '// Add SID revision.
      strSid.Append(sidBytes(0).ToString())

      '// Next six bytes are SID authority value.
      If sidBytes(6) <> 0 OrElse sidBytes(5) <> 0 Then
        Dim strAuth As String = String.Format _
                ("0x{0:2x}{1:2x}{2:2x}{3:2x}{4:2x}{5:2x}", _
                CType(sidBytes(1), Int16), _
                CType(sidBytes(2), Int16), _
                CType(sidBytes(3), Int16), _
                CType(sidBytes(4), Int16), _
                CType(sidBytes(5), Int16), _
                CType(sidBytes(6), Int16))
        strSid.Append("-")
        strSid.Append(strAuth)

      Else

        Dim iVal As Int64 = CType(sidBytes(1), Int32) _
          + CType((sidBytes(2) << 8), Int32) _
          + CType((sidBytes(3) << 16), Int32) _
          + CType((sidBytes(4) << 24), Int32)
        strSid.Append("-")
        strSid.Append(iVal.ToString())
      End If

      '// Get sub authority count...
      Dim iSubCount As Int32 = Convert.ToInt32(sidBytes(7))
      Dim idxAuth As Int32 = 0

      For i As Int32 = 0 To iSubCount - 1
        idxAuth = 8 + i * 4
        Dim iSubAuth As UInt32 = BitConverter.ToUInt32(sidBytes, idxAuth)
        strSid.Append("-")
        strSid.Append(iSubAuth.ToString())
      Next

    Catch ex As Exception
      Return ""
    End Try

    Return strSid.ToString()
  End Function

End Class
