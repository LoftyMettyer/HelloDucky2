Option Strict Off
Option Explicit On
Module modFileSecurity
	
	' Constants used within our API calls. Refer to the MSDN for more
	' information on how/what these constants are used for.
	
	' Memory constants used through various memory API calls.
	Public Const GMEM_MOVEABLE As Integer = &H2
	Public Const LMEM_FIXED As Integer = &H0
	Public Const LMEM_ZEROINIT As Integer = &H40
	Public Const LPTR As Decimal = (LMEM_FIXED + LMEM_ZEROINIT)
	Public Const GENERIC_READ As Integer = &H80000000
	Public Const GENERIC_ALL As Integer = &H10000000
	Public Const GENERIC_EXECUTE As Integer = &H20000000
	Public Const GENERIC_WRITE As Integer = &H40000000
	
	' The file/security API call constants.
	' Refer to the MSDN for more information on how/what these constants
	' are used for.
	Public Const DACL_SECURITY_INFORMATION As Integer = &H4
	Public Const SECURITY_DESCRIPTOR_REVISION As Short = 1
	Public Const SECURITY_DESCRIPTOR_MIN_LENGTH As Short = 20
	Public Const SD_SIZE As Decimal = (65536 + SECURITY_DESCRIPTOR_MIN_LENGTH)
	Public Const ACL_REVISION2 As Short = 2
	Public Const ACL_REVISION As Short = 2
	Public Const MAXDWORD As Integer = &HFFFFFFFF
	Public Const SidTypeUser As Short = 1
	Public Const AclSizeInformation As Short = 2
	
	'  The following are the inherit flags that go into the AceFlags field
	'  of an Ace header.
	
	Public Const OBJECT_INHERIT_ACE As Integer = &H1
	Public Const CONTAINER_INHERIT_ACE As Integer = &H2
	Public Const NO_PROPAGATE_INHERIT_ACE As Integer = &H4
	Public Const INHERIT_ONLY_ACE As Integer = &H8
	Public Const INHERITED_ACE As Integer = &H10
	Public Const VALID_INHERIT_FLAGS As Integer = &H1F
	Public Const DELETE As Integer = &H10000
	
	' Structures used by our API calls.
	' Refer to the MSDN for more information on how/what these
	' structures are used for.
	Structure ACE_HEADER
		Dim AceType As Byte
		Dim AceFlags As Byte
		Dim AceSize As Short
	End Structure
	
	
	Public Structure ACCESS_DENIED_ACE
		Dim Header As ACE_HEADER
		Dim Mask As Integer
		Dim SidStart As Integer
	End Structure
	
	Structure ACCESS_ALLOWED_ACE
		Dim Header As ACE_HEADER
		Dim Mask As Integer
		Dim SidStart As Integer
	End Structure
	
	Structure ACL
		Dim AclRevision As Byte
		Dim Sbz1 As Byte
		Dim AclSize As Short
		Dim AceCount As Short
		Dim Sbz2 As Short
	End Structure
	
	Structure ACL_SIZE_INFORMATION
		Dim AceCount As Integer
		Dim AclBytesInUse As Integer
		Dim AclBytesFree As Integer
	End Structure
	
	Structure SECURITY_DESCRIPTOR
		Dim Revision As Byte
		Dim Sbz1 As Byte
		Dim Control As Integer
		Dim Owner As Integer
		Dim Group As Integer
		Dim sACL As ACL
		Dim Dacl As ACL
	End Structure
	
	' API calls used within this sample. Refer to the MSDN for more
	' information on how/what these APIs do.
	
	Declare Function GetComputerName Lib "kernel32"  Alias "GetComputerNameA"(ByVal lpBuffer As String, ByRef nSize As Integer) As Integer
	
	Declare Function GetUserName Lib "advapi32.dll"  Alias "GetUserNameA"(ByVal lpBuffer As String, ByRef nSize As Integer) As Integer
	
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Declare Function LookupAccountName Lib "advapi32.dll"  Alias "LookupAccountNameA"(ByRef lpSystemName As String, ByVal lpAccountName As String, ByRef sid As Any, ByRef cbSid As Integer, ByVal ReferencedDomainName As String, ByRef cbReferencedDomainName As Integer, ByRef peUse As Integer) As Integer
	
	'UPGRADE_WARNING: Structure SECURITY_DESCRIPTOR may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Declare Function InitializeSecurityDescriptor Lib "advapi32.dll" (ByRef pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal dwRevision As Integer) As Integer
	
	Declare Function GetSecurityDescriptorDacl Lib "advapi32.dll" (ByRef pSecurityDescriptor As Byte, ByRef lpbDaclPresent As Integer, ByRef pDacl As Integer, ByRef lpbDaclDefaulted As Integer) As Integer
	
	Declare Function GetFileSecurityN Lib "advapi32.dll"  Alias "GetFileSecurityA"(ByVal lpFileName As String, ByVal RequestedInformation As Integer, ByVal pSecurityDescriptor As Integer, ByVal nLength As Integer, ByRef lpnLengthNeeded As Integer) As Integer
	
	Declare Function GetFileSecurity Lib "advapi32.dll"  Alias "GetFileSecurityA"(ByVal lpFileName As String, ByVal RequestedInformation As Integer, ByRef pSecurityDescriptor As Byte, ByVal nLength As Integer, ByRef lpnLengthNeeded As Integer) As Integer
	
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Declare Function GetAclInformation Lib "advapi32.dll" (ByVal pAcl As Integer, ByRef pAclInformation As Any, ByVal nAclInformationLength As Integer, ByVal dwAclInformationClass As Integer) As Integer
	
	Public Declare Function EqualSid Lib "advapi32.dll" (ByRef pSid1 As Byte, ByVal pSid2 As Integer) As Integer
	
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Declare Function GetLengthSid Lib "advapi32.dll" (ByRef pSid As Any) As Integer
	
	Declare Function InitializeAcl Lib "advapi32.dll" (ByRef pAcl As Byte, ByVal nAclLength As Integer, ByVal dwAclRevision As Integer) As Integer
	
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Declare Function GetAce Lib "advapi32.dll" (ByVal pAcl As Integer, ByVal dwAceIndex As Integer, ByRef pace As Any) As Integer
	
	Declare Function AddAce Lib "advapi32.dll" (ByVal pAcl As Integer, ByVal dwAceRevision As Integer, ByVal dwStartingAceIndex As Integer, ByVal pAceList As Integer, ByVal nAceListLength As Integer) As Integer
	
	Declare Function AddAccessAllowedAce Lib "advapi32.dll" (ByRef pAcl As Byte, ByVal dwAceRevision As Integer, ByVal AccessMask As Integer, ByRef pSid As Byte) As Integer
	
	Public Declare Function AddAccessDeniedAce Lib "advapi32.dll" (ByRef pAcl As Byte, ByVal dwAceRevision As Integer, ByVal AccessMask As Integer, ByRef pSid As Byte) As Integer
	
	'UPGRADE_WARNING: Structure SECURITY_DESCRIPTOR may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Declare Function SetSecurityDescriptorDacl Lib "advapi32.dll" (ByRef pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal bDaclPresent As Integer, ByRef pDacl As Byte, ByVal bDaclDefaulted As Integer) As Integer
	
	'UPGRADE_WARNING: Structure SECURITY_DESCRIPTOR may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Declare Function SetFileSecurity Lib "advapi32.dll"  Alias "SetFileSecurityA"(ByVal lpFileName As String, ByVal SecurityInformation As Integer, ByRef pSecurityDescriptor As SECURITY_DESCRIPTOR) As Integer
	
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Declare Sub CopyMemory Lib "kernel32"  Alias "RtlMoveMemory"(ByRef hpvDest As Any, ByVal hpvSource As Integer, ByVal cbCopy As Integer)
	
	Public Sub SetAccess(ByRef sUsername As String, ByRef sFileName As String, ByRef lMask As Integer)
		Dim lResult As Integer ' Result of various API calls.
		Dim I As Short ' Used in looping.
		Dim bUserSid(255) As Byte ' This will contain your SID.
		Dim bTempSid(255) As Byte ' This will contain the Sid of each ACE in the ACL .
		Dim sSystemName As String ' Name of this computer system.
		
		Dim lSystemNameLength As Integer ' Length of string that contains
		' the name of this system.
		
		Dim lLengthUserName As Integer ' Max length of user name.
		
		'Dim sUserName As String * 255  ' String to hold the current user
		' name.
		
		
		Dim lUserSID As Integer ' Used to hold the SID of the
		' current user.
		
		Dim lTempSid As Integer ' Used to hold the SID of each ACE in the ACL
		Dim lUserSIDSize As Integer ' Size of the SID.
		Dim sDomainName As New VB6.FixedLengthString(255) ' Domain the user belongs to.
		Dim lDomainNameLength As Integer ' Length of domain name needed.
		
		Dim lSIDType As Integer ' The type of SID info we are
		' getting back.
		
		Dim sFileSD As SECURITY_DESCRIPTOR ' SD of the file we want.
		
		Dim bSDBuf() As Byte ' Buffer that holds the security
		' descriptor for this file.
		
		Dim lFileSDSize As Integer ' Size of the File SD.
		Dim lSizeNeeded As Integer ' Size needed for SD for file.
		
		
		Dim sNewSD As SECURITY_DESCRIPTOR ' New security descriptor.
		
		Dim sACL As ACL ' Used in grabbing the DACL from
		' the File SD.
		
		Dim lDaclPresent As Integer ' Used in grabbing the DACL from
		' the File SD.
		
		Dim lDaclDefaulted As Integer ' Used in grabbing the DACL from
		' the File SD.
		
		Dim sACLInfo As ACL_SIZE_INFORMATION ' Used in grabbing the ACL
		' from the File SD.
		
		Dim lACLSize As Integer ' Size of the ACL structure used
		' to get the ACL from the File SD.
		
		Dim pAcl As Integer ' Current ACL for this file.
		Dim lNewACLSize As Integer ' Size of new ACL to create.
		Dim bNewACL() As Byte ' Buffer to hold new ACL.
		
		Dim sCurrentACE As ACCESS_ALLOWED_ACE ' Current ACE.
		Dim pCurrentAce As Integer ' Our current ACE.
		
		Dim nRecordNumber As Integer
		
		' Get the SID of the user. (Refer to the MSDN for more information on SIDs
		' and their function/purpose in the operating system.) Get the SID of this
		' user by using the LookupAccountName API. In order to use the SID
		' of the current user account, call the LookupAccountName API
		' twice. The first time is to get the required sizes of the SID
		' and the DomainName string. The second call is to actually get
		' the desired information.
		
		lResult = LookupAccountName(vbNullString, sUsername, bUserSid(0), 255, sDomainName.Value, lDomainNameLength, lSIDType)
		
		' Now set the sDomainName string buffer to its proper size before
		' calling the API again.
		sDomainName.Value = Space(lDomainNameLength)
		
		' Call the LookupAccountName again to get the actual SID for user.
		lResult = LookupAccountName(vbNullString, sUsername, bUserSid(0), 255, sDomainName.Value, lDomainNameLength, lSIDType)
		
		' Return value of zero means the call to LookupAccountName failed;
		' test for this before you continue.
		If (lResult = 0) Then
			'msgbox "Error: Unable to Lookup the Current User Account: " & sUsername
			Exit Sub
		End If
		
		' You now have the SID for the user who is logged on.
		' The SID is of interest since it will get the security descriptor
		' for the file that the user is interested in.
		' The GetFileSecurity API will retrieve the Security Descriptor
		' for the file. However, you must call this API twice: once to get
		' the proper size for the Security Descriptor and once to get the
		' actual Security Descriptor information.
		
		lResult = GetFileSecurityN(sFileName, DACL_SECURITY_INFORMATION, 0, 0, lSizeNeeded)
		
		' Redimension the Security Descriptor buffer to the proper size.
		ReDim bSDBuf(lSizeNeeded)
		
		' Now get the actual Security Descriptor for the file.
		lResult = GetFileSecurity(sFileName, DACL_SECURITY_INFORMATION, bSDBuf(0), lSizeNeeded, lSizeNeeded)
		
		' A return code of zero means the call failed; test for this
		' before continuing.
		If (lResult = 0) Then
			'msgbox "Error: Unable to Get the File Security Descriptor"
			Exit Sub
		End If
		
		' Call InitializeSecurityDescriptor to build a new SD for the
		' file.
		lResult = InitializeSecurityDescriptor(sNewSD, SECURITY_DESCRIPTOR_REVISION)
		
		' A return code of zero means the call failed; test for this
		' before continuing.
		If (lResult = 0) Then
			'msgbox "Error: Unable to Initialize New Security Descriptor"
			Exit Sub
		End If
		
		' You now have the file's SD and a new Security Descriptor
		' that will replace the current one. Next, pull the DACL from
		' the SD. To do so, call the GetSecurityDescriptorDacl API
		' function.
		
		lResult = GetSecurityDescriptorDacl(bSDBuf(0), lDaclPresent, pAcl, lDaclDefaulted)
		
		' A return code of zero means the call failed; test for this
		' before continuing.
		If (lResult = 0) Then
			'msgbox "Error: Unable to Get DACL from File Security " & "Descriptor"
			Exit Sub
		End If
		
		' You have the file's SD, and want to now pull the ACL from the
		' SD. To do so, call the GetACLInformation API function.
		' See if ACL exists for this file before getting the ACL
		' information.
		If (lDaclPresent = False) Then
			'msgbox "Error: No ACL Information Available for this File"
			Exit Sub
		End If
		
		' Attempt to get the ACL from the file's Security Descriptor.
		'UPGRADE_WARNING: Couldn't resolve default property of object sACLInfo. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lResult = GetAclInformation(pAcl, sACLInfo, Len(sACLInfo), 2)
		
		' A return code of zero means the call failed; test for this
		' before continuing.
		If (lResult = 0) Then
			'msgbox "Error: Unable to Get ACL from File Security Descriptor"
			Exit Sub
		End If
		
		' Now that you have the ACL information, compute the new ACL size
		' requirements.
		lNewACLSize = sACLInfo.AclBytesInUse + (Len(sCurrentACE) + GetLengthSid(bUserSid(0))) * 2 - 4
		
		' Resize our new ACL buffer to its proper size.
		ReDim bNewACL(lNewACLSize)
		
		' Use the InitializeAcl API function call to initialize the new
		' ACL.
		lResult = InitializeAcl(bNewACL(0), lNewACLSize, ACL_REVISION)
		
		' A return code of zero means the call failed; test for this
		' before continuing.
		If (lResult = 0) Then
			'msgbox "Error: Unable to Initialize New ACL"
			Exit Sub
		End If
		
		' If a DACL is present, copy it to a new DACL.
		If (lDaclPresent) Then
			
			' Copy the ACEs from the file to the new ACL.
			If (sACLInfo.AceCount > 0) Then
				
				' Grab each ACE and stuff them into the new ACL.
				nRecordNumber = 0
				For I = 0 To (sACLInfo.AceCount - 1)
					
					' Attempt to grab the next ACE.
					lResult = GetAce(pAcl, I, pCurrentAce)
					
					' Make sure you have the current ACE under question.
					If (lResult = 0) Then
						'msgbox "Error: Unable to Obtain ACE (" & I & ")"
						Exit Sub
					End If
					
					' You have a pointer to the ACE. Place it
					' into a structure, so you can get at its size.
					'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object sCurrentACE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CopyMemory(sCurrentACE, pCurrentAce, LenB(sCurrentACE))
					
					'Skip adding the ACE to the ACL if this is same usersid
					lTempSid = pCurrentAce + 8
					If EqualSid(bUserSid(0), lTempSid) = 0 Then
						
						' Now that you have the ACE, add it to the new ACL.
						'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
						lResult = AddAce(VarPtr(bNewACL(0)), ACL_REVISION, MAXDWORD, pCurrentAce, sCurrentACE.Header.AceSize)
						
						' Make sure you have the current ACE under question.
						If (lResult = 0) Then
							'msgbox "Error: Unable to Add ACE to New ACL"
							Exit Sub
						End If
						nRecordNumber = nRecordNumber + 1
					End If
					
				Next I
				
				' You have now rebuilt a new ACL and want to add it to
				' the newly created DACL.
				lResult = AddAccessAllowedAce(bNewACL(0), ACL_REVISION, lMask, bUserSid(0))
				
				' Make sure added the ACL to the DACL.
				If (lResult = 0) Then
					'msgbox "Error: Unable to Add ACL to DACL"
					Exit Sub
				End If
				
				'If it's directory, we need to add inheritance staff.
				If GetAttr(sFileName) And FileAttribute.Directory Then
					
					' Attempt to grab the next ACE which is what we just added.
					'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
					lResult = GetAce(VarPtr(bNewACL(0)), nRecordNumber, pCurrentAce)
					
					' Make sure you have the current ACE under question.
					If (lResult = 0) Then
						'msgbox "Error: Unable to Obtain ACE (" & I & ")"
						Exit Sub
					End If
					' You have a pointer to the ACE. Place it
					' into a structure, so you can get at its size.
					'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object sCurrentACE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CopyMemory(sCurrentACE, pCurrentAce, LenB(sCurrentACE))
					sCurrentACE.Header.AceFlags = OBJECT_INHERIT_ACE + INHERIT_ONLY_ACE 'NO_PROPAGATE_INHERIT_ACE
					'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
					'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
					CopyMemory(pCurrentAce, VarPtr(sCurrentACE), LenB(sCurrentACE))
					
					'add another ACE for files
					lResult = AddAccessAllowedAce(bNewACL(0), ACL_REVISION, lMask, bUserSid(0))
					
					' Make sure added the ACL to the DACL.
					If (lResult = 0) Then
						'msgbox "Error: Unable to Add ACL to DACL"
						Exit Sub
					End If
					
					' Attempt to grab the next ACE.
					'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
					lResult = GetAce(VarPtr(bNewACL(0)), nRecordNumber + 1, pCurrentAce)
					
					' Make sure you have the current ACE under question.
					If (lResult = 0) Then
						'msgbox "Error: Unable to Obtain ACE (" & I & ")"
						Exit Sub
					End If
					
					'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object sCurrentACE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CopyMemory(sCurrentACE, pCurrentAce, LenB(sCurrentACE))
					sCurrentACE.Header.AceFlags = CONTAINER_INHERIT_ACE
					'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
					'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
					CopyMemory(pCurrentAce, VarPtr(sCurrentACE), LenB(sCurrentACE))
				End If
				
				
				' Set the file's Security Descriptor to the new DACL.
				lResult = SetSecurityDescriptorDacl(sNewSD, 1, bNewACL(0), 0)
				
				' Make sure you set the SD to the new DACL.
				If (lResult = 0) Then
					'msgbox "Error: " & "Unable to Set New DACL to Security Descriptor"
					Exit Sub
				End If
				
				' The final step is to add the Security Descriptor back to
				' the file!
				lResult = SetFileSecurity(sFileName, DACL_SECURITY_INFORMATION, sNewSD)
				
				' Make sure you added the Security Descriptor to the file!
				If (lResult = 0) Then
					'msgbox "Error: Unable to Set New Security Descriptor " _
					'& " to File : " & sFileName
					'msgbox Err.LastDllError
				Else
					'msgbox "Updated Security Descriptor on File: " _
					'& sFileName
				End If
				
			End If
			
		End If
		
	End Sub
End Module