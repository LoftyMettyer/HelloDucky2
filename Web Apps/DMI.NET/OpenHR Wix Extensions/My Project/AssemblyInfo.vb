Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports Microsoft.Tools.WindowsInstallerXml

' General Information about an assembly is controlled through the following 
' set of attributes. Change these attribute values to modify the information
' associated with an assembly.

' Review the values of the assembly attributes

<Assembly: AssemblyTitle("OpenHR Wix Extensions")> 
<Assembly: AssemblyDescription("")> 
<Assembly: AssemblyCompany("")> 
<Assembly: AssemblyProduct("OpenHR Wix Extensions")> 
<Assembly: AssemblyCopyright("Copyright ©  2014")> 
<Assembly: AssemblyTrademark("")> 

<Assembly: ComVisible(False)>

'The following GUID is for the ID of the typelib if this project is exposed to COM
<Assembly: Guid("2c49cbe8-fa8f-48b7-ad5a-041c2e518db6")> 

' Version information for an assembly consists of the following four values:
'
'      Major Version
'      Minor Version 
'      Build Number
'      Revision
'
' You can specify all the values or you can default the Build and Revision Numbers 
' by using the '*' as shown below:
' <Assembly: AssemblyVersion("1.0.*")> 

<Assembly: AssemblyVersion("1.0.0.0")> 
<Assembly: AssemblyFileVersion("1.0.0.0")> 

<Assembly: AssemblyDefaultWixExtension(GetType(WixFileVersionExtension))> 