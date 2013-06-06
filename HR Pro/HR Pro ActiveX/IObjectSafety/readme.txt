To create the 'objsafe.tlb' type library:

1) Get latest of "$/HR Pro v2/HR Pro ActiveX/IObjectSafety" VSS project from the "HR Pro Version 2" VSS Database.

2) Open command prompt.

3) Change the current directory to the location of the 'MKTYPLIB.EXE' and 'objsafe.odl' files.

4) Run objSafe.BAT or execute the following command:
   "MKTYPLIB objsafe.odl /tlb objsafe.tlb /win32 /nocpp"

5) Copy the 'objsafe.tlb' file to the "<Windows>\System32" folder. 
   
NOTE: The 'objsafe.tlb' file does not necessarily need to be in the 'System32' folder. However, by doing this VB projects automatically pick up the reference to this file, as opposed to browsing to the file in project references.

---------------------------------
For further information reference Microsoft's Article 182598 Revision 3.