
This readme is referring to the subfolder "OLEGuids", which contains two files:

-	OLEGuids.tlb
	
	This is ONLY needed in the IDE.
	Copy that file into the windows system directory.
	In order to use that type library, it is necessary to load the OLEGuids.tlb in the IDE.
	(Project -> References... -> Browse -> select file OLEGuids.tlb -> check item "OLE Guid and interface definitions")
	
	Info: This is a amended version of the original from vbaccelerator.
	(http://www.vbaccelerator.com/home/VB/Type_Libraries/OLE_GUID_and_Interface_Definitions/article.asp)
	
	The uuid and library name differs from the original to prevent conflicts.
	
-	OLEGuids.odl
	
	That is the file containing the source code.