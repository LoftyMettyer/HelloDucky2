Attribute VB_Name = "modQAddressSpecifics"
Option Explicit

' What type of AFD is installed
Public Enum QAddressTypes
  QADDRESS_DISABLED = 0
  QADDRESS_RAPID = 1
  QADDRESS_PRO3 = 2
  QADDRESS_WORLDWIDE = 3
  QADDRESS_PRO4 = 4
  QADDRESS_PRO5 = 5
End Enum

Public giQAddressEnabled As QAddressTypes
Public mlngHandle As Long
Global Const qapro_AREARECODED = -9811

' From demo file
Global Const qaerr_FATAL = -1000
Global Const qaerr_NOMEMORY = -1001
Global Const qaerr_INITOOLARGE = -1005
Global Const qaerr_ININOEXTEND = -1006
Global Const qaerr_FILEOPEN = -1010
Global Const qaerr_FILEEXIST = -1011
Global Const qaerr_FILEREAD = -1012
Global Const qaerr_FILEWRITE = -1013
Global Const qaerr_FILEDELETE = -1014
Global Const qaerr_FILEACCESS = -1016
Global Const qaerr_FILEVERSION = -1017
Global Const qaerr_FILEHANDLE = -1018
Global Const qaerr_FILECREATE = -1019
Global Const qaerr_FILERENAME = -1020
Global Const qaerr_FILEEXPIRED = -1021
Global Const qaerr_FILENOTDEMO = -1022
Global Const qaerr_READFAIL = -1025
Global Const qaerr_WRITEFAIL = -1026
Global Const qaerr_BADDRIVE = -1027
Global Const qaerr_BADDIR = -1028
Global Const qaerr_DIRCREATE = -1029
Global Const qaerr_BADOPTION = -1030
Global Const qaerr_BADINIFILE = -1031
Global Const qaerr_BADLOGFILE = -1032
Global Const qaerr_BADMEMORY = -1033
Global Const qaerr_BADHOTKEY = -1034
Global Const qaerr_HOTKEYUSED = -1035
Global Const qaerr_BADRESOURCE = -1036
Global Const qaerr_BADDATADIR = -1037
Global Const qaerr_BADTEMPDIR = -1038
Global Const qaerr_NOTDEFINED = -1040
Global Const qaerr_DUPLICATE = -1041
Global Const qaerr_BADACTION = -1042
Global Const qaerr_CCFAILURE = -1050
Global Const qaerr_CCBADCODE = -1051
Global Const qaerr_CCACCESS = -1052
Global Const qaerr_CCNODONGLE = -1053
Global Const qaerr_CCNOUNITS = -1054
Global Const qaerr_CCNOMETER = -1055
Global Const qaerr_CCNOFEATURE = -1056
Global Const qaerr_CCINSTALL = -1060
Global Const qaerr_CCEXPIRED = -1061
Global Const qaerr_CCDATETIME = -1062
Global Const qaerr_CCUSERLIMIT = -1063
Global Const qaerr_CCACTIVATE = -1064
Global Const qaerr_CCBADDRIVE = -1065
Global Const qaerr_UNAUTHORISED = -1070
Global Const qaerr_NOTHREAD = -1080
Global Const qaerr_NOTLSMEMORY = -1081
Global Const qaerr_NOTASK = -1090

Global Const qaerr_RAPIDOPEN = -9800
Global Const qaerr_NONRAPIDFILE = -9801
Global Const qaerr_INVALIDAREA = -9802
Global Const qaerr_AREALEVEL = -9803
Global Const qaerr_DISTRICTLEVEL = -9804
Global Const qaerr_SECTORLEVEL = -9805
Global Const qaerr_HALFSECTORLEVEL = -9806
Global Const qaerr_NOCODES = -9807
Global Const qaerr_NOAREADATA = -9809
Global Const qaerr_NUMBEREDFLAT = -9810
Global Const qaerr_POSTCODERECODED = -9811
Global Const qaerr_SUBSMADE = -9813
Global Const qaerr_BADPOSTCODECHAR = -9814
Global Const qaerr_RAPIDNOTSTARTED = -9815

Declare Sub QAInitialise Lib "QAPUIEB.DLL" (ByVal vi1 As Long)
Declare Sub QAErrorMessage Lib "QAPUIEB.DLL" (ByVal vi1 As Long, ByVal rs2 As String, ByVal vi3 As Long)
Declare Function QAErrorLevel Lib "QAPUIEB.DLL" (ByVal vi1 As Long) As Long
Declare Function QAErrorHistory Lib "QAPUIEB.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal rs3 As String, ByVal vi4 As Long) As Long
Declare Sub QAVersionInfo Lib "QAPUIEB.DLL" (ByVal rs1 As String, ByVal vi2 As Long)
Declare Function QADataInfo Lib "QAPUIEB.DLL" (ByVal rs1 As String, ByVal vi2 As Long, ri3 As Long) As Long
Declare Function QASystemInfo Lib "QAPUIEB.DLL" (ByVal vi1 As Long, ByVal rs2 As String, ByVal vi3 As Long) As Long
Declare Sub QAUpdateKey Lib "QAPUIEB.DLL" (ByVal rs1 As String, ByVal vi2 As Long)
Declare Function QAUpdateCode Lib "QAPUIEB.DLL" (ByVal vs1 As String) As Long
Declare Function QALicenseInfo Lib "QAPUIEB.DLL" (ri1 As Long, ri2 As Long, ri3 As Long) As Long
Declare Function QAAuthorise Lib "QAPUIEB.DLL" (ByVal vs1 As String, ByVal vl2 As Long) As Long


Global Const matchlevel_NOLOOKUP = 0
Global Const matchlevel_AREANOTFOUND = 2
Global Const matchlevel_AREA = 3
Global Const matchlevel_DISTRICT = 4
Global Const matchlevel_SECTOR = 5
Global Const matchlevel_HALFSECTOR = 6
Global Const matchlevel_POSTCODE = 7
Global Const matchlevel_DELIVPT = 8
Global Const storelevel_NECESSARYPOST = 0
Global Const storelevel_EXACTPOST = 1
Global Const storelevel_DELIVPT = 2
Global Const storelevel_HOUSE = 3
Global Const storelevel_NAME = 4
Global Const dpst_EMPTY = 0
Global Const dpst_CLOSED = 1
Global Const dpst_OPENFAILED = 2
Global Const dpst_OPENING = 3
Global Const dpst_ATTACHING = 4
Global Const dpst_OPEN = 5
Global Const dpst_ATTACHED = 6
Global Const dpfmt_STANDARD = 0
Global Const dpfmt_DISPLAY = 1
Global Const dpfmt_LANDRANGER = 2
Global Const dpfmt_10KM = 4
Global Const dpfmt_1KM = 8
Global Const dpfmt_100M = 12
Global Const dpfmt_10M = 16
Global Const dpfmt_1M = 20
Global Const dpfmt_ADDP = 24
Global Const dp_GRIDRESBITS = 28
Global Const dpfmt_SPACEPAD = 32
Global Const dpfmt_OLDGRID = 64
Global Const dpfmt_OLDAKEY = 65
Global Const dpfmt_BMWAKEY = 66
Global Const dpfmt_DELV1 = 67
Global Const dpfmt_DELV2 = 68
Global Const codetype_TEXT = 0
Global Const codetype_BINARY = 1
Global Const qaerr_DATNOTSTARTED = -1700
Global Const qaerr_NODATASETSOPEN = -1701
Global Const qaerr_DATISSTARTED = -1702
Global Const qaerr_NOTOPEN = -1703
Global Const qaerr_CLOSEERRORS = -1704
Global Const qaerr_ISOPEN = -1705
Global Const qaerr_NEEDLATERREV = -1709
Global Const qaerr_CANTOPENINFFILE = -1710
Global Const qaerr_BADINFFILE = -1711
Global Const qaerr_CANTOPENDATFILE = -1712
Global Const qaerr_BADDATFILE = -1713
Global Const qaerr_NOTRAILER = -1714
Global Const qaerr_FILEMISMATCH = -1715
Global Const qaerr_COPYCONTROL = -1716
Global Const qaerr_NOUNITS = -1717
Global Const qaerr_DATASETCONFIG = -1718
Global Const qaerr_V1DATASET = -1719
Global Const qaerr_POSTCODEINVALID = -1720
Global Const qaerr_DPPINVALID = -1721
Global Const qaerr_DATAFORMAT = -1722
Global Const qaerr_DATASETNOTFOUND = -1730
Global Const qaerr_IDEMPTY = -1731
Global Const qaerr_IDINUSE = -1732
Global Const qaerr_IDINVALID = -1733
Global Const qaerr_CODESELINVALID = -1734
Global Const qaerr_ITEMSELINVALID = -1735
Global Const qaerr_ITEMNAMEINVALID = -1736
Global Const qaerr_INFIDINVALID = -1737
Global Const qaerr_NOLOOKUP = -1740
Global Const qaerr_NOCODE = -1741
Global Const qaerr_CODETRUNC = -1742
Global Const qaerr_NOINF = -1743
Global Const qaerr_BLANKINF = -1744
Global Const qaerr_INFTRUNC = -1745
Global Const qaerr_CODEINFTRUNC = -1746
Global Const qaerr_ITEMTRUNC = -1747
Global Const qaerr_CODESONLY = -1748
Global Const qaerr_NOBUFFER = -1749
Global Const qaerr_BUFFERTOOSMALL = -1750
Global Const qaerr_NODPSDATA = -1760
Global Const qaerr_POSTCODENOTFOUND = -1761
Global Const qaerr_DPSNOTFOUND = -1762
Global Const qaerr_DPSINVALID = -1763
Global Const dpinf_TITLE = 0
Global Const dpinf_VERSION = 1
Global Const dpinf_STORELEVEL = 2
Global Const dpinf_MULTCODES = 3
Global Const dpinf_VARILENCODES = 4
Global Const dpinf_CODETYPE = 5
Global Const dpinf_METERED = 6
Global Const dpinf_CODENAME = 7
Global Const dpinf_MAXCODELEN = 8
Global Const dpinf_CODEITEMSEP = 9
Global Const dpinf_INFNAME = 10
Global Const dpinf_MAXINFLEN = 11
Global Const dpinf_INFITEMSEP = 12

Declare Function QADP_Startup Lib "QAPUIEB.DLL" (ByVal vs1 As String, ByVal vs2 As String) As Long
Declare Function QADP_Shutdown Lib "QAPUIEB.DLL" () As Long
Declare Function QADP_Open Lib "QAPUIEB.DLL" (ByVal vs1 As String, ByVal vs2 As String, ByVal vs3 As String, ByVal vi4 As Long) As Long
Declare Function QADP_Close Lib "QAPUIEB.DLL" (ByVal vi1 As Long) As Long
Declare Function QADP_DataSpec Lib "QAPUIEB.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal rs3 As String, ByVal vi4 As Long, ByVal rs5 As String, ByVal vi6 As Long) As Long
Declare Function QADP_ItemCount Lib "QAPUIEB.DLL" (ByVal vi1 As Long) As Long
Declare Function QADP_ID Lib "QAPUIEB.DLL" (ByVal vs1 As String) As Long
Declare Function QADP_Name Lib "QAPUIEB.DLL" (ByVal vi1 As Long, ByVal rs2 As String, ByVal vi3 As Long) As Long
Declare Function QADP_ItemName Lib "QAPUIEB.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal rs3 As String, ByVal vi4 As Long) As Long
Declare Function QADP_MaxID Lib "QAPUIEB.DLL" () As Long
Declare Function QADP_OpenCount Lib "QAPUIEB.DLL" () As Long
Declare Function QADP_LookUp Lib "QAPUIEB.DLL" (ByVal vs1 As String, ByVal vi2 As Long) As Long
Declare Function QADP_LookUpOne Lib "QAPUIEB.DLL" (ByVal vi1 As Long, ByVal vs2 As String, ByVal vi3 As Long) As Long
Declare Function QADP_EndLookUp Lib "QAPUIEB.DLL" () As Long
Declare Function QADP_CodeCount Lib "QAPUIEB.DLL" (ByVal vi1 As Long) As Long
Declare Function QADP_Get Lib "QAPUIEB.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal rs3 As String, ByVal vi4 As Long, ByVal rs5 As String, ByVal vi6 As Long) As Long
Declare Function QADP_GetItem Lib "QAPUIEB.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal vi3 As Long, ByVal rs4 As String, ByVal vi5 As Long, ByVal rs6 As String, ByVal vi7 As Long) As Long
Declare Function QADP_GetItemByName Lib "QAPUIEB.DLL" (ByVal vs1 As String, ByVal vi2 As Long, ByVal vs3 As String, ByVal rs4 As String, ByVal vi5 As Long) As Long
Declare Function QADP_GetDPP Lib "QAPUIEB.DLL" (ByVal vs1 As String) As Long
Declare Function QADP_Status Lib "QAPUIEB.DLL" (ByVal vi1 As Long) As Long
Declare Function QADP_State Lib "QAPUIEB.DLL" (ByVal vi1 As Long) As Long
Declare Function QADP_MatchLevel Lib "QAPUIEB.DLL" (ByVal vi1 As Long) As Long
Declare Function QADP_Release Lib "QAPUIEB.DLL" (ByVal vi1 As Long) As Long

Global Const qaerr_CANCELLED = -9982
Global Const qaerr_TIMEDOUT = -10001
Global Const qaopevent_START = 0
Global Const qaopevent_PROGRESSINFO = 1
Global Const qaopevent_POLLCANCEL = 2
Global Const qaopevent_STOP = 3

Declare Function QAProKey_Open Lib "QAPUIEB.DLL" () As Long
Declare Function QAProKey_Close Lib "QAPUIEB.DLL" () As Long
Declare Function QAProKey_Search Lib "QAPUIEB.DLL" (ByVal vi1 As Long, ByVal vs2 As String, ByVal vs3 As String, ri4 As Long) As Long
Declare Function QAProKey_Get Lib "QAPUIEB.DLL" (ByVal vi1 As Long, ByVal rs2 As String, ri3 As Long, ri4 As Long, ByVal rs5 As String, ByVal rs6 As String) As Long

Declare Function QARapid_Open Lib "QARUIEB.DLL" (ByVal vs1 As String, ByVal vs2 As String) As Long
Declare Sub QARapid_Close Lib "QARUIEB.DLL" ()
Declare Function QARapid_ChangeFormat Lib "QARUIEB.DLL" (ByVal vs1 As String) As Long
Declare Function QARapid_Search Lib "QARUIEB.DLL" (ByVal vs1 As String) As Long
Declare Sub QARapid_EndSearch Lib "QARUIEB.DLL" ()
Declare Function QARapid_Count Lib "QARUIEB.DLL" () As Long
Declare Function QARapid_ListItem Lib "QARUIEB.DLL" (ByVal vl1 As Long, ByVal rs2 As String, ByVal vi3 As Long, ByVal vi4 As Long) As Long
Declare Function QARapid_AddrLine Lib "QARUIEB.DLL" (ByVal vl1 As Long, ByVal vi2 As Long, ByVal rs3 As String, ByVal vi4 As Long) As Long
Declare Function QARapid_FormatLine Lib "QARUIEB.DLL" (ByVal vl1 As Long, ByVal vi2 As Long, ByVal rs3 As String, ByVal vi4 As Long) As Long
Declare Function QARapid_FormatAddr Lib "QARUIEB.DLL" (ByVal vl1 As Long, ByVal rs2 As String, ByVal vi3 As Long) As Long

Global Const qaerr_PRODATAEXPIRED = -9966
Global Const qaerr_STREETSVERSION = -9967
Global Const qaerr_TOOSHORT = -9968
Global Const qaerr_PREMISESVERSION = -9969
Global Const qaerr_INDEXVERSION = -9970
Global Const qaerr_TOOLONG = -9971
Global Const qaerr_BADINFO = -9972
Global Const qaerr_NESTEDTOODEEP = -9973
Global Const qaerr_BADPHASE = -9974
Global Const qaerr_STREETSOPEN = -9975
Global Const qaerr_PREMISESOPEN = -9976
Global Const qaerr_INDEXOPEN = -9977
Global Const qaerr_NOMATCH = -9978
Global Const qaerr_TOOMANYCODES = -9979
Global Const qaerr_TOOMANYMATCHES = -9980
Global Const qaerr_NOSUCHAREA = -9981
Global Const qaerr_NODP = -9983
Global Const qaerr_NONUMBER = -9984
Global Const qaerr_CANTSTEP = -9987
Global Const qaerr_SYNTAX = -9988
Global Const qaerr_BADFILEFORMAT = -9989
Global Const qaerr_SURNAMEINDEXOPEN = -9990
Global Const qaerr_FORENAMEINDEXOPEN = -9991
Global Const qaerr_BADSEARCHDESC = -10000
Global Const qapro_STEPINFO = 0
Global Const qapro_RANGEINFO = 1
Global Const qapro_TYPEINFO = 2
Global Const qapro_DPPINFO = 3
Global Const qapro_QUALITYINFO = 4

Declare Function QAPro_Open Lib "QAPUIEB.DLL" (ByVal vs1 As String, ByVal vs2 As String) As Long
Declare Sub QAPro_Close Lib "QAPUIEB.DLL" ()
Declare Sub QAPro_SetTimeout Lib "QAPUIEB.DLL" (ByVal vl1 As Long)
Declare Function QAPro_ChangeFormat Lib "QAPUIEB.DLL" (ByVal vs1 As String) As Long
Declare Function QAPro_Search Lib "QAPUIEB.DLL" (ByVal vs1 As String) As Long
Declare Function QAPro_SearchDPP Lib "QAPUIEB.DLL" (ByVal vs1 As String, ByVal vi2 As Long) As Long
Declare Sub QAPro_EndSearch Lib "QAPUIEB.DLL" ()
Declare Function QAPro_Count Lib "QAPUIEB.DLL" () As Long
Declare Function QAPro_ListItem Lib "QAPUIEB.DLL" (ByVal vl1 As Long, ByVal rs2 As String, ByVal vi3 As Long, ByVal vi4 As Long) As Long
Declare Function QAPro_StepIn Lib "QAPUIEB.DLL" (ByVal vl1 As Long) As Long
Declare Function QAPro_StepOut Lib "QAPUIEB.DLL" () As Long
Declare Function QAPro_Select Lib "QAPUIEB.DLL" (ByVal vs1 As String, ByVal vi2 As Long, ByVal vi3 As Long) As Long
Declare Function QAPro_EndSelect Lib "QAPUIEB.DLL" () As Long
Declare Function QAPro_Pick Lib "QAPUIEB.DLL" (ByVal vs1 As String, ByVal vi2 As Long) As Long
Declare Function QAPro_EndPick Lib "QAPUIEB.DLL" () As Long
Declare Function QAPro_Back Lib "QAPUIEB.DLL" () As Long
Declare Function QAPro_First Lib "QAPUIEB.DLL" (ByVal vs1 As String, ByVal vi2 As Long, rl3 As Long) As Long
Declare Function QAPro_AddrLine Lib "QAPUIEB.DLL" (ByVal vl1 As Long, ByVal vi2 As Long, ByVal vs3 As String, ByVal rs4 As String, ByVal vi5 As Long) As Long
Declare Function QAPro_FormatLine Lib "QAPUIEB.DLL" (ByVal vl1 As Long, ByVal vi2 As Long, ByVal vs3 As String, ByVal rs4 As String, ByVal vi5 As Long) As Long
Declare Function QAPro_FormatAddr Lib "QAPUIEB.DLL" (ByVal vl1 As Long, ByVal vs2 As String, ByVal rs3 As String, ByVal vi4 As Long) As Long
Declare Function QAPro_FormatCount Lib "QAPUIEB.DLL" () As Long
Declare Function QAPro_GetItemInfo Lib "QAPUIEB.DLL" (ByVal vl1 As Long, ByVal vi2 As Long, rl3 As Long) As Long
Declare Function QAPro_GetDPP Lib "QAPUIEB.DLL" (ByVal vl1 As Long, ByVal vs2 As String, ri3 As Long) As Long
Declare Sub QAPro_CCUserUpdate Lib "QAPUIEB.DLL" ()
Declare Function QAPro_CCReadCounter Lib "QAPUIEB.DLL" () As Long

Global Const qaattribs_NONE = 0
Global Const qaattribs_NODIALOG = 1
Global Const qaattribs_NOWAITENTER = 2
Global Const qaattribs_NOLAYOUTBUTTON = 4
Global Const qaattribs_NOHELPBUTTON = 8
Global Const qaattribs_NOWARNCC = 16
Global Const qaattribs_SUPPRESSNAMES = 32
Global Const qaattribs_NONAMES = 64
Global Const qaattribs_KEEPRESULTS = 16384
Global Const qaattribs_DEMOMODE = 32768
Global Const qaret_SUCCESS = 1
Global Const qaret_OVERFLOW = 2
Global Const qaret_FIELDTRUNCATED = 4
Global Const qaret_POSTCODERECODED = 8
Global Const qaret_NUMBEREDFLAT = 16
Global Const qaret_SUBSMADE = 32
Global Const qaerr_UINOTSTARTED = -1900
Global Const qaerr_UISTARTED = -1901
Global Const qaerr_NORESULTS = -1920
Global Const qaerr_MOREMATCHES = -1921
Global Const qaerr_SEARCHRECODED = -1922
Global Const qaerr_NOTEXACT = -1923
Global Const qaerr_DATALOST = -1924

Declare Function QAProUI_Startup Lib "QAPUIEB.DLL" (ByVal vs1 As String, ByVal vs2 As String, ByVal vs3 As String, ByVal vl4 As Long) As Long
Declare Function QAProUI_DPPopup Lib "QAPUIEB.DLL" (ByVal vs1 As String, ByVal rs2 As String, ByVal vi3 As Long, ByVal vs4 As String, ri5 As Long) As Long
Declare Function QAProUI_Popup Lib "QAPUIEB.DLL" (ByVal vs1 As String, ByVal rs2 As String, ByVal vi3 As Long) As Long
Declare Sub QAProUI_Shutdown Lib "QAPUIEB.DLL" (ByVal vi1 As Long)
Declare Function QAProUI_Config Lib "QAPUIEB.DLL" (ByVal vs1 As String) As Long
Declare Sub QAProUI_IniSection Lib "QAPUIEB.DLL" (ByVal rs1 As String, ByVal vi2 As Long)

' NPG20081126 Fault 13364
Global Const qaerr_INITINSTANCE = -1002
Global Const qaerr_BADINTERFACE = -1003
Global Const qaerr_FILETOOLARGE = -1008
Global Const qaerr_FILECHGDETECT = -1009
Global Const qaerr_FILETIMEGET = -1023
Global Const qaerr_FILETIMESET = -1024
Global Const qaerr_BADDATE = -1045
Global Const qaerr_BADTIMEZONE = -1046
Global Const qaerr_CCINVALID = -1057
Global Const qaerr_CCREGISTER = -1066
Global Const qaerr_NOLOCALEFILE = -1074
Global Const qaerr_BADLOCALEFILE = -1075
Global Const qaerr_BADLOCALE = -1076
Global Const qaerr_BADCODEPAGE = -1077
Global Const qaerr_RESOURCEFAIL = -1078
Global Const qaerr_TOOMANYINSTANCES = -4501
Global Const qaerr_MAXRESOURCES = -4503
Global Const qaerr_OPENFAILURE = -4551
Global Const qaerr_APIHANDLE = -4552
Global Const qaerr_OUTOFSEQUENCE = -4553
Global Const qaerr_BUSYHANDLE = -4554
Global Const qaerr_BADINDEX = -4556
Global Const qaerr_BADVALUE = -4557
Global Const qaerr_BADPARAM = -4558
Global Const qaerr_PARAMTRUNCATED = -4559
Global Const qaerr_NOENGINE = -4560
Global Const qaerr_BADLAYOUT = -4561
Global Const qaerr_BADSTEP = -4562
Global Const qaerr_DATASETNOTAVAILABLE = -4570
Global Const qaerr_LICENSINGFAILURE = -4571
Global Const qaerr_NOACTIVEDATASET = -4572
Global Const qaerr_BADCOUNTRYLIST = -4573
Global Const qaerr_DATAMAPNOTAVAILABLE = -4574
Global Const qaerr_SERVERCONNLOST = -4580
Global Const qaerr_SERVERFULL = -4581
Global Const qaerr_BADMONIKER = -4590
Global Const qaerr_MONIKEREXPIRED = -4591
Global Const NO_HANDLE = 0
Global Const qalibflags_NONE = 0
Global Const qavalue_FALSE = 0
Global Const qavalue_TRUE = 1
Global Const qaengine_SINGLELINE = 1
Global Const qaengine_TYPEDOWN = 2
Global Const qaengine_BATCH = 3
Global Const qaengine_VERIFICATION = 4
Global Const qaengine_KEYFINDER = 5
Global Const qaengopt_DEFAULT = 0
Global Const qaengopt_ASYNCSEARCH = 1
Global Const qaengopt_ASYNCSTEPIN = 2
Global Const qaengopt_ASYNCREFINE = 3
Global Const qaengopt_THRESHOLD = 6
Global Const qaengopt_TIMEOUT = 7
Global Const qaengopt_SEARCHINTENSITY = 8
Global Const qaintensity_EXACT = 0
Global Const qaintensity_CLOSE = 1
Global Const qaintensity_EXTENSIVE = 2
Global Const qastate_NOSEARCH = 1
Global Const qastate_STILLSEARCHING = 2
Global Const qastate_TIMEOUT = 4
Global Const qastate_SEARCHCANCELLED = 8
Global Const qastate_MAXMATCHES = 16
Global Const qastate_OVERTHRESHOLD = 32
Global Const qastate_LARGEPOTENTIAL = 64
Global Const qastate_MOREOTHERMATCHES = 128
Global Const qastate_REFINING = 256
Global Const qastate_AUTOSTEPINSAFE = 512
Global Const qastate_AUTOSTEPINPASTCLOSE = 1024
Global Const qastate_CANSTEPOUT = 2048
Global Const qastate_AUTOFORMATSAFE = 4096
Global Const qastate_AUTOFORMATPASTCLOSE = 8192
Global Const qassint_PICKLISTSIZE = 1
Global Const qassint_POTENTIALMATCHES = 2
Global Const qassint_SEARCHSTATE = 3
Global Const qassint_ISNOSEARCH = 4
Global Const qassint_ISSTILLSEARCHING = 5
Global Const qassint_ISTIMEOUT = 6
Global Const qassint_ISSEARCHCANCELLED = 7
Global Const qassint_ISMAXMATCHES = 8
Global Const qassint_ISOVERTHRESHOLD = 9
Global Const qassint_ISLARGEPOTENTIAL = 10
Global Const qassint_ISMOREOTHERMATCHES = 11
Global Const qassint_ISREFINING = 12
Global Const qassint_ISAUTOSTEPINSAFE = 17
Global Const qassint_ISAUTOSTEPINPASTCLOSE = 18
Global Const qassint_CANSTEPOUT = 19
Global Const qassint_ISAUTOFORMATSAFE = 20
Global Const qassint_ISAUTOFORMATPASTCLOSE = 21
Global Const qaresult_FULLADDRESS = 1
Global Const qaresult_MULTIPLES = 2
Global Const qaresult_CANSTEP = 4
Global Const qaresult_ALIASMATCH = 8
Global Const qaresult_POSTCODERECODED = 16
Global Const qaresult_CROSSBORDERMATCH = 32
Global Const qaresult_DUMMYPOBOX = 64
Global Const qaresult_NAME = 256
Global Const qaresult_INFORMATION = 1024
Global Const qaresult_WARNINFORMATION = 2048
Global Const qaresult_INCOMPLETEADDR = 4096
Global Const qaresult_UNRESOLVABLERANGE = 8192
Global Const qaresult_INCLUDESUSERDATA = 16384
Global Const qaresult_PHANTOMPRIMARYPOINT = 32768
Global Const qaresult_RESOLVEDPPP = 65536
Global Const qaresultstr_DATAID = 1
Global Const qaresultstr_DESCRIPTION = 2
Global Const qaresultstr_PARTIALADDRESS = 3
Global Const qaresultstr_MATCHTYPE = 5
Global Const qaresultint_CONFIDENCE = 6
Global Const qaresultint_UNUSEDLINES = 7
Global Const qaresultint_POSTCODEACTION = 8
Global Const qaresultint_ADDRESSACTION = 9
Global Const qaresultint_GENERICINFO = 10
Global Const qaresultint_COUNTRYINFO1 = 11
Global Const qaresultint_COUNTRYINFO2 = 12
Global Const qaresultint_ISFULLADDRESS = 13
Global Const qaresultint_ISMULTIPLES = 14
Global Const qaresultint_ISCANSTEP = 15
Global Const qaresultint_ISALIASMATCH = 16
Global Const qaresultint_ISPOSTCODERECODED = 17
Global Const qaresultint_ISCROSSBORDERMATCH = 18
Global Const qaresultint_ISDUMMYPOBOX = 19
Global Const qaresultint_ISNAME = 20
Global Const qaresultint_ISINFORMATION = 21
Global Const qaresultint_ISWARNINFORMATION = 22
Global Const qaresultint_ISINCOMPLETEADDR = 23
Global Const qaresultint_ISUNRESOLVABLERANGE = 24
Global Const qaresultint_ISINCLUDEUSERDATA = 25
Global Const qaresultstr_ADDRMATCHCODE = 26
Global Const qaresultstr_POSTCODEMATCHED = 27
Global Const qaresultint_ISPHANTOMPRIMARYPOINT = 28
Global Const qapromptint_LINECOUNT = 1
Global Const qapromptint_DYNAMIC = 2
Global Const qadatastr_ID = 1
Global Const qadatastr_DESCRIPTION = 2
Global Const qadatastr_BASE = 3
Global Const qalicwarn_NONE = 0
Global Const qalicwarn_DATAEXPIRING = 10
Global Const qalicwarn_LICENCEEXPIRING = 20
Global Const qalicwarn_CLICKSLOW = 25
Global Const qalicwarn_EVALUATION = 30
Global Const qalicwarn_NOCLICKS = 35
Global Const qalicwarn_DATAEXPIRED = 40
Global Const qalicwarn_EVALLICENCEEXPIRED = 50
Global Const qalicwarn_FULLLICENCEEXPIRED = 60
Global Const qalicwarn_LICENCENOTFOUND = 70
Global Const qalicwarn_DATAUNREADABLE = 80
Global Const qalicencestr_ID = 1
Global Const qalicencestr_DESCRIPTION = 2
Global Const qalicencestr_COPYRIGHT = 3
Global Const qalicencestr_VERSION = 4
Global Const qalicencestr_BASECOUNTRY = 5
Global Const qalicencestr_STATUS = 6
Global Const qalicencestr_SERVER = 7
Global Const qalicenceint_WARNINGLEVEL = 8
Global Const qalicenceint_DAYSLEFT = 9
Global Const qalicenceint_DATADAYSLEFT = 10
Global Const qalicenceint_LICENCEDAYSLEFT = 11
Global Const qacancelflag_NONE = 0
Global Const qacancelflag_BLOCKING = 1
Global Const qaformat_OVERFLOW = 1
Global Const qaformat_TRUNCATED = 2
Global Const qaformat_DPVVALID = 16
Global Const qaformat_DPVINVALID = 32
Global Const qaformat_DPVLOCKED = 64
Global Const qaformatted_NONE = 0
Global Const qaformatted_NAME = 1
Global Const qaformatted_ADDRESS = 2
Global Const qaformatted_ANCILLARY = 4
Global Const qaformatted_DATAPLUS = 8
Global Const qaformatted_TRUNCATED = 16
Global Const qaformatted_OVERFLOW = 32
Global Const qaformatted_DATAPLUSSYNTAX = 64
Global Const qaformatted_DATAPLUSEXPIRED = 128
Global Const qaformatted_DATAPLUSBLANK = 256
Global Const qaformatted_UNMATCHED = 512
Global Const qasysinfo_SYSTEM = 1
Global Const qaunusedstr_TEXT = 1
Global Const qaunusedint_COMPLETENESS = 2
Global Const qaunusedint_TYPE = 3
Global Const qaunusedint_POSITION = 4
Global Const qaunusedint_ISCAREOF = 5
Global Const qaunusedint_ISPREMSUFFIX = 6

Declare Function QA_SetLibraryFlags Lib "QAUPIED.DLL" (ByVal vl1 As Long) As Long
Declare Function QA_Open Lib "QAUPIED.DLL" (ByVal vs1 As String, ByVal vs2 As String, ri3 As Long) As Long
Declare Function QA_Close Lib "QAUPIED.DLL" (ByVal vi1 As Long) As Long
Declare Sub QA_Shutdown Lib "QAUPIED.DLL" ()
Declare Function QA_SetEngine Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long) As Long
Declare Function QA_GetEngine Lib "QAUPIED.DLL" (ByVal vi1 As Long, ri2 As Long) As Long
Declare Function QA_SetEngineOption Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal vl3 As Long) As Long
Declare Function QA_GetEngineOption Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, rl3 As Long) As Long
Declare Function QA_GetEngineStatus Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ri3 As Long) As Long
Declare Function QA_GetPromptStatus Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ri3 As Long, ByVal rs4 As String, ByVal vi5 As Long) As Long
Declare Function QA_GetPrompt Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal rs3 As String, ByVal vi4 As Long, ri5 As Long, ByVal rs6 As String, ByVal vi7 As Long) As Long
Declare Function QA_Search Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vs2 As String) As Long
Declare Function QA_CancelSearch Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vl2 As Long) As Long
Declare Function QA_EndSearch Lib "QAUPIED.DLL" (ByVal vi1 As Long) As Long
Declare Function QA_GetSearchStatus Lib "QAUPIED.DLL" (ByVal vi1 As Long, ri2 As Long, ri3 As Long, rl4 As Long) As Long
Declare Function QA_GetSearchStatusDetail Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, rl3 As Long, ByVal rs4 As String, ByVal vi5 As Long) As Long
Declare Function QA_StepIn Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long) As Long
Declare Function QA_StepOut Lib "QAUPIED.DLL" (ByVal vi1 As Long) As Long
Declare Function QA_GetResult Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal rs3 As String, ByVal vi4 As Long, ri5 As Long, rl6 As Long) As Long
Declare Function QA_GetResultDetail Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal vi3 As Long, rl4 As Long, ByVal rs5 As String, ByVal vi6 As Long) As Long
Declare Function QA_FormatResult Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal vs3 As String, ri4 As Long, rl5 As Long) As Long
Declare Function QA_GetFormattedLine Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal rs3 As String, ByVal vi4 As Long, ByVal rs5 As String, ByVal vi6 As Long, rl7 As Long) As Long
Declare Function QA_GetExampleCount Lib "QAUPIED.DLL" (ByVal vi1 As Long, ri2 As Long) As Long
Declare Function QA_FormatExample Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal rs3 As String, ByVal vi4 As Long, ri5 As Long, rl6 As Long) As Long
Declare Function QA_GetLayoutCount Lib "QAUPIED.DLL" (ByVal vi1 As Long, ri2 As Long) As Long
Declare Function QA_GetLayout Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal rs3 As String, ByVal vi4 As Long, ByVal rs5 As String, ByVal vi6 As Long) As Long
Declare Function QA_GetActiveLayout Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal rs2 As String, ByVal vi3 As Long) As Long
Declare Function QA_SetActiveLayout Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vs2 As String) As Long
Declare Function QA_GetDataCount Lib "QAUPIED.DLL" (ByVal vi1 As Long, ri2 As Long) As Long
Declare Function QA_GetData Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal rs3 As String, ByVal vi4 As Long, ByVal rs5 As String, ByVal vi6 As Long) As Long
Declare Function QA_GetDataDetail Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal vi3 As Long, rl4 As Long, ByVal rs5 As String, ByVal vi6 As Long) As Long
Declare Function QA_GetLicensingCount Lib "QAUPIED.DLL" (ByVal vi1 As Long, ri2 As Long, rl3 As Long) As Long
Declare Function QA_GetLicensingDetail Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal vi3 As Long, rl4 As Long, ByVal rs5 As String, ByVal vi6 As Long) As Long
Declare Function QA_GetActiveData Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal rs2 As String, ByVal vi3 As Long) As Long
Declare Function QA_SetActiveData Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vs2 As String) As Long
Declare Function QA_GenerateSystemInfo Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ri3 As Long) As Long
Declare Function QA_GetSystemInfo Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal rs3 As String, ByVal vi4 As Long) As Long
Declare Function QA_ErrorMessage Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal rs2 As String, ByVal vi3 As Long) As Long
Declare Function QA_IsFlagSet Lib "QAUPIED.DLL" (ByVal vl1 As Long, ByVal vl2 As Long) As Long

' NPG Fault 13364
Private arrHistory(10) As String
Public aFullAddress(2, 4) As String

Public Sub modQAShowMappedFields(TableID As Long, FieldName As String, PostCode As String, frmForm As Form)

On Error GoTo QAShowMappedFieldsError

  Dim rs As ADODB.Recordset     'Recordset containing mapped fields (columnIDs)
  Dim sSQL As String            'source of recordset
  Dim fIndividual As Boolean    'individual or merged address fields

  'Let the user know something is happening
  Screen.MousePointer = vbHourglass
  
  If giQAddressEnabled = QADDRESS_PRO5 Then
    ' Quick Address Pro v5+ so use new screen...
    Load frmQAProMain
   
    'Set the source of the recordset
    sSQL = "SELECT * from asrsyscolumns WHERE tableid = " & frmForm.TableID & _
    " AND columnname = '" & FieldName & "'"
  
    'Load recordset
    Set rs = datGeneral.GetRecords(sSQL)
  
    'Go through each field (different ones depending on the value of fIndividual)
    'If there is a valid column mapped then set the tag property of the relevant text
    'box on the Afd form, otherwise, disable the checkbox and the text field on the Afd form.
    If Not rs.BOF And Not rs.EOF Then
      
      fIndividual = rs.Fields("QAindividual")
      
      If Not fIndividual Then
        If rs.Fields("QAaddress") <> 0 Then
          frmAFDFields.txtMergedAddress.Tag = rs.Fields("QAaddress")
        Else
          frmAFDFields.txtMergedAddress.Tag = 0
        End If
      Else
        aFullAddress(2, 0) = IIf(rs.Fields("QAproperty") <> 0, rs.Fields("QAproperty"), 0)  ' tag for address 1
        aFullAddress(2, 1) = IIf(rs.Fields("QAstreet") <> 0, rs.Fields("QAstreet"), 0)  ' tag for address 2
        aFullAddress(2, 2) = IIf(rs.Fields("QAtown") <> 0, rs.Fields("QAtown"), 0)  ' tag for town
        aFullAddress(2, 3) = IIf(rs.Fields("QAcounty") <> 0, rs.Fields("QAcounty"), 0)  ' tag for county
      End If
      
      'Clear recordset reference
      Set rs = Nothing
    Else
      'Here is no data is found in the recordset...should never happen, but just incase
      Set rs = Nothing
      Exit Sub
    End If
   
    'Call the Afd routines. If they fail, exit sub, if not, show the Afd form
    If frmQAProMain.InitialiseQA(PostCode, fIndividual, frmForm, FieldName) = False Then
      Screen.MousePointer = vbDefault
      Exit Sub
    End If
    
    frmQAProMain.txtInput = Trim(UCase(PostCode))
     
    'Return mousepointer to normal
    Screen.MousePointer = vbDefault
     
    'Show the Afd form
    frmQAProMain.Show vbModal
            
    Unload frmQAProMain
      
  Else
    'NPG - all other address types
    'Load the Afd form
    Load frmAFDFields
    
    'Set the source of the recordset
    sSQL = "SELECT * from asrsyscolumns WHERE tableid = " & frmForm.TableID & _
    " AND columnname = '" & FieldName & "'"
  
    'Load recordset
    Set rs = datGeneral.GetRecords(sSQL)
  
    'Go through each field (different ones depending on the value of fIndividual)
    'If there is a valid column mapped then set the tag property of the relevant text
    'box on the Afd form, otherwise, disable the checkbox and the text field on the Afd form.
    If Not rs.BOF And Not rs.EOF Then
      fIndividual = rs.Fields("QAindividual")
      
      ' Disable Names & Numbers Functionality (AFD Only)
      frmAFDFields.txtMergedForename.Tag = 0
      frmAFDFields.chkMergedForename.Value = False
      frmAFDFields.chkMergedForename.Enabled = False
      frmAFDFields.txtMergedForename.Enabled = False
      frmAFDFields.txtMergedForename.BackColor = &H8000000F
  
      frmAFDFields.txtMergedInitials.Tag = 0
      frmAFDFields.chkMergedInitials.Value = False
      frmAFDFields.chkMergedInitials.Enabled = False
      frmAFDFields.txtMergedInitials.Enabled = False
      frmAFDFields.txtMergedInitials.BackColor = &H8000000F
      
      frmAFDFields.txtMergedSurname.Tag = 0
      frmAFDFields.chkMergedSurname.Value = False
      frmAFDFields.chkMergedSurname.Enabled = False
      frmAFDFields.txtMergedSurname.Enabled = False
      frmAFDFields.txtMergedSurname.BackColor = &H8000000F
  
      frmAFDFields.txtMergedTelephone.Tag = 0
      frmAFDFields.chkMergedTelephone.Value = False
      frmAFDFields.chkMergedTelephone.Enabled = False
      frmAFDFields.txtMergedTelephone.Enabled = False
      frmAFDFields.txtMergedTelephone.BackColor = &H8000000F
      
      frmAFDFields.txtForename.Tag = 0
      frmAFDFields.chkForename.Value = False
      frmAFDFields.chkForename.Enabled = False
      frmAFDFields.txtForename.Enabled = False
      frmAFDFields.txtForename.BackColor = &H8000000F
  
      frmAFDFields.txtInitials.Tag = 0
      frmAFDFields.chkInitials.Value = False
      frmAFDFields.chkInitials.Enabled = False
      frmAFDFields.txtInitials.Enabled = False
      frmAFDFields.txtInitials.BackColor = &H8000000F
  
      frmAFDFields.txtSurname.Tag = 0
      frmAFDFields.chkSurname.Value = False
      frmAFDFields.chkSurname.Enabled = False
      frmAFDFields.txtSurname.Enabled = False
      frmAFDFields.txtSurname.BackColor = &H8000000F
  
      frmAFDFields.txtTelephone.Tag = 0
      frmAFDFields.chkTelephone.Value = False
      frmAFDFields.chkTelephone.Enabled = False
      frmAFDFields.txtTelephone.Enabled = False
      frmAFDFields.txtTelephone.BackColor = &H8000000F
      
      If Not fIndividual Then
        
        If rs.Fields("QAaddress") <> 0 Then
          frmAFDFields.txtMergedAddress.Tag = rs.Fields("QAaddress")
        Else
          frmAFDFields.txtMergedAddress.Tag = 0
          frmAFDFields.chkMergedAddress.Value = False
          frmAFDFields.chkMergedAddress.Enabled = False
          frmAFDFields.txtMergedAddress.Enabled = False
          frmAFDFields.txtMergedAddress.BackColor = &H8000000F
        End If
        
      Else
        
        If rs.Fields("QAproperty") <> 0 Then
          frmAFDFields.txtProperty.Tag = rs.Fields("QAproperty")
        Else
          frmAFDFields.txtProperty.Tag = 0
          frmAFDFields.chkProperty.Value = False
          frmAFDFields.chkProperty.Enabled = False
          frmAFDFields.txtProperty.Enabled = False
          frmAFDFields.txtProperty.BackColor = &H8000000F
  
        End If
        
        If rs.Fields("QAstreet") <> 0 Then
          frmAFDFields.txtStreet.Tag = rs.Fields("QAstreet")
        Else
          frmAFDFields.txtStreet.Tag = 0
          frmAFDFields.chkStreet.Value = False
          frmAFDFields.chkStreet.Enabled = False
          frmAFDFields.txtStreet.Enabled = False
          frmAFDFields.txtStreet.BackColor = &H8000000F
  
        End If
        
        If rs.Fields("QAlocality") <> 0 Then
          frmAFDFields.txtLocality.Tag = rs.Fields("QAlocality")
        Else
          frmAFDFields.txtLocality.Tag = 0
          frmAFDFields.chkLocality.Value = False
          frmAFDFields.chkLocality.Enabled = False
          frmAFDFields.txtLocality.Enabled = False
          frmAFDFields.txtLocality.BackColor = &H8000000F
  
        End If
        
        If rs.Fields("QAtown") <> 0 Then
          frmAFDFields.txtTown.Tag = rs.Fields("QAtown")
        Else
          frmAFDFields.txtTown.Tag = 0
          frmAFDFields.chkTown.Value = False
          frmAFDFields.chkTown.Enabled = False
          frmAFDFields.txtTown.Enabled = False
          frmAFDFields.txtTown.BackColor = &H8000000F
  
        End If
        
        If rs.Fields("QAcounty") <> 0 Then
          frmAFDFields.txtCounty.Tag = rs.Fields("QAcounty")
        Else
          frmAFDFields.txtCounty.Tag = 0
          frmAFDFields.chkCounty.Value = False
          frmAFDFields.chkCounty.Enabled = False
          frmAFDFields.txtCounty.Enabled = False
          frmAFDFields.txtCounty.BackColor = &H8000000F
  
        End If
            
      End If
      'Clear recordset reference
      Set rs = Nothing
    Else
      'Here is no data is found in the recordset...should never happen, but just incase
      Set rs = Nothing
      Exit Sub
    End If
  
    'Call the Afd routines. If they fail, exit sub, if not, show the Afd form
    If frmAFDFields.InitialiseQA(PostCode, fIndividual, frmForm, FieldName) = False Then
      Screen.MousePointer = vbDefault
      Exit Sub
    End If
    
    'Return mousepointer to normal
    Screen.MousePointer = vbDefault
    
    'Show the Afd form
    frmAFDFields.Show vbModal
  End If

QAShowMappedFieldsResume:
Exit Sub

QAShowMappedFieldsError:

COAMsgBox "Error : " & Err.Number & " - " & Err.Description & " - modQAddressSpecifics.modQAShowMappedFields", vbOKOnly, "Error"
Resume QAShowMappedFieldsResume

End Sub

Public Function QAddressGetPostcodes(pstrPostcode As String, ByRef pobjQAPostcodes() As DataMgr.PostCode) As Long
  
  Dim iResult As Integer
  Dim iCount As Integer
  Dim lNoOfItems As Long
  Dim lGetItemInfoReturn As Long
  Dim lResult As Long
  Dim lOpen As Long
  Dim lngDataCount As Long
  Dim lSearchReturn As Long
  
  ReDim pobjQAPostcodes(0)
  
  'Call Quick Address GetPostcode routines
  Select Case giQAddressEnabled
    
    ' Quick Address Rapid
    Case QADDRESS_RAPID
      
      lOpen = QARapid_Open("", "")
      QARapid_Search (MakeQAddressString(pstrPostcode))
      lNoOfItems = QARapid_Count
      
      If lNoOfItems > 0 Then
        For iCount = 0 To lNoOfItems - 1
          GetQARapidPostcodes iCount, pobjQAPostcodes
        Next iCount
      End If
      
      QARapid_Close
    
    ' Quick Address Version 3.X
    Case QADDRESS_PRO3
      QAPro_Close
      QAInitialise (1)
      lOpen = QAPro_Open("", "")
      QAPro_Search (MakeQAddressString(pstrPostcode))
      'Count the number of addresses in the list
      lNoOfItems = QAPro_Count
      If lNoOfItems > 0 Then
        If lNoOfItems = 1 Then
          lNoOfItems = QAPro_StepIn(iCount)
        End If
        For iCount = 0 To lNoOfItems - 1
          GetQAProSubItems iCount, pobjQAPostcodes
        Next iCount
      End If
      
      QAPro_Close
    
    ' Quick Address Worldwide
    Case QADDRESS_WORLDWIDE
    
    
    ' Quick Address Version 4.X
    Case QADDRESS_PRO4

  End Select

  QAddressGetPostcodes = UBound(pobjQAPostcodes)

End Function

Private Sub GetQAProSubItems(ByVal lListIndex As Long, ByRef pobjQAPostcodes() As DataMgr.PostCode)

  Dim rsBuffer As String * 200
  Dim vlBufLen As Long
  Dim lListItemReturn As Long
  Dim lItemNumber As Long
  Dim lGetItemInfoReturn As Long
  Dim lResult As Long
  Dim bStepInto As Boolean
  Dim lNoOfItems As Long
  Dim lFormatCountReturn As Long
  Dim iStepInto As Integer
  Dim iCount As Integer
  Dim iSeekNumber As Integer
  Dim strBuilding As String
  Dim strThoroughFare As String
  
  ' Find out if the picklist can be stepped into
  'lGetItemInfoReturn = QAPro_GetItemInfo(lListIndex, qapro_STEPINFO, lResult)
  'bStepInto = (lGetItemInfoReturn < 0)
   '     lListIndex = 0

'  lNoOfItems = QAPro_StepIn(lListIndex)
  lNoOfItems = QAPro_StepIn(lListIndex)
 
  If lNoOfItems = qaerr_CANTSTEP Then
    lNoOfItems = 1
  End If

  vlBufLen = 200
  
  For lItemNumber = 0 To (lNoOfItems - 1)
  
  'Debug.Assert lItemNumber <> 7
  
    iStepInto = QAPro_StepIn(lItemNumber)
    
    If Not (iStepInto = qaerr_CANTSTEP) Then
  
      ' Step into the sub items
      For iCount = 0 To iStepInto
        GetQAProSubItems iCount, pobjQAPostcodes
      Next iCount
      QAPro_StepOut
      
    Else
       
      ' Add this postcode to the array
      iSeekNumber = IIf(lNoOfItems > 1, lItemNumber, lListIndex)
      
      lListItemReturn = QAPro_ListItem(iSeekNumber, rsBuffer, 0, vlBufLen)

      
      If QAPro_FormatCount > 0 Then

        'Stuff the postcode into the postcode object
        ReDim Preserve pobjQAPostcodes(UBound(pobjQAPostcodes) + 1)
        
        ' Stuff that doesn't apply to Quick Address
        pobjQAPostcodes(UBound(pobjQAPostcodes)).FirstName = ""
        pobjQAPostcodes(UBound(pobjQAPostcodes)).Initial2 = ""
        pobjQAPostcodes(UBound(pobjQAPostcodes)).Organisation = ""
        pobjQAPostcodes(UBound(pobjQAPostcodes)).Phone = ""
        pobjQAPostcodes(UBound(pobjQAPostcodes)).Surname = ""

        ' Postcode
        rsBuffer = Space(200)
        QAPro_AddrLine iSeekNumber, 11, "", rsBuffer, 200
        pobjQAPostcodes(UBound(pobjQAPostcodes)).PostCode = Trim(UnMakeCString(rsBuffer))

        ' House Number
        rsBuffer = Space(200)
        QAPro_AddrLine iSeekNumber, 4, "", rsBuffer, 200
        pobjQAPostcodes(UBound(pobjQAPostcodes)).HouseNo = Trim(UnMakeCString(rsBuffer))

        ' Building
        rsBuffer = Space(200)
        QAPro_AddrLine iSeekNumber, 3, "", rsBuffer, 200
        strBuilding = Trim(UnMakeCString(rsBuffer))

        ' Sub Premesis
        rsBuffer = Space(200)
        QAPro_AddrLine iSeekNumber, 2, "", rsBuffer, 200
        pobjQAPostcodes(UBound(pobjQAPostcodes)).Building = Trim(UnMakeCString(rsBuffer)) & " " & strBuilding
  
        ' ThoroughFare
        rsBuffer = Space(200)
        QAPro_AddrLine iSeekNumber, 5, "", rsBuffer, 200
        strThoroughFare = Trim(UnMakeCString(rsBuffer))
  
        ' Street
        rsBuffer = Space(200)
        QAPro_AddrLine iSeekNumber, 6, "", rsBuffer, 200
        pobjQAPostcodes(UBound(pobjQAPostcodes)).Street = strThoroughFare & IIf(Len(strThoroughFare) > 0, ", ", "") & Trim(UnMakeCString(rsBuffer))
        
        ' Locality
        rsBuffer = Space(200)
        QAPro_AddrLine iSeekNumber, 8, "", rsBuffer, 200
        pobjQAPostcodes(UBound(pobjQAPostcodes)).Locality = Trim(UnMakeCString(rsBuffer))
        
        ' Town
        rsBuffer = Space(200)
        QAPro_AddrLine iSeekNumber, 9, "", rsBuffer, 200
        pobjQAPostcodes(UBound(pobjQAPostcodes)).Town = Trim(UnMakeCString(rsBuffer))
        
        ' County
        rsBuffer = Space(200)
        QAPro_AddrLine iSeekNumber, 10, "", rsBuffer, 200
        pobjQAPostcodes(UBound(pobjQAPostcodes)).County = Trim(UnMakeCString(rsBuffer))
      
      End If
    End If
     
  Next lItemNumber

  ' Step out of these sub items
  If lNoOfItems > 1 Then
    QAPro_StepOut
  End If

End Sub


Private Sub GetQARapidPostcodes(ByVal lListIndex As Long, ByRef pobjQAPostcodes() As DataMgr.PostCode)

    Dim rsBuffer As String * 200
    Dim vlBufLen As Long
    Dim lListItemReturn As Long
    Dim lItemNumber As Long
    Dim lGetItemInfoReturn As Long
    Dim lResult As Long
    Dim strThoroughFare As String
    
        vlBufLen = 200
        
        'For lItemNumber = 0 To (lListIndex - 1)
        
            'get each address for the picklist
            rsBuffer = Space(200)
            lListItemReturn = QARapid_ListItem(lListIndex, rsBuffer, 0, vlBufLen)
            
            If lListItemReturn >= 0 Then

        'Stuff the postcode into the postcode object
        ReDim Preserve pobjQAPostcodes(UBound(pobjQAPostcodes) + 1)
        
        ' Stuff that doesn't apply to Quick Address
        pobjQAPostcodes(UBound(pobjQAPostcodes)).FirstName = ""
        pobjQAPostcodes(UBound(pobjQAPostcodes)).Initial2 = ""
        pobjQAPostcodes(UBound(pobjQAPostcodes)).Organisation = ""
        pobjQAPostcodes(UBound(pobjQAPostcodes)).Phone = ""
        pobjQAPostcodes(UBound(pobjQAPostcodes)).Surname = ""

        ' Postcode
        rsBuffer = Space(200)
        QARapid_AddrLine lListIndex, 11, rsBuffer, 200
        pobjQAPostcodes(UBound(pobjQAPostcodes)).PostCode = Trim(UnMakeCString(rsBuffer))

        ' House Number
        pobjQAPostcodes(UBound(pobjQAPostcodes)).HouseNo = ""

        ' Building
        pobjQAPostcodes(UBound(pobjQAPostcodes)).Building = ""
  
        ' ThoroughFare
        rsBuffer = Space(200)
        QARapid_AddrLine lListIndex, 5, rsBuffer, 200
        strThoroughFare = Trim(UnMakeCString(rsBuffer))
  
        ' Street
        rsBuffer = Space(200)
        QARapid_AddrLine lListIndex, 6, rsBuffer, 200
        pobjQAPostcodes(UBound(pobjQAPostcodes)).Street = strThoroughFare & IIf(Len(strThoroughFare) > 0, ", ", "") & Trim(UnMakeCString(rsBuffer))
        
        ' Locality
        rsBuffer = Space(200)
        QARapid_AddrLine lListIndex, 8, rsBuffer, 200
        pobjQAPostcodes(UBound(pobjQAPostcodes)).Locality = Trim(UnMakeCString(rsBuffer))
        
        ' Town
        rsBuffer = Space(200)
        QARapid_AddrLine lListIndex, 9, rsBuffer, 200
        pobjQAPostcodes(UBound(pobjQAPostcodes)).Town = Trim(UnMakeCString(rsBuffer))
        
        ' County
        rsBuffer = Space(200)
        QARapid_AddrLine lListIndex, 10, rsBuffer, 200
        pobjQAPostcodes(UBound(pobjQAPostcodes)).County = Trim(UnMakeCString(rsBuffer))
                      
                
            End If
            
        'Next
End Sub

' FUNCTIONS FROM QUICK ADDRESS - slightly amended

' Make a string suitable for use by the DLLs
Public Function MakeQAddressString(ByVal sArg As String) As String
  
  MakeQAddressString = sArg & Chr(0)

End Function

' Removes the null character from the end of the string returned from the DLLs
Private Function UnMakeCString(ByVal sArg As String) As String
    
    Dim iNulIndex As Integer

    iNulIndex = InStr(sArg, Chr$(0))

    If iNulIndex > 0 Then
        UnMakeCString = Mid$(sArg, 1, iNulIndex - 1)
    Else
        UnMakeCString = sArg
    End If

End Function

' Bring up a dialog box based on the Error Number passed to it
Private Function ErrorMessage(ByVal lErrorNo As Long) As String
    
    Dim rsBuffer As String * 100
    Dim vlBufLen As Long
    
    vlBufLen = 100
    
    If lErrorNo < 0 Then
        Call QAErrorMessage(lErrorNo, rsBuffer, vlBufLen)
        ErrorMessage = "Error: " & Str$(lErrorNo) & "." & Chr(10) & Chr(13) & UnMakeCString(rsBuffer)
    End If

End Function

' Return the textual description of the error based on the Error Number passed to it
Private Function WarningMessage(ByVal lErrorNo As Long) As String
    
    Dim rsBuffer As String * 100
    Dim vlBufLen As Long
    
    vlBufLen = 100
    
    If lErrorNo < 0 Then
        Call QAErrorMessage(lErrorNo, rsBuffer, vlBufLen)
        WarningMessage = rsBuffer
    End If

End Function

