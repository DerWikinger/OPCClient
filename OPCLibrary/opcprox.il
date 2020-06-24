
//  Microsoft (R) .NET Framework IL Disassembler.  Version 4.8.3928.0
//  Copyright (c) Microsoft Corporation.  All rights reserved.



// Metadata version: v2.0.50727
.assembly extern mscorlib
{
  .publickeytoken = (B7 7A 5C 56 19 34 E0 89 )                         // .z\V.4..
  .ver 4:0:0:0
}
.assembly opcprox
{
  .custom instance void [mscorlib]System.Runtime.InteropServices.ImportedFromTypeLibAttribute::.ctor(string) = ( 01 00 05 4F 50 43 44 41 00 00 )                   // ...OPCDA..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeLibVersionAttribute::.ctor(int32,
                                                                                                int32) = ( 01 00 03 00 00 00 00 00 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 33 42 35 34 30 42 35 31 2D 30 33 37 38   // ..$3B540B51-0378
                                                                                                  2D 34 35 35 31 2D 41 44 43 43 2D 45 41 39 42 31   // -4551-ADCC-EA9B1
                                                                                                  30 34 33 30 32 42 46 00 00 )                      // 04302BF..
  .hash algorithm 0x00008004
  .ver 3:0:0:0
}
.module opcprox.dll
// MVID: {AD31B456-9556-444E-8699-9682C2802764}
.imagebase 0x00400000
.file alignment 0x00000200
.stackreserve 0x00100000
.subsystem 0x0003       // WINDOWS_CUI
.corflags 0x00000001    //  ILONLY
// Image base: 0x05350000


// =============== CLASS MEMBERS DECLARATION ===================

.class interface public abstract auto ansi import opcprox.CATID_OPCDAServer10
{
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 36 33 44 35 46 34 33 30 2D 43 46 45 34   // ..$63D5F430-CFE4
                                                                                                  2D 31 31 44 31 2D 42 32 43 38 2D 30 30 36 30 30   // -11D1-B2C8-00600
                                                                                                  38 33 42 41 31 46 42 00 00 )                      // 83BA1FB..
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(valuetype [mscorlib]System.Runtime.InteropServices.ComInterfaceType) = ( 01 00 01 00 00 00 00 00 ) 
} // end of class opcprox.CATID_OPCDAServer10

.class interface public abstract auto ansi import opcprox.CATID_OPCDAServer20
{
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 36 33 44 35 46 34 33 32 2D 43 46 45 34   // ..$63D5F432-CFE4
                                                                                                  2D 31 31 44 31 2D 42 32 43 38 2D 30 30 36 30 30   // -11D1-B2C8-00600
                                                                                                  38 33 42 41 31 46 42 00 00 )                      // 83BA1FB..
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(valuetype [mscorlib]System.Runtime.InteropServices.ComInterfaceType) = ( 01 00 01 00 00 00 00 00 ) 
} // end of class opcprox.CATID_OPCDAServer20

.class interface public abstract auto ansi import opcprox.CATID_OPCDAServer30
{
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(valuetype [mscorlib]System.Runtime.InteropServices.ComInterfaceType) = ( 01 00 01 00 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 43 43 36 30 33 36 34 32 2D 36 36 44 37   // ..$CC603642-66D7
                                                                                                  2D 34 38 46 31 2D 42 36 39 41 2D 42 36 32 35 45   // -48F1-B69A-B625E
                                                                                                  37 33 36 35 32 44 37 00 00 )                      // 73652D7..
} // end of class opcprox.CATID_OPCDAServer30

.class interface public abstract auto ansi import opcprox.CATID_XMLDAServer10
{
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 33 30 39 38 45 44 41 34 2D 41 30 30 36   // ..$3098EDA4-A006
                                                                                                  2D 34 38 42 32 2D 41 32 37 46 2D 32 34 37 34 35   // -48B2-A27F-24745
                                                                                                  33 39 35 39 34 30 38 00 00 )                      // 3959408..
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(valuetype [mscorlib]System.Runtime.InteropServices.ComInterfaceType) = ( 01 00 01 00 00 00 00 00 ) 
} // end of class opcprox.CATID_XMLDAServer10

.class public abstract auto ansi sealed opcprox.OPCDA_Constants
       extends [mscorlib]System.Object
{
  .field public static literal  marshal( lpwstr) string OPC_CATEGORY_DESCRIPTION_DA10 = "OPC Data Access Servers Version 1.0"
  .field public static literal  marshal( lpwstr) string OPC_CATEGORY_DESCRIPTION_DA20 = "OPC Data Access Servers Version 2.0"
  .field public static literal  marshal( lpwstr) string OPC_CATEGORY_DESCRIPTION_DA30 = "OPC Data Access Servers Version 3.0"
  .field public static literal  marshal( lpwstr) string OPC_CATEGORY_DESCRIPTION_XMLDA10 = "OPC XML Data Access Servers Version 1.0"
  .field public static literal uint32 OPC_READABLE
  .field public static literal uint32 OPC_WRITEABLE
  .field public static literal uint32 OPC_BROWSE_HASCHILDREN
  .field public static literal uint32 OPC_BROWSE_ISITEM
  .field public static literal  marshal( lpwstr) string OPC_TYPE_SYSTEM_OPCBINARY = "OPCBinary"
  .field public static literal  marshal( lpwstr) string OPC_TYPE_SYSTEM_XMLSCHEMA = "XMLSchema"
  .field public static literal  marshal( lpwstr) string OPC_CONSISTENCY_WINDOW_UNKNOWN = "Unknown"
  .field public static literal  marshal( lpwstr) string OPC_CONSISTENCY_WINDOW_NOT_CONSISTENT = "Not Consistent"
  .field public static literal  marshal( lpwstr) string OPC_WRITE_BEHAVIOR_BEST_EFFORT = "Best Effort"
  .field public static literal  marshal( lpwstr) string OPC_WRITE_BEHAVIOR_ALL_OR_NOTHING = "All or Nothing"
} // end of class opcprox.OPCDA_Constants

.class public abstract auto ansi sealed opcprox.OPCDA_Qualities
       extends [mscorlib]System.Object
{
  .field public static literal uint16 OPC_QUALITY_MASK = uint16(0x00C0)
  .field public static literal uint16 OPC_STATUS_MASK = uint16(0x00FC)
  .field public static literal uint16 OPC_LIMIT_MASK = uint16(0x0003)
  .field public static literal uint16 OPC_QUALITY_BAD = uint16(0x0000)
  .field public static literal uint16 OPC_QUALITY_UNCERTAIN = uint16(0x0040)
  .field public static literal uint16 OPC_QUALITY_GOOD = uint16(0x00C0)
  .field public static literal uint16 OPC_QUALITY_CONFIG_ERROR = uint16(0x0004)
  .field public static literal uint16 OPC_QUALITY_NOT_CONNECTED = uint16(0x0008)
  .field public static literal uint16 OPC_QUALITY_DEVICE_FAILURE = uint16(0x000C)
  .field public static literal uint16 OPC_QUALITY_SENSOR_FAILURE = uint16(0x0010)
  .field public static literal uint16 OPC_QUALITY_LAST_KNOWN = uint16(0x0014)
  .field public static literal uint16 OPC_QUALITY_COMM_FAILURE = uint16(0x0018)
  .field public static literal uint16 OPC_QUALITY_OUT_OF_SERVICE = uint16(0x001C)
  .field public static literal uint16 OPC_QUALITY_WAITING_FOR_INITIAL_DATA = uint16(0x0020)
  .field public static literal uint16 OPC_QUALITY_LAST_USABLE = uint16(0x0044)
  .field public static literal uint16 OPC_QUALITY_SENSOR_CAL = uint16(0x0050)
  .field public static literal uint16 OPC_QUALITY_EGU_EXCEEDED = uint16(0x0054)
  .field public static literal uint16 OPC_QUALITY_SUB_NORMAL = uint16(0x0058)
  .field public static literal uint16 OPC_QUALITY_LOCAL_OVERRIDE = uint16(0x00D8)
  .field public static literal uint16 OPC_LIMIT_OK = uint16(0x0000)
  .field public static literal uint16 OPC_LIMIT_LOW = uint16(0x0001)
  .field public static literal uint16 OPC_LIMIT_HIGH = uint16(0x0002)
  .field public static literal uint16 OPC_LIMIT_CONST = uint16(0x0003)
} // end of class opcprox.OPCDA_Qualities

.class public abstract auto ansi sealed opcprox.OPCDA_Properties
       extends [mscorlib]System.Object
{
  .field public static literal uint32 OPC_PROPERTY_DATATYPE
  .field public static literal uint32 OPC_PROPERTY_VALUE
  .field public static literal uint32 OPC_PROPERTY_QUALITY
  .field public static literal uint32 OPC_PROPERTY_TIMESTAMP
  .field public static literal uint32 OPC_PROPERTY_ACCESS_RIGHTS
  .field public static literal uint32 OPC_PROPERTY_SCAN_RATE
  .field public static literal uint32 OPC_PROPERTY_EU_TYPE
  .field public static literal uint32 OPC_PROPERTY_EU_INFO
  .field public static literal uint32 OPC_PROPERTY_EU_UNITS
  .field public static literal uint32 OPC_PROPERTY_DESCRIPTION
  .field public static literal uint32 OPC_PROPERTY_HIGH_EU
  .field public static literal uint32 OPC_PROPERTY_LOW_EU
  .field public static literal uint32 OPC_PROPERTY_HIGH_IR
  .field public static literal uint32 OPC_PROPERTY_LOW_IR
  .field public static literal uint32 OPC_PROPERTY_CLOSE_LABEL
  .field public static literal uint32 OPC_PROPERTY_OPEN_LABEL
  .field public static literal uint32 OPC_PROPERTY_TIMEZONE
  .field public static literal uint32 OPC_PROPERTY_CONDITION_STATUS
  .field public static literal uint32 OPC_PROPERTY_ALARM_QUICK_HELP
  .field public static literal uint32 OPC_PROPERTY_ALARM_AREA_LIST
  .field public static literal uint32 OPC_PROPERTY_PRIMARY_ALARM_AREA
  .field public static literal uint32 OPC_PROPERTY_CONDITION_LOGIC
  .field public static literal uint32 OPC_PROPERTY_LIMIT_EXCEEDED
  .field public static literal uint32 OPC_PROPERTY_DEADBAND
  .field public static literal uint32 OPC_PROPERTY_HIHI_LIMIT
  .field public static literal uint32 OPC_PROPERTY_HI_LIMIT
  .field public static literal uint32 OPC_PROPERTY_LO_LIMIT
  .field public static literal uint32 OPC_PROPERTY_LOLO_LIMIT
  .field public static literal uint32 OPC_PROPERTY_CHANGE_RATE_LIMIT
  .field public static literal uint32 OPC_PROPERTY_DEVIATION_LIMIT
  .field public static literal uint32 OPC_PROPERTY_SOUND_FILE
  .field public static literal uint32 OPC_PROPERTY_TYPE_SYSTEM_ID
  .field public static literal uint32 OPC_PROPERTY_DICTIONARY_ID
  .field public static literal uint32 OPC_PROPERTY_TYPE_ID
  .field public static literal uint32 OPC_PROPERTY_DICTIONARY
  .field public static literal uint32 OPC_PROPERTY_TYPE_DESCRIPTION
  .field public static literal uint32 OPC_PROPERTY_CONSISTENCY_WINDOW
  .field public static literal uint32 OPC_PROPERTY_WRITE_BEHAVIOR
  .field public static literal uint32 OPC_PROPERTY_UNCONVERTED_ITEM_ID
  .field public static literal uint32 OPC_PROPERTY_UNFILTERED_ITEM_ID
  .field public static literal uint32 OPC_PROPERTY_DATA_FILTER_VALUE
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_DATATYPE = "Item Canonical Data Type"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_VALUE = "Item Value"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_QUALITY = "Item Quality"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_TIMESTAMP = "Item Timestamp"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_ACCESS_RIGHTS = "Item Access Rights"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_SCAN_RATE = "Server Scan Rate"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_EU_TYPE = "Item EU Type"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_EU_INFO = "Item EU Info"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_EU_UNITS = "EU Units"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_DESCRIPTION = "Item Description"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_HIGH_EU = "High EU"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_LOW_EU = "Low EU"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_HIGH_IR = "High Instrument Range"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_LOW_IR = "Low Instrument Range"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_CLOSE_LABEL = "Contact Close Label"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_OPEN_LABEL = "Contact Open Label"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_TIMEZONE = "Item Timezone"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_CONDITION_STATUS = "Condition Status"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_ALARM_QUICK_HELP = "Alarm Quick Help"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_ALARM_AREA_LIST = "Alarm Area List"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_PRIMARY_ALARM_AREA = "Primary Alarm Area"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_CONDITION_LOGIC = "Condition Logic"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_LIMIT_EXCEEDED = "Limit Exceeded"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_DEADBAND = "Deadband"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_HIHI_LIMIT = "HiHi Limit"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_HI_LIMIT = "Hi Limit"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_LO_LIMIT = "Lo Limit"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_LOLO_LIMIT = "LoLo Limit"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_CHANGE_RATE_LIMIT = "Rate of Change Limit"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_DEVIATION_LIMIT = "Deviation Limit"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_SOUND_FILE = "Sound File"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_TYPE_SYSTEM_ID = "Type System ID"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_DICTIONARY_ID = "Dictionary ID"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_TYPE_ID = "Type ID"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_DICTIONARY = "Dictionary"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_TYPE_DESCRIPTION = "Type Description"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_CONSISTENCY_WINDOW = "Consistency Window"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_WRITE_BEHAVIOR = "Write Behavior"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_UNCONVERTED_ITEM_ID = "Unconverted Item ID"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_UNFILTERED_ITEM_ID = "Unfiltered Item ID"
  .field public static literal  marshal( lpwstr) string OPC_PROPERTY_DESC_DATA_FILTER_VALUE = "Data Filter Value"
} // end of class opcprox.OPCDA_Properties

.class interface public abstract auto ansi import opcprox.IOPCServer
{
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(valuetype [mscorlib]System.Runtime.InteropServices.ComInterfaceType) = ( 01 00 01 00 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 33 39 43 31 33 41 34 44 2D 30 31 31 45   // ..$39C13A4D-011E
                                                                                                  2D 31 31 44 30 2D 39 36 37 35 2D 30 30 32 30 41   // -11D0-9675-0020A
                                                                                                  46 44 38 41 44 42 33 00 00 )                      // FD8ADB3..
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComConversionLossAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public hidebysig newslot abstract virtual 
          instance void  AddGroup([in] string  marshal( lpwstr) szName,
                                  [in] int32 bActive,
                                  [in] uint32 dwRequestedUpdateRate,
                                  [in] uint32 hClientGroup,
                                  [in] int32& pTimeBias,
                                  [in] float32& pPercentDeadband,
                                  [in] uint32 dwLCID,
                                  [out] uint32& phServerGroup,
                                  [out] uint32& pRevisedUpdateRate,
                                  [in] valuetype [mscorlib]System.Guid& riid,
                                  [out] object&  marshal( iunknown ) ppUnk) runtime managed internalcall
  {
  } // end of method IOPCServer::AddGroup

  .method public hidebysig newslot abstract virtual 
          instance void  GetErrorString([in] int32  marshal( error) dwError,
                                        [in] uint32 dwLocale,
                                        [out] string&  marshal( lpwstr) ppString) runtime managed internalcall
  {
  } // end of method IOPCServer::GetErrorString

  .method public hidebysig newslot abstract virtual 
          instance void  GetGroupByName([in] string  marshal( lpwstr) szName,
                                        [in] valuetype [mscorlib]System.Guid& riid,
                                        [out] object&  marshal( iunknown ) ppUnk) runtime managed internalcall
  {
  } // end of method IOPCServer::GetGroupByName

  .method public hidebysig newslot abstract virtual 
          instance void  GetStatus([out] native int ppServerStatus) runtime managed internalcall
  {
  } // end of method IOPCServer::GetStatus

  .method public hidebysig newslot abstract virtual 
          instance void  RemoveGroup([in] uint32 hServerGroup,
                                     [in] int32 bForce) runtime managed internalcall
  {
  } // end of method IOPCServer::RemoveGroup

  .method public hidebysig newslot abstract virtual 
          instance void  CreateGroupEnumerator([in] valuetype opcprox.tagOPCENUMSCOPE dwScope,
                                               [in] valuetype [mscorlib]System.Guid& riid,
                                               [out] object&  marshal( iunknown ) ppUnk) runtime managed internalcall
  {
  } // end of method IOPCServer::CreateGroupEnumerator

} // end of class opcprox.IOPCServer

.class public auto ansi sealed opcprox.tagOPCENUMSCOPE
       extends [mscorlib]System.Enum
{
  .field public specialname rtspecialname int32 value__
  .field public static literal valuetype opcprox.tagOPCENUMSCOPE OPC_ENUM_PRIVATE_CONNECTIONS = int32(0x00000001)
  .field public static literal valuetype opcprox.tagOPCENUMSCOPE OPC_ENUM_PUBLIC_CONNECTIONS = int32(0x00000002)
  .field public static literal valuetype opcprox.tagOPCENUMSCOPE OPC_ENUM_ALL_CONNECTIONS = int32(0x00000003)
  .field public static literal valuetype opcprox.tagOPCENUMSCOPE OPC_ENUM_PRIVATE = int32(0x00000004)
  .field public static literal valuetype opcprox.tagOPCENUMSCOPE OPC_ENUM_PUBLIC = int32(0x00000005)
  .field public static literal valuetype opcprox.tagOPCENUMSCOPE OPC_ENUM_ALL = int32(0x00000006)
} // end of class opcprox.tagOPCENUMSCOPE

.class public sequential ansi sealed beforefieldinit opcprox.tagOPCSERVERSTATUS
       extends [mscorlib]System.ValueType
{
  .pack 4
  .size 0
  .field public valuetype opcprox._FILETIME ftStartTime
  .field public valuetype opcprox._FILETIME ftCurrentTime
  .field public valuetype opcprox._FILETIME ftLastUpdateTime
  .field public valuetype opcprox.tagOPCSERVERSTATE dwServerState
  .field public uint32 dwGroupCount
  .field public uint32 dwBandWidth
  .field public uint16 wMajorVersion
  .field public uint16 wMinorVersion
  .field public uint16 wBuildNumber
  .field public uint16 wReserved
  .field public  marshal( lpwstr) string szVendorInfo
} // end of class opcprox.tagOPCSERVERSTATUS

.class public sequential ansi sealed beforefieldinit opcprox._FILETIME
       extends [mscorlib]System.ValueType
{
  .pack 4
  .size 0
  .field public uint32 dwLowDateTime
  .field public uint32 dwHighDateTime
} // end of class opcprox._FILETIME

.class public auto ansi sealed opcprox.tagOPCSERVERSTATE
       extends [mscorlib]System.Enum
{
  .field public specialname rtspecialname int32 value__
  .field public static literal valuetype opcprox.tagOPCSERVERSTATE OPC_STATUS_RUNNING = int32(0x00000001)
  .field public static literal valuetype opcprox.tagOPCSERVERSTATE OPC_STATUS_FAILED = int32(0x00000002)
  .field public static literal valuetype opcprox.tagOPCSERVERSTATE OPC_STATUS_NOCONFIG = int32(0x00000003)
  .field public static literal valuetype opcprox.tagOPCSERVERSTATE OPC_STATUS_SUSPENDED = int32(0x00000004)
  .field public static literal valuetype opcprox.tagOPCSERVERSTATE OPC_STATUS_TEST = int32(0x00000005)
  .field public static literal valuetype opcprox.tagOPCSERVERSTATE OPC_STATUS_COMM_FAULT = int32(0x00000006)
} // end of class opcprox.tagOPCSERVERSTATE

.class interface public abstract auto ansi import opcprox.IOPCServerPublicGroups
{
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(valuetype [mscorlib]System.Runtime.InteropServices.ComInterfaceType) = ( 01 00 01 00 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 33 39 43 31 33 41 34 45 2D 30 31 31 45   // ..$39C13A4E-011E
                                                                                                  2D 31 31 44 30 2D 39 36 37 35 2D 30 30 32 30 41   // -11D0-9675-0020A
                                                                                                  46 44 38 41 44 42 33 00 00 )                      // FD8ADB3..
  .method public hidebysig newslot abstract virtual 
          instance void  GetPublicGroupByName([in] string  marshal( lpwstr) szName,
                                              [in] valuetype [mscorlib]System.Guid& riid,
                                              [out] object&  marshal( iunknown ) ppUnk) runtime managed internalcall
  {
  } // end of method IOPCServerPublicGroups::GetPublicGroupByName

  .method public hidebysig newslot abstract virtual 
          instance void  RemovePublicGroup([in] uint32 hServerGroup,
                                           [in] int32 bForce) runtime managed internalcall
  {
  } // end of method IOPCServerPublicGroups::RemovePublicGroup

} // end of class opcprox.IOPCServerPublicGroups

.class interface public abstract auto ansi import opcprox.IOPCBrowseServerAddressSpace
{
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(valuetype [mscorlib]System.Runtime.InteropServices.ComInterfaceType) = ( 01 00 01 00 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 33 39 43 31 33 41 34 46 2D 30 31 31 45   // ..$39C13A4F-011E
                                                                                                  2D 31 31 44 30 2D 39 36 37 35 2D 30 30 32 30 41   // -11D0-9675-0020A
                                                                                                  46 44 38 41 44 42 33 00 00 )                      // FD8ADB3..
  .method public hidebysig newslot abstract virtual 
          instance void  QueryOrganization([out] valuetype opcprox.tagOPCNAMESPACETYPE& pNameSpaceType) runtime managed internalcall
  {
  } // end of method IOPCBrowseServerAddressSpace::QueryOrganization

  .method public hidebysig newslot abstract virtual 
          instance void  ChangeBrowsePosition([in] valuetype opcprox.tagOPCBROWSEDIRECTION dwBrowseDirection,
                                              [in] string  marshal( lpwstr) szString) runtime managed internalcall
  {
  } // end of method IOPCBrowseServerAddressSpace::ChangeBrowsePosition

  .method public hidebysig newslot abstract virtual 
          instance void  BrowseOPCItemIDs([in] valuetype opcprox.tagOPCBROWSETYPE dwBrowseFilterType,
                                          [in] string  marshal( lpwstr) szFilterCriteria,
                                          [in] uint16 vtDataTypeFilter,
                                          [in] uint32 dwAccessRightsFilter,
                                          [out] class opcprox.IEnumString&  marshal( interface ) ppIEnumString) runtime managed internalcall
  {
  } // end of method IOPCBrowseServerAddressSpace::BrowseOPCItemIDs

  .method public hidebysig newslot abstract virtual 
          instance void  GetItemID([in] string  marshal( lpwstr) szItemDataID,
                                   [out] string&  marshal( lpwstr) szItemID) runtime managed internalcall
  {
  } // end of method IOPCBrowseServerAddressSpace::GetItemID

  .method public hidebysig newslot abstract virtual 
          instance void  BrowseAccessPaths([in] string  marshal( lpwstr) szItemID,
                                           [out] class opcprox.IEnumString&  marshal( interface ) ppIEnumString) runtime managed internalcall
  {
  } // end of method IOPCBrowseServerAddressSpace::BrowseAccessPaths

} // end of class opcprox.IOPCBrowseServerAddressSpace

.class public auto ansi sealed opcprox.tagOPCNAMESPACETYPE
       extends [mscorlib]System.Enum
{
  .field public specialname rtspecialname int32 value__
  .field public static literal valuetype opcprox.tagOPCNAMESPACETYPE OPC_NS_HIERARCHIAL = int32(0x00000001)
  .field public static literal valuetype opcprox.tagOPCNAMESPACETYPE OPC_NS_FLAT = int32(0x00000002)
} // end of class opcprox.tagOPCNAMESPACETYPE

.class public auto ansi sealed opcprox.tagOPCBROWSEDIRECTION
       extends [mscorlib]System.Enum
{
  .field public specialname rtspecialname int32 value__
  .field public static literal valuetype opcprox.tagOPCBROWSEDIRECTION OPC_BROWSE_UP = int32(0x00000001)
  .field public static literal valuetype opcprox.tagOPCBROWSEDIRECTION OPC_BROWSE_DOWN = int32(0x00000002)
  .field public static literal valuetype opcprox.tagOPCBROWSEDIRECTION OPC_BROWSE_TO = int32(0x00000003)
} // end of class opcprox.tagOPCBROWSEDIRECTION

.class public auto ansi sealed opcprox.tagOPCBROWSETYPE
       extends [mscorlib]System.Enum
{
  .field public specialname rtspecialname int32 value__
  .field public static literal valuetype opcprox.tagOPCBROWSETYPE OPC_BRANCH = int32(0x00000001)
  .field public static literal valuetype opcprox.tagOPCBROWSETYPE OPC_LEAF = int32(0x00000002)
  .field public static literal valuetype opcprox.tagOPCBROWSETYPE OPC_FLAT = int32(0x00000003)
} // end of class opcprox.tagOPCBROWSETYPE

.class interface public abstract auto ansi import opcprox.IEnumString
{
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 30 30 31 30 31 2D 30 30 30 30   // ..$00000101-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(valuetype [mscorlib]System.Runtime.InteropServices.ComInterfaceType) = ( 01 00 01 00 00 00 00 00 ) 
  .method public hidebysig newslot abstract virtual 
          instance void  RemoteNext([in] uint32 celt,
                                    [out] string&  marshal( lpwstr) rgelt,
                                    [out] uint32& pceltFetched) runtime managed internalcall
  {
  } // end of method IEnumString::RemoteNext

  .method public hidebysig newslot abstract virtual 
          instance void  Skip([in] uint32 celt) runtime managed internalcall
  {
  } // end of method IEnumString::Skip

  .method public hidebysig newslot abstract virtual 
          instance void  Reset() runtime managed internalcall
  {
  } // end of method IEnumString::Reset

  .method public hidebysig newslot abstract virtual 
          instance void  Clone([out] class opcprox.IEnumString&  marshal( interface ) ppenum) runtime managed internalcall
  {
  } // end of method IEnumString::Clone

} // end of class opcprox.IEnumString

.class interface public abstract auto ansi import opcprox.IOPCGroupStateMgt
{
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(valuetype [mscorlib]System.Runtime.InteropServices.ComInterfaceType) = ( 01 00 01 00 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 33 39 43 31 33 41 35 30 2D 30 31 31 45   // ..$39C13A50-011E
                                                                                                  2D 31 31 44 30 2D 39 36 37 35 2D 30 30 32 30 41   // -11D0-9675-0020A
                                                                                                  46 44 38 41 44 42 33 00 00 )                      // FD8ADB3..
  .method public hidebysig newslot abstract virtual 
          instance void  GetState([out] uint32& pUpdateRate,
                                  [out] int32& pActive,
                                  [out] string&  marshal( lpwstr) ppName,
                                  [out] int32& pTimeBias,
                                  [out] float32& pPercentDeadband,
                                  [out] uint32& pLCID,
                                  [out] uint32& phClientGroup,
                                  [out] uint32& phServerGroup) runtime managed internalcall
  {
  } // end of method IOPCGroupStateMgt::GetState

  .method public hidebysig newslot abstract virtual 
          instance void  SetState([in] uint32& pRequestedUpdateRate,
                                  [out] uint32& pRevisedUpdateRate,
                                  [in] int32& pActive,
                                  [in] int32& pTimeBias,
                                  [in] float32& pPercentDeadband,
                                  [in] uint32& pLCID,
                                  [in] uint32& phClientGroup) runtime managed internalcall
  {
  } // end of method IOPCGroupStateMgt::SetState

  .method public hidebysig newslot abstract virtual 
          instance void  SetName([in] string  marshal( lpwstr) szName) runtime managed internalcall
  {
  } // end of method IOPCGroupStateMgt::SetName

  .method public hidebysig newslot abstract virtual 
          instance void  CloneGroup([in] string  marshal( lpwstr) szName,
                                    [in] valuetype [mscorlib]System.Guid& riid,
                                    [out] object&  marshal( iunknown ) ppUnk) runtime managed internalcall
  {
  } // end of method IOPCGroupStateMgt::CloneGroup

} // end of class opcprox.IOPCGroupStateMgt

.class interface public abstract auto ansi import opcprox.IOPCPublicGroupStateMgt
{
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(valuetype [mscorlib]System.Runtime.InteropServices.ComInterfaceType) = ( 01 00 01 00 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 33 39 43 31 33 41 35 31 2D 30 31 31 45   // ..$39C13A51-011E
                                                                                                  2D 31 31 44 30 2D 39 36 37 35 2D 30 30 32 30 41   // -11D0-9675-0020A
                                                                                                  46 44 38 41 44 42 33 00 00 )                      // FD8ADB3..
  .method public hidebysig newslot abstract virtual 
          instance void  GetState([out] int32& pPublic) runtime managed internalcall
  {
  } // end of method IOPCPublicGroupStateMgt::GetState

  .method public hidebysig newslot abstract virtual 
          instance void  MoveToPublic() runtime managed internalcall
  {
  } // end of method IOPCPublicGroupStateMgt::MoveToPublic

} // end of class opcprox.IOPCPublicGroupStateMgt

.class interface public abstract auto ansi import opcprox.IOPCSyncIO
{
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 33 39 43 31 33 41 35 32 2D 30 31 31 45   // ..$39C13A52-011E
                                                                                                  2D 31 31 44 30 2D 39 36 37 35 2D 30 30 32 30 41   // -11D0-9675-0020A
                                                                                                  46 44 38 41 44 42 33 00 00 )                      // FD8ADB3..
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(valuetype [mscorlib]System.Runtime.InteropServices.ComInterfaceType) = ( 01 00 01 00 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComConversionLossAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public hidebysig newslot abstract virtual 
          instance void  Read([in] valuetype opcprox.tagOPCDATASOURCE dwSource,
                              [in] uint32 dwCount,
                              [in] uint32& phServer,
                              [out] native int& ppItemValues,
                              [out] native int& ppErrors) runtime managed internalcall
  {
  } // end of method IOPCSyncIO::Read

  .method public hidebysig newslot abstract virtual 
          instance void  Write([in] uint32 dwCount,
                               [in] uint32& phServer,
                               [in] object&  marshal( struct) pItemValues,
                               [out] native int& ppErrors) runtime managed internalcall
  {
  } // end of method IOPCSyncIO::Write

} // end of class opcprox.IOPCSyncIO

.class public auto ansi sealed opcprox.tagOPCDATASOURCE
       extends [mscorlib]System.Enum
{
  .field public specialname rtspecialname int32 value__
  .field public static literal valuetype opcprox.tagOPCDATASOURCE OPC_DS_CACHE = int32(0x00000001)
  .field public static literal valuetype opcprox.tagOPCDATASOURCE OPC_DS_DEVICE = int32(0x00000002)
} // end of class opcprox.tagOPCDATASOURCE

.class public sequential ansi sealed beforefieldinit opcprox.tagOPCITEMSTATE
       extends [mscorlib]System.ValueType
{
  .pack 8
  .size 0
  .field public uint32 hClient
  .field public valuetype opcprox._FILETIME ftTimeStamp
  .field public uint16 wQuality
  .field public uint16 wReserved
  .field public  marshal( struct) object vDataValue
} // end of class opcprox.tagOPCITEMSTATE

.class interface public abstract auto ansi import opcprox.IOPCAsyncIO
{
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 33 39 43 31 33 41 35 33 2D 30 31 31 45   // ..$39C13A53-011E
                                                                                                  2D 31 31 44 30 2D 39 36 37 35 2D 30 30 32 30 41   // -11D0-9675-0020A
                                                                                                  46 44 38 41 44 42 33 00 00 )                      // FD8ADB3..
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(valuetype [mscorlib]System.Runtime.InteropServices.ComInterfaceType) = ( 01 00 01 00 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComConversionLossAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public hidebysig newslot abstract virtual 
          instance void  Read([in] uint32 dwConnection,
                              [in] valuetype opcprox.tagOPCDATASOURCE dwSource,
                              [in] uint32 dwCount,
                              [in] uint32& phServer,
                              [out] uint32& pTransactionID,
                              [out] native int ppErrors) runtime managed internalcall
  {
  } // end of method IOPCAsyncIO::Read

  .method public hidebysig newslot abstract virtual 
          instance void  Write([in] uint32 dwConnection,
                               [in] uint32 dwCount,
                               [in] uint32& phServer,
                               [in] object&  marshal( struct) pItemValues,
                               [out] uint32& pTransactionID,
                               [out] native int ppErrors) runtime managed internalcall
  {
  } // end of method IOPCAsyncIO::Write

  .method public hidebysig newslot abstract virtual 
          instance void  Refresh([in] uint32 dwConnection,
                                 [in] valuetype opcprox.tagOPCDATASOURCE dwSource,
                                 [out] uint32& pTransactionID) runtime managed internalcall
  {
  } // end of method IOPCAsyncIO::Refresh

  .method public hidebysig newslot abstract virtual 
          instance void  Cancel([in] uint32 dwTransactionID) runtime managed internalcall
  {
  } // end of method IOPCAsyncIO::Cancel

} // end of class opcprox.IOPCAsyncIO

.class interface public abstract auto ansi import opcprox.IOPCDataCallback
{
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(valuetype [mscorlib]System.Runtime.InteropServices.ComInterfaceType) = ( 01 00 01 00 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 33 39 43 31 33 41 37 30 2D 30 31 31 45   // ..$39C13A70-011E
                                                                                                  2D 31 31 44 30 2D 39 36 37 35 2D 30 30 32 30 41   // -11D0-9675-0020A
                                                                                                  46 44 38 41 44 42 33 00 00 )                      // FD8ADB3..
  .method public hidebysig newslot abstract virtual 
          instance void  OnDataChange([in] uint32 dwTransid,
                                      [in] uint32 hGroup,
                                      [in] int32  marshal( error) hrMasterquality,
                                      [in] int32  marshal( error) hrMastererror,
                                      [in] uint32 dwCount,
                                      [in] uint32& phClientItems,
                                      [in] object&  marshal( struct) pvValues,
                                      [in] uint16& pwQualities,
                                      [in] valuetype opcprox._FILETIME& pftTimeStamps,
                                      [in] int32&  marshal( error) pErrors) runtime managed internalcall
  {
  } // end of method IOPCDataCallback::OnDataChange

  .method public hidebysig newslot abstract virtual 
          instance void  OnReadComplete([in] uint32 dwTransid,
                                        [in] uint32 hGroup,
                                        [in] int32  marshal( error) hrMasterquality,
                                        [in] int32  marshal( error) hrMastererror,
                                        [in] uint32 dwCount,
                                        [in] uint32& phClientItems,
                                        [in] object&  marshal( struct) pvValues,
                                        [in] uint16& pwQualities,
                                        [in] valuetype opcprox._FILETIME& pftTimeStamps,
                                        [in] int32&  marshal( error) pErrors) runtime managed internalcall
  {
  } // end of method IOPCDataCallback::OnReadComplete

  .method public hidebysig newslot abstract virtual 
          instance void  OnWriteComplete([in] uint32 dwTransid,
                                         [in] uint32 hGroup,
                                         [in] int32  marshal( error) hrMastererr,
                                         [in] uint32 dwCount,
                                         [in] uint32& pClienthandles,
                                         [in] int32&  marshal( error) pErrors) runtime managed internalcall
  {
  } // end of method IOPCDataCallback::OnWriteComplete

  .method public hidebysig newslot abstract virtual 
          instance void  OnCancelComplete([in] uint32 dwTransid,
                                          [in] uint32 hGroup) runtime managed internalcall
  {
  } // end of method IOPCDataCallback::OnCancelComplete

} // end of class opcprox.IOPCDataCallback

.class interface public abstract auto ansi import opcprox.IOPCItemMgt
{
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComConversionLossAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(valuetype [mscorlib]System.Runtime.InteropServices.ComInterfaceType) = ( 01 00 01 00 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 33 39 43 31 33 41 35 34 2D 30 31 31 45   // ..$39C13A54-011E
                                                                                                  2D 31 31 44 30 2D 39 36 37 35 2D 30 30 32 30 41   // -11D0-9675-0020A
                                                                                                  46 44 38 41 44 42 33 00 00 )                      // FD8ADB3..
  .method public hidebysig newslot abstract virtual 
          instance void  AddItems([in] uint32 dwCount,
                                  [in] native int pItemArray,
                                  [out] native int& ppAddResults,
                                  [out] native int& ppErrors) runtime managed internalcall
  {
  } // end of method IOPCItemMgt::AddItems

  .method public hidebysig newslot abstract virtual 
          instance void  ValidateItems([in] uint32 dwCount,
                                       [in] valuetype opcprox.tagOPCITEMDEF& pItemArray,
                                       [in] int32 bBlobUpdate,
                                       [out] native int ppValidationResults,
                                       [out] native int ppErrors) runtime managed internalcall
  {
  } // end of method IOPCItemMgt::ValidateItems

  .method public hidebysig newslot abstract virtual 
          instance void  RemoveItems([in] uint32 dwCount,
                                     [in] uint32& phServer,
                                     [out] native int& ppErrors) runtime managed internalcall
  {
  } // end of method IOPCItemMgt::RemoveItems

  .method public hidebysig newslot abstract virtual 
          instance void  SetActiveState([in] uint32 dwCount,
                                        [in] uint32& phServer,
                                        [in] int32 bActive,
                                        [out] native int ppErrors) runtime managed internalcall
  {
  } // end of method IOPCItemMgt::SetActiveState

  .method public hidebysig newslot abstract virtual 
          instance void  SetClientHandles([in] uint32 dwCount,
                                          [in] uint32& phServer,
                                          [in] uint32& phClient,
                                          [out] native int ppErrors) runtime managed internalcall
  {
  } // end of method IOPCItemMgt::SetClientHandles

  .method public hidebysig newslot abstract virtual 
          instance void  SetDatatypes([in] uint32 dwCount,
                                      [in] uint32& phServer,
                                      [in] uint16& pRequestedDatatypes,
                                      [out] native int ppErrors) runtime managed internalcall
  {
  } // end of method IOPCItemMgt::SetDatatypes

  .method public hidebysig newslot abstract virtual 
          instance void  CreateEnumerator([in] valuetype [mscorlib]System.Guid& riid,
                                          [out] object&  marshal( iunknown ) ppUnk) runtime managed internalcall
  {
  } // end of method IOPCItemMgt::CreateEnumerator

} // end of class opcprox.IOPCItemMgt

.class public sequential ansi sealed beforefieldinit opcprox.tagOPCITEMDEF
       extends [mscorlib]System.ValueType
{
  .pack 4
  .size 0
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComConversionLossAttribute::.ctor() = ( 01 00 00 00 ) 
  .field public  marshal( lpwstr) string szAccessPath
  .field public  marshal( lpwstr) string szItemID
  .field public int32 bActive
  .field public uint32 hClient
  .field public uint32 dwBlobSize
  .field public native int pBlob
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComConversionLossAttribute::.ctor() = ( 01 00 00 00 ) 
  .field public uint16 vtRequestedDataType
  .field public uint16 wReserved
} // end of class opcprox.tagOPCITEMDEF

.class public sequential ansi sealed beforefieldinit opcprox.tagOPCITEMRESULT
       extends [mscorlib]System.ValueType
{
  .pack 4
  .size 0
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComConversionLossAttribute::.ctor() = ( 01 00 00 00 ) 
  .field public uint32 hServer
  .field public uint16 vtCanonicalDataType
  .field public uint16 wReserved
  .field public uint32 dwAccessRights
  .field public uint32 dwBlobSize
  .field public native int pBlob
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComConversionLossAttribute::.ctor() = ( 01 00 00 00 ) 
} // end of class opcprox.tagOPCITEMRESULT

.class interface public abstract auto ansi import opcprox.IEnumOPCItemAttributes
{
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 33 39 43 31 33 41 35 35 2D 30 31 31 45   // ..$39C13A55-011E
                                                                                                  2D 31 31 44 30 2D 39 36 37 35 2D 30 30 32 30 41   // -11D0-9675-0020A
                                                                                                  46 44 38 41 44 42 33 00 00 )                      // FD8ADB3..
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(valuetype [mscorlib]System.Runtime.InteropServices.ComInterfaceType) = ( 01 00 01 00 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComConversionLossAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public hidebysig newslot abstract virtual 
          instance void  Next([in] uint32 celt,
                              [out] native int ppItemArray,
                              [out] uint32& pceltFetched) runtime managed internalcall
  {
  } // end of method IEnumOPCItemAttributes::Next

  .method public hidebysig newslot abstract virtual 
          instance void  Skip([in] uint32 celt) runtime managed internalcall
  {
  } // end of method IEnumOPCItemAttributes::Skip

  .method public hidebysig newslot abstract virtual 
          instance void  Reset() runtime managed internalcall
  {
  } // end of method IEnumOPCItemAttributes::Reset

  .method public hidebysig newslot abstract virtual 
          instance void  Clone([out] class opcprox.IEnumOPCItemAttributes&  marshal( interface ) ppEnumItemAttributes) runtime managed internalcall
  {
  } // end of method IEnumOPCItemAttributes::Clone

} // end of class opcprox.IEnumOPCItemAttributes

.class public sequential ansi sealed beforefieldinit opcprox.tagOPCITEMATTRIBUTES
       extends [mscorlib]System.ValueType
{
  .pack 8
  .size 0
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComConversionLossAttribute::.ctor() = ( 01 00 00 00 ) 
  .field public  marshal( lpwstr) string szAccessPath
  .field public  marshal( lpwstr) string szItemID
  .field public int32 bActive
  .field public uint32 hClient
  .field public uint32 hServer
  .field public uint32 dwAccessRights
  .field public uint32 dwBlobSize
  .field public native int pBlob
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComConversionLossAttribute::.ctor() = ( 01 00 00 00 ) 
  .field public uint16 vtRequestedDataType
  .field public uint16 vtCanonicalDataType
  .field public valuetype opcprox.tagOPCEUTYPE dwEUType
  .field public  marshal( struct) object vEUInfo
} // end of class opcprox.tagOPCITEMATTRIBUTES

.class public auto ansi sealed opcprox.tagOPCEUTYPE
       extends [mscorlib]System.Enum
{
  .field public specialname rtspecialname int32 value__
  .field public static literal valuetype opcprox.tagOPCEUTYPE OPC_NOENUM = int32(0x00000000)
  .field public static literal valuetype opcprox.tagOPCEUTYPE OPC_ANALOG = int32(0x00000001)
  .field public static literal valuetype opcprox.tagOPCEUTYPE OPC_ENUMERATED = int32(0x00000002)
} // end of class opcprox.tagOPCEUTYPE

.class interface public abstract auto ansi import opcprox.IOPCAsyncIO2
{
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(valuetype [mscorlib]System.Runtime.InteropServices.ComInterfaceType) = ( 01 00 01 00 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComConversionLossAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 33 39 43 31 33 41 37 31 2D 30 31 31 45   // ..$39C13A71-011E
                                                                                                  2D 31 31 44 30 2D 39 36 37 35 2D 30 30 32 30 41   // -11D0-9675-0020A
                                                                                                  46 44 38 41 44 42 33 00 00 )                      // FD8ADB3..
  .method public hidebysig newslot abstract virtual 
          instance void  Read([in] uint32 dwCount,
                              [in] uint32& phServer,
                              [in] uint32 dwTransactionID,
                              [out] uint32& pdwCancelID,
                              [out] native int& ppErrors) runtime managed internalcall
  {
  } // end of method IOPCAsyncIO2::Read

  .method public hidebysig newslot abstract virtual 
          instance void  Write([in] uint32 dwCount,
                               [in] uint32& phServer,
                               [in] object&  marshal( struct) pItemValues,
                               [in] uint32 dwTransactionID,
                               [out] uint32& pdwCancelID,
                               [out] native int ppErrors) runtime managed internalcall
  {
  } // end of method IOPCAsyncIO2::Write

  .method public hidebysig newslot abstract virtual 
          instance void  Refresh2([in] valuetype opcprox.tagOPCDATASOURCE dwSource,
                                  [in] uint32 dwTransactionID,
                                  [out] uint32& pdwCancelID) runtime managed internalcall
  {
  } // end of method IOPCAsyncIO2::Refresh2

  .method public hidebysig newslot abstract virtual 
          instance void  Cancel2([in] uint32 dwCancelID) runtime managed internalcall
  {
  } // end of method IOPCAsyncIO2::Cancel2

  .method public hidebysig newslot abstract virtual 
          instance void  SetEnable([in] int32 bEnable) runtime managed internalcall
  {
  } // end of method IOPCAsyncIO2::SetEnable

  .method public hidebysig newslot abstract virtual 
          instance void  GetEnable([out] int32& pbEnable) runtime managed internalcall
  {
  } // end of method IOPCAsyncIO2::GetEnable

} // end of class opcprox.IOPCAsyncIO2

.class interface public abstract auto ansi import opcprox.IOPCItemProperties
{
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(valuetype [mscorlib]System.Runtime.InteropServices.ComInterfaceType) = ( 01 00 01 00 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 33 39 43 31 33 41 37 32 2D 30 31 31 45   // ..$39C13A72-011E
                                                                                                  2D 31 31 44 30 2D 39 36 37 35 2D 30 30 32 30 41   // -11D0-9675-0020A
                                                                                                  46 44 38 41 44 42 33 00 00 )                      // FD8ADB3..
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComConversionLossAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public hidebysig newslot abstract virtual 
          instance void  QueryAvailableProperties([in] string  marshal( lpwstr) szItemID,
                                                  [out] uint32& pdwCount,
                                                  [out] native int& ppPropertyIDs,
                                                  [out] native int& ppDescriptions,
                                                  [out] native int& ppvtDataTypes) runtime managed internalcall
  {
  } // end of method IOPCItemProperties::QueryAvailableProperties

  .method public hidebysig newslot abstract virtual 
          instance void  GetItemProperties([in] string  marshal( lpwstr) szItemID,
                                           [in] uint32 dwCount,
                                           [in] native int pdwPropertyIDs,
                                           [out] native int& ppvData,
                                           [out] native int& ppErrors) runtime managed internalcall
  {
  } // end of method IOPCItemProperties::GetItemProperties

  .method public hidebysig newslot abstract virtual 
          instance void  LookupItemIDs([in] string  marshal( lpwstr) szItemID,
                                       [in] uint32 dwCount,
                                       [in] uint32& pdwPropertyIDs,
                                       [out] native int& ppszNewItemIDs,
                                       [out] native int& ppErrors) runtime managed internalcall
  {
  } // end of method IOPCItemProperties::LookupItemIDs

} // end of class opcprox.IOPCItemProperties

.class interface public abstract auto ansi import opcprox.IOPCItemDeadbandMgt
{
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(valuetype [mscorlib]System.Runtime.InteropServices.ComInterfaceType) = ( 01 00 01 00 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 35 39 34 36 44 41 39 33 2D 38 42 33 39   // ..$5946DA93-8B39
                                                                                                  2D 34 45 43 38 2D 41 42 33 44 2D 41 41 37 33 44   // -4EC8-AB3D-AA73D
                                                                                                  46 35 42 43 38 36 46 00 00 )                      // F5BC86F..
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComConversionLossAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public hidebysig newslot abstract virtual 
          instance void  SetItemDeadband([in] uint32 dwCount,
                                         [in] uint32& phServer,
                                         [in] float32& pPercentDeadband,
                                         [out] native int ppErrors) runtime managed internalcall
  {
  } // end of method IOPCItemDeadbandMgt::SetItemDeadband

  .method public hidebysig newslot abstract virtual 
          instance void  GetItemDeadband([in] uint32 dwCount,
                                         [in] uint32& phServer,
                                         [out] native int ppPercentDeadband,
                                         [out] native int ppErrors) runtime managed internalcall
  {
  } // end of method IOPCItemDeadbandMgt::GetItemDeadband

  .method public hidebysig newslot abstract virtual 
          instance void  ClearItemDeadband([in] uint32 dwCount,
                                           [in] uint32& phServer,
                                           [out] native int ppErrors) runtime managed internalcall
  {
  } // end of method IOPCItemDeadbandMgt::ClearItemDeadband

} // end of class opcprox.IOPCItemDeadbandMgt

.class interface public abstract auto ansi import opcprox.IOPCItemSamplingMgt
{
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(valuetype [mscorlib]System.Runtime.InteropServices.ComInterfaceType) = ( 01 00 01 00 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComConversionLossAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 33 45 32 32 44 33 31 33 2D 46 30 38 42   // ..$3E22D313-F08B
                                                                                                  2D 34 31 41 35 2D 38 36 43 38 2D 39 35 45 39 35   // -41A5-86C8-95E95
                                                                                                  43 42 34 39 46 46 43 00 00 )                      // CB49FFC..
  .method public hidebysig newslot abstract virtual 
          instance void  SetItemSamplingRate([in] uint32 dwCount,
                                             [in] uint32& phServer,
                                             [in] uint32& pdwRequestedSamplingRate,
                                             [out] native int ppdwRevisedSamplingRate,
                                             [out] native int ppErrors) runtime managed internalcall
  {
  } // end of method IOPCItemSamplingMgt::SetItemSamplingRate

  .method public hidebysig newslot abstract virtual 
          instance void  GetItemSamplingRate([in] uint32 dwCount,
                                             [in] uint32& phServer,
                                             [out] native int ppdwSamplingRate,
                                             [out] native int ppErrors) runtime managed internalcall
  {
  } // end of method IOPCItemSamplingMgt::GetItemSamplingRate

  .method public hidebysig newslot abstract virtual 
          instance void  ClearItemSamplingRate([in] uint32 dwCount,
                                               [in] uint32& phServer,
                                               [out] native int ppErrors) runtime managed internalcall
  {
  } // end of method IOPCItemSamplingMgt::ClearItemSamplingRate

  .method public hidebysig newslot abstract virtual 
          instance void  SetItemBufferEnable([in] uint32 dwCount,
                                             [in] uint32& phServer,
                                             [in] int32& pbEnable,
                                             [out] native int ppErrors) runtime managed internalcall
  {
  } // end of method IOPCItemSamplingMgt::SetItemBufferEnable

  .method public hidebysig newslot abstract virtual 
          instance void  GetItemBufferEnable([in] uint32 dwCount,
                                             [in] uint32& phServer,
                                             [out] native int ppbEnable,
                                             [out] native int ppErrors) runtime managed internalcall
  {
  } // end of method IOPCItemSamplingMgt::GetItemBufferEnable

} // end of class opcprox.IOPCItemSamplingMgt

.class interface public abstract auto ansi import opcprox.IOPCBrowse
{
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 33 39 32 32 37 30 30 34 2D 41 31 38 46   // ..$39227004-A18F
                                                                                                  2D 34 42 35 37 2D 38 42 30 41 2D 35 32 33 35 36   // -4B57-8B0A-52356
                                                                                                  37 30 46 34 34 36 38 00 00 )                      // 70F4468..
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComConversionLossAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(valuetype [mscorlib]System.Runtime.InteropServices.ComInterfaceType) = ( 01 00 01 00 00 00 00 00 ) 
  .method public hidebysig newslot abstract virtual 
          instance void  GetProperties([in] uint32 dwItemCount,
                                       [in] string&  marshal( lpwstr) pszItemIDs,
                                       [in] int32 bReturnPropertyValues,
                                       [in] uint32 dwPropertyCount,
                                       [in] uint32& pdwPropertyIDs,
                                       [out] native int ppItemProperties) runtime managed internalcall
  {
  } // end of method IOPCBrowse::GetProperties

  .method public hidebysig newslot abstract virtual 
          instance void  Browse([in] string  marshal( lpwstr) szItemID,
                                [in][out] string&  marshal( lpwstr) pszContinuationPoint,
                                [in] uint32 dwMaxElementsReturned,
                                [in] valuetype opcprox.tagOPCBROWSEFILTER dwBrowseFilter,
                                [in] string  marshal( lpwstr) szElementNameFilter,
                                [in] string  marshal( lpwstr) szVendorFilter,
                                [in] int32 bReturnAllProperties,
                                [in] int32 bReturnPropertyValues,
                                [in] uint32 dwPropertyCount,
                                [in] uint32& pdwPropertyIDs,
                                [out] int32& pbMoreElements,
                                [out] uint32& pdwCount,
                                [out] native int& ppBrowseElements) runtime managed internalcall
  {
  } // end of method IOPCBrowse::Browse

} // end of class opcprox.IOPCBrowse

.class public auto ansi sealed opcprox.tagOPCBROWSEFILTER
       extends [mscorlib]System.Enum
{
  .field public specialname rtspecialname int32 value__
  .field public static literal valuetype opcprox.tagOPCBROWSEFILTER OPC_BROWSE_FILTER_ALL = int32(0x00000001)
  .field public static literal valuetype opcprox.tagOPCBROWSEFILTER OPC_BROWSE_FILTER_BRANCHES = int32(0x00000002)
  .field public static literal valuetype opcprox.tagOPCBROWSEFILTER OPC_BROWSE_FILTER_ITEMS = int32(0x00000003)
} // end of class opcprox.tagOPCBROWSEFILTER

.class public sequential ansi sealed beforefieldinit opcprox.tagOPCITEMPROPERTIES
       extends [mscorlib]System.ValueType
{
  .pack 4
  .size 0
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComConversionLossAttribute::.ctor() = ( 01 00 00 00 ) 
  .field public  marshal( error) int32 hrErrorID
  .field public uint32 dwNumProperties
  .field public native int pItemProperties
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComConversionLossAttribute::.ctor() = ( 01 00 00 00 ) 
  .field public uint32 dwReserved
} // end of class opcprox.tagOPCITEMPROPERTIES

.class public sequential ansi sealed beforefieldinit opcprox.tagOPCBROWSEELEMENT
       extends [mscorlib]System.ValueType
{
  .pack 4
  .size 0
  .field public  marshal( lpwstr) string szName
  .field public  marshal( lpwstr) string szItemID
  .field public uint32 dwFlagValue
  .field public uint32 dwReserved
  .field public valuetype opcprox.tagOPCITEMPROPERTIES ItemProperties
} // end of class opcprox.tagOPCBROWSEELEMENT

.class public sequential ansi sealed beforefieldinit opcprox.tagOPCITEMPROPERTY
       extends [mscorlib]System.ValueType
{
  .pack 8
  .size 0
  .field public uint16 vtDataType
  .field public uint16 wReserved
  .field public uint32 dwPropertyID
  .field public  marshal( lpwstr) string szItemID
  .field public  marshal( lpwstr) string szDescription
  .field public  marshal( struct) object vValue
  .field public  marshal( error) int32 hrErrorID
  .field public uint32 dwReserved
} // end of class opcprox.tagOPCITEMPROPERTY

.class interface public abstract auto ansi import opcprox.IOPCItemIO
{
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComConversionLossAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 38 35 43 30 42 34 32 37 2D 32 38 39 33   // ..$85C0B427-2893
                                                                                                  2D 34 43 42 43 2D 42 44 37 38 2D 45 35 46 43 35   // -4CBC-BD78-E5FC5
                                                                                                  31 34 36 46 30 38 46 00 00 )                      // 146F08F..
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(valuetype [mscorlib]System.Runtime.InteropServices.ComInterfaceType) = ( 01 00 01 00 00 00 00 00 ) 
  .method public hidebysig newslot abstract virtual 
          instance void  Read([in] uint32 dwCount,
                              [in] string&  marshal( lpwstr) pszItemIDs,
                              [in] uint32& pdwMaxAge,
                              [out] native int ppvValues,
                              [out] native int ppwQualities,
                              [out] native int ppftTimeStamps,
                              [out] native int ppErrors) runtime managed internalcall
  {
  } // end of method IOPCItemIO::Read

  .method public hidebysig newslot abstract virtual 
          instance void  WriteVQT([in] uint32 dwCount,
                                  [in] string&  marshal( lpwstr) pszItemIDs,
                                  [in] valuetype opcprox.tagOPCITEMVQT& pItemVQT,
                                  [out] native int ppErrors) runtime managed internalcall
  {
  } // end of method IOPCItemIO::WriteVQT

} // end of class opcprox.IOPCItemIO

.class public sequential ansi sealed beforefieldinit opcprox.tagOPCITEMVQT
       extends [mscorlib]System.ValueType
{
  .pack 8
  .size 0
  .field public  marshal( struct) object vDataValue
  .field public int32 bQualitySpecified
  .field public uint16 wQuality
  .field public uint16 wReserved
  .field public int32 bTimeStampSpecified
  .field public uint32 dwReserved
  .field public valuetype opcprox._FILETIME ftTimeStamp
} // end of class opcprox.tagOPCITEMVQT

.class interface public abstract auto ansi import opcprox.IOPCSyncIO2
       implements opcprox.IOPCSyncIO
{
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(valuetype [mscorlib]System.Runtime.InteropServices.ComInterfaceType) = ( 01 00 01 00 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 37 33 30 46 35 46 30 46 2D 35 35 42 31   // ..$730F5F0F-55B1
                                                                                                  2D 34 43 38 31 2D 39 45 31 38 2D 46 46 38 41 30   // -4C81-9E18-FF8A0
                                                                                                  39 30 34 45 31 46 41 00 00 )                      // 904E1FA..
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComConversionLossAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public hidebysig newslot abstract virtual 
          instance void  Read([in] valuetype opcprox.tagOPCDATASOURCE dwSource,
                              [in] uint32 dwCount,
                              [in] uint32& phServer,
                              [out] native int ppItemValues,
                              [out] native int ppErrors) runtime managed internalcall
  {
  } // end of method IOPCSyncIO2::Read

  .method public hidebysig newslot abstract virtual 
          instance void  Write([in] uint32 dwCount,
                               [in] uint32& phServer,
                               [in] object&  marshal( struct) pItemValues,
                               [out] native int ppErrors) runtime managed internalcall
  {
  } // end of method IOPCSyncIO2::Write

  .method public hidebysig newslot abstract virtual 
          instance void  ReadMaxAge([in] uint32 dwCount,
                                    [in] uint32& phServer,
                                    [in] uint32& pdwMaxAge,
                                    [out] native int ppvValues,
                                    [out] native int ppwQualities,
                                    [out] native int ppftTimeStamps,
                                    [out] native int ppErrors) runtime managed internalcall
  {
  } // end of method IOPCSyncIO2::ReadMaxAge

  .method public hidebysig newslot abstract virtual 
          instance void  WriteVQT([in] uint32 dwCount,
                                  [in] uint32& phServer,
                                  [in] valuetype opcprox.tagOPCITEMVQT& pItemVQT,
                                  [out] native int ppErrors) runtime managed internalcall
  {
  } // end of method IOPCSyncIO2::WriteVQT

} // end of class opcprox.IOPCSyncIO2

.class interface public abstract auto ansi import opcprox.IOPCAsyncIO3
       implements opcprox.IOPCAsyncIO2
{
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 39 36 37 42 39 37 42 2D 33 36 45 46   // ..$0967B97B-36EF
                                                                                                  2D 34 32 33 45 2D 42 36 46 38 2D 36 42 46 46 31   // -423E-B6F8-6BFF1
                                                                                                  45 34 30 44 33 39 44 00 00 )                      // E40D39D..
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(valuetype [mscorlib]System.Runtime.InteropServices.ComInterfaceType) = ( 01 00 01 00 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComConversionLossAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public hidebysig newslot abstract virtual 
          instance void  Read([in] uint32 dwCount,
                              [in] uint32& phServer,
                              [in] uint32 dwTransactionID,
                              [out] uint32& pdwCancelID,
                              [out] native int ppErrors) runtime managed internalcall
  {
  } // end of method IOPCAsyncIO3::Read

  .method public hidebysig newslot abstract virtual 
          instance void  Write([in] uint32 dwCount,
                               [in] uint32& phServer,
                               [in] object&  marshal( struct) pItemValues,
                               [in] uint32 dwTransactionID,
                               [out] uint32& pdwCancelID,
                               [out] native int ppErrors) runtime managed internalcall
  {
  } // end of method IOPCAsyncIO3::Write

  .method public hidebysig newslot abstract virtual 
          instance void  Refresh2([in] valuetype opcprox.tagOPCDATASOURCE dwSource,
                                  [in] uint32 dwTransactionID,
                                  [out] uint32& pdwCancelID) runtime managed internalcall
  {
  } // end of method IOPCAsyncIO3::Refresh2

  .method public hidebysig newslot abstract virtual 
          instance void  Cancel2([in] uint32 dwCancelID) runtime managed internalcall
  {
  } // end of method IOPCAsyncIO3::Cancel2

  .method public hidebysig newslot abstract virtual 
          instance void  SetEnable([in] int32 bEnable) runtime managed internalcall
  {
  } // end of method IOPCAsyncIO3::SetEnable

  .method public hidebysig newslot abstract virtual 
          instance void  GetEnable([out] int32& pbEnable) runtime managed internalcall
  {
  } // end of method IOPCAsyncIO3::GetEnable

  .method public hidebysig newslot abstract virtual 
          instance void  ReadMaxAge([in] uint32 dwCount,
                                    [in] uint32& phServer,
                                    [in] uint32& pdwMaxAge,
                                    [in] uint32 dwTransactionID,
                                    [out] uint32& pdwCancelID,
                                    [out] native int ppErrors) runtime managed internalcall
  {
  } // end of method IOPCAsyncIO3::ReadMaxAge

  .method public hidebysig newslot abstract virtual 
          instance void  WriteVQT([in] uint32 dwCount,
                                  [in] uint32& phServer,
                                  [in] valuetype opcprox.tagOPCITEMVQT& pItemVQT,
                                  [in] uint32 dwTransactionID,
                                  [out] uint32& pdwCancelID,
                                  [out] native int ppErrors) runtime managed internalcall
  {
  } // end of method IOPCAsyncIO3::WriteVQT

  .method public hidebysig newslot abstract virtual 
          instance void  RefreshMaxAge([in] uint32 dwMaxAge,
                                       [in] uint32 dwTransactionID,
                                       [out] uint32& pdwCancelID) runtime managed internalcall
  {
  } // end of method IOPCAsyncIO3::RefreshMaxAge

} // end of class opcprox.IOPCAsyncIO3

.class interface public abstract auto ansi import opcprox.IOPCGroupStateMgt2
       implements opcprox.IOPCGroupStateMgt
{
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 38 45 33 36 38 36 36 36 2D 44 37 32 45   // ..$8E368666-D72E
                                                                                                  2D 34 46 37 38 2D 38 37 45 44 2D 36 34 37 36 31   // -4F78-87ED-64761
                                                                                                  31 43 36 31 43 39 46 00 00 )                      // 1C61C9F..
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(valuetype [mscorlib]System.Runtime.InteropServices.ComInterfaceType) = ( 01 00 01 00 00 00 00 00 ) 
  .method public hidebysig newslot abstract virtual 
          instance void  GetState([out] uint32& pUpdateRate,
                                  [out] int32& pActive,
                                  [out] string&  marshal( lpwstr) ppName,
                                  [out] int32& pTimeBias,
                                  [out] float32& pPercentDeadband,
                                  [out] uint32& pLCID,
                                  [out] uint32& phClientGroup,
                                  [out] uint32& phServerGroup) runtime managed internalcall
  {
  } // end of method IOPCGroupStateMgt2::GetState

  .method public hidebysig newslot abstract virtual 
          instance void  SetState([in] uint32& pRequestedUpdateRate,
                                  [out] uint32& pRevisedUpdateRate,
                                  [in] int32& pActive,
                                  [in] int32& pTimeBias,
                                  [in] float32& pPercentDeadband,
                                  [in] uint32& pLCID,
                                  [in] uint32& phClientGroup) runtime managed internalcall
  {
  } // end of method IOPCGroupStateMgt2::SetState

  .method public hidebysig newslot abstract virtual 
          instance void  SetName([in] string  marshal( lpwstr) szName) runtime managed internalcall
  {
  } // end of method IOPCGroupStateMgt2::SetName

  .method public hidebysig newslot abstract virtual 
          instance void  CloneGroup([in] string  marshal( lpwstr) szName,
                                    [in] valuetype [mscorlib]System.Guid& riid,
                                    [out] object&  marshal( iunknown ) ppUnk) runtime managed internalcall
  {
  } // end of method IOPCGroupStateMgt2::CloneGroup

  .method public hidebysig newslot abstract virtual 
          instance void  SetKeepAlive([in] uint32 dwKeepAliveTime,
                                      [out] uint32& pdwRevisedKeepAliveTime) runtime managed internalcall
  {
  } // end of method IOPCGroupStateMgt2::SetKeepAlive

  .method public hidebysig newslot abstract virtual 
          instance void  GetKeepAlive([out] uint32& pdwKeepAliveTime) runtime managed internalcall
  {
  } // end of method IOPCGroupStateMgt2::GetKeepAlive

} // end of class opcprox.IOPCGroupStateMgt2


// =============================================================

// *********** DISASSEMBLY COMPLETE ***********************
