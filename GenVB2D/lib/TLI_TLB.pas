unit TLI_TLB;

// ************************************************************************ //
// WARNING                                                                    
// -------                                                                    
// The types declared in this file were generated from data read from a       
// Type Library. If this type library is explicitly or indirectly (via        
// another type library referring to this type library) re-imported, or the   
// 'Refresh' command of the Type Library Editor activated while editing the   
// Type Library, the contents of this file will be regenerated and all        
// manual modifications will be lost.                                         
// ************************************************************************ //

// PASTLWTR : 1.2
// File generated on 2/4/2008 8:31:45 PM from Type Library described below.

// ************************************************************************  //
// Type Lib: C:\Windows\system32\TLBINF32.DLL (1)
// LIBID: {8B217740-717D-11CE-AB5B-D41203C10000}
// LCID: 0
// Helpfile: C:\Windows\system32\tlbinf32.chm
// HelpString: TypeLib Information
// DepndLst: 
//   (1) v2.0 stdole, (C:\Windows\system32\stdole2.tlb)
// Errors:
//   Hint: Parameter 'Object' of _TLIApplication.InterfaceInfoFromObject changed to 'Object_'
//   Hint: Parameter 'Object' of _TLIApplication.InvokeHook changed to 'Object_'
//   Hint: Parameter 'Object' of _TLIApplication.InvokeHookArray changed to 'Object_'
//   Hint: Parameter 'Object' of _TLIApplication.InvokeHookSub changed to 'Object_'
//   Hint: Parameter 'Object' of _TLIApplication.InvokeHookArraySub changed to 'Object_'
//   Hint: Parameter 'Object' of _TLIApplication.ClassInfoFromObject changed to 'Object_'
//   Hint: Parameter 'Object' of _TLIApplication.InvokeID changed to 'Object_'
//   Error creating palette bitmap of (TSearchHelper) : Server C:\Windows\system32\TLBINF32.DLL contains no icons
//   Error creating palette bitmap of (TTypeLibInfo) : Server C:\Windows\system32\TLBINF32.DLL contains no icons
// ************************************************************************ //
// *************************************************************************//
// NOTE:                                                                      
// Items guarded by $IFDEF_LIVE_SERVER_AT_DESIGN_TIME are used by properties  
// which return objects that may need to be explicitly created via a function 
// call prior to any access via the property. These items have been disabled  
// in order to prevent accidental use from within the object inspector. You   
// may enable them by defining LIVE_SERVER_AT_DESIGN_TIME or by selectively   
// removing them from the $IFDEF blocks. However, such items must still be    
// programmatically created via a method of the appropriate CoClass before    
// they can be used.                                                          
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
interface

uses Windows, ActiveX, Classes, Graphics, OleServer, StdVCL, Variants;
  

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  TLIMajorVersion = 1;
  TLIMinorVersion = 0;

  LIBID_TLI: TGUID = '{8B217740-717D-11CE-AB5B-D41203C10000}';

  IID_VarTypeInfo: TGUID = '{8B21774B-717D-11CE-AB5B-D41203C10000}';
  IID_ParameterInfo: TGUID = '{8B217749-717D-11CE-AB5B-D41203C10000}';
  IID_Parameters: TGUID = '{8B21774A-717D-11CE-AB5B-D41203C10000}';
  IID_MemberInfo: TGUID = '{8B217747-717D-11CE-AB5B-D41203C10000}';
  IID_Members: TGUID = '{8B217748-717D-11CE-AB5B-D41203C10000}';
  IID_InterfaceInfo: TGUID = '{8B217741-717D-11CE-AB5B-D41203C10000}';
  IID__BaseTypeInfos: TGUID = '{8B217750-717D-11CE-AB5B-D41203C10000}';
  IID_Interfaces: TGUID = '{8B217742-717D-11CE-AB5B-D41203C10000}';
  IID_CoClassInfo: TGUID = '{8B217743-717D-11CE-AB5B-D41203C10000}';
  IID_CoClasses: TGUID = '{8B217744-717D-11CE-AB5B-D41203C10000}';
  IID_ConstantInfo: TGUID = '{8B21774D-717D-11CE-AB5B-D41203C10000}';
  IID_Constants: TGUID = '{8B21774C-717D-11CE-AB5B-D41203C10000}';
  IID_DeclarationInfo: TGUID = '{8B21774F-717D-11CE-AB5B-D41203C10000}';
  IID_Declarations: TGUID = '{8B21774E-717D-11CE-AB5B-D41203C10000}';
  IID__SearchHelper: TGUID = '{8B217751-717D-11CE-AB5B-D41203C10000}';
  CLASS_SearchHelper: TGUID = '{8B217752-717D-11CE-AB5B-D41203C10000}';
  IID_TypeInfo: TGUID = '{8B217759-717D-11CE-AB5B-D41203C10000}';
  IID__TypeLibInfo: TGUID = '{8B217745-717D-11CE-AB5B-D41203C10000}';
  CLASS_TypeLibInfo: TGUID = '{8B217746-717D-11CE-AB5B-D41203C10000}';
  IID_SearchItem: TGUID = '{8B217756-717D-11CE-AB5B-D41203C10000}';
  IID_SearchResults: TGUID = '{8B217757-717D-11CE-AB5B-D41203C10000}';
  IID_ListBoxNotification: TGUID = '{8B217758-717D-11CE-AB5B-D41203C10000}';
  IID_CustomSort: TGUID = '{8B21775F-717D-11CE-AB5B-D41203C10000}';
  IID_CustomFilter: TGUID = '{8B217760-717D-11CE-AB5B-D41203C10000}';
  IID__TLIApplication: TGUID = '{8B21775D-717D-11CE-AB5B-D41203C10000}';
  CLASS_TLIApplication: TGUID = '{8B21775E-717D-11CE-AB5B-D41203C10000}';
  IID_TypeInfos: TGUID = '{8B21775A-717D-11CE-AB5B-D41203C10000}';
  IID_RecordInfo: TGUID = '{8B21775B-717D-11CE-AB5B-D41203C10000}';
  IID_Records: TGUID = '{8B21775C-717D-11CE-AB5B-D41203C10000}';
  IID_IntrinsicAliasInfo: TGUID = '{8B217761-717D-11CE-AB5B-D41203C10000}';
  IID_IntrinsicAliases: TGUID = '{8B217762-717D-11CE-AB5B-D41203C10000}';
  IID_CustomData: TGUID = '{8B217763-717D-11CE-AB5B-D41203C10000}';
  IID_CustomDataCollection: TGUID = '{8B217764-717D-11CE-AB5B-D41203C10000}';
  IID_UnionInfo: TGUID = '{8B217765-717D-11CE-AB5B-D41203C10000}';
  IID_Unions: TGUID = '{8B217766-717D-11CE-AB5B-D41203C10000}';

// *********************************************************************//
// Declaration of Enumerations defined in Type Library                    
// *********************************************************************//
// Constants for enum TliErrors
type
  TliErrors = TOleEnum;
const
  tliErrNoCurrentTypelib = $80040201;
  tliErrCantLoadLibrary = $80040202;
  tliErrTypeLibNotRegistered = $80040203;
  tliErrSearchResultsChanged = $80040204;
  tliErrNotApplicable = $80040205;
  tliErrIncompatibleData = $80040206;
  tliErrIncompatibleSearchType = $80040207;
  tliErrIncompatibleTypeKind = $80040208;
  tliErrInaccessibleImportLib = $80040209;
  tliErrNoDefaultValue = $8004020A;
  tliErrDataNotAvailable = $8004020B;
  tliErrNotAnEntryPoint = $8004020C;
  tliErrStopFiltering = $8004020D;
  tliErrArrayBoundsNotAvailable = $8004020E;
  tliErrSearchResultsNotSorted = $8004020F;
  tliErrTypeNotArray = $80040210;

// Constants for enum TypeFlags
type
  TypeFlags = TOleEnum;
const
  TYPEFLAG_NONE = $00000000;
  TYPEFLAG_FAPPOBJECT = $00000001;
  TYPEFLAG_FCANCREATE = $00000002;
  TYPEFLAG_FLICENSED = $00000004;
  TYPEFLAG_FPREDECLID = $00000008;
  TYPEFLAG_FHIDDEN = $00000010;
  TYPEFLAG_FCONTROL = $00000020;
  TYPEFLAG_FDUAL = $00000040;
  TYPEFLAG_FNONEXTENSIBLE = $00000080;
  TYPEFLAG_FOLEAUTOMATION = $00000100;
  TYPEFLAG_FRESTRICTED = $00000200;
  TYPEFLAG_FAGGREGATABLE = $00000400;
  TYPEFLAG_FREPLACEABLE = $00000800;
  TYPEFLAG_FDISPATCHABLE = $00001000;
  TYPEFLAG_FREVERSEBIND = $00002000;
  TYPEFLAG_FPROXY = $00004000;
  TYPEFLAG_DEFAULTFILTER = $00000210;
  TYPEFLAG_COCLASSATTRIBUTES = $0000063F;
  TYPEFLAG_INTERFACEATTRIBUTES = $00007BD0;
  TYPEFLAG_DISPATCHATTRIBUTES = $00005A90;
  TYPEFLAG_ALIASATTRIBUTES = $00000210;
  TYPEFLAG_MODULEATTRIBUTES = $00000210;
  TYPEFLAG_ENUMATTRIBUTES = $00000210;
  TYPEFLAG_RECORDATTRIBUTES = $00000210;
  TYPEFLAG_UNIONATTRIBUTES = $00000210;

// Constants for enum ImplTypeFlags
type
  ImplTypeFlags = TOleEnum;
const
  IMPLTYPEFLAG_FDEFAULT = $00000001;
  IMPLTYPEFLAG_FSOURCE = $00000002;
  IMPLTYPEFLAG_FRESTRICTED = $00000004;
  IMPLTYPEFLAG_FDEFAULTVTABLE = $00000008;

// Constants for enum TypeKinds
type
  TypeKinds = TOleEnum;
const
  TKIND_ENUM = $00000000;
  TKIND_RECORD = $00000001;
  TKIND_MODULE = $00000002;
  TKIND_INTERFACE = $00000003;
  TKIND_DISPATCH = $00000004;
  TKIND_COCLASS = $00000005;
  TKIND_ALIAS = $00000006;
  TKIND_UNION = $00000007;
  TKIND_MAX = $00000008;

// Constants for enum FuncFlags
type
  FuncFlags = TOleEnum;
const
  FUNCFLAG_NONE = $00000000;
  FUNCFLAG_FRESTRICTED = $00000001;
  FUNCFLAG_FSOURCE = $00000002;
  FUNCFLAG_FBINDABLE = $00000004;
  FUNCFLAG_FREQUESTEDIT = $00000008;
  FUNCFLAG_FDISPLAYBIND = $00000010;
  FUNCFLAG_FDEFAULTBIND = $00000020;
  FUNCFLAG_FHIDDEN = $00000040;
  FUNCFLAG_FUSESGETLASTERROR = $00000080;
  FUNCFLAG_FDEFAULTCOLLELEM = $00000100;
  FUNCFLAG_FUIDEFAULT = $00000200;
  FUNCFLAG_FNONBROWSABLE = $00000400;
  FUNCFLAG_FREPLACEABLE = $00000800;
  FUNCFLAG_FIMMEDIATEBIND = $00001000;
  FUNCFLAG_DEFAULTFILTER = $00000041;

// Constants for enum VarFlags
type
  VarFlags = TOleEnum;
const
  VARFLAG_NONE = $00000000;
  VARFLAG_FREADONLY = $00000001;
  VARFLAG_FSOURCE = $00000002;
  VARFLAG_FBINDABLE = $00000004;
  VARFLAG_FREQUESTEDIT = $00000008;
  VARFLAG_FDISPLAYBIND = $00000010;
  VARFLAG_FDEFAULTBIND = $00000020;
  VARFLAG_FHIDDEN = $00000040;
  VARFLAG_FRESTRICTED = $00000080;
  VARFLAG_FDEFAULTCOLLELEM = $00000100;
  VARFLAG_FUIDEFAULT = $00000200;
  VARFLAG_FNONBROWSABLE = $00000400;
  VARFLAG_FREPLACEABLE = $00000800;
  VARFLAG_FIMMEDIATEBIND = $00001000;
  VARFLAG_DEFAULTFILTER = $000000C0;

// Constants for enum SysKinds
type
  SysKinds = TOleEnum;
const
  SYS_WIN16 = $00000000;
  SYS_WIN32 = $00000001;
  SYS_MAC = $00000002;

// Constants for enum LibFlags
type
  LibFlags = TOleEnum;
const
  LIBFLAG_FRESTRICTED = $00000001;
  LIBFLAG_FCONTROL = $00000002;
  LIBFLAG_FHIDDEN = $00000004;
  LIBFLAG_FHASDISKIMAGE = $00000008;

// Constants for enum InvokeKinds
type
  InvokeKinds = TOleEnum;
const
  INVOKE_UNKNOWN = $00000000;
  INVOKE_FUNC = $00000001;
  INVOKE_PROPERTYGET = $00000002;
  INVOKE_PROPERTYPUT = $00000004;
  INVOKE_PROPERTYPUTREF = $00000008;
  INVOKE_EVENTFUNC = $00000010;
  INVOKE_CONST = $00000020;

// Constants for enum IDLFlags
type
  IDLFlags = TOleEnum;
const
  IDLFLAG_NONE = $00000000;
  IDLFLAG_FIN = $00000001;
  IDLFLAG_FOUT = $00000002;
  IDLFLAG_FLCID = $00000004;
  IDLFLAG_FRETVAL = $00000008;

// Constants for enum ParamFlags
type
  ParamFlags = TOleEnum;
const
  PARAMFLAG_NONE = $00000000;
  PARAMFLAG_FIN = $00000001;
  PARAMFLAG_FOUT = $00000002;
  PARAMFLAG_FLCID = $00000004;
  PARAMFLAG_FRETVAL = $00000008;
  PARAMFLAG_FOPT = $00000010;
  PARAMFLAG_FHASDEFAULT = $00000020;
  PARAMFLAG_FHASCUSTDATA = $00000040;

// Constants for enum DescKinds
type
  DescKinds = TOleEnum;
const
  DESCKIND_NONE = $00000000;
  DESCKIND_FUNCDESC = $00000001;
  DESCKIND_VARDESC = $00000002;

// Constants for enum TliVarType
type
  TliVarType = TOleEnum;
const
  VT_EMPTY = $00000000;
  VT_NULL = $00000001;
  VT_I2 = $00000002;
  VT_I4 = $00000003;
  VT_R4 = $00000004;
  VT_R8 = $00000005;
  VT_CY = $00000006;
  VT_DATE = $00000007;
  VT_BSTR = $00000008;
  VT_DISPATCH = $00000009;
  VT_ERROR = $0000000A;
  VT_BOOL = $0000000B;
  VT_VARIANT = $0000000C;
  VT_UNKNOWN = $0000000D;
  VT_DECIMAL = $0000000E;
  VT_I1 = $00000010;
  VT_UI1 = $00000011;
  VT_UI2 = $00000012;
  VT_UI4 = $00000013;
  VT_I8 = $00000014;
  VT_UI8 = $00000015;
  VT_INT = $00000016;
  VT_UINT = $00000017;
  VT_VOID = $00000018;
  VT_HRESULT = $00000019;
  VT_PTR = $0000001A;
  VT_SAFEARRAY = $0000001B;
  VT_CARRAY = $0000001C;
  VT_USERDEFINED = $0000001D;
  VT_LPSTR = $0000001E;
  VT_LPWSTR = $0000001F;
  VT_RECORD = $00000024;
  VT_FILETIME = $00000040;
  VT_BLOB = $00000041;
  VT_STREAM = $00000042;
  VT_STORAGE = $00000043;
  VT_STREAMED_OBJECT = $00000044;
  VT_STORED_OBJECT = $00000045;
  VT_BLOB_OBJECT = $00000046;
  VT_CF = $00000047;
  VT_CLSID = $00000048;
  VT_VECTOR = $00001000;
  VT_ARRAY = $00002000;
  VT_BYREF = $00004000;
  VT_RESERVED = $00008000;

// Constants for enum TliSearchTypes
type
  TliSearchTypes = TOleEnum;
const
  tliStDefault = $00001000;
  tliStClasses = $00000001;
  tliStEvents = $00000002;
  tliStConstants = $00000004;
  tliStDeclarations = $00000008;
  tliStAppObject = $00000010;
  tliStRecords = $00000020;
  tliStIntrinsicAliases = $00000040;
  tliStUnions = $00000080;
  tliStAll = $000000EF;

// Constants for enum TliWindowTypes
type
  TliWindowTypes = TOleEnum;
const
  tliWtListBox = $00000000;
  tliWtComboBox = $00000001;

// Constants for enum TliItemDataTypes
type
  TliItemDataTypes = TOleEnum;
const
  tliIdtMemberID = $00000000;
  tliIdtInvokeKinds = $00000001;

// Constants for enum TliCustomFilterAction
type
  TliCustomFilterAction = TOleEnum;
const
  tliCfaLeave = $00000000;
  tliCfaDuplicate = $00000001;
  tliCfaExtract = $00000002;
  tliCfaDelete = $00000003;

// Constants for enum CallConvs
type
  CallConvs = TOleEnum;
const
  CC_FASTCALL = $00000000;
  CC_CDECL = $00000001;
  CC_MSCPASCAL = $00000002;
  CC_PASCAL = $00000002;
  CC_MACPASCAL = $00000003;
  CC_STDCALL = $00000004;
  CC_FPFASTCALL = $00000005;
  CC_SYSCALL = $00000006;
  CC_MPWCDECL = $00000007;
  CC_MPWPASCAL = $00000008;
  CC_MAX = $00000009;

type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  VarTypeInfo = interface;
  VarTypeInfoDisp = dispinterface;
  ParameterInfo = interface;
  ParameterInfoDisp = dispinterface;
  Parameters = interface;
  ParametersDisp = dispinterface;
  MemberInfo = interface;
  MemberInfoDisp = dispinterface;
  Members = interface;
  MembersDisp = dispinterface;
  InterfaceInfo = interface;
  InterfaceInfoDisp = dispinterface;
  _BaseTypeInfos = interface;
  _BaseTypeInfosDisp = dispinterface;
  Interfaces = interface;
  InterfacesDisp = dispinterface;
  CoClassInfo = interface;
  CoClassInfoDisp = dispinterface;
  CoClasses = interface;
  CoClassesDisp = dispinterface;
  ConstantInfo = interface;
  ConstantInfoDisp = dispinterface;
  Constants = interface;
  ConstantsDisp = dispinterface;
  DeclarationInfo = interface;
  DeclarationInfoDisp = dispinterface;
  Declarations = interface;
  DeclarationsDisp = dispinterface;
  _SearchHelper = interface;
  _SearchHelperDisp = dispinterface;
  TypeInfo = interface;
  TypeInfoDisp = dispinterface;
  _TypeLibInfo = interface;
  _TypeLibInfoDisp = dispinterface;
  SearchItem = interface;
  SearchItemDisp = dispinterface;
  SearchResults = interface;
  SearchResultsDisp = dispinterface;
  ListBoxNotification = interface;
  ListBoxNotificationDisp = dispinterface;
  CustomSort = interface;
  CustomSortDisp = dispinterface;
  CustomFilter = interface;
  CustomFilterDisp = dispinterface;
  _TLIApplication = interface;
  _TLIApplicationDisp = dispinterface;
  TypeInfos = interface;
  TypeInfosDisp = dispinterface;
  RecordInfo = interface;
  RecordInfoDisp = dispinterface;
  Records = interface;
  RecordsDisp = dispinterface;
  IntrinsicAliasInfo = interface;
  IntrinsicAliasInfoDisp = dispinterface;
  IntrinsicAliases = interface;
  IntrinsicAliasesDisp = dispinterface;
  CustomData = interface;
  CustomDataDisp = dispinterface;
  CustomDataCollection = interface;
  CustomDataCollectionDisp = dispinterface;
  UnionInfo = interface;
  UnionInfoDisp = dispinterface;
  Unions = interface;
  UnionsDisp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  SearchHelper = _SearchHelper;
  TypeLibInfo = _TypeLibInfo;
  TLIApplication = _TLIApplication;


// *********************************************************************//
// Declaration of structures, unions and aliases.                         
// *********************************************************************//
  POleVariant1 = ^OleVariant; {*}
  PPSafeArray1 = ^PSafeArray; {*}
  PPSafeArray2 = ^PSafeArray; {*}
  PPSafeArray3 = ^PSafeArray; {*}
  PWideString1 = ^WideString; {*}


// *********************************************************************//
// Interface: VarTypeInfo
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B21774B-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  VarTypeInfo = interface(IDispatch)
    ['{8B21774B-717D-11CE-AB5B-D41203C10000}']
    function Me: VarTypeInfo; safecall;
    procedure _placeholder_destructor; safecall;
    function Get__OldVarType: HResult; safecall;
    function Get_TypeInfo: TypeInfo; safecall;
    function Get_TypeInfoNumber: Smallint; safecall;
    function Get_TypedVariant: OleVariant; safecall;
    function Get_IsExternalType: WordBool; safecall;
    function Get_TypeLibInfoExternal: TypeLibInfo; safecall;
    function Get_PointerLevel: Smallint; safecall;
    function Get_VarType: TliVarType; safecall;
    function ArrayBounds(out Bounds: PSafeArray): Smallint; safecall;
    function Get_ElementPointerLevel: Smallint; safecall;
    property _OldVarType: HResult read Get__OldVarType;
    property TypeInfo: TypeInfo read Get_TypeInfo;
    property TypeInfoNumber: Smallint read Get_TypeInfoNumber;
    property TypedVariant: OleVariant read Get_TypedVariant;
    property IsExternalType: WordBool read Get_IsExternalType;
    property TypeLibInfoExternal: TypeLibInfo read Get_TypeLibInfoExternal;
    property PointerLevel: Smallint read Get_PointerLevel;
    property VarType: TliVarType read Get_VarType;
    property ElementPointerLevel: Smallint read Get_ElementPointerLevel;
  end;

// *********************************************************************//
// DispIntf:  VarTypeInfoDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B21774B-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  VarTypeInfoDisp = dispinterface
    ['{8B21774B-717D-11CE-AB5B-D41203C10000}']
    function Me: VarTypeInfo; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    property _OldVarType: HResult readonly dispid 1610743810;
    property TypeInfo: TypeInfo readonly dispid 1610743811;
    property TypeInfoNumber: Smallint readonly dispid 1610743812;
    property TypedVariant: OleVariant readonly dispid 1610743813;
    property IsExternalType: WordBool readonly dispid 1610743814;
    property TypeLibInfoExternal: TypeLibInfo readonly dispid 1610743815;
    property PointerLevel: Smallint readonly dispid 1610743816;
    property VarType: TliVarType readonly dispid 0;
    function ArrayBounds(out Bounds: {??PSafeArray}OleVariant): Smallint; dispid 1610743818;
    property ElementPointerLevel: Smallint readonly dispid 1610743819;
  end;

// *********************************************************************//
// Interface: ParameterInfo
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217749-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  ParameterInfo = interface(IDispatch)
    ['{8B217749-717D-11CE-AB5B-D41203C10000}']
    function Me: ParameterInfo; safecall;
    procedure _placeholder_destructor; safecall;
    function Get_Name: WideString; safecall;
    function Get_Optional: WordBool; safecall;
    function Get__OldFlags: HResult; safecall;
    function Get_VarTypeInfo: VarTypeInfo; safecall;
    function Get_Default: WordBool; safecall;
    function Get_DefaultValue: OleVariant; safecall;
    function Get_HasCustomData: WordBool; safecall;
    function Get_CustomDataCollection: CustomDataCollection; safecall;
    function Get_Flags: ParamFlags; safecall;
    property Name: WideString read Get_Name;
    property Optional: WordBool read Get_Optional;
    property _OldFlags: HResult read Get__OldFlags;
    property VarTypeInfo: VarTypeInfo read Get_VarTypeInfo;
    property Default: WordBool read Get_Default;
    property DefaultValue: OleVariant read Get_DefaultValue;
    property HasCustomData: WordBool read Get_HasCustomData;
    property CustomDataCollection: CustomDataCollection read Get_CustomDataCollection;
    property Flags: ParamFlags read Get_Flags;
  end;

// *********************************************************************//
// DispIntf:  ParameterInfoDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217749-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  ParameterInfoDisp = dispinterface
    ['{8B217749-717D-11CE-AB5B-D41203C10000}']
    function Me: ParameterInfo; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    property Name: WideString readonly dispid 0;
    property Optional: WordBool readonly dispid 1610743811;
    property _OldFlags: HResult readonly dispid 1610743812;
    property VarTypeInfo: VarTypeInfo readonly dispid 1610743813;
    property Default: WordBool readonly dispid 1610743814;
    property DefaultValue: OleVariant readonly dispid 1610743815;
    property HasCustomData: WordBool readonly dispid 1610743816;
    property CustomDataCollection: CustomDataCollection readonly dispid 1610743817;
    property Flags: ParamFlags readonly dispid 1610743818;
  end;

// *********************************************************************//
// Interface: Parameters
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B21774A-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  Parameters = interface(IDispatch)
    ['{8B21774A-717D-11CE-AB5B-D41203C10000}']
    function Me: Parameters; safecall;
    procedure _placeholder_destructor; safecall;
    function _NewEnum: IUnknown; safecall;
    function Get_Item(Index: Smallint): ParameterInfo; safecall;
    function Get_Count: Smallint; safecall;
    function Get_OptionalCount: Smallint; safecall;
    function Get_DefaultCount: Smallint; safecall;
    property Item[Index: Smallint]: ParameterInfo read Get_Item; default;
    property Count: Smallint read Get_Count;
    property OptionalCount: Smallint read Get_OptionalCount;
    property DefaultCount: Smallint read Get_DefaultCount;
  end;

// *********************************************************************//
// DispIntf:  ParametersDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B21774A-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  ParametersDisp = dispinterface
    ['{8B21774A-717D-11CE-AB5B-D41203C10000}']
    function Me: Parameters; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    function _NewEnum: IUnknown; dispid -4;
    property Item[Index: Smallint]: ParameterInfo readonly dispid 0; default;
    property Count: Smallint readonly dispid 1610743812;
    property OptionalCount: Smallint readonly dispid 1610743813;
    property DefaultCount: Smallint readonly dispid 1610743814;
  end;

// *********************************************************************//
// Interface: MemberInfo
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217747-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  MemberInfo = interface(IDispatch)
    ['{8B217747-717D-11CE-AB5B-D41203C10000}']
    function Me: MemberInfo; safecall;
    procedure _placeholder_destructor; safecall;
    function Get_Name: WideString; safecall;
    function Get__OldHelpString: HResult; safecall;
    function Get_HelpContext: Integer; safecall;
    function Get_HelpFile: WideString; safecall;
    function Get_Parameters: Parameters; safecall;
    function Get_ReturnType: VarTypeInfo; safecall;
    function Get_AttributeMask: Smallint; safecall;
    function Get_AttributeStrings(out AttributeArray: PSafeArray): Smallint; safecall;
    function Get__OldDescKind: HResult; safecall;
    function Get_Value: OleVariant; safecall;
    function Get_MemberId: Integer; safecall;
    function Get_VTableOffset: Smallint; safecall;
    function Get_InvokeKind: InvokeKinds; safecall;
    function Get_CustomDataCollection: CustomDataCollection; safecall;
    function Get_DescKind: DescKinds; safecall;
    procedure GetDllEntry(out DllName: WideString; out EntryName: WideString; out Ordinal: Smallint); safecall;
    function Get_HelpString(LCID: Integer): WideString; safecall;
    function Get_HelpStringDll(LCID: Integer): WideString; safecall;
    function Get_HelpStringContext: Integer; safecall;
    function Get_CallConv: CallConvs; safecall;
    property Name: WideString read Get_Name;
    property _OldHelpString: HResult read Get__OldHelpString;
    property HelpContext: Integer read Get_HelpContext;
    property HelpFile: WideString read Get_HelpFile;
    property Parameters: Parameters read Get_Parameters;
    property ReturnType: VarTypeInfo read Get_ReturnType;
    property AttributeMask: Smallint read Get_AttributeMask;
    property AttributeStrings[out AttributeArray: PSafeArray]: Smallint read Get_AttributeStrings;
    property _OldDescKind: HResult read Get__OldDescKind;
    property Value: OleVariant read Get_Value;
    property MemberId: Integer read Get_MemberId;
    property VTableOffset: Smallint read Get_VTableOffset;
    property InvokeKind: InvokeKinds read Get_InvokeKind;
    property CustomDataCollection: CustomDataCollection read Get_CustomDataCollection;
    property DescKind: DescKinds read Get_DescKind;
    property HelpString[LCID: Integer]: WideString read Get_HelpString;
    property HelpStringDll[LCID: Integer]: WideString read Get_HelpStringDll;
    property HelpStringContext: Integer read Get_HelpStringContext;
    property CallConv: CallConvs read Get_CallConv;
  end;

// *********************************************************************//
// DispIntf:  MemberInfoDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217747-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  MemberInfoDisp = dispinterface
    ['{8B217747-717D-11CE-AB5B-D41203C10000}']
    function Me: MemberInfo; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    property Name: WideString readonly dispid 0;
    property _OldHelpString: HResult readonly dispid 1610743811;
    property HelpContext: Integer readonly dispid 1610743812;
    property HelpFile: WideString readonly dispid 1610743813;
    property Parameters: Parameters readonly dispid 1610743814;
    property ReturnType: VarTypeInfo readonly dispid 1610743815;
    property AttributeMask: Smallint readonly dispid 1610743816;
    property AttributeStrings[out AttributeArray: {??PSafeArray}OleVariant]: Smallint readonly dispid 1610743817;
    property _OldDescKind: HResult readonly dispid 1610743818;
    property Value: OleVariant readonly dispid 1610743819;
    property MemberId: Integer readonly dispid 1610743820;
    property VTableOffset: Smallint readonly dispid 1610743821;
    property InvokeKind: InvokeKinds readonly dispid 1610743822;
    property CustomDataCollection: CustomDataCollection readonly dispid 1610743823;
    property DescKind: DescKinds readonly dispid 1610743824;
    procedure GetDllEntry(out DllName: WideString; out EntryName: WideString; out Ordinal: Smallint); dispid 1610743825;
    property HelpString[LCID: Integer]: WideString readonly dispid 1610743826;
    property HelpStringDll[LCID: Integer]: WideString readonly dispid 1610743827;
    property HelpStringContext: Integer readonly dispid 1610743828;
    property CallConv: CallConvs readonly dispid 1610743829;
  end;

// *********************************************************************//
// Interface: Members
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217748-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  Members = interface(IDispatch)
    ['{8B217748-717D-11CE-AB5B-D41203C10000}']
    function Me: Members; safecall;
    procedure _placeholder_destructor; safecall;
    function _NewEnum: IUnknown; safecall;
    function Get_Item(Index: Smallint): MemberInfo; safecall;
    function Get_Count: Smallint; safecall;
    procedure Set_FuncFilter(retVal: FuncFlags); safecall;
    function Get_FuncFilter: FuncFlags; safecall;
    procedure Set_VarFilter(retVal: VarFlags); safecall;
    function Get_VarFilter: VarFlags; safecall;
    procedure _OldFillList; safecall;
    function Get_GetFilteredMembers(ShowUnderscore: WordBool): SearchResults; safecall;
    function GetFilteredMembersDirect(hWnd: SYSINT; WindowType: TliWindowTypes; 
                                      ItemDataType: TliItemDataTypes; ShowUnderscore: WordBool): Smallint; safecall;
    property Item[Index: Smallint]: MemberInfo read Get_Item; default;
    property Count: Smallint read Get_Count;
    property FuncFilter: FuncFlags read Get_FuncFilter write Set_FuncFilter;
    property VarFilter: VarFlags read Get_VarFilter write Set_VarFilter;
    property GetFilteredMembers[ShowUnderscore: WordBool]: SearchResults read Get_GetFilteredMembers;
  end;

// *********************************************************************//
// DispIntf:  MembersDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217748-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  MembersDisp = dispinterface
    ['{8B217748-717D-11CE-AB5B-D41203C10000}']
    function Me: Members; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    function _NewEnum: IUnknown; dispid -4;
    property Item[Index: Smallint]: MemberInfo readonly dispid 0; default;
    property Count: Smallint readonly dispid 1610743812;
    property FuncFilter: FuncFlags dispid 1610743813;
    property VarFilter: VarFlags dispid 1610743815;
    procedure _OldFillList; dispid 1610743817;
    property GetFilteredMembers[ShowUnderscore: WordBool]: SearchResults readonly dispid 1610743818;
    function GetFilteredMembersDirect(hWnd: SYSINT; WindowType: TliWindowTypes; 
                                      ItemDataType: TliItemDataTypes; ShowUnderscore: WordBool): Smallint; dispid 1610743819;
  end;

// *********************************************************************//
// Interface: InterfaceInfo
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217741-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  InterfaceInfo = interface(IDispatch)
    ['{8B217741-717D-11CE-AB5B-D41203C10000}']
    function Me: InterfaceInfo; safecall;
    procedure _placeholder_destructor; safecall;
    function Get_Name: WideString; safecall;
    function Get_GUID: WideString; safecall;
    function Get__OldHelpString: HResult; safecall;
    function Get_HelpContext: Integer; safecall;
    function Get_HelpFile: WideString; safecall;
    function Get_AttributeMask: Smallint; safecall;
    function Get_AttributeStrings(out AttributeArray: PSafeArray): Smallint; safecall;
    function Get__OldTypeKind: HResult; safecall;
    function Get_TypeKindString: WideString; safecall;
    function Get_TypeInfoNumber: Smallint; safecall;
    function Get_VTableInterface: InterfaceInfo; safecall;
    function Get_GetMember(Index: OleVariant): MemberInfo; safecall;
    function Get_Members: Members; safecall;
    function Get_Parent: TypeLibInfo; safecall;
    function Get_ImpliedInterfaces: Interfaces; safecall;
    procedure _DefaultInterface; safecall;
    procedure _DefaultEventInterface; safecall;
    function Get_TypeKind: TypeKinds; safecall;
    function Get_ResolvedType: VarTypeInfo; safecall;
    function Get_CustomDataCollection: CustomDataCollection; safecall;
    function Get_HelpString(LCID: Integer): WideString; safecall;
    function Get_ITypeInfo: IUnknown; safecall;
    function Get_MajorVersion: Smallint; safecall;
    function Get_MinorVersion: Smallint; safecall;
    function Get_HelpStringDll(LCID: Integer): WideString; safecall;
    function Get_HelpStringContext: Integer; safecall;
    property Name: WideString read Get_Name;
    property GUID: WideString read Get_GUID;
    property _OldHelpString: HResult read Get__OldHelpString;
    property HelpContext: Integer read Get_HelpContext;
    property HelpFile: WideString read Get_HelpFile;
    property AttributeMask: Smallint read Get_AttributeMask;
    property AttributeStrings[out AttributeArray: PSafeArray]: Smallint read Get_AttributeStrings;
    property _OldTypeKind: HResult read Get__OldTypeKind;
    property TypeKindString: WideString read Get_TypeKindString;
    property TypeInfoNumber: Smallint read Get_TypeInfoNumber;
    property VTableInterface: InterfaceInfo read Get_VTableInterface;
    property GetMember[Index: OleVariant]: MemberInfo read Get_GetMember;
    property Members: Members read Get_Members;
    property Parent: TypeLibInfo read Get_Parent;
    property ImpliedInterfaces: Interfaces read Get_ImpliedInterfaces;
    property TypeKind: TypeKinds read Get_TypeKind;
    property ResolvedType: VarTypeInfo read Get_ResolvedType;
    property CustomDataCollection: CustomDataCollection read Get_CustomDataCollection;
    property HelpString[LCID: Integer]: WideString read Get_HelpString;
    property ITypeInfo: IUnknown read Get_ITypeInfo;
    property MajorVersion: Smallint read Get_MajorVersion;
    property MinorVersion: Smallint read Get_MinorVersion;
    property HelpStringDll[LCID: Integer]: WideString read Get_HelpStringDll;
    property HelpStringContext: Integer read Get_HelpStringContext;
  end;

// *********************************************************************//
// DispIntf:  InterfaceInfoDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217741-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  InterfaceInfoDisp = dispinterface
    ['{8B217741-717D-11CE-AB5B-D41203C10000}']
    function Me: InterfaceInfo; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    property Name: WideString readonly dispid 0;
    property GUID: WideString readonly dispid 1610743811;
    property _OldHelpString: HResult readonly dispid 1610743812;
    property HelpContext: Integer readonly dispid 1610743813;
    property HelpFile: WideString readonly dispid 1610743814;
    property AttributeMask: Smallint readonly dispid 1610743815;
    property AttributeStrings[out AttributeArray: {??PSafeArray}OleVariant]: Smallint readonly dispid 1610743816;
    property _OldTypeKind: HResult readonly dispid 1610743817;
    property TypeKindString: WideString readonly dispid 1610743818;
    property TypeInfoNumber: Smallint readonly dispid 1610743819;
    property VTableInterface: InterfaceInfo readonly dispid 1610743820;
    property GetMember[Index: OleVariant]: MemberInfo readonly dispid 1610743821;
    property Members: Members readonly dispid 1610743822;
    property Parent: TypeLibInfo readonly dispid 1610743823;
    property ImpliedInterfaces: Interfaces readonly dispid 1610743824;
    procedure _DefaultInterface; dispid 1610743825;
    procedure _DefaultEventInterface; dispid 1610743826;
    property TypeKind: TypeKinds readonly dispid 1610743827;
    property ResolvedType: VarTypeInfo readonly dispid 1610743828;
    property CustomDataCollection: CustomDataCollection readonly dispid 1610743829;
    property HelpString[LCID: Integer]: WideString readonly dispid 1610743830;
    property ITypeInfo: IUnknown readonly dispid 1610743831;
    property MajorVersion: Smallint readonly dispid 1610743832;
    property MinorVersion: Smallint readonly dispid 1610743833;
    property HelpStringDll[LCID: Integer]: WideString readonly dispid 1610743834;
    property HelpStringContext: Integer readonly dispid 1610743835;
  end;

// *********************************************************************//
// Interface: _BaseTypeInfos
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217750-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  _BaseTypeInfos = interface(IDispatch)
    ['{8B217750-717D-11CE-AB5B-D41203C10000}']
    function Me: TypeInfos; safecall;
    procedure _placeholder_destructor; safecall;
    function _NewEnum: IUnknown; safecall;
    function Get_Count: Smallint; safecall;
    property Count: Smallint read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  _BaseTypeInfosDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217750-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  _BaseTypeInfosDisp = dispinterface
    ['{8B217750-717D-11CE-AB5B-D41203C10000}']
    function Me: TypeInfos; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    function _NewEnum: IUnknown; dispid -4;
    property Count: Smallint readonly dispid 1610743811;
  end;

// *********************************************************************//
// Interface: Interfaces
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217742-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  Interfaces = interface(_BaseTypeInfos)
    ['{8B217742-717D-11CE-AB5B-D41203C10000}']
    function Get_Item(Index: Smallint): InterfaceInfo; safecall;
    function Get_IndexedItem(TypeInfoNumber: Smallint): InterfaceInfo; safecall;
    function Get_NamedItem(var TypeInfoName: WideString): InterfaceInfo; safecall;
    property Item[Index: Smallint]: InterfaceInfo read Get_Item; default;
    property IndexedItem[TypeInfoNumber: Smallint]: InterfaceInfo read Get_IndexedItem;
    property NamedItem[var TypeInfoName: WideString]: InterfaceInfo read Get_NamedItem;
  end;

// *********************************************************************//
// DispIntf:  InterfacesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217742-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  InterfacesDisp = dispinterface
    ['{8B217742-717D-11CE-AB5B-D41203C10000}']
    property Item[Index: Smallint]: InterfaceInfo readonly dispid 0; default;
    property IndexedItem[TypeInfoNumber: Smallint]: InterfaceInfo readonly dispid 1610809345;
    property NamedItem[var TypeInfoName: WideString]: InterfaceInfo readonly dispid 1610809346;
    function Me: TypeInfos; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    function _NewEnum: IUnknown; dispid -4;
    property Count: Smallint readonly dispid 1610743811;
  end;

// *********************************************************************//
// Interface: CoClassInfo
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217743-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  CoClassInfo = interface(IDispatch)
    ['{8B217743-717D-11CE-AB5B-D41203C10000}']
    function Me: CoClassInfo; safecall;
    procedure _placeholder_destructor; safecall;
    function Get_Name: WideString; safecall;
    function Get_GUID: WideString; safecall;
    function Get__OldHelpString: HResult; safecall;
    function Get_HelpContext: Integer; safecall;
    function Get_HelpFile: WideString; safecall;
    function Get_AttributeMask: Smallint; safecall;
    function Get_AttributeStrings(out AttributeArray: PSafeArray): Smallint; safecall;
    function Get__OldTypeKind: HResult; safecall;
    function Get_TypeKindString: WideString; safecall;
    function Get_TypeInfoNumber: Smallint; safecall;
    procedure _placeholder_VTableInterface; safecall;
    procedure _placeholder_GetMember; safecall;
    procedure _placeholder_Members; safecall;
    function Get_Parent: TypeLibInfo; safecall;
    function Get_Interfaces: Interfaces; safecall;
    function Get_DefaultInterface: InterfaceInfo; safecall;
    function Get_DefaultEventInterface: InterfaceInfo; safecall;
    function Get_TypeKind: TypeKinds; safecall;
    function Get_ResolvedType: VarTypeInfo; safecall;
    function Get_CustomDataCollection: CustomDataCollection; safecall;
    function Get_HelpString(LCID: Integer): WideString; safecall;
    function Get_ITypeInfo: IUnknown; safecall;
    function Get_MajorVersion: Smallint; safecall;
    function Get_MinorVersion: Smallint; safecall;
    function Get_HelpStringDll(LCID: Integer): WideString; safecall;
    function Get_HelpStringContext: Integer; safecall;
    property Name: WideString read Get_Name;
    property GUID: WideString read Get_GUID;
    property _OldHelpString: HResult read Get__OldHelpString;
    property HelpContext: Integer read Get_HelpContext;
    property HelpFile: WideString read Get_HelpFile;
    property AttributeMask: Smallint read Get_AttributeMask;
    property AttributeStrings[out AttributeArray: PSafeArray]: Smallint read Get_AttributeStrings;
    property _OldTypeKind: HResult read Get__OldTypeKind;
    property TypeKindString: WideString read Get_TypeKindString;
    property TypeInfoNumber: Smallint read Get_TypeInfoNumber;
    property Parent: TypeLibInfo read Get_Parent;
    property Interfaces: Interfaces read Get_Interfaces;
    property DefaultInterface: InterfaceInfo read Get_DefaultInterface;
    property DefaultEventInterface: InterfaceInfo read Get_DefaultEventInterface;
    property TypeKind: TypeKinds read Get_TypeKind;
    property ResolvedType: VarTypeInfo read Get_ResolvedType;
    property CustomDataCollection: CustomDataCollection read Get_CustomDataCollection;
    property HelpString[LCID: Integer]: WideString read Get_HelpString;
    property ITypeInfo: IUnknown read Get_ITypeInfo;
    property MajorVersion: Smallint read Get_MajorVersion;
    property MinorVersion: Smallint read Get_MinorVersion;
    property HelpStringDll[LCID: Integer]: WideString read Get_HelpStringDll;
    property HelpStringContext: Integer read Get_HelpStringContext;
  end;

// *********************************************************************//
// DispIntf:  CoClassInfoDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217743-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  CoClassInfoDisp = dispinterface
    ['{8B217743-717D-11CE-AB5B-D41203C10000}']
    function Me: CoClassInfo; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    property Name: WideString readonly dispid 0;
    property GUID: WideString readonly dispid 1610743811;
    property _OldHelpString: HResult readonly dispid 1610743812;
    property HelpContext: Integer readonly dispid 1610743813;
    property HelpFile: WideString readonly dispid 1610743814;
    property AttributeMask: Smallint readonly dispid 1610743815;
    property AttributeStrings[out AttributeArray: {??PSafeArray}OleVariant]: Smallint readonly dispid 1610743816;
    property _OldTypeKind: HResult readonly dispid 1610743817;
    property TypeKindString: WideString readonly dispid 1610743818;
    property TypeInfoNumber: Smallint readonly dispid 1610743819;
    procedure _placeholder_VTableInterface; dispid 1610743820;
    procedure _placeholder_GetMember; dispid 1610743821;
    procedure _placeholder_Members; dispid 1610743822;
    property Parent: TypeLibInfo readonly dispid 1610743823;
    property Interfaces: Interfaces readonly dispid 1610743824;
    property DefaultInterface: InterfaceInfo readonly dispid 1610743825;
    property DefaultEventInterface: InterfaceInfo readonly dispid 1610743826;
    property TypeKind: TypeKinds readonly dispid 1610743827;
    property ResolvedType: VarTypeInfo readonly dispid 1610743828;
    property CustomDataCollection: CustomDataCollection readonly dispid 1610743829;
    property HelpString[LCID: Integer]: WideString readonly dispid 1610743830;
    property ITypeInfo: IUnknown readonly dispid 1610743831;
    property MajorVersion: Smallint readonly dispid 1610743832;
    property MinorVersion: Smallint readonly dispid 1610743833;
    property HelpStringDll[LCID: Integer]: WideString readonly dispid 1610743834;
    property HelpStringContext: Integer readonly dispid 1610743835;
  end;

// *********************************************************************//
// Interface: CoClasses
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217744-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  CoClasses = interface(_BaseTypeInfos)
    ['{8B217744-717D-11CE-AB5B-D41203C10000}']
    function Get_Item(Index: Smallint): CoClassInfo; safecall;
    function Get_IndexedItem(TypeInfoNumber: Smallint): CoClassInfo; safecall;
    function Get_NamedItem(var TypeInfoName: WideString): CoClassInfo; safecall;
    property Item[Index: Smallint]: CoClassInfo read Get_Item; default;
    property IndexedItem[TypeInfoNumber: Smallint]: CoClassInfo read Get_IndexedItem;
    property NamedItem[var TypeInfoName: WideString]: CoClassInfo read Get_NamedItem;
  end;

// *********************************************************************//
// DispIntf:  CoClassesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217744-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  CoClassesDisp = dispinterface
    ['{8B217744-717D-11CE-AB5B-D41203C10000}']
    property Item[Index: Smallint]: CoClassInfo readonly dispid 0; default;
    property IndexedItem[TypeInfoNumber: Smallint]: CoClassInfo readonly dispid 1610809345;
    property NamedItem[var TypeInfoName: WideString]: CoClassInfo readonly dispid 1610809346;
    function Me: TypeInfos; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    function _NewEnum: IUnknown; dispid -4;
    property Count: Smallint readonly dispid 1610743811;
  end;

// *********************************************************************//
// Interface: ConstantInfo
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B21774D-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  ConstantInfo = interface(IDispatch)
    ['{8B21774D-717D-11CE-AB5B-D41203C10000}']
    function Me: TypeInfo; safecall;
    procedure _placeholder_destructor; safecall;
    function Get_Name: WideString; safecall;
    function Get_GUID: WideString; safecall;
    function Get__OldHelpString: HResult; safecall;
    function Get_HelpContext: Integer; safecall;
    function Get_HelpFile: WideString; safecall;
    function Get_AttributeMask: Smallint; safecall;
    function Get_AttributeStrings(out AttributeArray: PSafeArray): Smallint; safecall;
    function Get__OldTypeKind: HResult; safecall;
    function Get_TypeKindString: WideString; safecall;
    function Get_TypeInfoNumber: Smallint; safecall;
    procedure _placeholder_VTableInterface; safecall;
    function Get_GetMember(Index: OleVariant): MemberInfo; safecall;
    function Get_Members: Members; safecall;
    function Get_Parent: TypeLibInfo; safecall;
    procedure _ImpliedInterfaces; safecall;
    procedure _DefaultInterface; safecall;
    procedure _DefaultEventInterface; safecall;
    function Get_TypeKind: TypeKinds; safecall;
    function Get_ResolvedType: VarTypeInfo; safecall;
    function Get_CustomDataCollection: CustomDataCollection; safecall;
    function Get_HelpString(LCID: Integer): WideString; safecall;
    function Get_ITypeInfo: IUnknown; safecall;
    function Get_MajorVersion: Smallint; safecall;
    function Get_MinorVersion: Smallint; safecall;
    function Get_HelpStringDll(LCID: Integer): WideString; safecall;
    function Get_HelpStringContext: Integer; safecall;
    property Name: WideString read Get_Name;
    property GUID: WideString read Get_GUID;
    property _OldHelpString: HResult read Get__OldHelpString;
    property HelpContext: Integer read Get_HelpContext;
    property HelpFile: WideString read Get_HelpFile;
    property AttributeMask: Smallint read Get_AttributeMask;
    property AttributeStrings[out AttributeArray: PSafeArray]: Smallint read Get_AttributeStrings;
    property _OldTypeKind: HResult read Get__OldTypeKind;
    property TypeKindString: WideString read Get_TypeKindString;
    property TypeInfoNumber: Smallint read Get_TypeInfoNumber;
    property GetMember[Index: OleVariant]: MemberInfo read Get_GetMember;
    property Members: Members read Get_Members;
    property Parent: TypeLibInfo read Get_Parent;
    property TypeKind: TypeKinds read Get_TypeKind;
    property ResolvedType: VarTypeInfo read Get_ResolvedType;
    property CustomDataCollection: CustomDataCollection read Get_CustomDataCollection;
    property HelpString[LCID: Integer]: WideString read Get_HelpString;
    property ITypeInfo: IUnknown read Get_ITypeInfo;
    property MajorVersion: Smallint read Get_MajorVersion;
    property MinorVersion: Smallint read Get_MinorVersion;
    property HelpStringDll[LCID: Integer]: WideString read Get_HelpStringDll;
    property HelpStringContext: Integer read Get_HelpStringContext;
  end;

// *********************************************************************//
// DispIntf:  ConstantInfoDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B21774D-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  ConstantInfoDisp = dispinterface
    ['{8B21774D-717D-11CE-AB5B-D41203C10000}']
    function Me: TypeInfo; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    property Name: WideString readonly dispid 0;
    property GUID: WideString readonly dispid 1610743811;
    property _OldHelpString: HResult readonly dispid 1610743812;
    property HelpContext: Integer readonly dispid 1610743813;
    property HelpFile: WideString readonly dispid 1610743814;
    property AttributeMask: Smallint readonly dispid 1610743815;
    property AttributeStrings[out AttributeArray: {??PSafeArray}OleVariant]: Smallint readonly dispid 1610743816;
    property _OldTypeKind: HResult readonly dispid 1610743817;
    property TypeKindString: WideString readonly dispid 1610743818;
    property TypeInfoNumber: Smallint readonly dispid 1610743819;
    procedure _placeholder_VTableInterface; dispid 1610743820;
    property GetMember[Index: OleVariant]: MemberInfo readonly dispid 1610743821;
    property Members: Members readonly dispid 1610743822;
    property Parent: TypeLibInfo readonly dispid 1610743823;
    procedure _ImpliedInterfaces; dispid 1610743824;
    procedure _DefaultInterface; dispid 1610743825;
    procedure _DefaultEventInterface; dispid 1610743826;
    property TypeKind: TypeKinds readonly dispid 1610743827;
    property ResolvedType: VarTypeInfo readonly dispid 1610743828;
    property CustomDataCollection: CustomDataCollection readonly dispid 1610743829;
    property HelpString[LCID: Integer]: WideString readonly dispid 1610743830;
    property ITypeInfo: IUnknown readonly dispid 1610743831;
    property MajorVersion: Smallint readonly dispid 1610743832;
    property MinorVersion: Smallint readonly dispid 1610743833;
    property HelpStringDll[LCID: Integer]: WideString readonly dispid 1610743834;
    property HelpStringContext: Integer readonly dispid 1610743835;
  end;

// *********************************************************************//
// Interface: Constants
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B21774C-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  Constants = interface(_BaseTypeInfos)
    ['{8B21774C-717D-11CE-AB5B-D41203C10000}']
    function Get_Item(Index: Smallint): ConstantInfo; safecall;
    function Get_IndexedItem(TypeInfoNumber: Smallint): ConstantInfo; safecall;
    function Get_NamedItem(var TypeInfoName: WideString): ConstantInfo; safecall;
    property Item[Index: Smallint]: ConstantInfo read Get_Item; default;
    property IndexedItem[TypeInfoNumber: Smallint]: ConstantInfo read Get_IndexedItem;
    property NamedItem[var TypeInfoName: WideString]: ConstantInfo read Get_NamedItem;
  end;

// *********************************************************************//
// DispIntf:  ConstantsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B21774C-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  ConstantsDisp = dispinterface
    ['{8B21774C-717D-11CE-AB5B-D41203C10000}']
    property Item[Index: Smallint]: ConstantInfo readonly dispid 0; default;
    property IndexedItem[TypeInfoNumber: Smallint]: ConstantInfo readonly dispid 1610809345;
    property NamedItem[var TypeInfoName: WideString]: ConstantInfo readonly dispid 1610809346;
    function Me: TypeInfos; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    function _NewEnum: IUnknown; dispid -4;
    property Count: Smallint readonly dispid 1610743811;
  end;

// *********************************************************************//
// Interface: DeclarationInfo
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B21774F-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  DeclarationInfo = interface(IDispatch)
    ['{8B21774F-717D-11CE-AB5B-D41203C10000}']
    function Me: DeclarationInfo; safecall;
    procedure _placeholder_destructor; safecall;
    function Get_Name: WideString; safecall;
    function Get_GUID: WideString; safecall;
    function Get__OldHelpString: HResult; safecall;
    function Get_HelpContext: Integer; safecall;
    function Get_HelpFile: WideString; safecall;
    function Get_AttributeMask: Smallint; safecall;
    function Get_AttributeStrings(out AttributeArray: PSafeArray): Smallint; safecall;
    function Get__OldTypeKind: HResult; safecall;
    function Get_TypeKindString: WideString; safecall;
    function Get_TypeInfoNumber: Smallint; safecall;
    procedure _placeholder_VTableInterface; safecall;
    function Get_GetMember(Index: OleVariant): MemberInfo; safecall;
    function Get_Members: Members; safecall;
    function Get_Parent: TypeLibInfo; safecall;
    procedure _ImpliedInterfaces; safecall;
    procedure _DefaultInterface; safecall;
    procedure _DefaultEventInterface; safecall;
    function Get_TypeKind: TypeKinds; safecall;
    function Get_ResolvedType: VarTypeInfo; safecall;
    function Get_CustomDataCollection: CustomDataCollection; safecall;
    function Get_HelpString(LCID: Integer): WideString; safecall;
    function Get_ITypeInfo: IUnknown; safecall;
    function Get_MajorVersion: Smallint; safecall;
    function Get_MinorVersion: Smallint; safecall;
    function Get_HelpStringDll(LCID: Integer): WideString; safecall;
    function Get_HelpStringContext: Integer; safecall;
    property Name: WideString read Get_Name;
    property GUID: WideString read Get_GUID;
    property _OldHelpString: HResult read Get__OldHelpString;
    property HelpContext: Integer read Get_HelpContext;
    property HelpFile: WideString read Get_HelpFile;
    property AttributeMask: Smallint read Get_AttributeMask;
    property AttributeStrings[out AttributeArray: PSafeArray]: Smallint read Get_AttributeStrings;
    property _OldTypeKind: HResult read Get__OldTypeKind;
    property TypeKindString: WideString read Get_TypeKindString;
    property TypeInfoNumber: Smallint read Get_TypeInfoNumber;
    property GetMember[Index: OleVariant]: MemberInfo read Get_GetMember;
    property Members: Members read Get_Members;
    property Parent: TypeLibInfo read Get_Parent;
    property TypeKind: TypeKinds read Get_TypeKind;
    property ResolvedType: VarTypeInfo read Get_ResolvedType;
    property CustomDataCollection: CustomDataCollection read Get_CustomDataCollection;
    property HelpString[LCID: Integer]: WideString read Get_HelpString;
    property ITypeInfo: IUnknown read Get_ITypeInfo;
    property MajorVersion: Smallint read Get_MajorVersion;
    property MinorVersion: Smallint read Get_MinorVersion;
    property HelpStringDll[LCID: Integer]: WideString read Get_HelpStringDll;
    property HelpStringContext: Integer read Get_HelpStringContext;
  end;

// *********************************************************************//
// DispIntf:  DeclarationInfoDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B21774F-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  DeclarationInfoDisp = dispinterface
    ['{8B21774F-717D-11CE-AB5B-D41203C10000}']
    function Me: DeclarationInfo; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    property Name: WideString readonly dispid 0;
    property GUID: WideString readonly dispid 1610743811;
    property _OldHelpString: HResult readonly dispid 1610743812;
    property HelpContext: Integer readonly dispid 1610743813;
    property HelpFile: WideString readonly dispid 1610743814;
    property AttributeMask: Smallint readonly dispid 1610743815;
    property AttributeStrings[out AttributeArray: {??PSafeArray}OleVariant]: Smallint readonly dispid 1610743816;
    property _OldTypeKind: HResult readonly dispid 1610743817;
    property TypeKindString: WideString readonly dispid 1610743818;
    property TypeInfoNumber: Smallint readonly dispid 1610743819;
    procedure _placeholder_VTableInterface; dispid 1610743820;
    property GetMember[Index: OleVariant]: MemberInfo readonly dispid 1610743821;
    property Members: Members readonly dispid 1610743822;
    property Parent: TypeLibInfo readonly dispid 1610743823;
    procedure _ImpliedInterfaces; dispid 1610743824;
    procedure _DefaultInterface; dispid 1610743825;
    procedure _DefaultEventInterface; dispid 1610743826;
    property TypeKind: TypeKinds readonly dispid 1610743827;
    property ResolvedType: VarTypeInfo readonly dispid 1610743828;
    property CustomDataCollection: CustomDataCollection readonly dispid 1610743829;
    property HelpString[LCID: Integer]: WideString readonly dispid 1610743830;
    property ITypeInfo: IUnknown readonly dispid 1610743831;
    property MajorVersion: Smallint readonly dispid 1610743832;
    property MinorVersion: Smallint readonly dispid 1610743833;
    property HelpStringDll[LCID: Integer]: WideString readonly dispid 1610743834;
    property HelpStringContext: Integer readonly dispid 1610743835;
  end;

// *********************************************************************//
// Interface: Declarations
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B21774E-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  Declarations = interface(_BaseTypeInfos)
    ['{8B21774E-717D-11CE-AB5B-D41203C10000}']
    function Get_Item(Index: Smallint): DeclarationInfo; safecall;
    function Get_IndexedItem(TypeInfoNumber: Smallint): DeclarationInfo; safecall;
    function Get_NamedItem(var TypeInfoName: WideString): DeclarationInfo; safecall;
    property Item[Index: Smallint]: DeclarationInfo read Get_Item; default;
    property IndexedItem[TypeInfoNumber: Smallint]: DeclarationInfo read Get_IndexedItem;
    property NamedItem[var TypeInfoName: WideString]: DeclarationInfo read Get_NamedItem;
  end;

// *********************************************************************//
// DispIntf:  DeclarationsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B21774E-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  DeclarationsDisp = dispinterface
    ['{8B21774E-717D-11CE-AB5B-D41203C10000}']
    property Item[Index: Smallint]: DeclarationInfo readonly dispid 0; default;
    property IndexedItem[TypeInfoNumber: Smallint]: DeclarationInfo readonly dispid 1610809345;
    property NamedItem[var TypeInfoName: WideString]: DeclarationInfo readonly dispid 1610809346;
    function Me: TypeInfos; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    function _NewEnum: IUnknown; dispid -4;
    property Count: Smallint readonly dispid 1610743811;
  end;

// *********************************************************************//
// Interface: _SearchHelper
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217751-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  _SearchHelper = interface(IDispatch)
    ['{8B217751-717D-11CE-AB5B-D41203C10000}']
    function Me: TypeLibInfo; safecall;
    procedure _placeholder_destructor; safecall;
    procedure _OldInit; safecall;
    function Get_CheckHaveMatch(const Name: WideString): WordBool; safecall;
    function Get_Init(SysKind: SysKinds; LCID: Integer): HResult; safecall;
    property CheckHaveMatch[const Name: WideString]: WordBool read Get_CheckHaveMatch;
    property Init[SysKind: SysKinds; LCID: Integer]: HResult read Get_Init;
  end;

// *********************************************************************//
// DispIntf:  _SearchHelperDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217751-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  _SearchHelperDisp = dispinterface
    ['{8B217751-717D-11CE-AB5B-D41203C10000}']
    function Me: TypeLibInfo; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    procedure _OldInit; dispid 1610743810;
    property CheckHaveMatch[const Name: WideString]: WordBool readonly dispid 1610743811;
    property Init[SysKind: SysKinds; LCID: Integer]: HResult readonly dispid 1610743812;
  end;

// *********************************************************************//
// Interface: TypeInfo
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217759-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  TypeInfo = interface(IDispatch)
    ['{8B217759-717D-11CE-AB5B-D41203C10000}']
    function Me: TypeInfo; safecall;
    procedure _placeholder_destructor; safecall;
    function Get_Name: WideString; safecall;
    function Get_GUID: WideString; safecall;
    function Get__OldHelpString: HResult; safecall;
    function Get_HelpContext: Integer; safecall;
    function Get_HelpFile: WideString; safecall;
    function Get_AttributeMask: Smallint; safecall;
    function Get_AttributeStrings(out AttributeArray: PSafeArray): Smallint; safecall;
    function Get__OldTypeKind: HResult; safecall;
    function Get_TypeKindString: WideString; safecall;
    function Get_TypeInfoNumber: Smallint; safecall;
    function Get_VTableInterface: InterfaceInfo; safecall;
    function Get_GetMember(Index: OleVariant): MemberInfo; safecall;
    function Get_Members: Members; safecall;
    function Get_Parent: TypeLibInfo; safecall;
    function Get_Interfaces: Interfaces; safecall;
    function Get_DefaultInterface: InterfaceInfo; safecall;
    function Get_DefaultEventInterface: InterfaceInfo; safecall;
    function Get_TypeKind: TypeKinds; safecall;
    function Get_ResolvedType: VarTypeInfo; safecall;
    function Get_CustomDataCollection: CustomDataCollection; safecall;
    function Get_HelpString(LCID: Integer): WideString; safecall;
    function Get_ITypeInfo: IUnknown; safecall;
    function Get_MajorVersion: Smallint; safecall;
    function Get_MinorVersion: Smallint; safecall;
    function Get_HelpStringDll(LCID: Integer): WideString; safecall;
    function Get_HelpStringContext: Integer; safecall;
    property Name: WideString read Get_Name;
    property GUID: WideString read Get_GUID;
    property _OldHelpString: HResult read Get__OldHelpString;
    property HelpContext: Integer read Get_HelpContext;
    property HelpFile: WideString read Get_HelpFile;
    property AttributeMask: Smallint read Get_AttributeMask;
    property AttributeStrings[out AttributeArray: PSafeArray]: Smallint read Get_AttributeStrings;
    property _OldTypeKind: HResult read Get__OldTypeKind;
    property TypeKindString: WideString read Get_TypeKindString;
    property TypeInfoNumber: Smallint read Get_TypeInfoNumber;
    property VTableInterface: InterfaceInfo read Get_VTableInterface;
    property GetMember[Index: OleVariant]: MemberInfo read Get_GetMember;
    property Members: Members read Get_Members;
    property Parent: TypeLibInfo read Get_Parent;
    property Interfaces: Interfaces read Get_Interfaces;
    property DefaultInterface: InterfaceInfo read Get_DefaultInterface;
    property DefaultEventInterface: InterfaceInfo read Get_DefaultEventInterface;
    property TypeKind: TypeKinds read Get_TypeKind;
    property ResolvedType: VarTypeInfo read Get_ResolvedType;
    property CustomDataCollection: CustomDataCollection read Get_CustomDataCollection;
    property HelpString[LCID: Integer]: WideString read Get_HelpString;
    property ITypeInfo: IUnknown read Get_ITypeInfo;
    property MajorVersion: Smallint read Get_MajorVersion;
    property MinorVersion: Smallint read Get_MinorVersion;
    property HelpStringDll[LCID: Integer]: WideString read Get_HelpStringDll;
    property HelpStringContext: Integer read Get_HelpStringContext;
  end;

// *********************************************************************//
// DispIntf:  TypeInfoDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217759-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  TypeInfoDisp = dispinterface
    ['{8B217759-717D-11CE-AB5B-D41203C10000}']
    function Me: TypeInfo; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    property Name: WideString readonly dispid 0;
    property GUID: WideString readonly dispid 1610743811;
    property _OldHelpString: HResult readonly dispid 1610743812;
    property HelpContext: Integer readonly dispid 1610743813;
    property HelpFile: WideString readonly dispid 1610743814;
    property AttributeMask: Smallint readonly dispid 1610743815;
    property AttributeStrings[out AttributeArray: {??PSafeArray}OleVariant]: Smallint readonly dispid 1610743816;
    property _OldTypeKind: HResult readonly dispid 1610743817;
    property TypeKindString: WideString readonly dispid 1610743818;
    property TypeInfoNumber: Smallint readonly dispid 1610743819;
    property VTableInterface: InterfaceInfo readonly dispid 1610743820;
    property GetMember[Index: OleVariant]: MemberInfo readonly dispid 1610743821;
    property Members: Members readonly dispid 1610743822;
    property Parent: TypeLibInfo readonly dispid 1610743823;
    property Interfaces: Interfaces readonly dispid 1610743824;
    property DefaultInterface: InterfaceInfo readonly dispid 1610743825;
    property DefaultEventInterface: InterfaceInfo readonly dispid 1610743826;
    property TypeKind: TypeKinds readonly dispid 1610743827;
    property ResolvedType: VarTypeInfo readonly dispid 1610743828;
    property CustomDataCollection: CustomDataCollection readonly dispid 1610743829;
    property HelpString[LCID: Integer]: WideString readonly dispid 1610743830;
    property ITypeInfo: IUnknown readonly dispid 1610743831;
    property MajorVersion: Smallint readonly dispid 1610743832;
    property MinorVersion: Smallint readonly dispid 1610743833;
    property HelpStringDll[LCID: Integer]: WideString readonly dispid 1610743834;
    property HelpStringContext: Integer readonly dispid 1610743835;
  end;

// *********************************************************************//
// Interface: _TypeLibInfo
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217745-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  _TypeLibInfo = interface(IDispatch)
    ['{8B217745-717D-11CE-AB5B-D41203C10000}']
    function Me: TypeLibInfo; safecall;
    procedure _placeholder_destructor; safecall;
    function Get_ContainingFile: WideString; safecall;
    procedure Set_ContainingFile(const retVal: WideString); safecall;
    procedure LoadRegTypeLib(const TypeLibGuid: WideString; MajorVersion: Smallint; 
                             MinorVersion: Smallint; LCID: Integer); safecall;
    function Get_Name: WideString; safecall;
    function Get__OldHelpString: HResult; safecall;
    function Get_HelpContext: Integer; safecall;
    function Get_HelpFile: WideString; safecall;
    function Get_GUID: WideString; safecall;
    function Get_LCID: Integer; safecall;
    function Get__OldSysKind: HResult; safecall;
    function Get_MajorVersion: Smallint; safecall;
    function Get_MinorVersion: Smallint; safecall;
    function Get_AttributeMask: Smallint; safecall;
    function Get_AttributeStrings(out AttributeArray: PSafeArray): Smallint; safecall;
    function Get_CoClasses: CoClasses; safecall;
    function Get_Interfaces: Interfaces; safecall;
    function Get_Constants: Constants; safecall;
    function Get_Declarations: Declarations; safecall;
    function Get_TypeInfoCount: Smallint; safecall;
    function Get__OldGetTypeKind: HResult; safecall;
    function Get_GetTypeInfo(var Index: OleVariant): TypeInfo; safecall;
    function Get_GetTypeInfoNumber(const Name: WideString): Smallint; safecall;
    function IsSameLibrary(const CheckLib: TypeLibInfo): WordBool; safecall;
    procedure _OldResetSearchCriteria; safecall;
    procedure _OldGetTypesWithMember; safecall;
    procedure _OldGetMembersWithSubString; safecall;
    procedure _OldGetTypesWithSubString; safecall;
    procedure _OldCaseTypeName; safecall;
    procedure _OldCaseMemberName; safecall;
    procedure _OldFillTypesList; safecall;
    procedure _OldFillTypesCombo; safecall;
    procedure _OldFillMemberList; safecall;
    procedure _OldAddClassTypeToList; safecall;
    procedure Set_AppObjString(const retVal: WideString); safecall;
    procedure Set_LibNum(retVal: Smallint); safecall;
    function Get_ShowLibName: WordBool; safecall;
    procedure Set_ShowLibName(retVal: WordBool); safecall;
    procedure _Set__OldListBoxNotification(const Param1: ListBoxNotification); safecall;
    function Get_GetTypeKind(TypeInfoNumber: Smallint): TypeKinds; safecall;
    function Get_SysKind: SysKinds; safecall;
    function Get_SearchDefault: TliSearchTypes; safecall;
    procedure Set_SearchDefault(retVal: TliSearchTypes); safecall;
    function CaseTypeName(var bstrName: WideString; SearchType: TliSearchTypes): TliSearchTypes; safecall;
    function CaseMemberName(var bstrName: WideString; SearchType: TliSearchTypes): WordBool; safecall;
    procedure ResetSearchCriteria(TypeFilter: TypeFlags; IncludeEmptyTypes: WordBool; 
                                  ShowUnderscore: WordBool); safecall;
    function GetTypesWithMember(const MemberName: WideString; var StartResults: SearchResults; 
                                SearchType: TliSearchTypes; Sort: WordBool; ShowUnderscore: WordBool): SearchResults; safecall;
    function GetTypesWithMemberDirect(const MemberName: WideString; hWnd: SYSINT; 
                                      WindowType: TliWindowTypes; SearchType: TliSearchTypes; 
                                      ShowUnderscore: WordBool): Smallint; safecall;
    function GetMembersWithSubString(const SubString: WideString; var StartResults: SearchResults; 
                                     SearchType: TliSearchTypes; SearchMiddle: WordBool; 
                                     var Helper: SearchHelper; Sort: WordBool; 
                                     ShowUnderscore: WordBool): SearchResults; safecall;
    function GetMembersWithSubStringDirect(const SubString: WideString; hWnd: SYSINT; 
                                           WindowType: TliWindowTypes; SearchType: TliSearchTypes; 
                                           SearchMiddle: WordBool; var Helper: SearchHelper; 
                                           ShowUnderscore: WordBool): Smallint; safecall;
    function GetTypesWithSubString(const SubString: WideString; var StartResults: SearchResults; 
                                   SearchType: TliSearchTypes; SearchMiddle: WordBool; 
                                   Sort: WordBool): SearchResults; safecall;
    function GetTypesWithSubStringDirect(const SubString: WideString; hWnd: SYSINT; 
                                         WindowType: TliWindowTypes; SearchType: TliSearchTypes; 
                                         SearchMiddle: WordBool): Smallint; safecall;
    function GetTypes(var StartResults: SearchResults; SearchType: TliSearchTypes; Sort: WordBool): SearchResults; safecall;
    function GetTypesDirect(hWnd: SYSINT; WindowType: TliWindowTypes; SearchType: TliSearchTypes): Smallint; safecall;
    function GetMembers(SearchData: Integer; ShowUnderscore: WordBool): SearchResults; safecall;
    function GetMembersDirect(SearchData: Integer; hWnd: SYSINT; WindowType: TliWindowTypes; 
                              ItemDataType: TliItemDataTypes; ShowUnderscore: WordBool): Smallint; safecall;
    procedure SetMemberFilters(FuncFilter: FuncFlags; VarFilter: VarFlags); safecall;
    function MakeSearchData(const TypeInfoName: WideString; SearchType: TliSearchTypes): Integer; safecall;
    function Get_TypeInfos: TypeInfos; safecall;
    function Get_Records: Records; safecall;
    function Get_IntrinsicAliases: IntrinsicAliases; safecall;
    function Get_CustomDataCollection: CustomDataCollection; safecall;
    function GetMemberInfo(SearchData: Integer; InvokeKinds: InvokeKinds; MemberId: Integer; 
                           const MemberName: WideString): MemberInfo; safecall;
    function Get_Unions: Unions; safecall;
    function AddTypes(var TypeInfoNumbers: PSafeArray; var StartResults: SearchResults; 
                      SearchType: TliSearchTypes; Sort: WordBool): SearchResults; safecall;
    function AddTypesDirect(var TypeInfoNumbers: PSafeArray; hWnd: SYSINT; 
                            WindowType: TliWindowTypes; SearchType: TliSearchTypes): Smallint; safecall;
    procedure FreeSearchCriteria; safecall;
    procedure Register(const HelpDir: WideString); safecall;
    procedure UnRegister; safecall;
    function Get_HelpString(LCID: Integer): WideString; safecall;
    function Get_AppObjString: WideString; safecall;
    function Get_LibNum: Smallint; safecall;
    function GetMembersWithSubStringEx(const SubString: WideString; 
                                       var InvokeGroupings: PSafeArray; 
                                       var StartResults: SearchResults; SearchType: TliSearchTypes; 
                                       SearchMiddle: WordBool; Sort: WordBool; 
                                       ShowUnderscore: WordBool): SearchResults; safecall;
    function GetTypesWithMemberEx(const MemberName: WideString; InvokeKind: InvokeKinds; 
                                  var StartResults: SearchResults; SearchType: TliSearchTypes; 
                                  Sort: WordBool; ShowUnderscore: WordBool): SearchResults; safecall;
    function Get_ITypeLib: IUnknown; safecall;
    procedure _Set_ITypeLib(const retVal: IUnknown); safecall;
    function Get_HelpStringDll(LCID: Integer): WideString; safecall;
    function Get_HelpStringContext: Integer; safecall;
    function Get_BestEquivalentType(const TypeInfoName: WideString): WideString; safecall;
    property ContainingFile: WideString read Get_ContainingFile write Set_ContainingFile;
    property Name: WideString read Get_Name;
    property _OldHelpString: HResult read Get__OldHelpString;
    property HelpContext: Integer read Get_HelpContext;
    property HelpFile: WideString read Get_HelpFile;
    property GUID: WideString read Get_GUID;
    property LCID: Integer read Get_LCID;
    property _OldSysKind: HResult read Get__OldSysKind;
    property MajorVersion: Smallint read Get_MajorVersion;
    property MinorVersion: Smallint read Get_MinorVersion;
    property AttributeMask: Smallint read Get_AttributeMask;
    property AttributeStrings[out AttributeArray: PSafeArray]: Smallint read Get_AttributeStrings;
    property CoClasses: CoClasses read Get_CoClasses;
    property Interfaces: Interfaces read Get_Interfaces;
    property Constants: Constants read Get_Constants;
    property Declarations: Declarations read Get_Declarations;
    property TypeInfoCount: Smallint read Get_TypeInfoCount;
    property _OldGetTypeKind: HResult read Get__OldGetTypeKind;
    property GetTypeInfo[var Index: OleVariant]: TypeInfo read Get_GetTypeInfo;
    property GetTypeInfoNumber[const Name: WideString]: Smallint read Get_GetTypeInfoNumber;
    property AppObjString: WideString read Get_AppObjString write Set_AppObjString;
    property LibNum: Smallint read Get_LibNum write Set_LibNum;
    property ShowLibName: WordBool read Get_ShowLibName write Set_ShowLibName;
    property GetTypeKind[TypeInfoNumber: Smallint]: TypeKinds read Get_GetTypeKind;
    property SysKind: SysKinds read Get_SysKind;
    property SearchDefault: TliSearchTypes read Get_SearchDefault write Set_SearchDefault;
    property TypeInfos: TypeInfos read Get_TypeInfos;
    property Records: Records read Get_Records;
    property IntrinsicAliases: IntrinsicAliases read Get_IntrinsicAliases;
    property CustomDataCollection: CustomDataCollection read Get_CustomDataCollection;
    property Unions: Unions read Get_Unions;
    property HelpString[LCID: Integer]: WideString read Get_HelpString;
    property ITypeLib: IUnknown read Get_ITypeLib write _Set_ITypeLib;
    property HelpStringDll[LCID: Integer]: WideString read Get_HelpStringDll;
    property HelpStringContext: Integer read Get_HelpStringContext;
    property BestEquivalentType[const TypeInfoName: WideString]: WideString read Get_BestEquivalentType;
  end;

// *********************************************************************//
// DispIntf:  _TypeLibInfoDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217745-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  _TypeLibInfoDisp = dispinterface
    ['{8B217745-717D-11CE-AB5B-D41203C10000}']
    function Me: TypeLibInfo; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    property ContainingFile: WideString dispid 1610743810;
    procedure LoadRegTypeLib(const TypeLibGuid: WideString; MajorVersion: Smallint; 
                             MinorVersion: Smallint; LCID: Integer); dispid 1610743812;
    property Name: WideString readonly dispid 0;
    property _OldHelpString: HResult readonly dispid 1610743814;
    property HelpContext: Integer readonly dispid 1610743815;
    property HelpFile: WideString readonly dispid 1610743816;
    property GUID: WideString readonly dispid 1610743817;
    property LCID: Integer readonly dispid 1610743818;
    property _OldSysKind: HResult readonly dispid 1610743819;
    property MajorVersion: Smallint readonly dispid 1610743820;
    property MinorVersion: Smallint readonly dispid 1610743821;
    property AttributeMask: Smallint readonly dispid 1610743822;
    property AttributeStrings[out AttributeArray: {??PSafeArray}OleVariant]: Smallint readonly dispid 1610743823;
    property CoClasses: CoClasses readonly dispid 1610743824;
    property Interfaces: Interfaces readonly dispid 1610743825;
    property Constants: Constants readonly dispid 1610743826;
    property Declarations: Declarations readonly dispid 1610743827;
    property TypeInfoCount: Smallint readonly dispid 1610743828;
    property _OldGetTypeKind: HResult readonly dispid 1610743829;
    property GetTypeInfo[var Index: OleVariant]: TypeInfo readonly dispid 1610743830;
    property GetTypeInfoNumber[const Name: WideString]: Smallint readonly dispid 1610743831;
    function IsSameLibrary(const CheckLib: TypeLibInfo): WordBool; dispid 1610743832;
    procedure _OldResetSearchCriteria; dispid 1610743833;
    procedure _OldGetTypesWithMember; dispid 1610743834;
    procedure _OldGetMembersWithSubString; dispid 1610743835;
    procedure _OldGetTypesWithSubString; dispid 1610743836;
    procedure _OldCaseTypeName; dispid 1610743837;
    procedure _OldCaseMemberName; dispid 1610743838;
    procedure _OldFillTypesList; dispid 1610743839;
    procedure _OldFillTypesCombo; dispid 1610743840;
    procedure _OldFillMemberList; dispid 1610743841;
    procedure _OldAddClassTypeToList; dispid 1610743842;
    property AppObjString: WideString dispid 1610743843;
    property LibNum: Smallint dispid 1610743844;
    property ShowLibName: WordBool dispid 1610743845;
    property GetTypeKind[TypeInfoNumber: Smallint]: TypeKinds readonly dispid 1610743848;
    property SysKind: SysKinds readonly dispid 1610743849;
    property SearchDefault: TliSearchTypes dispid 1610743850;
    function CaseTypeName(var bstrName: WideString; SearchType: TliSearchTypes): TliSearchTypes; dispid 1610743852;
    function CaseMemberName(var bstrName: WideString; SearchType: TliSearchTypes): WordBool; dispid 1610743853;
    procedure ResetSearchCriteria(TypeFilter: TypeFlags; IncludeEmptyTypes: WordBool; 
                                  ShowUnderscore: WordBool); dispid 1610743854;
    function GetTypesWithMember(const MemberName: WideString; var StartResults: SearchResults; 
                                SearchType: TliSearchTypes; Sort: WordBool; ShowUnderscore: WordBool): SearchResults; dispid 1610743855;
    function GetTypesWithMemberDirect(const MemberName: WideString; hWnd: SYSINT; 
                                      WindowType: TliWindowTypes; SearchType: TliSearchTypes; 
                                      ShowUnderscore: WordBool): Smallint; dispid 1610743856;
    function GetMembersWithSubString(const SubString: WideString; var StartResults: SearchResults; 
                                     SearchType: TliSearchTypes; SearchMiddle: WordBool; 
                                     var Helper: SearchHelper; Sort: WordBool; 
                                     ShowUnderscore: WordBool): SearchResults; dispid 1610743857;
    function GetMembersWithSubStringDirect(const SubString: WideString; hWnd: SYSINT; 
                                           WindowType: TliWindowTypes; SearchType: TliSearchTypes; 
                                           SearchMiddle: WordBool; var Helper: SearchHelper; 
                                           ShowUnderscore: WordBool): Smallint; dispid 1610743858;
    function GetTypesWithSubString(const SubString: WideString; var StartResults: SearchResults; 
                                   SearchType: TliSearchTypes; SearchMiddle: WordBool; 
                                   Sort: WordBool): SearchResults; dispid 1610743859;
    function GetTypesWithSubStringDirect(const SubString: WideString; hWnd: SYSINT; 
                                         WindowType: TliWindowTypes; SearchType: TliSearchTypes; 
                                         SearchMiddle: WordBool): Smallint; dispid 1610743860;
    function GetTypes(var StartResults: SearchResults; SearchType: TliSearchTypes; Sort: WordBool): SearchResults; dispid 1610743861;
    function GetTypesDirect(hWnd: SYSINT; WindowType: TliWindowTypes; SearchType: TliSearchTypes): Smallint; dispid 1610743862;
    function GetMembers(SearchData: Integer; ShowUnderscore: WordBool): SearchResults; dispid 1610743863;
    function GetMembersDirect(SearchData: Integer; hWnd: SYSINT; WindowType: TliWindowTypes; 
                              ItemDataType: TliItemDataTypes; ShowUnderscore: WordBool): Smallint; dispid 1610743864;
    procedure SetMemberFilters(FuncFilter: FuncFlags; VarFilter: VarFlags); dispid 1610743865;
    function MakeSearchData(const TypeInfoName: WideString; SearchType: TliSearchTypes): Integer; dispid 1610743866;
    property TypeInfos: TypeInfos readonly dispid 1610743867;
    property Records: Records readonly dispid 1610743868;
    property IntrinsicAliases: IntrinsicAliases readonly dispid 1610743869;
    property CustomDataCollection: CustomDataCollection readonly dispid 1610743870;
    function GetMemberInfo(SearchData: Integer; InvokeKinds: InvokeKinds; MemberId: Integer; 
                           const MemberName: WideString): MemberInfo; dispid 1610743871;
    property Unions: Unions readonly dispid 1610743872;
    function AddTypes(var TypeInfoNumbers: {??PSafeArray}OleVariant; 
                      var StartResults: SearchResults; SearchType: TliSearchTypes; Sort: WordBool): SearchResults; dispid 1610743873;
    function AddTypesDirect(var TypeInfoNumbers: {??PSafeArray}OleVariant; hWnd: SYSINT; 
                            WindowType: TliWindowTypes; SearchType: TliSearchTypes): Smallint; dispid 1610743874;
    procedure FreeSearchCriteria; dispid 1610743875;
    procedure Register(const HelpDir: WideString); dispid 1610743876;
    procedure UnRegister; dispid 1610743877;
    property HelpString[LCID: Integer]: WideString readonly dispid 1610743878;
    function GetMembersWithSubStringEx(const SubString: WideString; 
                                       var InvokeGroupings: {??PSafeArray}OleVariant; 
                                       var StartResults: SearchResults; SearchType: TliSearchTypes; 
                                       SearchMiddle: WordBool; Sort: WordBool; 
                                       ShowUnderscore: WordBool): SearchResults; dispid 1610743881;
    function GetTypesWithMemberEx(const MemberName: WideString; InvokeKind: InvokeKinds; 
                                  var StartResults: SearchResults; SearchType: TliSearchTypes; 
                                  Sort: WordBool; ShowUnderscore: WordBool): SearchResults; dispid 1610743882;
    property ITypeLib: IUnknown dispid 1610743883;
    property HelpStringDll[LCID: Integer]: WideString readonly dispid 1610743885;
    property HelpStringContext: Integer readonly dispid 1610743886;
    property BestEquivalentType[const TypeInfoName: WideString]: WideString readonly dispid 1610743887;
  end;

// *********************************************************************//
// Interface: SearchItem
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217756-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  SearchItem = interface(IDispatch)
    ['{8B217756-717D-11CE-AB5B-D41203C10000}']
    function Me: SearchItem; safecall;
    procedure _placeholder_destructor; safecall;
    function Get_Name: WideString; safecall;
    function Get_SearchData: Integer; safecall;
    function Get_TypeInfoNumber: Smallint; safecall;
    function Get__OldLibNum: Byte; safecall;
    function Get_SearchType: TliSearchTypes; safecall;
    function Get_MemberId: Integer; safecall;
    function Get_InvokeKinds: InvokeKinds; safecall;
    function Get_NamePtrW: Integer; safecall;
    function Get_LibNum: Smallint; safecall;
    function Get_Constant: WordBool; safecall;
    function Get_Hidden: WordBool; safecall;
    function Get_InvokeGroup: Smallint; safecall;
    property Name: WideString read Get_Name;
    property SearchData: Integer read Get_SearchData;
    property TypeInfoNumber: Smallint read Get_TypeInfoNumber;
    property _OldLibNum: Byte read Get__OldLibNum;
    property SearchType: TliSearchTypes read Get_SearchType;
    property MemberId: Integer read Get_MemberId;
    property InvokeKinds: InvokeKinds read Get_InvokeKinds;
    property NamePtrW: Integer read Get_NamePtrW;
    property LibNum: Smallint read Get_LibNum;
    property Constant: WordBool read Get_Constant;
    property Hidden: WordBool read Get_Hidden;
    property InvokeGroup: Smallint read Get_InvokeGroup;
  end;

// *********************************************************************//
// DispIntf:  SearchItemDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217756-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  SearchItemDisp = dispinterface
    ['{8B217756-717D-11CE-AB5B-D41203C10000}']
    function Me: SearchItem; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    property Name: WideString readonly dispid 0;
    property SearchData: Integer readonly dispid 1610743811;
    property TypeInfoNumber: Smallint readonly dispid 1610743812;
    property _OldLibNum: Byte readonly dispid 1610743813;
    property SearchType: TliSearchTypes readonly dispid 1610743814;
    property MemberId: Integer readonly dispid 1610743815;
    property InvokeKinds: InvokeKinds readonly dispid 1610743816;
    property NamePtrW: Integer readonly dispid 1610743817;
    property LibNum: Smallint readonly dispid 1610743818;
    property Constant: WordBool readonly dispid 1610743819;
    property Hidden: WordBool readonly dispid 1610743820;
    property InvokeGroup: Smallint readonly dispid 1610743821;
  end;

// *********************************************************************//
// Interface: SearchResults
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217757-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  SearchResults = interface(IDispatch)
    ['{8B217757-717D-11CE-AB5B-D41203C10000}']
    function Me: SearchResults; safecall;
    procedure _placeholder_destructor; safecall;
    function _NewEnum: IUnknown; safecall;
    function Get__OldItem(Index: Smallint): SearchItem; safecall;
    function Get__OldCount: Smallint; safecall;
    function Get_Sorted: WordBool; safecall;
    procedure Sort(const CustomSort: CustomSort); safecall;
    function Filter(const CustomFilter: CustomFilter; var AppendExtractedTo: SearchResults; 
                    const StartAfter: SearchItem): SearchResults; safecall;
    function Get_Item(Index: Integer): SearchItem; safecall;
    function Get_Count: Integer; safecall;
    function LocateSorted(const CustomSort: CustomSort): Integer; safecall;
    function Locate(const SearchString: WideString; const CustomSort: CustomSort; 
                    StartAfter: Integer): Integer; safecall;
    property _OldItem[Index: Smallint]: SearchItem read Get__OldItem;
    property _OldCount: Smallint read Get__OldCount;
    property Sorted: WordBool read Get_Sorted;
    property Item[Index: Integer]: SearchItem read Get_Item; default;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  SearchResultsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217757-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  SearchResultsDisp = dispinterface
    ['{8B217757-717D-11CE-AB5B-D41203C10000}']
    function Me: SearchResults; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    function _NewEnum: IUnknown; dispid -4;
    property _OldItem[Index: Smallint]: SearchItem readonly dispid 1610743811;
    property _OldCount: Smallint readonly dispid 1610743812;
    property Sorted: WordBool readonly dispid 1610743813;
    procedure Sort(const CustomSort: CustomSort); dispid 1610743814;
    function Filter(const CustomFilter: CustomFilter; var AppendExtractedTo: SearchResults; 
                    const StartAfter: SearchItem): SearchResults; dispid 1610743815;
    property Item[Index: Integer]: SearchItem readonly dispid 0; default;
    property Count: Integer readonly dispid 1610743817;
    function LocateSorted(const CustomSort: CustomSort): Integer; dispid 1610743818;
    function Locate(const SearchString: WideString; const CustomSort: CustomSort; 
                    StartAfter: Integer): Integer; dispid 1610743819;
  end;

// *********************************************************************//
// Interface: ListBoxNotification
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217758-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  ListBoxNotification = interface(IDispatch)
    ['{8B217758-717D-11CE-AB5B-D41203C10000}']
    procedure OnAddString(lpstr: Integer; fUnicode: WordBool); safecall;
  end;

// *********************************************************************//
// DispIntf:  ListBoxNotificationDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217758-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  ListBoxNotificationDisp = dispinterface
    ['{8B217758-717D-11CE-AB5B-D41203C10000}']
    procedure OnAddString(lpstr: Integer; fUnicode: WordBool); dispid 1610743808;
  end;

// *********************************************************************//
// Interface: CustomSort
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B21775F-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  CustomSort = interface(IDispatch)
    ['{8B21775F-717D-11CE-AB5B-D41203C10000}']
    procedure Compare(const Item1: SearchItem; const Item2: SearchItem; var Compare: Integer); safecall;
  end;

// *********************************************************************//
// DispIntf:  CustomSortDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B21775F-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  CustomSortDisp = dispinterface
    ['{8B21775F-717D-11CE-AB5B-D41203C10000}']
    procedure Compare(const Item1: SearchItem; const Item2: SearchItem; var Compare: Integer); dispid 1610743808;
  end;

// *********************************************************************//
// Interface: CustomFilter
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217760-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  CustomFilter = interface(IDispatch)
    ['{8B217760-717D-11CE-AB5B-D41203C10000}']
    procedure Visit(const Item: SearchItem; var Action: TliCustomFilterAction); safecall;
  end;

// *********************************************************************//
// DispIntf:  CustomFilterDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217760-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  CustomFilterDisp = dispinterface
    ['{8B217760-717D-11CE-AB5B-D41203C10000}']
    procedure Visit(const Item: SearchItem; var Action: TliCustomFilterAction); dispid 1610743808;
  end;

// *********************************************************************//
// Interface: _TLIApplication
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B21775D-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  _TLIApplication = interface(IDispatch)
    ['{8B21775D-717D-11CE-AB5B-D41203C10000}']
    function Me: TLIApplication; safecall;
    procedure _placeholder_destructor; safecall;
    function TypeLibInfoFromFile(const FileName: WideString): TypeLibInfo; safecall;
    function InterfaceInfoFromObject(const Object_: IDispatch): InterfaceInfo; safecall;
    function Get_ListBoxNotification: ListBoxNotification; safecall;
    procedure _Set_ListBoxNotification(const retVal: ListBoxNotification); safecall;
    function Get_ResolveAliases: WordBool; safecall;
    procedure Set_ResolveAliases(retVal: WordBool); safecall;
    function InvokeHook(const Object_: IDispatch; NameOrID: OleVariant; InvokeKind: InvokeKinds; 
                        var ReverseArgList: PSafeArray): OleVariant; safecall;
    function InvokeHookArray(const Object_: IDispatch; NameOrID: OleVariant; 
                             InvokeKind: InvokeKinds; var ReverseArgList: PSafeArray): OleVariant; safecall;
    procedure InvokeHookSub(const Object_: IDispatch; NameOrID: OleVariant; 
                            InvokeKind: InvokeKinds; var ReverseArgList: PSafeArray); safecall;
    procedure InvokeHookArraySub(const Object_: IDispatch; NameOrID: OleVariant; 
                                 InvokeKind: InvokeKinds; var ReverseArgList: PSafeArray); safecall;
    function ClassInfoFromObject(const Object_: IUnknown): TypeInfo; safecall;
    function InvokeID(const Object_: IDispatch; const Name: WideString): Integer; safecall;
    function Get_InvokeLCID: Integer; safecall;
    procedure Set_InvokeLCID(retVal: Integer); safecall;
    function TypeInfoFromITypeInfo(const ptinfo: IUnknown): TypeInfo; safecall;
    function TypeLibInfoFromITypeLib(const pITypeLib: IUnknown): TypeLibInfo; safecall;
    function TypeLibInfoFromRegistry(const TypeLibGuid: WideString; MajorVersion: Smallint; 
                                     MinorVersion: Smallint; LCID: Integer): TypeLibInfo; safecall;
    function TypeInfoFromRecordVariant(var RecordVariant: OleVariant): TypeInfo; safecall;
    function Get_RecordField(var RecordVariant: OleVariant; var FieldName: WideString): OleVariant; safecall;
    procedure Set_RecordField(var RecordVariant: OleVariant; var FieldName: WideString; 
                              var retVal: OleVariant); safecall;
    procedure _Set_RecordField(var RecordVariant: OleVariant; var FieldName: WideString; 
                               var retVal: OleVariant); safecall;
    property ListBoxNotification: ListBoxNotification read Get_ListBoxNotification write _Set_ListBoxNotification;
    property ResolveAliases: WordBool read Get_ResolveAliases write Set_ResolveAliases;
    property InvokeLCID: Integer read Get_InvokeLCID write Set_InvokeLCID;
  end;

// *********************************************************************//
// DispIntf:  _TLIApplicationDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B21775D-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  _TLIApplicationDisp = dispinterface
    ['{8B21775D-717D-11CE-AB5B-D41203C10000}']
    function Me: TLIApplication; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    function TypeLibInfoFromFile(const FileName: WideString): TypeLibInfo; dispid 1610743810;
    function InterfaceInfoFromObject(const Object_: IDispatch): InterfaceInfo; dispid 1610743811;
    property ListBoxNotification: ListBoxNotification dispid 1610743812;
    property ResolveAliases: WordBool dispid 1610743814;
    function InvokeHook(const Object_: IDispatch; NameOrID: OleVariant; InvokeKind: InvokeKinds; 
                        var ReverseArgList: {??PSafeArray}OleVariant): OleVariant; dispid 1610743816;
    function InvokeHookArray(const Object_: IDispatch; NameOrID: OleVariant; 
                             InvokeKind: InvokeKinds; var ReverseArgList: {??PSafeArray}OleVariant): OleVariant; dispid 1610743817;
    procedure InvokeHookSub(const Object_: IDispatch; NameOrID: OleVariant; 
                            InvokeKind: InvokeKinds; var ReverseArgList: {??PSafeArray}OleVariant); dispid 1610743818;
    procedure InvokeHookArraySub(const Object_: IDispatch; NameOrID: OleVariant; 
                                 InvokeKind: InvokeKinds; 
                                 var ReverseArgList: {??PSafeArray}OleVariant); dispid 1610743819;
    function ClassInfoFromObject(const Object_: IUnknown): TypeInfo; dispid 1610743820;
    function InvokeID(const Object_: IDispatch; const Name: WideString): Integer; dispid 1610743821;
    property InvokeLCID: Integer dispid 1610743822;
    function TypeInfoFromITypeInfo(const ptinfo: IUnknown): TypeInfo; dispid 1610743824;
    function TypeLibInfoFromITypeLib(const pITypeLib: IUnknown): TypeLibInfo; dispid 1610743825;
    function TypeLibInfoFromRegistry(const TypeLibGuid: WideString; MajorVersion: Smallint; 
                                     MinorVersion: Smallint; LCID: Integer): TypeLibInfo; dispid 1610743826;
    function TypeInfoFromRecordVariant(var RecordVariant: OleVariant): TypeInfo; dispid 1610743827;
    function RecordField(var RecordVariant: OleVariant; var FieldName: WideString): OleVariant; dispid 1610743828;
  end;

// *********************************************************************//
// Interface: TypeInfos
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B21775A-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  TypeInfos = interface(_BaseTypeInfos)
    ['{8B21775A-717D-11CE-AB5B-D41203C10000}']
    function Get_Item(Index: Smallint): TypeInfo; safecall;
    function Get_IndexedItem(TypeInfoNumber: Smallint): TypeInfo; safecall;
    function Get_NamedItem(var TypeInfoName: WideString): TypeInfo; safecall;
    property Item[Index: Smallint]: TypeInfo read Get_Item; default;
    property IndexedItem[TypeInfoNumber: Smallint]: TypeInfo read Get_IndexedItem;
    property NamedItem[var TypeInfoName: WideString]: TypeInfo read Get_NamedItem;
  end;

// *********************************************************************//
// DispIntf:  TypeInfosDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B21775A-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  TypeInfosDisp = dispinterface
    ['{8B21775A-717D-11CE-AB5B-D41203C10000}']
    property Item[Index: Smallint]: TypeInfo readonly dispid 0; default;
    property IndexedItem[TypeInfoNumber: Smallint]: TypeInfo readonly dispid 1610809345;
    property NamedItem[var TypeInfoName: WideString]: TypeInfo readonly dispid 1610809346;
    function Me: TypeInfos; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    function _NewEnum: IUnknown; dispid -4;
    property Count: Smallint readonly dispid 1610743811;
  end;

// *********************************************************************//
// Interface: RecordInfo
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B21775B-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  RecordInfo = interface(ConstantInfo)
    ['{8B21775B-717D-11CE-AB5B-D41203C10000}']
  end;

// *********************************************************************//
// DispIntf:  RecordInfoDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B21775B-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  RecordInfoDisp = dispinterface
    ['{8B21775B-717D-11CE-AB5B-D41203C10000}']
    function Me: TypeInfo; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    property Name: WideString readonly dispid 0;
    property GUID: WideString readonly dispid 1610743811;
    property _OldHelpString: HResult readonly dispid 1610743812;
    property HelpContext: Integer readonly dispid 1610743813;
    property HelpFile: WideString readonly dispid 1610743814;
    property AttributeMask: Smallint readonly dispid 1610743815;
    property AttributeStrings[out AttributeArray: {??PSafeArray}OleVariant]: Smallint readonly dispid 1610743816;
    property _OldTypeKind: HResult readonly dispid 1610743817;
    property TypeKindString: WideString readonly dispid 1610743818;
    property TypeInfoNumber: Smallint readonly dispid 1610743819;
    procedure _placeholder_VTableInterface; dispid 1610743820;
    property GetMember[Index: OleVariant]: MemberInfo readonly dispid 1610743821;
    property Members: Members readonly dispid 1610743822;
    property Parent: TypeLibInfo readonly dispid 1610743823;
    procedure _ImpliedInterfaces; dispid 1610743824;
    procedure _DefaultInterface; dispid 1610743825;
    procedure _DefaultEventInterface; dispid 1610743826;
    property TypeKind: TypeKinds readonly dispid 1610743827;
    property ResolvedType: VarTypeInfo readonly dispid 1610743828;
    property CustomDataCollection: CustomDataCollection readonly dispid 1610743829;
    property HelpString[LCID: Integer]: WideString readonly dispid 1610743830;
    property ITypeInfo: IUnknown readonly dispid 1610743831;
    property MajorVersion: Smallint readonly dispid 1610743832;
    property MinorVersion: Smallint readonly dispid 1610743833;
    property HelpStringDll[LCID: Integer]: WideString readonly dispid 1610743834;
    property HelpStringContext: Integer readonly dispid 1610743835;
  end;

// *********************************************************************//
// Interface: Records
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B21775C-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  Records = interface(_BaseTypeInfos)
    ['{8B21775C-717D-11CE-AB5B-D41203C10000}']
    function Get_Item(Index: Smallint): RecordInfo; safecall;
    function Get_IndexedItem(TypeInfoNumber: Smallint): RecordInfo; safecall;
    function Get_NamedItem(var TypeInfoName: WideString): RecordInfo; safecall;
    property Item[Index: Smallint]: RecordInfo read Get_Item; default;
    property IndexedItem[TypeInfoNumber: Smallint]: RecordInfo read Get_IndexedItem;
    property NamedItem[var TypeInfoName: WideString]: RecordInfo read Get_NamedItem;
  end;

// *********************************************************************//
// DispIntf:  RecordsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B21775C-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  RecordsDisp = dispinterface
    ['{8B21775C-717D-11CE-AB5B-D41203C10000}']
    property Item[Index: Smallint]: RecordInfo readonly dispid 0; default;
    property IndexedItem[TypeInfoNumber: Smallint]: RecordInfo readonly dispid 1610809345;
    property NamedItem[var TypeInfoName: WideString]: RecordInfo readonly dispid 1610809346;
    function Me: TypeInfos; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    function _NewEnum: IUnknown; dispid -4;
    property Count: Smallint readonly dispid 1610743811;
  end;

// *********************************************************************//
// Interface: IntrinsicAliasInfo
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217761-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  IntrinsicAliasInfo = interface(IDispatch)
    ['{8B217761-717D-11CE-AB5B-D41203C10000}']
    function Me: IntrinsicAliasInfo; safecall;
    procedure _placeholder_destructor; safecall;
    function Get_Name: WideString; safecall;
    function Get_GUID: WideString; safecall;
    function Get__OldHelpString: HResult; safecall;
    function Get_HelpContext: Integer; safecall;
    function Get_HelpFile: WideString; safecall;
    function Get_AttributeMask: Smallint; safecall;
    function Get_AttributeStrings(out AttributeArray: PSafeArray): Smallint; safecall;
    function Get__OldTypeKind: HResult; safecall;
    function Get_TypeKindString: WideString; safecall;
    function Get_TypeInfoNumber: Smallint; safecall;
    procedure _placeholder_VTableInterface; safecall;
    procedure _placeholder_GetMember; safecall;
    procedure _placeholder_Members; safecall;
    function Get_Parent: TypeLibInfo; safecall;
    procedure _ImpliedInterfaces; safecall;
    procedure _DefaultInterface; safecall;
    procedure _DefaultEventInterface; safecall;
    function Get_TypeKind: TypeKinds; safecall;
    function Get_ResolvedType: VarTypeInfo; safecall;
    function Get_CustomDataCollection: CustomDataCollection; safecall;
    function Get_HelpString(LCID: Integer): WideString; safecall;
    function Get_ITypeInfo: IUnknown; safecall;
    function Get_MajorVersion: Smallint; safecall;
    function Get_MinorVersion: Smallint; safecall;
    function Get_HelpStringDll(LCID: Integer): WideString; safecall;
    function Get_HelpStringContext: Integer; safecall;
    property Name: WideString read Get_Name;
    property GUID: WideString read Get_GUID;
    property _OldHelpString: HResult read Get__OldHelpString;
    property HelpContext: Integer read Get_HelpContext;
    property HelpFile: WideString read Get_HelpFile;
    property AttributeMask: Smallint read Get_AttributeMask;
    property AttributeStrings[out AttributeArray: PSafeArray]: Smallint read Get_AttributeStrings;
    property _OldTypeKind: HResult read Get__OldTypeKind;
    property TypeKindString: WideString read Get_TypeKindString;
    property TypeInfoNumber: Smallint read Get_TypeInfoNumber;
    property Parent: TypeLibInfo read Get_Parent;
    property TypeKind: TypeKinds read Get_TypeKind;
    property ResolvedType: VarTypeInfo read Get_ResolvedType;
    property CustomDataCollection: CustomDataCollection read Get_CustomDataCollection;
    property HelpString[LCID: Integer]: WideString read Get_HelpString;
    property ITypeInfo: IUnknown read Get_ITypeInfo;
    property MajorVersion: Smallint read Get_MajorVersion;
    property MinorVersion: Smallint read Get_MinorVersion;
    property HelpStringDll[LCID: Integer]: WideString read Get_HelpStringDll;
    property HelpStringContext: Integer read Get_HelpStringContext;
  end;

// *********************************************************************//
// DispIntf:  IntrinsicAliasInfoDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217761-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  IntrinsicAliasInfoDisp = dispinterface
    ['{8B217761-717D-11CE-AB5B-D41203C10000}']
    function Me: IntrinsicAliasInfo; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    property Name: WideString readonly dispid 0;
    property GUID: WideString readonly dispid 1610743811;
    property _OldHelpString: HResult readonly dispid 1610743812;
    property HelpContext: Integer readonly dispid 1610743813;
    property HelpFile: WideString readonly dispid 1610743814;
    property AttributeMask: Smallint readonly dispid 1610743815;
    property AttributeStrings[out AttributeArray: {??PSafeArray}OleVariant]: Smallint readonly dispid 1610743816;
    property _OldTypeKind: HResult readonly dispid 1610743817;
    property TypeKindString: WideString readonly dispid 1610743818;
    property TypeInfoNumber: Smallint readonly dispid 1610743819;
    procedure _placeholder_VTableInterface; dispid 1610743820;
    procedure _placeholder_GetMember; dispid 1610743821;
    procedure _placeholder_Members; dispid 1610743822;
    property Parent: TypeLibInfo readonly dispid 1610743823;
    procedure _ImpliedInterfaces; dispid 1610743824;
    procedure _DefaultInterface; dispid 1610743825;
    procedure _DefaultEventInterface; dispid 1610743826;
    property TypeKind: TypeKinds readonly dispid 1610743827;
    property ResolvedType: VarTypeInfo readonly dispid 1610743828;
    property CustomDataCollection: CustomDataCollection readonly dispid 1610743829;
    property HelpString[LCID: Integer]: WideString readonly dispid 1610743830;
    property ITypeInfo: IUnknown readonly dispid 1610743831;
    property MajorVersion: Smallint readonly dispid 1610743832;
    property MinorVersion: Smallint readonly dispid 1610743833;
    property HelpStringDll[LCID: Integer]: WideString readonly dispid 1610743834;
    property HelpStringContext: Integer readonly dispid 1610743835;
  end;

// *********************************************************************//
// Interface: IntrinsicAliases
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217762-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  IntrinsicAliases = interface(_BaseTypeInfos)
    ['{8B217762-717D-11CE-AB5B-D41203C10000}']
    function Get_Item(Index: Smallint): IntrinsicAliasInfo; safecall;
    function Get_IndexedItem(TypeInfoNumber: Smallint): IntrinsicAliasInfo; safecall;
    function Get_NamedItem(var TypeInfoName: WideString): IntrinsicAliasInfo; safecall;
    property Item[Index: Smallint]: IntrinsicAliasInfo read Get_Item; default;
    property IndexedItem[TypeInfoNumber: Smallint]: IntrinsicAliasInfo read Get_IndexedItem;
    property NamedItem[var TypeInfoName: WideString]: IntrinsicAliasInfo read Get_NamedItem;
  end;

// *********************************************************************//
// DispIntf:  IntrinsicAliasesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217762-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  IntrinsicAliasesDisp = dispinterface
    ['{8B217762-717D-11CE-AB5B-D41203C10000}']
    property Item[Index: Smallint]: IntrinsicAliasInfo readonly dispid 0; default;
    property IndexedItem[TypeInfoNumber: Smallint]: IntrinsicAliasInfo readonly dispid 1610809345;
    property NamedItem[var TypeInfoName: WideString]: IntrinsicAliasInfo readonly dispid 1610809346;
    function Me: TypeInfos; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    function _NewEnum: IUnknown; dispid -4;
    property Count: Smallint readonly dispid 1610743811;
  end;

// *********************************************************************//
// Interface: CustomData
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217763-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  CustomData = interface(IDispatch)
    ['{8B217763-717D-11CE-AB5B-D41203C10000}']
    function Me: CustomData; safecall;
    procedure _placeholder_destructor; safecall;
    function Get_GUID: WideString; safecall;
    function Get_Value: OleVariant; safecall;
    property GUID: WideString read Get_GUID;
    property Value: OleVariant read Get_Value;
  end;

// *********************************************************************//
// DispIntf:  CustomDataDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217763-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  CustomDataDisp = dispinterface
    ['{8B217763-717D-11CE-AB5B-D41203C10000}']
    function Me: CustomData; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    property GUID: WideString readonly dispid 1610743810;
    property Value: OleVariant readonly dispid 1610743811;
  end;

// *********************************************************************//
// Interface: CustomDataCollection
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217764-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  CustomDataCollection = interface(IDispatch)
    ['{8B217764-717D-11CE-AB5B-D41203C10000}']
    function Me: CustomDataCollection; safecall;
    procedure _placeholder_destructor; safecall;
    function _NewEnum: IUnknown; safecall;
    function Get_Item(Index: Smallint): CustomData; safecall;
    function Get_Count: Smallint; safecall;
    property Item[Index: Smallint]: CustomData read Get_Item; default;
    property Count: Smallint read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  CustomDataCollectionDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217764-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  CustomDataCollectionDisp = dispinterface
    ['{8B217764-717D-11CE-AB5B-D41203C10000}']
    function Me: CustomDataCollection; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    function _NewEnum: IUnknown; dispid -4;
    property Item[Index: Smallint]: CustomData readonly dispid 0; default;
    property Count: Smallint readonly dispid 1610743812;
  end;

// *********************************************************************//
// Interface: UnionInfo
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217765-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  UnionInfo = interface(ConstantInfo)
    ['{8B217765-717D-11CE-AB5B-D41203C10000}']
  end;

// *********************************************************************//
// DispIntf:  UnionInfoDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217765-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  UnionInfoDisp = dispinterface
    ['{8B217765-717D-11CE-AB5B-D41203C10000}']
    function Me: TypeInfo; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    property Name: WideString readonly dispid 0;
    property GUID: WideString readonly dispid 1610743811;
    property _OldHelpString: HResult readonly dispid 1610743812;
    property HelpContext: Integer readonly dispid 1610743813;
    property HelpFile: WideString readonly dispid 1610743814;
    property AttributeMask: Smallint readonly dispid 1610743815;
    property AttributeStrings[out AttributeArray: {??PSafeArray}OleVariant]: Smallint readonly dispid 1610743816;
    property _OldTypeKind: HResult readonly dispid 1610743817;
    property TypeKindString: WideString readonly dispid 1610743818;
    property TypeInfoNumber: Smallint readonly dispid 1610743819;
    procedure _placeholder_VTableInterface; dispid 1610743820;
    property GetMember[Index: OleVariant]: MemberInfo readonly dispid 1610743821;
    property Members: Members readonly dispid 1610743822;
    property Parent: TypeLibInfo readonly dispid 1610743823;
    procedure _ImpliedInterfaces; dispid 1610743824;
    procedure _DefaultInterface; dispid 1610743825;
    procedure _DefaultEventInterface; dispid 1610743826;
    property TypeKind: TypeKinds readonly dispid 1610743827;
    property ResolvedType: VarTypeInfo readonly dispid 1610743828;
    property CustomDataCollection: CustomDataCollection readonly dispid 1610743829;
    property HelpString[LCID: Integer]: WideString readonly dispid 1610743830;
    property ITypeInfo: IUnknown readonly dispid 1610743831;
    property MajorVersion: Smallint readonly dispid 1610743832;
    property MinorVersion: Smallint readonly dispid 1610743833;
    property HelpStringDll[LCID: Integer]: WideString readonly dispid 1610743834;
    property HelpStringContext: Integer readonly dispid 1610743835;
  end;

// *********************************************************************//
// Interface: Unions
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217766-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  Unions = interface(_BaseTypeInfos)
    ['{8B217766-717D-11CE-AB5B-D41203C10000}']
    function Get_Item(Index: Smallint): UnionInfo; safecall;
    function Get_IndexedItem(TypeInfoNumber: Smallint): UnionInfo; safecall;
    function Get_NamedItem(var TypeInfoName: WideString): UnionInfo; safecall;
    property Item[Index: Smallint]: UnionInfo read Get_Item; default;
    property IndexedItem[TypeInfoNumber: Smallint]: UnionInfo read Get_IndexedItem;
    property NamedItem[var TypeInfoName: WideString]: UnionInfo read Get_NamedItem;
  end;

// *********************************************************************//
// DispIntf:  UnionsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {8B217766-717D-11CE-AB5B-D41203C10000}
// *********************************************************************//
  UnionsDisp = dispinterface
    ['{8B217766-717D-11CE-AB5B-D41203C10000}']
    property Item[Index: Smallint]: UnionInfo readonly dispid 0; default;
    property IndexedItem[TypeInfoNumber: Smallint]: UnionInfo readonly dispid 1610809345;
    property NamedItem[var TypeInfoName: WideString]: UnionInfo readonly dispid 1610809346;
    function Me: TypeInfos; dispid 1610743808;
    procedure _placeholder_destructor; dispid 1610743809;
    function _NewEnum: IUnknown; dispid -4;
    property Count: Smallint readonly dispid 1610743811;
  end;

// *********************************************************************//
// The Class CoSearchHelper provides a Create and CreateRemote method to          
// create instances of the default interface _SearchHelper exposed by              
// the CoClass SearchHelper. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoSearchHelper = class
    class function Create: _SearchHelper;
    class function CreateRemote(const MachineName: string): _SearchHelper;
  end;


// *********************************************************************//
// OLE Server Proxy class declaration
// Server Object    : TSearchHelper
// Help String      : Helper object for GetMembersWithSubString and multiple TypeLibs
// Default Interface: _SearchHelper
// Def. Intf. DISP? : No
// Event   Interface: 
// TypeFlags        : (2) CanCreate
// *********************************************************************//
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  TSearchHelperProperties= class;
{$ENDIF}
  TSearchHelper = class(TOleServer)
  private
    FIntf:        _SearchHelper;
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    FProps:       TSearchHelperProperties;
    function      GetServerProperties: TSearchHelperProperties;
{$ENDIF}
    function      GetDefaultInterface: _SearchHelper;
  protected
    procedure InitServerData; override;
    function Get_CheckHaveMatch(const Name: WideString): WordBool;
    function Get_Init(SysKind: SysKinds; LCID: Integer): HResult;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: _SearchHelper);
    procedure Disconnect; override;
    function Me: TypeLibInfo;
    procedure _OldInit;
    property DefaultInterface: _SearchHelper read GetDefaultInterface;
    property CheckHaveMatch[const Name: WideString]: WordBool read Get_CheckHaveMatch;
    property Init[SysKind: SysKinds; LCID: Integer]: HResult read Get_Init;
  published
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    property Server: TSearchHelperProperties read GetServerProperties;
{$ENDIF}
  end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
// *********************************************************************//
// OLE Server Properties Proxy Class
// Server Object    : TSearchHelper
// (This object is used by the IDE's Property Inspector to allow editing
//  of the properties of this server)
// *********************************************************************//
 TSearchHelperProperties = class(TPersistent)
  private
    FServer:    TSearchHelper;
    function    GetDefaultInterface: _SearchHelper;
    constructor Create(AServer: TSearchHelper);
  protected
    function Get_CheckHaveMatch(const Name: WideString): WordBool;
    function Get_Init(SysKind: SysKinds; LCID: Integer): HResult;
  public
    property DefaultInterface: _SearchHelper read GetDefaultInterface;
  published
  end;
{$ENDIF}


// *********************************************************************//
// The Class CoTypeLibInfo provides a Create and CreateRemote method to          
// create instances of the default interface _TypeLibInfo exposed by              
// the CoClass TypeLibInfo. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoTypeLibInfo = class
    class function Create: _TypeLibInfo;
    class function CreateRemote(const MachineName: string): _TypeLibInfo;
  end;


// *********************************************************************//
// OLE Server Proxy class declaration
// Server Object    : TTypeLibInfo
// Help String      : TypeLib information
// Default Interface: _TypeLibInfo
// Def. Intf. DISP? : No
// Event   Interface: 
// TypeFlags        : (2) CanCreate
// *********************************************************************//
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  TTypeLibInfoProperties= class;
{$ENDIF}
  TTypeLibInfo = class(TOleServer)
  private
    FIntf:        _TypeLibInfo;
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    FProps:       TTypeLibInfoProperties;
    function      GetServerProperties: TTypeLibInfoProperties;
{$ENDIF}
    function      GetDefaultInterface: _TypeLibInfo;
  protected
    procedure InitServerData; override;
    function Get_ContainingFile: WideString;
    procedure Set_ContainingFile(const retVal: WideString);
    function Get_Name: WideString;
    function Get_HelpContext: Integer;
    function Get_HelpFile: WideString;
    function Get_GUID: WideString;
    function Get_LCID: Integer;
    function Get_MajorVersion: Smallint;
    function Get_MinorVersion: Smallint;
    function Get_AttributeMask: Smallint;
    function Get_AttributeStrings(out AttributeArray: PSafeArray): Smallint;
    function Get_CoClasses: CoClasses;
    function Get_Interfaces: Interfaces;
    function Get_Constants: Constants;
    function Get_Declarations: Declarations;
    function Get_TypeInfoCount: Smallint;
    function Get_GetTypeInfo(var Index: OleVariant): TypeInfo;
    function Get_GetTypeInfoNumber(const Name: WideString): Smallint;
    procedure Set_AppObjString(const retVal: WideString);
    procedure Set_LibNum(retVal: Smallint);
    function Get_ShowLibName: WordBool;
    procedure Set_ShowLibName(retVal: WordBool);
    function Get_GetTypeKind(TypeInfoNumber: Smallint): TypeKinds;
    function Get_SysKind: SysKinds;
    function Get_SearchDefault: TliSearchTypes;
    procedure Set_SearchDefault(retVal: TliSearchTypes);
    function Get_TypeInfos: TypeInfos;
    function Get_Records: Records;
    function Get_IntrinsicAliases: IntrinsicAliases;
    function Get_CustomDataCollection: CustomDataCollection;
    function Get_Unions: Unions;
    function Get_HelpString(LCID: Integer): WideString;
    function Get_AppObjString: WideString;
    function Get_LibNum: Smallint;
    function Get_ITypeLib: IUnknown;
    procedure _Set_ITypeLib(const retVal: IUnknown);
    function Get_HelpStringDll(LCID: Integer): WideString;
    function Get_HelpStringContext: Integer;
    function Get_BestEquivalentType(const TypeInfoName: WideString): WideString;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: _TypeLibInfo);
    procedure Disconnect; override;
    function Me: TypeLibInfo;
    procedure LoadRegTypeLib(const TypeLibGuid: WideString; MajorVersion: Smallint; 
                             MinorVersion: Smallint; LCID: Integer);
    function IsSameLibrary(const CheckLib: TypeLibInfo): WordBool;
    function CaseTypeName(var bstrName: WideString; SearchType: TliSearchTypes): TliSearchTypes;
    function CaseMemberName(var bstrName: WideString; SearchType: TliSearchTypes): WordBool;
    procedure ResetSearchCriteria(TypeFilter: TypeFlags; IncludeEmptyTypes: WordBool; 
                                  ShowUnderscore: WordBool);
    function GetTypesWithMember(const MemberName: WideString; var StartResults: SearchResults; 
                                SearchType: TliSearchTypes; Sort: WordBool; ShowUnderscore: WordBool): SearchResults;
    function GetTypesWithMemberDirect(const MemberName: WideString; hWnd: SYSINT; 
                                      WindowType: TliWindowTypes; SearchType: TliSearchTypes; 
                                      ShowUnderscore: WordBool): Smallint;
    function GetMembersWithSubString(const SubString: WideString; var StartResults: SearchResults; 
                                     SearchType: TliSearchTypes; SearchMiddle: WordBool; 
                                     var Helper: SearchHelper; Sort: WordBool; 
                                     ShowUnderscore: WordBool): SearchResults;
    function GetMembersWithSubStringDirect(const SubString: WideString; hWnd: SYSINT; 
                                           WindowType: TliWindowTypes; SearchType: TliSearchTypes; 
                                           SearchMiddle: WordBool; var Helper: SearchHelper; 
                                           ShowUnderscore: WordBool): Smallint;
    function GetTypesWithSubString(const SubString: WideString; var StartResults: SearchResults; 
                                   SearchType: TliSearchTypes; SearchMiddle: WordBool; 
                                   Sort: WordBool): SearchResults;
    function GetTypesWithSubStringDirect(const SubString: WideString; hWnd: SYSINT; 
                                         WindowType: TliWindowTypes; SearchType: TliSearchTypes; 
                                         SearchMiddle: WordBool): Smallint;
    function GetTypes(var StartResults: SearchResults; SearchType: TliSearchTypes; Sort: WordBool): SearchResults;
    function GetTypesDirect(hWnd: SYSINT; WindowType: TliWindowTypes; SearchType: TliSearchTypes): Smallint;
    function GetMembers(SearchData: Integer; ShowUnderscore: WordBool): SearchResults;
    function GetMembersDirect(SearchData: Integer; hWnd: SYSINT; WindowType: TliWindowTypes; 
                              ItemDataType: TliItemDataTypes; ShowUnderscore: WordBool): Smallint;
    procedure SetMemberFilters(FuncFilter: FuncFlags; VarFilter: VarFlags);
    function MakeSearchData(const TypeInfoName: WideString; SearchType: TliSearchTypes): Integer;
    function GetMemberInfo(SearchData: Integer; InvokeKinds: InvokeKinds; MemberId: Integer; 
                           const MemberName: WideString): MemberInfo;
    function AddTypes(var TypeInfoNumbers: PSafeArray; var StartResults: SearchResults; 
                      SearchType: TliSearchTypes; Sort: WordBool): SearchResults;
    function AddTypesDirect(var TypeInfoNumbers: PSafeArray; hWnd: SYSINT; 
                            WindowType: TliWindowTypes; SearchType: TliSearchTypes): Smallint;
    procedure FreeSearchCriteria;
    procedure Register(const HelpDir: WideString);
    procedure UnRegister;
    function GetMembersWithSubStringEx(const SubString: WideString; 
                                       var InvokeGroupings: PSafeArray; 
                                       var StartResults: SearchResults; SearchType: TliSearchTypes; 
                                       SearchMiddle: WordBool; Sort: WordBool; 
                                       ShowUnderscore: WordBool): SearchResults;
    function GetTypesWithMemberEx(const MemberName: WideString; InvokeKind: InvokeKinds; 
                                  var StartResults: SearchResults; SearchType: TliSearchTypes; 
                                  Sort: WordBool; ShowUnderscore: WordBool): SearchResults;
    property DefaultInterface: _TypeLibInfo read GetDefaultInterface;
    property Name: WideString read Get_Name;
    property HelpContext: Integer read Get_HelpContext;
    property HelpFile: WideString read Get_HelpFile;
    property GUID: WideString read Get_GUID;
    property LCID: Integer read Get_LCID;
    property MajorVersion: Smallint read Get_MajorVersion;
    property MinorVersion: Smallint read Get_MinorVersion;
    property AttributeMask: Smallint read Get_AttributeMask;
    property AttributeStrings[out AttributeArray: PSafeArray]: Smallint read Get_AttributeStrings;
    property CoClasses: CoClasses read Get_CoClasses;
    property Interfaces: Interfaces read Get_Interfaces;
    property Constants: Constants read Get_Constants;
    property Declarations: Declarations read Get_Declarations;
    property TypeInfoCount: Smallint read Get_TypeInfoCount;
    property GetTypeInfo[var Index: OleVariant]: TypeInfo read Get_GetTypeInfo;
    property GetTypeInfoNumber[const Name: WideString]: Smallint read Get_GetTypeInfoNumber;
    property GetTypeKind[TypeInfoNumber: Smallint]: TypeKinds read Get_GetTypeKind;
    property SysKind: SysKinds read Get_SysKind;
    property TypeInfos: TypeInfos read Get_TypeInfos;
    property Records: Records read Get_Records;
    property IntrinsicAliases: IntrinsicAliases read Get_IntrinsicAliases;
    property CustomDataCollection: CustomDataCollection read Get_CustomDataCollection;
    property Unions: Unions read Get_Unions;
    property HelpString[LCID: Integer]: WideString read Get_HelpString;
    property ITypeLib: IUnknown read Get_ITypeLib write _Set_ITypeLib;
    property HelpStringDll[LCID: Integer]: WideString read Get_HelpStringDll;
    property HelpStringContext: Integer read Get_HelpStringContext;
    property BestEquivalentType[const TypeInfoName: WideString]: WideString read Get_BestEquivalentType;
    property ContainingFile: WideString read Get_ContainingFile write Set_ContainingFile;
    property AppObjString: WideString read Get_AppObjString write Set_AppObjString;
    property LibNum: Smallint read Get_LibNum write Set_LibNum;
    property ShowLibName: WordBool read Get_ShowLibName write Set_ShowLibName;
    property SearchDefault: TliSearchTypes read Get_SearchDefault write Set_SearchDefault;
  published
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    property Server: TTypeLibInfoProperties read GetServerProperties;
{$ENDIF}
  end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
// *********************************************************************//
// OLE Server Properties Proxy Class
// Server Object    : TTypeLibInfo
// (This object is used by the IDE's Property Inspector to allow editing
//  of the properties of this server)
// *********************************************************************//
 TTypeLibInfoProperties = class(TPersistent)
  private
    FServer:    TTypeLibInfo;
    function    GetDefaultInterface: _TypeLibInfo;
    constructor Create(AServer: TTypeLibInfo);
  protected
    function Get_ContainingFile: WideString;
    procedure Set_ContainingFile(const retVal: WideString);
    function Get_Name: WideString;
    function Get_HelpContext: Integer;
    function Get_HelpFile: WideString;
    function Get_GUID: WideString;
    function Get_LCID: Integer;
    function Get_MajorVersion: Smallint;
    function Get_MinorVersion: Smallint;
    function Get_AttributeMask: Smallint;
    function Get_AttributeStrings(out AttributeArray: PSafeArray): Smallint;
    function Get_CoClasses: CoClasses;
    function Get_Interfaces: Interfaces;
    function Get_Constants: Constants;
    function Get_Declarations: Declarations;
    function Get_TypeInfoCount: Smallint;
    function Get_GetTypeInfo(var Index: OleVariant): TypeInfo;
    function Get_GetTypeInfoNumber(const Name: WideString): Smallint;
    procedure Set_AppObjString(const retVal: WideString);
    procedure Set_LibNum(retVal: Smallint);
    function Get_ShowLibName: WordBool;
    procedure Set_ShowLibName(retVal: WordBool);
    function Get_GetTypeKind(TypeInfoNumber: Smallint): TypeKinds;
    function Get_SysKind: SysKinds;
    function Get_SearchDefault: TliSearchTypes;
    procedure Set_SearchDefault(retVal: TliSearchTypes);
    function Get_TypeInfos: TypeInfos;
    function Get_Records: Records;
    function Get_IntrinsicAliases: IntrinsicAliases;
    function Get_CustomDataCollection: CustomDataCollection;
    function Get_Unions: Unions;
    function Get_HelpString(LCID: Integer): WideString;
    function Get_AppObjString: WideString;
    function Get_LibNum: Smallint;
    function Get_ITypeLib: IUnknown;
    procedure _Set_ITypeLib(const retVal: IUnknown);
    function Get_HelpStringDll(LCID: Integer): WideString;
    function Get_HelpStringContext: Integer;
    function Get_BestEquivalentType(const TypeInfoName: WideString): WideString;
  public
    property DefaultInterface: _TypeLibInfo read GetDefaultInterface;
  published
    property ContainingFile: WideString read Get_ContainingFile write Set_ContainingFile;
    property AppObjString: WideString read Get_AppObjString write Set_AppObjString;
    property LibNum: Smallint read Get_LibNum write Set_LibNum;
    property ShowLibName: WordBool read Get_ShowLibName write Set_ShowLibName;
    property SearchDefault: TliSearchTypes read Get_SearchDefault write Set_SearchDefault;
  end;
{$ENDIF}


// *********************************************************************//
// The Class CoTLIApplication provides a Create and CreateRemote method to          
// create instances of the default interface _TLIApplication exposed by              
// the CoClass TLIApplication. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoTLIApplication = class
    class function Create: _TLIApplication;
    class function CreateRemote(const MachineName: string): _TLIApplication;
  end;

procedure Register;

resourcestring
  dtlServerPage = 'ActiveX';

  dtlOcxPage = 'ActiveX';

implementation

uses ComObj;

class function CoSearchHelper.Create: _SearchHelper;
begin
  Result := CreateComObject(CLASS_SearchHelper) as _SearchHelper;
end;

class function CoSearchHelper.CreateRemote(const MachineName: string): _SearchHelper;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_SearchHelper) as _SearchHelper;
end;

procedure TSearchHelper.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{8B217752-717D-11CE-AB5B-D41203C10000}';
    IntfIID:   '{8B217751-717D-11CE-AB5B-D41203C10000}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TSearchHelper.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as _SearchHelper;
  end;
end;

procedure TSearchHelper.ConnectTo(svrIntf: _SearchHelper);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TSearchHelper.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TSearchHelper.GetDefaultInterface: _SearchHelper;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL. Component is not connected to Server. You must call ''Connect'' or ''ConnectTo'' before this operation');
  Result := FIntf;
end;

constructor TSearchHelper.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps := TSearchHelperProperties.Create(Self);
{$ENDIF}
end;

destructor TSearchHelper.Destroy;
begin
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps.Free;
{$ENDIF}
  inherited Destroy;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
function TSearchHelper.GetServerProperties: TSearchHelperProperties;
begin
  Result := FProps;
end;
{$ENDIF}

function TSearchHelper.Get_CheckHaveMatch(const Name: WideString): WordBool;
begin
    Result := DefaultInterface.CheckHaveMatch[Name];
end;

function TSearchHelper.Get_Init(SysKind: SysKinds; LCID: Integer): HResult;
begin
    Result := DefaultInterface.Get_Init(SysKind, LCID);
end;

function TSearchHelper.Me: TypeLibInfo;
begin
  Result := DefaultInterface.Me;
end;

procedure TSearchHelper._OldInit;
begin
  DefaultInterface._OldInit;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
constructor TSearchHelperProperties.Create(AServer: TSearchHelper);
begin
  inherited Create;
  FServer := AServer;
end;

function TSearchHelperProperties.GetDefaultInterface: _SearchHelper;
begin
  Result := FServer.DefaultInterface;
end;

function TSearchHelperProperties.Get_CheckHaveMatch(const Name: WideString): WordBool;
begin
    Result := DefaultInterface.CheckHaveMatch[Name];
end;

function TSearchHelperProperties.Get_Init(SysKind: SysKinds; LCID: Integer): HResult;
begin
    Result := DefaultInterface.Get_Init(SysKind, LCID);
end;

{$ENDIF}

class function CoTypeLibInfo.Create: _TypeLibInfo;
begin
  Result := CreateComObject(CLASS_TypeLibInfo) as _TypeLibInfo;
end;

class function CoTypeLibInfo.CreateRemote(const MachineName: string): _TypeLibInfo;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_TypeLibInfo) as _TypeLibInfo;
end;

procedure TTypeLibInfo.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{8B217746-717D-11CE-AB5B-D41203C10000}';
    IntfIID:   '{8B217745-717D-11CE-AB5B-D41203C10000}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TTypeLibInfo.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as _TypeLibInfo;
  end;
end;

procedure TTypeLibInfo.ConnectTo(svrIntf: _TypeLibInfo);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TTypeLibInfo.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TTypeLibInfo.GetDefaultInterface: _TypeLibInfo;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL. Component is not connected to Server. You must call ''Connect'' or ''ConnectTo'' before this operation');
  Result := FIntf;
end;

constructor TTypeLibInfo.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps := TTypeLibInfoProperties.Create(Self);
{$ENDIF}
end;

destructor TTypeLibInfo.Destroy;
begin
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps.Free;
{$ENDIF}
  inherited Destroy;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
function TTypeLibInfo.GetServerProperties: TTypeLibInfoProperties;
begin
  Result := FProps;
end;
{$ENDIF}

function TTypeLibInfo.Get_ContainingFile: WideString;
begin
    Result := DefaultInterface.ContainingFile;
end;

procedure TTypeLibInfo.Set_ContainingFile(const retVal: WideString);
  { Warning: The property ContainingFile has a setter and a getter whose
    types do not match. Delphi was unable to generate a property of
    this sort and so is using a Variant as a passthrough. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.ContainingFile := retVal;
end;

function TTypeLibInfo.Get_Name: WideString;
begin
    Result := DefaultInterface.Name;
end;

function TTypeLibInfo.Get_HelpContext: Integer;
begin
    Result := DefaultInterface.HelpContext;
end;

function TTypeLibInfo.Get_HelpFile: WideString;
begin
    Result := DefaultInterface.HelpFile;
end;

function TTypeLibInfo.Get_GUID: WideString;
begin
    Result := DefaultInterface.GUID;
end;

function TTypeLibInfo.Get_LCID: Integer;
begin
    Result := DefaultInterface.LCID;
end;

function TTypeLibInfo.Get_MajorVersion: Smallint;
begin
    Result := DefaultInterface.MajorVersion;
end;

function TTypeLibInfo.Get_MinorVersion: Smallint;
begin
    Result := DefaultInterface.MinorVersion;
end;

function TTypeLibInfo.Get_AttributeMask: Smallint;
begin
    Result := DefaultInterface.AttributeMask;
end;

function TTypeLibInfo.Get_AttributeStrings(out AttributeArray: PSafeArray): Smallint;
begin
    Result := DefaultInterface.AttributeStrings[AttributeArray];
end;

function TTypeLibInfo.Get_CoClasses: CoClasses;
begin
    Result := DefaultInterface.CoClasses;
end;

function TTypeLibInfo.Get_Interfaces: Interfaces;
begin
    Result := DefaultInterface.Interfaces;
end;

function TTypeLibInfo.Get_Constants: Constants;
begin
    Result := DefaultInterface.Constants;
end;

function TTypeLibInfo.Get_Declarations: Declarations;
begin
    Result := DefaultInterface.Declarations;
end;

function TTypeLibInfo.Get_TypeInfoCount: Smallint;
begin
    Result := DefaultInterface.TypeInfoCount;
end;

function TTypeLibInfo.Get_GetTypeInfo(var Index: OleVariant): TypeInfo;
begin
    Result := DefaultInterface.GetTypeInfo[Index];
end;

function TTypeLibInfo.Get_GetTypeInfoNumber(const Name: WideString): Smallint;
begin
    Result := DefaultInterface.GetTypeInfoNumber[Name];
end;

procedure TTypeLibInfo.Set_AppObjString(const retVal: WideString);
  { Warning: The property AppObjString has a setter and a getter whose
    types do not match. Delphi was unable to generate a property of
    this sort and so is using a Variant as a passthrough. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.AppObjString := retVal;
end;

procedure TTypeLibInfo.Set_LibNum(retVal: Smallint);
begin
  DefaultInterface.Set_LibNum(retVal);
end;

function TTypeLibInfo.Get_ShowLibName: WordBool;
begin
    Result := DefaultInterface.ShowLibName;
end;

procedure TTypeLibInfo.Set_ShowLibName(retVal: WordBool);
begin
  DefaultInterface.Set_ShowLibName(retVal);
end;

function TTypeLibInfo.Get_GetTypeKind(TypeInfoNumber: Smallint): TypeKinds;
begin
    Result := DefaultInterface.GetTypeKind[TypeInfoNumber];
end;

function TTypeLibInfo.Get_SysKind: SysKinds;
begin
    Result := DefaultInterface.SysKind;
end;

function TTypeLibInfo.Get_SearchDefault: TliSearchTypes;
begin
    Result := DefaultInterface.SearchDefault;
end;

procedure TTypeLibInfo.Set_SearchDefault(retVal: TliSearchTypes);
begin
  DefaultInterface.Set_SearchDefault(retVal);
end;

function TTypeLibInfo.Get_TypeInfos: TypeInfos;
begin
    Result := DefaultInterface.TypeInfos;
end;

function TTypeLibInfo.Get_Records: Records;
begin
    Result := DefaultInterface.Records;
end;

function TTypeLibInfo.Get_IntrinsicAliases: IntrinsicAliases;
begin
    Result := DefaultInterface.IntrinsicAliases;
end;

function TTypeLibInfo.Get_CustomDataCollection: CustomDataCollection;
begin
    Result := DefaultInterface.CustomDataCollection;
end;

function TTypeLibInfo.Get_Unions: Unions;
begin
    Result := DefaultInterface.Unions;
end;

function TTypeLibInfo.Get_HelpString(LCID: Integer): WideString;
begin
    Result := DefaultInterface.HelpString[LCID];
end;

function TTypeLibInfo.Get_AppObjString: WideString;
begin
    Result := DefaultInterface.AppObjString;
end;

function TTypeLibInfo.Get_LibNum: Smallint;
begin
    Result := DefaultInterface.LibNum;
end;

function TTypeLibInfo.Get_ITypeLib: IUnknown;
begin
    Result := DefaultInterface.ITypeLib;
end;

procedure TTypeLibInfo._Set_ITypeLib(const retVal: IUnknown);
  { Warning: The property ITypeLib has a setter and a getter whose
    types do not match. Delphi was unable to generate a property of
    this sort and so is using a Variant as a passthrough. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.ITypeLib := retVal;
end;

function TTypeLibInfo.Get_HelpStringDll(LCID: Integer): WideString;
begin
    Result := DefaultInterface.HelpStringDll[LCID];
end;

function TTypeLibInfo.Get_HelpStringContext: Integer;
begin
    Result := DefaultInterface.HelpStringContext;
end;

function TTypeLibInfo.Get_BestEquivalentType(const TypeInfoName: WideString): WideString;
begin
    Result := DefaultInterface.BestEquivalentType[TypeInfoName];
end;

function TTypeLibInfo.Me: TypeLibInfo;
begin
  Result := DefaultInterface.Me;
end;

procedure TTypeLibInfo.LoadRegTypeLib(const TypeLibGuid: WideString; MajorVersion: Smallint; 
                                      MinorVersion: Smallint; LCID: Integer);
begin
  DefaultInterface.LoadRegTypeLib(TypeLibGuid, MajorVersion, MinorVersion, LCID);
end;

function TTypeLibInfo.IsSameLibrary(const CheckLib: TypeLibInfo): WordBool;
begin
  Result := DefaultInterface.IsSameLibrary(CheckLib);
end;

function TTypeLibInfo.CaseTypeName(var bstrName: WideString; SearchType: TliSearchTypes): TliSearchTypes;
begin
  Result := DefaultInterface.CaseTypeName(bstrName, SearchType);
end;

function TTypeLibInfo.CaseMemberName(var bstrName: WideString; SearchType: TliSearchTypes): WordBool;
begin
  Result := DefaultInterface.CaseMemberName(bstrName, SearchType);
end;

procedure TTypeLibInfo.ResetSearchCriteria(TypeFilter: TypeFlags; IncludeEmptyTypes: WordBool; 
                                           ShowUnderscore: WordBool);
begin
  DefaultInterface.ResetSearchCriteria(TypeFilter, IncludeEmptyTypes, ShowUnderscore);
end;

function TTypeLibInfo.GetTypesWithMember(const MemberName: WideString; 
                                         var StartResults: SearchResults; 
                                         SearchType: TliSearchTypes; Sort: WordBool; 
                                         ShowUnderscore: WordBool): SearchResults;
begin
  Result := DefaultInterface.GetTypesWithMember(MemberName, StartResults, SearchType, Sort, 
                                                ShowUnderscore);
end;

function TTypeLibInfo.GetTypesWithMemberDirect(const MemberName: WideString; hWnd: SYSINT; 
                                               WindowType: TliWindowTypes; 
                                               SearchType: TliSearchTypes; ShowUnderscore: WordBool): Smallint;
begin
  Result := DefaultInterface.GetTypesWithMemberDirect(MemberName, hWnd, WindowType, SearchType, 
                                                      ShowUnderscore);
end;

function TTypeLibInfo.GetMembersWithSubString(const SubString: WideString; 
                                              var StartResults: SearchResults; 
                                              SearchType: TliSearchTypes; SearchMiddle: WordBool; 
                                              var Helper: SearchHelper; Sort: WordBool; 
                                              ShowUnderscore: WordBool): SearchResults;
begin
  Result := DefaultInterface.GetMembersWithSubString(SubString, StartResults, SearchType, 
                                                     SearchMiddle, Helper, Sort, ShowUnderscore);
end;

function TTypeLibInfo.GetMembersWithSubStringDirect(const SubString: WideString; hWnd: SYSINT; 
                                                    WindowType: TliWindowTypes; 
                                                    SearchType: TliSearchTypes; 
                                                    SearchMiddle: WordBool; 
                                                    var Helper: SearchHelper; 
                                                    ShowUnderscore: WordBool): Smallint;
begin
  Result := DefaultInterface.GetMembersWithSubStringDirect(SubString, hWnd, WindowType, SearchType, 
                                                           SearchMiddle, Helper, ShowUnderscore);
end;

function TTypeLibInfo.GetTypesWithSubString(const SubString: WideString; 
                                            var StartResults: SearchResults; 
                                            SearchType: TliSearchTypes; SearchMiddle: WordBool; 
                                            Sort: WordBool): SearchResults;
begin
  Result := DefaultInterface.GetTypesWithSubString(SubString, StartResults, SearchType, 
                                                   SearchMiddle, Sort);
end;

function TTypeLibInfo.GetTypesWithSubStringDirect(const SubString: WideString; hWnd: SYSINT; 
                                                  WindowType: TliWindowTypes; 
                                                  SearchType: TliSearchTypes; SearchMiddle: WordBool): Smallint;
begin
  Result := DefaultInterface.GetTypesWithSubStringDirect(SubString, hWnd, WindowType, SearchType, 
                                                         SearchMiddle);
end;

function TTypeLibInfo.GetTypes(var StartResults: SearchResults; SearchType: TliSearchTypes; 
                               Sort: WordBool): SearchResults;
begin
  Result := DefaultInterface.GetTypes(StartResults, SearchType, Sort);
end;

function TTypeLibInfo.GetTypesDirect(hWnd: SYSINT; WindowType: TliWindowTypes; 
                                     SearchType: TliSearchTypes): Smallint;
begin
  Result := DefaultInterface.GetTypesDirect(hWnd, WindowType, SearchType);
end;

function TTypeLibInfo.GetMembers(SearchData: Integer; ShowUnderscore: WordBool): SearchResults;
begin
  Result := DefaultInterface.GetMembers(SearchData, ShowUnderscore);
end;

function TTypeLibInfo.GetMembersDirect(SearchData: Integer; hWnd: SYSINT; 
                                       WindowType: TliWindowTypes; ItemDataType: TliItemDataTypes; 
                                       ShowUnderscore: WordBool): Smallint;
begin
  Result := DefaultInterface.GetMembersDirect(SearchData, hWnd, WindowType, ItemDataType, 
                                              ShowUnderscore);
end;

procedure TTypeLibInfo.SetMemberFilters(FuncFilter: FuncFlags; VarFilter: VarFlags);
begin
  DefaultInterface.SetMemberFilters(FuncFilter, VarFilter);
end;

function TTypeLibInfo.MakeSearchData(const TypeInfoName: WideString; SearchType: TliSearchTypes): Integer;
begin
  Result := DefaultInterface.MakeSearchData(TypeInfoName, SearchType);
end;

function TTypeLibInfo.GetMemberInfo(SearchData: Integer; InvokeKinds: InvokeKinds; 
                                    MemberId: Integer; const MemberName: WideString): MemberInfo;
begin
  Result := DefaultInterface.GetMemberInfo(SearchData, InvokeKinds, MemberId, MemberName);
end;

function TTypeLibInfo.AddTypes(var TypeInfoNumbers: PSafeArray; var StartResults: SearchResults; 
                               SearchType: TliSearchTypes; Sort: WordBool): SearchResults;
begin
  Result := DefaultInterface.AddTypes(TypeInfoNumbers, StartResults, SearchType, Sort);
end;

function TTypeLibInfo.AddTypesDirect(var TypeInfoNumbers: PSafeArray; hWnd: SYSINT; 
                                     WindowType: TliWindowTypes; SearchType: TliSearchTypes): Smallint;
begin
  Result := DefaultInterface.AddTypesDirect(TypeInfoNumbers, hWnd, WindowType, SearchType);
end;

procedure TTypeLibInfo.FreeSearchCriteria;
begin
  DefaultInterface.FreeSearchCriteria;
end;

procedure TTypeLibInfo.Register(const HelpDir: WideString);
begin
  DefaultInterface.Register(HelpDir);
end;

procedure TTypeLibInfo.UnRegister;
begin
  DefaultInterface.UnRegister;
end;

function TTypeLibInfo.GetMembersWithSubStringEx(const SubString: WideString; 
                                                var InvokeGroupings: PSafeArray; 
                                                var StartResults: SearchResults; 
                                                SearchType: TliSearchTypes; SearchMiddle: WordBool; 
                                                Sort: WordBool; ShowUnderscore: WordBool): SearchResults;
begin
  Result := DefaultInterface.GetMembersWithSubStringEx(SubString, InvokeGroupings, StartResults, 
                                                       SearchType, SearchMiddle, Sort, 
                                                       ShowUnderscore);
end;

function TTypeLibInfo.GetTypesWithMemberEx(const MemberName: WideString; InvokeKind: InvokeKinds; 
                                           var StartResults: SearchResults; 
                                           SearchType: TliSearchTypes; Sort: WordBool; 
                                           ShowUnderscore: WordBool): SearchResults;
begin
  Result := DefaultInterface.GetTypesWithMemberEx(MemberName, InvokeKind, StartResults, SearchType, 
                                                  Sort, ShowUnderscore);
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
constructor TTypeLibInfoProperties.Create(AServer: TTypeLibInfo);
begin
  inherited Create;
  FServer := AServer;
end;

function TTypeLibInfoProperties.GetDefaultInterface: _TypeLibInfo;
begin
  Result := FServer.DefaultInterface;
end;

function TTypeLibInfoProperties.Get_ContainingFile: WideString;
begin
    Result := DefaultInterface.ContainingFile;
end;

procedure TTypeLibInfoProperties.Set_ContainingFile(const retVal: WideString);
  { Warning: The property ContainingFile has a setter and a getter whose
    types do not match. Delphi was unable to generate a property of
    this sort and so is using a Variant as a passthrough. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.ContainingFile := retVal;
end;

function TTypeLibInfoProperties.Get_Name: WideString;
begin
    Result := DefaultInterface.Name;
end;

function TTypeLibInfoProperties.Get_HelpContext: Integer;
begin
    Result := DefaultInterface.HelpContext;
end;

function TTypeLibInfoProperties.Get_HelpFile: WideString;
begin
    Result := DefaultInterface.HelpFile;
end;

function TTypeLibInfoProperties.Get_GUID: WideString;
begin
    Result := DefaultInterface.GUID;
end;

function TTypeLibInfoProperties.Get_LCID: Integer;
begin
    Result := DefaultInterface.LCID;
end;

function TTypeLibInfoProperties.Get_MajorVersion: Smallint;
begin
    Result := DefaultInterface.MajorVersion;
end;

function TTypeLibInfoProperties.Get_MinorVersion: Smallint;
begin
    Result := DefaultInterface.MinorVersion;
end;

function TTypeLibInfoProperties.Get_AttributeMask: Smallint;
begin
    Result := DefaultInterface.AttributeMask;
end;

function TTypeLibInfoProperties.Get_AttributeStrings(out AttributeArray: PSafeArray): Smallint;
begin
    Result := DefaultInterface.AttributeStrings[AttributeArray];
end;

function TTypeLibInfoProperties.Get_CoClasses: CoClasses;
begin
    Result := DefaultInterface.CoClasses;
end;

function TTypeLibInfoProperties.Get_Interfaces: Interfaces;
begin
    Result := DefaultInterface.Interfaces;
end;

function TTypeLibInfoProperties.Get_Constants: Constants;
begin
    Result := DefaultInterface.Constants;
end;

function TTypeLibInfoProperties.Get_Declarations: Declarations;
begin
    Result := DefaultInterface.Declarations;
end;

function TTypeLibInfoProperties.Get_TypeInfoCount: Smallint;
begin
    Result := DefaultInterface.TypeInfoCount;
end;

function TTypeLibInfoProperties.Get_GetTypeInfo(var Index: OleVariant): TypeInfo;
begin
    Result := DefaultInterface.GetTypeInfo[Index];
end;

function TTypeLibInfoProperties.Get_GetTypeInfoNumber(const Name: WideString): Smallint;
begin
    Result := DefaultInterface.GetTypeInfoNumber[Name];
end;

procedure TTypeLibInfoProperties.Set_AppObjString(const retVal: WideString);
  { Warning: The property AppObjString has a setter and a getter whose
    types do not match. Delphi was unable to generate a property of
    this sort and so is using a Variant as a passthrough. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.AppObjString := retVal;
end;

procedure TTypeLibInfoProperties.Set_LibNum(retVal: Smallint);
begin
  DefaultInterface.Set_LibNum(retVal);
end;

function TTypeLibInfoProperties.Get_ShowLibName: WordBool;
begin
    Result := DefaultInterface.ShowLibName;
end;

procedure TTypeLibInfoProperties.Set_ShowLibName(retVal: WordBool);
begin
  DefaultInterface.Set_ShowLibName(retVal);
end;

function TTypeLibInfoProperties.Get_GetTypeKind(TypeInfoNumber: Smallint): TypeKinds;
begin
    Result := DefaultInterface.GetTypeKind[TypeInfoNumber];
end;

function TTypeLibInfoProperties.Get_SysKind: SysKinds;
begin
    Result := DefaultInterface.SysKind;
end;

function TTypeLibInfoProperties.Get_SearchDefault: TliSearchTypes;
begin
    Result := DefaultInterface.SearchDefault;
end;

procedure TTypeLibInfoProperties.Set_SearchDefault(retVal: TliSearchTypes);
begin
  DefaultInterface.Set_SearchDefault(retVal);
end;

function TTypeLibInfoProperties.Get_TypeInfos: TypeInfos;
begin
    Result := DefaultInterface.TypeInfos;
end;

function TTypeLibInfoProperties.Get_Records: Records;
begin
    Result := DefaultInterface.Records;
end;

function TTypeLibInfoProperties.Get_IntrinsicAliases: IntrinsicAliases;
begin
    Result := DefaultInterface.IntrinsicAliases;
end;

function TTypeLibInfoProperties.Get_CustomDataCollection: CustomDataCollection;
begin
    Result := DefaultInterface.CustomDataCollection;
end;

function TTypeLibInfoProperties.Get_Unions: Unions;
begin
    Result := DefaultInterface.Unions;
end;

function TTypeLibInfoProperties.Get_HelpString(LCID: Integer): WideString;
begin
    Result := DefaultInterface.HelpString[LCID];
end;

function TTypeLibInfoProperties.Get_AppObjString: WideString;
begin
    Result := DefaultInterface.AppObjString;
end;

function TTypeLibInfoProperties.Get_LibNum: Smallint;
begin
    Result := DefaultInterface.LibNum;
end;

function TTypeLibInfoProperties.Get_ITypeLib: IUnknown;
begin
    Result := DefaultInterface.ITypeLib;
end;

procedure TTypeLibInfoProperties._Set_ITypeLib(const retVal: IUnknown);
  { Warning: The property ITypeLib has a setter and a getter whose
    types do not match. Delphi was unable to generate a property of
    this sort and so is using a Variant as a passthrough. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.ITypeLib := retVal;
end;

function TTypeLibInfoProperties.Get_HelpStringDll(LCID: Integer): WideString;
begin
    Result := DefaultInterface.HelpStringDll[LCID];
end;

function TTypeLibInfoProperties.Get_HelpStringContext: Integer;
begin
    Result := DefaultInterface.HelpStringContext;
end;

function TTypeLibInfoProperties.Get_BestEquivalentType(const TypeInfoName: WideString): WideString;
begin
    Result := DefaultInterface.BestEquivalentType[TypeInfoName];
end;

{$ENDIF}

class function CoTLIApplication.Create: _TLIApplication;
begin
  Result := CreateComObject(CLASS_TLIApplication) as _TLIApplication;
end;

class function CoTLIApplication.CreateRemote(const MachineName: string): _TLIApplication;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_TLIApplication) as _TLIApplication;
end;

procedure Register;
begin
  RegisterComponents(dtlServerPage, [TSearchHelper, TTypeLibInfo]);
end;

end.
