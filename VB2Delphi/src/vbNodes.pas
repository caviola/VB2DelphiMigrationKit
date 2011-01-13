// The MIT License
//
// Copyright (c) 2011 Albert Almeida (caviola@gmail.com)
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.

unit vbNodes;

interface

  uses
    Classes,
    Contnrs,
    Types,
    HashList;

  const
    UNKNOWN_TYPE  = $0000;
    BOOLEAN_TYPE  = $0001;
    BYTE_TYPE     = $0002;
    CURRENCY_TYPE = $0004;
    DATE_TYPE     = $0008;
    DOUBLE_TYPE   = $0010;
    INTEGER_TYPE  = $0020;
    LONG_TYPE     = $0040;
    OBJECT_TYPE   = $0080;
    SINGLE_TYPE   = $0100;
    STRING_TYPE   = $0200;
    VARIANT_TYPE  = $0400;
    ANY_TYPE      = $0800;
    ARRAY_TYPE    = $8000;

  type

    TvbNodeList = class;

    TvbNodeKind = (
      // Definitions
      CLASS_DEF,
      CLASS_MODULE_DEF,
      FORM_MODULE_DEF,
      CONST_DEF,
      DLL_FUNC_DEF,
      ENUM_DEF,
      EVENT_DEF,
      FUNC_DEF,
      LABEL_DEF,
      LIB_MODULE_DEF,
      PARAM_DEF,
      PROJECT_DEF,
      PROPERTY_DEF,
      RECORD_DEF,
      STD_MODULE_DEF,
      VAR_DEF,

      // Names
      LABEL_REF,
      QUALIFIED_NAME,
      SIMPLE_NAME,

      // Type reference
      TYPE_REF,

      // Expressions,
      ABS_EXPR,
      ADDRESSOF_EXPR,
      ARRAY_EXPR,
      BINARY_EXPR,
      BOOL_LIT_EXPR,
      CALL_OR_INDEXER_EXPR,
      CAST_EXPR,
      DATE_EXPR,
      DATE_LIT_EXPR,
      DICT_ACCESS_EXPR,
      DOEVENTS_EXPR,
      FIX_EXPR,
      FLOAT_LIT_EXPR,
      FUNC_RESULT_EXPR,
      INPUT_EXPR,
      INT_EXPR,
      INT_LIT_EXPR,
      LBOUND_EXPR,
      LEN_EXPR,
      LENB_EXPR,
      ME_EXPR,
      MEMBER_ACCESS_EXPR,
      MID_EXPR,
      NAME_EXPR,
      NAMED_ARG_EXPR,
      NEW_EXPR,
      NOTHING_EXPR,
      PRINT_ARG_EXPR,
      SEEK_EXPR,
      SGN_EXPR,
      STRING_EXPR,
      STRING_LIT_EXPR,
      TYPEOF_EXPR,
      UBOUND_EXPR,
      UNARY_EXPR,

      // Statements
      ASSIGN_STMT,
      CALL_STMT,
      CASE_CLAUSE,
      CASE_STMT,
      CIRCLE_STMT,
      CLOSE_STMT,
      DATE_STMT,
      DEBUG_ASSERT_STMT,
      DEBUG_PRINT_STMT,
      DO_LOOP_STMT,
      END_STMT,
      ERASE_STMT,
      ERROR_STMT,
      EXIT_LOOP_STMT,
      EXIT_STMT,
      FOR_STMT,
      FOREACH_STMT,
      GET_STMT,
      GOTO_OR_GOSUB_STMT,
      IF_STMT,
      INPUT_STMT,
      LABEL_DEF_STMT,
      LINE_INPUT_STMT,
      LINE_STMT,
      LOCK_STMT,
      MID_ASSIGN_STMT,
      NAME_STMT,
      ON_ERROR_STMT,
      ONGOTO_OR_GOSUB_STMT,
      OPEN_STMT,
      PRINT_METHOD_STMT,
      PRINT_STMT,
      PSET_STMT,
      PUT_STMT,
      RAISE_EVENT_STMT,
      RANGE_CASE_CLAUSE,
      REDIM_STMT,
      REL_CASE_CLAUSE,
      RESUME_STMT,
      RETURN_STMT,
      SEEK_STMT,
      SELECT_CASE_STMT,
      STMT_BLOCK,
      STOP_STMT,
      TIME_STMT,
      UNLOCK_STMT,
      WIDTH_STMT,
      WITH_STMT,
      WRITE_STMT,

      nkImplIntf // implemented interface
    );

    TvbNodeKinds = set of TvbNodeKind;

    TvbNodeFlags = set of (
      dfFriend,                  // friend declaration
      dfPrivate,                 // private declaration
      dfPublic,                  // public declaration
      dfStatic,                  // static declaration
      dfGlobal,                  // global declaration
      dfConst,                   // a constant (Const) declaration
      dfDim,                     // a variable (Dim) declaration
      dfWithEvents,
      dfNew,                     // an auto-instance object declaration
      dfExternal,                // imported from an external library
      dfPropSet,                 // property set
      dfPropLet,                 // property let
      dfPropGet,                 // property get
      dfFormControl,
      nfReDim_DeclaredLocalVar,  // a dynamic array declared with ReDim
      nfArgByRef,
      nfArgByVal,
      for_StepMinus1,            // a For statement with Step -1
      for_StepN,                 // a For statement with Step N
      for_StepMinusN,            // a For statement with Step -N
      expr_Const,                // a constant expression
      expr_UsedWithArgs,         // the expression is used in a TvbCallOrIndexerExpr.
      expr_MustBeFunc,           // the name that the expression refers to must be a function
      select_AsIfElse,           // convert a Select Case statement to if..then
      enum_MissingValues,        // an enum type in which at least one constant has no explicit value.
      assign_LHS,                // the node is the LHS of an assignment expression
      int_Octal,
      int_Hex,
      sfx_Int,                   // integer literal suffixed with ''  
      sfx_Long,                  // integer literal suffixed with ''
      sfx_Single,                // numeric literal suffixed with ''
      sfx_Curr, 
      pfByVal,                   // pass by value
      pfByRef,                   // pass by reference
      pfParamArray,              // an optional list of parameters
      pfOptional,                // optional parameter
      _pfPChar,                  // string parameter of a DLL function
      _pfWordBool,               // boolean parameter of a DLL function
      _pfAnyByVal,
      _pfAnyByRef,
      _pfPointerType,
      expr_isNotMethod,
      mapAsFunctionCall,          // convert a property put to a function call
      map_OnlyName,
      map_Default,
      map_Full 
      );

    TvbType = class;

    TvbScope = class;

    IvbType = interface;

    TvbNodeVisitor = class;

    TvbNode =
      class
        protected
          fKind         : TvbNodeKind;
          fFlags        : TvbNodeFlags;
          fName1        : string; // the translation expression for this node
          fNodeType     : IvbType;
          fTopCmts      : TStringDynArray;
          fDownCmts     : TStringDynArray;
          fTopRightCmt  : string;
          fDownRightCmt : string;
        protected
          constructor Create( kind : TvbNodeKind );
        public
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); virtual;
          property NodeKind  : TvbNodeKind  read fKind;
          property NodeFlags : TvbNodeFlags read fFlags    write fFlags;
          property NodeType  : IvbType      read fNodeType write fNodeType;
          property Name1     : string       read fName1    write fName1;
          // These properties hold the comments around a node, e.g:
          //    'TopCmd_1
          //    'TopCmd_2
          //    Dim Name As String 'TopRightCmt
          property TopCmts      : TStringDynArray read fTopCmts      write fTopCmts;
          property DownCmts     : TStringDynArray read fDownCmts     write fDownCmts;
          property TopRightCmt  : string          read fTopRightCmt  write fTopRightCmt;
          property DownRightCmt : string          read fDownRightCmt write fDownRightCmt;
      end;

    TvbNodeList =
      class
        private
          fNodes : TObjectList;
          function GetNodeCount: integer;
        protected
          function GetNode( i : integer ) : TvbNode; virtual;
        public
          constructor Create( OwnsNodes : boolean = true );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil );
          function Add( node : TvbNode ) : TvbNode; overload;
          property Nodes[i : integer] : TvbNode read GetNode; default;
          property Count              : integer read GetNodeCount;
      end;

    TvbDef         = class;
    TvbConstDef    = class;
    TvbVarDef      = class;
    TvbFuncDef     = class;
    TvbTypeDef     = class;
    TvbSimpleName  = class;

    TvbNameBindingKind = (
      TYPE_BINDING,             // name bound to type declaration(record or enum)
      LABEL_BINDING,            // name bound to a label declaration
      OBJECT_BINDING,           // name bound to a var, const, function or property
      EVENT_BINDING,            // name bound to an event declaration
      PROJECT_OR_MODULE_BINDING // name bound to project or module declaration
      );

    // Stores all posible entity declarations to which a name may be bound
    // in a given scope.
    //
    TvbNameBindings =
      class
        public
          Bindings : array[TvbNameBindingKind] of TvbDef;
      end;

    TvbLookupObjectFlags = set of (
      // The declaration must take arguments.
      // So it must be a function/array or property.
      lofMustTakeArgs,
      // The declaration must be an RValue.
      // So it must be an array or property-put.
      lofMustBeRValue,
      // The declaration must be a function.
      lofMustBeFunc
      );

    TvbClassDef = class;

    // An scope is a region of the program where entities are declared.
    // VB defines the following kinds of scopes:
    //
    //    1) global or project scope
    //    2) module scope
    //    3) type(record and enum) scope
    //    4) function scope 
    //
    // VB defines the following kinds of entities:
    //
    //    1) project
    //    2) modules(standard and object)
    //    3) functions
    //    4) properties
    //    5) variables
    //    6) constants
    //    7) types(records and enumerated constants).
    //    8) events
    //
    // In a given scope two enities of the same kind cannot have the same name,
    // but two entities of different kinds can happily share a name.
    // For example, in a module we can have a global varaible named "UserData"
    // and a record type also named "UserData", but we cannot have a record
    // named "PersonInfo" and an enum named "PersonInfo". Neither can we have
    // a module named like the project.
    //
    // Therefore for a given name an scope stores all posible entities that
    // may be bound to it and we represent this association with a
    // TvbNameBindings instance.
    //
    // An scope(child) may be contained within another scope(parent) and we
    // capture this relationship by defining a Parent property with allows
    // a child scope to have access to its parent.
    // Example, the function scope(where local vars and consts are declared)
    // is contained inside the module scope, so, the "function scope" is the
    // children and the "module scope" is the parent. The "function scope" has
    // a Parent property that links it to the "module scope".
    //
    // To determine the entity to which a name refers, clients use the Lookup
    // function. Given that a name may be bound to different kinds of entities,
    // the client must specify what kind of entity should be matched.
    // An unqualified name is initially searched for in the scope where the Lookup
    // operaton occurs and if not found there, it's searched for recursively
    // in the parent scopes. Qualified names are searched for recursively
    // in the scope of their qualifiers and eventually the top qualifier
    // in the scope where the Lookup operation occurs.
    //
    TvbScope =
      class
        protected
          {$IFDEF DEBUG}
          fName : string;
          {$ENDIF}
          fParent      : TvbScope;
          fBaseScope   : TvbScope;
          fBindings    : THashList;
          fTraits      : TCaseInsensitiveTraits;
          fHaveBinding : array[TvbNameBindingKind] of boolean;
          function FindBinding( const name : string ) : TvbNameBindings;
        public
          constructor Create( parent : TvbScope = nil{$IFDEF DEBUG}; const name : string = '' {$ENDIF} );
          destructor Destroy; override;
          function Bind( decl : TvbDef; kind : TvbNameBindingKind ) : boolean;
          function Lookup( kind          : TvbNameBindingKind;
                           const name    : string;
                           var decl      : TvbDef;
                           searchParents : boolean = true ) : boolean; virtual;

          // Use to specify the certain aspects of the object we are looking for.
          function LookupObject( const name    : string;
                                 var decl      : TvbDef;
                                 flags         : TvbLookupObjectFlags = [];
                                 searchParents : boolean = true ) : boolean; virtual;
                                 
          // Declares the class members in this scope.
          // Remember this scope DOES NOT take ownership of those members.
          procedure BindMembers( cls : TvbClassDef );
          property Parent    : TvbScope read fParent    write fParent;
          property BaseScope : TvbScope read fBaseScope write fBaseScope;
          {$IFDEF DEBUG}
          property Name : string read fName write fName;
          {$ENDIF}
      end;

    TvbProjectScope  = class( TvbScope );
    TvbModuleScope   = class( TvbScope );
    TvbTypeScope     = class( TvbScope );
    TvbFunctionScope = class( TvbScope );

    TvbModuleDef = class;

    // A "name" represents the use/reference of a declaration.
    TvbName =
      class( TvbNode )
        private
          fName : string;
          fDef : TvbDef;
        protected
          constructor Create( nk : TvbNodeKind; const name : string );
        public
          property Def  : TvbDef read fDef write fDef;
          property Name : string read fName;
      end;

    TvbSimpleName =
      class( TvbName )
        public
          constructor Create( const name : string );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
      end;

    TvbQualifiedName =
      class( TvbName )
        private
          fQualifier : TvbName;
        public
          constructor Create( const name : string; qual : TvbName = nil );
          destructor Destroy; override;
          property Qualifier : TvbName read fQualifier write fQualifier;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
      end;

    TvbDef =
      class( TvbNode )
        private
          fName : string;
          procedure SetName( const name : string );
        protected
          constructor Create( kind        : TvbNodeKind;
                              const flags : TvbNodeFlags = [];
                              const name  : string       = '';
                              const name1 : string       = '' );
        public
          property Name : string read fName write SetName;
      end;

    TvbDllFuncDef           = class;
    TvbRecordDef            = class;
    TvbEnumDef              = class;
    TvbEventDef             = class;
    TvbPropertyDef          = class;
    TvbImplementedInterface = class;
    TvbLibModuleDef         = class;

    TvbImplementedInterface =
      class( TvbNode )
        private
          fName  : TvbQualifiedName;
          fFuncs : TvbNodeList;
        protected
          function GetFunctionCount : integer;
          function GetFunction( i : integer ) : TvbFuncDef;
        public
          constructor Create( name : TvbQualifiedName );
          destructor Destroy; override;
          procedure AddImplementedFunction( func : TvbFuncDef );
          property Name                       : TvbQualifiedName read fName;
          property ImplFunctionCount          : integer          read GetFunctionCount;
          property ImplFunctions[i : integer] : TvbFuncDef       read GetFunction;
      end;

    TvbProjectKind = ( vbpEXE, vbpActiveX_EXE, vbpActiveX_DLL);

    TvbStdModuleDef = class;

    TvbProject =
      class( TvbNode )
        private
          fAutoIncrementVer       : string;
          fBoundsCheck            : boolean;
          fCompilationType        : byte;
          fExeName32              : string;
          fFavorPentiumPro        : boolean;
          fFDIVCheck              : boolean;
          fFloatPointCheck        : boolean;
          fHelpContextID          : string;
          fHelpFile               : string;
          fIconForm               : string;
          fImportedModules        : TvbNodeList;
          fKind                   : TvbProjectKind;
          fMajorVer               : integer;
          fMinorVer               : integer;
          fModules                : TvbNodeList;
          fName                   : string;
          fOptimizationType       : byte;
          fOverflowCheck          : boolean;
          fResFile32              : string;
          fRevisionVer            : integer;
          fScope                  : TvbScope;
          fStartup                : string;
          fTitle                  : string;
          fVersionComments        : string;
          fVersionCompanyName     : string;
          fVersionFileDescription : string;
          fVersionLegalCopyright  : string;
          fVersionLegalTrademarks : string;
          fVersionProductName     : string;
          function GetModuleCount : integer;
          function GetModule( i : integer ) : TvbModuleDef;
          function GetImportedModule( i : integer ) : TvbStdModuleDef;
          function GetImportedModuleCount : integer;
        public
          constructor Create;
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          function AddModule( m : TvbModuleDef ) : TvbModuleDef;
          function AddImportedModule( m : TvbStdModuleDef ) : TvbStdModuleDef;
          property ImportedModule[i : integer] : TvbStdModuleDef read GetImportedModule;
          property ImportedModuleCount         : integer         read GetImportedModuleCount;
          property Module[i : integer]         : TvbModuleDef    read GetModule;
          property ModuleCount                 : integer         read GetModuleCount;
          property Name                        : string          read fName                   write fName;
          property optAutoIncrementVer         : string          read fAutoIncrementVer       write fAutoIncrementVer;
          property optBoundsCheck              : boolean         read fBoundsCheck            write fBoundsCheck;
          property optCompilationType          : byte            read fCompilationType        write fCompilationType;
          property optExeName32                : string          read fExeName32              write fExeName32;
          property optFavorPentiumPro          : boolean         read fFavorPentiumPro        write fFavorPentiumPro;
          property optFDIVCheck                : boolean         read fFDIVCheck              write fFDIVCheck;
          property optFloatPointCheck          : boolean         read fFloatPointCheck        write fFloatPointCheck;
          property optHelpContextID            : string          read fHelpContextID          write fHelpContextID;
          property optHelpFile                 : string          read fHelpFile               write fHelpFile;
          property optIconForm                 : string          read fIconForm               write fIconForm;
          property optMajorVer                 : integer         read fMajorVer               write fMajorVer;
          property optMinorVer                 : integer         read fMinorVer               write fMinorVer;
          property optOptimizationType         : byte            read fOptimizationType       write fOptimizationType;
          property optOverflowCheck            : boolean         read fOverflowCheck          write fOverflowCheck;
          property optProjectKind              : TvbProjectKind  read fKind                   write fKind;
          property optResFile32                : string          read fResFile32              write fResFile32;
          property optRevisionVer              : integer         read fRevisionVer            write fRevisionVer;
          property optStartup                  : string          read fStartup                write fStartup;
          property optTitle                    : string          read fTitle                  write fTitle;
          property optVersionComments          : string          read fVersionComments        write fVersionComments;
          property optVersionCompanyName       : string          read fVersionCompanyName     write fVersionCompanyName;
          property optVersionFileDescription   : string          read fVersionFileDescription write fVersionFileDescription;
          property optVersionLegalCopyright    : string          read fVersionLegalCopyright  write fVersionLegalCopyright;
          property optVersionLegalTrademarks   : string          read fVersionLegalTrademarks write fVersionLegalTrademarks;
          property optVersionProductName       : string          read fVersionProductName     write fVersionProductName;
          property Scope                       : TvbScope        read fScope;
      end;

    TvbAbstractFuncDef = class;

    TvbModuleDef =
      class( TvbDef )
        protected
          fClasses    : TvbNodeList;
          fConsts     : TvbNodeList;
          fDllFuncs   : TvbNodeList;
          fEnums      : TvbNodeList;
          fRelPath    : string;
          fLibModules : TvbNodeList;
          fRecords    : TvbNodeList;
          fScope      : TvbScope;
          fProject    : TvbProject;
          fOptionBase : byte;
          constructor Create( nk : TvbNodeKind; const name, relativePath : string; ParentScope : TvbScope );
          function GetClass( i : integer ) : TvbClassDef; virtual;
          function GetClassCount : integer; virtual;
          function GetConst( i : integer ) : TvbConstDef; virtual;
          function GetConstCount : integer; virtual;
          function GetDllFunc( i : integer ) : TvbDllFuncDef; virtual;
          function GetDllFuncCount : integer; virtual;
          function GetEnum( i : integer ) : TvbEnumDef;
          function GetEnumCount : integer;
          function GetEvent( i : integer ) : TvbEventDef; virtual; abstract;
          function GetEventCount : integer; virtual; abstract;
          function GetLibModule( i : integer ) : TvbLibModuleDef; virtual;
          function GetLibModuleCount : integer; virtual;
          function GetMethod( i : integer ) : TvbAbstractFuncDef; virtual; abstract;
          function GetMethodCount : integer; virtual; abstract;
          function GetProperty(i: integer): TvbPropertyDef; virtual; abstract;
          function GetPropCount: integer; virtual; abstract;
          function GetRecord( i : integer) : TvbRecordDef;
          function GetRecordCount : integer;
          function GetVar( i : integer ) : TvbVarDef; virtual; abstract;
          function GetVarCount : integer; virtual; abstract;
        public
          destructor Destroy; override;
          function AddClass( c : TvbClassDef ) : TvbClassDef; virtual;
          function AddConstant( cst : TvbConstDef ) : TvbConstDef; virtual;
          function AddDllFunc( df : TvbDllFuncDef ) : TvbDllFuncDef; virtual;
          function AddEnum( en : TvbEnumDef ) : TvbEnumDef; virtual;
          function AddEvent( e : TvbEventDef ) : TvbEventDef; virtual; abstract;
          function AddLibModule( lm : TvbLibModuleDef ) : TvbLibModuleDef; virtual;
          function AddMethod( m : TvbAbstractFuncDef ) : TvbAbstractFuncDef; virtual; abstract;
          function AddProperty( p : TvbPropertyDef ) : TvbPropertyDef; virtual; abstract;
          function AddRecord( s : TvbRecordDef ) : TvbRecordDef; virtual;
          function AddVar( v : TvbVarDef ) : TvbVarDef; virtual; abstract;
          property ClassCount             : integer            read GetClassCount;
          property Classes[i : integer]   : TvbClassDef        read GetClass;
          property ConstCount             : integer            read GetConstCount;
          property Consts[i : integer]    : TvbConstDef        read GetConst;
          property DllFuncCount           : integer            read GetDllFuncCount;
          property DllFuncs[i : integer]  : TvbDllFuncDef      read GetDllFunc;
          property Enum[i : integer]      : TvbEnumDef         read GetEnum;
          property EnumCount              : integer            read GetEnumCount;
          property Event[i : integer]     : TvbEventDef        read GetEvent;
          property EventCount             : integer            read GetEventCount;
          property LibModule[i : integer] : TvbLibModuleDef    read GetLibModule;
          property LibModuleCount         : integer            read GetLibModuleCount;
          property Method[i : integer]    : TvbAbstractFuncDef read GetMethod;
          property MethodCount            : integer            read GetMethodCount;
          property OptionBase             : byte               read fOptionBase write fOptionBase;
          property Prop[i : integer]      : TvbPropertyDef     read GetProperty;
          property PropCount              : integer            read GetPropCount;
          property RecordCount            : integer            read GetRecordCount;
          property Records[i : integer]   : TvbRecordDef       read GetRecord;
          property RelativePath           : string             read fRelPath    write fRelPath;
          property Scope                  : TvbScope           read fScope;
          property VarCount               : integer            read GetVarCount;
          property Vars[i : integer]      : TvbVarDef          read GetVar;
      end;

    TvbStdModuleDef =
      class( TvbModuleDef )
        protected
          fMethods : TvbNodeList;
          fProps   : TvbNodeList;
          fVars    : TvbNodeList;
          function GetEvent( i : integer ) : TvbEventDef; override;
          function GetEventCount : integer; override;
          function GetMethod( i : integer ) : TvbAbstractFuncDef; override;
          function GetMethodCount : integer; override;
          function GetProperty(i: integer): TvbPropertyDef; override;
          function GetPropCount: integer; override;
          function GetVar( i : integer ) : TvbVarDef; override;
          function GetVarCount : integer; override;
        public
          constructor Create( const name, relativePath : string; ParentScope : TvbScope = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          function AddEvent( e : TvbEventDef ) : TvbEventDef; override;
          function AddMethod( m : TvbAbstractFuncDef ) : TvbAbstractFuncDef; override;
          function AddProperty( p : TvbPropertyDef ) : TvbPropertyDef; override;
          function AddVar( v : TvbVarDef ) : TvbVarDef; override;
      end;

    TvbClassModuleDef =
      class( TvbModuleDef )
        protected
          fClassDef : TvbClassDef;
          function GetEvent( i : integer ) : TvbEventDef; override;
          function GetEventCount : integer; override;
          function GetMethod( i : integer ) : TvbAbstractFuncDef; override;
          function GetMethodCount : integer; override;
          function GetProperty(i: integer): TvbPropertyDef; override;
          function GetPropCount: integer; override;
          function GetVar( i : integer ) : TvbVarDef; override;
          function GetVarCount : integer; override;
          function CreateClassInstance( const name : string ) : TvbClassDef; virtual;
        public
          constructor Create( const name, relativePath : string; ParentScope : TvbScope );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          function AddEvent( e : TvbEventDef ) : TvbEventDef; override;
          function AddMethod( m : TvbAbstractFuncDef ) : TvbAbstractFuncDef; override;
          function AddProperty( p : TvbPropertyDef ) : TvbPropertyDef; override;
          function AddVar( v : TvbVarDef ) : TvbVarDef; override;
          property ClassDef : TvbClassDef read fClassDef;
      end;

    TvbFormDef = class;

    TvbFormModuleDef =
      class( TvbClassModuleDef )
        private
          fFormDef : TvbFormDef;
          fFormVar : TvbVarDef;
        protected
          function CreateClassInstance( const name : string ) : TvbClassDef; override;
        public
          constructor Create( const name, relativePath : string; ParentScope : TvbScope );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          // this is the definition of the form class type.
          property FormDef : TvbFormDef read fFormDef;
          property FormVar : TvbVarDef  read fFormVar;
      end;

    TvbExpr =
      class( TvbNode )
        private
          fDef   : TvbDef;
          fParen : boolean;
        protected
          fOp1 : TvbExpr;
          fOp2 : TvbExpr;
          constructor Create( nk : TvbNodeKind; op1 : TvbExpr ); overload;
          constructor Create( nk : TvbNodeKind; op1, op2 : TvbExpr ); overload;
          function GetIsRValue : boolean; virtual;
          procedure SetIsRValue( const value : boolean ); virtual;
          function GetIsUsedWithEmptyParens : boolean; virtual;
          procedure SetIsUsedWithEmptyParens( const Value : boolean ); virtual;
          function GetIsUsedWithArgs : boolean; virtual;
          procedure SetIsUsedWithArgs( const Value : boolean ); virtual;
        public
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property Op1           : TvbExpr read fOp1; // first operand(used by unary and binary exprs)
          property Op2           : TvbExpr read fOp2; // second operand(only used by binary exprs)
          property ExprType      : IvbType read fNodeType;
          property Parenthesized : boolean read fParen write fParen;
          // This is the definition that this name refers to(if applicable).
          // Note that this is not the same as the expression's type.
          property Def  : TvbDef  read fDef write fDef;
          // Tells/sets whether the expression is the LHS of an assignment.
          property IsRValue : boolean read GetIsRValue write SetIsRValue;
          //property IsConstant : boolean read GetIsConstant write SetIsConstant;
          property IsUsedWithArgs        : boolean read GetIsUsedWithArgs        write SetIsUsedWithArgs;
          property IsUsedWithEmptyParens : boolean read GetIsUsedWithEmptyParens write SetIsUsedWithEmptyParens;
      end;

    // LBound( array )
    TvbLBoundExpr =
      class( TvbExpr )
        public
          constructor Create( arr : TvbExpr; dim : TvbExpr = nil );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property ArrayVar : TvbExpr read fOp1 write fOp1;
          property Dim      : TvbExpr read fOp2 write fOp2;
      end;

    // UBound( array )
    TvbUBoundExpr =
      class( TvbExpr )
        public
          constructor Create( arr : TvbExpr; dim : TvbExpr = nil );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property ArrayVar : TvbExpr read fOp1 write fOp1;
          property Dim      : TvbExpr read fOp2 write fOp2;
      end;

    TvbCastTo = (
      BOOLEAN_CAST,
      BYTE_CAST,
      CURRENCY_CAST,
      DATE_CAST,
      DECIMAL_CAST,
      DOUBLE_CAST,
      INTEGER_CAST,
      LONG_CAST,
      SINGLE_CAST,
      STRING_CAST,
      VARIANT_CAST,
      VDATE_CAST,
      VERROR_CAST);

    TvbCastExpr =
      class( TvbExpr )
        private
          fConvertTo : TvbCastTo;
        public
          constructor Create( expr : TvbExpr; conv : TvbCastTo );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property CastTo : TvbCastTo read fConvertTo;
      end;

    TvbExprList = class( TvbNodeList );

    // Dim
    TvbVarDef =
      class( TvbDef )
        public
          constructor Create( const flags : TvbNodeFlags = [];
                              const name  : string       = '';
                              const typ   : IvbType   = nil;
                              const name1 : string       = ''); 
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property VarType : IvbType read fNodeType write fNodeType;
      end;

    // Const
    TvbConstDef =
      class( TvbDef )
        private
          fValue : TvbExpr;
        public
          constructor Create( const flags : TvbNodeFlags    = [];
                              const name  : string          = '';
                              const typ   : IvbType      = nil;
                              val         : TvbExpr         = nil;
                              const name1 : string          = '' ); 
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property ConstType : IvbType read fNodeType write fNodeType;
          property Value     : TvbExpr    read fValue    write fValue;
      end;

    // [Optional, ParamArray, ByVal/ByRef] paramName [As paramType] [=defVal]
    TvbParamDef =
      class( TvbDef )
        private
          fDefVal : TvbExpr;
        public
          constructor Create( flags      : TvbNodeFlags = [pfByRef];
                              const name : string       = '';
                              const typ  : IvbType   = nil;
                              defVal     : TvbExpr      = nil ); overload;
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          // These are the VB flags specified in the parameter declaration.
          property ParamType    : IvbType read fNodeType write fNodeType;
          property DefaultValue : TvbExpr    read fDefVal   write fDefVal;
      end;

    TvbParamList =
      class
        private
          fOptParamCount : integer;
          fOwnsParams    : boolean;
          fParamArrayIdx : integer;
          fParams        : TStringList;
          function GetParam( i : integer ) : TvbParamDef;
          function GetParamByName( n : string ) : TvbParamDef;
          function GetParamCount : integer;
        public
          constructor Create( OwnsParams : boolean );
          //constructor Create( p : TvbParamDef ); overload;
          destructor Destroy; override;
          procedure AddParam( p : TvbParamDef );
          // Number of optional parameters.
          property OptParamCount : integer read fOptParamCount;
          // Index of the parameter flagged with ParamArray.
          // It's -1 if there is none.
          property ParamArrayIndex : integer read fParamArrayIdx;
          property ParamByName[n : string] : TvbParamDef read GetParamByName;
          property ParamCount              : integer     read GetParamCount;
          property Params[i : integer]     : TvbParamDef read GetParam; default;
      end;

    TvbStmt = class;
    TvbLabelDef = class;

    TvbStmtBlock = class;

    // Represents the definition of a subroutine/function.
    TvbAbstractFuncDef =
      class( TvbDef )
        protected
          function GetConst( i : integer ) : TvbConstDef; virtual;
          function GetConstCount : integer; virtual;
          function GetVar( i : integer ) : TvbVarDef; virtual;
          function GetVarCount : integer; virtual;
          function GetLabel( i : integer ) : string; virtual;
          function GetLabelCount : integer; virtual;
          function GetParamArrayIndex : integer; virtual;
          function GetParam(i: integer): TvbParamDef; virtual;
          function GetParamCount: integer; virtual;
          function GetScope : TvbScope; virtual;
          function GetStaticVarCount : integer; virtual;
          function GetStmtBlock : TvbStmtBlock; virtual;
          function GetParams : TvbParamList; virtual;
          procedure SetStmtBlock( block : TvbStmtBlock ); virtual;
        public
          function GetArgc : integer; virtual;
          procedure AddLabel( lab : TvbLabelDef ); virtual;
          procedure AddLocalConst( cst : TvbConstDef ); virtual;
          procedure AddLocalVar( const name : string; typ : TvbType ); overload; virtual;
          procedure AddLocalVar( v : TvbVarDef ); overload; virtual;
          procedure AddParam( p : TvbParamDef ); virtual;
          property Argc                : integer      read GetArgc;
          property ConstCount          : integer      read GetConstCount;
          property Consts[i : integer] : TvbConstDef  read GetConst;
          property LabelCount          : integer      read GetLabelCount;
          property Labels[i : integer] : string       read GetLabel;
          property Param[i : integer]  : TvbParamDef  read GetParam;
          // Index of the one and only parameter flagged with ParamArray.
          // This is -1 if there is none.
          property ParamArrayIndex     : integer      read GetParamArrayIndex;
          property ParamCount          : integer      read GetParamCount;
          property Params              : TvbParamList read GetParams;
          property ReturnType          : IvbType      read fNodeType    write fNodeType;
          property Scope               : TvbScope     read GetScope;
          property StaticVarCount      : integer      read GetStaticVarCount;
          property StmtBlock           : TvbStmtBlock read GetStmtBlock write SetStmtBlock;
          property VarCount            : integer      read GetVarCount;
          property Vars[i : integer]   : TvbVarDef    read GetVar;
      end;

    TvbFuncDef =
      class( TvbAbstractFuncDef )
        private
          fBlock          : TvbStmtBlock;
          fLabels         : TStringList;
          fLocalConsts    : TStringList;
          fLocalVars      : TStringList;
          fParams         : TvbParamList;
          fScope          : TvbScope;
          fStaticVarCount : integer;
        protected
          function GetConst( i : integer ) : TvbConstDef; override;
          function GetConstCount : integer; override;
          function GetVar( i : integer ) : TvbVarDef; override;
          function GetVarCount : integer; override;
          //function GetVarByName( const n : string ) : TvbVarDef; override;
          //function GetConstByName( const n : string ) : TvbConstDef; override;
          function GetLabel( i : integer ) : string; override;
          function GetLabelCount : integer; override;
          function GetParamArrayIndex : integer; override;
          function GetParam(i: integer): TvbParamDef; override;
          function GetParamCount: integer; override;
          function GetParams : TvbParamList; override;
          function GetScope : TvbScope; override;
          function GetStaticVarCount : integer; override;
          function GetStmtBlock : TvbStmtBlock; override;
          procedure SetStmtBlock( block : TvbStmtBlock ); override;
        public
          constructor Create( flags       : TvbNodeFlags = [];
                              const name  : string       = '';
                              params      : TvbParamList = nil;
                              const typ   : IvbType   = nil;
                              stmts       : TvbStmtBlock = nil;
                              const name1 : string       = '' );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          function GetArgc : integer; override;
          procedure AddLabel( lab : TvbLabelDef ); override;
          procedure AddLocalConst( cst : TvbConstDef ); override;
          procedure AddLocalVar( const name : string; typ : TvbType ); overload; override;
          procedure AddLocalVar( v : TvbVarDef ); overload; override;
          procedure AddParam( p : TvbParamDef ); override;
      end;

    // Represents the definition of a function exported from a DLL.
    TvbDllFuncDef =
      class( TvbAbstractFuncDef )
        private
          fParams : TvbParamList;
          fDllName : string;
          fAlias   : string;
        protected
          function GetParam( i : integer ) : TvbParamDef; override;
          function GetParamCount : integer; override;
          function GetParams : TvbParamList; override;
        public
          constructor Create( flags         : TvbNodeFlags = [];
                              const name    : string       = '';
                              const dllName : string       = '';
                              const alias   : string       = '';
                              params        : TvbParamList = nil;
                              const typ     : IvbType      = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          procedure AddParam( p : TvbParamDef ); override;
          property DllName : string read fDllName write fDllName;
          property Alias   : string read fAlias   write fAlias;
      end;

    TvbExternalFuncDef =
      class( TvbAbstractFuncDef )
        private
          fArgc : integer;
        public
          constructor Create( flags       : TvbNodeFlags = [];
                              const name  : string       = '';
                              const name1 : string       = '';
                              argc        : integer      = 0;
                              const typ   : IvbType      = nil );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;                              
          function GetArgc : integer; override;
      end;

    TvbPropertyDef =
      class( TvbDef )
        private
          fGetterFunc : TvbAbstractFuncDef;
          fSetterFunc : TvbAbstractFuncDef;
          fLetterFunc : TvbAbstractFuncDef;
          fParams     : TvbParamList;
          function GetParam( i : integer ) : TvbParamDef;
          function GetParamCount : integer;
        public
          constructor Create( const name : string; flags : TvbNodeFlags = [] );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          procedure SetParamsAndType;
          property GetFn              : TvbAbstractFuncDef read fGetterFunc write fGetterFunc;
          property LetFn              : TvbAbstractFuncDef read fLetterFunc write fLetterFunc;
          property Param[i : integer] : TvbParamDef        read GetParam;
          property ParamCount         : integer            read GetParamCount;
          property Params             : TvbParamList       read fParams;
          property PropType           : IvbType            read fNodeType   write fNodeType;
          property SetFn              : TvbAbstractFuncDef read fSetterFunc write fSetterFunc;
      end;

    TvbTypeDef =
      class( TvbDef )
        protected
          fScope : TvbScope;
          constructor Create( kind        : TvbNodeKind;
                              const flags : TvbNodeFlags = [];
                              const name  : string       = '';
                              const name1 : string       = '' );
        public
          destructor Destroy; override;
          property Scope : TvbScope read fScope;
      end;

    TvbClassDef =
      class( TvbTypeDef )
        private
          fBaseClass       : IvbType;
          fClassInitialize : TvbAbstractFuncDef;
          fClassTerminate  : TvbAbstractFuncDef;
          fEvents          : TvbNodeList;
          fMethods         : TvbNodeList;
          fMeType          : IvbType;
          fProps           : TvbNodeList;
          fVars            : TvbNodeList;
          function GetEventCount : integer;
          function GetEvents( i : integer ) : TvbEventDef;
          function GetMethod( i : integer ) : TvbAbstractFuncDef;
          function GetMethodCount : integer;
          function GetProperty( i : integer ) : TvbPropertyDef;
          function GetPropCount : integer;
          function GetVarCount : integer;
          function GetVar( i : integer ) : TvbVarDef;
        public
          constructor Create( const flags     : TvbNodeFlags = [];
                              const name      : string       = '';
                              const name1     : string       = '';
                              const baseClass : IvbType      = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          procedure AddEvent( e : TvbEventDef );
          procedure AddMethod( m : TvbAbstractFuncDef );
          procedure AddProperty( p : TvbPropertyDef );
          procedure AddVar( v : TvbVarDef );
          property BaseClass           : IvbType            read fBaseClass       write fBaseClass;
          property ClassInitialize     : TvbAbstractFuncDef read fClassInitialize write fClassInitialize;
          property ClassTerminate      : TvbAbstractFuncDef read fClassTerminate  write fClassTerminate;
          property Event[i : integer]  : TvbEventDef        read GetEvents;
          property EventCount          : integer            read GetEventCount;
          property Method[i : integer] : TvbAbstractFuncDef read GetMethod;
          property MethodCount         : integer            read GetMethodCount;
          property MeType              : IvbType            read fMeType;
          property Prop[i : integer]   : TvbPropertyDef     read GetProperty;
          property PropCount           : integer            read GetPropCount;
          property VarCount            : integer            read GetVarCount;
          property Vars[i : integer]   : TvbVarDef          read GetVar;
      end;

    TvbFormControlList = TvbNodeList;

    TvbControlPropertyKind = ( cpkSimple, cpkControl, cpkResource );

    TvbControl = class;

    TvbControlProperty =
      class
        private
          fControl : TvbControl;
          fKind    : TvbControlPropertyKind;
          fName    : string;
          fValue   : olevariant;
        public
          constructor Create( kind : TvbControlPropertyKind; const name : string = '' );
          destructor Destroy; override;
          property Name : string read fName write fName;
          property Kind : TvbControlPropertyKind read fKind;
          // value for simple properties and the offset into the FRX file.
          property Value : olevariant read fValue write fValue;
          // for when property is a control.
          property Control : TvbControl read fControl write fControl;
      end;

    TvbControl =
      class( TvbVarDef )
        private
          //fParent   : TvbControl;
          fControls : TStringList;
          fProps    : TStringList;
          function GetControl( i : integer ) : TvbControl;
          function GetControlCount : integer;
          function GetProp( i : integer ) : TvbControlProperty;
          function GetPropCount : integer;
          function GetPropByName( const name : string ) : TvbControlProperty;
        public
          constructor Create( const name : string; typ : IvbType );
          destructor Destroy; override;
          // adds a control property(BeginProperty) to this control.
          // checks if another property with the same name exists and returns it. 
          function AddObjectProperty( const name : string ) : TvbControl;
          function AddProperty( const name : string; kind : TvbControlPropertyKind = cpkSimple ) : TvbControlProperty;
          procedure AddControl( ctl : TvbControl );
          property Control[i : integer]            : TvbControl         read GetControl;
          property ControlCount                    : integer            read GetControlCount;
          property Prop[i : integer]               : TvbControlProperty read GetProp;
          property PropByName[const name : string] : TvbControlProperty read GetPropByName;
          property PropCount                       : integer            read GetPropCount;
      end;

    TvbFormDef =
      class( TvbClassDef )
        private
          fRootControl : TvbControl;
          // this is simple a map with all the controls defined in the form.
          fControls : TStringList;
          function GetControl( i : integer ) : TvbControl;
          function GetControlCount : integer;
        public
          constructor Create( const flags : TvbNodeFlags = [];
                              const name  : string       = '';
                              const name1 : string       = '' );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          procedure OnNewControl( control : TvbControl );
          // this is the form description which is a tree of parent/child relationships
          // among controls.
          // the type of this control should be "VB.Form" or "VB.MDIForm". 
          property RootControl : TvbControl read fRootControl write fRootControl;
          property Control[i : integer] : TvbControl read GetControl;
          property ControlCount         : integer    read GetControlCount;
      end;

    // Represents the definition of a record/struct.
    TvbRecordDef =
      class( TvbTypeDef )
        private
          fFields : TvbNodeList;
          function GetField( i : integer ) : TvbVarDef;
          function GetFieldCount : integer;
        public
          constructor Create( flags       : TvbNodeFlags = [];
                              const name  : string       = '';
                              const name1 : string       = '' ); 
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          procedure AddField( f : TvbVarDef );
          property FieldCount         : integer   read GetFieldCount;
          property Field[i : integer] : TvbVarDef read GetField;
      end;

    // Represents the definition of an enumerated type.
    // It consists of a group of integer constants.
    TvbEnumDef =
      class( TvbTypeDef )
        private
          fEnums : TvbNodeList;
          function GetEnum( i : integer ) : TvbConstDef;
          function GetEnumCount : integer;
        public
          constructor Create( flags       : TvbNodeFlags = [];
                              const name  : string       = '';
                              const name1 : string       = '' ); 
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          procedure AddEnum( e : TvbConstDef );
          property EnumCount         : integer     read GetEnumCount;
          property Enum[i : integer] : TvbConstDef read GetEnum;
      end;

    // Represents the definition of a MODULE from an external library.
    // A MODULE consist of a group of static functions or constants.
    TvbLibModuleDef  =
      class( TvbDef )
        private
          fConsts : TvbNodeList;
          fFuncs  : TvbNodeList;
          fProps  : TvbNodeList;
          fScope  : TvbScope;
          fVars   : TvbNodeList;
          function GetConst(i: integer): TvbConstDef;
          function GetConstCount: integer;
          function GetFunction(i: integer): TvbAbstractFuncDef;
          function GetFunctionCount: integer;
          function GetProperty( i : integer ) : TvbPropertyDef;
          function GetPropCount : integer;
          function GetVarCount: integer;
          function GetVar( i : integer ) : TvbVarDef;
        public
          constructor Create( const flags : TvbNodeFlags = [];
                              const name  : string       = '';
                              const name1 : string       = '' );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          procedure AddConstant( cst : TvbConstDef );
          procedure AddMethod( func : TvbAbstractFuncDef );
          procedure AddProperty( p : TvbPropertyDef );
          procedure AddVar( vr : TvbVarDef );
          property ConstCount             : integer            read GetConstCount;
          property Consts[i : integer]    : TvbConstDef        read GetConst;
          property FunctionCount          : integer            read GetFunctionCount;
          property Functions[i : integer] : TvbAbstractFuncDef read GetFunction;
          property Prop[i : integer]      : TvbPropertyDef     read GetProperty;
          property PropCount              : integer            read GetPropCount;
          property Scope                  : TvbScope           read fScope;
          property VarCount               : integer            read GetVarCount;
          property Vars[i : integer]      : TvbVarDef          read GetVar;
      end;

    // Represents the definition of an event in a class module.
    TvbEventDef =
      class( TvbDef )
        private
          fParams : TvbParamList;
          function GetParam( i : integer ) : TvbParamDef;
          function GetParamCount : integer;
        public
          constructor Create( const name : string = ''; params : TvbParamList = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          procedure AddParam( p : TvbParamDef );
          property Params             : TvbParamList read fParams write fParams;
          property ParamCount         : integer      read GetParamCount;
          property Param[i : integer] : TvbParamDef  read GetParam;
      end;

    TvbLabelDef =
      class( TvbDef )
        private
          fPrevLabel : TvbLabelDef;
          fNextLabel : TvbLabelDef;
        public
          constructor Create( const name : string );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          // The label that follows this one if any.
          property PrevLabel : TvbLabelDef read fPrevLabel write fPrevLabel;
          // The label that precedes this one if any.
          property NextLabel : TvbLabelDef read fNextLabel write fNextLabel;
      end;

    TvbArrayBounds =
      class( TvbNode )
        private
          fLower : TvbExpr;
          fUpper : TvbExpr;
        public
          constructor Create( l, u : TvbExpr );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property Lower : TvbExpr read fLower;
          property Upper : TvbExpr read fUpper;
      end;

    TvbArrayBoundList = TvbNodeList;

    IvbType =
      interface
        ['{17F9CB98-0296-4098-8A08-446568EC0BFE}']
        function GetArrayBoundCount : integer;
        function GetArrayBounds( i : integer) : TvbArrayBounds;
        function GetCode  : cardinal;
        function GetName  : string;
        function GetName1 : string;
        procedure SetName1( const value : string );
        function GetStringLen : TvbExpr;
        function GetTypeDef : TvbTypeDef;
        procedure SetTypeDef( td : TvbTypeDef );
        function GetParentModule : TvbModuleDef;
        // The number of bounds for an static array type.
        // Must be 0 for a dynamic array.
        property ArrayBoundCount          : integer        read GetArrayBoundCount;
        property ArrayBounds[i : integer] : TvbArrayBounds read GetArrayBounds;
        // The size of an static string.
        // Must be nil for a dynamic string.
        property StringLen : TvbExpr read GetStringLen;
        // This is the definition of an user-defined type(record/enum or class).
        property TypeDef      : TvbTypeDef   read GetTypeDef write SetTypeDef;
        property Code         : cardinal     read GetCode;
        property Name         : string       read GetName;
        property Name1        : string       read GetName1 write SetName1;
        property ParentModule : TvbModuleDef read GetParentModule;
      end;

    TvbType =
      class( TInterfacedObject, IvbType )
        private
          fBounds : TvbArrayBoundList;
          fModule : TvbModuleDef;
          fDef    : TvbTypeDef;
          fStrLen : TvbExpr;
          fCode   : cardinal;
          fName   : string;
          fName1  : string;
        public
          constructor Create( typeCode : cardinal );
          // Create an array of the specified intrincic type.
          constructor CreateArray( baseType : cardinal ); overload;
          // Create an array of an UDT given its name.
          constructor CreateArray( const typeName : string ); overload;
          // Create an array of an UDT given its type definition.
          constructor CreateArray( typeDef : TvbTypeDef ); overload;
          // Create a vector of the base type specified.
          constructor CreateVector( const baseType : IvbType; bounds : TvbArrayBoundList );
          // Create an static string type.
          constructor CreateString( len : TvbExpr = nil );
          // Create an UDT given its name.
          constructor CreateUDT( const typeName : string ); overload;
          // Create an UDT given its type definition.
          constructor CreateUDT( typeDef : TvbTypeDef ); overload;
          destructor Destroy; override;
          procedure AddBound( lower, upper : TvbExpr );
          // IvbType
          function GetArrayBoundCount : integer;
          function GetArrayBounds( i : integer ) : TvbArrayBounds;
          function GetStringLen: TvbExpr;
          function GetTypeDef: TvbTypeDef;
          function GetCode: cardinal;
          function GetName: string;
          procedure SetTypeDef( td : TvbTypeDef );
          function GetName1 : string;
          procedure SetName1( const value : string );
          function GetParentModule : TvbModuleDef;          
      end;

    TvbOperator = (
      ADD_OP,         // addition
      AND_OP,         // logical and bitwise AND
      DIV_OP,         // division
      EQ_OP,          // assignment and equality test
      EQV_OP,         // bitwise equivalence
      GREATER_EQ_OP,  // greater or equal
      GREATER_OP,     // greater
      IMP_OP,         // bitwise implication
      INT_DIV_OP,     // integer division
      IS_OP,          // equality
      LESS_EQ_OP,     // less or equaal
      LESS_OP,        // less
      LIKE_OP,        // string pattern match
      MOD_OP,         // division remainder
      MULT_OP,        // multiplication
      NEW_OP,         // instance creation
      NOT_EQ_OP,      // not equal(different)
      NOT_OP,         // logical and bitwise NOT
      OR_OP,          // logical and bitwise OR
      POWER_OP,       // power
      STR_CONCAT_OP,  // string concatenation
      SUB_OP,         // substraction
      XOR_OP          // bitwise exclusive OR
      );

    TvbArgList =
      class
        private
          fArgIndex : TStringList;
          fArgs     : TvbNodeList;
          fFirstNamedArgIdx : integer;
          function GetArg( i : integer ) : TvbExpr;
          function GetNamedArg( const name : string ) : TvbExpr;
          function GetArgCount : integer;
        public
          constructor Create;
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil );
          procedure Add( arg : TvbExpr );
          procedure Resolve( CurrentScope : TvbScope );
          property Arg[i : integer]              : TvbExpr read GetArg; default;
          property ArgCount                      : integer read GetArgCount;
          property FirstNamedArgIndex            : integer read fFirstNamedArgIdx;
          property NamedArg[const name : string] : TvbExpr read GetNamedArg;
      end;

    // func()
    // func(expr1, expr2)
    // arr(index)
    TvbCallOrIndexerExpr =
      class( TvbExpr )
        private
          fArgs : TvbArgList;
        protected
          function GetIsRValue : boolean; override;
          procedure SetIsRValue( const value : boolean ); override;
          function GetIsUsedWithEmptyParens : boolean; override;
          procedure SetIsUsedWithEmptyParens( const Value : boolean ); override;
          function GetIsUsedWithArgs : boolean; override;
          procedure SetIsUsedWithArgs( const Value : boolean ); override;
        public
          constructor Create( funcOrArray : TvbExpr; args : TvbArgList = nil; isFunc : boolean = false );
          destructor Destroy; override;
          procedure AddArg( arg : TvbExpr );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property FuncOrVar : TvbExpr    read fOp1  write fOp1;
          property Args      : TvbArgList read fArgs write fArgs;
      end;

    // AddressOf funcName
    TvbAddressOfExpr =
      class( TvbExpr )
        private
          fFuncName : TvbName;
        public
          constructor Create( aFuncName : TvbName );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property FuncName : TvbName read fFuncName;
      end;

    TvbUnaryExpr =
      class( TvbExpr )
        private
          fOperator : TvbOperator;
        public
          constructor Create( op : TvbOperator; expr : TvbExpr );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property Oper : TvbOperator read fOperator;
          property Expr : TvbExpr     read fOp1;
      end;

    // TypeOf objectName Is typeName
    TvbTypeOfExpr =
      class( TvbExpr )
        private
          fTypeName : TvbName;
        public
          constructor Create( obj : TvbExpr; typeName : TvbName );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property ObjectExpr : TvbExpr read fOp1;
          property Name   : TvbName read fTypeName;
      end;

    TvbBinaryExpr =
      class( TvbExpr )
        private
          fOperator : TvbOperator;
        public
          constructor Create( op : TvbOperator; left, right : TvbExpr; paren : boolean = false );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property Oper      : TvbOperator read fOperator;
          property LeftExpr  : TvbExpr     read fOp1;
          property RightExpr : TvbExpr     read fOp2;
      end;

    // var
    TvbNameExpr =
      class( TvbExpr )
        private
          fName : string;
        public
          constructor Create( const name : string );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property Name : string  read fName;
      end;

    // var.member
    // .member
    TvbMemberAccessExpr =
      class( TvbExpr )
        private
          fMember : string;
          fWithExpr : TvbExpr;
        public
          constructor Create( const aMember : string;
                              aTarget       : TvbExpr = nil;
                              withExpr      : TvbExpr = nil );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property Member : string read fMember;
          // When Target is nil, the target is the expression of the
          // most recent "With" statement.
          property Target   : TvbExpr read fOp1;
          property WithExpr : TvbExpr read fWithExpr write fWithExpr;
      end;

    // DictAccess: Target!Name => Target.DefaultProperty["Name"]
    //
    TvbDictAccessExpr =
      class( TvbExpr )
        private
          fName     : string;
          fWithExpr : TvbExpr;
        public
          constructor Create( const name : string;
                              aTarget    : TvbExpr = nil;
                              withExpr   : TvbExpr = nil );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property Name     : string  read fName write fName;
          property Target   : TvbExpr read fOp1;
          property WithExpr : TvbExpr read fOp2  write fOp2;
      end;

    // Me
    TvbMeExpr =
      class( TvbExpr )
        public
          constructor Create( typ : IvbType = nil );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
      end;

    TvbFuncResultExpr =
      class( TvbExpr )
        public
          constructor Create( typ : IvbType );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
      end;

    // Type suffixes for identifiers and numeric literals.
    // NOTE: DO NOT alter the declaration order because we build a lookup
    // table from this.
    //
    TvbTypeSfx = (
      NO_SFX,       // no suffix at all
      STRING_SFX,   // string type suffix
      INTEGER_SFX,  // integer type suffix
      LONG_SFX,     // long type suffix
      SINGLE_SFX,   // single type suffix
      DOUBLE_SFX,   // double type suffix
      CURRENCY_SFX  // currency type suffix
      );

    TvbIntLit =
      class( TvbExpr )
        private
          fValue : integer;
        public
          constructor Create( val  : integer );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property Value   : integer    read fValue;
      end;

    TvbFloatLit =
      class( TvbExpr )
        private
          fValue : double;
        public
          constructor Create( val : double );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property Value   : double read fValue;
      end;

    TvbNothingLit =
      class( TvbExpr )
        public
          constructor Create;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
      end;

    // True
    // False
    TvbBoolLit =
      class( TvbExpr )
        private
          fValue : boolean;
        public
          constructor Create( val : boolean );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property Value : boolean read fValue;
      end;

    TvbStringLit =
      class( TvbExpr )
        private
          fValue : string;
        public
          constructor Create( const val : string );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property Value : string read fValue;
      end;

    TvbDateLit =
      class( TvbExpr )
        private
          fValue : TDateTime;
        public
          constructor Create( const val : TDateTime );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property Value : TDateTime read fValue;
      end;

    // New className
    TvbNewExpr =
      class( TvbExpr )
        public
          constructor Create( typ : TvbExpr ); overload;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property TypeExpr : TvbExpr read fOp1 write fOp1;
      end;

    // Mid$(string, start[,position])
    TvbMidExpr =
      class( TvbExpr )
        private
          fLen : TvbExpr;
        public
          constructor Create( str   : TvbExpr = nil;
                              start : TvbExpr = nil;
                              len   : TvbExpr = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property Len   : TvbExpr read fLen write fLen;
          property Start : TvbExpr read fOp2 write fOp2;
          property Str   : TvbExpr read fOp1 write fOp1;
      end;

    TvbStringExpr =
      class( TvbExpr )
        public
          constructor Create( count : TvbExpr = nil; ch : TvbExpr = nil );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property Char  : TvbExpr read fOp2 write fOp2;
          property Count : TvbExpr read fOp1 write fOp1;
      end;

    TvbIntExpr =
      class( TvbExpr )
        public
          constructor Create( num : TvbExpr = nil );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property Number : TvbExpr read fOp1 write fOp1;
      end;

    TvbFixExpr =
      class( TvbExpr )
        public
          constructor Create( num : TvbExpr = nil );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property Number : TvbExpr read fOp1 write fOp1;
      end;

    TvbAbsExpr =
      class( TvbExpr )
        public
          constructor Create( num : TvbExpr = nil );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property Number : TvbExpr read fOp1 write fOp1;
      end;

    TvbLenExpr =
      class( TvbExpr )
        public
          constructor Create( strOrVar : TvbExpr = nil );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property StrOrVar : TvbExpr read fOp1 write fOp1;
      end;

    TvbLenBExpr =
      class( TvbExpr )
        public
          constructor Create( strOrVar : TvbExpr = nil );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property StrOrVar : TvbExpr read fOp1 write fOp1;
      end;

    TvbSgnExpr =
      class( TvbExpr )
        public
          constructor Create( num : TvbExpr = nil );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property Number : TvbExpr read fOp1 write fOp1;
      end;

    TvbSeekExpr =
      class( TvbExpr )
        public
          constructor Create( fileNum : TvbExpr = nil );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property FileNum : TvbExpr read fOp1 write fOp1;
      end;

    TvbDateExpr =
      class( TvbExpr )
        public
          constructor Create;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
      end;

    TvbArrayExpr =
      class( TvbExpr )
        private
          fValues : TvbExprList;
          function GetValue( i : integer ) : TvbExpr;
          function GetValueCount : integer;
        public
          constructor Create;
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          procedure AddValue( val : TvbExpr );
          property ValueCount          : integer read GetValueCount;
          property Values[i : integer] : TvbExpr read GetValue;
      end;

    TvbInputExpr =
      class( TvbExpr )
        public
          constructor Create;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property CharCount : TvbExpr read fOp1 write fOp1;
          property FileNum   : TvbExpr read fOp2 write fOp2;
      end;

    TvbNamedArgExpr =
      class( TvbExpr )
        private
          fArgName : string;
        public
          constructor Create( const argName : string; value : TvbExpr );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property ArgName : string  read fArgName write fArgName;
          property Value   : TvbExpr read fOp1     write fOp1;
      end;

    TvbDoEventsExpr =
      class( TvbExpr )
        public
          constructor Create;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
      end;

    TvbStmt =
      class( TvbNode )
        private
          fLabel    : TvbLabelDef;
          fNextStmt : TvbStmt;
          fPrevStmt : TvbStmt;
        public
          // The label(s) in front of this statement if any.
          property FrontLabel : TvbLabelDef read fLabel write fLabel;
          // The statement that follows this one in the block.
          property NextStmt : TvbStmt read fNextStmt write fNextStmt;
          // The statement that precedes this one in the block.
          property PrevStmt : TvbStmt read fPrevStmt write fPrevStmt;
      end;

    TvbNameStmt = class;

    TvbExitBlock = (
      exitDo,
      exitFor,
      exitSub,
      exitFunction,
      exitProperty
      );

    // Exit Sub
    // Exit Function
    // Exit Property
    TvbExitStmt =
      class( TvbStmt )
        public
          constructor Create;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
      end;

    // Exit Do
    // Exit For
    TvbExitLoopStmt =
      class( TvbStmt )
        public
          constructor Create;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
      end;

    // ident:
    // number:
    // number
    TvbLabelDefStmt =
      class( TvbStmt )
        public
          constructor Create( lab : TvbLabelDef );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
      end;

    TvbAssignKind = ( asSet, asLet, asLSet, asRSet );

    TvbAssignStmt =
      class( TvbStmt )
        private
          fKind  : TvbAssignKind;
          fLHS   : TvbExpr;
          fValue : TvbExpr;
        public
          constructor Create( assign : TvbAssignKind;
                              dest   : TvbExpr = nil;
                              source : TvbExpr = nil ); overload;
          constructor Create( dest : TvbExpr; source : TvbExpr ); overload;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          destructor Destroy; override;
          property AssignKind : TvbAssignKind read fKind  write fKind;
          property LHS        : TvbExpr       read fLHS   write fLHS;
          property Value      : TvbExpr       read fValue write fValue;
      end;

    TvbCircleStmt =
      class( TvbStmt )
        private
          fObject     : TvbExpr;
          fX          : TvbExpr;
          fY          : TvbExpr;
          fR          : TvbExpr;
          fColor      : TvbExpr;
          fStartAngle : TvbExpr;
          fEndAngle   : TvbExpr;
          fAspect     : TvbExpr;
          fStep       : boolean;
        public
          constructor Create( obj        : TvbExpr = nil;
                              x          : TvbExpr = nil;
                              y          : TvbExpr = nil;
                              r          : TvbExpr = nil;
                              step       : boolean = true;
                              color      : TvbExpr = nil;
                              startAngle : TvbExpr = nil;
                              endAngle   : TvbExpr = nil;
                              aspect     : TvbExpr = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property ObjectRef  : TvbExpr read fObject     write fObject;
          property Aspect     : TvbExpr read fAspect     write fAspect;
          property Color      : TvbExpr read fColor      write fColor;
          property EndAngle   : TvbExpr read fEndAngle   write fEndAngle;
          property R          : TvbExpr read fR          write fR;
          property StartAngle : TvbExpr read fStartAngle write fStartAngle;
          property Step       : boolean read fStep       write fStep;
          property X          : TvbExpr read fX          write fX;
          property Y          : TvbExpr read fY          write fY;
      end;

    TvbLineStmt =
      class( TvbStmt )
        private
          fObject : TvbExpr;
          fX1     : TvbExpr;
          fY1     : TvbExpr;
          fX2     : TvbExpr;
          fY2     : TvbExpr;
          fColor  : TvbExpr;
          fStep   : boolean;
          fFilled : boolean;
          fBoxed  : boolean;
        public
          constructor Create( obj    : TvbExpr = nil;
                              x1     : TvbExpr = nil;
                              y1     : TvbExpr = nil;
                              x2     : TvbExpr = nil;
                              y2     : TvbExpr = nil;
                              color  : TvbExpr = nil;
                              step   : boolean = false;
                              filled : boolean = true;
                              boxed  : boolean = true );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property ObjectRef : TvbExpr read fObject write fObject;
          property X1        : TvbExpr read fX1     write fX1;
          property Y1        : TvbExpr read fY1     write fY1;
          property X2        : TvbExpr read fX2     write fX2;
          property Y2        : TvbExpr read fY2     write fY2;
          property Color     : TvbExpr read fColor  write fColor;
          property Step      : boolean read fStep   write fStep;
          property Filled    : boolean read fFilled write fFilled;
          property Boxed     : boolean read fBoxed  write fBoxed;
      end;

    TvbPSetStmt =
      class( TvbStmt )
        private
          fObject : TvbExpr;
          fX      : TvbExpr;
          fY      : TvbExpr;
          fColor  : TvbExpr;
        public
          constructor Create( obj   : TvbExpr = nil;
                              x     : TvbExpr = nil;
                              y     : TvbExpr = nil;
                              color : TvbExpr = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property ObjectRef : TvbExpr read fObject write fObject;
          property X         : TvbExpr read fX      write fX;
          property Y         : TvbExpr read fY      write fY;
          property Color     : TvbExpr read fColor  write fColor;
      end;

    // Method1 arg1, arg2,...argN
    // Call Method1(arg1, arg2,...argN)
    TvbCallStmt =
      class( TvbStmt )
        private
          fMethod : TvbExpr;
        public
          constructor Create( aMethod : TvbExpr = nil ); overload;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          destructor Destroy; override;
          property Method : TvbExpr read fMethod write fMethod;
      end;

    TvbGetStmt =
      class( TvbStmt )
        private
          fFileNum : TvbExpr;
          fRecNum  : TvbExpr;
          fVar     : TvbExpr;
        public
          constructor Create( fileNum : TvbExpr = nil;
                              rec     : TvbExpr = nil;
                              varName : TvbExpr = nil );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          destructor Destroy; override;
          property FileNum : TvbExpr read fFileNum write fFileNum;
          property RecNum  : TvbExpr read fRecNum  write fRecNum;
          property VarName : TvbExpr read fVar     write fVar;
      end;

    TvbLineInputStmt =
      class( TvbStmt )
        private
          fFileNum : TvbExpr;
          fVar     : TvbExpr;
        public
          constructor Create( fileNum : TvbExpr = nil; varName : TvbExpr = nil );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          destructor Destroy; override;
          property FileNum : TvbExpr read fFileNum write fFileNum;
          property VarName : TvbExpr read fVar     write fVar;
      end;

    TvbInputStmt =
      class( TvbStmt )
        private
          fFileNum : TvbExpr;
          fVars    : TvbExprList;
          function GetVar( i : integer ) : TvbExpr;
          function GetVarCount : integer;
        public
          constructor Create( fileNum : TvbExpr = nil; aVarNames : TvbExprList = nil );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          destructor Destroy; override;
          procedure AddVar( v : TvbExpr );
          property FileNum           : TvbExpr read fFileNum write fFileNum;
          property VarCount          : integer read GetVarCount;
          property Vars[i : integer] : TvbExpr read GetVar;
      end;

    TvbPrintArgKind = ( pakSpc, pakTab, pakExpr );

    TvbPrintStmtExpr =
      class( TvbExpr )
        private
          fArgKind : TvbPrintArgKind;
        public
          constructor Create( anArgKind : TvbPrintArgKind = pakExpr; expr : TvbExpr = nil );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          destructor Destroy; override;
          property ArgKind : TvbPrintArgKind read fArgKind write fArgKind;
          property Expr    : TvbExpr         read fOp1     write fOp1;
      end;

    TvbAbstractPrintStmt =
      class( TvbStmt )
        private
          fExprs     : TvbNodeList;
          fLineBreak : boolean;
          function GetExprCount : integer;
          function GetExpr( i : integer ) : TvbPrintStmtExpr;
        public
          destructor Destroy; override;
          function AddExpr( kind : TvbPrintArgKind ) : TvbPrintStmtExpr;
          property Expr[i : integer]  : TvbPrintStmtExpr read GetExpr;
          property ExprCount          : integer          read GetExprCount;
          property LineBreak          : boolean          read fLineBreak write fLineBreak;
      end;

    TvbPrintStmt =
      class( TvbAbstractPrintStmt )
        private
          fFileNum : TvbExpr;
        public
          constructor Create;
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property FileNum : TvbExpr read fFileNum write fFileNum;
      end;

    TvbPrintMethod =
      class( TvbAbstractPrintStmt )
        private
          fObj : TvbExpr;
        public
          constructor Create;
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property Obj : TvbExpr read fObj write fObj;
      end;

    TvbCloseStmt =
      class( TvbStmt )
        private
          fFileNum : TvbExpr;
        public
          constructor Create( fileNum : TvbExpr = nil );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          destructor Destroy; override;
          property FileNum : TvbExpr read fFileNum write fFileNum;
      end;

    TvbLockStmt =
      class( TvbStmt )
        private
          fFileNum  : TvbExpr;
          fStartRec : TvbExpr;
          fEndRec   : TvbExpr;
        public
          constructor Create( fileNum  : TvbExpr = nil;
                              startRec : TvbExpr = nil;
                              endRec   : TvbExpr = nil );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          destructor Destroy; override;
          property EndRec   : TvbExpr read fEndRec   write fEndRec;
          property FileNum  : TvbExpr read fFileNum  write fFileNum;
          property StartRec : TvbExpr read fStartRec write fStartRec;
      end;

    TvbUnlockStmt =
      class(TvbStmt)
        private
          fEndRec   : TvbExpr;
          fFileNum  : TvbExpr;
          fStartRec : TvbExpr;
        public
          constructor Create( fileNum  : TvbExpr = nil;
                              startRec : TvbExpr = nil;
                              endRec   : TvbExpr = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property EndRec   : TvbExpr read fEndRec   write fEndRec;
          property FileNum  : TvbExpr read fFileNum  write fFileNum;
          property StartRec : TvbExpr read fStartRec write fStartRec;
      end;

    TvbPutStmt =
      class( TvbStmt )
        private
          fFileNum : TvbExpr;
          fRecNum  : TvbExpr;
          fVar     : TvbExpr;
        public
          constructor Create( fileNum : TvbExpr = nil;
                              rec     : TvbExpr = nil;
                              varName : TvbExpr = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property FileNum : TvbExpr read fFileNum write fFileNum;
          property RecNum  : TvbExpr read fRecNum  write fRecNum;
          property VarName : TvbExpr read fVar     write fVar;
      end;

    TvbSeekStmt =
      class( TvbStmt )
        private
          fFileNum : TvbExpr;
          fPos     : TvbExpr;
        public
          constructor Create( fileNum : TvbExpr = nil; aPos : TvbExpr = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property FileNum : TvbExpr read fFileNum write fFileNum;
          property Pos     : TvbExpr read fPos     write fPos;
      end;

    TvbWidthStmt =
      class( TvbStmt )
        private
          fFileNum  : TvbExpr;
          fRecWidth : TvbExpr;
        public
          constructor Create( fileNum : TvbExpr = nil; width : TvbExpr = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property FileNum : TvbExpr read fFileNum  write fFileNum;
          property Width   : TvbExpr read fRecWidth write fRecWidth;
      end;

    TvbWriteStmt =
      class( TvbStmt )
        private
          fFileNum : TvbExpr;
          fValues  : TvbExprList;
          function GetValueCount : integer;
          function GetValue( i : integer ) : TvbExpr;
        public
          constructor Create( fileNum : TvbExpr = nil; values : TvbExprList = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          procedure AddValue( val : TvbExpr );
          property FileNum            : TvbExpr read fFileNum write fFileNum;
          property ValueCount         : integer read GetValueCount;
          property Value[i : integer] : TvbExpr read GetValue;
      end;

    TvbErrorStmt =
      class( TvbStmt )
        private
          fErrorNum : TvbExpr;
        public
          constructor Create( err : TvbExpr = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property ErrorNumber : TvbExpr read fErrorNum write fErrorNum;
      end;

    TvbDateStmt =
      class( TvbStmt )
        private
          fNewDate : TvbExpr;
        public
          constructor Create( newDate : TvbExpr = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property NewDate : TvbExpr read fNewDate write fNewDate;
      end;

    TvbTimeStmt =
      class( TvbStmt )
        private
          fNewTime : TvbExpr;
        public
          constructor Create( newTime : TvbExpr );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property NewTime : TvbExpr read fNewTime write fNewTime;
      end;

    TvbNameStmt =
      class( TvbStmt )
        private
          fOldPath : TvbExpr;
          fNewPath : TvbExpr;
        public
          constructor Create( oldPath : TvbExpr = nil; newPath : TvbExpr = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property NewPath : TvbExpr read fNewPath write fNewPath;
          property OldPath : TvbExpr read fOldPath write fOldPath;
      end;

    TvbStopStmt =
      class( TvbStmt )
        public
          constructor Create;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
      end;

    TvbDebugAssertStmt =
      class( TvbStmt )
        private
          fBoolExpr : TvbExpr;
        public
          constructor Create( cond : TvbExpr = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property BoolExpr : TvbExpr read fBoolExpr write fBoolExpr;
      end;

    TvbDebugPrintStmt =
      class( TvbAbstractPrintStmt )
        public
          constructor Create;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
      end;

    // A clause for a CASE statement in the form: Expression
    TvbCaseClause =
      class( TvbNode )
        private
          fExpr : TvbExpr;
        public
          constructor Create( nk : TvbNodeKind; expr : TvbExpr );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property Expr : TvbExpr read fExpr;
      end;

    // A clause in the form: [Is] Operator Expression
    TvbRelationalCaseClause =
      class( TvbCaseClause )
        private
          fOper : TvbOperator;
        public
          constructor Create( anOper : TvbOperator; expr : TvbExpr );
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property Operator : TvbOperator read fOper;
      end;

    // A clause in the form: Expression [To Expression]
    TvbRangeCaseClause =
      class( TvbCaseClause )
        private
          fUpperExpr : TvbExpr;
        public
          constructor Create( aLowerExpr, aUpperExpr : TvbExpr );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property UpperExpr : TvbExpr read fUpperExpr write fUpperExpr;
      end;

    // A "Case" within a "Select Case" statement
    TvbCaseStmt =
      class( TvbStmt )
        private
          fClauses : TvbNodeList;
          fStmts   : TvbStmtBlock;
          function GetClause( i : integer ) : TvbCaseClause;
          function GetClauseCount : integer;
        public
          constructor Create;
          destructor Destroy; override;
          function AddCaseClause( caseClause : TvbCaseClause ) : TvbCaseClause;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property ClauseCount         : integer       read GetClauseCount;
          property Clause[i : integer] : TvbCaseClause read GetClause;
          property StmtBlock           : TvbStmtBlock  read fStmts   write fStmts;
      end;

    // Select Case
    TvbSelectCaseStmt =
      class( TvbStmt )
        private
          fExpr     : TvbExpr;
          fCaseElse : TvbStmtBlock;
          fCases    : TvbNodeList;
          function GetCase( i : integer ) : TvbCaseStmt;
          function GetCaseCount : integer;
        public
          constructor Create( expr : TvbExpr = nil );
          destructor Destroy; override;
          function AddCaseStmt( caseStmt : TvbCaseStmt ) : TvbCaseStmt;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property CaseCount          : integer      read GetCaseCount;
          property CaseElse           : TvbStmtBlock read fCaseElse write fCaseElse;
          property Cases[i : integer] : TvbCaseStmt  read GetCase;
          property Expr               : TvbExpr      read fExpr     write fExpr;
      end;

    // If..Then..ElseIf..Else
    //
    TvbIfStmt =
      class( TvbStmt )
        private
          fExpr    : TvbExpr;
          fThen    : TvbStmtBlock;
          fElse    : TvbStmtBlock;
          fElseIfs : TvbNodeList;
          function GetElseIf(i: integer): TvbIfStmt;
          function GetElseIfCount: integer;
        public
          constructor Create( expr      : TvbExpr      = nil;
                              thenStmts : TvbStmtBlock = nil;
                              elseStmts : TvbStmtBlock = nil;
                              elseIfs   : TvbNodeList  = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          function AddElseIf : TvbIfStmt;
          property ElseBlock           : TvbStmtBlock read fElse write fElse;
          property ElseIfCount         : integer      read GetElseIfCount;
          property ElseIf[i : integer] : TvbIfStmt    read GetElseIf;
          property Expr                : TvbExpr      read fExpr write fExpr;
          property ThenBlock           : TvbStmtBlock read fThen write fThen;
      end;

    // dlWhile    - Loop while condition is true.
    // dlUntil    - Loop until condition is true.
    // dlInfinite - Loop indefinitely.
    //
    TvbDoLoop = ( doWhile, doUntil, doInfinite );

    // Do..Loop
    // While..Wend
    //
    TvbDoLoopStmt =
      class( TvbStmt )
        private
          fLoop        : TvbDoLoop;
          fTestAtBegin : boolean;
          fCond        : TvbExpr;
          fBlock       : TvbStmtBlock;
        public
          constructor Create( aLoop       : TvbDoLoop    = doInfinite;
                              TestAtBegin : boolean      = true;
                              aCond       : TvbExpr      = nil;
                              Stmts       : TvbStmtBlock = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property Condition   : TvbExpr      read fCond        write fCond;
          property DoLoop      : TvbDoLoop    read fLoop        write fLoop;
          property StmtBlock   : TvbStmtBlock read fBlock       write fBlock;
          property TestAtBegin : boolean      read fTestAtBegin write fTestAtBegin;
      end;

    TvbForDir = ( fdTo, fdDownto );

    // For
    //
    TvbForStmt =
      class( TvbStmt )
        private
          fCounter  : TvbExpr;
          fInitVal  : TvbExpr;
          fFinalVal : TvbExpr;
          fStepVal  : TvbExpr;
          fBlock    : TvbStmtBlock;
        public
          constructor Create( counter : TvbExpr      = nil;
                              init    : TvbExpr      = nil;
                              max     : TvbExpr      = nil;
                              step    : TvbExpr      = nil;
                              stmts   : TvbStmtBlock = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property Counter    : TvbExpr      read fCounter  write fCounter;
          property FinalValue : TvbExpr      read fFinalVal write fFinalVal;
          property InitValue  : TvbExpr      read fInitVal  write fInitVal;
          property StepValue  : TvbExpr      read fStepVal  write fStepVal;
          property StmtBlock  : TvbStmtBlock read fBlock    write fBlock;
      end;

    TvbForeachStmt =
      class( TvbStmt )
        private
          fElement : TvbExpr;
          fGroup   : TvbExpr;
          fBlock   : TvbStmtBlock;
        public
          constructor Create( elem  : TvbExpr      = nil;
                              group : TvbExpr      = nil;
                              stmts : TvbStmtBlock = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property Element   : TvbExpr      read fElement write fElement;
          property Group     : TvbExpr      read fGroup   write fGroup;
          property StmtBlock : TvbStmtBlock read fBlock   write fBlock;
      end;

    TvbWithStmt =
      class( TvbStmt )
        private
          fObjOrRec : TvbExpr;
          fBlock    : TvbStmtBlock;
        public
          constructor Create( expr : TvbExpr = nil; stmts : TvbStmtBlock = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property ObjOrRec  : TvbExpr      read fObjOrRec write fObjOrRec;
          property StmtBlock : TvbStmtBlock read fBlock    write fBlock;
      end;

    TvbMidAssignStmt =
      class( TvbStmt )
        private
          fVarName  : TvbExpr;
          fStartPos : TvbExpr;
          fRep      : TvbExpr;
          fLen      : TvbExpr;
        public
          constructor Create( varName   : TvbExpr = nil;
                              aStartPos : TvbExpr = nil;
                              aLen      : TvbExpr = nil;
                              aRep      : TvbExpr = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property Length      : TvbExpr read fLen      write fLen;
          property Replacement : TvbExpr read fRep      write fRep;
          property StartPos    : TvbExpr read fStartPos write fStartPos;
          property VarName     : TvbExpr read fVarName  write fVarName;
      end;

    TvbReturnStmt =
      class( TvbStmt )
        public
          constructor Create;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
      end;

    // RaiseEvent
    TvbRaiseEventStmt =
      class( TvbStmt )
        private
          fEventName : string;
          fArgs      : TvbArgList;
        public
          constructor Create( const event : string = ''; args : TvbArgList = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property EventArgs : TvbArgList read fArgs      write fArgs;
          property EventName : string     read fEventName write fEventName;
      end;

    TvbGotoOrGosub = ( ggGoTo, ggGoSub );

    // Goto
    // GoSub
    TvbGotoOrGosubStmt =
      class( TvbStmt )
        private
          fKind  : TvbGotoOrGosub;
          fLabel : TvbSimpleName;
        public
          constructor Create( kind : TvbGotoOrGosub = ggGoTo; labelName : TvbSimpleName = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property GotoOrGosub : TvbGotoOrGosub read fKind  write fKind;
          property LabelName   : TvbSimpleName  read fLabel write fLabel;
      end;

    TvbResumeKind = ( rkNext, rkLabel, rkZero );

    // Resume Next
    // Resume line
    // Resume [0]
    TvbResumeStmt =
      class( TvbStmt )
        private
          fKind  : TvbResumeKind;
          fLabel : TvbSimpleName;
        public
          constructor Create( kind : TvbResumeKind = rkZero; labelName : TvbSimpleName = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property ResumeKind : TvbResumeKind read fKind  write fKind;
          property LabelName  : TvbSimpleName read fLabel write fLabel;
      end;

    TvbEndStmt =
      class( TvbStmt )
        public
          constructor Create;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
      end;

    TvbOnErrorDo = ( doResumeNext, doGoto0, doGotoLabel );

    // On Error..Resume
    TvbOnErrorStmt =
      class( TvbStmt )
        private
          fLabel   : TvbSimpleName;
          fOnError : TvbOnErrorDo;
        public
          constructor Create( onErrKind : TvbOnErrorDo = doGoto0; labelName : TvbSimpleName = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property OnErrorDo : TvbOnErrorDo  read fOnError write fOnError;
          property LabelName : TvbSimpleName read fLabel   write fLabel;
      end;

    // On..GoTo
    // On..GoSub
    TvbOnGotoOrGosubStmt =
      class( TvbStmt )
        private
          fExpr        : TvbExpr;
          fGotoOrGosub : TvbGotoOrGosub;
          fLabels      : TvbNodeList;
          function GetLabelCount : integer;
          function GetLabelName( i : integer ) : TvbSimpleName;
        public
          constructor Create( expr : TvbExpr = nil; gg : TvbGotoOrGosub = ggGoSub );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          procedure AddLabel( lbl : TvbSimpleName );
          property Expr                   : TvbExpr        read fExpr        write fExpr;
          property GotoOrGosub            : TvbGotoOrGosub read fGotoOrGosub write fGotoOrGosub;
          property LabelCount             : integer        read GetLabelCount;
          property LabelName[i : integer] : TvbSimpleName  read GetLabelName;
      end;

    // ReDim
    TvbReDimStmt =
      class( TvbStmt )
        private
          fArrVar   : TvbVarDef;
          fPreserve : boolean;
        public
          constructor Create( preserve : boolean = false; arrVar : TvbVarDef = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property ArrayVar : TvbVarDef read fArrVar   write fArrVar;
          property Preserve : boolean   read fPreserve write fPreserve;
      end;

    // Erase
    TvbEraseStmt =
      class( TvbStmt )
        private
          fArrVar : TvbExpr;
        public
          constructor Create( arrVar : TvbExpr = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property ArrayVar : TvbExpr read fArrVar write fArrVar;
      end;

    TvbOpenMode   = ( omAppend, omBinary, omInput, omOutput, omRandom );
    TvbOpenAccess = ( oaRead, oaWrite, oaReadWrite );
    TvbOpenLock   = ( olShared, olLockRead, olLockWrite, olLockReadWrite );

    // Open
    TvbOpenStmt =
      class( TvbStmt )
        private
          fFilePath : TvbExpr;
          fMode     : TvbOpenMode;
          fAccess   : TvbOpenAccess;
          fLock     : TvbOpenLock;
          fFileNum  : TvbExpr;
          fRecLen   : TvbExpr;
        public
          constructor Create( filePath : TvbExpr       = nil;
                              mode     : TvbOpenMode   = omRandom;
                              access   : TvbOpenAccess = oaReadWrite;
                              lock     : TvbOpenLock   = olLockWrite;
                              fileNum  : TvbExpr       = nil;
                              recLen   : TvbExpr       = nil );
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          property Access   : TvbOpenAccess read fAccess   write fAccess;
          property FileNum  : TvbExpr       read fFileNum  write fFileNum;
          property FilePath : TvbExpr       read fFilePath write fFilePath;
          property Lock     : TvbOpenLock   read fLock     write fLock;
          property Mode     : TvbOpenMode   read fMode     write fMode;
          property RecLen   : TvbExpr       read fRecLen   write fRecLen;
      end;

    TvbStmtBlock =
      class( TvbStmt )
        private
          fFirstStmt : TvbStmt;
          fLastStmt  : TvbStmt;
          function GetCount: integer;
        public
          constructor Create;
          destructor Destroy; override;
          procedure Accept( Visitor : TvbNodeVisitor; Context : TObject = nil ); override;
          function Add( stmt : TvbStmt ) : TvbStmt;

          function AddCallStmt : TvbCallStmt;
          function AddCircleStmt : TvbCircleStmt;
          function AddCloseStmt : TvbCloseStmt;
          function AddDateStmt : TvbDateStmt;
          function AddDebugAssertStmt : TvbDebugAssertStmt;
          function AddDebugPrintStmt : TvbDebugPrintStmt;
          function AddDoLoopStmt : TvbDoLoopStmt;
          function AddEraseStmt : TvbEraseStmt;
          function AddErrorStmt : TvbErrorStmt;
          function AddForeachStmt : TvbForeachStmt;
          function AddForStmt : TvbForStmt;
          function AddGetStmt : TvbGetStmt;
          function AddGotoOrGosubStmt : TvbGotoOrGosubStmt;
          function AddIfStmt : TvbIfStmt;
          function AddInputStmt : TvbInputStmt;
          function AddLineInputStmt : TvbLineInputStmt;
          function AddLineStmt : TvbLineStmt;
          function AddLockStmt : TvbLockStmt;
          function AddMidAssignStmt : TvbMidAssignStmt;
          function AddNameStmt : TvbNameStmt;
          function AddOnErrorStmt : TvbOnErrorStmt;
          function AddOnGotoOrGosubStmt : TvbOnGotoOrGosubStmt;
          function AddOpenStmt : TvbOpenStmt;
          function AddPrintStmt : TvbPrintStmt;
          function AddPrintMethod : TvbPrintMethod;
          function AddPSetStmt : TvbPSetStmt;
          function AddPutStmt : TvbPutStmt;
          function AddRaiseEventStmt : TvbRaiseEventStmt;
          function AddReDimStmt : TvbReDimStmt;
          function AddResumeStmt : TvbResumeStmt;
          function AddSeekStmt : TvbSeekStmt;
          function AddSelectCaseStmt : TvbSelectCaseStmt;
          function AddStopStmt : TvbStopStmt;
          function AddUnlockStmt : TvbUnlockStmt;
          function AddWhileStmt : TvbDoLoopStmt;
          function AddWidthStmt : TvbWidthStmt;
          function AddWithStmt : TvbWithStmt;
          function AddWriteStmt : TvbWriteStmt;
          // First statement within this block.
          property FirstStmt : TvbStmt read fFirstStmt write fFirstStmt;
          // Last statement within this block.
          property LastStmt : TvbStmt read fLastStmt write fLastStmt;
          // The number of statements in this block.
          property Count : integer read GetCount;
      end;

    TvbNodeVisitor =
      class
        public
          procedure OnAbsExpr( Node : TvbAbsExpr; Context : TObject = nil ); virtual;
          procedure OnAddressOfExpr( Node : TvbAddressOfExpr; Context : TObject = nil ); virtual; 
          procedure OnArrayExpr( Node : TvbArrayExpr; Context : TObject = nil ); virtual; 
          procedure OnAssignStmt( Node : TvbAssignStmt; Context : TObject = nil ); virtual;
          procedure OnBinaryExpr( Node : TvbBinaryExpr; Context : TObject = nil ); virtual; 
          procedure OnBoolLitExpr( Node : TvbBoolLit; Context : TObject = nil ); virtual; 
          procedure OnCallOrIndexerExpr( Node : TvbCallOrIndexerExpr; Context : TObject = nil ); virtual; 
          procedure OnCallStmt( Node : TvbCallStmt; Context : TObject = nil ); virtual; 
          procedure OnCastExpr( Node : TvbCastExpr; Context : TObject = nil ); virtual; 
          procedure OnCircleStmt( Node : TvbCircleStmt; Context : TObject = nil ); virtual; 
          procedure OnClassDef( Node : TvbClassDef; Context : TObject = nil ); virtual; 
          procedure OnClassModuleDef( Node : TvbClassModuleDef; Context : TObject = nil ); virtual; 
          procedure OnCloseStmt( Node : TvbCloseStmt; Context : TObject = nil ); virtual; 
          procedure OnConstDef( Node : TvbConstDef; Context : TObject = nil ); virtual; 
          procedure OnDateExpr( Node : TvbDateExpr; Context : TObject = nil ); virtual; 
          procedure OnDateLitExpr( Node : TvbDateLit; Context : TObject = nil ); virtual;  
          procedure OnDateStmt( Node : TvbDateStmt; Context : TObject = nil ); virtual; 
          procedure OnDebugAssertStmt( Node : TvbDebugAssertStmt; Context : TObject = nil ); virtual; 
          procedure OnDebugPrintStmt( Node : TvbDebugPrintStmt; Context : TObject = nil ); virtual; 
          procedure OnDictAccessExpr( Node : TvbDictAccessExpr; Context : TObject = nil ); virtual; 
          procedure OnDllFuncDef( Node : TvbDllFuncDef; Context : TObject = nil ); virtual; 
          procedure OnDoEventsExpr( Node : TvbDoEventsExpr; Context : TObject = nil ); virtual; 
          procedure OnDoLoopStmt( Node : TvbDoLoopStmt; Context : TObject = nil ); virtual; 
          procedure OnEndStmt( Node : TvbEndStmt; Context : TObject = nil ); virtual; 
          procedure OnEnumDef( Node : TvbEnumDef; Context : TObject = nil ); virtual; 
          procedure OnEraseStmt( Node : TvbEraseStmt; Context : TObject = nil ); virtual; 
          procedure OnErrorStmt( Node : TvbErrorStmt; Context : TObject = nil ); virtual; 
          procedure OnEventDef( Node : TvbEventDef; Context : TObject = nil ); virtual; 
          procedure OnExitStmt( Node : TvbExitStmt; Context : TObject = nil ); virtual;
          procedure OnExitLoopStmt( Node : TvbExitLoopStmt; Context : TObject = nil ); virtual;
          procedure OnFixExpr( Node : TvbFixExpr; Context : TObject = nil ); virtual; 
          procedure OnFloatLitExpr( Node : TvbFloatLit; Context : TObject = nil ); virtual; 
          procedure OnForeachStmt( Node : TvbForeachStmt; Context : TObject = nil ); virtual; 
          procedure OnForStmt( Node : TvbForStmt; Context : TObject = nil ); virtual;
          procedure OnFormDef( Node : TvbFormDef; Context : TObject = nil ); virtual;
          procedure OnFormModuleDef( Node : TvbFormModuleDef; Context : TObject = nil ); virtual;
          procedure OnFunctionDef( Node : TvbFuncDef; Context : TObject = nil ); virtual;
          procedure OnFuncResultExpr( Node : TvbFuncResultExpr; Context : TObject = nil ); virtual; 
          procedure OnGetStmt( Node : TvbGetStmt; Context : TObject = nil ); virtual; 
          procedure OnGotoOrGosubStmt( Node : TvbGotoOrGosubStmt; Context : TObject = nil ); virtual; 
          procedure OnIfStmt( Node : TvbIfStmt; Context : TObject = nil ); virtual;
          procedure OnInputExpr( Node : TvbInputExpr; Context : TObject = nil ); virtual;
          procedure OnInputStmt( Node : TvbInputStmt; Context : TObject = nil ); virtual; 
          procedure OnLabelDef( Node : TvbLabelDef; Context : TObject = nil ); virtual; 
          procedure OnLabelDefStmt( Node : TvbLabelDefStmt; Context : TObject = nil ); virtual; 
          //procedure OnLabelRef( Node : tvblabel; Context : TObject = nil ); virtual; 
          procedure OnLBoundExpr( Node : TvbLBoundExpr; Context : TObject = nil ); virtual; 
          procedure OnLenBExpr( Node : TvbLenBExpr; Context : TObject = nil ); virtual; 
          procedure OnLenExpr( Node : TvbLenExpr; Context : TObject = nil ); virtual; 
          procedure OnLibModuleDef( Node : TvbLibModuleDef; Context : TObject = nil ); virtual; 
          procedure OnLineStmt( Node : TvbLineStmt; Context : TObject = nil ); virtual; 
          procedure OnLineInputStmt( Node : TvbLineInputStmt; Context : TObject = nil ); virtual; 
          procedure OnLockStmt( Node : TvbLockStmt; Context : TObject = nil ); virtual; 
          procedure OnMeExpr( Node : TvbMeExpr; Context : TObject = nil ); virtual; 
          procedure OnMemberAccessExpr( Node : TvbMemberAccessExpr; Context : TObject = nil ); virtual; 
          procedure OnMidAssignStmt( Node : TvbMidAssignStmt; Context : TObject = nil ); virtual; 
          procedure OnMidExpr( Node : TvbMidExpr; Context : TObject = nil ); virtual; 
          procedure OnNameExpr( Node : TvbNameExpr; Context : TObject = nil ); virtual;
          procedure OnNamedArgExpr( Node : TvbNamedArgExpr; Context : TObject = nil ); virtual; 
          procedure OnNameStmt( Node : TvbNameStmt; Context : TObject = nil ); virtual; 
          procedure OnNewExpr( Node : TvbNewExpr; Context : TObject = nil ); virtual; 
          procedure OnNothingExpr( Node : TvbNothingLit; Context : TObject = nil ); virtual; 
          procedure OnOnErrorStmt( Node : TvbOnErrorStmt; Context : TObject = nil ); virtual; 
          procedure OnOnGotoOrGosubStmt( Node : TvbOnGotoOrGosubStmt; Context : TObject = nil ); virtual;  
          procedure OnOpenStmt( Node : TvbOpenStmt; Context : TObject = nil ); virtual; 
          procedure OnParamDef( Node : TvbParamDef; Context : TObject = nil ); virtual; 
          procedure OnPrintArgExpr( Node : TvbPrintStmtExpr; Context : TObject = nil ); virtual; 
          procedure OnPrintStmt( Node : TvbPrintStmt; Context : TObject = nil ); virtual; 
          procedure OnProjectDef( Node : TvbProject; Context : TObject = nil ); virtual; 
          procedure OnPropertyDef( Node : TvbPropertyDef; Context : TObject = nil ); virtual; 
          procedure OnPSetStmt( Node : TvbPSetStmt; Context : TObject = nil ); virtual; 
          procedure OnPutStmt( Node : TvbPutStmt; Context : TObject = nil ); virtual;
          procedure OnQualifiedName( Node : TvbQualifiedName; Context : TObject = nil ); virtual; 
          procedure OnRaiseEventStmt( Node : TvbRaiseEventStmt; Context : TObject = nil ); virtual; 
          procedure OnRecordDef( Node : TvbRecordDef; Context : TObject = nil ); virtual; 
          procedure OnReDimStmt( Node : TvbReDimStmt; Context : TObject = nil ); virtual; 
          procedure OnResumeStmt( Node : TvbResumeStmt; Context : TObject = nil ); virtual;
          procedure OnReturnStmt( Node : TvbReturnStmt; Context : TObject = nil ); virtual;
          procedure OnSelectCaseStmt( Node : TvbSelectCaseStmt; Context : TObject = nil ); virtual;
          procedure OnCaseStmt( Node : TvbCaseStmt; Context : TObject = nil ); virtual;
          procedure OnRelationalCaseClause( Node : TvbRelationalCaseClause; Context : TObject = nil ); virtual;
          procedure OnRangeCaseClause( Node : TvbRangeCaseClause; Context : TObject = nil ); virtual;
          procedure OnCaseClause( Node : TvbCaseClause; Context : TObject = nil ); virtual;
          procedure OnSeekExpr( Node : TvbSeekExpr; Context : TObject = nil ); virtual; 
          procedure OnSeekStmt( Node : TvbSeekStmt; Context : TObject = nil ); virtual;
          procedure OnSgnExpr( Node : TvbSgnExpr; Context : TObject = nil ); virtual;
          procedure OnSimpleName( Node : TvbSimpleName; Context : TObject = nil ); virtual;
          procedure OnStdModuleDef( Node : TvbStdModuleDef; Context : TObject = nil ); virtual;
          procedure OnStmtBlock( Node : TvbStmtBlock; Context : TObject = nil ); virtual;
          procedure OnStopStmt( Node : TvbStopStmt; Context : TObject = nil ); virtual;
          procedure OnStringExpr( Node : TvbStringExpr; Context : TObject = nil ); virtual;
          procedure OnStringLitExpr( Node : TvbStringLit; Context : TObject = nil ); virtual;
          procedure OnTimeStmt( Node : TvbTimeStmt; Context : TObject = nil ); virtual;
          procedure OnIntegerLitExpr( Node : TvbIntLit; Context : TObject = nil ); virtual;
          procedure OnIntExpr( Node : TvbIntExpr; Context : TObject = nil ); virtual; 
          procedure OnTypeOfExpr( Node : TvbTypeOfExpr; Context : TObject = nil ); virtual; 
          procedure OnTypeRef( Node : TvbType; Context : TObject = nil ); virtual;
          procedure OnUboundExpr( Node : TvbUBoundExpr; Context : TObject = nil ); virtual; 
          procedure OnUnaryExpr( Node : TvbUnaryExpr; Context : TObject = nil ); virtual; 
          procedure OnUnlockStmt( Node : TvbUnlockStmt; Context : TObject = nil ); virtual; 
          procedure OnVarDef( Node : TvbVarDef; Context : TObject = nil ); virtual; 
          procedure OnWidthStmt( Node : TvbWidthStmt; Context : TObject = nil ); virtual; 
          procedure OnWithStmt( Node : TvbWithStmt; Context : TObject = nil ); virtual; 
          procedure OnWriteStmt( Node : TvbWriteStmt; Context : TObject = nil ); virtual; 
      end;

  var
    vbAnyType          : IvbType;
    vbBooleanType      : IvbType;
    vbByteType         : IvbType;
    vbCurrencyType     : IvbType;
    vbDateType         : IvbType;
    vbDoubleType       : IvbType;
    vbIntegerType      : IvbType;
    vbLongType         : IvbType;
    vbObjectType       : IvbType;
    vbSingleType       : IvbType;
    vbStringType       : IvbType;
    vbVariantType      : IvbType;
    vbVariantArrayType : IvbType;
    _vbUnknownType     : IvbType; 

    // Dynamic arrays of simple types.
    vbBooleanDynArrayType  : IvbType;
    vbByteDynArrayType     : IvbType;
    vbCurrencyDynArrayType : IvbType;
    vbDateDynArrayType     : IvbType;
    vbDoubleDynArrayType   : IvbType;
    vbIntegerDynArrayType  : IvbType;
    vbLongDynArrayType     : IvbType;
    vbObjectDynArrayType   : IvbType;
    vbSingleDynArrayType   : IvbType;
    vbStringDynArrayType   : IvbType;
    vbVariantDynArrayType  : IvbType;

  function IsArrayType( const T : IvbType ) : boolean;
  function IsFloatingType( const T : IvbType ) : boolean;
  function IsIntegerEqTo( E : TvbExpr; value : integer ) : boolean;
  function IsIntegralType( const T : IvbType ) : boolean;
  function IsNameEqTo( E : TvbExpr; const name : string ) : boolean;
  function IsStringEqTo( E : TvbExpr; const value : string; caseSensitive : boolean = false ) : boolean;
  function IsStringType( const T : IvbType ) : boolean;
  function IsStaticArrayType( const T : IvbType ) : boolean;

implementation

  uses    
    SysUtils,
    Variants;

  function IsIntegerEqTo( E : TvbExpr; value : integer ) : boolean;
    begin
      result := ( E.NodeKind = INT_LIT_EXPR ) and ( TvbIntLit( E ).Value = value );
    end;

  function IsStringEqTo( E : TvbExpr; const value : string; caseSensitive : boolean ) : boolean;
    begin
      result := false;
      if E.NodeKind <> STRING_LIT_EXPR
        then exit;
      if caseSensitive
        then result := TvbStringLit( E ).Value = value
        else result := SameText( TvbStringLit( E ).Value, value );
    end;

  function IsNameEqTo( E : TvbExpr; const name : string ) : boolean;
    begin
      result := false;
      if E.NodeKind <> NAME_EXPR
        then exit;
      result := SameText( TvbNameExpr( E ).Name, name );
    end;

  function IsIntegralType( const T : IvbType ) : boolean;
    begin
      result :=
        ( T <> nil ) and
        ( ( T = vbByteType ) or
          ( T = vbBooleanType ) or
          ( T = vbIntegerType ) or
          ( T = vbLongType ) or
          ( ( T.TypeDef <> nil ) and ( T.TypeDef.NodeKind = ENUM_DEF ) ) );
    end;

  function IsFloatingType( const T : IvbType ) : boolean;
    begin
      result :=
        ( T <> nil ) and
        ( ( T = vbDateType ) or
          ( T = vbSingleType ) or
          ( T = vbCurrencyType ) or
          ( T = vbDoubleType ) );
    end;

  function IsArrayType( const T : IvbType ) : boolean;
    begin
      result := ( T <> nil ) and ( T.Code and ARRAY_TYPE <> 0 );
    end;

  function IsStringType( const T : IvbType ) : boolean;
    begin
      result := ( T <> nil ) and ( T.Code and STRING_TYPE <> 0 );
    end;

  function IsStaticArrayType( const T : IvbType ) : boolean;
    begin
      result :=
        ( T <> nil ) and
        ( T.Code and ARRAY_TYPE <> 0 ) and
        ( T.ArrayBoundCount <> 0 );
    end;

  { TvbNode }

  procedure TvbNode.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
    end;

  constructor TvbNode.Create( kind : TvbNodeKind );
    begin
      fKind := kind;
    end;

  { TvbType }

  constructor TvbType.Create( typeCode : cardinal );
    begin
      fCode := typeCode;
    end;

  constructor TvbType.CreateUDT( const typeName : string );
    begin
      fCode  := 0;
      fName  := typeName;
      fName1 := typeName;
    end;

  constructor TvbType.CreateString( len : TvbExpr );
    begin
      fCode   := STRING_TYPE;
      fStrLen := len;
    end;

  destructor TvbType.Destroy;
    begin
      FreeAndNil( fStrLen );
      FreeAndNil( fBounds );
      inherited;
    end;

  function TvbType.GetArrayBoundCount : integer;
    begin
      if fBounds <> nil
        then result := fBounds.Count
        else result := 0;
    end;

  function TvbType.GetArrayBounds( i : integer ) : TvbArrayBounds;
    begin
      if fBounds = nil
        then result := nil
        else result := TvbArrayBounds( fBounds.Nodes[i] );
    end;

  procedure TvbType.AddBound( lower, upper : TvbExpr );
    begin
      if fBounds = nil
        then fBounds := TvbArrayBoundList.Create;
      fBounds.Add( TvbArrayBounds.Create( lower, upper ) );
    end;

  function TvbType.GetStringLen : TvbExpr;
    begin
      result := fStrLen;
    end;

  function TvbType.GetTypeDef : TvbTypeDef;
    begin
      result := fDef;
    end;

  function TvbType.GetCode : cardinal;
    begin
      result := fCode;
    end;

  function TvbType.GetName : string;
    begin
      result := fName;
    end;

  constructor TvbType.CreateUDT( typeDef : TvbTypeDef );
    begin
      CreateUDT( typeDef.Name );
      SetTypeDef( typeDef );
    end;

  procedure TvbType.SetTypeDef( td : TvbTypeDef );
    begin
      fDef := td;
      fName1 := td.Name1;
    end;

  function TvbType.GetParentModule : TvbModuleDef;
    begin
      result := fModule;
    end;

  constructor TvbType.CreateArray( baseType : cardinal );
    begin
      fCode := baseType or ARRAY_TYPE;
    end;

  constructor TvbType.CreateArray( const typeName : string );
    begin
      fCode := ARRAY_TYPE;
      fName := typeName;
      fName1 := typeName;
    end;

  constructor TvbType.CreateArray( typeDef : TvbTypeDef );
    begin
      fCode := ARRAY_TYPE;
      SetTypeDef( typeDef );
    end;

  constructor TvbType.CreateVector( const baseType : IvbType; bounds : TvbArrayBoundList );
    begin
      assert( bounds <> nil );
      fCode   := baseType.Code or ARRAY_TYPE;
      fBounds := bounds;
      fDef    := baseType.TypeDef;
      fName   := baseType.Name;
      fName1  := baseType.Name1;
    end;

  function TvbType.GetName1 : string;
    begin
      result := fName1;
    end;

  procedure TvbType.SetName1( const value : string );
    begin
      fName1 := value;
    end;

  { TvbArrayBounds }

  procedure TvbArrayBounds.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      //Visitor.onar
    end;

  constructor TvbArrayBounds.Create( l, u : TvbExpr );
    begin
      fLower := l;
      fUpper := u;
    end;

  destructor TvbArrayBounds.Destroy;
    begin
      FreeAndNil( fLower );
      FreeAndNil( fUpper );
      inherited;
    end;

  { TvbIntLit }

  procedure TvbIntLit.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnIntegerLitExpr( self, Context );
    end;

  constructor TvbIntLit.Create( val : integer );
    begin
      inherited Create( INT_LIT_EXPR );
      fValue := val;
    end;

  { TvbFloatLit }

  procedure TvbFloatLit.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnFloatLitExpr( self, Context );
    end;

  constructor TvbFloatLit.Create( val : double );
    begin
      inherited Create( FLOAT_LIT_EXPR );
      fValue := val;
    end;

  { TvbBoolLit }

  procedure TvbBoolLit.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnBoolLitExpr( self, Context );
    end;

  constructor TvbBoolLit.Create( val : boolean );
    begin
      inherited Create( BOOL_LIT_EXPR );
      fValue    := val;
      fNodeType := vbBooleanType;
    end;

  { TvbStringLit }

  procedure TvbStringLit.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnStringLitExpr( self, Context );
    end;

  constructor TvbStringLit.Create( const val : string );
    begin
      inherited Create( STRING_LIT_EXPR );
      fValue := val;
      fNodeType := vbStringType;
    end;

  { TvbParamDef }

  procedure TvbParamDef.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnParamDef( self, Context );
    end;

  constructor TvbParamDef.Create( flags      : TvbNodeFlags;
                                  const name : string;
                                  const typ  : IvbType;
                                  defVal     : TvbExpr );
    begin
      inherited Create( PARAM_DEF, [], name );
      fNodeType := typ;
      fDefVal   := defVal;
      NodeFlags := NodeFlags + flags;
    end;

  destructor TvbParamDef.Destroy;
    begin
      FreeAndNil( fDefVal );
      fNodeType := nil;
      inherited;
    end;

  { TvbDef }

  constructor TvbDef.Create( kind : TvbNodeKind; const flags : TvbNodeFlags; const name, name1 : string );
    begin
      inherited Create( kind );
      fFlags := flags;
      fName  := name;
      if name1 <> ''
        then fName1 := name1
        else fName1 := name;
    end;

  procedure TvbDef.SetName( const name : string );
    begin
      fName  := name;
      fName1 := name;
    end;

  { TvbNothingLit }

  procedure TvbNothingLit.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
    end;

  constructor TvbNothingLit.Create;
    begin
      inherited Create( NOTHING_EXPR );
    end;

  { TvbCallOrIndexerExpr }

  procedure TvbCallOrIndexerExpr.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnCallOrIndexerExpr( self, Context );
    end;

  procedure TvbCallOrIndexerExpr.AddArg( arg : TvbExpr );
    begin
      if fArgs = nil
        then fArgs := TvbArgList.Create;
      fArgs.Add( arg );
    end;

  constructor TvbCallOrIndexerExpr.Create( funcOrArray : TvbExpr; args : TvbArgList; isFunc : boolean );
    begin
      inherited Create( CALL_OR_INDEXER_EXPR );
      assert( funcOrArray <> nil );
      fOp1 := funcOrArray;
      fArgs := args;
      IsUsedWithEmptyParens := isFunc;
    end;

  destructor TvbCallOrIndexerExpr.Destroy;
    begin
      FreeAndNil( fArgs );
      inherited;
    end;

  function TvbCallOrIndexerExpr.GetIsRValue : boolean;
    begin
      assert( fOp1 <> nil );
      result := assign_LHS in fOp1.fFlags;
    end;

  function TvbCallOrIndexerExpr.GetIsUsedWithArgs : boolean;
    begin
      assert( fOp1 <> nil );
      result := expr_UsedWithArgs in fOp1.fFlags;
    end;

  function TvbCallOrIndexerExpr.GetIsUsedWithEmptyParens : boolean;
    begin
      assert( fOp1 <> nil );
      result := expr_MustBeFunc in fOp1.fFlags;
    end;

  procedure TvbCallOrIndexerExpr.SetIsRValue( const value : boolean );
    begin
      assert( fOp1 <> nil );
      Include( fOp1.fFlags, assign_LHS );
      Include( fFlags, expr_isNotMethod );
    end;

  procedure TvbCallOrIndexerExpr.SetIsUsedWithArgs( const Value : boolean );
    begin
      assert( fOp1 <> nil );
      if value
        then Include( fOp1.fFlags, expr_UsedWithArgs )
        else Exclude( fOp1.fFlags, expr_UsedWithArgs );
    end;

  procedure TvbCallOrIndexerExpr.SetIsUsedWithEmptyParens( const Value : boolean );
    begin
      assert( fOp1 <> nil );
      if value
        then Include( fOp1.fFlags, expr_MustBeFunc )
        else Exclude( fOp1.fFlags, expr_MustBeFunc );
    end;

  { TvbBinaryExpr }

  procedure TvbBinaryExpr.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnBinaryExpr( self, Context );
    end;

  constructor TvbBinaryExpr.Create( op : TvbOperator; left, right : TvbExpr; paren : boolean );
    begin
      inherited Create( BINARY_EXPR );
      fOperator := op;
      fOp1      := left;
      fOp2      := right;
      fParen    := paren;
    end;

  { TvbUnaryExpr }

  procedure TvbUnaryExpr.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnUnaryExpr( self, Context );
    end;

  constructor TvbUnaryExpr.Create( op : TvbOperator; expr : TvbExpr );
    begin
      inherited Create( UNARY_EXPR );
      fOperator := op;
      fOp1      := expr;
    end;

  { TvbLabelDefStmt }

  procedure TvbLabelDefStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnLabelDefStmt( self, Context );
    end;

  constructor TvbLabelDefStmt.Create( lab : TvbLabelDef );
    begin
      inherited Create( LABEL_DEF_STMT );
      fLabel := lab;
    end;

  { TvbAssignStmt }

  constructor TvbAssignStmt.Create( assign : TvbAssignKind; dest, source : TvbExpr );
    begin
      inherited Create( ASSIGN_STMT );
      fKind  := assign;
      fValue := source;
      fLHS   := dest;
    end;

  constructor TvbAssignStmt.Create( dest, source : TvbExpr );
    begin
      Create( asLet, dest, source );
    end;

  procedure TvbAssignStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnAssignStmt( self, Context );
    end;

  destructor TvbAssignStmt.Destroy;
    begin
      FreeAndNil( fValue );
      FreeAndNil( fLHS );
      inherited;
    end;

  { TvbIfStmt }

  procedure TvbIfStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnIfStmt( self, Context );
    end;

  function TvbIfStmt.AddElseIf : TvbIfStmt;
    begin
      if fElseIfs = nil
        then fElseIfs := TvbNodeList.Create;
      result := TvbIfStmt( fElseIfs.Add( TvbIfStmt.Create ) );
    end;

  constructor TvbIfStmt.Create( expr      : TvbExpr;
                                thenStmts : TvbStmtBlock;
                                elseStmts : TvbStmtBlock;
                                elseIfs   : TvbNodeList );
    begin
      inherited Create( IF_STMT );
      fExpr    := expr;
      fThen    := thenStmts;
      fElse    := elseStmts;
      fElseIfs := elseIfs;
    end;

  destructor TvbIfStmt.Destroy;
    begin
      FreeAndNil( fExpr );
      FreeAndNil( fThen );
      FreeAndNil( fElse );
      FreeAndNil( fElseIfs );
      inherited;
    end;

  function TvbIfStmt.GetElseIf( i : integer ) : TvbIfStmt;
    begin
      result := TvbIfStmt( fElseIfs.Nodes[i] );
    end;

  function TvbIfStmt.GetElseIfCount : integer;
    begin
      if fElseIfs = nil
        then result := 0
        else result := fElseIfs.Count;
    end;

  { TvbCaseClause }

  procedure TvbCaseClause.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnCaseClause( self, Context );
    end;

  constructor TvbCaseClause.Create( nk : TvbNodeKind; expr : TvbExpr );
    begin
      inherited Create( nk );
      fExpr := expr;
    end;

  destructor TvbCaseClause.Destroy;
    begin
      FreeAndNil( fExpr );
      inherited;
    end;

  { TvbRelationalCaseClause }

  procedure TvbRelationalCaseClause.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnRelationalCaseClause( self, Context );
    end;

  constructor TvbRelationalCaseClause.Create( anOper : TvbOperator; expr : TvbExpr );
    begin
      inherited Create( REL_CASE_CLAUSE, expr );
      fOper := anOper;
    end;

  { TvbRangeCaseClause }

  procedure TvbRangeCaseClause.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnRangeCaseClause( self, Context );
    end;

  constructor TvbRangeCaseClause.Create( aLowerExpr, aUpperExpr : TvbExpr );
    begin
      inherited Create( RANGE_CASE_CLAUSE, aLowerExpr );
      fUpperExpr := aUpperExpr;
    end;

  destructor TvbRangeCaseClause.Destroy;
    begin
      FreeAndNil( fUpperExpr );
      inherited;
    end;

  { TvbCaseStmt }

  procedure TvbCaseStmt.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnCaseStmt( self, Context );
    end;

  function TvbCaseStmt.AddCaseClause( caseClause : TvbCaseClause ) : TvbCaseClause;
    begin
      fClauses.Add( caseClause );
      result := caseClause;
    end;

  constructor TvbCaseStmt.Create;
    begin
      inherited Create( CASE_STMT );
      fClauses := TvbNodeList.Create;
    end;

  destructor TvbCaseStmt.Destroy;
    begin
      FreeAndNil( fClauses );
      FreeAndNil( fStmts );
      inherited;
    end;

  function TvbCaseStmt.GetClause( i : integer ) : TvbCaseClause;
    begin
      result := TvbCaseClause( fClauses[i] );
    end;

  function TvbCaseStmt.GetClauseCount : integer;
    begin
      result := fClauses.Count;
    end;

  { TvbSelectCaseStmt }

  procedure TvbSelectCaseStmt.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnSelectCaseStmt( self, Context );
    end;

  function TvbSelectCaseStmt.AddCaseStmt( caseStmt : TvbCaseStmt ) : TvbCaseStmt;
    begin
      fCases.Add( caseStmt );
      result := caseStmt;
    end;

  constructor TvbSelectCaseStmt.Create( expr : TvbExpr );
    begin
      inherited Create( SELECT_CASE_STMT );
      fCases := TvbNodeList.Create;
      fExpr  := expr;
    end;

  destructor TvbSelectCaseStmt.Destroy;
    begin
      FreeAndNil( fExpr );
      FreeAndNil( fCases );
      FreeAndNil( fCaseElse );
      inherited;
    end;

  function TvbSelectCaseStmt.GetCase( i : integer ) : TvbCaseStmt;
    begin
      result := TvbCaseStmt( fCases[i] );
    end;

  function TvbSelectCaseStmt.GetCaseCount : integer;
    begin
      result := fCases.Count;
    end;

  { TvbDoLoopStmt }

  procedure TvbDoLoopStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnDoLoopStmt( self, Context );
    end;

  constructor TvbDoLoopStmt.Create( aLoop       : TvbDoLoop;
                                    TestAtBegin : boolean;
                                    aCond       : TvbExpr;
                                    stmts       : TvbStmtBlock );
    begin
      inherited Create( DO_LOOP_STMT );
      fLoop        := aLoop;
      fCond        := aCond;
      fTestAtBegin := TestAtBegin;
      fBlock       := Stmts;
    end;

  destructor TvbDoLoopStmt.Destroy;
    begin
      FreeAndNil( fCond );
      FreeAndNil( fBlock );
      inherited;
    end;

  { TvbWithStmt }

  procedure TvbWithStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnWithStmt( self, Context );
    end;

  constructor TvbWithStmt.Create( expr : TvbExpr; stmts : TvbStmtBlock );
    begin
      inherited Create( WITH_STMT );
      fObjOrRec  := expr;
      fBlock := stmts;
    end;

  destructor TvbWithStmt.Destroy;
    begin
      FreeAndNil( fObjOrRec );
      FreeAndNil( fBlock );
      inherited;
    end;

  { TvbExitStmt }

  procedure TvbExitStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnExitStmt( self, Context );
    end;

  constructor TvbExitStmt.Create;
    begin
      inherited Create( EXIT_STMT );
    end;

  { TvbSimpleName }

  procedure TvbSimpleName.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnSimpleName( self, Context );
    end;

  constructor TvbSimpleName.Create( const name : string );
    begin
      inherited Create( SIMPLE_NAME, name );
    end;

  { TvbQualifiedName }

  procedure TvbQualifiedName.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnQualifiedName( self, Context );
    end;

  constructor TvbQualifiedName.Create( const name : string; qual : TvbName );
    begin
      inherited Create( QUALIFIED_NAME, name );
      fQualifier := qual;
    end;

  destructor TvbQualifiedName.Destroy;
    begin
      FreeAndNil( fQualifier );
      inherited;
    end;

  { TvbNameExpr }

  procedure TvbNameExpr.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnNameExpr( self, Context );
    end;

  constructor TvbNameExpr.Create( const name : string );
    begin
      inherited Create( NAME_EXPR );
      fName := name;
    end;

  { TvbAddressOfExpr }

  procedure TvbAddressOfExpr.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnAddressOfExpr( self, Context );
    end;

  constructor TvbAddressOfExpr.Create( aFuncName : TvbName );
    begin
      inherited Create( ADDRESSOF_EXPR );
      fFuncName := aFuncName;
    end;

  destructor TvbAddressOfExpr.Destroy;
    begin
      FreeAndNil( fFuncName );
      inherited;
    end;

  { TvbMemberAccessExpr }

  procedure TvbMemberAccessExpr.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnMemberAccessExpr( self, Context );
    end;

  constructor TvbMemberAccessExpr.Create( const aMember : string; aTarget : TvbExpr; withExpr : TvbExpr );
    begin
      inherited Create( MEMBER_ACCESS_EXPR );
      fMember   := aMember;
      fOp1      := aTarget;
      fWithExpr := withExpr;
    end;

  { TvbDictAccessExpr }

  procedure TvbDictAccessExpr.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnDictAccessExpr( self, Context );
    end;

  constructor TvbDictAccessExpr.Create( const name : string; aTarget : TvbExpr; withExpr : TvbExpr );
    begin
      inherited Create( DICT_ACCESS_EXPR );
      fName     := name;
      fOp1      := aTarget;
      fWithExpr := withExpr;
    end;

  { TvbMeExpr }

  procedure TvbMeExpr.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnMeExpr( self, Context );
    end;

  constructor TvbMeExpr.Create( typ : IvbType );
    begin
      inherited Create( ME_EXPR );
      fNodeType := typ;
    end;

  { TvbReturnStmt }

  procedure TvbReturnStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnReturnStmt( self, Context );
    end;

  constructor TvbReturnStmt.Create;
    begin
      inherited Create( RETURN_STMT );
    end;

  { TvbGotoOrGosubStmt }

  procedure TvbGotoOrGosubStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnGotoOrGosubStmt( self, Context );
    end;

  constructor TvbGotoOrGosubStmt.Create( kind : TvbGotoOrGosub; labelName : TvbSimpleName );
    begin
      inherited Create( GOTO_OR_GOSUB_STMT );
      fKind  := kind;
      fLabel := labelName;
    end;

  destructor TvbGotoOrGosubStmt.Destroy;
     begin
      FreeAndNil( fLabel );
      inherited;
    end;

  { TvbNewExpr }

  procedure TvbNewExpr.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnNewExpr( self, Context );
    end;

  constructor TvbNewExpr.Create( typ : TvbExpr );
    begin
      inherited Create( NEW_EXPR, typ );
    end;

  { TvbOnGotoOrGosubStmt }

  constructor TvbOnGotoOrGosubStmt.Create( expr : TvbExpr; gg : TvbGotoOrGosub );
    begin
      inherited Create( ONGOTO_OR_GOSUB_STMT );
      fExpr        := expr;
      fGotoOrGosub := gg;
      fLabels      := TvbNodeList.Create;
    end;

  destructor TvbOnGotoOrGosubStmt.Destroy;
    begin
      FreeAndNil( fExpr );
      FreeAndNil( fLabels );
      inherited;
    end;

  procedure TvbOnGotoOrGosubStmt.AddLabel( lbl : TvbSimpleName );
    begin
      fLabels.Add( lbl );
    end;

  function TvbOnGotoOrGosubStmt.GetLabelCount : integer;
    begin
      result := fLabels.Count;
    end;

  function TvbOnGotoOrGosubStmt.GetLabelName( i : integer ) : TvbSimpleName;
    begin
      result := TvbSimpleName( fLabels.Nodes[i] );
    end;

  procedure TvbOnGotoOrGosubStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnOnGotoOrGosubStmt( self, Context );
    end;

  { TvbOnErrorStmt }

  procedure TvbOnErrorStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnOnErrorStmt( self, Context );
    end;

  constructor TvbOnErrorStmt.Create( onErrKind : TvbOnErrorDo; labelName : TvbSimpleName );
    begin
      inherited Create( ON_ERROR_STMT );
      fOnError := onErrKind;
      fLabel   := labelName;
    end;

  destructor TvbOnErrorStmt.Destroy;
    begin
      FreeAndNil( fLabel );
      inherited;
    end;

  { TvbRaiseEventStmt }

  procedure TvbRaiseEventStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnRaiseEventStmt( self, Context );
    end;

  constructor TvbRaiseEventStmt.Create( const event : string; args : TvbArgList );
    begin
      inherited Create( RAISE_EVENT_STMT );
      fEventName := event;
      fArgs      := args;
    end;

  destructor TvbRaiseEventStmt.Destroy;
    begin
      FreeAndNil( fArgs );
      inherited;
    end;

  { TvbReDimStmt }

  procedure TvbReDimStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnReDimStmt( self, Context );
    end;

  constructor TvbReDimStmt.Create( preserve : boolean; arrVar : TvbVarDef );
    begin
      inherited Create( REDIM_STMT );
      fArrVar   := arrVar;
      fPreserve := preserve;
    end;

  destructor TvbReDimStmt.Destroy;
    begin
      // IMPORTANT! We destroy the variable object ONLY IF it isn't flagged
      // because if if is, the variable is destroy by the containing function.
      if not ( nfReDim_DeclaredLocalVar in fArrVar.NodeFlags )
        then FreeAndNil( fArrVar );
      inherited;
    end;

  { TvbOpenStmt }

  procedure TvbOpenStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnOpenStmt( self, Context );
    end;

  constructor TvbOpenStmt.Create( filePath : TvbExpr;
                                  mode     : TvbOpenMode;
                                  access   : TvbOpenAccess;
                                  lock     : TvbOpenLock;
                                  fileNum  : TvbExpr;
                                  recLen   : TvbExpr);
    begin
      inherited Create( OPEN_STMT );
      fFilePath := filePath;
      fMode     := mode;
      fAccess   := access;
      fLock     := lock;
      fFileNum  := fileNum;
      fRecLen   := recLen;
    end;

  destructor TvbOpenStmt.Destroy;
    begin
      FreeAndNil( fFilePath );
      FreeAndNil( fFileNum );
      FreeAndNil( fRecLen );
      inherited;
    end;

  { TvbArgList }

  constructor TvbArgList.Create;
    begin
      fArgIndex         := nil;
      fArgs             := TvbNodeList.Create;
      fFirstNamedArgIdx := -1;
    end;

  destructor TvbArgList.Destroy;
    begin
      FreeAndNil( fArgIndex );
      FreeAndNil( fArgs );
      inherited;
    end;

  procedure TvbArgList.Add( arg : TvbExpr );
    var
      namedArg : TvbNamedArgExpr absolute arg;
    begin
      // if arg=nil, this is an omitted argument
      if ( arg <> nil ) and ( arg.NodeKind = NAMED_ARG_EXPR )
        then
          begin
            if fArgIndex = nil
              then
                begin
                  fArgIndex               := TStringList.Create;
                  fArgIndex.CaseSensitive := false;
                  fArgIndex.Sorted        := true;
                  fArgIndex.Duplicates    := dupError;
                  fFirstNamedArgIdx       := fArgs.Count;
                end;
            assert( namedArg.Value <> nil );
            fArgIndex.AddObject( namedArg.ArgName, namedArg.Value );
          end;
      fArgs.Add( arg );
    end;

  function TvbArgList.GetArg( i : integer ) : TvbExpr;
    begin
      if (i < fArgs.Count) and ( fArgs[i] <> nil )
        then
          if fArgs[i].NodeKind = NAMED_ARG_EXPR
            then result := TvbNamedArgExpr( fArgs[i] ).Value
            else result := TvbExpr( fArgs[i] )
        else result := nil;
    end;

  function TvbArgList.GetNamedArg( const name : string ) : TvbExpr;
    var
      i : integer;
    begin
      if ( fArgIndex <> nil ) and fArgIndex.Find( name, i )
        then result := TvbExpr( fArgIndex.Objects[i] )
        else result := nil;
    end;

  function TvbArgList.GetArgCount : integer;
    begin
      result := fArgs.Count;
    end;

  procedure TvbArgList.Resolve( CurrentScope : TvbScope );
    begin
      assert( fArgs <> nil );
      //fArgs.ResolveExprs( currentScope );
    end;

  procedure TvbArgList.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    var
      i : integer;
    begin
      for i := 0 to fArgs.Count - 1 do
        if fArgs[i] <> nil
          then TvbExpr( fArgs[i] ).Accept( Visitor, Context );
    end;

  { TvbNodeList }

  procedure TvbNodeList.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    var
      i : integer;
    begin
      for i := 0 to fNodes.Count - 1 do
        begin
          assert( fNodes[i] <> nil );
          TvbNode( fNodes[i] ).Accept( Visitor, Context );
        end;
    end;

  function TvbNodeList.Add( node : TvbNode ) : TvbNode;
    begin
      result := node;
      fNodes.Add( node );
    end;

  constructor TvbNodeList.Create( OwnsNodes : boolean );
    begin
      fNodes := TObjectList.Create( OwnsNodes );
    end;

  destructor TvbNodeList.Destroy;
    begin
      FreeAndNil( fNodes );
      inherited;
    end;

  function TvbNodeList.GetNode( i : integer ) : TvbNode;
    begin
      result := TvbNode( fNodes[i] );
    end;

  function TvbNodeList.GetNodeCount : integer;
    begin
      result := fNodes.Count;
    end;

  { TvbVarDef }

  procedure TvbVarDef.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnVarDef( self, Context );
    end;

  constructor TvbVarDef.Create( const flags : TvbNodeFlags; const name : string; const typ : IvbType; const name1 : string );
    begin
      inherited Create( VAR_DEF, flags, name, name1 );
      fNodeType := typ;
    end;

  destructor TvbVarDef.Destroy;
    begin
      fNodeType := nil;
      inherited;
    end;

  { TvbConstDef }

  procedure TvbConstDef.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnConstDef( self, Context );
    end;

  constructor TvbConstDef.Create( const flags : TvbNodeFlags;
                                  const name  : string;
                                  const typ   : IvbType;
                                  val         : TvbExpr;
                                  const name1 : string );
    begin
      inherited Create( CONST_DEF, flags, name, name1 );
      fValue    := val;
      fNodeType := typ;
    end;

  destructor TvbConstDef.Destroy;
    begin
      FreeAndNil( fValue );
      inherited;
    end;

  { TvbTypeDef }

  constructor TvbTypeDef.Create( kind        : TvbNodeKind;
                                 const flags : TvbNodeFlags;
                                 const name  : string;
                                 const name1 : string );
    begin
      inherited Create( kind, flags, name, name1 );
      fScope := TvbScope.Create;
      {$IFDEF DEBUG}
      fScope.Name := name;
      {$ENDIF}
    end;

  destructor TvbTypeDef.Destroy;
    begin
      FreeAndNil( fScope );
      inherited;
    end;

  { TvbEnumDef }

  procedure TvbEnumDef.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnEnumDef( self, Context );
    end;

  procedure TvbEnumDef.AddEnum( e : TvbConstDef );
    begin
      assert( e.NodeKind = CONST_DEF );
      fEnums.Add( e );
      fScope.Bind( e, OBJECT_BINDING );
    end;

  constructor TvbEnumDef.Create( flags : TvbNodeFlags; const name, name1 : string );
    begin
      inherited Create( ENUM_DEF, flags, name, name1 );
      fEnums := TvbNodeList.Create;
    end;

  destructor TvbEnumDef.Destroy;
    begin
      FreeAndNil( fEnums );
      inherited;
    end;

  function TvbEnumDef.GetEnum( i : integer ) : TvbConstDef;
    begin
      result := TvbConstDef( fEnums[i] );
    end;

  function TvbEnumDef.GetEnumCount : integer;
    begin
      result := fEnums.Count;
    end;

  { TvbRecordDef }

  procedure TvbRecordDef.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnRecordDef( self, Context );
    end;

  procedure TvbRecordDef.AddField( f : TvbVarDef );
    begin
      assert( f.NodeKind = VAR_DEF );
      fFields.Add( f );
      fScope.Bind( f, OBJECT_BINDING );
    end;

  constructor TvbRecordDef.Create( flags : TvbNodeFlags; const name, name1 : string );
    begin
      inherited Create( RECORD_DEF, flags, name, name1 );
      fFields := TvbNodeList.Create;
    end;

  destructor TvbRecordDef.Destroy;
    begin
      FreeAndNil( fFields );
      inherited;
    end;

  function TvbRecordDef.GetField( i : integer ) : TvbVarDef;
    begin
      result := TvbVarDef( fFields[i] );
    end;

  function TvbRecordDef.GetFieldCount : integer;
    begin
      result := fFields.Count;
    end;

  { TvbForStmt }

  procedure TvbForStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnForStmt( self, Context );
    end;

  constructor TvbForStmt.Create( counter : TvbExpr;
                                 init    : TvbExpr;
                                 max     : TvbExpr;
                                 step    : TvbExpr;
                                 stmts   : TvbStmtBlock );
    begin
      inherited Create( FOR_STMT );
      fCounter  := counter;
      fInitVal  := init;
      fFinalVal := max;
      fStepVal  := step;
      fBlock    := Stmts;
    end;

  destructor TvbForStmt.Destroy;
    begin
      FreeAndNil( fCounter );
      FreeAndNil( fInitVal );
      FreeAndNil( fFinalVal );
      FreeAndNil( fStepVal );
      FreeAndNil( fBlock );
      inherited;
    end;

  { TvbForeachStmt }

  procedure TvbForeachStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnForeachStmt( self, Context );
    end;

  constructor TvbForeachStmt.Create( elem, group : TvbExpr; stmts : TvbStmtBlock );
    begin
      inherited Create( FOREACH_STMT );
      fElement := elem;
      fGroup   := group;
      fBlock   := stmts;
    end;

  destructor TvbForeachStmt.Destroy;
    begin
      FreeAndNil( fElement );
      FreeAndNil( fGroup );
      FreeAndNil( fBlock );
      inherited;
    end;

  { TvbMidAssignStmt }

  procedure TvbMidAssignStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnMidAssignStmt( self, Context );
    end;

  constructor TvbMidAssignStmt.Create( varName, aStartPos, aLen, aRep : TvbExpr );
    begin
      inherited Create( MID_ASSIGN_STMT );
      fVarName  := varName;
      fStartPos := aStartPos;
      fLen      := aLen;
      fRep      := aRep;
    end;

  destructor TvbMidAssignStmt.Destroy;
    begin
      FreeAndNil( fVarName );
      FreeAndNil( fStartPos );
      FreeAndNil( fRep );
      FreeAndNil( fLen );
      inherited;
    end;

  { TvbEraseStmt }

  procedure TvbEraseStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnEraseStmt( self, Context );
    end;

  constructor TvbEraseStmt.Create( arrVar : TvbExpr );
    begin
      inherited Create( ERASE_STMT );
      fArrVar := arrVar;
    end;

  destructor TvbEraseStmt.Destroy;
    begin
      FreeAndNil( fArrVar );
      inherited;
    end;

  { TvbEventDef }

  procedure TvbEventDef.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnEventDef( self, Context );
    end;

  procedure TvbEventDef.AddParam( p : TvbParamDef );
    begin
      assert( p <> nil );
      if fParams = nil
        then fParams := TvbParamList.Create( true );
      fParams.AddParam( p );
    end;

  constructor TvbEventDef.Create( const name : string; params : TvbParamList );
    begin
      inherited Create( EVENT_DEF, [], name );
      fParams := params; 
    end;

  destructor TvbEventDef.Destroy;
    begin
      FreeAndNil( fParams );
      inherited;
    end;

  function TvbEventDef.GetParam( i : integer ) : TvbParamDef;
    begin
      result := fParams.Params[i];
    end;

  function TvbEventDef.GetParamCount : integer;
    begin
      if fParams = nil
        then result := 0
        else result := fParams.ParamCount;
    end;

  { TvbFuncDef }

  constructor TvbFuncDef.Create( flags       : TvbNodeFlags;
                                 const name  : string;
                                 params      : TvbParamList;
                                 const typ   : IvbType;
                                 stmts       : TvbStmtBlock;
                                 const name1 : string );
    begin
      inherited Create( FUNC_DEF, flags, name, name1 );
      fParams         := params;
      fNodeType       := typ;
      fBlock          := stmts;
      fStaticVarCount := 0;
      fScope          := TvbScope.Create;
      {$IFDEF DEBUG}
      fScope.Name := name;
      {$ENDIF}
    end;

  destructor TvbFuncDef.Destroy;
    var
      i : integer;
    begin
      FreeAndNil( fParams );
      FreeAndNil( fBlock );
      if fLocalVars <> nil
        then
          for i := 0 to fLocalVars.Count - 1 do
            fLocalVars.Objects[i].Free;
      FreeAndNil( fLocalVars );
      if fLocalConsts <> nil
        then
          for i := 0 to fLocalConsts.Count - 1 do
            fLocalConsts.Objects[i].Free;
      FreeAndNil( fLocalConsts );
      if fLabels <> nil
        then
          for i := 0 to fLabels.Count - 1 do
            fLabels.Objects[i].Free;
      FreeAndNil( fLabels );
      FreeAndNil( fScope );
      inherited;
    end;

  procedure TvbFuncDef.AddLocalConst( cst : TvbConstDef );
    begin
      assert( cst <> nil );
      if fLocalConsts = nil
        then  fLocalConsts := TStringList.Create;
      fLocalConsts.AddObject( cst.Name, cst );
      fScope.Bind( cst, OBJECT_BINDING );
    end;

  procedure TvbFuncDef.AddLocalVar( v : TvbVarDef );
    begin
      assert( v <> nil );
      assert( v.VarType <> nil );
      if fLocalVars = nil
        then  fLocalVars := TStringList.Create;
      fLocalVars.AddObject( v.Name, v );
      if dfStatic in v.NodeFlags
        then inc( fStaticVarCount );
      fScope.Bind( v, OBJECT_BINDING );
    end;

  function TvbFuncDef.GetConst( i : integer ) : TvbConstDef;
    begin
      result := TvbConstDef( fLocalConsts.Objects[i] );
    end;

  function TvbFuncDef.GetConstCount : integer;
    begin
      if fLocalConsts = nil
        then result := 0
        else result := fLocalConsts.Count;
    end;

  function TvbFuncDef.GetVar( i : integer ) : TvbVarDef;
    begin
      result := TvbVarDef( fLocalVars.Objects[i] );
    end;

  function TvbFuncDef.GetVarCount : integer;
    begin
      if fLocalVars = nil
        then result := 0
        else result := fLocalVars.Count;
    end;

  function TvbFuncDef.GetLabel( i : integer ) : string;
    begin
      result := fLabels.Strings[i];
    end;

  function TvbFuncDef.GetLabelCount : integer;
    begin
      if fLabels = nil
        then result := 0
        else result := fLabels.Count;
    end;

  procedure TvbFuncDef.AddLabel( lab : TvbLabelDef );
    begin
      assert( lab <> nil );
      if fLabels = nil
        then  fLabels := TStringList.Create;
      fLabels.AddObject( lab.Name, lab );
      fScope.Bind( lab, LABEL_BINDING );
    end;

  procedure TvbFuncDef.AddLocalVar( const name : string; typ : TvbType );
    begin
      AddLocalVar( TvbVarDef.Create( [], name, typ ) );
    end;

  function TvbFuncDef.GetParamArrayIndex : integer;
    begin
      if fParams = nil
        then result := -1
        else result := fParams.ParamArrayIndex;
    end;

  procedure TvbFuncDef.AddParam( p : TvbParamDef );
    begin
      assert( p <> nil );
      assert( p.ParamType <> nil );
      if fParams = nil
        then fParams := TvbParamList.Create( true );
      fParams.AddParam( p );
      fScope.Bind( p, OBJECT_BINDING );
    end;

  procedure TvbFuncDef.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnFunctionDef( self, Context );
    end;

  function TvbFuncDef.GetParam( i : integer ) : TvbParamDef;
    begin
      result := fParams[i];
    end;

  function TvbFuncDef.GetParamCount : integer;
    begin
      if fParams <> nil
        then result := fParams.ParamCount
        else result := 0;
    end;

  function TvbFuncDef.GetParams: TvbParamList;
    begin
      result := fParams;
    end;

  function TvbFuncDef.GetScope: TvbScope;
    begin
      result := fScope;
    end;

  function TvbFuncDef.GetStaticVarCount: integer;
    begin
      result := fStaticVarCount;
    end;

  function TvbFuncDef.GetStmtBlock: TvbStmtBlock;
    begin
      result := fBlock;
    end;

  procedure TvbFuncDef.SetStmtBlock( block : TvbStmtBlock );
    begin
      fBlock := block;
    end;

  function TvbFuncDef.GetArgc: integer;
    begin
      result := GetParamCount;
    end;

  { TvbAbstractFuncDef }

  procedure TvbAbstractFuncDef.AddLabel( lab : TvbLabelDef );
    begin
      assert( false );
    end;

  procedure TvbAbstractFuncDef.AddLocalConst( cst : TvbConstDef );
    begin
      assert( false );
    end;

  procedure TvbAbstractFuncDef.AddLocalVar( v : TvbVarDef );
    begin
      assert( false );
    end;

  procedure TvbAbstractFuncDef.AddLocalVar( const name : string; typ : TvbType );
    begin
      assert( false );
    end;

  procedure TvbAbstractFuncDef.AddParam( p : TvbParamDef );
    begin
      assert( false );
    end;

  function TvbAbstractFuncDef.GetLabel( i : integer ) : string;
    begin
      assert( false );
      result := '';
    end;

  function TvbAbstractFuncDef.GetLabelCount : integer;
    begin
      result := 0;
    end;

  function TvbAbstractFuncDef.GetConst( i : integer ) : TvbConstDef;
    begin
      assert( false );
      result := nil;
    end;

  function TvbAbstractFuncDef.GetConstCount : integer;
    begin
      result := 0;
    end;

  function TvbAbstractFuncDef.GetVar( i : integer ) : TvbVarDef;
    begin
      assert( false );
      result := nil;
    end;

  function TvbAbstractFuncDef.GetVarCount : integer;
    begin
      result := 0;
    end;

  function TvbAbstractFuncDef.GetParam( i : integer ) : TvbParamDef;
    begin
      assert( false );
      result := nil;
    end;

  function TvbAbstractFuncDef.GetParamArrayIndex : integer;
    begin
      result := -1;
    end;

  function TvbAbstractFuncDef.GetParamCount : integer;
    begin
      result := 0;
    end;

  function TvbAbstractFuncDef.GetParams : TvbParamList;
    begin
      result := nil;
    end;

  function TvbAbstractFuncDef.GetScope : TvbScope;
    begin
      result := nil;
    end;

  function TvbAbstractFuncDef.GetStaticVarCount : integer;
    begin
      result := 0;
    end;

  function TvbAbstractFuncDef.GetStmtBlock : TvbStmtBlock;
    begin
      result := nil;
    end;

  procedure TvbAbstractFuncDef.SetStmtBlock( block : TvbStmtBlock );
    begin
      assert( false );
    end;

  function TvbAbstractFuncDef.GetArgc : integer;
    begin
      result := 0;
    end;

  { TvbDllFuncDef }

  procedure TvbDllFuncDef.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnDllFuncDef( self, Context );
    end;

  procedure TvbDllFuncDef.AddParam( p : TvbParamDef );
    begin
      if fParams = nil
        then fParams := TvbParamList.Create( true );
      fParams.AddParam( p );
      if ( p.ParamType.Code = STRING_TYPE ) and
         ( pfByVal in p.NodeFlags )
        then p.NodeFlags := p.NodeFlags + [_pfPChar]
        else
      if p.ParamType = vbBooleanType
        then p.NodeFlags := p.NodeFlags + [_pfWordBool]
        else
      if p.ParamType = vbAnyType
        then
          if pfByVal in p.NodeFlags
            then p.NodeFlags := p.NodeFlags + [_pfAnyByVal]
            else p.NodeFlags := p.NodeFlags + [_pfAnyByRef];
    end;

  constructor TvbDllFuncDef.Create( flags         : TvbNodeFlags;
                                    const name    : string;
                                    const dllName : string;
                                    const alias   : string;
                                    params        : TvbParamList;
                                    const typ     : IvbType );
    begin
      inherited Create( DLL_FUNC_DEF, flags, name, name1 );
      fParams   := params;
      fNodeType := typ;
      fDllName  := dllName;
      fAlias    := alias;
    end;

  destructor TvbDllFuncDef.Destroy;
    begin
      FreeAndNil( fParams );
      inherited;
    end;

  function TvbDllFuncDef.GetParam( i : integer ) : TvbParamDef;
    begin
      if fParams = nil
        then result := nil
        else result := fParams[i];
    end;

  function TvbDllFuncDef.GetParamCount : integer;
    begin
      if fParams = nil
        then result := 0
        else result := fParams.ParamCount;
    end;

  function TvbDllFuncDef.GetParams : TvbParamList;
    begin
      result := fParams;
    end;

  { TvbExternalFuncDef }

  procedure TvbExternalFuncDef.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
      //Visitor.OnFunctionDef( self);
    end;

  constructor TvbExternalFuncDef.Create( flags       : TvbNodeFlags;
                                         const name  : string;
                                         const name1 : string;
                                         argc        : integer;
                                         const typ   : IvbType );
    begin
      inherited Create( FUNC_DEF, flags, name, name1 );
      fNodeType := typ;
      fArgc := argc;
    end;

  function TvbExternalFuncDef.GetArgc : integer;
    begin
      result := fArgc;
    end;

  { TvbPropertyDef }

  procedure TvbPropertyDef.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnPropertyDef( self, Context );
    end;

  constructor TvbPropertyDef.Create( const name : string; flags : TvbNodeFlags );
    begin
      inherited Create( PROPERTY_DEF, flags, name );
      fParams := TvbParamList.Create( false );
    end;

  destructor TvbPropertyDef.Destroy;
    begin
      // Note that we don't destroy the parameter objects inside the
      // string list because they are owned by the getter/setter functions.
      FreeAndNil( fParams );
      FreeAndNil( fGetterFunc );
      FreeAndNil( fSetterFunc );
      FreeAndNil( fLetterFunc );
      inherited;
    end;

  function TvbPropertyDef.GetParam( i : integer ) : TvbParamDef;
    begin
      result := fParams.Params[i];
    end;

  function TvbPropertyDef.GetParamCount : integer;
    begin
      result := fParams.ParamCount;
    end;

  procedure TvbPropertyDef.SetParamsAndType;
    var
      i, lastParamIdx : integer;
    begin
      if fGetterFunc <> nil
        then
          begin
            fNodeType := fGetterFunc.ReturnType;
            for i := 0 to fGetterFunc.ParamCount - 1 do
              fParams.AddParam( fGetterFunc.Param[i] );
          end
        else
      if fSetterFunc <> nil
        then
          begin
            lastParamIdx := fSetterFunc.ParamCount - 1;
            if lastParamIdx >= 0
              then
                begin
                  fNodeType := fSetterFunc.Param[lastParamIdx].ParamType;
                  for i := 0 to lastParamIdx - 1 do
                    fParams.AddParam( fSetterFunc.Param[i] );
                end;
          end
        else
      if fLetterFunc <> nil
        then
          begin
            lastParamIdx := fLetterFunc.ParamCount - 1;
            if lastParamIdx >= 0
              then
                begin
                  fNodeType := fLetterFunc.Param[lastParamIdx].ParamType;
                  for i := 0 to lastParamIdx - 1 do
                    fParams.AddParam( fLetterFunc.Param[i] );
                end;
          end;
    end;

  { TvbTypeOfExpr }

  procedure TvbTypeOfExpr.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnTypeOfExpr( self, Context );
    end;

  constructor TvbTypeOfExpr.Create( obj : TvbExpr; typeName : TvbName );
    begin
      inherited Create( TYPEOF_EXPR );
      fOp1      := obj;
      fTypeName := typeName;
    end;

  destructor TvbTypeOfExpr.Destroy;
    begin
      FreeAndNil( fTypeName );
      inherited;
    end;

  { TvbImplementedInterface }

  procedure TvbImplementedInterface.AddImplementedFunction( func : TvbFuncDef );
    begin
      if not Assigned( fFuncs )
        then fFuncs := TvbNodeList.Create; // owned list
      fFuncs.Add( func );
    end;

  constructor TvbImplementedInterface.Create( name : TvbQualifiedName );
    begin
      inherited Create( nkImplIntf );
      fName := name;
    end;

  destructor TvbImplementedInterface.Destroy;
    begin
      FreeAndNil( fName );
      FreeAndNil( fFuncs );
      inherited;
    end;

  function TvbImplementedInterface.GetFunctionCount : integer;
    begin
      if Assigned( fFuncs )
        then Result := fFuncs.Count
        else Result := 0;
    end;

  function TvbImplementedInterface.GetFunction( i : integer ) : TvbFuncDef;
    begin
      Result := nil;
      if fFuncs <> nil
        then
          if i < fFuncs.Count
            then Result := TvbFuncDef( fFuncs.Nodes[i] );
    end;

  { TvbParamList }

  procedure TvbParamList.AddParam( p : TvbParamDef );
    begin
      fParams.AddObject( p.Name, p );
      if pfParamArray in p.NodeFlags
        then
          begin
            assert( fParamArrayIdx = -1 );
            assert( fOptParamCount = 0 );
            fParamArrayIdx := fParams.Count - 1;
          end
        else
      if pfOptional in p.NodeFlags
        then inc( fOptParamCount );
    end;

//  constructor TvbParamList.Create( p : TvbParamDef );
//    begin
//      Create;
//      AddParam( p );
//    end;

  constructor TvbParamList.Create( OwnsParams : boolean );
    begin
      fOwnsParams := OwnsParams;
      fParams := TStringList.Create;
      fParamArrayIdx := -1;
      fOptParamCount := 0;
    end;

  destructor TvbParamList.Destroy;
    var
      i : integer;
    begin
      if fOwnsParams
        then
          for i := 0 to fParams.Count - 1 do
            fParams.Objects[i].Free;
      fParams.Free;
      inherited;
    end;

  function TvbParamList.GetParam( i : integer ) : TvbParamDef;
    begin
      if i < fParams.Count
        then result := TvbParamDef( fParams.Objects[i] )
        else result := nil;
    end;

  function TvbParamList.GetParamByName( n : string ) : TvbParamDef;
    var
      i : integer;
    begin
      i := fParams.IndexOf( n );
      if i <> -1
        then result := TvbParamDef( fParams.Objects[i] )
        else result := nil;
    end;

  function TvbParamList.GetParamCount : integer;
    begin
      result := fParams.Count;
    end;

  { TvbCloseStmt }

  procedure TvbCloseStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnCloseStmt( self, Context );
    end;

  constructor TvbCloseStmt.Create( fileNum : TvbExpr );
    begin
      inherited Create( CLOSE_STMT );
      fFileNum := fileNum;
    end;

  destructor TvbCloseStmt.Destroy;
    begin
      FreeAndNil( fFileNum );
      inherited;
    end;

  { TvbGetStmt }

  procedure TvbGetStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
    end;

  constructor TvbGetStmt.Create( fileNum, rec, varName : TvbExpr );
    begin
      inherited Create( GET_STMT );
      fFileNum := fileNum;
      fRecNum  := rec;
      fVar := varName;
    end;

  destructor TvbGetStmt.Destroy;
    begin
      FreeAndNil( fFileNum );
      FreeAndNil( fRecNum );
      FreeAndNil( fVar );
      inherited;
    end;

  { TvbLineInputStmt }

  procedure TvbLineInputStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnLineInputStmt( self, Context );
    end;

  constructor TvbLineInputStmt.Create(fileNum, varName: TvbExpr);
    begin
      inherited Create( LINE_INPUT_STMT );
      fFileNum := fileNum;
      fVar     := varName;
    end;

  destructor TvbLineInputStmt.Destroy;
    begin
      FreeAndNil( fFileNum );
      FreeAndNil( fVar );
      inherited;
    end;

  { TvbLockStmt }

  procedure TvbLockStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnLockStmt( self, Context );
    end;

  constructor TvbLockStmt.Create( fileNum, startRec, endRec : TvbExpr );
    begin
      inherited Create( LOCK_STMT );
      fFileNum  := fileNum;
      fStartRec := startRec;
      fEndRec   := endRec;
    end;

  destructor TvbLockStmt.Destroy;
    begin
      FreeAndNil( fFileNum );
      FreeAndNil( fStartRec );
      FreeAndNil( fEndRec );
      inherited;
    end;

  { TvbUnlockStmt }

  procedure TvbUnlockStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnUnlockStmt( self, Context );
    end;

  constructor TvbUnlockStmt.Create( fileNum, startRec, endRec : TvbExpr );
    begin
      inherited Create( UNLOCK_STMT );
      fFileNum := fileNum;
      fStartRec := startRec;
      fEndRec := endRec;
    end;

  destructor TvbUnlockStmt.Destroy;
    begin
      FreeAndNil( fFileNum );
      FreeAndNil( fStartRec );
      FreeAndNil( fEndRec );
      inherited;
    end;

  { TvbPutStmt }

  procedure TvbPutStmt.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnPutStmt( self, Context );
    end;

  constructor TvbPutStmt.Create( fileNum, rec, varName : TvbExpr );
    begin
      inherited Create( PUT_STMT );
      fFileNum := fileNum;
      fRecNum  := rec;
      fVar := varName;
    end;

  destructor TvbPutStmt.Destroy;
    begin
      FreeAndNil( fFileNum );
      FreeAndNil( fRecNum );
      FreeAndNil( fVar );
      inherited;
    end;

  { TvbSeekStmt }

  procedure TvbSeekStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnSeekStmt( self, Context );
    end;

  constructor TvbSeekStmt.Create( fileNum, aPos : TvbExpr );
    begin
      inherited Create( SEEK_STMT );
      fFileNum := fileNum;
      fPos     := aPos;
    end;

  destructor TvbSeekStmt.Destroy;
    begin
      FreeAndNil( fFileNum );
      FreeAndNil( fPos );
      inherited;
    end;

  { TvbWidthStmt }

  procedure TvbWidthStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnWidthStmt( self, Context );
    end;

  constructor TvbWidthStmt.Create( fileNum, width : TvbExpr );
    begin
      inherited Create( WIDTH_STMT );
      fFileNum := fileNum;
      fRecWidth := width;
    end;

  destructor TvbWidthStmt.Destroy;
    begin
      FreeAndNil( fFileNum );
      FreeAndNil( fRecWidth );
      inherited;
    end;

  { TvbWriteStmt }

  procedure TvbWriteStmt.AddValue( val : TvbExpr );
    begin
      if fValues = nil
        then fValues := TvbExprList.Create;
      fValues.Add( val );
    end;

  constructor TvbWriteStmt.Create( fileNum : TvbExpr; values : TvbExprList );
    begin
      inherited Create( WRITE_STMT );
      fFileNum := fileNum;
      fValues  := values;
    end;

  destructor TvbWriteStmt.Destroy;
    begin
      FreeAndNil( fFileNum );
      FreeAndNil( fValues );
      inherited;
    end;

  function TvbWriteStmt.GetValueCount : integer;
    begin
      if fValues = nil
        then result := 0
        else result := fValues.Count;
    end;

  function TvbWriteStmt.GetValue( i : integer ) : TvbExpr;
    begin
      if fValues = nil
        then result := nil
        else result := TvbExpr( fValues.Nodes[i] );
    end;

  procedure TvbWriteStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnWriteStmt( self, Context );
    end;

  { TvbErrorStmt }

  procedure TvbErrorStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnErrorStmt( self, Context );
    end;

  constructor TvbErrorStmt.Create( err : TvbExpr );
    begin
      inherited Create( ERROR_STMT );
      fErrorNum := err;
    end;

  destructor TvbErrorStmt.Destroy;
    begin
      FreeAndNil( fErrorNum );
      inherited;
    end;

  { TvbTimeStmt }

  procedure TvbTimeStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnTimeStmt( self, Context );
    end;

  constructor TvbTimeStmt.Create( newTime : TvbExpr );
    begin
      inherited Create( TIME_STMT );
      fNewTime := newTime;
    end;

  destructor TvbTimeStmt.Destroy;
    begin
      FreeAndNil( fNewTime );
      inherited;
    end;

  { TvbStopStmt }

  procedure TvbStopStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnStopStmt( self, Context );
    end;

  constructor TvbStopStmt.Create;
    begin
      inherited Create( STOP_STMT );
    end;

  { TvbDateStmt }

  procedure TvbDateStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnDateStmt( self, Context );
    end;

  constructor TvbDateStmt.Create( newDate : TvbExpr );
    begin
      inherited Create( DATE_STMT );
      fNewDate := newDate;
    end;

  destructor TvbDateStmt.Destroy;
    begin
      FreeAndNil( fNewDate );
      inherited;
    end;

  { TvbNameStmt }

  procedure TvbNameStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnNameStmt( self, Context );
    end;

  constructor TvbNameStmt.Create( oldPath, newPath : TvbExpr );
    begin
      inherited Create( NAME_STMT );
      fOldPath := oldPath;
      fNewPath := newPath;
    end;

  destructor TvbNameStmt.Destroy;
    begin
      FreeAndNil( fOldPath );
      FreeAndNil( fNewPath );
      inherited;
    end;

  { TvbInputStmt }

  procedure TvbInputStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnInputStmt( self, Context );
    end;

  procedure TvbInputStmt.AddVar( v : TvbExpr );
    begin
      if fVars = nil
        then fVars := TvbExprList.Create;
      fVars.Add( v );
    end;

  constructor TvbInputStmt.Create( fileNum : TvbExpr; aVarNames : TvbExprList );
    begin
      inherited Create( INPUT_STMT );
      fFileNum := fileNum;
      fVars := aVarNames;
    end;

  destructor TvbInputStmt.Destroy;
    begin
      FreeAndNil( fFileNum );
      FreeAndNil( fVars );
      inherited;
    end;

  function TvbInputStmt.GetVar( i : integer ) : TvbExpr;
    begin
      if fVars = nil
        then result := nil
        else result := TvbExpr( fVars.Nodes[i] );
    end;

  function TvbInputStmt.GetVarCount : integer;
    begin
      if fVars = nil
        then result := 0
        else result := fVars.Count;
    end;

  { TvbDebugAssertStmt }

  procedure TvbDebugAssertStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnDebugAssertStmt( self, Context );
    end;

  constructor TvbDebugAssertStmt.Create( cond : TvbExpr );
    begin
      inherited Create( DEBUG_ASSERT_STMT );
      fBoolExpr := cond;
    end;

  destructor TvbDebugAssertStmt.Destroy;
    begin
      FreeAndNil( fBoolExpr );
      inherited;
    end;

  { TvbDebugPrintStmt }

  procedure TvbDebugPrintStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnDebugPrintStmt( self, Context );
    end;

  constructor TvbDebugPrintStmt.Create;
    begin
      inherited Create( DEBUG_PRINT_STMT );
    end;

  { TvbResumeStmt }

  procedure TvbResumeStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnResumeStmt( self, Context );
    end;

  constructor TvbResumeStmt.Create( kind : TvbResumeKind; labelName : TvbSimpleName );
    begin
      inherited Create( RESUME_STMT );
      fKind  := kind;
      fLabel := labelName;
    end;

  destructor TvbResumeStmt.Destroy;
    begin
      FreeAndNil( fLabel );
      inherited;
    end;

  { TvbPrintStmtExpr }

  procedure TvbPrintStmtExpr.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnPrintArgExpr( self, Context );
    end;

  constructor TvbPrintStmtExpr.Create( anArgKind : TvbPrintArgKind; expr : TvbExpr );
    begin
      inherited Create( PRINT_ARG_EXPR );
      fArgKind := anArgKind;
      fOp1 := expr;
    end;

  destructor TvbPrintStmtExpr.Destroy;
    begin
      inherited;
    end;

  { TvbAbstractPrintStmt }

  destructor TvbAbstractPrintStmt.Destroy;
    begin
      FreeAndNil( fExprs );
      inherited;
    end;

  function TvbAbstractPrintStmt.GetExprCount : integer;
    begin
      if fExprs = nil
        then result := 0
        else result := fExprs.Count;
    end;

  function TvbAbstractPrintStmt.GetExpr( i : integer ) : TvbPrintStmtExpr;
    begin
      if fExprs = nil
        then result := nil
        else
          begin
            assert( fExprs.Nodes[i].NodeKind = PRINT_ARG_EXPR );
            result := TvbPrintStmtExpr( fExprs.Nodes[i] );
          end;
    end;

  function TvbAbstractPrintStmt.AddExpr( kind : TvbPrintArgKind ) : TvbPrintStmtExpr;
    begin
      if fExprs = nil
        then fExprs := TvbNodeList.Create;
      result := TvbPrintStmtExpr.Create( kind );
      try
        fExprs.Add( result )
      except
        FreeAndNil( result );
      end;
    end;

  { TvbCallStmt }

  procedure TvbCallStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnCallStmt( self, Context );
    end;

  constructor TvbCallStmt.Create( aMethod : TvbExpr );
    begin
      inherited Create( CALL_STMT );
      fMethod := aMethod;
    end;

  destructor TvbCallStmt.Destroy;
    begin
      FreeAndNil( fMethod );
      inherited;
    end;

  { TvbDateLit }

  procedure TvbDateLit.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnDateLitExpr( self, Context );
    end;

  constructor TvbDateLit.Create( const val : TDateTime );
    begin
      inherited Create( DATE_LIT_EXPR );
      fValue    := val;
      fNodeType := vbDateType;
    end;

  { TvbEndStmt }

  procedure TvbEndStmt.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnEndStmt( self, Context );
    end;

  constructor TvbEndStmt.Create;
    begin
      inherited Create( END_STMT );
    end;

  { TvbUBoundExpr }

  procedure TvbUBoundExpr.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnUboundExpr( self, Context );
    end;

  constructor TvbUBoundExpr.Create( arr, dim : TvbExpr );
    begin
      inherited Create( UBOUND_EXPR, arr, dim );
      fNodeType := vbLongType;
    end;

  { TvbLBoundExpr }

  procedure TvbLBoundExpr.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnLBoundExpr( self, Context );
    end;

  constructor TvbLBoundExpr.Create( arr, dim : TvbExpr );
    begin
      inherited Create( LBOUND_EXPR, arr, dim );
      fNodeType := vbLongType;
    end;

  { TvbExpr }

  constructor TvbExpr.Create( nk : TvbNodeKind; op1 : TvbExpr );
    begin
      inherited Create( nk );
      fOp1 := op1;
    end;

  procedure TvbExpr.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
    end;

  constructor TvbExpr.Create( nk : TvbNodeKind; op1, op2 : TvbExpr );
    begin
      inherited Create( nk );
      fOp1 := op1;
      fOp2 := op2;
    end;

  destructor TvbExpr.Destroy;
    begin
      FreeAndNil( fOp1 );
      FreeAndNil( fOp2 );
      inherited;
    end;

  function TvbExpr.GetIsRValue : boolean;
    begin
      result := assign_LHS in fFlags;
    end;

  procedure TvbExpr.SetIsRValue( const value : boolean );
    begin
      Include( fFlags, assign_LHS );
    end;

  function TvbExpr.GetIsUsedWithArgs : boolean;
    begin
      result := expr_UsedWithArgs in fFlags;
    end;

  procedure TvbExpr.SetIsUsedWithArgs( const Value : boolean );
    begin
      Include( fFlags, expr_UsedWithArgs );
    end;

  function TvbExpr.GetIsUsedWithEmptyParens : boolean;
    begin
      result := expr_MustBeFunc in fFlags;
    end;

  procedure TvbExpr.SetIsUsedWithEmptyParens( const Value : boolean );
    begin
      Include( fFlags, expr_MustBeFunc );
    end;

  { TvbProject }

  procedure TvbProject.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnProjectDef( self, Context );
    end;

  function TvbProject.AddImportedModule( m : TvbStdModuleDef ) : TvbStdModuleDef;
    begin
      result := m;
      fImportedModules.Add( m );
      m.Scope.fParent := fScope;
      fScope.Bind( m, PROJECT_OR_MODULE_BINDING );
    end;

  function TvbProject.AddModule( m : TvbModuleDef ) : TvbModuleDef;
    begin
      result := m;
      m.Scope.fParent := fScope;
      fModules.Add( m );
      fScope.Bind( m, PROJECT_OR_MODULE_BINDING );
    end;

  constructor TvbProject.Create;
    begin
      inherited Create( PROJECT_DEF );
      fKind            := vbpEXE;
      fModules         := TvbNodeList.Create;
      fImportedModules := TvbNodeList.Create( false ); // doesn't own its objects
      fScope           := TvbScope.Create;
    end;

  destructor TvbProject.Destroy;
    begin
      FreeAndNil( fModules );
      FreeAndNil( fImportedModules );
      FreeAndNil( fScope );
      inherited;
    end;

  function TvbProject.GetImportedModule( i : integer ) : TvbStdModuleDef;
    begin
      result := TvbStdModuleDef( fImportedModules[i] );
    end;

  function TvbProject.GetImportedModuleCount : integer;
    begin
      result := fImportedModules.Count;
    end;

  function TvbProject.GetModule( i : integer ) : TvbModuleDef;
    begin
      result := TvbModuleDef( fModules[i] );
    end;

  function TvbProject.GetModuleCount : integer;
    begin
      result := fModules.Count;
    end;

  { TvbScope }

  constructor TvbScope.Create( parent : TvbScope{$IFDEF DEBUG}; const name : string {$ENDIF} );
    begin
      fParent   := parent;
      fTraits   := TCaseInsensitiveTraits.Create;
      fBindings := THashList.Create( fTraits, 100 );
      {$IFDEF DEBUG}
      fName := name;
      {$ENDIF}
    end;

  destructor TvbScope.Destroy;
    begin
      fBindings.Iterate( nil, Iterate_FreeObjects);
      FreeAndNil( fBindings );
      FreeAndNil( fTraits );
      inherited;
    end;

  function TvbScope.FindBinding( const name : string ) : TvbNameBindings;
    begin
      result := TvbNameBindings( fBindings[name] );
      if result = nil
        then
          begin
            result := TvbNameBindings.Create;
            fBindings.Add( name, result );
          end;
    end;

  function TvbScope.Lookup( kind : TvbNameBindingKind; const name : string; var decl : TvbDef; searchParents: boolean ) : boolean;
    var
      b : TvbNameBindings;
    begin
      result := false;
      decl := nil;
      // first search locally in this scope.
      if fHaveBinding[kind]
        then
          begin
            b := TvbNameBindings( fBindings[name] );
            if b <> nil
              then decl := b.Bindings[kind];
          end;
      // if name was not found locally, search locally in the base scope
      // then in parent scopes if indicated.
      if decl = nil
        then
          begin
            if fBaseScope <> nil
              then result := fBaseScope.Lookup( kind, name, decl, false );
            if not result and searchParents and ( fParent <> nil )
              then result := fParent.Lookup( kind, name, decl, true );
          end
        else result := true;
      {$IFDEF DEBUG}
      if result = true
        then assert( decl <> nil );
      {$ENDIF}
    end;

  function TvbScope.Bind( decl : TvbDef; kind : TvbNameBindingKind ) : boolean;
    begin
      result := true;
      with FindBinding( decl.Name ) do
        if Bindings[kind] = nil
          then
            begin
              Bindings[kind]     := decl;
              fHaveBinding[kind] := true;
            end
          else result := false;
    end;

  procedure TvbScope.BindMembers( cls : TvbClassDef );
    var
      i : integer;
    begin
      with cls do
        begin
          for i := 0 to VarCount - 1 do
            Bind( Vars[i], OBJECT_BINDING );
          for i := 0 to MethodCount - 1 do
            Bind( Method[i], OBJECT_BINDING );
          for i := 0 to PropCount - 1 do
            Bind( Prop[i], OBJECT_BINDING );
          for i := 0 to EventCount - 1 do
            Bind( Event[i], EVENT_BINDING );
        end;
    end;

  function TvbScope.LookupObject( const name    : string;
                                  var decl      : TvbDef;
                                  flags         : TvbLookupObjectFlags;
                                  searchParents : boolean ) : boolean;
    var
      funcDef : TvbAbstractFuncDef absolute decl;
      varDef  : TvbVarDef absolute decl;
      propDef : TvbPropertyDef absolute decl;
      b       : TvbNameBindings;
    begin
      result := false;
      decl := nil;

      if fHaveBinding[OBJECT_BINDING]
        then
          begin
            b := TvbNameBindings( fBindings[name] );
            if b <> nil   // name defined in the scope?
              then
                begin
                  // check if we have an OBJECT_BINDING with the requested name.
                  decl := b.Bindings[OBJECT_BINDING];
                  if decl <> nil
                    then
                      // check if the decl matches the requested criteria.
                      case decl.NodeKind of
                        DLL_FUNC_DEF,
                        FUNC_DEF :
                          begin
                            if ( lofMustTakeArgs in flags ) and
                               ( funcDef.Argc = 0 )
                              then decl := nil
                              else
                            // functions can't be rvalues.
                            if lofMustBeRValue in flags
                              then decl := nil;
                          end;
                        PARAM_DEF,
                        VAR_DEF :
                          if lofMustBeFunc in flags
                            then decl := nil
                            else
                          // if the caller requested parameters and the var isn't
                          // an array, the declaration is no good.
                          // vars are always rvalues.
                          if ( lofMustTakeArgs in flags ) and
                             ( varDef.VarType <> nil ) and
                             ( varDef.VarType.Code and ARRAY_TYPE = 0 )
                            then decl := nil;
                        PROPERTY_DEF :
                          begin
                            // if the caller requested an rvalue this property must
                            // have a setter or letter function to be good.
                            if lofMustBeRValue in flags
                              then
                                begin
                                  if propDef.SetFn <> nil
                                    then decl := propDef.SetFn
                                    else decl := propDef.LetFn;
                                end
                              else decl := propDef.GetFn;

                            // note that if at this point decl <> NIL
                            // it's a function(setter, letter or getter).
                            if ( decl <> nil ) and
                               ( lofMustTakeArgs in flags ) and
                               ( funcDef.Argc = 0 )
                              then decl := nil;
                          end;
                        CONST_DEF :
                          begin
                            // a constant doesn't satisfy any of the flags,
                            // so they are never valid in those contexts.
                            decl := nil
                          end;
                        // else assert( flags = [] );
                      end;
                end;
          end;

      // if name was not found locally, search locally in the base scope
      // then in parent scopes if indicated.
      if decl = nil
        then
          begin
            if fBaseScope <> nil
              then result := fBaseScope.LookupObject( name, decl, flags, false );
            if not result and searchParents and ( fParent <> nil )
              then result := fParent.LookupObject( name, decl, flags, true );
          end
        else result := true;
      {$IFDEF DEBUG}
      if result = true
        then assert( decl <> nil );
      {$ENDIF}
    end;

  { TvbCastExpr }

  procedure TvbCastExpr.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnCastExpr( self, Context );
    end;

  constructor TvbCastExpr.Create( expr : TvbExpr; conv : TvbCastTo );
    begin
      inherited Create( CAST_EXPR, expr );
      fConvertTo := conv;
    end;

  { TvbMidExpr }

  procedure TvbMidExpr.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnMidExpr( self, Context );
    end;

  constructor TvbMidExpr.Create( str, start, len : TvbExpr );
    begin
      inherited Create( MID_EXPR );
      fOp1      := str;
      fOp2      := start;
      fLen      := len;
      fNodeType := vbStringType;
    end;

  destructor TvbMidExpr.Destroy;
    begin
      FreeAndNil( fLen );
      inherited;
    end;

  { TvbStringExpr }

  procedure TvbStringExpr.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnStringExpr( self, Context );
    end;

  constructor TvbStringExpr.Create( count, ch : TvbExpr );
    begin
      inherited Create( STRING_EXPR );
      fOp1      := count;
      fOp2      := ch;
      fNodeType := vbStringType;
    end;

  { TvbIntExpr }

  procedure TvbIntExpr.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnIntExpr( self, Context );
    end;

  constructor TvbIntExpr.Create( num : TvbExpr );
    begin
      inherited Create( INT_EXPR );
      fOp1 := num;
      fNodeType := vbLongType;
    end;

  { TvbFixExpr }

  procedure TvbFixExpr.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnFixExpr( self, Context );
    end;

  constructor TvbFixExpr.Create( num : TvbExpr );
    begin
      inherited Create( FIX_EXPR );
      fOp1 := num;
      fNodeType := vbLongType;
    end;

  { TvbAbsExpr }

  procedure TvbAbsExpr.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnAbsExpr( self, context );
    end;

  constructor TvbAbsExpr.Create( num : TvbExpr );
    begin
      inherited Create( ABS_EXPR );
      fOp1 := num;
    end;

  { TvbLenExpr }

  procedure TvbLenExpr.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnLenExpr( self, Context );
    end;

  constructor TvbLenExpr.Create( strOrVar : TvbExpr );
    begin
      inherited Create( LEN_EXPR );
      fOp1      := strOrVar;
      fNodeType := vbLongType;
    end;

  { TvbLenBExpr }

  procedure TvbLenBExpr.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnLenBExpr( self, Context );
    end;

  constructor TvbLenBExpr.Create( strOrVar : TvbExpr );
    begin
      inherited Create( LENB_EXPR );
      fOp1      := strOrVar;
      fNodeType := vbLongType;
    end;

  { TvbSgnExpr }

  procedure TvbSgnExpr.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnSgnExpr( self, Context );
    end;

  constructor TvbSgnExpr.Create( num : TvbExpr );
    begin
      inherited Create( SGN_EXPR );
      fOp1      := num;
      fNodeType := vbIntegerType;
    end;

  { TvbSeekExpr }

  procedure TvbSeekExpr.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnSeekExpr( self, Context );
    end;

  constructor TvbSeekExpr.Create( fileNum : TvbExpr );
    begin
      inherited Create( SEEK_EXPR );
      fOp1      := fileNum;
      fNodeType := vbLongType;
    end;

  { TvbArrayExpr }

  procedure TvbArrayExpr.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnArrayExpr( self, Context );
    end;

  procedure TvbArrayExpr.AddValue( val : TvbExpr );
    begin
      fValues.Add( val );
    end;

  constructor TvbArrayExpr.Create;
    begin
      inherited Create( ARRAY_EXPR );
      fValues := TvbExprList.Create( true );
    end;

  destructor TvbArrayExpr.Destroy;
    begin
      FreeAndNil( fValues );
      inherited;
    end;

  function TvbArrayExpr.GetValue( i : integer ) : TvbExpr;
    begin
      result := TvbExpr( fValues[i] );
    end;

  function TvbArrayExpr.GetValueCount : integer;
    begin
      result := fValues.Count;
    end;

  { TvbStmtBlock }

  procedure TvbStmtBlock.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnStmtBlock( self, Context );
    end;

  function TvbStmtBlock.Add( stmt : TvbStmt ) : TvbStmt;
    begin
      assert( stmt <> nil );
      assert( stmt.PrevStmt = nil );
      if fFirstStmt = nil
        then
          begin
            fFirstStmt := stmt;
            fLastStmt  := stmt;
          end
        else
          begin
            stmt.PrevStmt      := fLastStmt;
            fLastStmt.NextStmt := stmt;
            fLastStmt          := stmt;
          end;
      result := stmt;
    end;

  function TvbStmtBlock.AddCallStmt: TvbCallStmt;
    begin
      result := TvbCallStmt( Add( TvbCallStmt.Create ) );
    end;

  function TvbStmtBlock.AddCircleStmt: TvbCircleStmt;
    begin
      result := TvbCircleStmt( Add( TvbCircleStmt.Create ) );
    end;

  function TvbStmtBlock.AddCloseStmt: TvbCloseStmt;
    begin
      result := TvbCloseStmt( Add( TvbCloseStmt.Create ) );
    end;

  function TvbStmtBlock.AddDateStmt : TvbDateStmt;
    begin
      result := TvbDateStmt( Add( TvbDateStmt.Create ) );
    end;

  function TvbStmtBlock.AddDebugAssertStmt : TvbDebugAssertStmt;
    begin
      result := TvbDebugAssertStmt( Add( TvbDebugAssertStmt.Create ) );
    end;

  function TvbStmtBlock.AddDebugPrintStmt : TvbDebugPrintStmt;
    begin
      result := TvbDebugPrintStmt( Add( TvbDebugPrintStmt.Create ) );
    end;

  function TvbStmtBlock.AddDoLoopStmt: TvbDoLoopStmt;
    begin
      result := TvbDoLoopStmt( Add( TvbDoLoopStmt.Create ) );
    end;

  function TvbStmtBlock.AddEraseStmt: TvbEraseStmt;
    begin
      result := TvbEraseStmt( Add( TvbEraseStmt.Create ) );
    end;

  function TvbStmtBlock.AddErrorStmt : TvbErrorStmt;
    begin
      result := TvbErrorStmt( Add( TvbErrorStmt.Create ) );
    end;

  function TvbStmtBlock.AddForeachStmt: TvbForeachStmt;
    begin
      result := TvbForeachStmt( Add( TvbForeachStmt.Create ) );
    end;

  function TvbStmtBlock.AddForStmt: TvbForStmt;
    begin
      result := TvbForStmt( Add( TvbForStmt.Create ) );
    end;

  function TvbStmtBlock.AddGetStmt: TvbGetStmt;
    begin
      result := TvbGetStmt( Add( TvbGetStmt.Create ) );
    end;

  function TvbStmtBlock.AddGotoOrGosubStmt: TvbGotoOrGosubStmt;
    begin
      result := TvbGotoOrGosubStmt( Add( TvbGotoOrGosubStmt.Create ) );
    end;

  function TvbStmtBlock.AddIfStmt : TvbIfStmt;
    begin
      result := TvbIfStmt( Add( TvbIfStmt.Create ) );
    end;

  function TvbStmtBlock.AddInputStmt: TvbInputStmt;
    begin
      result := TvbInputStmt( Add( TvbInputStmt.Create ) );
    end;

  function TvbStmtBlock.AddLineInputStmt : TvbLineInputStmt;
    begin
      result := TvbLineInputStmt( Add( TvbLineInputStmt.Create ) );
    end;

  function TvbStmtBlock.AddLineStmt : TvbLineStmt;
    begin
      result := TvbLineStmt( Add( TvbLineStmt.Create ) );
    end;

  function TvbStmtBlock.AddLockStmt: TvbLockStmt;
    begin
      result := TvbLockStmt( Add( TvbLockStmt.Create ) );
    end;

  function TvbStmtBlock.AddMidAssignStmt : TvbMidAssignStmt;
    begin
      result := TvbMidAssignStmt( Add( TvbMidAssignStmt.Create ) );
    end;

  function TvbStmtBlock.AddNameStmt : TvbNameStmt;
    begin
      result := TvbNameStmt( Add( TvbNameStmt.Create ) );
    end;

  function TvbStmtBlock.AddOnErrorStmt: TvbOnErrorStmt;
    begin
      result := TvbOnErrorStmt( Add( TvbOnErrorStmt.Create ) );
    end;

  function TvbStmtBlock.AddOnGotoOrGosubStmt: TvbOnGotoOrGosubStmt;
    begin
      result := TvbOnGotoOrGosubStmt( Add( TvbOnGotoOrGosubStmt.Create ) );
    end;

  function TvbStmtBlock.AddOpenStmt : TvbOpenStmt;
    begin
      result := TvbOpenStmt( Add( TvbOpenStmt.Create ) );
    end;

  function TvbStmtBlock.AddPrintStmt : TvbPrintStmt;
    begin
      result := TvbPrintStmt( Add( TvbPrintStmt.Create ) );
    end;

  function TvbStmtBlock.AddPrintMethod : TvbPrintMethod;
    begin
      result := TvbPrintMethod( Add( TvbPrintMethod.Create ) );
    end;

  function TvbStmtBlock.AddPSetStmt : TvbPSetStmt;
    begin
      result := TvbPSetStmt( Add( TvbPSetStmt.Create ) );
    end;

  function TvbStmtBlock.AddPutStmt: TvbPutStmt;
    begin
      result := TvbPutStmt( Add( TvbPutStmt.Create ) );
    end;

  function TvbStmtBlock.AddRaiseEventStmt: TvbRaiseEventStmt;
    begin
      result := TvbRaiseEventStmt( Add( TvbRaiseEventStmt.Create ) );
    end;

  function TvbStmtBlock.AddReDimStmt: TvbReDimStmt;
    begin
      result := TvbReDimStmt( Add( TvbReDimStmt.Create ) );
    end;

  function TvbStmtBlock.AddResumeStmt : TvbResumeStmt;
    begin
      result := TvbResumeStmt( Add( TvbResumeStmt.Create ) );
    end;

  function TvbStmtBlock.AddSeekStmt: TvbSeekStmt;
    begin
      result := TvbSeekStmt( Add( TvbSeekStmt.Create ) );
    end;

  function TvbStmtBlock.AddSelectCaseStmt : TvbSelectCaseStmt;
    begin
      result := TvbSelectCaseStmt( Add( TvbSelectCaseStmt.Create ) );
    end;

  function TvbStmtBlock.AddStopStmt : TvbStopStmt;
    begin
      result := TvbStopStmt( Add( TvbStopStmt.Create ) );
    end;

  function TvbStmtBlock.AddUnlockStmt : TvbUnlockStmt;
    begin
      result := TvbUnlockStmt( Add( TvbUnlockStmt.Create ) );
    end;

  function TvbStmtBlock.AddWhileStmt: TvbDoLoopStmt;
    begin
      result := TvbDoLoopStmt( Add( TvbDoLoopStmt.Create ) );
    end;

  function TvbStmtBlock.AddWidthStmt: TvbWidthStmt;
    begin
      result := TvbWidthStmt( Add( TvbWidthStmt.Create ) );
    end;

  function TvbStmtBlock.AddWithStmt: TvbWithStmt;
    begin
      result := TvbWithStmt( Add( TvbWithStmt.Create ) );
    end;

  function TvbStmtBlock.AddWriteStmt: TvbWriteStmt;
    begin
      result := TvbWriteStmt( Add( TvbWriteStmt.Create ) );
    end;

  constructor TvbStmtBlock.Create;
    begin
      inherited Create( STMT_BLOCK );
    end;

  destructor TvbStmtBlock.Destroy;
    var
      next : TvbStmt;
    begin
      while fFirstStmt <> nil do
        begin
          next := fFirstStmt.NextStmt;
          fFirstStmt.Free;
          fFirstStmt := next;
        end;
      inherited;
    end;

  function TvbStmtBlock.GetCount : integer;
    begin
      result := 0;
    end;

  { TvbName }

  constructor TvbName.Create( nk : TvbNodeKind; const name : string );
    begin
      inherited Create( nk );
      fName  := name;
      fName1 := name;
    end;

  { TvbLabelDef }

  procedure TvbLabelDef.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnLabelDef( self, Context );
    end;

  constructor TvbLabelDef.Create( const name : string );
    begin
      inherited Create( LABEL_DEF, [], name );
    end;

  { TvbDateExpr }

  procedure TvbDateExpr.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnDateExpr( self, Context );
    end;

  constructor TvbDateExpr.Create;
    begin
      inherited Create( DATE_EXPR );
    end;

  { TvbFuncResultExpr }

  procedure TvbFuncResultExpr.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnFuncResultExpr( self, Context );
    end;

  constructor TvbFuncResultExpr.Create( typ : IvbType );
    begin
      inherited Create( FUNC_RESULT_EXPR );
      fNodeType := typ;
    end;

  { TvbInputExpr }

  procedure TvbInputExpr.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnInputExpr( self, Context );
    end;

  constructor TvbInputExpr.Create;
    begin
      inherited Create( INPUT_EXPR );
    end;

  { TvbDoEventsExpr }

  procedure TvbDoEventsExpr.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnDoEventsExpr( self, Context );
    end;

  constructor TvbDoEventsExpr.Create;
    begin
      inherited Create( DOEVENTS_EXPR );
    end;

  { TvbLibModuleDef }

  procedure TvbLibModuleDef.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnLibModuleDef( self, Context );
    end;

  procedure TvbLibModuleDef.AddConstant( cst : TvbConstDef );
    begin
      if fConsts = nil
        then fConsts := TvbNodeList.Create;
      fConsts.Add( cst );
      fScope.Bind( cst, OBJECT_BINDING );
    end;

  procedure TvbLibModuleDef.AddMethod( func : TvbAbstractFuncDef );
    begin
      if fFuncs = nil
        then fFuncs := TvbNodeList.Create;
      fFuncs.Add( func );
      fScope.Bind( func, OBJECT_BINDING );        
    end;

  procedure TvbLibModuleDef.AddVar( vr : TvbVarDef );
    begin
      if fVars = nil
        then fVars := TvbNodeList.Create;
      fVars.Add( vr );
      fScope.Bind( vr, OBJECT_BINDING );
    end;

  constructor TvbLibModuleDef.Create( const flags : TvbNodeFlags;
                                          const name  : string;
                                          const name1 : string );
    begin
      inherited Create( LIB_MODULE_DEF, flags, name, name1 );
      fScope := TvbScope.Create;
    end;

  destructor TvbLibModuleDef.Destroy;
    begin
      FreeAndNil( fScope );
      FreeAndNil( fVars );
      FreeAndNil( fConsts );
      FreeAndNil( fFuncs );
      FreeAndNil( fProps );
      inherited;
    end;

  function TvbLibModuleDef.GetConst( i : integer ) : TvbConstDef;
    begin
      assert( fConsts <> nil );
      result := TvbConstDef( fConsts[i] );
    end;

  function TvbLibModuleDef.GetConstCount : integer;
    begin
      if fConsts = nil
        then result := 0
        else result := fConsts.Count;
    end;

  function TvbLibModuleDef.GetFunction( i : integer ) : TvbAbstractFuncDef;
    begin
      assert( fFuncs <> nil );
      result := TvbAbstractFuncDef( fFuncs[i] );
    end;

  function TvbLibModuleDef.GetFunctionCount : integer;
    begin
      if fFuncs = nil
        then result := 0
        else result := fFuncs.Count;
    end;

  function TvbLibModuleDef.GetVarCount : integer;
    begin
      if fVars = nil
        then result := 0
        else result := fVars.Count;
    end;

  function TvbLibModuleDef.GetVar( i : integer ) : TvbVarDef;
    begin
      assert( fVars <> nil );
      result := TvbVarDef( fVars[i] );
    end;

  function TvbLibModuleDef.GetProperty( i : integer ) : TvbPropertyDef;
    begin
      assert( fProps <> nil );
      result := TvbPropertyDef( fProps[i] );
    end;

  function TvbLibModuleDef.GetPropCount : integer;
    begin
      if fProps = nil
        then result := 0
        else result := fProps.Count;
    end;

  procedure TvbLibModuleDef.AddProperty( p : TvbPropertyDef );
    begin
      if fProps = nil
        then fProps := TvbNodeList.Create;
      fProps.Add( p );
    end;

  { TvbClassDef }

  procedure TvbClassDef.Accept( Visitor: TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnClassDef( self, Context );
    end;

  procedure TvbClassDef.AddEvent( e : TvbEventDef );
    begin
      fEvents.Add( e );
      fScope.Bind( e, EVENT_BINDING );
    end;

  procedure TvbClassDef.AddMethod( m : TvbAbstractFuncDef );
    begin
      fMethods.Add( m );
      // we make this check because external functions have no scope.
      if m.Scope <> nil
        then m.Scope.fParent := fScope;
      fScope.Bind( m, OBJECT_BINDING );
    end;

  procedure TvbClassDef.AddProperty( p : TvbPropertyDef );
    begin
      fProps.Add( p );
//      // is the property has got setter/getter functions set their parent scopes
//      // to the class scope.
//      if p.SetFn <> nil
//        then p.SetFn.Scope.Parent := fScope;
//      if p.GetFn <> nil
//        then p.GetFn.Scope.Parent := fScope;
      fScope.Bind( p, OBJECT_BINDING );
    end;

  procedure TvbClassDef.AddVar( v : TvbVarDef );
    begin
      fVars.Add( v );
      fScope.Bind( v, OBJECT_BINDING );
    end;

  constructor TvbClassDef.Create( const flags     : TvbNodeFlags;
                                  const name      : string;
                                  const name1     : string;
                                  const baseClass : IvbType );

    begin
      inherited Create( CLASS_DEF, flags, name, name1 );
      fVars      := TvbNodeList.Create;
      fEvents    := TvbNodeList.Create;
      fMethods   := TvbNodeList.Create;
      fProps     := TvbNodeList.Create;
      fMeType    := TvbType.CreateUDT( self );
      fBaseClass := baseClass;
    end;

  destructor TvbClassDef.Destroy;
    begin
      FreeAndNil( fVars );
      FreeAndNil( fEvents );
      FreeAndNil( fMethods );
      FreeAndNil( fProps );
      FreeAndNil( fClassInitialize );
      FreeAndNil( fClassTerminate );
      inherited;
    end;

  function TvbClassDef.GetEventCount : integer;
    begin
      result := fEvents.Count;
    end;

  function TvbClassDef.GetEvents( i : integer ) : TvbEventDef;
    begin
      result := TvbEventDef( fEvents[i] );
    end;

  function TvbClassDef.GetMethod( i : integer) : TvbAbstractFuncDef;
    begin
      result := TvbAbstractFuncDef( fMethods[i] );
    end;

  function TvbClassDef.GetMethodCount : integer;
    begin
      result := fMethods.Count;
    end;

  function TvbClassDef.GetProperty( i : integer ) : TvbPropertyDef;
    begin
      result := TvbPropertyDef( fProps[i] );
    end;

  function TvbClassDef.GetPropCount : integer;
    begin
      result := fProps.Count;
    end;

  function TvbClassDef.GetVarCount : integer;
    begin
      result := fVars.Count;
    end;

  function TvbClassDef.GetVar( i : integer ) : TvbVarDef;
    begin
      result := TvbVarDef( fVars[i] );
    end;

  { TvbModuleDef }

  function TvbModuleDef.AddClass( c : TvbClassDef ) : TvbClassDef;
    begin
      result := c;
      assert( c.NodeKind = CLASS_DEF );
      if fClasses = nil
        then fClasses := TvbNodeList.Create;
      c.Scope.fParent := fScope;
      fClasses.Add( c );
      fScope.Bind( c, TYPE_BINDING );
    end;

  function TvbModuleDef.AddConstant( cst : TvbConstDef ) : TvbConstDef;
    begin
      result := cst;
      assert( cst.NodeKind = CONST_DEF );
      if fConsts = nil
        then fConsts := TvbNodeList.Create;
      fConsts.Add( cst );
      fScope.Bind( cst, OBJECT_BINDING );
    end;

  function TvbModuleDef.AddDllFunc( df : TvbDllFuncDef ) : TvbDllFuncDef;
    begin
      result := df;
      if fDllFuncs = nil
        then fDllFuncs := TvbNodeList.Create;
      fDllFuncs.Add( df );
      fScope.Bind( df, OBJECT_BINDING );
    end;

  function TvbModuleDef.AddEnum( en : TvbEnumDef ) : TvbEnumDef;
    begin
      result := en;
      if fEnums = nil
        then fEnums := TvbNodeList.Create;
      en.Scope.fParent := fScope;
      fEnums.Add( en );
      fScope.Bind( en, TYPE_BINDING );
    end;

  function TvbModuleDef.AddRecord( s : TvbRecordDef ) : TvbRecordDef;
    begin
      result := s;
      if fRecords = nil
        then fRecords := TvbNodeList.Create;
      s.Scope.fParent := fScope;
      fRecords.Add( s );
      fScope.Bind( s, TYPE_BINDING );
    end;

  constructor TvbModuleDef.Create( nk : TvbNodeKind; const name, relativePath : string; ParentScope : TvbScope );
    begin
      inherited Create( nk, [], name );
      fRelPath := relativePath;
      fScope := TvbScope.Create( ParentScope );
      {$IFDEF DEBUG}
      fScope.Name := fName + '(M)';
      {$ENDIF}
    end;

  destructor TvbModuleDef.Destroy;
    begin
      FreeAndNil( fScope );
      FreeAndNil( fConsts );
      FreeAndNil( fEnums );
      FreeAndNil( fRecords );
      FreeAndNil( fDllFuncs );
      FreeAndNil( fClasses );
      FreeAndNil( fLibModules );
      inherited;
    end;

  function TvbModuleDef.GetClass( i : integer ) : TvbClassDef;
    begin
      result := TvbClassDef( fClasses[i] );
    end;

  function TvbModuleDef.GetClassCount : integer;
    begin
      if fClasses = nil
        then result := 0
        else result := fClasses.Count;
    end;

  function TvbModuleDef.GetConst( i : integer ) : TvbConstDef;
    begin
      result := TvbConstDef( fConsts[i] );
    end;

  function TvbModuleDef.GetConstCount : integer;
    begin
      if fConsts = nil
        then result := 0
        else result := fConsts.Count;
    end;

  function TvbModuleDef.GetDllFunc( i : integer ) : TvbDllFuncDef;
    begin
      result := TvbDllFuncDef( fDllFuncs[i] );
    end;

  function TvbModuleDef.GetDllFuncCount : integer;
    begin
      if fDllFuncs = nil
        then result := 0
        else result := fDllFuncs.Count;
    end;

  function TvbModuleDef.GetEnum( i : integer ) : TvbEnumDef;
    begin
      result := TvbEnumDef( fEnums[i] );
    end;

  function TvbModuleDef.GetEnumCount : integer;
    begin
      if fEnums = nil
        then result := 0
        else result := fEnums.Count;
    end;

  function TvbModuleDef.GetLibModule( i : integer ) : TvbLibModuleDef;
    begin
      result := TvbLibModuleDef( fLibModules[i] );
    end;

  function TvbModuleDef.GetLibModuleCount : integer;
    begin
      if fLibModules = nil
        then result := 0
        else result := fLibModules.Count;        
    end;

  function TvbModuleDef.GetRecord( i : integer ) : TvbRecordDef;
    begin
      result := TvbRecordDef( fRecords[i] );
    end;

  function TvbModuleDef.GetRecordCount : integer;
    begin
      if fRecords = nil
        then result := 0
        else result := fRecords.Count;
    end;

  function TvbModuleDef.AddLibModule( lm : TvbLibModuleDef ) : TvbLibModuleDef;
    begin
      result := lm;
      if fLibModules = nil
        then fLibModules := TvbNodeList.Create;
      lm.Scope.fParent := fScope;
      fLibModules.Add( lm );
      fScope.Bind( lm, PROJECT_OR_MODULE_BINDING );
    end;

  { TvbStdModuleDef }

  procedure TvbStdModuleDef.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnStdModuleDef( self, Context );
    end;

  constructor TvbStdModuleDef.Create( const name, relativePath : string; ParentScope : TvbScope );
    begin
      inherited Create( STD_MODULE_DEF, Name, relativePath, ParentScope );
    end;

  function TvbStdModuleDef.AddMethod( m : TvbAbstractFuncDef ) : TvbAbstractFuncDef;
    begin
      result := m;
      if fMethods = nil
        then fMethods := TvbNodeList.Create;
      m.Scope.fParent := fScope;
      fMethods.Add( m );
      fScope.Bind( m, OBJECT_BINDING );
    end;

  destructor TvbStdModuleDef.Destroy;
    begin
      FreeAndNil( fProps );
      FreeAndNil( fVars );
      FreeAndNil( fMethods );
      inherited;
    end;

  function TvbStdModuleDef.GetProperty( i : integer ) : TvbPropertyDef;
    begin
      result := TvbPropertyDef( fProps[i] );
    end;

  function TvbStdModuleDef.GetPropCount : integer;
    begin
      if fProps = nil
        then result := 0
        else result := fProps.Count;
    end;

  function TvbStdModuleDef.GetMethod( i : integer ) : TvbAbstractFuncDef;
    begin
      result := TvbAbstractFuncDef( fMethods[i] );
    end;

  function TvbStdModuleDef.GetMethodCount : integer;
    begin
      if fMethods = nil
        then result := 0
        else result := fMethods.Count;
    end;

  function TvbStdModuleDef.AddProperty( p : TvbPropertyDef ) : TvbPropertyDef;
    begin
      result := p;
      if fProps = nil
        then fProps := TvbNodeList.Create;
      fProps.Add( p );
      fScope.Bind( p, OBJECT_BINDING );
    end;

  function TvbStdModuleDef.GetVar( i : integer ) : TvbVarDef;
    begin
      result := TvbVarDef( fVars[i] );
    end;

  function TvbStdModuleDef.GetVarCount: integer;
    begin
      if fVars = nil
        then result := 0
        else result := fVars.Count;
    end;

  function TvbStdModuleDef.AddVar( v : TvbVarDef ) : TvbVarDef;
    begin
      result := v;
      if fVars = nil
        then fVars := TvbNodeList.Create;
      fVars.Add( v );
      fScope.Bind( v, OBJECT_BINDING );
    end;

  function TvbStdModuleDef.AddEvent( e : TvbEventDef ) : TvbEventDef;
    begin
      result := nil;
      assert( true );
    end;

  function TvbStdModuleDef.GetEvent( i : integer ) : TvbEventDef;
    begin
      result := nil;
    end;

  function TvbStdModuleDef.GetEventCount : integer;
    begin
      result := 0;
    end;

  { TvbClassModuleDef }

  procedure TvbClassModuleDef.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnClassModuleDef( self, Context );
    end;

  function TvbClassModuleDef.AddEvent( e : TvbEventDef ) : TvbEventDef;
    begin
      result := e;
      fClassDef.AddEvent( e );
    end;

  function TvbClassModuleDef.AddMethod( m : TvbAbstractFuncDef ) : TvbAbstractFuncDef;
    begin
      if SameText( m.Name, 'Class_Initialize' )
        then fClassDef.ClassInitialize := m
        else
      if SameText( m.Name, 'Class_Terminate' )
        then fClassDef.ClassTerminate := m
        else fClassDef.AddMethod( m );
      result := m;
    end;

  function TvbClassModuleDef.AddProperty( p : TvbPropertyDef ) : TvbPropertyDef;
    begin
      result := p;
      fClassDef.AddProperty( p );
    end;

  function TvbClassModuleDef.AddVar( v : TvbVarDef ) : TvbVarDef;
    begin
      result := v;
      fClassDef.AddVar( v );
    end;

  constructor TvbClassModuleDef.Create( const name, relativePath : string; ParentScope : TvbScope );
    begin
      inherited Create( CLASS_MODULE_DEF, Name, relativePath, ParentScope );
      fClassDef := AddClass( CreateClassInstance( name ) );
    end;

  function TvbClassModuleDef.GetEvent( i : integer ) : TvbEventDef;
    begin
      result := fClassDef.Event[i];
    end;

  function TvbClassModuleDef.GetEventCount : integer;
    begin
      result := fClassDef.EventCount;
    end;

  function TvbClassModuleDef.GetMethod( i : integer ) : TvbAbstractFuncDef;
    begin
      result := fClassDef.Method[i];
    end;

  function TvbClassModuleDef.GetMethodCount : integer;
    begin
      result := fClassDef.MethodCount;
    end;

  function TvbClassModuleDef.GetProperty( i : integer ) : TvbPropertyDef;
    begin
      result := fClassDef.Prop[i];
    end;

  function TvbClassModuleDef.GetPropCount : integer;
    begin
      result := fClassDef.PropCount;
    end;

  function TvbClassModuleDef.GetVar( i : integer ) : TvbVarDef;
    begin
      result := fClassDef.Vars[i];
    end;

  function TvbClassModuleDef.GetVarCount : integer;
    begin
      result := fClassDef.VarCount;
    end;

  function TvbClassModuleDef.CreateClassInstance( const name : string ) : TvbClassDef;
    begin
      result := TvbClassDef.Create( [dfPublic], name );
    end;

  { TvbNodeVisitor }

  procedure TvbNodeVisitor.OnAbsExpr( Node : TvbAbsExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnAddressOfExpr( Node : TvbAddressOfExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnArrayExpr( Node : TvbArrayExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnAssignStmt( Node : TvbAssignStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnBinaryExpr( Node : TvbBinaryExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnBoolLitExpr( Node : TvbBoolLit; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnCallOrIndexerExpr( Node : TvbCallOrIndexerExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnCallStmt( Node : TvbCallStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnCastExpr( Node : TvbCastExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnCircleStmt( Node : TvbCircleStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnClassDef( Node : TvbClassDef; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnClassModuleDef( Node : TvbClassModuleDef; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnCloseStmt( Node : TvbCloseStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnConstDef( Node : TvbConstDef; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnDateExpr( Node : TvbDateExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnDateLitExpr( Node : TvbDateLit; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnDateStmt( Node : TvbDateStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnDebugAssertStmt( Node : TvbDebugAssertStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnDebugPrintStmt( Node : TvbDebugPrintStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnDictAccessExpr( Node : TvbDictAccessExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnDllFuncDef( Node : TvbDllFuncDef; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnDoEventsExpr( Node : TvbDoEventsExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnDoLoopStmt( Node : TvbDoLoopStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnEndStmt( Node : TvbEndStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnEnumDef( Node : TvbEnumDef; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnEraseStmt( Node : TvbEraseStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnErrorStmt( Node : TvbErrorStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnEventDef( Node : TvbEventDef; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnExitStmt( Node : TvbExitStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnFixExpr( Node : TvbFixExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnFloatLitExpr( Node : TvbFloatLit; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnForeachStmt( Node : TvbForeachStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnForStmt( Node : TvbForStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnFunctionDef( Node : TvbFuncDef; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnFuncResultExpr( Node : TvbFuncResultExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnGetStmt( Node : TvbGetStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnGotoOrGosubStmt( Node : TvbGotoOrGosubStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnIfStmt( Node : TvbIfStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnInputExpr( Node : TvbInputExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnInputStmt( Node : TvbInputStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnLabelDef( Node : TvbLabelDef; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnLabelDefStmt( Node : TvbLabelDefStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnLBoundExpr( Node : TvbLBoundExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnLenBExpr( Node : TvbLenBExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnLenExpr( Node : TvbLenExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnLibModuleDef( Node : TvbLibModuleDef; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnLineStmt( Node : TvbLineStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnLineInputStmt( Node : TvbLineInputStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnLockStmt( Node : TvbLockStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnMeExpr( Node : TvbMeExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnMemberAccessExpr( Node : TvbMemberAccessExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnMidAssignStmt( Node : TvbMidAssignStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnMidExpr( Node : TvbMidExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnNameExpr( Node : TvbNameExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnNameStmt( Node : TvbNameStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnNewExpr( Node : TvbNewExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnNothingExpr( Node : TvbNothingLit; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnOnErrorStmt( Node : TvbOnErrorStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnOnGotoOrGosubStmt( Node : TvbOnGotoOrGosubStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnOpenStmt( Node : TvbOpenStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnParamDef( Node : TvbParamDef; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnPrintArgExpr( Node : TvbPrintStmtExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnPrintStmt( Node : TvbPrintStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnProjectDef( Node : TvbProject; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnPropertyDef( Node : TvbPropertyDef; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnPSetStmt( Node : TvbPSetStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnPutStmt( Node : TvbPutStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnQualifiedName( Node : TvbQualifiedName; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnRaiseEventStmt( Node : TvbRaiseEventStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnRecordDef( Node : TvbRecordDef; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnReDimStmt( Node : TvbReDimStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnResumeStmt( Node : TvbResumeStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnReturnStmt( Node : TvbReturnStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnSeekExpr( Node : TvbSeekExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnSeekStmt( Node : TvbSeekStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnSgnExpr( Node : TvbSgnExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnSimpleName( Node : TvbSimpleName; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnStdModuleDef( Node : TvbStdModuleDef; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnStmtBlock( Node : TvbStmtBlock; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnStopStmt( Node : TvbStopStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnStringExpr( Node : TvbStringExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnStringLitExpr( Node : TvbStringLit; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnTimeStmt( Node : TvbTimeStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnIntegerLitExpr( Node : TvbIntLit; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnIntExpr( Node : TvbIntExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnTypeRef( Node : TvbType; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnTypeOfExpr( Node : TvbTypeOfExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnUboundExpr( Node : TvbUBoundExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnUnaryExpr( Node : TvbUnaryExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnUnlockStmt( Node : TvbUnlockStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnVarDef( Node : TvbVarDef; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnWidthStmt( Node : TvbWidthStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnWithStmt( Node : TvbWithStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnWriteStmt( Node : TvbWriteStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnCaseClause( Node : TvbCaseClause; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnCaseStmt( Node : TvbCaseStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnRangeCaseClause( Node : TvbRangeCaseClause; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnRelationalCaseClause( Node : TvbRelationalCaseClause; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnSelectCaseStmt( Node : TvbSelectCaseStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnExitLoopStmt( Node : TvbExitLoopStmt; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnNamedArgExpr( Node : TvbNamedArgExpr; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnFormModuleDef( Node : TvbFormModuleDef; Context : TObject );
    begin
    end;

  procedure TvbNodeVisitor.OnFormDef( Node : TvbFormDef; Context : TObject );
    begin
    end;

{ TvbFormModuleDef }

  procedure TvbFormModuleDef.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnFormModuleDef( self, Context );
    end;

  constructor TvbFormModuleDef.Create( const name, relativePath : string; ParentScope : TvbScope );
    begin
      inherited;
      fKind := FORM_MODULE_DEF;
      // note that we create the var without type.
      // it's the responsability of form module's owner to set the type. 
      fFormVar := TvbVarDef.Create( [], Name );
      fScope.Bind( fFormDef, OBJECT_BINDING );
    end;

  function TvbFormModuleDef.CreateClassInstance( const name : string ) : TvbClassDef;
    begin
      fFormDeF := TvbFormDef.Create( [dfPublic], name );
      result := fFormDef;
    end;

  destructor TvbFormModuleDef.Destroy;
    begin
      FreeAndNil( fFormVar );
      inherited;
    end;

  { TvbExitLoopStmt }

  procedure TvbExitLoopStmt.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnExitLoopStmt( self, Context );
    end;

  constructor TvbExitLoopStmt.Create;
    begin
      inherited Create( EXIT_LOOP_STMT );
    end;

  { TvbNamedArgExpr }

  procedure TvbNamedArgExpr.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
    end;

  constructor TvbNamedArgExpr.Create( const argName : string; value : TvbExpr );
    begin
      inherited Create( NAMED_ARG_EXPR, value );
      fArgName := argName;
    end;

  { TvbFormDef }

  procedure TvbFormDef.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
      Visitor.OnFormDef( self, Context );
    end;

  constructor TvbFormDef.Create( const flags : TvbNodeFlags; const name, name1 : string );
    begin
      inherited Create( flags, name, name1, nil );
      fControls := TStringList.Create;
      fControls.Sorted := true;
      fControls.Duplicates := dupAccept;
      //fKind := FORM_DEF;
    end;

  destructor TvbFormDef.Destroy;
    begin
      // the controls inside this list are owned by fRootControl.
      FreeAndNil( fControls );
      FreeAndNil( fRootControl );
      inherited;
    end;

  function TvbFormDef.GetControl( i : integer ) : TvbControl;
    begin
      result := TvbControl( fControls.Objects[i] );
    end;

  function TvbFormDef.GetControlCount : integer;
    begin
      result := fControls.Count;
    end;

  procedure TvbFormDef.OnNewControl( control : TvbControl );
    var
      ctl : TvbDef;
    begin
      fControls.AddObject( control.Name, control );
      // before inserting the control in the form's scope
      // check that there is no another control with the same name.
      // duplicate controls occur when there are control arrays in the form.
      // these are controls that share the same name and have different values
      // in their "Index" properties.
      if not fScope.Lookup( OBJECT_BINDING, control.Name, ctl, false )
        then fScope.Bind( control, OBJECT_BINDING );
    end;

  { TvbControl }

  procedure TvbControl.AddControl( ctl : TvbControl );
    begin
      assert( ctl <> nil );
      fControls.AddObject( ctl.Name, ctl );
    end;

  function TvbControl.AddObjectProperty( const name : string ) : TvbControl;
    begin
      result := AddProperty( name, cpkControl ).Control;
    end;

  function TvbControl.AddProperty( const name : string; kind : TvbControlPropertyKind ) : TvbControlProperty;
    var
      i : integer;
    begin
      // if we already have a property with the given name return it to the
      // caller. otherwise create a new property.
      if fProps.Find( name, i )
        then result := TvbControlProperty( fProps.Objects[i] )
        else
          begin
            result := TvbControlProperty.Create( kind, name );
            try
              fProps.AddObject( name, result );
            except
              FreeAndNil( result );
              raise;
            end;
          end;
    end;

  constructor TvbControl.Create( const name : string; typ : IvbType );
    begin
      inherited Create( [dfPublic], name, typ );
      fControls               := TStringList.Create;
      fControls.Sorted        := true;
      fControls.Duplicates    := dupAccept;
      fControls.CaseSensitive := false;
      fProps                  := TStringList.Create;
      fProps.Sorted           := true;
      fProps.Duplicates       := dupError;
      fProps.CaseSensitive    := false;
    end;

  destructor TvbControl.Destroy;
    var
      i : integer;
    begin
      for i := 0 to fControls.Count - 1 do
        fControls.Objects[i].Free;
      for i := 0 to fProps.Count - 1 do
        fProps.Objects[i].Free;
      FreeAndNil( fControls );
      FreeAndNil( fProps );
      inherited;
    end;

  function TvbControl.GetControl( i : integer ) : TvbControl;
    begin
      result := TvbControl( fControls.Objects[i] );
    end;

  function TvbControl.GetControlCount : integer;
    begin
      result := fControls.Count;
    end;

  function TvbControl.GetProp( i : integer ) : TvbControlProperty;
    begin
      result := TvbControlProperty( fProps.Objects[i] );
    end;

  function TvbControl.GetPropByName( const name : string ) : TvbControlProperty;
    var
      i : integer;
    begin
      if fProps.Find( name, i )
        then result := TvbControlProperty( fProps.Objects[i] )
        else result := nil;
    end;

  function TvbControl.GetPropCount: integer;
    begin
      result := fProps.Count;
    end;

  { TvbCircleStmt }

  procedure TvbCircleStmt.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
    end;

  constructor TvbCircleStmt.Create( obj, x, y, r : TvbExpr;
                                    step : boolean;
                                    color, startAngle, endAngle, aspect : TvbExpr );
    begin
      inherited Create( CIRCLE_STMT );
      fObject := obj;
      fX := x;
      fY := y;
      fR := r;
      fStep := step;
      fColor := color;
      fStartAngle := startAngle;
      fEndAngle := endAngle;
      fAspect := aspect;
    end;

  destructor TvbCircleStmt.Destroy;
    begin
      FreeAndNil( fX );
      FreeAndNil( fY );
      FreeAndNil( fR );
      FreeAndNil( fObject );
      FreeAndNil( fColor );
      FreeAndNil( fStartAngle );
      FreeAndNil( fEndAngle );
      FreeAndNil( fAspect );
      inherited;
    end;

  { TvbLineStmt }

  procedure TvbLineStmt.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
    end;

  constructor TvbLineStmt.Create( obj, x1, y1, x2, y2, color : TvbExpr;
                                  step, filled, boxed : boolean);
    begin
      inherited Create( LINE_STMT );
      fX1 := x1;
      fY1 := y1;
      fX2 := x2;
      fY2 := y2;
      fObject := obj;
      fColor := color;
      fStep := step;
      fFilled := filled;
      fBoxed := boxed;
    end;

  destructor TvbLineStmt.Destroy;
    begin
      FreeAndNil( fX1 );
      FreeAndNil( fY1 );
      FreeAndNil( fX2 );
      FreeAndNil( fY2 );
      FreeAndNil( fColor );
      FreeAndNil( fObject );
      inherited;
    end;

  { TvbPSetStmt }

  procedure TvbPSetStmt.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
    end;

  constructor TvbPSetStmt.Create( obj, x, y, color : TvbExpr );
    begin
      inherited Create( PSET_STMT );
      fX := x;
      fY := y;
      fColor := color;
      fObject := obj;
    end;

  destructor TvbPSetStmt.Destroy;
    begin
      FreeAndNil( fX );
      FreeAndNil( fY );
      FreeAndNil( fObject );
      FreeAndNil( fColor );
      inherited;
    end;

  { TvbControlProperty }

  constructor TvbControlProperty.Create( kind : TvbControlPropertyKind; const name : string );
    begin
      fKind := kind;
      fName := name;
      if kind = cpkControl
        then fControl := TvbControl.Create( name, nil ); 
    end;

  destructor TvbControlProperty.Destroy;
    begin
      FreeAndNil( fControl );
      inherited;
    end;

  { TvbPrintStmt }

  procedure TvbPrintStmt.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
    end;

  constructor TvbPrintStmt.Create;
    begin
      inherited Create( PRINT_STMT );
    end;

  destructor TvbPrintStmt.Destroy;
    begin
      FreeAndNil( fFileNum );
      inherited;
    end;

  { TvbPrintMethod }

  procedure TvbPrintMethod.Accept( Visitor : TvbNodeVisitor; Context : TObject );
    begin
    end;

  constructor TvbPrintMethod.Create;
    begin
      inherited Create( PRINT_METHOD_STMT );
    end;

  destructor TvbPrintMethod.Destroy;
    begin
      FreeAndNil( fObj );
      inherited;
    end;

initialization

  vbAnyType          := TvbType.Create( ANY_TYPE );
  vbBooleanType      := TvbType.Create( BOOLEAN_TYPE );
  vbByteType         := TvbType.Create( BYTE_TYPE );
  vbCurrencyType     := TvbType.Create( CURRENCY_TYPE );
  vbDateType         := TvbType.Create( DATE_TYPE );
  vbDoubleType       := TvbType.Create( DOUBLE_TYPE );
  vbIntegerType      := TvbType.Create( INTEGER_TYPE );
  vbLongType         := TvbType.Create( LONG_TYPE );
  vbObjectType       := TvbType.Create( OBJECT_TYPE );
  vbSingleType       := TvbType.Create( SINGLE_TYPE );
  vbStringType       := TvbType.Create( STRING_TYPE );
  vbVariantType      := TvbType.Create( VARIANT_TYPE );
  vbVariantArrayType := TvbType.CreateArray( VARIANT_TYPE );
  _vbUnknownType     := TvbType.Create( UNKNOWN_TYPE );

  vbBooleanDynArrayType  := TvbType.CreateArray( BOOLEAN_TYPE );
  vbByteDynArrayType     := TvbType.CreateArray( BYTE_TYPE );
  vbCurrencyDynArrayType := TvbType.CreateArray( CURRENCY_TYPE );
  vbDateDynArrayType     := TvbType.CreateArray( DATE_TYPE );
  vbDoubleDynArrayType   := TvbType.CreateArray( DOUBLE_TYPE );
  vbIntegerDynArrayType  := TvbType.CreateArray( INTEGER_TYPE );
  vbLongDynArrayType     := TvbType.CreateArray( LONG_TYPE );
  vbObjectDynArrayType   := TvbType.CreateArray( OBJECT_TYPE );
  vbSingleDynArrayType   := TvbType.CreateArray( SINGLE_TYPE );
  vbStringDynArrayType   := TvbType.CreateArray( STRING_TYPE );
  vbVariantDynArrayType  := TvbType.CreateArray( VARIANT_TYPE );

finalization

  vbAnyType      := nil;
  vbBooleanType  := nil;
  vbByteType     := nil;
  vbCurrencyType := nil;
  vbDateType     := nil;
  vbDoubleType   := nil;
  vbIntegerType  := nil;
  vbLongType     := nil;
  vbObjectType   := nil;
  vbSingleType   := nil;
  vbStringType   := nil;
  vbVariantType  := nil;
  _vbUnknownType := nil;

  vbBooleanDynArrayType  := nil;
  vbByteDynArrayType     := nil;
  vbCurrencyDynArrayType := nil;
  vbDateDynArrayType     := nil;
  vbDoubleDynArrayType   := nil;
  vbIntegerDynArrayType  := nil;
  vbLongDynArrayType     := nil;
  vbObjectDynArrayType   := nil;
  vbSingleDynArrayType   := nil;
  vbStringDynArrayType   := nil;
  vbVariantDynArrayType  := nil;

end.








































































































































































































































