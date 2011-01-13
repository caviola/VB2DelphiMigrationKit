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

unit vbParser;

interface

  uses
    SysUtils,
    Classes,
    vbTokens,
    vbLexer,
    vbNodes,
    vbClasses,
    vbBuffers,
    vbOptions,
    CollHash,
    RegExpr;

  type

    TvbParseFormDefinitionBeginEvent = procedure ( form : TvbFormDef ) of object;
    TvbParseFormDefinitionEndEvent   = procedure of object;
    TvbFormDefinitionProblemEvent    = procedure ( line, column : cardinal; const msg : string ) of object;
    TvbParseFormControlBeginEvent    = procedure ( control : TvbControl ) of object;
    TvbParseFormControlEndEvent      = procedure of object;

    TvbParseModuleBeginEvent = procedure ( const name, relativePath : string ) of object;
    TvbParseModuleEndEvent   = procedure of object;
    TvbParseModuleFailEvent  = procedure of object;

    TvbParseProjectBeginEvent = procedure ( const name, relativePath : string ) of object;
    TvbParseProjectEndEvent   = procedure of object;
    TvbParseProjectFailEvent  = procedure of object;

    TvbOptionCompare = ( ocBinary, ocText, ocDatabase );

    TvbModuleOptions =
      class
        private
          fCompare       : TvbOptionCompare;
          fDefaultTypes  : array[char] of IvbType;
          fExplicit      : boolean;
          fPrivateModule : boolean;
        public
          constructor Create;
          procedure SetDefaults;
          procedure DefineType( fromChar, toChar : char; typk : IvbType );
          function GetDefinedType( const name : string ) : IvbType;
          property optCompare       : TvbOptionCompare read fCompare       write fCompare;
          property optExplicit      : boolean          read fExplicit      write fExplicit;
          property optPrivateModule : boolean          read fPrivateModule write fPrivateModule;
      end;
  
    TvbParser =
      class
        private
          // Hold the current statement block.
          // The statement parsing functions add to this block instance.
          fCurrBlock : TvbStmtBlock;
          // The current Sub/Function/Property being parsed.
          fCurrFunc : TvbAbstractFuncDef;
          // The name of the current "Function".
          // We use this to determine when an expression is refering to the
          // function's result value.
          fCurrFuncName : string;
          // The current module being parsed.
          // This is an instance of TvbStdModuleDef, TvbClassModuleDef or TvbFormModuleDef.
          fCurrModule : TvbModuleDef;
          fCurrNodeFlags : TvbNodeFlags;   
          // The current scope.
          // When parsing a standard module, this is the module's scope.
          // But when parsing a class/form, this is the scope of the class/form object.
          fCurrScope : TvbScope;
          // The list of comments atop the current declaration/statement.
          fCurrTopCmts : TStringList;    
          // The most-recent "With" statement expression.
          fCurrWithExpr : TvbExpr;
          // The most-recent linked list of label declarations atop an statement
          // or at the end of a block, in which case a TvbLabelDeclStmt is inserted
          // to capture this labels.
          fFrontLabel : TvbLabelDef;
          // The last label in the linked list.
          // We use this to avoid traversing the label linked list to add a new one.
          fLastLabel : TvbLabelDef;
          fLexer : TvbLexer;
          // Module-level options for the module being parsed.
          // Before starting parsing a module, these options are reset to their
          // default values. 
          fModuleOpts : TvbModuleOptions;
          fMonitor : IvbProgressMonitor;
          // General conversion options.
          fOptions : TvbOptions;
          fPropertyMap : TvbPropertyMap;
          fRegexp : TRegExpr;
          fResourceProvider : IvbResourceProvider;
          fTempVarCount : cardinal;
          // The token being currently analized.
          fToken : TvbToken;
          fTypeMap : TvbModuleTypeMap;
          
          fOnFormDefinitionProblem    : TvbFormDefinitionProblemEvent;
          fOnParseModuleBegin         : TvbParseModuleBeginEvent;
          fOnParseModuleEnd           : TvbParseModuleEndEvent;
          fOnParseModuleFail          : TvbParseModuleFailEvent;
          fOnParseProjectBegin        : TvbParseProjectBeginEvent;
          fOnParseProjectEnd          : TvbParseProjectEndEvent;
          fOnParseProjectFail         : TvbParseProjectFailEvent;
          fOnParseFormDefinitionBegin : TvbParseFormDefinitionBeginEvent;
          fOnParseFormDefinitionEnd   : TvbParseFormDefinitionEndEvent;
          fOnParseFormControlBegin    : TvbParseFormControlBeginEvent;
          fOnParseFormControlEnd      : TvbParseFormControlEndEvent;

          procedure Error( const msg : string );
          procedure CheckError( cond : boolean; const msg : string );
        protected
          procedure Log( const msg : string; line : cardinal; col : cardinal; logType : TvbLogType = ltInfo ); overload;
          procedure Log( const msg : string; logType : TvbLogType = ltInfo ); overload;
        
          function CheckRightCmt : string;
          procedure CheckMoveCommentsToTop;

          // Token handling
          function CheckIdent( const name : string ) : boolean;
          function CheckToken( tk : TvbTokenKind ) : boolean; overload;
          function CheckToken( tks : TvbTokenKinds; var TokenMatch : TvbTokenKind ) : boolean; overload;
          function GetIdent( var sfx : TvbTypeSfx; AllowKeyword : boolean = false ) : string;
          function PeekToken : TvbToken;
          procedure GetToken;
          procedure NeedIdent( const name : string ); overload;
          procedure NeedToken( tk : TvbTokenKind );

          function CheckFormIdent( const name : string ) : boolean;
          function CheckFormToken( tk : TvbTokenKind ) : boolean; 
          procedure NeedFormIdent( const name : string );
          procedure NeedFormToken( tk : TvbTokenKind );
          procedure GetFormToken;

          procedure ParseModule;
          procedure SkipBlankLines;
          procedure SkipVBAttrs;

          // Declaration parsing
          function ParseConstDef( AllowKeyword : boolean ) : TvbConstDef;
          function ParseDllFuncDef : TvbDllFuncDef;
          function ParseEnumDef : TvbEnumDef;
          function ParseEventDef : TvbEventDef;
          procedure ParseFunctionDef;
          function ParseName : TvbName;
          function ParseParamDef : TvbParamDef;
          procedure ParsePropertyFuncDef;
          function ParseRecordDef : TvbRecordDef;
          function ParseBaseType : IvbType;
          function ParseVarDef( AllowKeyword, MustBeTyped : boolean ) : TvbVarDef;
          procedure ParseDeclFlags;
          procedure ParseDef;
          procedure ParseDefStmt;
          procedure ParseOptionStmt;

          // Expressions parsing
          function ParseExpr( parseNameArgs : boolean = false ) : TvbExpr;
          function ParseTermExpr( parseNamedArgs : boolean = false ) : TvbExpr;
          //function ParseMidExpr : TvbMidExpr;
          function CalcBinaryExpr( op : TvbOperator; left, right : TvbExpr ) : TvbExpr;
          function GetTempVarName : string;
          function ParseArgList : TvbArgList;
          function ParseAssignStmt( assign : TvbAssignKind ) : TvbAssignStmt;
          function ParseLabelRef : TvbSimpleName;
          procedure ParsePrintArgList( printStmt : TvbAbstractPrintStmt );
          function ParseSelectStmt : TvbSelectCaseStmt;
          function ParseStmtBlock( TermTokens : TvbTokenKinds ) : TvbStmtBlock;
          procedure ParseCallStmt;
          procedure ParseCloseStmt;
          procedure ParseDateStmt;
          procedure ParseDebugAssertStmt;
          procedure ParseDebugPrintStmt;
          procedure ParseDoLoopStmt;
          procedure ParseEndStmt;
          procedure ParseEraseStmt;
          procedure ParseForeachStmt;
          procedure ParseFormDefinition( form : TvbFormDef );
          procedure ParseForStmt;
          procedure ParseGetStmt;
          procedure ParseGotoOrGosubStmt;
          procedure ParseIfStmt;
          procedure ParseInputStmt;
          procedure ParseLocalVarOrConst;
          procedure ParseLockStmt;
          procedure ParseMidAssignStmt;
          procedure ParseOnErrorStmt;
          procedure ParseOnGotoOrGosubStmt;
          procedure ParseOpenStmt;
          procedure ParsePrintStmt;
          procedure ParsePrintMethod( obj : TvbExpr = nil );
          procedure ParsePrintStmtOrMethod( obj : TvbExpr = nil );
          procedure ParseProjectOption( const Option, Value : string; Project : TvbProject );
          procedure ParsePutStmt;
          procedure ParseRaiseEventStmt;
          procedure ParseReDimStmt;
          procedure ParseResumeStmt;
          procedure ParseReturnStmt;
          procedure ParseSeekStmt;
          procedure ParseStopStmt;
          procedure ParseUnlockStmt;
          procedure ParseWhileStmt;
          procedure ParseWithStmt;
          procedure ParseWriteStmt;
          procedure ParseWidthStmt;
          procedure ParseErrorStmt;
          procedure ParseNameStmt;
          procedure ParseCircleStmt( obj : TvbExpr = nil );
          procedure ParseLineStmt( obj : TvbExpr = nil );
          procedure ParsePSetStmt( obj : TvbExpr = nil );
          procedure ParseLineInputStmt;

          // Form definition parsing.
          function ParseControl( form : TvbFormDef ) : TvbControl;
          procedure ParseObjectProperty( owner : TvbControl );
          procedure ParseProperty( owner : TvbControl );
          procedure SkipFormLine;

          procedure DoFormDefinitionProblem( const msg : string; skipLine : boolean = true );
          procedure DoParseModuleBegin( const name, relativePath : string );
          procedure DoParseModuleEnd;
          procedure DoParseModuleFail;
          procedure DoParseProjectBegin( const name, relativePath : string );
          procedure DoParseProjectEnd;
          procedure DoParseProjectFail;
          procedure DoParseFormContronBegin( control : TvbControl );
          procedure DoParseFormContronEnd;

          procedure DoParseFormDefinitionBegin( form : TvbFormDef );
          procedure DoParseFormDefinitionEnd;

          function IsModuleLevelDef( const name : string ) : boolean;
        public
          constructor Create( opts : TvbOptions; resourceProvider : IvbResourceProvider; const monitor : IvbProgressMonitor = nil );
          destructor Destroy; override;
          function ParseClassModule( const Name, relPath : string; Source : TvbCustomTextBuffer ) : TvbModuleDef;
          function ParseFormModule( const Name, relPath : string; Source : TvbCustomTextBuffer ) : TvbModuleDef;
          function ParseProject( Lines : TStringList; const relPath : string ) : TvbProject;
          function ParseStdModule( const Name, relPath : string; Source : TvbCustomTextBuffer ) : TvbModuleDef;

          // event fired by the parser
          property OnFormDefinitionProblem    : TvbFormDefinitionProblemEvent    read fOnFormDefinitionProblem    write fOnFormDefinitionProblem;
          property OnParseFormControlBegin    : TvbParseFormControlBeginEvent    read fOnParseFormControlBegin    write fOnParseFormControlBegin;
          property OnParseFormControlEnd      : TvbParseFormControlEndEvent      read fOnParseFormControlEnd      write fOnParseFormControlEnd;
          property OnParseFormDefinitionBegin : TvbParseFormDefinitionBeginEvent read fOnParseFormDefinitionBegin write fOnParseFormDefinitionBegin;
          property OnParseFormDefinitionEnd   : TvbParseFormDefinitionEndEvent   read fOnParseFormDefinitionEnd   write fOnParseFormDefinitionEnd;
          property OnParseModuleBegin         : TvbParseModuleBeginEvent         read fOnParseModuleBegin         write fOnParseModuleBegin;
          property OnParseModuleEnd           : TvbParseModuleEndEvent           read fOnParseModuleEnd           write fOnParseModuleEnd;
          property OnParseModuleFail          : TvbParseModuleFailEvent          read fOnParseModuleFail          write fOnParseModuleFail;
          property OnParseProjectBegin        : TvbParseProjectBeginEvent        read fOnParseProjectBegin        write fOnParseProjectBegin;
          property OnParseProjectEnd          : TvbParseProjectEndEvent          read fOnParseProjectEnd          write fOnParseProjectEnd;
          property OnParseProjectFail         : TvbParseProjectFailEvent         read fOnParseProjectFail         write fOnParseProjectFail;
      end;

implementation

  uses
    Types,
    ArrayUtils,
    Collections,
    CollWrappers, Variants;

  var
    // Table to map type suffixes to intrinsic types.
    TypeSfx2Type : array[STRING_SFX..CURRENCY_SFX] of IvbType;

  type

    TvbTokenEqvs =         
      record
        eqvOperator   : TvbOperator;
        eqvSimpleType : IvbType;
        eqvDeclFlags  : TvbNodeFlags;
        eqvParamFlags : TvbNodeFlags;
      end;

  var
    TokenEqvTable : array[TvbTokenKind] of TvbTokenEqvs;

  { TvbParser }

  constructor TvbParser.Create( opts : TvbOptions; resourceProvider : IvbResourceProvider; const monitor : IvbProgressMonitor );
    begin
      assert( opts <> nil );
      fRegexp           := TRegExpr.Create;
      fCurrTopCmts      := TStringList.Create;
      fLexer            := TvbLexer.Create;
      fModuleOpts       := TvbModuleOptions.Create;
      fPropertyMap      := TvbPropertyMap.Create;
      fTypeMap          := TvbModuleTypeMap.Create;
      fOptions          := opts;
      fResourceProvider := resourceProvider;
      fMonitor          := monitor;
    end;

  procedure TvbParser.GetToken;
    begin
      fToken := fLexer.GetToken;
      with fToken do
        case tokenKind of
          tkEnd :
            case fLexer.PeekToken.tokenKind of
              tkFunction :
                begin
                  tokenKind := _tkEndFunction;
                  fLexer.GetToken;
                end;
              tkIf :
                begin
                  tokenKind := _tkEndIf;
                  fLexer.GetToken;
                end;
              tkProperty :
                begin
                  tokenKind := _tkEndProperty;
                  fLexer.GetToken;
                end;
              tkSelect :
                begin
                  tokenKind := _tkEndSelect;
                  fLexer.GetToken;
                end;
              tkSub :
                begin
                  tokenKind := _tkEndSub;
                  fLexer.GetToken;
                end;
              tkWith :
                begin
                  tokenKind := _tkEndWith;
                  fLexer.GetToken;
                end;
            end;
        end;
    end;

  function TvbParser.PeekToken : TvbToken;
    begin
      result := fLexer.PeekToken;
    end;

  function TvbParser.CheckIdent( const name : string ) : boolean;
    begin
      result := false;
      if ( fToken.tokenKind = tkIdent ) and SameText( fToken.tokenText, name )
        then
          begin
            result := true;
            GetToken;
          end;
    end;

  function TvbParser.CheckToken( tk : TvbTokenKind ) : boolean;
    begin
      result := false;
      if fToken.tokenKind = tk
        then
          begin
            result := true;
            GetToken;
          end;
    end;

  function TvbParser.CheckToken( tks : TvbTokenKinds; var TokenMatch : TvbTokenKind ) : boolean;
    begin
      result := false;
      if fToken.tokenKind in tks
        then
          begin
            TokenMatch := fToken.tokenKind;
            result := true;
            GetToken;
          end;
    end;

  function TvbParser.GetIdent( var sfx : TvbTypeSfx; AllowKeyword : boolean ) : string;
    begin
      with fToken do
        if ( tokenKind = tkIdent ) or
           ( AllowKeyword and ( tokenKind in [tkAbs..tkXor] ) )
          then
            begin
              result := tokenText;
              if TF_INTEGER_SFX in tokenFlags
                then sfx := INTEGER_SFX
                else
              if TF_LONG_SFX in tokenFlags
                then sfx := LONG_SFX
                else
              if TF_SINGLE_SFX in tokenFlags
                then sfx := SINGLE_SFX
                else
              if TF_DOUBLE_SFX in tokenFlags
                then sfx := DOUBLE_SFX
                else
              if TF_STRING_SFX in tokenFlags
                then sfx := STRING_SFX
                else
              if TF_CURRENCY_SFX in tokenFlags
                then sfx := CURRENCY_SFX
                else sfx := NO_SFX;
              GetToken;
            end
          else Error( 'identifier expected' );
    end;

  procedure TvbParser.NeedIdent( const name : string ); 
    begin
      if not CheckIdent( name )
        then Error( Format( 'expected "%s"', [name] ) );
    end;

  procedure TvbParser.NeedToken( tk : TvbTokenKind );
    begin
      if not CheckToken( tk )
        then Error( Format( 'expected "%s"', [GetTokenText( tk )] ) );
    end;

  procedure TvbParser.SkipBlankLines;
    begin
      while fToken.tokenKind = tkEOL do
        GetToken;
    end;

  procedure TvbParser.SkipVBAttrs;
    begin
      while CheckToken( tkAttribute ) do
        repeat
          GetToken;
        until CheckToken( tkEOL );
    end;

  function TvbParser.ParseName : TvbName;
    var
      sfx : TvbTypeSfx;
    begin
      result := nil;
      with fToken do
        begin
          if tokenKind in [tkIdent, tkAbs..tkXor]
            then
              begin
                result := TvbSimpleName.Create( tokenText );
                GetToken;
              end
            else Error( 'name expected' );
          try
            while CheckToken( tkDot ) do
              result := TvbQualifiedName.Create( GetIdent( sfx, true ), result );
          except
            result.Free;
            raise;
          end;
        end;
    end;

  function TvbParser.ParseBaseType : IvbType;
    var
      n : string;
    begin
      with fToken do
        if TF_TYPE in tokenFlags
          then
            if tokenKind = tkString
              then
                begin
                  GetToken;
                  if CheckToken( tkMult ) // '*'
                    then result := TvbType.CreateString( ParseTermExpr )
                    else result := vbStringType;
                end
              else
                begin
                  result := TokenEqvTable[tokenKind].eqvSimpleType;
                  GetToken;
                end
          else
            begin
              n := '';
              repeat
                 if tokenKind in [tkIdent, tkAbs..tkXor]
                  then n := n + tokenText
                  else Error( 'type name expected' );
                 GetToken;
                 if CheckToken( tkDot )
                   then n := n + '.'
                   else break;
              until false;
              result := fTypeMap.GetType( n );
            end;
    end;

  function TvbParser.ParseLabelRef : TvbSimpleName;
    var
      s : string;
    begin
      with fToken do
        case tokenKind of
          tkIdent :
            begin
              s := tokenText;
              GetToken;
            end;
          tkIntLit :
            begin
              // We treat numeric labels as if they were strings.
              s := IntToStr( tokenInt );
              GetToken;
            end;
          else Error( 'invalid label name' );
        end;
      result := TvbSimpleName.Create( s );
    end;

  procedure TvbParser.Error( const msg : string );
    begin
      raise EvbSourceError.Create( msg, fLexer.TokenLine, fLexer.TokenColumn );
    end;

  // Option Base {0 | 1}
  // Option Compare {"Binary" | "Text" | "Database"}
  // Option Explicit
  // Option Private Module
  procedure TvbParser.ParseOptionStmt;
    begin
      GetToken;
      with fToken, fModuleOpts do
        if CheckToken( tkPrivate )
          then
            begin
              NeedIdent( 'Module' );
              optPrivateModule := true;
            end
          else
        if CheckIdent( 'Base' )
          then
            begin
              CheckError( tokenKind = tkIntLit, '"Option Base" must be 0 or 1' );
              fCurrModule.OptionBase := tokenInt and 1; // ensure only 0 or 1
              GetToken;
            end
          else
        if CheckIdent( 'Explicit' )
          then optExplicit := true
          else
        if CheckIdent( 'Compare' )
          then
            begin
              if SameText( tokenText, 'Binary' )
                then optCompare := ocBinary
                else
              if SameText( tokenText, 'Text' )
                then optCompare := ocText
                else
              if SameText( tokenText, 'Database' )
                then optCompare := ocDatabase
                else Error( 'invalid "Option Compare"' );
              GetToken;
            end
          else Error( 'invalid "Option"' );
    end;

  // DefBool
  // DefByte
  // DefCur
  // DefDate
  // DefDbl
  // DefInt
  // DefLng
  // DefObj
  // DefSng
  // DefStr
  // DefVar
  procedure TvbParser.ParseDefStmt;
    var
      typ    : IvbType;
      suffix : TvbTypeSfx;
      fromc  : char;
      toc    : char;
    begin
      with fToken do
        begin
          typ := TokenEqvTable[tokenKind].eqvSimpleType;
          GetToken;
          repeat
            fromc := GetIdent( suffix, false )[1];      // take only first char
            if CheckToken( tkMinus )
              then toc := GetIdent( suffix, false )[1]  // take only first char
              else toc := fromc;
            fModuleOpts.DefineType( fromc, toc, typ );
          until not CheckToken( tkComma );
          NeedToken( tkEOL );
        end;
    end;

  // [Public | Private | Friend | Static | WithEvents | Dim | Const | Global]
  procedure TvbParser.ParseDeclFlags;
    begin
      fCurrNodeFlags := [];
      with fToken do
        while tokenKind in
          [tkFriend,
           tkPrivate,
           tkPublic,
           tkStatic,
           tkWithEvents,
           tkGlobal,
           tkDim,
           tkConst] do
          begin
            fCurrNodeFlags := fCurrNodeFlags + TokenEqvTable[tokenKind].eqvDeclFlags;
            GetToken;
          end;
    end;

  // [Optional] [ByVal | ByRef] [ParamArray] ident[( )] [As type] [= defaultvalue]
  function TvbParser.ParseParamDef : TvbParamDef;
    var
      flags   : TvbNodeFlags;
      sfx     : TvbTypeSfx;
      isArray : boolean;
      pt      : IvbType;
    begin
      isArray := false;
      flags   := [];

      with fToken do
        repeat
          case tokenKind of
            tkByRef :
              begin
                assert( not ( pfByVal in flags ) );
                assert( not ( pfByRef in flags ) );
                Include( flags, pfByRef );
              end;
            tkByVal :
              begin
                assert( not ( pfByVal in flags ) );
                assert( not ( pfByRef in flags ) );
                Include( flags, pfByVal );
              end;
            tkOptional :
              begin
                assert( not ( pfOptional in flags ) );
                Include( flags, pfOptional );
              end;
            tkParamArray :
              begin
                assert( not ( pfByVal in flags ) );
                assert( not ( pfByRef in flags ) );
                assert( not ( pfOptional in flags ) );
                flags := [pfParamArray];
                GetToken;
                break;
              end;
            else
              begin
                // Parameters are ByRef by default.
                if not ( pfByVal in flags ) and not fOptions.optParamDefaultByVal
                  then Include( flags, pfByRef );
                break;
              end;
          end;
          GetToken;
        until false;

      result := TvbParamDef.Create( flags, GetIdent( sfx ) );
      with result do
        try
          if pfParamArray in flags
            then
              begin
                NeedToken( tkOParen );
                NeedToken( tkCParen );
                if CheckToken( tkAs )
                  then NeedToken( tkVariant );
                ParamType := vbVariantArrayType;
              end
            else
              begin
                if CheckToken( tkOParen )
                  then
                    begin
                      isArray := true;
                      NeedToken( tkCParen );
                    end;

                if sfx <> NO_SFX
                  then pt := TypeSfx2Type[sfx]
                  else
                    if CheckToken( tkAs )  // explicit or default type?
                      then pt := ParseBaseType
                      else pt := fModuleOpts.GetDefinedType( Name );

                if CheckToken( tkEq )
                  then DefaultValue := ParseExpr;

                ParamType := CreateType( pt, isArray );
              end;
        except
          Free;
          raise;
        end;
    end;

  // ident[([[lower To] upper [, [lower To] upper] ..])] [As [New] type]
  function TvbParser.ParseVarDef( AllowKeyword, MustBeTyped : boolean ) : TvbVarDef;
    var
      bounds  : TvbNodeList;
      expr    : TvbExpr;
      isArray : boolean;
      sfx     : TvbTypeSfx;
      vt      : IvbType;
    begin
      bounds := nil;
      isArray := false;
      
      // Variables are private by default.
      // I'm refering here to variables declared at the module level which
      // become global(public or private) variables in standard modules and
      // get/set properties in object modules.
      if not ( dfPublic in fCurrNodeFlags )
        then Include( fCurrNodeFlags, dfPrivate );

      // Variables declard inside a static function are also static.
      if ( fCurrFunc <> nil ) and ( dfStatic in fCurrFunc.NodeFlags )
        then Include( fCurrNodeFlags, dfStatic );

      result := TvbVarDef.Create( fCurrNodeFlags, GetIdent( sfx, AllowKeyword ) );
      with result do
        try
          TopCmts := StrDynArrayFromStrList( fCurrTopCmts );
          if CheckToken( tkOParen ) // is this an array?
            then
              begin
                isArray := true;
                if not CheckToken( tkCParen ) // with dimensions?
                  then
                    begin
                      bounds := TvbNodeList.Create;
                      try
                        repeat
                          // Get the first dimension.
                          expr := ParseExpr;
                          try
                            if CheckToken( tkTo )
                              then bounds.Add( TvbArrayBounds.Create( expr, ParseExpr ) )
                              else bounds.Add( TvbArrayBounds.Create( nil, expr ) );
                            expr := nil;
                          except
                            expr.Free;
                            raise;
                          end;
                        until not CheckToken( tkComma );
                        NeedToken( tkCParen );
                      except
                        bounds.Free;
                        raise;
                      end;
                    end;
              end;

          if sfx <> NO_SFX
            then vt := TypeSfx2Type[sfx]
            else
              begin
                if MustBeTyped
                  then NeedToken( tkAs );
                if MustBeTyped or CheckToken( tkAs )
                  then
                    begin
                      if CheckToken( tkNew )
                        then NodeFlags := NodeFlags + [dfNew];
                      vt := ParseBaseType;
                    end
                  else vt := fModuleOpts.GetDefinedType( Name );
              end;

          VarType := CreateType( vt, isArray, bounds );

          TopRightCmt := CheckRightCmt;
        except
          Free;
          raise;
        end;
    end;

  // Const ident[As type] = expr
  //
  function TvbParser.ParseConstDef( AllowKeyword : boolean ) : TvbConstDef;
    var
      sfx : TvbTypeSfx;
    begin
      if not ( dfPrivate in fCurrNodeFlags )
        then Include( fCurrNodeFlags, dfPublic );

      result := TvbConstDef.Create( fCurrNodeFlags );
      with result do
        try
          TopCmts := StrDynArrayFromStrList( fCurrTopCmts );
          Name := GetIdent( sfx, AllowKeyword );
          if sfx <> NO_SFX
            then ConstType := TypeSfx2Type[sfx]
            else
          if CheckToken( tkAs )
            then ConstType := ParseBaseType;
          NeedToken( tkEq );
          Value       := ParseExpr;
          TopRightCmt := CheckRightCmt;
        except
          Free;
          raise;
        end;
    end;

  // Function | Sub ident [(arglist)] [As type[()]]
  //   [statements]
  // End Function | Sub
  procedure TvbParser.ParseFunctionDef;
    var
      termTokens : TvbTokenKinds;
      sfx        : TvbTypeSfx;
      isFunc     : boolean;
      prevScope  : TvbScope;
      rt         : IvbType;
      isArray    : boolean;
      name       : string;
    begin
      isFunc        := false;
      isArray       := false;
      fTempVarCount := 0;
      // Functions are public by default.
      if not ( dfPrivate in fCurrNodeFlags )
        then Include( fCurrNodeFlags, dfPublic );
      if fToken.tokenKind = tkSub
        then termTokens := [_tkEndSub]
        else
          begin
            isFunc := true;
            termTokens := [_tkEndFunction];
          end;

      GetToken;
      name := GetIdent( sfx );

      fCurrFunc  := fCurrModule.AddMethod( TvbFuncDef.Create( fCurrNodeFlags, name ) );
      prevScope  := fCurrScope;
      fCurrScope := fCurrFunc.Scope;

      with fCurrFunc do
        begin
          //writeln( Name );
          if CheckToken( tkOParen ) and not CheckToken( tkCParen ) // function takes arguemnts?
            then
              begin
                repeat
                  AddParam( ParseParamDef );
                until not CheckToken( tkComma );
                NeedToken( tkCParen );
              end;

          if isFunc
            then
              begin
                fCurrFuncName := Name;
                if sfx <> NO_SFX  // return type specified as sfx in function name?
                  then rt := TypeSfx2Type[sfx]
                  else
                if CheckToken( tkAs )  // explicit return type?
                  then
                    begin
                      rt := ParseBaseType;
                      if CheckToken( tkOParen) // returns an array?
                        then
                          begin
                            isArray := true;
                            NeedToken( tkCParen );
                          end;
                    end
                  else rt := fModuleOpts.GetDefinedType( Name );
                ReturnType := CreateType( rt, isArray );
              end;

          TopRightCmt := CheckRightCmt;
          NeedToken( tkEOL );
          SkipVBAttrs;

          CheckMoveCommentsToTop;
          
          TopCmts := StrDynArrayFromStrList( fCurrTopCmts );
          fCurrTopCmts.Clear;
          
          StmtBlock := ParseStmtBlock( termTokens );
          fCurrFunc := nil;
          if isFunc
            then NeedToken( _tkEndFunction )
            else NeedToken( _tkEndSub );
          DownRightCmt := CheckRightCmt;
        end;
      fCurrScope := prevScope;
      fCurrFuncName := '';
    end;

  // Property Set | Get | Let name [([arglist,] reference)] [As type]
  //   [statements]
  // End Property
  procedure TvbParser.ParsePropertyFuncDef;
    var
      sfx       : TvbTypeSfx;
      prevScope : TvbScope;
      isArray   : boolean;
      rt        : IvbType;
      name      : string;
    begin
      isArray       := false;
      fTempVarCount := 0;

      if not ( dfPrivate in fCurrNodeFlags )
        then Include( fCurrNodeFlags, dfPublic );

      if fCurrModule.NodeKind = STD_MODULE_DEF
        then Include( fCurrNodeFlags, mapAsFunctionCall );

      GetToken;
      case fToken.tokenKind of
        tkSet : fCurrNodeFlags := fCurrNodeFlags + [dfPropSet];
        tkGet : fCurrNodeFlags := fCurrNodeFlags + [dfPropGet];
        tkLet : fCurrNodeFlags := fCurrNodeFlags + [dfPropLet];
        else Error( '"Set", "Get" or "Let" expected' );
      end;
      GetToken;
      name := GetIdent( sfx );

      fCurrFunc := fPropertyMap.AddPropFunc( TvbFuncDef.Create( fCurrNodeFlags, name ) );
      // the parent scope of this function is the current scope.
      fCurrFunc.Scope.Parent := fCurrScope;
      // save the current scope to restore it later.
      prevScope  := fCurrScope;
      // set the current scope to the new function's scope.
      fCurrScope := fCurrFunc.Scope;

      with fCurrFunc do
        begin
          if CheckToken( tkOParen ) and not CheckToken( tkCParen )
            then
              begin
                repeat
                  AddParam( ParseParamDef );
                until not CheckToken( tkComma );
                NeedToken( tkCParen );
              end;

          if dfPropGet in NodeFlags
            then
              begin
                fCurrFuncName := Name;
                if sfx <> NO_SFX  // return type specified as sfx in property name?
                  then rt := TypeSfx2Type[sfx]
                  else
                if CheckToken( tkAs )  // explicit return type?
                  then
                    begin
                      rt := ParseBaseType;
                      if CheckToken( tkOParen) // returns an array?
                        then
                          begin
                            isArray := true;
                            NeedToken( tkCParen );
                          end;
                    end
                  else rt := fModuleOpts.GetDefinedType( Name );
                ReturnType := CreateType( rt, isArray );
              end;

          TopRightCmt := CheckRightCmt;
          NeedToken( tkEOL );
          SkipVBAttrs;

          CheckMoveCommentsToTop;
          TopCmts := StrDynArrayFromStrList( fCurrTopCmts );
          fCurrTopCmts.Clear;         

          StmtBlock := ParseStmtBlock( [_tkEndProperty] );
          fCurrFunc := nil;
          NeedToken( _tkEndProperty );
        end;
      fCurrScope := prevScope;
      fCurrFuncName := '';
    end;

  // Declare Sub name Lib "libName" [Alias "aliasName"] [([arglist])]
  // Declare Function name Lib "libName" [Alias "aliasName"] [([arglist])] [As type]
  function TvbParser.ParseDllFuncDef : TvbDllFuncDef;
    var
      sfx : TvbTypeSfx;
    begin
      // Declare statements are public by default.
      if not ( dfPrivate in fCurrNodeFlags )
        then Include( fCurrNodeFlags, dfPublic );
      result := TvbDllFuncDef.Create( fCurrNodeFlags );
      with result, fToken do
        try
          GetToken;
          if not CheckToken( tkSub ) and not CheckToken( tkFunction )
            then Error( '"Sub" or "Function" expected' );

          Name := GetIdent( sfx );
          NeedIdent( 'Lib' );
          CheckError( tokenKind = tkStringLit, 'DLL name expected' );
          DllName := tokenText;

          GetToken;
          if CheckIdent( 'Alias' )
            then
              begin
                CheckError( tokenKind = tkStringLit, 'function alias expected' );
                Alias := tokenText;
                GetToken;
              end;

          if CheckToken( tkOParen ) and not CheckToken( tkCParen )
            then
              begin
                repeat
                  AddParam( ParseParamDef );
                until not CheckToken( tkComma );
                NeedToken( tkCParen );
              end;

          if sfx <> NO_SFX
            then ReturnType := TypeSfx2Type[sfx]
            else
          if CheckToken( tkAs )
            then ReturnType := ParseBaseType;

          TopRightCmt := CheckRightCmt;
          NeedToken( tkEOL );
          SkipVBAttrs;

          CheckMoveCommentsToTop;
          TopCmts := StrDynArrayFromStrList( fCurrTopCmts );
        except
          Free;
          raise;
        end;
    end;

  // Type ident
  //   elementname [([subscripts])] As type
  //   [elementname [([subscripts])] As type]
  //   . . .
  // End Type
  function TvbParser.ParseRecordDef : TvbRecordDef;
    var
      sfx       : TvbTypeSfx;
      prevScope : TvbScope;
    begin
      // Records are public by default.
      if not ( dfPrivate in fCurrNodeFlags )
        then Include( fCurrNodeFlags, dfPublic );
      result     := TvbRecordDef.Create( fCurrNodeFlags );
      prevScope  := fCurrScope;
      fCurrScope := result.Scope;
      with result, fToken do
        try
          TopCmts := StrDynArrayFromStrList( fCurrTopCmts );
          GetToken;
          Name        := GetIdent( sfx );
          TopRightCmt := CheckRightCmt;
          NeedToken( tkEOL );
          SkipVBAttrs;
          fCurrTopCmts.Clear;

          repeat
            if tokenKind = tkEOL
              then GetToken
              else
            if tokenKind = tkComment
              then
                begin
                  fCurrTopCmts.Add( tokenText );
                  GetToken;
                end
              else
            if CheckToken( tkEnd )
              then
                begin
                  NeedToken( tkType );
                  break;
                end
              else
                begin
                  AddField( ParseVarDef( true, true ) );
                  NeedToken( tkEOL );
                  SkipVBAttrs;
                  fCurrTopCmts.Clear;
                end;
          until false;

          DownCmts     := StrDynArrayFromStrList( fCurrTopCmts );
          DownRightCmt := CheckRightCmt;
        except
          Free;
          raise;
        end;
      fCurrScope := prevScope;
    end;

  // Enum name
  //   membername [= constantexpression]
  //   membername [= constantexpression]
  //   . . .
  // End Enum
  function TvbParser.ParseEnumDef : TvbEnumDef;
    var
      sfx       : TvbTypeSfx;
      prevScope : TvbScope;

    function ParseEnumConst : TvbConstDef;
      begin
        result := TvbConstDef.Create( [dfPublic], GetIdent( sfx ), vbLongType );
        with result do
          try
            if CheckToken( tkEq )
              then Value := ParseExpr;
            TopCmts     := StrDynArrayFromStrList( fCurrTopCmts );
            TopRightCmt := CheckRightCmt;
            NeedToken( tkEOL );
            SkipVBAttrs;
          except
            result.Free;
            raise;
          end;
      end;

    begin
      if not ( dfPrivate in fCurrNodeFlags )
        then Include( fCurrNodeFlags, dfPublic );
      result     := TvbEnumDef.Create( fCurrNodeFlags );
      prevScope  := fCurrScope;
      fCurrScope := result.Scope;
      with result, fToken do
        try
          TopCmts := StrDynArrayFromStrList( fCurrTopCmts );
          GetToken;
          Name := GetIdent( sfx );
          TopRightCmt:= CheckRightCmt;
          NeedToken( tkEOL );
          SkipVBAttrs;
          fCurrTopCmts.Clear;
          repeat
            if tokenKind = tkEOL
              then GetToken
              else
            if tokenKind = tkComment
              then
                begin
                  fCurrTopCmts.Add( tokenText );
                  GetToken;
                end
              else
            if CheckToken( tkEnd )
              then
                begin
                  NeedToken( tkEnum );
                  CheckError( EnumCount > 0, 'enumerated type without enumerated constants' );
                  break;
                end
              else
                begin
                  AddEnum( ParseEnumConst );
                  fCurrTopCmts.Clear;
                end;
          until false;
          DownCmts     := StrDynArrayFromStrList( fCurrTopCmts );
          DownRightCmt := CheckRightCmt;
        except
          result.Free;
          raise;
        end;
      fCurrScope := prevScope;
    end;

  function TvbParser.ParseExpr( parseNameArgs : boolean ) : TvbExpr;

    function ParseUnaryExpr : TvbExpr;
      var
        tok : TvbTokenKind;
      begin
        if CheckToken( [tkPlus, tkMinus, tkNot], tok )
          then result := TvbUnaryExpr.Create( TokenEqvTable[tok].eqvOperator, ParseTermExpr )
          else
        // NOTE: VB6 supports AddressOf only in function arguments.
        if CheckToken( tkAddressOf )
          then result := TvbAddressOfExpr.Create( ParseName )
          else
        if CheckToken( tkNew )
          then result := TvbNewExpr.Create( ParseTermExpr )
          else result := ParseTermExpr( parseNameArgs );
        assert( result <> nil );
      end;


    function ParseMultExpr : TvbExpr;
      var
        tok : TvbTokenKind;
      begin
        result := ParseUnaryExpr;          
        assert( result <> nil );
        try
          while CheckToken( [tkMult, tkIntDiv, tkFloatDiv, tkMod, tkPower], tok ) do
            result := TvbBinaryExpr.Create(
              TokenEqvTable[tok].eqvOperator, result, ParseUnaryExpr );
        except
          result.Free;
          raise;
        end;
        {$IFDEF DEBUG}
        if result is TvbBinaryExpr
          then
            begin
              assert( TvbBinaryExpr( result ).LeftExpr <> nil );
              assert( TvbBinaryExpr( result ).RightExpr <> nil );
            end;          
        {$ENDIF}
      end;

    function ParseAddExpr : TvbExpr;
      var
        tok : TvbTokenKind;
      begin
        result := ParseMultExpr;
        assert( result <> nil );
        try
          while CheckToken( [tkPlus, tkMinus, tkConcat], tok ) do
            result := TvbBinaryExpr.Create(
              TokenEqvTable[tok].eqvOperator, result, ParseMultExpr );
        except
          result.Free;
          raise;
        end;
        {$IFDEF DEBUG}
        if result is TvbBinaryExpr
          then
            begin
              assert( TvbBinaryExpr( result ).LeftExpr <> nil );
              assert( TvbBinaryExpr( result ).RightExpr <> nil );
            end;
        {$ENDIF}
      end;

    function ParseRelationalExpr : TvbExpr;
      var
        tok : TvbTokenKind;
        putParens : boolean;
      begin
        putParens := false;
        result    := ParseAddExpr;
        assert( result <> nil );
        try
          while CheckToken( [tkLess,
                             tkLessEq,
                             tkGreater,
                             tkGreaterEq,
                             tkEq,
                             tkNotEq,
                             tkIs,
                             tkLike], tok ) do
            begin
              putParens := true;
              result    := TvbBinaryExpr.Create( TokenEqvTable[tok].eqvOperator, result, ParseAddExpr );
            end;
          if putParens
            then result.Parenthesized := true;
        except
          result.Free;
          raise;
        end;
        {$IFDEF DEBUG}
        if result is TvbBinaryExpr
          then
            begin
              assert( TvbBinaryExpr( result ).LeftExpr <> nil );
              assert( TvbBinaryExpr( result ).RightExpr <> nil );
            end;          
        {$ENDIF}
      end;

    function ParseLogicalExpr : TvbExpr;
      var
        tok : TvbTokenKind;
      begin
        result := ParseRelationalExpr;
        assert( result <> nil );
        try
          while CheckToken( [tkAnd, tkOr, tkXor, tkEqv, tkImp], tok ) do
            result := TvbBinaryExpr.Create(
              TokenEqvTable[tok].eqvOperator, result, ParseRelationalExpr );
        except
          result.Free;
          raise;
        end;
        {$IFDEF DEBUG}
        if result is TvbBinaryExpr
          then
            begin
              assert( TvbBinaryExpr( result ).LeftExpr <> nil );
              assert( TvbBinaryExpr( result ).RightExpr <> nil );
            end;          
        {$ENDIF}
      end;

    begin
      result := ParseLogicalExpr;
    end;

  function TvbParser.ParseTermExpr( parseNamedArgs : boolean ) : TvbExpr;
    var
      typ       : TvbName;
      obj       : TvbExpr;
      callOrIdx : TvbCallOrIndexerExpr absolute result;
      ident     : string;

    function ParseExprInParens : TvbExpr;
      begin
        GetToken;
        NeedToken( tkOParen );
        result := ParseExpr;
        assert( result <> nil );
        try
          NeedToken( tkCParen );
        except
          result.Free;
          raise;
        end;
      end;

    function ParseStringExpr : TvbStringExpr;
      begin
        GetToken;
        NeedToken( tkOParen );
        result := TvbStringExpr.Create( ParseExpr );
        with result do
          try
            NeedToken( tkComma );
            Char := ParseExpr;
            NeedToken( tkCParen );
          except
            Free;
            raise;
          end;
      end;

    function ParseArrayExpr : TvbArrayExpr;
      begin
        GetToken;
        NeedToken( tkOParen );
        result := TvbArrayExpr.Create;
        with result do
          try
            if not CheckToken( tkCParen )
              then
                repeat
                  AddValue( ParseExpr );
                  if CheckToken( tkCParen )
                    then break;
                  NeedToken( tkComma );
                until false;
          except
            Free;
            raise;
          end;  
      end;

    function ParseInputExpr : TvbInputExpr;
      begin
        GetToken;
        NeedToken( tkOParen );
        result := TvbInputExpr.Create;
        with result do
          try
            CharCount := ParseExpr;
            NeedToken( tkComma );
            CheckToken( tkHash );
            FileNum := ParseExpr;
            NeedToken( tkCParen );
          except
            Free;
            raise;
          end;
      end;

    begin
      result := nil;
      try
        with fToken do
          begin
            case tokenKind of
              tkIntLit :
                begin
                  result := TvbIntLit.Create( tokenInt );
                  if TF_HEX in tokenFlags
                    then result.NodeFlags := result.NodeFlags + [int_Hex]
                    else
                  if TF_OCTAL in tokenFlags
                    then result.NodeFlags := result.NodeFlags + [int_Octal];
                  if TF_LONG_SFX in tokenFlags
                    then result.NodeFlags := result.NodeFlags + [sfx_Long];
                  GetToken;
                end;
              tkFloatLit :
                begin
                  result := TvbFloatLit.Create( tokenFloat );
                  if TF_SINGLE_SFX in tokenFlags
                    then result.NodeFlags := result.NodeFlags + [sfx_Single]
                    else
                  if TF_CURRENCY_SFX in tokenFlags
                    then result.NodeFlags := result.NodeFlags + [sfx_Curr];
                  GetToken;
                end;
              tkStringLit :
                begin
                  result := TvbStringLit.Create( tokenText );
                  GetToken;
                end;
              tkTrue,
              tkFalse :
                begin
                  result := TvbBoolLit.Create( tokenKind = tkTrue );
                  GetToken;
                end;
              tkNothing :
                begin
                  result := TvbNothingLit.Create;
                  GetToken;
                end;
              tkMe :
                begin
                  result := TvbMeExpr.Create( fTypeMap.GetType( fCurrModule.Name ) );
                  GetToken;
                end;
              tkDateLit :
                begin
                  result := TvbDateLit.Create( tokenDate );
                  GetToken;
                end;
              tkTypeOf: // 'TypeOf' objectName 'Is' typeName
                begin
                  GetToken;
                  typ := nil;
                  obj := ParseTermExpr;
                  try
                    NeedToken( tkIs );
                    typ    := ParseName;
                    result := TvbTypeOfExpr.Create( obj, typ );
                  except
                    obj.Free;
                    typ.Free;
                    raise;
                  end;
                end;
              tkCBool  : result := TvbCastExpr.Create( ParseExprInParens, BOOLEAN_CAST );
              tkCByte  : result := TvbCastExpr.Create( ParseExprInParens, BYTE_CAST );
              tkCDate  : result := TvbCastExpr.Create( ParseExprInParens, DATE_CAST );
              tkCDbl   : result := TvbCastExpr.Create( ParseExprInParens, DOUBLE_CAST );
              tkCDec   : result := TvbCastExpr.Create( ParseExprInParens, DECIMAL_CAST );
              tkCInt   : result := TvbCastExpr.Create( ParseExprInParens, INTEGER_CAST );
              tkCLng   : result := TvbCastExpr.Create( ParseExprInParens, LONG_CAST );
              tkCSng   : result := TvbCastExpr.Create( ParseExprInParens, SINGLE_CAST );
              tkCStr   : result := TvbCastExpr.Create( ParseExprInParens, STRING_CAST );
              tkCVar   : result := TvbCastExpr.Create( ParseExprInParens, VARIANT_CAST );
              tkCVErr  : result := TvbCastExpr.Create( ParseExprInParens, VERROR_CAST );
              tkLBound :
                begin
                  GetToken;
                  NeedToken( tkOParen );
                  result := TvbLBoundExpr.Create( ParseExpr );
                  try
                    if CheckToken( tkComma )
                      then TvbLBoundExpr( result ).Dim := ParseExpr;
                    NeedToken( tkCParen );
                  except
                    FreeAndNil( result );
                    raise;
                  end;
                end;
              tkUBound :
                begin
                  GetToken;
                  NeedToken( tkOParen );
                  result := TvbUBoundExpr.Create( ParseExpr );
                  try
                    if CheckToken( tkComma )
                      then TvbUBoundExpr( result ).Dim := ParseExpr;
                    NeedToken( tkCParen );
                  except
                    FreeAndNil( result );
                    raise;
                  end;
                end;
              tkString : result := ParseStringExpr;
              tkInt    : result := TvbIntExpr.Create( ParseExprInParens );
              tkFix    : result := TvbFixExpr.Create( ParseExprInParens );
              tkAbs    : result := TvbAbsExpr.Create( ParseExprInParens );
              tkLen    : result := TvbLenExpr.Create( ParseExprInParens );
              tkLenB   : result := TvbLenBExpr.Create( ParseExprInParens );
              tkSgn    : result := TvbSgnExpr.Create( ParseExprInParens );
              tkSeek   : result := TvbSeekExpr.Create( ParseExprInParens );
              tkArray  : result := ParseArrayExpr;
              tkDate :
                begin
                  result := TvbDateExpr.Create;
                  GetToken;
                end;
              tkDoEvents :
                begin
                  GetToken;
                  if CheckToken( tkOParen )
                    then NeedToken( tkCParen );
                  result := TvbDoEventsExpr.Create;
                end;
              tkIdent :
                begin
                  ident := tokenText;
                  GetToken;
//                  if TF_STRING_SFX in tokenFlags
//                    then s := s + '$';
//                  if SameText( tokenText, 'Mid$' )
//                    then result := ParseMidExpr

                  // if we've got the current function name and it's not
                  // followed by paranthesis, this expression is refering to the
                  // function's result value.
                  if SameText( ident, fCurrFuncName ) and ( tokenKind <> tkOParen )
                    then result := TvbFuncResultExpr.Create( fCurrFunc.ReturnType )
                    else
                  // if we were told to parse named arguments and a ':=' follows,
                  // parse a named argument expression. 
                  if parseNamedArgs and CheckToken( tkColonEq )
                    then
                      begin
                        if tokenKind in [tkByVal, tkByRef]
                          then GetToken; // TODO: add the flags to the expression
                        result := TvbNamedArgExpr.Create( ident, ParseExpr )
                      end
                    else
                      // otherwise this is simply a name expression.
                      result := TvbNameExpr.Create( ident );
                end;
              tkInput,
              tkInputB : result := ParseInputExpr;
              tkOParen : // parenthesized expression '(' expr ')'
                begin
                  GetToken;
                  result := ParseExpr;
                  result.Parenthesized := true;
                  NeedToken( tkCParen );
                end;
            end;

            while true do
              if CheckToken( tkDot )
                then
                  begin
                    if not ( tokenKind in [tkIdent, tkAbs..tkXor] )
                      then Error( 'member name expected' );
                    if result = nil
                      then result := TvbMemberAccessExpr.Create( tokenText, nil, fCurrWithExpr )
                      else result := TvbMemberAccessExpr.Create( tokenText, result );
                    GetToken;
                  end
                else
              if CheckToken( tkBang )  // dictionary access operator
                then
                  begin
                    if not ( tokenKind in [tkIdent, tkAbs..tkXor] )
                      then Error( 'name expected' );
                    if result = nil
                      then result := TvbDictAccessExpr.Create( tokenText, nil, fCurrWithExpr )
                      else result := TvbDictAccessExpr.Create( tokenText, result );
                    GetToken;
                  end
                else
              if ( result <> nil ) and CheckToken( tkOParen ) // method call or array indexer
                then
                  begin
                    if not CheckToken( tkCParen )
                      then
                        begin
                          result := TvbCallOrIndexerExpr.Create( result, ParseArgList );
                          result.IsUsedWithArgs := true;
                          NeedToken( tkCParen );
                        end
                      else
                        // since the target is used with empty parenthesis we
                        // indicate the expression is refering to a function.
                        result := TvbCallOrIndexerExpr.Create( result, nil, true );
                  end
                else break;
          end;
          assert( result <> nil );
      except
        result.Free;
        raise;
      end;
    end;

  // Event procedurename [(arglist)]
  function TvbParser.ParseEventDef : TvbEventDef;
    var
      sfx : TvbTypeSfx;
    begin
      result := TvbEventDef.Create;
      with result do
        try
          TopCmts := StrDynArrayFromStrList( fCurrTopCmts );
          GetToken;
          Name := GetIdent( sfx );

          if CheckToken( tkOParen ) and not CheckToken( tkCParen )
            then
              begin
                repeat
                  AddParam( ParseParamDef );
                until not CheckToken( tkComma );
                NeedToken( tkCParen );
              end;

          TopRightCmt := CheckRightCmt;
          NeedToken( tkEOL );
          SkipVBAttrs;
        except
          Free;
          raise;
        end;
    end;

  // Parses the "Width" statement.
  // The caller already consumed "Width #".
  procedure TvbParser.ParseWidthStmt;
    begin
      with fCurrBlock.AddWidthStmt do
        begin
          TopCmts    := StrDynArrayFromStrList( fCurrTopCmts );
          FrontLabel := fFrontLabel;
          FileNum := ParseExpr;
          NeedToken( tkComma );
          Width := ParseExpr;
          TopRightCmt := CheckRightCmt;
        end
    end;

  procedure TvbParser.ParseErrorStmt;
    begin
      with fCurrBlock.AddErrorStmt do
        begin
          TopCmts     := StrDynArrayFromStrList( fCurrTopCmts );
          FrontLabel  := fFrontLabel;
          ErrorNumber := ParseExpr;
          TopRightCmt := CheckRightCmt;
        end
    end;

  procedure TvbParser.ParseNameStmt;
    begin
      with fCurrBlock.AddNameStmt do
        begin
          TopCmts    := StrDynArrayFromStrList( fCurrTopCmts );
          FrontLabel := fFrontLabel;
          OldPath := ParseExpr;
          NeedToken( tkAs );
          NewPath := ParseExpr;
          TopRightCmt := CheckRightCmt;
        end;
    end;

  // Circle [Step](x, y), radius[, [color], [start], [end] [, aspect]]
  procedure TvbParser.ParseCircleStmt( obj : TvbExpr = nil );
    begin
      with fCurrBlock.AddCircleStmt do
        begin
          TopCmts    := StrDynArrayFromStrList( fCurrTopCmts );
          FrontLabel := fFrontLabel;
          ObjectRef := obj;
          Step := CheckIdent( 'Step' );
          NeedToken( tkOParen );
          X := ParseExpr;
          NeedToken( tkComma );
          Y := ParseExpr;
          NeedToken( tkCParen );
          NeedToken( tkComma );
          R := ParseExpr;
          if CheckToken( tkComma )
            then
              with fToken do
                begin
                  if tokenKind <> tkComma
                    then Color := ParseExpr;
                  if CheckToken( tkComma )
                    then
                      begin
                        if tokenKind <> tkComma
                          then StartAngle := ParseExpr;
                        if CheckToken( tkComma )
                          then
                            begin
                              if tokenKind <> tkComma
                                then EndAngle := ParseExpr;
                              if CheckToken( tkComma )
                                then Aspect := ParseExpr;
                            end;
                      end;
                end;
          TopRightCmt := CheckRightCmt;
        end;
    end;

  // Line [(x1, y1 )]-[Step](x2, y2)[, [color] [,B | BF]]
  procedure TvbParser.ParseLineStmt( obj : TvbExpr = nil );
    var
      BF : string;
    begin
      //NOTE: i don't know the defaults for boxed and filled.
      with fCurrBlock.AddLineStmt do
        begin
          TopCmts    := StrDynArrayFromStrList( fCurrTopCmts );
          FrontLabel := fFrontLabel;
          ObjectRef := obj;
          if CheckToken( tkOParen )
            then
              begin
                X1 := ParseExpr;
                NeedToken( tkComma );
                Y1 := ParseExpr;
                NeedToken( tkCParen );
              end;
          NeedToken( tkMinus );
          Step := CheckIdent( 'Step' );
          NeedToken( tkOParen );
          X2 := ParseExpr;
          NeedToken( tkComma );
          Y2 := ParseExpr;
          NeedToken( tkCParen );
          if CheckToken( tkComma )
            then
              begin
                if fToken.tokenKind <> tkComma
                  then Color := ParseExpr;
                if CheckToken( tkComma )
                  then
                    begin
                      BF := uppercase( fToken.tokenText );
                      GetToken;
                      if BF = 'B'
                        then
                          begin
                            Boxed  := true;
                            Filled := false;
                          end
                        else
                      if BF = 'BF'
                        then
                          begin
                            Boxed  := true;
                            Filled := true;
                          end
                        else Error( 'expected "B" or "BF"' );
                    end;
              end;
          TopRightCmt := CheckRightCmt;
        end;
    end;

  // PSet (x, y)[, color]
  procedure TvbParser.ParsePSetStmt( obj : TvbExpr = nil );
    begin
      with fCurrBlock.AddPSetStmt do
        begin
          TopCmts    := StrDynArrayFromStrList( fCurrTopCmts );
          FrontLabel := fFrontLabel;
          ObjectRef := obj;
          NeedToken( tkOParen );
          X := ParseExpr;
          NeedToken( tkComma );
          Y := ParseExpr;
          NeedToken( tkCParen );
          if CheckToken( tkComma )
            then Color := ParseExpr;
          TopRightCmt := CheckRightCmt;
        end;
    end;

  procedure TvbParser.ParseLineInputStmt;
    begin
      with fCurrBlock.AddLineInputStmt do
        begin
          TopCmts := StrDynArrayFromStrList( fCurrTopCmts );
          FrontLabel := fFrontLabel;
          CheckToken( tkHash );
          FileNum := ParseExpr;
          NeedToken( tkComma );
          VarName     := ParseExpr;
          TopRightCmt := CheckRightCmt;
        end;
    end;
  // Parses a block of statements until we meet a token from the TermTokens set.
  // Always returns a new statement block object even if it's empty.
  // If there are hanging comments at the end of the block they are stored
  // in the statement block's DownCmts property.
  // If there a hanging labels at the end of the block, a TvbLabelDefStmt is
  // added to the block to capture those labels.
  function TvbParser.ParseStmtBlock( TermTokens : TvbTokenKinds ) : TvbStmtBlock;
    var
      skipTokens    : TvbTokenKinds;
      prevStmtBlock : TvbStmtBlock;
      lab           : TvbLabelDef;

    procedure AddStmt( stmt : TvbStmt;
                       collectTopCmts     : boolean = true;
                       collectTopRightCmt : boolean = true );
      begin
        fCurrBlock.Add( stmt );
        if collectTopCmts
          then stmt.TopCmts := StrDynArrayFromStrList( fCurrTopCmts );
        if collectTopRightCmt
          then stmt.TopRightCmt := CheckRightCmt;
        stmt.FrontLabel := fFrontLabel;
      end;

    procedure ParseAssignOrCallStmt;
      const
        sCircle = 'Circle';
        sError  = 'Error';
        sLine   = 'Line';
        sMid    = 'Mid';
        sName   = 'Name';
        sPrint  = 'Print';
        sPSet   = 'PSet';
        sWidth  = 'Width';
      var
        expr : TvbExpr;
        args : TvbArgList;
      begin
        expr := nil;
        args := nil;
        with fToken do
          try
            // first check for the special/legacy statements that start
            // with identifiers.
            if tokenKind = tkIdent
              then
                // first check for the current function name.
                if SameText( tokenText, fCurrFuncName )
                  then
                    begin
                      GetToken;
                      if tokenKind <> tkOParen
                        then expr := TvbFuncResultExpr.Create( fCurrFunc.ReturnType )
                        else expr := TvbNameExpr.Create( tokenText )
                    end
                  else
                // check for "Width" statement.
                if SameText( tokenText, sWidth )
                  then
                    begin
                      GetToken;
                      if CheckToken( tkHash )
                        then
                          begin
                            ParseWidthStmt;
                            exit;
                          end
                        else expr := TvbNameExpr.Create( sWidth );
                    end
                  else
                if SameText( tokenText, sMid )
                  then
                    begin
                      GetToken;
                      // if we have a parenthesis without leading whitespaces
                      // parse the "Mid" statement.
                      if not ( TF_SPACED in tokenFlags ) and CheckToken( tkOParen )
                        then
                          begin
                            ParseMidAssignStmt;
                            exit;
                          end
                        else expr := TvbNameExpr.Create( sMid );
                    end
                  else
                if SameText( tokenText, sError )
                  then
                    begin
                      GetToken;
                      if not ( tokenKind in [tkDot, tkBang, tkEq] )
                        then
                          begin
                            ParseErrorStmt;
                            exit;
                          end
                        else expr := TvbNameExpr.Create( sError );
                    end
                  else
                // check for the "Name" statement.
                if SameText( tokenText, sName )
                  then
                    begin
                      GetToken;
                      if not ( tokenKind in [tkDot, tkBang, tkEq] )
                        then
                          begin
                            ParseNameStmt;
                            exit;
                          end
                        else expr := TvbNameExpr.Create( sName );
                    end
                  else
                // check for the "Line" method or "Line Input" statement.
                if SameText( tokenText, sLine )
                  then
                    begin
                      GetToken;
                      if not ( tokenKind in [tkDot, tkBang, tkEq] )
                        then
                          begin
                            if CheckToken( tkInput )
                              then ParseLineInputStmt
                              else ParseLineStmt( expr );
                            exit;
                          end
                        else expr := TvbNameExpr.Create( sLine );
                    end
                  else
                    begin
                      expr := TvbNameExpr.Create( tokenText );  // simply an identifier
                      GetToken;
                    end
            else
          if CheckToken( tkMe )
            then expr := TvbMeExpr.Create( fTypeMap.GetType( fCurrModule.Name ) )
            else
          if CheckToken( tkDoEvents )
            then expr := TvbDoEventsExpr.Create;
            

            // this loop parses contiguous member-access, dictionary-access and
            // indexer expressions. e.g:
            //      foo.bar
            //      foo(arg1, arg2).bar!name
            //      foo!bar
            repeat
              // since '.', '!' and '(', can't be spaced in this context we
              // first make sure that the token is NOT SPACED.
              if ( expr <> nil ) and ( TF_SPACED in tokenFlags )
                then break;

              if CheckToken( tkDot )  // check for member-access
                then
                  begin
                    if not ( tokenKind in [tkIdent, tkAbs..tkXor] )
                      then Error( 'identifier expected' );
                    // check for legacy Line method syntax
                    if SameText( tokenText, sLine )
                      then
                        begin
                          GetToken;
                          if not ( tokenKind in [tkDot, tkBang, tkEq] )
                            then
                              begin
                                ParseLineStmt( expr );
                                exit;
                              end
                            else
                          if expr = nil
                            then expr := TvbMemberAccessExpr.Create( sLine, nil, fCurrWithExpr )
                            else expr := TvbMemberAccessExpr.Create( sLine, expr );
                        end
                      else
                    if CheckToken( tkPSet )
                      then
                        if not ( tokenKind in [tkDot, tkBang, tkEq] )
                          then
                            begin
                              ParsePSetStmt( expr );
                              exit;
                            end
                          else
                        if expr = nil
                          then expr := TvbMemberAccessExpr.Create( sPSet, nil, fCurrWithExpr )
                          else expr := TvbMemberAccessExpr.Create( sPSet, expr )
                      else
                    if CheckToken( tkCircle )
                      then
                        if not ( tokenKind in [tkDot, tkBang, tkEq] )
                          then
                            begin
                              ParseCircleStmt( expr );
                              exit;
                            end
                          else
                        if expr = nil
                          then expr := TvbMemberAccessExpr.Create( sCircle, nil, fCurrWithExpr )
                          else expr := TvbMemberAccessExpr.Create( sCircle, expr )
                      else
                    if CheckToken( tkPrint )
                      then
                        if not ( tokenKind in [tkDot, tkBang, tkEq] )
                          then
                            begin
                              ParsePrintMethod( expr );
                              exit;
                            end
                          else
                        if expr = nil
                          then expr := TvbMemberAccessExpr.Create( sPrint, nil, fCurrWithExpr )
                          else expr := TvbMemberAccessExpr.Create( sPrint, expr )
                      else
                        begin
                          if expr = nil
                            then expr := TvbMemberAccessExpr.Create( tokenText, nil, fCurrWithExpr )
                            else expr := TvbMemberAccessExpr.Create( tokenText, expr );
                          GetToken;
                        end;
                  end
                else
              if CheckToken( tkBang )   // check for a dictionary access
                then
                  begin
                    if not ( tokenKind in [tkIdent, tkAbs..tkXor] )
                      then Error( 'name expected' );
                    if expr = nil
                      then expr := TvbDictAccessExpr.Create( tokenText, nil, fCurrWithExpr )
                      else expr := TvbDictAccessExpr.Create( tokenText, expr );
                    GetToken;
                  end
                else
              // finally check for an indexer expression if we've already
              // parsed the left-side.
              if ( expr <> nil ) and CheckToken( tkOParen )
                then
                  begin
                    if not CheckToken( tkCParen )
                      then
                        begin
                          expr := TvbCallOrIndexerExpr.Create( expr, ParseArgList );
                          expr.IsUsedWithArgs := true;
                          NeedToken( tkCParen );
                        end
                      else
                        begin
                          // note that we create this call/indexer expression
                          // indicating that the target must be a function since
                          // it was used with "()".
                          expr := TvbCallOrIndexerExpr.Create( expr, nil, true );
                        end;
                  end
                else break;
            until false;

            // if at this point we haven't parsed an expression
            // the user provided something invalid.
            if expr = nil
              then Error( 'invaid statement' );

            if CheckToken( tkEq )  // an assignment?
              then
                begin
                  expr.IsRValue := true;
                  AddStmt( TvbAssignStmt.Create( expr, ParseExpr ) );
                end
              else  // otherwise this is a subroutine call
                begin
                  if not ( fToken.tokenKind in [tkEOL, tkColon, tkComment] )
                    then args := ParseArgList;
                  AddStmt( TvbCallStmt.Create( TvbCallOrIndexerExpr.Create( expr, args ) ) );
                end;
          except
            FreeAndNil( expr );
            FreeAndNil( args );
            raise;
          end;    
      end;

    procedure ParseStmt;
      begin
        with fToken do
          case tokenKind of
            tkIf : ParseIfStmt;
            tkFor :
              begin
                GetToken;
                if tokenKind = tkEach
                  then ParseForeachStmt
                  else ParseForStmt;
              end;
            tkWhile  : ParseWhileStmt;
            tkSelect : ParseSelectStmt;
            tkWith   : ParseWithStmt;
            tkCall   : ParseCallStmt;

            tkIdent,
            tkDot,
            tkDoEvents,
            tkMe : ParseAssignOrCallStmt;

            tkReDim  : ParseReDimStmt;
            tkDo     : ParseDoLoopStmt;
            tkGoSub,
            tkGoTo   : ParseGotoOrGosubStmt;
            tkResume : ParseResumeStmt;
            tkReturn : ParseReturnStmt;
            tkExit :
              begin
                GetToken;
                if tokenKind in [tkDo, tkFor]
                  then AddStmt( TvbExitLoopStmt.Create )
                  else
                if tokenKind in [tkSub, tkFunction, tkProperty]
                  then AddStmt( TvbExitStmt.Create )
                  else Error( 'expected "Do", "For", "Function", "Property", or "Sub"' );
                GetToken;
              end;
            tkOn :
              begin
                GetToken;
                if CheckToken( tkLocal ) or SameText( tokenText, 'Error' )
                  then ParseOnErrorStmt
                  else ParseOnGotoOrGosubStmt;
              end;

            tkRaiseEvent : ParseRaiseEventStmt;
            tkCircle :
              begin
                GetToken;
                ParseCircleStmt;
              end;
            tkPSet :
              begin
                GetToken;
                ParsePSetStmt;
              end;              
            tkEnd   : ParseEndStmt;
            tkInput : ParseInputStmt;
            tkPrint :
              begin
                GetToken;
                if CheckToken( tkHash )
                  then ParsePrintStmt
                  else ParsePrintMethod;
              end;
            tkOpen   : ParseOpenStmt;
            tkErase  : ParseEraseStmt;
            tkStop   : ParseStopStmt;
            tkGet    : ParseGetStmt;
            tkClose  : ParseCloseStmt;
            tkLock   : ParseLockStmt;
            tkUnlock : ParseUnlockStmt;
            tkPut    : ParsePutStmt;
            tkSeek   : ParseSeekStmt;
            tkWrite  : ParseWriteStmt;
            tkDate   : ParseDateStmt;
            tkSet    : ParseAssignStmt( asSet );
            tkLet    : ParseAssignStmt( asLet );
            tkLSet   : ParseAssignStmt( asLSet );
            tkRSet   : ParseAssignStmt( asRSet );
            tkDebug :
              begin
                // Debug isn't really a VB object. Debug.Assert and
                // Debug.Print are special language constructs.
                GetToken;
                NeedToken( tkDot );
                if CheckToken( tkPrint )
                  then ParseDebugPrintStmt
                  else
                if CheckIdent( 'Assert')
                  then ParseDebugAssertStmt
                  else Error( '"Assert" or "Print" expected' );
              end;
            else Error( 'invalid statement' );
          end;
      end;

    begin
      Include( TermTokens, _tkEOF );
      Include( TermTokens, tkComment );
      skipTokens    := [tkEOL, tkColon];
      skipTokens    := skipTokens - TermTokens * skipTokens;
      prevStmtBlock := fCurrBlock;
      fCurrTopCmts.Clear;
      fCurrBlock := TvbStmtBlock.Create;
      fFrontLabel   := nil;
      with fToken do
        try
          repeat
            // Check for a label declaration.
            if TF_LABEL in tokenFlags
              then
                begin
                  if tokenKind = tkIdent
                    then lab := TvbLabelDef.Create( tokenText )
                    else lab := TvbLabelDef.Create( IntToStr( tokenInt ) );
                  try
                    GetToken;

                    lab.TopCmts     := StrDynArrayFromStrList( fCurrTopCmts );
                    lab.TopRightCmt := CheckRightCmt;
                    fCurrTopCmts.Clear;

                    // Link this label to the previous one.
                    // This way we'll end up with a linked-list of all labels
                    // in front of an statement.
                    if fFrontLabel = nil
                      then
                        begin
                          fFrontLabel := lab;
                          fLastLabel  := lab;
                        end
                      else
                        begin
                          lab.PrevLabel        := fLastLabel;
                          fLastLabel.NextLabel := lab;
                          fLastLabel           := lab;
                        end;
                    fCurrFunc.AddLabel( lab );
                  except
                    FreeAndNil( lab );
                    raise;
                  end;
                end
              else
            if tokenKind in [tkDim, tkConst, tkStatic]
              then ParseLocalVarOrConst
              else
            if tokenKind in skipTokens
              then GetToken
              else
            if tokenKind = tkComment
              then
                begin
                  fCurrTopCmts.Add( tokenText );
                  GetToken;
                end
              else
            if tokenKind in TermTokens
              then break
              else
                begin
                  ParseStmt;
                  fCurrTopCmts.Clear;
                  fFrontLabel := nil;
                end;
          until false;

          // At this point, fCurrTopCmts has all the comments at the end of
          // the block. Now they become the DownCmts of the statement block.
          fCurrBlock.DownCmts := StrDynArrayFromStrList( fCurrTopCmts );
          fCurrTopCmts.Clear;

          // Check if there are labels at the end of the block that are
          // not bound to any statement. In this case, we add a label
          // declaration statement to capture those labels.
          if fFrontLabel <> nil
            then
              begin
                fCurrBlock.Add( TvbLabelDefStmt.Create( fFrontLabel ) );
                fFrontLabel := nil;
              end;
        except
          FreeAndNil( fCurrBlock );
          raise;
        end;
      result     := fCurrBlock;
      fCurrBlock := prevStmtBlock;
    end;

  procedure TvbParser.ParseDef;
    begin
      //Writeln( fToken.tokenText );
      with fCurrModule do
        begin
          ParseDeclFlags;
          case fToken.tokenKind of
            tkSub,
            tkFunction : ParseFunctionDef;
            tkDeclare  : AddDllFunc( ParseDllFuncDef );
            tkType     : AddRecord( ParseRecordDef );
            tkEnum     : AddEnum( ParseEnumDef );
            tkImplements :
              begin
                repeat
                  GetToken;
                until fToken.tokenKind in [tkEOL, _tkEOF];
              end;
            tkEvent    : AddEvent( ParseEventDef );
            tkProperty : ParsePropertyFuncDef;
            else
              if ( fToken.tokenKind = tkIdent ) and
                 ( fCurrNodeFlags * [dfGlobal, dfPublic, dfPrivate , dfStatic, dfDim, dfConst] <> [] )
                then
                  begin
                    repeat
                      if dfConst in fCurrNodeFlags
                        then AddConstant( ParseConstDef( false ) )
                        else AddVar( ParseVarDef( false, false ) );
                      fCurrTopCmts.Clear;                        
                    until not CheckToken( tkComma );
                    NeedToken( tkEOL );
                    SkipVBAttrs;
                  end
                else Error( 'invalid declaration' );
          end;
        end;
    end;

  function TvbParser.ParseStdModule( const Name, relPath : string; Source : TvbCustomTextBuffer ) : TvbModuleDef;
    var
      oldScope : TvbScope;
    begin
      DoParseModuleBegin( Name, relPath );

      fModuleOpts.SetDefaults;
      oldScope := fCurrScope;
      fCurrModule := TvbStdModuleDef.Create( Name, relPath, fCurrScope );
      fCurrModule.RelativePath := relPath;
      try
        fCurrScope := fCurrModule.Scope;
        fLexer.PushBuffer( Source );
        try
          GetToken;
          ParseModule;
        finally
          fLexer.PopBuffer;
        end;
      except
        fCurrModule.Free;
        raise;
      end;
      result     := fCurrModule;
      fCurrScope := oldScope;

      DoParseModuleEnd;
    end;

  function TvbParser.ParseArgList : TvbArgList;

    function ParseArg : TvbExpr;
      var
        flg : TvbNodeFlags;
      begin
        if CheckToken( tkByVal )
          then flg := [nfArgByVal]
          else
        if CheckToken( tkByRef )
          then flg := [nfArgByRef]
          else flg := [];
        result := ParseExpr( true );
        result.NodeFlags := result.NodeFlags + flg;
      end;

    begin
      result := TvbArgList.Create;
      with fToken do
        try
          repeat
            // collect omitted arguments
            while CheckToken( tkComma ) do
              result.Add( nil );            
            result.Add( ParseArg );
          until not CheckToken( tkComma );
        except
          result.Free;
          raise;
        end;
    end;

  // Call funcName
  procedure TvbParser.ParseCallStmt;
    begin
      GetToken;
      with fCurrBlock.AddCallStmt do
        begin
          Method := ParseTermExpr;
          TopCmts     := StrDynArrayFromStrList( fCurrTopCmts );
          TopRightCmt := CheckRightCmt;
          FrontLabel  := fFrontLabel;
        end;
    end;

  // If condition Then [statements] [Else elsestatements]
  // If condition Then
  //   [statements]
  // [ElseIf condition-n Then
  //   [elseifstatements] ...
  // [Else
  //   [elsestatements]]
  // End If
  procedure TvbParser.ParseIfStmt;
    var
      elif            : TvbIfStmt;
      elseTopRightCmt : string;
    begin
      elif := nil;
      GetToken;
      with fCurrBlock.AddIfStmt do
        begin
          TopCmts     := StrDynArrayFromStrList( fCurrTopCmts );
          FrontLabel  := fFrontLabel;
          Expr        := ParseExpr;
          NeedToken( tkThen );
          TopRightCmt := CheckRightCmt;
          if CheckToken( tkEOL )
            then
              begin
                ThenBlock := ParseStmtBlock( [tkElseIf, tkElse, _tkEndIf] );

                while CheckToken( tkElseIf ) do
                  with AddElseIf do
                    begin
                      Expr := ParseExpr;
                      NeedToken( tkThen );
                      TopRightCmt := CheckRightCmt;
                      ThenBlock := ParseStmtBlock( [tkElseIf, tkElse, _tkEndIf] );
                      assert( ThenBlock <> nil );
                    end;
                    
                if CheckToken( tkElse )
                  then
                    begin
                      elseTopRightCmt := CheckRightCmt;
                      ElseBlock := ParseStmtBlock( [_tkEndIf] );
                      assert( ElseBlock <> nil );
                      ElseBlock.TopRightCmt := elseTopRightCmt;
                      NeedToken( _tkEndIf );
                      ElseBlock.DownRightCmt := CheckRightCmt;
                    end
                  else
                    begin
                      NeedToken( _tkEndIf );
                      if elif <> nil
                        then elif.DownRightCmt := CheckRightCmt
                        else DownRightCmt := CheckRightCmt;
                    end;
              end
            else
              begin
                ThenBlock := ParseStmtBlock( [tkElse, tkEOL] );
                if CheckToken( tkElse )
                  then ElseBlock := ParseStmtBlock( [tkEOL] );
              end;
        end;
    end;

  // Select Case testexpression
  //   [Case expressionlist-n
  //     [statements-n]] ...
  //   [Case Else
  //     [elsestatements]]
  // End Select
  function TvbParser.ParseSelectStmt : TvbSelectCaseStmt;

    // Expression
    // [Is] RelationalOperator Expression
    // Expression [To Expression]
    //
    function ParseCaseClause : TvbCaseClause;
      var
        e        : TvbExpr;
        isClause : boolean;
        op       : TvbOperator;
      begin
        result := nil;
        isClause := CheckToken( tkIs );
        if fToken.tokenKind in
          [tkEq,
           tkNotEq,
           tkLess,
           tkLessEq,
           tkGreater,
           tkGreaterEq]
          then
            begin
              op := TokenEqvTable[fToken.tokenKind].eqvOperator;
              GetToken;
              result := TvbRelationalCaseClause.Create( op, ParseExpr );
            end
          else
        if not isClause
          then
            begin
              e := ParseExpr;
              try
                if CheckToken( tkTo )
                  then result := TvbRangeCaseClause.Create( e, ParseExpr )
                  else result := TvbCaseClause.Create( CASE_CLAUSE, e );
              except
                e.Free;
                raise;
              end;
            end
          else Error( 'invalid case expression' );
      end;

    var
      caseCmts  : TStringList;
    begin
      GetToken;
      NeedToken( tkCase );
      caseCmts := nil;
      result   := TvbSelectCaseStmt.Create( ParseExpr );
      with result do
        try
          caseCmts    := TStringList.Create;
          FrontLabel  := fFrontLabel;
          TopRightCmt := CheckRightCmt;
          
          NeedToken( tkEOL );
          repeat
            case fToken.tokenKind of
              tkEOL : GetToken;  // skip empty lines
              tkComment :
                begin
                  caseCmts.Add( fToken.tokenText );
                  GetToken;
                end;
              tkCase :
                begin
                  GetToken;
                  if CheckToken( tkElse )
                    then
                      begin
                        CaseElse := ParseStmtBlock( [_tkEndSelect] );
                        break;
                      end;
                  with AddCaseStmt( TvbCaseStmt.Create ) do
                    begin
                      AddCaseClause( ParseCaseClause );
                      while CheckToken( tkComma ) do
                        AddCaseClause( ParseCaseClause );
                      TopRightCmt := CheckRightCmt;
                      StmtBlock   := ParseStmtBlock( [tkCase, _tkEndSelect] );
                      TopCmts     := StrDynArrayFromStrList( caseCmts );
                      caseCmts.Clear;
                    end;
                end;
              _tkEndSelect : break;
              else break;
            end;
          until false;
          if caseCmts.Count > 0
            then result.DownCmts := StrDynArrayFromStrList( caseCmts );            
          FreeAndNil( caseCmts );
          NeedToken( _tkEndSelect );
          DownRightCmt := CheckRightCmt;
        except
          FreeAndNil( result );
          FreeAndNil( caseCmts );
          raise;
        end;
        fCurrBlock.Add( result );
    end;

  // Do [{While | Until} condition]
  //   [statements]
  // Loop
  // Do
  //   [statements]
  // Loop [{While | Until} condition]
  procedure TvbParser.ParseDoLoopStmt;
    var
      loop : TvbDoLoopStmt;

    function CheckCondition : boolean;
      begin
        if fToken.tokenKind in [tkWhile, tkUntil]
          then
            begin
              case fToken.tokenKind of
                tkWhile : loop.DoLoop := doWhile;
                tkUntil : loop.DoLoop := doUntil;
              end;
              GetToken;
              loop.Condition := ParseExpr;
              result := true;
            end
          else result := false;
      end;

    begin
      GetToken;
      loop := fCurrBlock.AddDoLoopStmt;
      with loop do
        begin
          TopCmts := StrDynArrayFromStrList( fCurrTopCmts );
          fCurrTopCmts.Clear;
          FrontLabel  := fFrontLabel;
          fFrontLabel := nil;
          TestAtBegin := CheckCondition;
          TopRightCmt := CheckRightCmt;
          StmtBlock   := ParseStmtBlock( [tkLoop] );
          NeedToken( tkLoop );
          if not TestAtBegin
            then CheckCondition;
          DownRightCmt := CheckRightCmt;
          if DoLoop = doInfinite
            then
              begin
                DoLoop    := doWhile;
                Condition := TvbBoolLit.Create( true );
              end;
        end;
    end;

  // While condition
  //   [statements]
  // Wend
  procedure TvbParser.ParseWhileStmt;
    begin
      GetToken;
      with fCurrBlock.AddWhileStmt do
        begin
          DoLoop      := doWhile;
          Condition   := ParseExpr;
          TestAtBegin := true;
          TopCmts     := StrDynArrayFromStrList( fCurrTopCmts );
          TopRightCmt := CheckRightCmt;
          FrontLabel  := fFrontLabel;
          StmtBlock   := ParseStmtBlock( [tkWend] );
          NeedToken( tkWend );
          DownRightCmt := CheckRightCmt;
        end;
    end;

  // For counter = start To end [Step step]
  //   [statements]
  // Next [counter]
  procedure TvbParser.ParseForStmt;
    begin
      with fCurrBlock.AddForStmt do
        begin
          Counter := ParseTermExpr;
          TopCmts := StrDynArrayFromStrList( fCurrTopCmts );
          FrontLabel := fFrontLabel;
          NeedToken( tkEq );
          InitValue := ParseExpr;
          NeedToken( tkTo );
          FinalValue := ParseExpr;
          if CheckIdent( 'Step' )
            then StepValue := ParseExpr;
          TopRightCmt := CheckRightCmt;
          StmtBlock := ParseStmtBlock( [tkNext] );
          NeedToken( tkNext );
          if not CheckToken( tkEOL) and ( fToken.tokenKind <> tkComment )
            then
              begin
                ParseTermExpr.Free;  // we don't need the Next argument
                if fToken.tokenKind = tkComma
                  then fToken.tokenKind := tkNext;
              end;
          DownRightCmt := CheckRightCmt;
        end;
    end;

  // For Each element In group
  //   [statements]
  // Next [element]
  procedure TvbParser.ParseForeachStmt;
    begin
      GetToken;
      with fCurrBlock.AddForeachStmt do
        begin
          Element    := ParseTermExpr;
          TopCmts    := StrDynArrayFromStrList( fCurrTopCmts );
          FrontLabel := fFrontLabel;
          NeedToken( tkIn );
          Group       := ParseTermExpr;
          TopRightCmt := CheckRightCmt;
          StmtBlock   := ParseStmtBlock( [tkNext] );
          NeedToken( tkNext );
          if not ( fToken.tokenKind in [tkEOL, tkComment, _tkEOF] )
            then
              begin
                ParseTermExpr.Free;  // we don't need the Next argument
                // COOL HACK! If now comes a comma as in 'Next A, B',
                // we substitute the comma by 'Next' so that it appears
                // we are parsing 'Next A Next B'.
                if fToken.tokenKind = tkComma
                  then fToken.tokenKind := tkNext;
              end;
          DownRightCmt := CheckRightCmt;
        end;
    end;

  // With expression
  //   [statements]
  // End With
  procedure TvbParser.ParseWithStmt;
    var
      prevWithExpr : TvbExpr;
    begin
      GetToken;
      with fCurrBlock.AddWithStmt do
        begin
          ObjOrRec      := ParseTermExpr;
          prevWithExpr  := fCurrWithExpr;
          fCurrWithExpr := ObjOrRec;
          TopCmts       := StrDynArrayFromStrList( fCurrTopCmts );
          TopRightCmt   := CheckRightCmt;
          FrontLabel    := fFrontLabel;
          StmtBlock     := ParseStmtBlock( [_tkEndWith] );
          NeedToken( _tkEndWith );
          DownRightCmt := CheckRightCmt;
          fCurrWithExpr := prevWithExpr;
        end;
    end;

  // Mid(stringvar, start[, length]) = string
  procedure TvbParser.ParseMidAssignStmt;
    begin
      with fCurrBlock.AddMidAssignStmt do
        begin
          VarName    := ParseTermExpr;
          TopCmts    := StrDynArrayFromStrList( fCurrTopCmts );
          FrontLabel := fFrontLabel;
          NeedToken( tkComma );
          StartPos := ParseExpr;
          if CheckToken( tkComma )
            then Length := ParseExpr;
          NeedToken( tkCParen );
          NeedToken( tkEq );
          Replacement := ParseExpr;
          TopRightCmt := CheckRightCmt;
        end;
    end;

  // LSet varname1 = varname2
  // RSet varname1 = varname2
  // Let varname1 = varname2
  // Set varname1 = varname2
  // varname1 = varname2
  function TvbParser.ParseAssignStmt( assign : TvbAssignKind ) : TvbAssignStmt;
    begin
      GetToken;
      result := TvbAssignStmt.Create( assign, ParseTermExpr );
      with result do
        try
          LHS.IsRValue := true;
          TopCmts      := StrDynArrayFromStrList( fCurrTopCmts );
          FrontLabel   := fFrontLabel;
          NeedToken( tkEq );
          Value      := ParseExpr;
          TopRightCmt := CheckRightCmt;
          FrontLabel  := fFrontLabel;
          fCurrBlock.Add( result );
        except
          Free;
          raise;
        end;
    end;

  // GoSub line
  // GoTo line
  procedure TvbParser.ParseGotoOrGosubStmt;
    begin
      with fCurrBlock.AddGotoOrGosubStmt do
        begin
          TopCmts    := StrDynArrayFromStrList( fCurrTopCmts );
          FrontLabel := fFrontLabel;
          if fToken.tokenKind = tkGoto
            then GotoOrGosub := ggGoTo
            else GotoOrGosub := ggGoSub;
          GetToken;
          LabelName   := ParseLabelRef;
          TopRightCmt := CheckRightCmt;
        end;
    end;

  // On Error GoTo line
  // On Error Resume Next
  // On Error GoTo 0
  procedure TvbParser.ParseOnErrorStmt;
    begin
      GetToken;
      with fCurrBlock.AddOnErrorStmt do
        begin
          TopCmts    := StrDynArrayFromStrList( fCurrTopCmts );
          FrontLabel := fFrontLabel;
          if CheckToken( tkResume )
            then
              begin
                NeedToken( tkNext );
                OnErrorDo := doResumeNext;
              end
            else
          if CheckToken( tkGoto )
            then
              if ( fToken.tokenKind = tkIntLit ) and ( fToken.tokenInt = 0 )
                then
                  begin
                    OnErrorDo := doGoto0;
                    GetToken;
                  end
                else
                  begin
                    OnErrorDo := doGotoLabel;
                    LabelName := ParseLabelRef;
                  end
            else Error( '"GoTo" or "Resume" expected' );
          TopRightCmt := CheckRightCmt;
        end;
    end;

  // On expression GoSub destinationlist
  // On expression GoTo destinationlist
  procedure TvbParser.ParseOnGotoOrGosubStmt;
    begin
      with fCurrBlock.AddOnGotoOrGosubStmt do
        begin
          Expr := ParseExpr;
          TopCmts    := StrDynArrayFromStrList( fCurrTopCmts );
          FrontLabel := fFrontLabel;
          if CheckToken( tkGoto )
            then GotoOrGosub := ggGoTo
            else
          if CheckToken( tkGoSub )
            then GotoOrGosub := ggGoSub
            else Error( '"GoTo" or "GoSub" expected' );
          repeat
            AddLabel( ParseLabelRef );
            CheckToken( tkComma );
          until fToken.tokenKind in [tkEOL, tkComment];
          TopRightCmt := CheckRightCmt;
        end;
    end;

  // Resume Next
  // Resume line
  // Resume [0]
  procedure TvbParser.ParseResumeStmt;
    begin
      GetToken;
      with fCurrBlock.AddResumeStmt do
        begin
          TopCmts    := StrDynArrayFromStrList( fCurrTopCmts );
          FrontLabel := fFrontLabel;
          case fToken.tokenKind of
            tkNext :
              begin
                ResumeKind := rkNext;
                GetToken;
              end;
            tkIntLit :
              if fToken.tokenInt <> 0
                then
                  begin
                    ResumeKind := rkLabel;
                    LabelName  := ParseLabelRef;
                  end
                else GetToken;
            tkIdent :
              begin
                ResumeKind := rkLabel;
                LabelName  := ParseLabelRef;
              end;
          end;
          TopRightCmt := CheckRightCmt;
        end;
    end;

  // RaiseEvent eventname [(argumentlist)]
  procedure TvbParser.ParseRaiseEventStmt;
    var
      sfx : TvbTypeSfx;
    begin
      GetToken;
      with fCurrBlock.AddRaiseEventStmt do
        begin
          EventName := GetIdent( sfx, true );
          TopCmts    := StrDynArrayFromStrList( fCurrTopCmts );
          FrontLabel := fFrontLabel;
          if CheckToken( tkOParen ) and not CheckToken( tkCParen )
            then
              begin
                EventArgs := ParseArgList;
                NeedToken( tkCParen );
              end;
          TopRightCmt := CheckRightCmt;
        end;
    end;

  // Open pathname [For mode] [Access access] [lock] As [#]filenumber [Len=reclength]
  procedure TvbParser.ParseOpenStmt;
    begin
      GetToken;
      with fCurrBlock.AddOpenStmt do
        begin
          FilePath := ParseExpr;
          TopCmts    := StrDynArrayFromStrList( fCurrTopCmts );
          FrontLabel := fFrontLabel;
          if CheckToken( tkFor )
            then
              if CheckToken( tkInput )
                then Mode := omInput
                else
              if CheckIdent( 'Append' )
                then Mode := omAppend
                else
              if CheckIdent( 'Binary' )
                then Mode := omBinary
                else
              if CheckIdent( 'Output' )
                then Mode := omOutput
                else
              if CheckIdent( 'Random' )
                then Mode := omRandom
                else Error( 'invalid Open access mode' );

          if CheckIdent( 'Access' )
            then
              if CheckIdent( 'Read' )
                then
                  if CheckToken( tkWrite )
                    then Access := oaReadWrite
                    else Access := oaRead
                else
              if CheckToken( tkWrite )
                then Access := oaWrite
                else Error( 'invalid Open lock' );

          if CheckToken( tkShared )
            then Lock := olShared
            else
          if CheckToken( tkLock )
            then
              if CheckIdent( 'Read' )
                then
                  if CheckToken( tkWrite )
                    then Lock := olLockReadWrite
                    else Lock := olLockRead
                else
              if CheckToken( tkWrite )
                then Lock := olLockWrite
                else Error( 'invalid Open lock' );

          NeedToken( tkAs );
          CheckToken( tkHash );

          FileNum := ParseExpr;

          if CheckToken( tkLen )
            then
              begin
                NeedToken( tkEq );
                RecLen := ParseExpr;
              end;
          TopRightCmt := CheckRightCmt;
        end;
    end;

  // ReDim [Preserve] varname(subscripts) [As type] [, varname(subscripts) [As type]]..
  procedure TvbParser.ParseReDimStmt;
    var
      pre : boolean;
      def : TvbDef;
    begin
      GetToken;
      pre := CheckToken( tkPreserve );
      repeat
        with fCurrBlock.AddReDimStmt do
          begin
            Preserve := pre;
            TopCmts := StrDynArrayFromStrList( fCurrTopCmts );
            fCurrTopCmts.Clear;
            FrontLabel := fFrontLabel;
            fLastLabel := nil; // only the first ReDim gets the labels

            // Clear the top comments so that ArrayVar gets none.
            fCurrTopCmts.Clear;
            ArrayVar := ParseVarDef( false, false );

            // Search for the array variable in the current scope(function)
            // and the parent scope(standard module or class definition).
            // If found, this ReDim is refering to it and it's only changing
            // dimensions. If not found, this ReDim is also declaring the
            // array in the function's scope.
            if not fCurrScope.Lookup( OBJECT_BINDING, ArrayVar.Name, def, false ) and
               ( ( fCurrScope.Parent <> nil ) and not fCurrScope.Parent.Lookup( OBJECT_BINDING, ArrayVar.Name, def, false ) )
              then
                begin
                  ArrayVar.NodeFlags := ArrayVar.NodeFlags + [nfReDim_DeclaredLocalVar];
                  fCurrFunc.AddLocalVar( ArrayVar );
                end;

            // If the variable has a TopRightCmt pass them to the ReDim stmt.
            // This TopRightCmt was collected by ParseVarDef.
            if ArrayVar.TopRightCmt <> ''
              then
                begin
                  TopRightCmt := ArrayVar.TopRightCmt;
                  ArrayVar.TopRightCmt := '';
                end;
          end;
      until not CheckToken( tkComma );
    end;

  // Erase arraylist
  procedure TvbParser.ParseEraseStmt;
    begin
      GetToken;
      repeat
        with fCurrBlock.AddEraseStmt do
          begin
            TopCmts := StrDynArrayFromStrList( fCurrTopCmts );
            fCurrTopCmts.Clear;
            FrontLabel := fFrontLabel;
            fLastLabel := nil; // only the first Erase gets the labels
            ArrayVar := ParseTermExpr;
            TopRightCmt := CheckRightCmt;
          end;
      until not CheckToken( tkComma );
    end;

  // Stop
  procedure TvbParser.ParseStopStmt;
    begin
      GetToken;
      with fCurrBlock.AddStopStmt do
        begin
          TopCmts     := StrDynArrayFromStrList( fCurrTopCmts );
          TopRightCmt := CheckRightCmt;
          FrontLabel  := fFrontLabel;
        end;
    end;

   // Close [[#]filenumber] [, [#]filenumber] ...
  procedure TvbParser.ParseCloseStmt;
    begin
      GetToken;
      repeat
        with fCurrBlock.AddCloseStmt do
          begin
            TopCmts := StrDynArrayFromStrList( fCurrTopCmts );
            fCurrTopCmts.Clear;
            FrontLabel := fFrontLabel;
            fLastLabel := nil; // only the first Close gets the labels

            // We may have a comment here if this Close statement is being
            // used without file numbers.
            if fToken.tokenKind = tkComment
              then
                begin
                  TopRightCmt := fToken.tokenText;
                  GetToken;
                end
              else
            if not CheckToken( tkEOL )
              then
                begin
                  CheckToken( tkHash );
                  FileNum     := ParseExpr;
                  TopRightCmt := CheckRightCmt;
                end;
          end;
      until not CheckToken( tkComma );
    end;

  // Date = newDate
  procedure TvbParser.ParseDateStmt;
    begin
      GetToken;
      NeedToken( tkEq );
      with fCurrBlock.AddDateStmt do
        begin
          NewDate     := ParseExpr;
          TopCmts     := StrDynArrayFromStrList( fCurrTopCmts );
          TopRightCmt := CheckRightCmt;
          FrontLabel  := fFrontLabel;
        end;
    end;

  // Get [#]filenumber, [recnumber], varname
  procedure TvbParser.ParseGetStmt;
    begin
      GetToken;
      CheckToken( tkHash );
      with fCurrBlock.AddGetStmt do
        begin
          TopCmts := StrDynArrayFromStrList( fCurrTopCmts );
          FrontLabel := fFrontLabel;
          FileNum := ParseExpr;
          NeedToken( tkComma );
          if not CheckToken( tkComma)
            then
              begin
                RecNum := ParseExpr;
                NeedToken( tkComma );
              end;
          VarName     := ParseExpr;
          TopRightCmt := CheckRightCmt;
        end;
    end;

  // Lock [#]filenumber[, recordrange]
  procedure TvbParser.ParseLockStmt;
    begin
      GetToken;
      CheckToken( tkHash );
      with fCurrBlock.AddLockStmt do
        begin
          FileNum := ParseExpr;
          TopCmts    := StrDynArrayFromStrList( fCurrTopCmts );
          FrontLabel := fFrontLabel;
          if CheckToken( tkComma )
            then
              begin
                StartRec := ParseExpr;
                if CheckToken( tkTo )
                  then EndRec := ParseExpr;
              end;
          TopRightCmt := CheckRightCmt;
        end;
    end;

  // Put [#]filenumber, [recnumber], varname
  procedure TvbParser.ParsePutStmt;
    begin
      GetToken;
      CheckToken( tkHash );
      with fCurrBlock.AddPutStmt do
        begin
          FileNum := ParseExpr;
          TopCmts    := StrDynArrayFromStrList( fCurrTopCmts );
          FrontLabel := fFrontLabel;
          NeedToken( tkComma );
          if not CheckToken( tkComma)
            then
              begin
                RecNum := ParseExpr;
                NeedToken( tkComma );
              end;
          VarName := ParseExpr;
          TopRightCmt := CheckRightCmt;
        end;
    end;

  // Seek [#]filenumber, position
  procedure TvbParser.ParseSeekStmt;
    begin
      GetToken;
      CheckToken( tkHash );
      with fCurrBlock.AddSeekStmt do
        begin
          FileNum := ParseExpr;
          TopCmts    := StrDynArrayFromStrList( fCurrTopCmts );
          FrontLabel := fFrontLabel;
          NeedToken( tkComma );
          Pos := ParseExpr;
          TopRightCmt := CheckRightCmt;
        end;
    end;

  // Unlock [#]filenumber[, recordrange]
  procedure TvbParser.ParseUnlockStmt;
    begin
      GetToken;
      CheckToken( tkHash );
      with fCurrBlock.AddUnlockStmt do
        begin
          FileNum := ParseExpr;
          TopCmts    := StrDynArrayFromStrList( fCurrTopCmts );
          FrontLabel := fFrontLabel;
          if CheckToken( tkComma )
            then
              begin
                StartRec := ParseExpr;
                if CheckToken( tkTo )
                  then EndRec := ParseExpr;
              end;
          TopRightCmt := CheckRightCmt;
        end;
    end;

  // Write #filenumber, [outputlist]
  procedure TvbParser.ParseWriteStmt;
    begin
      GetToken;
      NeedToken( tkHash );
      with fCurrBlock.AddWriteStmt do
        begin
          FileNum := ParseExpr;
          TopCmts    := StrDynArrayFromStrList( fCurrTopCmts );
          FrontLabel := fFrontLabel;
          NeedToken( tkComma );
          if fToken.tokenKind <> tkEOL
            then
              repeat
                AddValue( ParseExpr );
              until not CheckToken( tkComma );
          TopRightCmt := CheckRightCmt;
        end;
    end;

  // Input #filenumber, varlist
  procedure TvbParser.ParseInputStmt;
    begin
      GetToken;
      NeedToken( tkHash );
      with fCurrBlock.AddInputStmt do
        begin
          FileNum := ParseExpr;
          TopCmts    := StrDynArrayFromStrList( fCurrTopCmts );
          FrontLabel := fFrontLabel;
          NeedToken( tkComma );
          repeat
            AddVar( ParseExpr );
          until not CheckToken( tkComma );
          TopRightCmt := CheckRightCmt;
        end;
    end;

  // Print #filenumber, [outputlist]
  // [object.]Print
  procedure TvbParser.ParsePrintStmtOrMethod( obj : TvbExpr );
    var
      printToFile : TvbPrintStmt;
      printToObj  : TvbPrintMethod;
    begin
      GetToken;
      if CheckToken( tkHash )
        then
          begin
          end
        else ParsePrintArgList( fCurrBlock.AddPrintMethod );
    end;

  // Debug.Assert booleanExpression
  procedure TvbParser.ParseDebugAssertStmt;
    begin
      with fCurrBlock.AddDebugAssertStmt do
        begin
          BoolExpr    := ParseExpr;
          TopCmts     := StrDynArrayFromStrList( fCurrTopCmts );
          TopRightCmt := CheckRightCmt;
          FrontLabel  := fFrontLabel;
        end;
    end;

  // Debug.Print [outputlist]
  procedure TvbParser.ParseDebugPrintStmt;
    begin
      ParsePrintArgList( fCurrBlock.AddDebugPrintStmt );
    end;

  // Parses the list of arguments(one at least) for Debug.Print and Print.
  // If the arguments terminate with ';' or ',' BreakLine will be false.
  // Otherewise it will be true.
  procedure TvbParser.ParsePrintArgList( printStmt : TvbAbstractPrintStmt );
    begin
      with printStmt, fToken do
        begin
          LineBreak := true;
          TopCmts := StrDynArrayFromStrList( fCurrTopCmts );
          FrontLabel := fFrontLabel;

          // parse the argument list.
          if not ( fToken.tokenKind in [tkEOL, tkColon, tkComment] )
            then
              repeat
                case tokenKind of
                  tkComma : // replace commas by "Tab"
                    begin
                      AddExpr( pakTab );
                      GetToken;
                      LineBreak := false;
                    end;
                  tkSemiColon : // ignore semicolons
                    begin
                      GetToken;
                      LineBreak := false;
                    end;
                  tkSpc :
                    begin
                      GetToken;
                      NeedToken( tkOParen ); // argument is required
                      AddExpr( pakSpc ).Expr := ParseExpr;
                      NeedToken( tkCParen);
                      LineBreak := true;
                    end;
                  tkTab :
                    begin
                      GetToken;
                      LineBreak := true;
                      if CheckToken( tkOParen )  // Tab argument is optional
                        then
                          begin
                            AddExpr( pakTab ).Expr := ParseExpr;
                            NeedToken( tkCParen );
                          end;
                    end;
                  else AddExpr( pakExpr ).Expr := ParseExpr;
                end;
              until tokenKind in [tkEOL, tkColon, tkComment];

          TopRightCmt := CheckRightCmt;
        end;
    end;

  function TvbParser.GetTempVarName : string;
    begin
      inc( fTempVarCount );
      result := 'var' + IntToStr( fTempVarCount );
    end;

  function TvbParser.CalcBinaryExpr( op : TvbOperator; left, right : TvbExpr ) : TvbExpr;
    var
      l, r : integer;

    function IsIntConst( expr : TvbExpr; var value : integer ) : boolean;
      var
        ue : TvbUnaryExpr absolute expr;
      begin
        result := true;
        if expr.NodeKind = INT_LIT_EXPR
          then value := TvbIntLit( expr ).Value
          else
        if ( expr.NodeKind = UNARY_EXPR ) and
           ( ue.Oper = SUB_OP ) and
           ( ue.Expr.NodeKind = INT_LIT_EXPR )
          then value := -TvbIntLit( ue.Expr ).Value
          else result := false;
      end;

    begin
      if IsIntConst( left, l ) and IsIntConst( right, r )
        then
          case op of
            ADD_OP     : result := TvbIntLit.Create( l + r );
            SUB_OP     : result := TvbIntLit.Create( l - r );
            MULT_OP    : result := TvbIntLit.Create( l * r );
            INT_DIV_OP : result := TvbIntLit.Create( l div r );
            else result := TvbBinaryExpr.Create( op, left, right );
          end
        else result := TvbBinaryExpr.Create( op, left, right );
    end;

  function TvbParser.ParseClassModule( const Name, relPath : string; Source : TvbCustomTextBuffer ) : TvbModuleDef;
    var
      oldScope : TvbScope;
    begin
      DoParseModuleBegin( Name, relPath );
      
      fModuleOpts.SetDefaults;
      oldScope   := fCurrScope;
      fCurrModule := TvbClassModuleDef.Create( Name, relPath, fCurrScope );
      try
        fCurrModule.RelativePath := relPath;
        assert( fCurrModule.ClassCount = 1 );
        // insert the definition of the class type in the type map.
        fTypeMap.DefineType( fCurrModule.Classes[0] );
        // set the current scope to the class's scope.
        fCurrScope := fCurrModule.Classes[0].Scope;
        fLexer.PushBuffer( Source );
        try
          GetToken;
          if CheckIdent( 'VERSION' )
            then
              begin
                NeedToken( tkFloatLit );
                NeedIdent( 'CLASS' );
                NeedToken( tkEOL );
                NeedIdent( 'BEGIN' );
                while not CheckToken( tkEnd ) do
                  GetToken;
              end;
          ParseModule;
          result := fCurrModule;
          fCurrScope := oldScope;
        finally
          fLexer.PopBuffer;
        end;
      except
        fCurrModule.Free;
        raise;
      end;

      DoParseModuleEnd;
    end;

  destructor TvbParser.Destroy;
    begin
      FreeAndNil( fCurrTopCmts );
      FreeAndNil( fModuleOpts );
      FreeAndNil( fLexer );
      FreeAndNil( fPropertyMap );
      FreeAndNil( fTypeMap );
      FreeAndNil( fRegExp );
      inherited;
    end;

  procedure TvbParser.ParseModule;
    var
      i : integer;
      prop : TvbPropertyDef;
    begin
      fCurrBlock    := nil;
      fCurrFunc     := nil;
      fCurrFuncName := '';
      fCurrWithExpr := nil;

      try
        repeat
          case fToken.tokenKind of
            tkEOL : GetToken;
            tkAttribute :
              repeat // skip whole line for the time being
                GetToken;
              until fToken.tokenKind = tkEOL;
            tkComment :
              begin
                fCurrTopCmts.Add( fToken.tokenText );
                GetToken;
              end;
            _tkEOF : break;
            tkDefBool,
            tkDefByte,
            tkDefCur,
            tkDefDate,
            tkDefDbl,
            tkDefInt,
            tkDefLng,
            tkDefObj,
            tkDefSng,
            tkDefStr,
            tkDefVar  : ParseDefStmt;
            tkOption  : ParseOptionStmt;
            else
              begin
                ParseDef;
                fCurrTopCmts.Clear;
              end;
          end;
        until false;

        // We'll have TopCmts if no declaration followed those comments.
        // Then those comments become DownCmts of the module.
        fCurrModule.DownCmts := StrDynArrayFromStrList( fCurrTopCmts );

        // copy all the properties in the property map to the current module.
        // but before call the method that set's the property parameters and
        // return type.
        for i := 0 to fPropertyMap.PropCount - 1 do
          begin
            prop := fPropertyMap[i];
            assert( prop <> nil );
            prop.SetParamsAndType;
            fCurrModule.AddProperty( prop );
          end;
        // reset the property map.
        fPropertyMap.Clear;
        
        // clear all type references created in the current module.
        fTypeMap.Clear;
      except
        on e : EvbSourceError do
          begin
//            if fMonitor <> nil
//              then fMonitor.ParseModuleFailed( e.Message, e.Line, e.Column );
            raise;
          end;
        on e : Exception do
          begin
            Log( e.Message, ltError );
            raise;
          end;
      end;
    end;

  // If the current token is a comment, return it and get the next token.
  // Otherwise return an empty string.
  function TvbParser.CheckRightCmt : string;
    begin
      if fToken.tokenKind = tkComment
        then
          begin
            result := fToken.tokenText;
            GetToken;
          end
        else result := '';
    end;

  procedure TvbParser.ParseEndStmt;           
    var
      stmt :  TvbEndStmt;
    begin
      GetToken;
      stmt := TvbEndStmt.Create;
      with stmt do
      try
        TopCmts     := StrDynArrayFromStrList( fCurrTopCmts );
        TopRightCmt := CheckRightCmt;
        FrontLabel  := fFrontLabel;
        fCurrBlock.Add( stmt );
      except
        Free;
        raise
      end;
    end;

  procedure TvbParser.ParseLocalVarOrConst;
    begin
      ParseDeclFlags;
      if dfConst in fCurrNodeFlags
        then
          repeat
            fCurrFunc.AddLocalConst( ParseConstDef( false ) );
            fCurrTopCmts.Clear;
          until not CheckToken( tkComma )
        else
          repeat
            fCurrFunc.AddLocalVar( ParseVarDef( false, false ) );
            fCurrTopCmts.Clear;
          until not CheckToken( tkComma );
    end;

  procedure TvbParser.ParseReturnStmt;
    var
      stmt : TvbReturnStmt;
    begin
      GetToken;
      stmt := TvbReturnStmt.Create;
      with stmt do
        try
          TopCmts     := StrDynArrayFromStrList( fCurrTopCmts );
          TopRightCmt := CheckRightCmt;
          FrontLabel  := fFrontLabel;
          fCurrBlock.Add( stmt );
        except
          Free;
          raise;
        end;
    end;

  procedure TvbParser.CheckError( cond : boolean; const msg : string );
    begin
      if not cond
        then Error( msg );
    end;

  function TvbParser.ParseFormModule( const Name, relPath : string; Source : TvbCustomTextBuffer ) : TvbModuleDef;
    var
      oldScope : TvbScope;
      form : TvbFormModuleDef;
    begin
      DoParseModuleBegin( Name, relPath );

      fModuleOpts.SetDefaults;
      oldScope := fCurrScope;
      form := TvbFormModuleDef.Create( Name,  relPath, fCurrScope );
      fCurrModule := form;
      try
        fCurrModule.RelativePath := relPath;
        assert( form.FormDef <> nil );
        // insert the definition of the form(class) type in the type map.
        fTypeMap.DefineType( form.FormDef );
        // this is the default form variable.
        // this is typically referenced from other modules to use this form.
        form.FormVar.VarType := fTypeMap.GetType( Name );
        // set the current scope to the form's scope.
        fCurrScope := form.FormDef.Scope;
        assert( fCurrScope.BaseScope = nil );

        fLexer.PushBuffer( Source );
        try
          GetToken;
          if CheckIdent( 'VERSION' )
            then
              begin
                NeedToken( tkFloatLit );
                NeedToken( tkEOL );
              end;
          ParseFormDefinition( form.FormDef );
          ParseModule;
          result := fCurrModule;
          fCurrScope := oldScope;
        finally
          fLexer.PopBuffer;
        end;
      except
        fCurrModule.Free;
        raise;
      end;

      DoParseModuleEnd;
    end;

  procedure TvbParser.CheckMoveCommentsToTop;
    begin
      if fOptions.optCommentsBelowFuncToTop
        then
          while fToken.tokenKind = tkComment do
            begin
              fCurrTopCmts.Add( fToken.tokenText );
              GetToken;
              CheckToken( tkEOL );
            end;
    end;

  function TvbParser.ParseProject( Lines : TStringList; const relPath : string ) : TvbProject;
    var
      i, j   : integer;
      line   : string;
      option : string;
      value  : string;
    begin
      DoParseProjectBegin( '', relPath );

      result := TvbProject.Create;
      try
        for j := 1 to Lines.Count - 1 do
          begin
            line := trim( Lines[j] );
            i := pos( '=', line );
            if ( line = '' ) or      // skip empty lines
               ( line[1] = ';' ) or  // and comments
               ( line[1] = '[' ) or  // and section names
               ( i = 0 )             // and non OPTION=VALUE kind of lines
              then continue;

            option := trim( lowercase( copy( line, 1, i - 1 ) ) );
            value  := trim( copy( line, i + 1, length( line ) ) );
            if value[1] = '"'
              then
                begin
                  delete( value, 1, 1 );
                  if value[length( value )] = '"'
                    then delete( value, length( value ), 1 );
                end;
            if value = ''
              then continue;

            ParseProjectOption( option, value, result );
          end;
      except
        DoParseProjectFail;
        FreeAndNil( result );
        raise;
      end;

      DoParseProjectEnd;
    end;

  procedure TvbParser.ParseProjectOption( const Option, Value : string; Project : TvbProject );
    const
      reReference =
        '\*\\G' + // \*G
        '(\{[\dA-F]{8}-[\dA-F]{4}-[\dA-F]{4}-[\dA-F]{4}-[\dA-F]{12}\})' + // GUID
        '#(\d+)\.(\d+)#' + // #MajorVer.MinorVer#
        '(\d+)' +          // Lcid  nnn
        '#(.+?)#' +        // #Typelib path#
        '(.*)';            // Typelib description
      reObject =
        '(\{[\dA-F]{8}-[\dA-F]{4}-[\dA-F]{4}-[\dA-F]{4}-[\dA-F]{12}\})' + // GUID
        '#(\d+)\.(\d+)#' + // #MajorVer.MinorVer#
        '(\d+);\s*' +      // Lcid  nnn
        '(.*)';            // library name
    var
      i             : integer;
      moduleName    : string;
      moduleRelPath : string; // module's file path relative to project's directory
      moduleDef     : TvbModuleDef;
      source        : TvbCustomTextBuffer;
      map           : TvbStdModuleDef;
    begin
      assert( Option <> '' );
      assert( Value <> '' );
      with Project do
        if ( Option = 'module' ) or
           ( Option = 'class' ) or
           ( Option = 'form' )
          then
            begin
              i := pos( ';', Value );
              if i <> 0
                then
                  begin
                    moduleName    := trim( copy( Value, 1, i - 1 ) );
                    moduleRelPath := trim( copy( Value, i + 1, length( Value ) ) );
                  end
                else
                  begin
                    moduleName    := ChangeFileExt( ExtractFileName( Value ), '' );
                    moduleRelPath := Value;
                  end;
              source := fResourceProvider.GetSourceCode( moduleRelPath );
              if source <> nil
                then
                  begin
                    //writeln( ModuleRelPath );
                    if Option = 'module'
                      then moduleDef := Project.AddModule( ParseStdModule( moduleName, moduleRelPath, source ) )
                      else
                    if Option = 'class'
                      then moduleDef := Project.AddModule( ParseClassModule( moduleName, moduleRelPath, source ) )
                      else
                    if Option = 'form'
                      then moduleDef := Project.AddModule( ParseFormModule( moduleName, moduleRelPath, source ) )
                      else assert( true );

                    assert( moduleDef <> nil );
                  end
                else Log( Format( 'couldn''t read "%s"', [moduleRelPath] ), ltWarn );
            end
          else
        if Option = 'reference'
          then
            begin
              fRegexp.Expression := reReference;
              if fRegexp.Exec( Value )
                then
                  begin
                    map := fResourceProvider.LoadTranslationMap( ExtractFileName( fRegexp.Match[5] ) );
                    if map <> nil
                      then Project.AddImportedModule( map );                      
                  end;
            end
          else
        if Option = 'object'
          then
            begin
              fRegexp.Expression := reObject;
              if fRegexp.Exec( Value )
                then
                  begin
                    map := fResourceProvider.LoadTranslationMap( ExtractFileName( fRegexp.Match[4] ) );
                    if map <> nil
                      then Project.AddImportedModule( map );
                  end;
            end
          else
        if Option = 'name'
          then Name := Value
          else
        if Option = 'exename32'
          then optExeName32 := Value;
    end;

  procedure TvbParser.Log( const msg : string; line : cardinal; col : cardinal; logType : TvbLogType = ltInfo );
    begin
      if fMonitor <> nil
        then fMonitor.Log( msg, line, col, logType );
    end;

  procedure TvbParser.Log( const msg : string; logType : TvbLogType = ltInfo );
    begin
      if fMonitor <> nil
        then fMonitor.Log( msg, logType );
    end;

  procedure TvbParser.SkipFormLine;
    begin
      while not ( fToken.tokenKind in [tkEOL, _tkEOF] ) do
        GetFormToken;
      CheckFormToken( tkEOL );
    end;

  procedure TvbParser.ParseObjectProperty( owner : TvbControl );
    begin
      repeat
        if CheckFormIdent( 'EndProperty' )
          then
            begin
              NeedFormToken( tkEOL );
              break;
            end
          else ParseProperty( owner );
      until false;
    end;

  // parses a property and adds it to 'owner'.
  procedure TvbParser.ParseProperty( owner : TvbControl );
    var
      name : string;
    begin
      assert( owner <> nil );
      with fToken do
        if CheckFormIdent( 'BeginProperty' )
          then
            begin
              CheckError( tokenKind in [tkIdent, tkAbs..tkXor], 'property name expected' );
              name := tokenText;
              GetFormToken;
              // skip whatever comes for the time being
              if tokenKind <> tkEOL
                then GetFormToken;
              NeedFormToken( tkEOL );
              ParseObjectProperty( owner.AddObjectProperty( name ) );
            end
          else
            begin
              repeat
                CheckError( tokenKind in [tkIdent, tkAbs..tkXor], 'property name expected' );
                name := tokenText;
                GetFormToken;
                // if now we see a '.' this means we are parsing a property like
                // 'Font.FontColor'. in this case we create an object property,
                // add it to 'owner' and set 'owner' to the newly created
                // control. this way when we later create the property it's created
                // in the correct 'owner'.
                if CheckFormToken( tkDot )
                  then owner := owner.AddObjectProperty( name )
                  else break;
              until false;

              //TODO: los parentesis que sigen a veces el nombre de la propiedad.
              NeedFormToken( tkEq );

              assert( owner <> nil );
              assert( name <> '' );

              // now we create a property named 'name' in control 'owner'.
              // and set its value.              
              with owner.AddProperty( name ) do
                case tokenKind of
                  _tkFormRes   : Value := tokenFormRes;
                  _tkFormCtrl  : Value := tokenFormCtrl;
                  tkIntLit     : Value := tokenInt;
                  tkFloatLit   : Value := tokenFloat;
                  tkDateLit    : Value := VarFromDateTime( tokenDate );
                  _tkBracketStr,
                  tkStringLit,
                  tkAbs..tkXor : Value := string( tokenText );
                  else
                    begin
                      DoFormDefinitionProblem( 'property vaue not supported' );
                      exit;
                    end;
                end;
              GetFormToken;
              CheckFormToken( tkComment );
              NeedFormToken( tkEOL );
            end;
    end;

  function TvbParser.ParseControl( form : TvbFormDef ) : TvbControl;
    var
      name, typ : string;
    begin
      result := nil;
      with fToken do
        try
          NeedFormIdent( 'Begin' );
          // parse the control type name
          typ := '';
          repeat
            CheckError( tokenKind in [tkIdent, tkAbs..tkXor], 'invalid type name' );
            typ := typ + tokenText;
            GetFormToken;
            if CheckFormToken( tkDot )
              then typ := typ + '.'
              else break;
          until false;

          // parse the control name
          CheckError( tokenKind in [tkIdent, tkAbs..tkXor], 'control name expected' );
          name := tokenText;
          GetFormToken;
          NeedFormToken( tkEOL );

          result := TvbControl.Create( name, fTypeMap.GetType( typ ) );

          DoParseFormContronBegin( result );
          form.OnNewControl( result );

          // parse properties and child controls.
          repeat
            if CheckFormToken( tkEnd )
              then
                begin
                  CheckFormToken( tkComment );
                  NeedFormToken( tkEOL );
                  break;
                end
              else
            if SameText( tokenText, 'Begin' )
              then result.AddControl( ParseControl( form ) )
              else ParseProperty( result );
          until false;

          DoParseFormContronEnd;
        except;
          FreeAndNil( result );
          raise;
        end;
    end;

//      beginNestingLevel := 0;
//      with fLexer do
//        repeat
//          with fToken do
//            if ( tokenKind = tkIdent ) and ( stricomp( tokenText, 'Begin' ) = 0 )
//              then
//                begin
//                  fToken := GetFormToken;
//                  inc( beginNestingLevel );
//
//                  if ( tokenKind = tkIdent ) and ( beginNestingLevel > 1 )
//                    then
//                      begin
//                        typeName := '';
//                        repeat
//                          typeName := typeName + tokenText;
//                          fToken := GetFormToken;
//                          if tokenKind = tkDot
//                            then
//                              begin
//                                typeName := typeName + '.';
//                                fToken := GetFormToken;
//                              end
//                            else break;
//                        until false;
//                        
//                        if tokenKind = tkIdent
//                          then
//                            begin
//                              ctlName := tokenText;
//                              form.AddControl( ctlName, fTypeMap.GetType( typeName ) );
//                              fToken := GetFormToken;
//                            end;
//                      end;
//                end
//              else
//            if tokenKind = tkEnd
//              then
//                begin
//                  fToken := GetFormToken;
//                  if tokenKind = tkComment
//                    then fToken := GetFormToken;
//                  dec( beginNestingLevel );
//                  if beginNestingLevel = 0
//                    then break
//                    else
//                  if beginNestingLevel < 0
//                    then Error( '"End" without "Begin"' );
//                end
//              else
//                begin
//                end;
//          fToken := GetFormToken;
//        until fToken.tokenKind in [_tkEOF]; // just in case!!
  procedure TvbParser.ParseFormDefinition( form : TvbFormDef );
    begin
      DoParseFormDefinitionBegin( form );

      fLexer.TopBuffer.buffFormDef := true;

      while CheckToken( tkObject ) do
        begin
          GetToken; // =
          GetToken; // {...}
          if not CheckToken( tkEOL )
            then
              begin
                GetToken; // ;
                GetToken; // library
                NeedToken( tkEOL );
              end;
        end;

      form.RootControl := ParseControl( form );

      fLexer.TopBuffer.buffFormDef := false;

      DoParseFormDefinitionEnd;
    end;

  // return TRUE if the given func/var/property is declared inside a function
  // or it's parent scope. remember that a function's parent module can be
  // a standard module or the class definition(TvbClassDef) of a class/form module.
  // this function should be called ONLY when parsing statements.
  function TvbParser.IsModuleLevelDef( const name : string ) : boolean;
    var
      def : TvbDef;
    begin
      assert( fCurrScope <> nil );
      assert( fCurrScope.Parent <> nil );
      result :=
        fCurrScope.Lookup( OBJECT_BINDING, name, def, false ) or
        fCurrScope.Parent.Lookup( OBJECT_BINDING, name, def, false );
    end;

  function TvbParser.CheckFormIdent( const name : string ) : boolean;
    begin
      if ( fToken.tokenKind = tkIdent ) and SameText( fToken.tokenText, name )
        then
          begin
            GetFormToken;
            result := true;
          end
        else result := false;
    end;

  function TvbParser.CheckFormToken( tk : TvbTokenKind ) : boolean;
    begin
      if fToken.tokenKind = tk
        then
          begin
            GetFormToken;
            result := true;
          end
        else result := false;
    end;

  procedure TvbParser.NeedFormIdent( const name : string );
    begin
      if ( fToken.tokenKind = tkIdent ) and SameText( fToken.tokenText, name )
        then GetFormToken
        else Error( Format('"%s" expected', [name] ) );        
    end;

  procedure TvbParser.NeedFormToken( tk : TvbTokenKind );
    begin
      if not CheckFormToken( tk )
        then Error( Format( 'expected "%s"', [GetTokenText( tk )] ) );
    end;

  procedure TvbParser.GetFormToken;
    begin
      fToken := fLexer.GetFormToken;
    end;

  procedure TvbParser.DoFormDefinitionProblem( const msg : string; skipLine : boolean );
    begin
      if assigned( fOnFormDefinitionProblem )
        then fOnFormDefinitionProblem( fLexer.TokenLine, fLexer.TokenColumn, msg );
      if skipLine
        then SkipFormLine;
    end;

  procedure TvbParser.DoParseModuleBegin( const name, relativePath : string );
    begin
      if assigned( fOnParseModuleBegin )
        then fOnParseModuleBegin( name, relativePath );
    end;

  procedure TvbParser.DoParseModuleEnd;
    begin
      if assigned( fOnParseModuleEnd )
        then fOnParseModuleEnd;
    end;

  procedure TvbParser.DoParseModuleFail;
    begin
      if assigned( fOnParseModuleFail )
        then fOnParseModuleFail;
    end;

  procedure TvbParser.DoParseProjectBegin( const name, relativePath : string );
    begin
      if assigned( fOnParseProjectBegin )
        then fOnParseProjectBegin( name, relativePath );
    end;

  procedure TvbParser.DoParseProjectEnd;
    begin
      if assigned( fOnParseProjectEnd )
        then fOnParseProjectEnd;
    end;

  procedure TvbParser.DoParseProjectFail;
    begin
      if assigned( fOnParseProjectFail )
        then fOnParseProjectFail;
    end;

  procedure TvbParser.DoParseFormDefinitionBegin( form : TvbFormDef );
    begin
      if assigned( fOnParseFormDefinitionBegin )
        then fOnParseFormDefinitionBegin( form );
    end;

  procedure TvbParser.DoParseFormDefinitionEnd;
    begin
      if assigned( fOnParseFormDefinitionEnd )
        then fOnParseFormDefinitionEnd;
    end;

  procedure TvbParser.DoParseFormContronBegin( control : TvbControl );
    begin
      if assigned( fOnParseFormControlBegin )
        then fOnParseFormControlBegin( control );
    end;

  procedure TvbParser.DoParseFormContronEnd;
    begin
      if assigned( fOnParseFormControlEnd )
        then fOnParseFormControlEnd;
    end;

  procedure TvbParser.ParsePrintMethod( obj : TvbExpr );
    var
      stmt : TvbPrintMethod;
    begin
      stmt := fCurrBlock.AddPrintMethod;
      stmt.Obj := obj;
      ParsePrintArgList( stmt );
    end;

  procedure TvbParser.ParsePrintStmt;
    var
      stmt : TvbPrintStmt;
    begin
      stmt := fCurrBlock.AddPrintStmt;
      stmt.FileNum := ParseExpr;
      ParsePrintArgList( stmt );
    end;

  { TvbModuleOptions }

  constructor TvbModuleOptions.Create;
    begin
      SetDefaults;
    end;

  procedure TvbModuleOptions.SetDefaults;
    var
      i : char;
    begin
      fExplicit      := false;
      fPrivateModule := false;
      fCompare       := ocBinary;
      for i := low( fDefaultTypes ) to high( fDefaultTypes ) do
        fDefaultTypes[i] := vbVariantType;
    end;

  function TvbModuleOptions.GetDefinedType( const name : string ) : IvbType;
    begin
      result := fDefaultTypes[name[1]];
    end;

  procedure TvbModuleOptions.DefineType( fromChar, toChar : char; typk : IvbType );
    var
      c : char;
    begin
      for c := fromChar to toChar do
        fDefaultTypes[c] := typk;
    end;

initialization

  // Operators
  TokenEqvTable[tkConcat].eqvOperator    := STR_CONCAT_OP;
  TokenEqvTable[tkMult].eqvOperator      := MULT_OP;
  TokenEqvTable[tkPlus].eqvOperator      := ADD_OP;
  TokenEqvTable[tkMinus].eqvOperator     := SUB_OP;
  TokenEqvTable[tkFloatDiv].eqvOperator  := DIV_OP;
  TokenEqvTable[tkIntDiv].eqvOperator    := INT_DIV_OP;
  TokenEqvTable[tkPower].eqvOperator     := POWER_OP;
  TokenEqvTable[tkEq].eqvOperator        := EQ_OP;
  TokenEqvTable[tkAnd].eqvOperator       := AND_OP;
  TokenEqvTable[tkIs].eqvOperator        := IS_OP;
  TokenEqvTable[tkLike].eqvOperator      := LIKE_OP;
  TokenEqvTable[tkEqv].eqvOperator       := EQV_OP;
  TokenEqvTable[tkImp].eqvOperator       := IMP_OP;
  TokenEqvTable[tkMod].eqvOperator       := MOD_OP;
  TokenEqvTable[tkOr].eqvOperator        := OR_OP;
  TokenEqvTable[tkXor].eqvOperator       := XOR_OP;
  TokenEqvTable[tkLess].eqvOperator      := LESS_OP;
  TokenEqvTable[tkLessEq].eqvOperator    := LESS_EQ_OP;
  TokenEqvTable[tkGreater].eqvOperator   := GREATER_OP;
  TokenEqvTable[tkGreaterEq].eqvOperator := GREATER_EQ_OP;
  TokenEqvTable[tkNotEq].eqvOperator     := NOT_EQ_OP;
  TokenEqvTable[tkNot].eqvOperator       := NOT_OP;

  // Access modifiers
  TokenEqvTable[tkFriend].eqvDeclFlags     := [dfFriend];
  TokenEqvTable[tkPrivate].eqvDeclFlags    := [dfPrivate];
  TokenEqvTable[tkPublic].eqvDeclFlags     := [dfPublic];
  TokenEqvTable[tkStatic].eqvDeclFlags     := [dfStatic];
  TokenEqvTable[tkGlobal].eqvDeclFlags     := [dfGlobal];
  TokenEqvTable[tkWithEvents].eqvDeclFlags := [dfWithEvents];
  TokenEqvTable[tkDim].eqvDeclFlags        := [dfDim];
  TokenEqvTable[tkConst].eqvDeclFlags      := [dfConst];

  // Argument modifiers
  TokenEqvTable[tkByVal].eqvParamFlags      := [pfByVal];
  TokenEqvTable[tkByRef].eqvParamFlags      := [pfByRef];
  TokenEqvTable[tkOptional].eqvParamFlags   := [pfOptional];
  TokenEqvTable[tkParamArray].eqvParamFlags := [pfParamArray];

  TokenEqvTable[tkDefBool].eqvSimpleType := vbBooleanType;
  TokenEqvTable[tkDefByte].eqvSimpleType := vbByteType;
  TokenEqvTable[tkDefCur].eqvSimpleType  := vbCurrencyType;
  TokenEqvTable[tkDefDate].eqvSimpleType := vbDateType;
  TokenEqvTable[tkDefDbl].eqvSimpleType  := vbDoubleType;
  TokenEqvTable[tkDefInt].eqvSimpleType  := vbIntegerType;
  TokenEqvTable[tkDefLng].eqvSimpleType  := vbLongType;
  TokenEqvTable[tkDefObj].eqvSimpleType  := vbObjectType;
  TokenEqvTable[tkDefSng].eqvSimpleType  := vbSingleType;
  TokenEqvTable[tkDefStr].eqvSimpleType  := vbStringType;
  TokenEqvTable[tkDefVar].eqvSimpleType  := vbVariantType;

  TokenEqvTable[tkBoolean].eqvSimpleType  := vbBooleanType;
  TokenEqvTable[tkByte].eqvSimpleType     := vbByteType;
  TokenEqvTable[tkInteger].eqvSimpleType  := vbIntegerType;
  TokenEqvTable[tkLong].eqvSimpleType     := vbLongType;
  TokenEqvTable[tkSingle].eqvSimpleType   := vbSingleType;
  TokenEqvTable[tkDouble].eqvSimpleType   := vbDoubleType;
  TokenEqvTable[tkDate].eqvSimpleType     := vbDateType;
  TokenEqvTable[tkCurrency].eqvSimpleType := vbCurrencyType;
  TokenEqvTable[tkVariant].eqvSimpleType  := vbVariantType;
  TokenEqvTable[tkObject].eqvSimpleType   := vbObjectType;
  TokenEqvTable[tkAny].eqvSimpleType      := vbAnyType;

  TypeSfx2Type[STRING_SFX]   := vbStringType;
  TypeSfx2Type[INTEGER_SFX]  := vbIntegerType;
  TypeSfx2Type[LONG_SFX]     := vbLongType;
  TypeSfx2Type[SINGLE_SFX]   := vbSingleType;
  TypeSfx2Type[DOUBLE_SFX]   := vbDoubleType;
  TypeSfx2Type[CURRENCY_SFX] := vbCurrencyType;

end.
















































































































