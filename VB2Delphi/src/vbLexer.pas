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

unit vbLexer;

interface

  uses
    Contnrs,
    vbClasses,
    vbBuffers,
    vbTokens;

  type

    TvbIfCtxType = ( typIf, typIfDef, typIfnDef, typElse, typElseIf );

    TvbIfCtx =
      class
        ctxType        : TvbIfCtxType;
        ctxSkipElses   : boolean;
        ctxWasSkipping : boolean;
        ctxPrev        : TvbIfCtx;
        constructor Create( typ : TvbIfCtxType; prevCtx : TvbIfCtx = nil );
      end;

    TvbLexer =
      class
        private
          fBuffer      : TvbCustomTextBuffer;
          fBufferStack : TObjectStack;
          fText        : array[0..vbMaxLine] of char;
          fTokenColumn : cardinal;
          fTokenLine   : cardinal;
          fNextToken   : TvbToken;
          fIfCtx       : TvbIfCtx;
          fToken       : TvbToken;
          function LexTokenDirect : TvbToken;
          //procedure SkipLineDirect;
          procedure DoConst;
          procedure DoElse;
          procedure DoElseIf;
          procedure DoEndIf;
          procedure DoIf;
          function CheckToken( tok : TvbTokenKind ) : boolean;
          procedure PushIfCtx( typ : TvbIfCtxType; skip : boolean );
          procedure HandleDirective;
        public
          constructor Create;
          destructor Destroy; override;
          procedure PushBuffer( buff : TvbCustomTextBuffer );
          procedure PopBuffer;
          function GetToken : TvbToken;
          function GetFormToken : TvbToken;
          function PeekToken : TvbToken;
          function ParseConstExpr : variant;
          property TokenColumn : cardinal            read fTokenColumn;
          property TokenLine   : cardinal            read fTokenLine;
          property TopBuffer   : TvbCustomTextBuffer read fBuffer;
      end;

implementation

  uses
    Variants,
    SysUtils,
    DateUtils,
    HashList;

  const

    vbKeywords : array[tkAbs..tkXor] of TvbToken = (
        ( tokenKind : tkAbs;        tokenFlags : [];                            tokenText : 'Abs' ),
        ( tokenKind : tkAddressOf;  tokenFlags : [];                            tokenText : 'AddressOf' ),
        ( tokenKind : tkAnd;        tokenFlags : [];                            tokenText : 'And' ),
        ( tokenKind : tkAny;        tokenFlags : [TF_TYPE];                     tokenText : 'Any' ),
        ( tokenKind : tkArray;      tokenFlags : [];                            tokenText : 'Array' ),
        ( tokenKind : tkAs;         tokenFlags : [];                            tokenText : 'As' ),
        ( tokenKind : tkAttribute;  tokenFlags : [];                            tokenText : 'Attribute' ),
        ( tokenKind : tkBoolean;    tokenFlags : [TF_TYPE];                     tokenText : 'Boolean' ),
        ( tokenKind : tkByRef;      tokenFlags : [];                            tokenText : 'ByRef' ),
        ( tokenKind : tkByte;       tokenFlags : [TF_TYPE];                     tokenText : 'Byte' ),
        ( tokenKind : tkByVal;      tokenFlags : [];                            tokenText : 'ByVal' ),
        ( tokenKind : tkCall;       tokenFlags : [];                            tokenText : 'Call' ),
        ( tokenKind : tkCase;       tokenFlags : [];                            tokenText : 'Case' ),
        ( tokenKind : tkCBool;      tokenFlags : [];                            tokenText : 'CBool' ),
        ( tokenKind : tkCByte;      tokenFlags : [];                            tokenText : 'CByte' ),
        ( tokenKind : tkCDate;      tokenFlags : [];                            tokenText : 'CDate' ),
        ( tokenKind : tkCDbl;       tokenFlags : [];                            tokenText : 'CDbl' ),
        ( tokenKind : tkCDec;       tokenFlags : [];                            tokenText : 'CDec' ),
        ( tokenKind : tkCInt;       tokenFlags : [];                            tokenText : 'CInt' ),
        ( tokenKind : tkCircle;     tokenFlags : [];                            tokenText : 'Circle' ),
        ( tokenKind : tkCLng;       tokenFlags : [];                            tokenText : 'CLng' ),
        ( tokenKind : tkClose;      tokenFlags : [];                            tokenText : 'Close' ),
        ( tokenKind : tkConst;      tokenFlags : [];                            tokenText : 'Const' ),
        ( tokenKind : tkCSng;       tokenFlags : [];                            tokenText : 'CSng' ),
        ( tokenKind : tkCStr;       tokenFlags : [];                            tokenText : 'CStr' ),
        ( tokenKind : tkCurrency;   tokenFlags : [TF_TYPE];                     tokenText : 'Currency' ),
        ( tokenKind : tkCVar;       tokenFlags : [];                            tokenText : 'CVar' ),
        ( tokenKind : tkCVErr;      tokenFlags : [];                            tokenText : 'CVErr' ),
        ( tokenKind : tkDate;       tokenFlags : [TF_TYPE];                     tokenText : 'Date' ),
        ( tokenKind : tkDebug;      tokenFlags : [];                            tokenText : 'Debug' ),
        ( tokenKind : tkDecimal;    tokenFlags : [TF_TYPE];                     tokenText : 'Decimal' ),
        ( tokenKind : tkDeclare;    tokenFlags : [];                            tokenText : 'Declare' ),
        ( tokenKind : tkDefBool;    tokenFlags : [];                            tokenText : 'DefBool' ),
        ( tokenKind : tkDefByte;    tokenFlags : [];                            tokenText : 'DefByte' ),
        ( tokenKind : tkDefCur;     tokenFlags : [];                            tokenText : 'DefCur' ),
        ( tokenKind : tkDefDate;    tokenFlags : [];                            tokenText : 'DefDate' ),
        ( tokenKind : tkDefDbl;     tokenFlags : [];                            tokenText : 'DefDbl' ),
        ( tokenKind : tkDefDec;     tokenFlags : [];                            tokenText : 'DefDec' ),
        ( tokenKind : tkDefInt;     tokenFlags : [];                            tokenText : 'DefInt' ),
        ( tokenKind : tkDefLng;     tokenFlags : [];                            tokenText : 'DefLng' ),
        ( tokenKind : tkDefObj;     tokenFlags : [];                            tokenText : 'DefObj' ),
        ( tokenKind : tkDefSng;     tokenFlags : [];                            tokenText : 'DefSng' ),
        ( tokenKind : tkDefStr;     tokenFlags : [];                            tokenText : 'DefStr' ),
        ( tokenKind : tkDefVar;     tokenFlags : [];                            tokenText : 'DefVar' ),
        ( tokenKind : tkDim;        tokenFlags : [];                            tokenText : 'Dim' ),
        ( tokenKind : tkDo;         tokenFlags : [TF_EXIT_BLOCK];               tokenText : 'Do' ),
        ( tokenKind : tkDoEvents;   tokenFlags : [];                            tokenText : 'DoEvents' ),
        ( tokenKind : tkDouble;     tokenFlags : [TF_TYPE];                     tokenText : 'Double' ),
        ( tokenKind : tkEach;       tokenFlags : [];                            tokenText : 'Each' ),
        ( tokenKind : tkElse;       tokenFlags : [];                            tokenText : 'Else' ),
        ( tokenKind : tkElseIf;     tokenFlags : [];                            tokenText : 'ElseIf' ),
        ( tokenKind : tkEnd;        tokenFlags : [];                            tokenText : 'End' ),
        ( tokenKind : tkEnum;       tokenFlags : [TF_END_BLOCK];                tokenText : 'Enum' ),
        ( tokenKind : tkEqv;        tokenFlags : [];                            tokenText : 'Eqv' ),
        ( tokenKind : tkErase;      tokenFlags : [];                            tokenText : 'Erase' ),
        ( tokenKind : tkEvent;      tokenFlags : [];                            tokenText : 'Event' ),
        ( tokenKind : tkExit;       tokenFlags : [];                            tokenText : 'Exit' ),
        ( tokenKind : tkFalse;      tokenFlags : [];                            tokenText : 'False' ),
        ( tokenKind : tkFix;        tokenFlags : [];                            tokenText : 'Fix' ),
        ( tokenKind : tkFor;        tokenFlags : [TF_EXIT_BLOCK];               tokenText : 'For' ),
        ( tokenKind : tkFriend;     tokenFlags : [];                            tokenText : 'Friend' ),
        ( tokenKind : tkFunction;   tokenFlags : [TF_END_BLOCK, TF_EXIT_BLOCK]; tokenText : 'Function' ),
        ( tokenKind : tkGet;        tokenFlags : [];                            tokenText : 'Get' ),
        ( tokenKind : tkGlobal;     tokenFlags : [];                            tokenText : 'Global' ),
        ( tokenKind : tkGoSub;      tokenFlags : [];                            tokenText : 'Gosub' ),
        ( tokenKind : tkGoTo;       tokenFlags : [];                            tokenText : 'Goto' ),
        ( tokenKind : tkIf;         tokenFlags : [TF_END_BLOCK];                tokenText : 'If' ),
        ( tokenKind : tkImp;        tokenFlags : [];                            tokenText : 'Imp' ),
        ( tokenKind : tkImplements; tokenFlags : [];                            tokenText : 'Implements' ),
        ( tokenKind : tkIn;         tokenFlags : [];                            tokenText : 'In' ),
        ( tokenKind : tkInput;      tokenFlags : [];                            tokenText : 'Input' ),
        ( tokenKind : tkInputB;     tokenFlags : [];                            tokenText : 'InputB' ),
        ( tokenKind : tkInt;        tokenFlags : [];                            tokenText : 'Int' ),
        ( tokenKind : tkInteger;    tokenFlags : [TF_TYPE];                     tokenText : 'Integer' ),
        ( tokenKind : tkIs;         tokenFlags : [];                            tokenText : 'Is' ),
        ( tokenKind : tkLBound;     tokenFlags : [];                            tokenText : 'LBound' ),
        ( tokenKind : tkLen;        tokenFlags : [];                            tokenText : 'Len' ),
        ( tokenKind : tkLenB;       tokenFlags : [];                            tokenText : 'LenB' ),
        ( tokenKind : tkLet;        tokenFlags : [];                            tokenText : 'Let' ),
        ( tokenKind : tkLike;       tokenFlags : [];                            tokenText : 'Like' ),
        ( tokenKind : tkLocal;      tokenFlags : [];                            tokenText : 'Local' ),
        ( tokenKind : tkLock;       tokenFlags : [];                            tokenText : 'Lock' ),
        ( tokenKind : tkLong;       tokenFlags : [TF_TYPE];                     tokenText : 'Long' ),
        ( tokenKind : tkLoop;       tokenFlags : [];                            tokenText : 'Loop' ),
        ( tokenKind : tkLSet;       tokenFlags : [];                            tokenText : 'LSet' ),
        ( tokenKind : tkMe;         tokenFlags : [];                            tokenText : 'Me' ),
        ( tokenKind : tkMod;        tokenFlags : [];                            tokenText : 'Mod' ),
        ( tokenKind : tkNew;        tokenFlags : [];                            tokenText : 'New' ),
        ( tokenKind : tkNext;       tokenFlags : [];                            tokenText : 'Next' ),
        ( tokenKind : tkNot;        tokenFlags : [];                            tokenText : 'Not' ),
        ( tokenKind : tkNothing;    tokenFlags : [];                            tokenText : 'Nothing' ),
        ( tokenKind : tkObject;     tokenFlags : [TF_TYPE];                     tokenText : 'Object' ),
        ( tokenKind : tkOn;         tokenFlags : [];                            tokenText : 'On' ),
        ( tokenKind : tkOpen;       tokenFlags : [];                            tokenText : 'Open' ),
        ( tokenKind : tkOption;     tokenFlags : [];                            tokenText : 'Option' ),
        ( tokenKind : tkOptional;   tokenFlags : [];                            tokenText : 'Optional' ),
        ( tokenKind : tkOr;         tokenFlags : [];                            tokenText : 'Or' ),
        ( tokenKind : tkParamArray; tokenFlags : [];                            tokenText : 'ParamArray' ),
        ( tokenKind : tkPreserve;   tokenFlags : [];                            tokenText : 'Preserve' ),
        ( tokenKind : tkPrint;      tokenFlags : [];                            tokenText : 'Print' ),
        ( tokenKind : tkPrivate;    tokenFlags : [];                            tokenText : 'Private' ),
        ( tokenKind : tkProperty;   tokenFlags : [TF_END_BLOCK, TF_EXIT_BLOCK]; tokenText : 'Property' ),
        ( tokenKind : tkPSet;       tokenFlags : [];                            tokenText : 'PSet' ),
        ( tokenKind : tkPublic;     tokenFlags : [];                            tokenText : 'Public' ),
        ( tokenKind : tkPut;        tokenFlags : [];                            tokenText : 'Put' ),
        ( tokenKind : tkRaiseEvent; tokenFlags : [];                            tokenText : 'RaiseEvent' ),
        ( tokenKind : tkReDim;      tokenFlags : [];                            tokenText : 'Redim' ),
        ( tokenKind : tkREM;        tokenFlags : [];                            tokenText : 'Rem' ),
        ( tokenKind : tkResume;     tokenFlags : [];                            tokenText : 'Resume' ),
        ( tokenKind : tkReturn;     tokenFlags : [];                            tokenText : 'Return' ),
        ( tokenKind : tkRSet;       tokenFlags : [];                            tokenText : 'RSet' ),
        ( tokenKind : tkSeek;       tokenFlags : [];                            tokenText : 'Seek' ),
        ( tokenKind : tkSelect;     tokenFlags : [TF_END_BLOCK];                tokenText : 'Select' ),
        ( tokenKind : tkSet;        tokenFlags : [];                            tokenText : 'Set' ),
        ( tokenKind : tkSgn;        tokenFlags : [];                            tokenText : 'Sgn' ),
        ( tokenKind : tkShared;     tokenFlags : [];                            tokenText : 'Shared' ),
        ( tokenKind : tkSingle;     tokenFlags : [TF_TYPE];                     tokenText : 'Single' ),
        ( tokenKind : tkSpc;        tokenFlags : [];                            tokenText : 'Spc' ),
        ( tokenKind : tkStatic;     tokenFlags : [];                            tokenText : 'Static' ),
        ( tokenKind : tkStop;       tokenFlags : [];                            tokenText : 'Stop' ),
        ( tokenKind : tkString;     tokenFlags : [TF_TYPE];                     tokenText : 'String' ),
        ( tokenKind : tkSub;        tokenFlags : [TF_END_BLOCK, TF_EXIT_BLOCK]; tokenText : 'Sub' ),
        ( tokenKind : tkTab;        tokenFlags : [];                            tokenText : 'Tab' ),
        ( tokenKind : tkThen;       tokenFlags : [];                            tokenText : 'Then' ),
        ( tokenKind : tkTo;         tokenFlags : [];                            tokenText : 'To' ),
        ( tokenKind : tkTrue;       tokenFlags : [];                            tokenText : 'True' ),
        ( tokenKind : tkType;       tokenFlags : [TF_END_BLOCK];                tokenText : 'Type' ),
        ( tokenKind : tkTypeOf;     tokenFlags : [];                            tokenText : 'TypeOf' ),
        ( tokenKind : tkUBound;     tokenFlags : [];                            tokenText : 'UBound' ),
        ( tokenKind : tkUnlock;     tokenFlags : [];                            tokenText : 'Unlock' ),
        ( tokenKind : tkUntil;      tokenFlags : [];                            tokenText : 'Until' ),
        ( tokenKind : tkVariant;    tokenFlags : [TF_TYPE];                     tokenText : 'Variant' ),
        ( tokenKind : tkWend;       tokenFlags : [];                            tokenText : 'Wend' ),
        ( tokenKind : tkWhile;      tokenFlags : [];                            tokenText : 'While' ),
        ( tokenKind : tkWith;       tokenFlags : [TF_END_BLOCK];                tokenText : 'With' ),
        ( tokenKind : tkWithEvents; tokenFlags : [];                            tokenText : 'WithEvents' ),
        ( tokenKind : tkWrite;      tokenFlags : [];                            tokenText : 'Write' ),
        ( tokenKind : tkXor;        tokenFlags : [];                            tokenText : 'Xor' ) );

        // These are also keywords but we are not using them right now.
        // They are the names of VB attributes applicable to VB declarations.
        // -----------------------------------------------------------------
        // VB_Base
        // VB_Control
        // VB_Creatable
        // VB_Customizable
        // VB_Description
        // VB_Exposed
        // VB_GlobalNameSpace
        // VB_HelpID
        // VB_Invoke
        // VB_MemberFlags
        // VB_Name
        // VB_PredeclaredId
        // VB_ProcData
        // VB_PropertyPut
        // VB_PropertyPutRef
        // VB_TemplateDerived
        // VB_UserMemId
        // VB_VarDescription
        // VB_VarHelpID
        // VB_VarMemberFlags
        // VB_VarProcData
        // VB_VarUserMemId

    csHexDigits          = ['0'..'9', 'a'..'f', 'A'..'F'];
    csDigits             = ['0'..'9'];
    csIdentStart         = ['a'..'z', 'A'..'Z'];
    csEscapedIdentStart  = csIdentStart + ['_'];
    csIdentChars         = csIdentStart + ['0'..'9', '_'];
    csOctalDigits        = ['0'..'7'];

    H2D : array['0'..'f'] of byte =
      ( 0, 1, 2,  3,  4,  5,  6,  7,  8,  9, 0, 0, 0, 0, 0,
        0, 0, 10, 11, 12, 13, 14, 15, 0, 0, 0, 0, 0, 0, 0,
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
        0, 0, 0, 0, 10, 11,  12, 13, 14, 15 );

  var
    KeywordHashTable : THashList;
    Traits           : TCaseInsensitiveTraits;
    MaxKeywordLen    : integer;

  { TvbLexer }

  constructor TvbLexer.Create;
    begin
      fBufferStack := TObjectStack.Create;
      fNextToken.tokenKind := _tkUnknown;
    end;

  destructor TvbLexer.Destroy;
    begin
      assert( fBufferStack.Count = 0 );
      FreeAndNil( fBufferStack );
      inherited;
    end;

  function TvbLexer.GetToken: TvbToken;
    begin
      if fNextToken.tokenKind <> _tkUnknown
        then
          begin
            result := fNextToken;
            fNextToken.tokenKind := _tkUnknown;
          end
        else
          begin
            with fToken do
              repeat
                fToken := LexTokenDirect;
                // A '#' at BOL starts a compiler directive.
                if ( tokenKind = tkHash ) and ( TF_BOL in tokenFlags )
                  then HandleDirective;
                if not fBuffer.buffSkipping
                  then break
                  else
                if tokenKind = _tkEOF
                  then
                    if fIfCtx <> nil
                      then raise EvbSourceError.Create( 'expected #End If', fTokenLine, fTokenColumn )
                      else break;
              until false;
            result := fToken;
          end;
    end;

  function TvbLexer.LexTokenDirect : TvbToken;
    var
      ch    : char;
      pch   : pchar;
      limit : pchar;
      space : boolean;
      i     : integer;

    procedure LexComment;
      var
        p : integer;
      begin
        p := 0;
        inc( pch );
        repeat
          // if we see #13 check if this is EOB or EOL.
          // if EOB, refill the buffer and break out if EOF.
          // if this is EOL simply break out.
          if pch[0] = #13
            then
              if pch >= limit
                then
                  with fBuffer do
                    begin
                      buffPos := pch;
                      Refill;
                      pch   := buffPos;
                      limit := buffLimit;
                      if pch = limit
                        then break;
                    end
                else break;

          fText[p] := pch[0];
          inc( pch );
          inc( p );
          if p > vbMaxLine
            then raise EvbSourceError.Create( 'comment too long', fTokenLine, fTokenColumn );
        until false;
        fText[p] := #0;
        result.tokenKind := tkComment;
        result.tokenText := fText;
      end;

    procedure LexIdentOrKeyword( checkKeyword : boolean = true );
      var
        p : integer;
        K : PvbToken;
      begin
        if limit - pch < vbMaxIdent
          then
            with fBuffer do
              begin
                buffPos := pch;
                Refill;
                pch   := buffPos;
                limit := buffLimit;
              end;

        p := 0;
        repeat
          fText[p] := pch[0];
          inc( pch );
          inc( p );
          if p > vbMaxIdent
            then raise EvbSourceError.Create( 'identifier too long', fTokenLine, fTokenColumn );
        until not ( pch[0] in csIdentChars );
        fText[p] := #0;

        // check for type suffixes.
        with result do
          begin
            tokenKind := tkIdent;
            tokenText := fText;

            if checkKeyword and ( p <= MaxKeywordLen ) and ( p >= 2 ) 
              then
                begin
                  k := KeywordHashTable[fText];
                  if k <> nil
                    then
                      begin
                        tokenKind  := k^.tokenKind;
                        tokenFlags := tokenFlags + k^.tokenFlags;
                      end;
                end;

            case pch[0] of
              '%' :
                begin
                  Include( tokenFlags, TF_INTEGER_SFX );
                  inc( pch );
                end;
              '&' :
                begin
                  Include( tokenFlags, TF_LONG_SFX );
                  inc( pch );
                end;
              '@' :
                begin
                  Include( tokenFlags, TF_CURRENCY_SFX );
                  inc( pch );
                end;
              '#' :
                begin
                  Include( tokenFlags, TF_DOUBLE_SFX );
                  inc( pch );
                end;
              '$' :
                begin
                  Include( tokenFlags, TF_STRING_SFX );
                  inc( pch );
                end;
              '!' :
                if ( not ( pch[1] in csIdentStart ) ) and ( pch[1] <> '[' )  // not bang operator?
                  then
                    begin
                      Include( tokenFlags, TF_SINGLE_SFX );
                      inc( pch );
                    end;
              ':' :
                if ( tokenKind = tkIdent ) and ( TF_BOL in tokenFlags )
                  then
                    begin
                      Include( tokenFlags, TF_LABEL );
                      inc( pch );
                    end;                  
            end;
          end;
      end;

    procedure LexEscapedIdent;
      var
        p : cardinal;
      begin
        if limit - pch < vbMaxIdent
          then
            with fBuffer do
              begin
                buffPos := pch;
                Refill;
                pch   := buffPos;
                limit := buffLimit;
              end;
        p := 0;
        repeat
          fText[p] := pch[0];
          inc( pch );
          inc( p );
          if p > vbMaxIdent
            then raise EvbSourceError.Create( 'identifier too long', fTokenLine, fTokenColumn );
        until pch[0] = ']';
        inc( pch );
        fText[p] := #0;
        result.tokenKind := tkIdent;
        result.tokenText := fText;
      end;

    procedure LexStringInBrackets;
      var
        p : cardinal;
      begin
        p := 0;
        repeat
          fText[p] := pch[0];
          inc( pch );
          inc( p );
          if p > 39
            then raise EvbSourceError.Create( 'string too long', fTokenLine, fTokenColumn );
        until pch[0] = '}';
        inc( pch );
        fText[p] := #0;
        result.tokenKind := _tkBracketStr;
        result.tokenText := fText;
      end;


    procedure LexNumber;
      var
        gotPoint : boolean;
        gotExp   : boolean;
        p        : pchar;
        i        : byte;
        v        : int64;       
      begin
        p        := pch - 1;
        gotPoint := false;
        gotExp   := false;
        if p[0] = '.'  // has a deciml part?
          then gotPoint := true;
        while pch[0] in csDigits do
          inc( pch );
        if ( pch[0] = '.' ) and not gotPoint and ( pch[1] in csDigits )
          then
            begin
              gotPoint := true;
              inc( pch, 2 );  // skip '.' and first following digit
              while pch[0] in csDigits do
                inc( pch );
            end;
        with result do
          begin
            if pch[0] in ['e', 'E', 'd', 'D']  // has an exponent part?
              then
                begin
                  gotExp := true;
                  // change 'D' to 'E' because StrToFloat expects 'E'.
                  if pch[0] in ['d', 'D']
                    then
                      begin
                        pch[0] := 'e';
                        Include( tokenFlags, TF_DOUBLE_SFX );
                      end;
                  inc( pch );
                  if pch[0] in ['+', '-']  // exponent explicitly signed?
                    then inc( pch );
                  while pch[0] in csDigits do
                    inc( pch );
                end;

            move( p[0], fText[0], pch - p );
            fText[pch - p] := #0;

            if gotPoint or gotExp
              then
                begin
                  tokenKind  := tkFloatLit;
                  if not TryStrToFloat( fText, tokenFloat )
                    then raise EvbSourceError.Create( 'invalid floating-point constant', fTokenLine, fTokenColumn );
                end
              else                                
                begin
                  tokenKind := tkIntLit;
                  // fast path for common case of a single digit number
                  if fText[1] = #0
                    then tokenInt := ord( fText[0] ) - ord( '0' )
                    else
                      begin
                        i := 0;
                        v := 0;
                        repeat
                          v := v * 10 + ord( fText[i] ) - ord( '0' );
                          inc( i );
                        until fText[i] = #0;
                        tokenInt := v;
                      end;

                  // if the number occurs at BOL and optionally followed by ':'
                  // then this is a label declaration.
                  if TF_BOL in tokenFlags
                    then
                      begin
                        Include( tokenFlags, TF_LABEL );
                        if pch[0] = ':'
                          then inc( pch );
                      end;                    
                end;

            // check for type suffixes.
            case pch[0] of
              '&' :
                if tokenKind = tkIntLit
                  then
                    begin
                      Include( tokenFlags, TF_LONG_SFX );
                      inc( pch );
                    end;
              '!' :
                begin
                  Include( tokenFlags, TF_SINGLE_SFX );
                  if tokenKind = tkIntLit
                    then tokenFloat := tokenInt;
                  inc( pch );
                end;
              '#' :
                begin
                  Include( tokenFlags, TF_DOUBLE_SFX );
                  if tokenKind = tkIntLit
                    then tokenFloat := tokenInt;
                  inc( pch );
                end;
              '@' :
                begin
                  Include( tokenFlags, TF_CURRENCY_SFX );
                  if tokenKind = tkIntLit
                    then tokenFloat := tokenInt;
                  inc( pch );
                end;
            end;
          end;
      end;

    // The VB6 IDE formats date literals like: #dd/mm/yyyy[ hh:mm:ss AM|PM]#
    procedure LexDateOrHash;
      var
        dd   : integer;
        HH   : integer; // 24h format
        mm   : integer;
        min  : integer;
        ss   : integer;
        yyyy : integer;
      begin
        // NOTE: Yes, I agree! This is code "a la Ansi C"....but its fast!
        repeat
          if pch[2] <> '/'
            then break;
          dd := // day
            ( ord( pch[0] ) - ord( '0' ) ) * 10 +
            ( ord( pch[1] ) - ord( '0' ) );
          if pch[5] <> '/'
            then break;
          mm := // month
            ( ord( pch[3] ) - ord( '0' ) ) * 10 +
            ( ord( pch[4] ) - ord( '0' ) );            
          if ( pch[10] <> '#' ) and ( pch[10] <> #32 )
            then break;          
          yyyy := // year
            ( ord( pch[6] ) - ord( '0' ) ) * 1000 +
            ( ord( pch[7] ) - ord( '0' ) ) * 100 +
            ( ord( pch[8] ) - ord( '0' ) ) * 10 +
            ( ord( pch[9] ) - ord( '0' ) );

          if pch[10] = #32
            then
              begin
                if pch[22] <> '#'
                  then break;
                HH := // hour
                  ( ord( pch[11] ) - ord( '0' ) ) * 10 +
                  ( ord( pch[12] ) - ord( '0' ) );
                if pch[13] <> ':'
                  then break;
                min := // minutes
                  ( ord( pch[14] ) - ord( '0' ) ) * 10 +
                  ( ord( pch[15] ) - ord( '0' ) );
                if pch[16] <> ':'
                  then break;
                ss := // seconds
                  ( ord( pch[17] ) - ord( '0' ) ) * 10 +
                  ( ord( pch[18] ) - ord( '0' ) );
                if pch[19] <> #32
                  then break;
                // AM or PM
                if ( ( pch[20] = 'a' ) or ( pch[20] = 'A' ) ) and
                   ( ( pch[21] = 'm' ) or ( pch[21] = 'M' ) )
                  then  // nothing...
                  else
                if ( ( pch[20] = 'p' ) or ( pch[20] = 'P' ) ) and
                   ( ( pch[21] = 'm' ) or ( pch[21] = 'M' ) )
                  then inc(  HH, 12  )
                  else break;
                inc(  pch, 23  );
              end
            else
              begin
                if pch[10] <> '#'
                  then break;
                HH  := 0;
                min := 0;
                ss  := 0;
                inc( pch, 11 );
              end;              
          result.tokenKind := tkDateLit;
          if not TryEncodeDateTime( yyyy, mm, dd, HH, min, ss, 0, result.tokenDate )
            then raise EvbSourceError.Create( 'invalid date constant', fTokenLine, fTokenColumn );            
          exit;
        until false;
        result.tokenKind := tkHash;
      end;

    procedure LexString;
      var
        p : integer;
      begin
        if limit - pch < vbMaxLine
          then
            with fBuffer do
              begin
                buffPos := pch;
                Refill;
                pch   := buffPos;
                limit := buffLimit;
              end;
        inc( pch );
        p := 0;
        repeat
          if pch[0] = '"'
            then
              if pch[1] = '"'
                then inc( pch )
                else
                  begin
                    inc( pch ); // skip closing double-quote
                    break;
                  end;
          fText[p] := pch[0];
          inc( pch );
          inc( p );
          if p > vbMaxLine
            then raise EvbSourceError.Create( 'string constant too long', fTokenLine, fTokenColumn );            
        until false;
        fText[p] := #0;
        result.tokenKind := tkStringLit;
        result.tokenText := fText;
      end;

    begin
      with fBuffer, result do
        begin
          pch        := buffPos;
          limit      := buffLimit;
          space      := false;
          tokenKind  := _tkUnknown;
          tokenText  := nil;
          fTokenLine := buffLineNo;
          if buffBOL
            then
              begin
                tokenFlags := [TF_BOL];
                buffBOL := false;
              end
            else tokenFlags := [];
          repeat
            case pch[0] of
              #32, #9 : // skip whitespaces
                begin
                  space := true;
                  repeat
                    inc( pch );
                  until ( pch[0] <> #32 ) and ( pch[0] <> #9 );
                end;
              #13  : // check for end-of-line
                begin
                  // since a valid newline combination is two chars(#13#10) we
                  // have to make sure there are at least two bytes left
                  // in the buffer. if there aren't, refill the buffer.
                  if limit - pch < 2
                    then
                      begin
                        buffPos := pch;
                        Refill;
                        pch   := buffPos;
                        limit := buffLimit;
                        if pch = limit
                          then
                            begin
                              tokenKind := _tkEOF;
                              break;
                            end;
                      end
                    else
                      begin
                        inc( pch );
                        inc( buffLineNo );
                        if pch[0] = #10
                          then inc( pch );
                        buffLine := pch;
                        buffBOL  := true;
                        tokenKind := tkEOL;
                        break;
                      end;
                end;
              '_' :
                if buffFormDef
                  then
                    begin
                      LexIdentOrKeyword( false );
                      break
                    end
                  else  // advance to next line
                    begin
                      inc( pch );
                      // since we are joining the current line wih the next one
                      // I take advantage of this fact and make sure the whole
                      // next line is present in the buffer.
                      if limit - pch < vbMaxLine
                        then
                          begin
                            buffPos := pch;
                            Refill;
                            pch   := buffPos;
                            limit := buffLimit;
                            if pch = limit
                              then raise EvbSourceError.Create( 'invalid use of line concatenation at end-of-file', buffLineNo, pch - buffLine + 1 );
                          end;

                      // there shouldn't be whitespace between '_' and EOL because
                      // the VB6 IDE removes them, but I check just in case.
                      while ( pch[0] = #32 ) or ( pch[0] = #9 ) do
                        inc( pch );
                      if pch[0] <> #13
                        then raise EvbSourceError.Create( 'line concatenation must be followed by end-of-line', buffLineNo, pch - buffLine + 1 );
                      inc( pch );
                      if pch[0] = #10
                        then inc( pch );

                      buffLine := pch;
                      inc( buffLineNo );
                    end;
              else
                begin                              
                  fTokenColumn := pch - buffLine + 1;
                  case pch[0] of
                    'a'..'z', 'A'..'Z' :
                      begin
                        LexIdentOrKeyword;
                        if tokenKind = tkRem
                          then LexComment;
                      end;
                    '"'  : LexString;
                    '''' : LexComment;
                    else
                      begin
                        // all the tokens we are about to check for have a
                        // maximum length of 32 bytes, so here we make sure
                        // the buffer contains at least 40 chars.
                        if limit - pch < 40
                          then
                            begin
                              buffPos := pch;
                              Refill;
                              pch   := buffPos;
                              limit := buffLimit;
                              if pch = limit
                                then
                                  begin
                                    tokenKind := _tkEOF;
                                    break
                                  end;                                
                            end;                          
                        ch := pch[0];
                        inc( pch );
                        case ch of
                          '0'..'9' : LexNumber;
                          '(' : tokenKind := tkOParen;
                          ')' : tokenKind := tkCParen;
                          '=' : tokenKind := tkEq;
                          ',' : tokenKind := tkComma;
                          '+' : tokenKind := tkPlus;
                          '-' : tokenKind := tkMinus;
                          '*' : tokenKind := tkMult;
                          '^' : tokenKind := tkPower;
                          '/' : tokenKind := tkFloatDiv;
                          '\' : tokenKind := tkIntDiv;
                          ';' : tokenKind := tkSemiColon;
                          '!' : tokenKind := tkBang;
                          '#' : LexDateOrHash;
                          '[' : LexEscapedIdent;
                          ':' :
                            if pch[0] = '='
                              then
                                begin
                                  tokenKind := tkColonEq;
                                  inc( pch );
                                end
                              else tokenKind := tkColon;
                          '>' :
                            if pch[0] = '='
                              then
                                begin
                                  tokenKind := tkGreaterEq;
                                  inc( pch );
                                end
                              else tokenKind := tkGreater;
                          '.' :
                            begin
                              tokenKind := tkDot;
                              if pch[0] in csDigits
                                then LexNumber;
                            end;
                          '&' :
                            begin
                              tokenKind := tkConcat;
                              i := 0;
                              case pch[0] of
                                'h', 'H' :  // an hexadecimal integer 
                                  begin
                                    inc( pch );  // skip 'h' or 'H'
                                    if not ( pch[0] in csHexDigits )
                                      then raise EvbSourceError.Create( 'no hexadecimal digit after "&H"', buffLineNo, pch - buffLine + 1 );
                                    repeat
                                      i := ( i shl 4 ) + H2D[pch[0]];
                                      inc( pch );
                                      if not ( pch[0] in csHexDigits )
                                        then break;
                                      if i >= $10000000
                                        then raise EvbSourceError.Create( 'overflow in hexadecimal constant', buffLineNo, pch - buffLine + 1 );
                                    until false;

                                    tokenKind := tkIntLit;
                                    tokenInt  := i;
                                    Include( tokenFlags, TF_HEX );
                                    if pch[0] = '&'
                                      then
                                        begin
                                          Include( tokenFlags, TF_LONG_SFX );
                                          inc( pch );
                                        end;
                                  end;
                                'o', 'O' :  // an octal integer
                                  begin
                                    inc( pch ); // skip 'o' or 'O'
                                    if not ( pch[0] in csOctalDigits )
                                      then raise EvbSourceError.Create( 'no octal digit after "&O"', buffLineNo, pch - buffLine + 1 );
                                    repeat
                                      i := ( i shl 3 ) + ( ord( pch[0] ) - ord( '0' ) );
                                      inc( pch );
                                      if not ( pch[0] in csOctalDigits )
                                        then break;
                                      if i >= $20000000
                                        then raise EvbSourceError.Create( 'overflow in octal constant', buffLineNo, pch - buffLine + 1 );
                                    until false;

                                    tokenKind := tkIntLit;
                                    tokenInt  := i;
                                    Include( tokenFlags, TF_OCTAL );
                                    if pch[0] = '&'
                                      then
                                        begin
                                          Include( tokenFlags, TF_LONG_SFX );
                                          inc( pch );
                                        end;
                                  end;
                              end;
                            end;
                          '<' :
                            begin
                              tokenKind := tkLess;
                              case pch[0] of
                                '>' :
                                  begin
                                    tokenKind := tkNotEq;
                                    inc( pch );
                                  end;
                                '=' :
                                  begin
                                    tokenKind := tkLessEq;
                                    inc( pch );
                                  end;
                              end;
                            end;
                          '$' : tokenKind := tkDollar;
                          '{' : LexStringInBrackets;
                        end;
                      end;
                  end;
                  break;  // token lexed, break out the loop.
                end;
            end;
          until false;
          buffPos := pch;
          if space
            then Include( tokenFlags, TF_SPACED );
        end;
    end;

  function TvbLexer.PeekToken : TvbToken;
    begin
      if fNextToken.tokenKind <> _tkUnknown
        then result := fNextToken
        else
          begin
            result := GetToken;
            fNextToken := result;
          end
    end;

  procedure TvbLexer.PopBuffer;
    begin
      if fBufferStack.Count > 0
        then fBuffer := TvbCustomTextBuffer( fBufferStack.Pop )
        else fBuffer := nil;
    end;

  procedure TvbLexer.PushBuffer( buff : TvbCustomTextBuffer );
    begin
      if fBuffer <> nil
        then fBufferStack.Push( fBuffer );
      fBuffer := buff;
    end;

  procedure CreateKeywordHashTable;
    var
      i : TvbTokenKind;
      p : PvbToken;
    begin
      MaxKeywordLen    := 0;
      Traits           := nil;
      KeywordHashTable := nil;
      try
        Traits           := TCaseInsensitiveTraits.Create;
        KeywordHashTable := THashList.Create( Traits, length( vbKeywords ) ) ;
        for i := low( vbKeywords ) to high( vbKeywords ) do
          begin
            if length( vbKeywords[i].tokenText ) > MaxKeywordLen
              then MaxKeywordLen := length( vbKeywords[i].tokenText );              
            p := @vbKeywords[i];
            KeywordHashTable.Add( vbKeywords[i].tokenText, p );
          end;
      except
        FreeAndNil( KeywordHashTable );
        FreeAndNil( Traits );
        raise;
      end;
    end;

  procedure DestroyKeywordHashTable;
    begin
      FreeAndNil( Traits );
      FreeAndNil( KeywordHashTable );
    end;

//  procedure TvbLexer.SkipLineDirect;
//    var
//      pch   : pchar;
//      limit : pchar;
//    begin
//      with fBuffer do
//        begin
//          pch := buffPos;
//          limit := buffLimit;
//          repeat
//            if pch[0] = #13
//              then
//                begin
//                  // if we reached the EOB refill the buffer.
//                  if pch >= limit
//                    then
//                      begin
//                        buffPos := pch;
//                        Refill;
//                        pch   := buffPos;
//                        limit := buffLimit;
//                        if pch = limit
//                          then break;
//                      end
//                    else  // we reached end-of-line
//                      begin
//                        inc( buffLineNo );
//                        inc( pch );
//                        if pch[0] = #10
//                          then inc( pch );
//                        buffLine := pch;
//                        buffBOL  := true;
//                        break;
//                      end;
//                end
//              else inc( pch );
//          until false;
//          buffPos := pch;
//        end;
//    end;

  procedure TvbLexer.DoIf;
    var
      skip : boolean;
    begin
      skip := true;
      fToken := LexTokenDirect;
      if not fBuffer.buffSkipping
        then
          begin
            skip := ( ParseConstExpr = 0 );
            if not CheckToken( tkThen )
              then raise EvbSourceError.Create( 'expected "Then"', fTokenLine, fTokenColumn );
          end;    
      PushIfCtx( typIf, skip );
    end;

  procedure TvbLexer.DoElseIf;
    begin
      fToken := LexTokenDirect;
      if fIfCtx = nil
        then raise EvbSourceError.Create( '#Else If without #If', fTokenLine, fTokenColumn );
      if fIfCtx.ctxType = typElse
        then raise EvbSourceError.Create( '#Else If after #Else', fTokenLine, fTokenColumn );
      fIfCtx.ctxType := typElseIf;
      if fIfCtx.ctxSkipElses
        then fBuffer.buffSkipping := true
        else
          begin
            fBuffer.buffSkipping := ( ParseConstExpr = 0 );
            if not CheckToken( tkThen )
              then raise EvbSourceError.Create( 'expected "Then"', fTokenLine, fTokenColumn );
            fIfCtx.ctxSkipElses := not fBuffer.buffSkipping;
          end;
    end;

  procedure TvbLexer.DoElse;
    begin
      fToken := LexTokenDirect;
      if fIfCtx = nil
        then raise EvbSourceError.Create( '#Else without #If', fTokenLine, fTokenColumn );
      if fIfCtx.ctxType = typElse
        then raise EvbSourceError.Create( '#Else after #Else', fTokenLine, fTokenColumn );
      fIfCtx.ctxType := typElse;
      fBuffer.buffSkipping := fIfCtx.ctxSkipElses;
      fIfCtx.ctxSkipElses := true;
    end;

  procedure TvbLexer.DoEndIf;
    var
      ctx : TvbIfCtx;
    begin
      if fIfCtx = nil
        then raise EvbSourceError.Create( '#End If without #If', fTokenLine, fTokenColumn );
      ctx := fIfCtx;
      fIfCtx := fIfCtx.ctxPrev;
      fBuffer.buffSkipping := ctx.ctxWasSkipping;
      ctx.Free;
    end;

  procedure TvbLexer.PushIfCtx( typ : TvbIfCtxType; skip : boolean );
    var
      newCtx : TvbIfCtx;
    begin
      newCtx := TvbIfCtx.Create( typ, fIfCtx );
      newCtx.ctxSkipElses := fBuffer.buffSkipping or not skip;
      newCtx.ctxWasSkipping := fBuffer.buffSkipping;
      fBuffer.buffSkipping := skip;
      fIfCtx := newCtx;
    end;

  procedure TvbLexer.DoConst;
    var
      name : string;
      t    : TvbToken;
      val  : variant;
    begin
      t := LexTokenDirect;
      with t do
        begin
          if tokenKind <> tkIdent
            then raise EvbSourceError.Create( 'constant name expected', fTokenLine, fTokenColumn );
          name := tokenText;
          t := LexTokenDirect;
          if tokenKind <> tkEq
            then raise EvbSourceError.Create( 'expected "="', fTokenLine, fTokenColumn );
          fToken := LexTokenDirect;
          val := ParseConstExpr;
          //TODO:SetConstValue
        end;
    end;

  function TvbLexer.ParseConstExpr : variant;

    function ParseLogicalOrExpr : variant; forward;

    function ParsePrimaryExpr : variant;
      begin
        result := 0;
        with fToken do
          case tokenKind of
            tkIntLit    : result := tokenInt;
            tkFloatLit  : result := tokenFloat;
            tkDateLit   : result := tokenDate;
            tkStringLit : result := string( tokenText );
            tkNothing   : result := 0;
            tkTrue      : result := true;
            tkFalse     : result := false;
            tkOParen :
              begin
                fToken := LexTokenDirect;
                result := ParseLogicalOrExpr;
                if tokenKind <> tkCParen
                  then EvbSourceError.Create( 'expected ")"', fTokenLine, fTokenColumn );
              end;
            tkIdent :
              if StrIComp( 'Win32', tokenText ) = 0
                then result := true
                else result := 0;
            //TODO: Abs, Int, Fix,...
            else raise EvbSourceError.Create( 'invalid constant expression', fTokenLine, fTokenColumn );
          end;
          fToken := LexTokenDirect;
      end;

    function ParseUnaryExpr : variant;
      begin
        if CheckToken( tkPlus )
          then result := +ParsePrimaryExpr
          else
        if CheckToken( tkMinus )
          then result := -ParsePrimaryExpr
          else
        if CheckToken( tkNot )
          then result := not ParsePrimaryExpr
          else result := ParsePrimaryExpr;
      end;

    function ParseMultExpr : variant;
      begin
        result := ParseUnaryExpr;
        while true do
          if CheckToken( tkMult )
            then result := result * ParseUnaryExpr
            else
          if CheckToken( tkIntDiv )
            then result := result div ParseUnaryExpr
            else
          if CheckToken( tkFloatDiv )
            then result := result / ParseUnaryExpr
            else
          if CheckToken( tkMod )
            then result := result mod ParseUnaryExpr
            else break;
      end;

    function ParseAddExpr : variant;
      begin
        result := ParseMultExpr;
        while true do
          if CheckToken( tkConcat )
            then result := string( result ) + string( ParseMultExpr )
            else
          if CheckToken( tkPlus )
            then result := result + ParseMultExpr
            else
          if CheckToken( tkMinus )
            then result := result - ParseMultExpr
            else break;
      end;

    function GetRelationalExpr : variant;
      begin
        result := ParseAddExpr;
        if CheckToken( tkLess )
          then result := result < ParseAddExpr
          else
        if CheckToken( tkGreater )
          then result := result > ParseAddExpr
          else
        if CheckToken( tkLessEq )
          then result := result <= ParseAddExpr
          else
        if CheckToken( tkGreaterEq )
          then result := result >= ParseAddExpr
          else
        if CheckToken( tkEq )
          then result := result = ParseAddExpr
          else
        if CheckToken( tkNotEq )
          then result := result <> ParseAddExpr;
      end;

    function GetLogicalANDExpr : variant;
      begin
        result := GetRelationalExpr;
        while CheckToken( tkAnd ) do
          result := result and GetRelationalExpr;
      end;
      
    function ParseLogicalOrExpr : variant;
      begin
        result := GetLogicalANDExpr;
        while CheckToken( tkOr ) do
          result := result and GetLogicalANDExpr;
      end;

    begin
      try
        result := ParseLogicalOrExpr;
      except
        on e : EVariantError do
          raise EvbSourceError.Create( 'invalid constant expression operands', fTokenLine, fTokenColumn );
        else raise;
      end;
    end;

  function TvbLexer.CheckToken( tok : TvbTokenKind ) : boolean;
    begin
      if fToken.tokenKind = tok
        then
          begin
            fToken := LexTokenDirect;
            result := true;
          end
        else result := false;
    end;

  procedure TvbLexer.HandleDirective;
    begin
      fToken := LexTokenDirect;
      with fToken do
        case tokenKind of
          tkConst  : DoConst;
          tkIf     : DoIf;
          tkElse   : DoElse;
          tkElseIf : DoElseIf;
          tkEnd :
            begin
              fToken := LexTokenDirect;
              if not CheckToken( tkIf )
                then raise EvbSourceError.Create( 'expected "If"', fTokenLine, fTokenColumn );
              DoEndIf;
            end;
          else raise EvbSourceError.Create( 'expected "Const", "If", "Else If", "Else" or "End If" ', fTokenLine, fTokenColumn );
        end;
      // if there is a comment at EOL skip it.
      CheckToken( tkComment );
    end;

  function TvbLexer.GetFormToken : TvbToken;
    var
      t1 : TvbToken;

    function LexResOffset : cardinal;
      begin
        result := 0;
        with fBuffer do
          while buffPos[0] in csHexDigits do
            begin
              result := result shl 4 + H2D[buffPos[0]];
              inc( buffPos );
            end;
      end;

    begin
      with fBuffer, result do
        begin
          tokenKind := _tkUnknown;
          t1 := LexTokenDirect;
          case t1.tokenKind of
            tkMinus :
              begin
                t1 := LexTokenDirect;
                if t1.tokenKind = tkIntLit
                  then tokenInt := -t1.tokenInt
                  else
                if t1.tokenKind = tkFloatLit
                  then tokenFloat := -t1.tokenFloat
                  else raise EvbSourceError.Create( 'invalid expression', fTokenLine, fTokenColumn );
                tokenKind := t1.tokenKind;
              end;
            tkDollar :
              begin
                t1 := LexTokenDirect;
                if t1.tokenKind <> tkStringLit
                  then exit;
                if buffPos[0] <> ':'
                  then exit;
                inc( buffPos );
                tokenFormRes := LexResOffset;
                tokenKind := _tkFormRes;
              end;
            tkPower :
              begin
                t1 := LexTokenDirect;
                tokenKind := _tkFormCtrl;
                tokenText := t1.tokenText;
              end;
            tkStringLit :
              begin
                if buffPos[0] = ':'
                  then
                    begin
                      inc( buffPos );
                      tokenFormRes := LexResOffset;
                      tokenKind := _tkFormRes;
                    end
                  else result := t1;
              end;
            else result := t1;
          end;
        end;
    end;

  { TvbIfCtx }

  constructor TvbIfCtx.Create( typ : TvbIfCtxType; prevCtx : TvbIfCtx );
    begin
      ctxType := typ;
      ctxPrev := prevCtx;
    end;

initialization
  CreateKeywordHashTable;

finalization
  DestroyKeywordHashTable;

end.




























