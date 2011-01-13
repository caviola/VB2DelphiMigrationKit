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
    vbClasses,
    Buffers,
    vbTokens;

  type

    TvbLexer =
      class
        private
          fLineNum : cardinal;
          fCurrLine : pchar;
          fBuffer    : TvbCustomTextBuffer;
          fNextToken : TvbToken;
          fBOL       : boolean; // buffer pointer at begining-of-line?
          procedure Fail( const msg : string );
          procedure SetBuffer( buff : TvbCustomTextBuffer );
        public
          constructor Create;
          procedure GetPosition( var line, column : cardinal );
          function GetToken : TvbToken;
          function PeekToken : TvbToken;
          property Buffer : TvbCustomTextBuffer read fBuffer write SetBuffer;
      end;

implementation

  uses
    magoConsts,
    Classes,
    NumBaseUtils,
    SysUtils,
    DateUtils,
    HashList;

  const

    vbKeywords : array[tkAbs..tkXor] of TvbToken = (
        ( tokenKind : tkAbs;        tokenText : 'Abs';        tokenFlags : [] ),
        ( tokenKind : tkAddressOf;  tokenText : 'AddressOf';  tokenFlags : [] ),
        ( tokenKind : tkAnd;        tokenText : 'And';        tokenFlags : [] ),
        ( tokenKind : tkAny;        tokenText : 'Any';        tokenFlags : [tfBuiltInType] ),
        ( tokenKind : tkArray;      tokenText : 'Array';      tokenFlags : [] ),
        ( tokenKind : tkAs;         tokenText : 'As';         tokenFlags : [] ),
        ( tokenKind : tkAttribute;  tokenText : 'Attribute';  tokenFlags : [] ),
        ( tokenKind : tkBoolean;    tokenText : 'Boolean';    tokenFlags : [tfBuiltInType] ),
        ( tokenKind : tkByRef;      tokenText : 'ByRef';      tokenFlags : [] ),
        ( tokenKind : tkByte;       tokenText : 'Byte';       tokenFlags : [tfBuiltInType] ),
        ( tokenKind : tkByVal;      tokenText : 'ByVal';      tokenFlags : [] ),
        ( tokenKind : tkCall;       tokenText : 'Call';       tokenFlags : [] ),
        ( tokenKind : tkCase;       tokenText : 'Case';       tokenFlags : [] ),
        ( tokenKind : tkCBool;      tokenText : 'CBool';      tokenFlags : [] ),
        ( tokenKind : tkCByte;      tokenText : 'CByte';      tokenFlags : [] ),
        ( tokenKind : tkCDate;      tokenText : 'CDate';      tokenFlags : [] ),
        ( tokenKind : tkCDbl;       tokenText : 'CDbl';       tokenFlags : [] ),
        ( tokenKind : tkCDec;       tokenText : 'CDec';       tokenFlags : [] ),
        ( tokenKind : tkCInt;       tokenText : 'CInt';       tokenFlags : [] ),
        ( tokenKind : tkCircle;     tokenText : 'Circle';     tokenFlags : [] ),
        ( tokenKind : tkCLng;       tokenText : 'CLng';       tokenFlags : [] ),
        ( tokenKind : tkClose;      tokenText : 'Close';      tokenFlags : [] ),
        ( tokenKind : tkConst;      tokenText : 'Const';      tokenFlags : [] ),
        ( tokenKind : tkCSng;       tokenText : 'CSng';       tokenFlags : [] ),
        ( tokenKind : tkCStr;       tokenText : 'CStr';       tokenFlags : [] ),
        ( tokenKind : tkCurrency;   tokenText : 'Currency';   tokenFlags : [tfBuiltInType] ),
        ( tokenKind : tkCVar;       tokenText : 'CVar';       tokenFlags : [] ),
        ( tokenKind : tkCVErr;      tokenText : 'CVErr';      tokenFlags : [] ),
        ( tokenKind : tkDate;       tokenText : 'Date';       tokenFlags : [tfBuiltInType] ),
        ( tokenKind : tkDebug;      tokenText : 'Debug';      tokenFlags : [] ),
        ( tokenKind : tkDecimal;    tokenText : 'Decimal';    tokenFlags : [tfBuiltInType] ),
        ( tokenKind : tkDeclare;    tokenText : 'Declare';    tokenFlags : [] ),
        ( tokenKind : tkDefBool;    tokenText : 'DefBool';    tokenFlags : [] ),
        ( tokenKind : tkDefByte;    tokenText : 'DefByte';    tokenFlags : [] ),
        ( tokenKind : tkDefCur;     tokenText : 'DefCur';     tokenFlags : [] ),
        ( tokenKind : tkDefDate;    tokenText : 'DefDate';    tokenFlags : [] ),
        ( tokenKind : tkDefDbl;     tokenText : 'DefDbl';     tokenFlags : [] ),
        ( tokenKind : tkDefDec;     tokenText : 'DefDec';     tokenFlags : [] ),
        ( tokenKind : tkDefInt;     tokenText : 'DefInt';     tokenFlags : [] ),
        ( tokenKind : tkDefLng;     tokenText : 'DefLng';     tokenFlags : [] ),
        ( tokenKind : tkDefObj;     tokenText : 'DefObj';     tokenFlags : [] ),
        ( tokenKind : tkDefSng;     tokenText : 'DefSng';     tokenFlags : [] ),
        ( tokenKind : tkDefStr;     tokenText : 'DefStr';     tokenFlags : [] ),
        ( tokenKind : tkDefVar;     tokenText : 'DefVar';     tokenFlags : [] ),
        ( tokenKind : tkDim;        tokenText : 'Dim';        tokenFlags : [] ),
        ( tokenKind : tkDo;         tokenText : 'Do';         tokenFlags : [tfExitBlock] ),
        ( tokenKind : tkDoEvents;   tokenText : 'DoEvents';   tokenFlags : [] ),
        ( tokenKind : tkDouble;     tokenText : 'Double';     tokenFlags : [tfBuiltInType] ),
        ( tokenKind : tkEach;       tokenText : 'Each';       tokenFlags : [] ),
        ( tokenKind : tkElse;       tokenText : 'Else';       tokenFlags : [] ),
        ( tokenKind : tkElseIf;     tokenText : 'ElseIf';     tokenFlags : [] ),
        ( tokenKind : tkEnd;        tokenText : 'End';        tokenFlags : [] ),
        ( tokenKind : tkEnum;       tokenText : 'Enum';       tokenFlags : [tfEndBlock] ),
        ( tokenKind : tkEqv;        tokenText : 'Eqv';        tokenFlags : [] ),
        ( tokenKind : tkErase;      tokenText : 'Erase';      tokenFlags : [] ),
        ( tokenKind : tkEvent;      tokenText : 'Event';      tokenFlags : [] ),
        ( tokenKind : tkExit;       tokenText : 'Exit';       tokenFlags : [] ),
        ( tokenKind : tkFalse;      tokenText : 'False';      tokenFlags : [] ),
        ( tokenKind : tkFix;        tokenText : 'Fix';        tokenFlags : [] ),
        ( tokenKind : tkFor;        tokenText : 'For';        tokenFlags : [tfExitBlock] ),
        ( tokenKind : tkFriend;     tokenText : 'Friend';     tokenFlags : [] ),
        ( tokenKind : tkFunction;   tokenText : 'Function';   tokenFlags : [tfEndBlock, tfExitBlock] ),
        ( tokenKind : tkGet;        tokenText : 'Get';        tokenFlags : [] ),
        ( tokenKind : tkGlobal;     tokenText : 'Global';     tokenFlags : [] ),
        ( tokenKind : tkGoSub;      tokenText : 'Gosub';      tokenFlags : [] ),
        ( tokenKind : tkGoTo;       tokenText : 'Goto';       tokenFlags : [] ),
        ( tokenKind : tkIf;         tokenText : 'If';         tokenFlags : [tfEndBlock] ),
        ( tokenKind : tkImp;        tokenText : 'Imp';        tokenFlags : [] ),
        ( tokenKind : tkImplements; tokenText : 'Implements'; tokenFlags : [] ),
        ( tokenKind : tkIn;         tokenText : 'In';         tokenFlags : [] ),
        ( tokenKind : tkInput;      tokenText : 'Input';      tokenFlags : [] ),
        ( tokenKind : tkInputB;     tokenText : 'InputB';     tokenFlags : [] ),
        ( tokenKind : tkInt;        tokenText : 'Int';        tokenFlags : [] ),
        ( tokenKind : tkInteger;    tokenText : 'Integer';    tokenFlags : [tfBuiltInType] ),
        ( tokenKind : tkIs;         tokenText : 'Is';         tokenFlags : [] ),
        ( tokenKind : tkLBound;     tokenText : 'LBound';     tokenFlags : [] ),
        ( tokenKind : tkLen;        tokenText : 'Len';        tokenFlags : [] ),
        ( tokenKind : tkLenB;       tokenText : 'LenB';       tokenFlags : [] ),
        ( tokenKind : tkLet;        tokenText : 'Let';        tokenFlags : [] ),
        ( tokenKind : tkLike;       tokenText : 'Like';       tokenFlags : [] ),
        ( tokenKind : tkLocal;      tokenText : 'Local';      tokenFlags : [] ),
        ( tokenKind : tkLock;       tokenText : 'Lock';       tokenFlags : [] ),
        ( tokenKind : tkLong;       tokenText : 'Long';       tokenFlags : [tfBuiltInType] ),
        ( tokenKind : tkLoop;       tokenText : 'Loop';       tokenFlags : [] ),
        ( tokenKind : tkLSet;       tokenText : 'LSet';       tokenFlags : [] ),
        ( tokenKind : tkMe;         tokenText : 'Me';         tokenFlags : [] ),
        ( tokenKind : tkMod;        tokenText : 'Mod';        tokenFlags : [] ),
        ( tokenKind : tkNew;        tokenText : 'New';        tokenFlags : [] ),
        ( tokenKind : tkNext;       tokenText : 'Next';       tokenFlags : [] ),
        ( tokenKind : tkNot;        tokenText : 'Not';        tokenFlags : [] ),
        ( tokenKind : tkNothing;    tokenText : 'Nothing';    tokenFlags : [] ),
        ( tokenKind : tkObject;     tokenText : 'Object';     tokenFlags : [tfBuiltInType] ),
        ( tokenKind : tkOn;         tokenText : 'On';         tokenFlags : [] ),
        ( tokenKind : tkOpen;       tokenText : 'Open';       tokenFlags : [] ),
        ( tokenKind : tkOption;     tokenText : 'Option';     tokenFlags : [] ),
        ( tokenKind : tkOptional;   tokenText : 'Optional';   tokenFlags : [] ),
        ( tokenKind : tkOr;         tokenText : 'Or';         tokenFlags : [] ),
        ( tokenKind : tkParamArray; tokenText : 'ParamArray'; tokenFlags : [] ),
        ( tokenKind : tkPreserve;   tokenText : 'Preserve';   tokenFlags : [] ),
        ( tokenKind : tkPrint;      tokenText : 'Print';      tokenFlags : [] ),
        ( tokenKind : tkPrivate;    tokenText : 'Private';    tokenFlags : [] ),
        ( tokenKind : tkProperty;   tokenText : 'Property';   tokenFlags : [tfEndBlock, tfExitBlock] ),
        ( tokenKind : tkPSet;       tokenText : 'PSet';       tokenFlags : [] ),
        ( tokenKind : tkPublic;     tokenText : 'Public';     tokenFlags : [] ),
        ( tokenKind : tkPut;        tokenText : 'Put';        tokenFlags : [] ),
        ( tokenKind : tkRaiseEvent; tokenText : 'RaiseEvent'; tokenFlags : [] ),
        ( tokenKind : tkReDim;      tokenText : 'Redim';      tokenFlags : [] ),
        ( tokenKind : tkREM;        tokenText : 'Rem';        tokenFlags : [] ),
        ( tokenKind : tkResume;     tokenText : 'Resume';     tokenFlags : [] ),
        ( tokenKind : tkReturn;     tokenText : 'Return';     tokenFlags : [] ),
        ( tokenKind : tkRSet;       tokenText : 'RSet';       tokenFlags : [] ),
        ( tokenKind : tkSeek;       tokenText : 'Seek';       tokenFlags : [] ),
        ( tokenKind : tkSelect;     tokenText : 'Select';     tokenFlags : [tfEndBlock] ),
        ( tokenKind : tkSet;        tokenText : 'Set';        tokenFlags : [] ),
        ( tokenKind : tkSgn;        tokenText : 'Sgn';        tokenFlags : [] ),
        ( tokenKind : tkShared;     tokenText : 'Shared';     tokenFlags : [] ),
        ( tokenKind : tkSingle;     tokenText : 'Single';     tokenFlags : [tfBuiltInType] ),
        ( tokenKind : tkSpc;        tokenText : 'Spc';        tokenFlags : [] ),
        ( tokenKind : tkStatic;     tokenText : 'Static';     tokenFlags : [] ),
        ( tokenKind : tkStop;       tokenText : 'Stop';       tokenFlags : [] ),
        ( tokenKind : tkString;     tokenText : 'String';     tokenFlags : [tfBuiltInType] ),
        ( tokenKind : tkSub;        tokenText : 'Sub';        tokenFlags : [tfEndBlock, tfExitBlock] ),
        ( tokenKind : tkTab;        tokenText : 'Tab';        tokenFlags : [] ),
        ( tokenKind : tkThen;       tokenText : 'Then';       tokenFlags : [] ),
        ( tokenKind : tkTo;         tokenText : 'To';         tokenFlags : [] ),
        ( tokenKind : tkTrue;       tokenText : 'True';       tokenFlags : [] ),
        ( tokenKind : tkType;       tokenText : 'Type';       tokenFlags : [tfEndBlock] ),
        ( tokenKind : tkTypeOf;     tokenText : 'TypeOf';     tokenFlags : [] ),
        ( tokenKind : tkUBound;     tokenText : 'UBound';     tokenFlags : [] ),
        ( tokenKind : tkUnlock;     tokenText : 'Unlock';     tokenFlags : [] ),
        ( tokenKind : tkUntil;      tokenText : 'Until';      tokenFlags : [] ),
        ( tokenKind : tkVariant;    tokenText : 'Variant';    tokenFlags : [tfBuiltInType] ),
        ( tokenKind : tkWend;       tokenText : 'Wend';       tokenFlags : [] ),
        ( tokenKind : tkWhile;      tokenText : 'While';      tokenFlags : [] ),
        ( tokenKind : tkWith;       tokenText : 'With';       tokenFlags : [tfEndBlock] ),
        ( tokenKind : tkWithEvents; tokenText : 'WithEvents'; tokenFlags : [] ),
        ( tokenKind : tkWrite;      tokenText : 'Write';      tokenFlags : [] ),
        ( tokenKind : tkXor;        tokenText : 'Xor';        tokenFlags : [] ) );

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

  var
    vbKeywordHashTable : THashList;
    Traits             : TCaseInsensitiveTraits;

  { TvbLexer }

  constructor TvbLexer.Create;
    begin
      SetBuffer( nil );
    end;

  procedure TvbLexer.SetBuffer( buff : TvbCustomTextBuffer );
    begin
      fLineNum := 0;
      if buff <> nil
        then fCurrLine := buff.buffPos;
      fBuffer := buff;
      fBOL := true;
      fNextToken.tokenKind := _tkUnknown;
    end;

  function TvbLexer.GetToken : TvbToken;
    var
      token : TvbToken absolute result;
      c     : char;
      p     : pchar;

    // The VB6 IDE formats date literals like this: #dd/mm/yyyy[ hh:mm:ss AM|PM]#
    //
    procedure LexDateOrHash;
      var
        day,
        month,
        year,
        hour,
        min,
        sec   : integer;
        p     : pchar;
        isAM  : boolean; // true if AM, false if PM
      label
        hash;

      function ReadDatePart( out aDatePart :  integer; aMaxLen : integer ) : boolean;
        var
          p : pchar;
        begin
          result := false;
          with fBuffer do
            if buffPos[0] in csDigits
              then
                begin
                  result := true;
                  p := buffPos;
                  repeat
                    inc( buffPos );
                  until not ( ( buffPos[0] in csDigits ) and ( ( buffPos - p + 1) <= aMaxLen ) );
                  aDatePart := StrToInt( StrLPas( p ) );
                end;
        end;

      begin
        with fBuffer do
          begin
            p := buffPos;
            if not ReadDatePart( day, 2 ) then goto hash;
            if buffPos[0] <> '/' then goto hash;
            inc( buffPos );
            if not ReadDatePart( month, 2 ) then goto hash;
            if buffPos[0] <> '/' then goto hash;
            inc( buffPos );
            if not ReadDatePart( year, 4 ) then goto hash;
            if buffPos[0] = #32
              then
                begin
                  inc( buffPos );
                  if not ReadDatePart( hour, 2 ) then goto hash;
                  if buffPos[0] <> ':' then goto hash;
                  inc( buffPos );
                  if not ReadDatePart( min, 2 ) then goto hash;
                  if buffPos[0] <> ':' then goto hash;
                  inc( buffPos );
                  if not ReadDatePart( sec, 2 ) then goto hash;
                  if buffPos[0] <> #32 then goto hash;
                  inc( buffPos );
                  if ( ( buffPos[0] = 'a' ) or ( buffPos[0] = 'A' ) ) and
                     ( ( buffPos[1] = 'm' ) or ( buffPos[1] = 'M' ) )
                    then isAM := true
                    else
                      if ( ( buffPos[0] = 'p' ) or ( buffPos[0] = 'P' ) ) and
                         ( ( buffPos[1] = 'm' ) or ( buffPos[1] = 'M' ) )
                        then isAM := false
                        else goto hash;
                  inc( buffPos, 2 );  // skip 'AM' or 'PM'
                end
              else
                begin
                  hour := 0;
                  min := 0;
                  sec := 0;
                  isAM := true;
                end;
            if buffPos[0] <> '#' then goto hash;
            inc( buffPos );
            if not isAM
              then inc( hour, 12 );
            token.tokenKind := tkDateLit;
            if not TryEncodeDateTime( year, month, day, hour, min, sec, 0, token.tokenDate )
              then Fail( 'invalid date literal' );
            exit;
          hash:
            token.tokenKind := tkHash;
            buffPos := p;
          end;
      end;

    procedure ReadIdent;
      var
        p : pchar;
      begin
        with fBuffer, token do
          begin
            tokenKind := tkIdent;
            p := buffPos - 1;
            while buffPos[0] in csIdentChars do
              inc( buffPos );
            tokenText := StrLPas( p );
          end;
      end;

    procedure LexEscapedIdent;
      begin
        with fBuffer do
          begin
            inc( buffPos );
            if buffPos[0] in csIdentStart
              then
                begin
                  ReadIdent;
                  Include( token.tokenFlags, tfEscapedIdent );
                  if buffPos[0] = ']'
                    then inc( buffPos );
                end;
          end;
      end;

    procedure LexString;
      var
        i : integer;  // 1-based index of first quote within string 
        s : pchar;
      begin
        i := -1;
        with fBuffer, token do
          begin
            tokenKind := tkStringLit;
            s := buffPos;
            repeat
              while not ( buffPos[0] in ['"', #13] ) do
                inc( buffPos );              
              if ( buffPos[0] = '"' ) and ( buffPos[1] = '"' )
                then
                  begin
                    inc( buffPos, 2 );
                    if i = -1
                      then i := buffPos - s;
                  end
                else break;
            until false;
            tokenText := StrLPas( s );
            inc( buffPos );
            if i <> -1  // string contains double quotes?
              then
                while i < length( tokenText ) do
                  begin
                    if tokenText[i] = '"'
                      then delete( tokenText, i, 1 );  // remove one of the two quotes
                    inc( i );
                  end;
          end;
      end;

    procedure LexNumber;
      var
        gotPoint,
        gotExp    : boolean;
        p         : pchar;
        s         : string;

      procedure CoerceToFloat( flg : TvbTokenFlag );
        begin
          Include( token.tokenFlags, flg );
          inc( fBuffer.buffPos );
          with token do
            if tokenKind = tkIntLit
              then
                begin
                  tokenKind  := tkFloatLit;
                  tokenFloat := tokenInt;
                  Exclude( tokenFlags, tfDec );
                end;
        end;

      begin
        with fBuffer, token do
          begin
            if c = '&'  // an hex or octal literal?
              then
                begin
                  tokenKind := tkIntLit;
                  c := buffPos[0]; // now c has 'H' or 'O'
                  inc( buffPos );  // skip 'H' or 'O'
                  p := buffPos;    // start pos of hex or octal literal
                  case c of
                    'h', 'H' :
                      begin
                        Include( tokenFlags, tfHex );
                        while buffPos[0] in csHexDigits do
                          inc( buffPos );
                        tokenInt := Hex2Dec( StrLPas( p ) );
                      end;
                    'o', 'O' :
                      begin
                        Include( tokenFlags, tfOctal );
                        while buffPos[0] in csOctalDigits do
                          inc( buffPos );
                        tokenInt := Oct2Dec( StrLPas( p ) );
                      end;
                  end;
                end
              else  // read a decimal or floating literal
                begin
                  p := buffPos - 1;
                  gotPoint := false;
                  gotExp := false;
                  if c = '.'  // has a deciml part?
                    then gotPoint := true;
                  while buffPos[0] in csDigits do
                    inc( buffPos );
                  if ( buffPos[0] = '.' ) and not gotPoint and ( buffPos[1] in csDigits )
                    then
                      begin
                        gotPoint := true; 
                        inc( buffPos );
                        while buffPos[0] in csDigits do
                          inc( buffPos );
                      end;
                  if buffPos[0] in ['e', 'E', 'd', 'D']  // has an exponent part?
                    then
                      begin
                        // HACK! StrToFloat doesn't support 'D' in the exponent part
                        // so we always change it to 'E'.
                        //
                        gotExp := true;
                        if buffPos[0] in ['d', 'D']
                          then
                            begin
                              buffPos[0] := 'e';
                              Include( tokenFlags, tfDoubleType );
                            end;
                        inc( buffPos );
                        if buffPos[0] in ['+', '-']  // exponent explicitly signed? 
                          then inc( buffPos );
                        while buffPos[0] in csDigits do
                          inc( buffPos );  
                      end;
                  s := StrLPas( p );
                  if gotPoint or gotExp  // did we get a floating literal?
                    then
                      begin
                        tokenKind  := tkFloatLit;
                        tokenFloat := StrToFloat( s );
                      end
                    else
                      begin
                        tokenKind := tkIntLit;
                        tokenInt  := StrToInt64( s );
                        Include( tokenFlags, tfDec );  // decimal integer literal
                      end;
                end;

            // If an integer literal is suffixed with the long type char "&"
            // flag the token.
            if ( buffPos[0] = '&' ) and ( tokenKind = tkIntLit )
              then
                begin
                  Include( tokenFlags, tfLongType );
                  inc( buffPos );
                end
              else
            // If a decimal integer or a floating literal is suffixed
            // with a floating type char, flag the token accordingly.
            if ( tfDec in tokenFlags ) or ( tokenKind = tkFloatLit )
              then
                case buffPos[0] of
                  '!' : CoerceToFloat( tfSingleType );
                  '#' : CoerceToFloat( tfDoubleType );
                  '@' : CoerceToFloat( tfCurrencyType );
                end;
          end;
      end;

    procedure LexIdent;
      var
        k : PvbToken;
      begin
        with fBuffer, token do
          begin
            ReadIdent;

            // Check if this is a keyword.
            k := vbKeywordHashTable[tokenText];
            if k <> nil
              then
                begin
                  tokenKind  := k^.tokenKind;
                  tokenFlags := tokenFlags + k^.tokenFlags;
                end;
                
            case buffPos[0] of
              '%' :
                begin
                  Include( tokenFlags, tfIntegerType );
                  inc( buffPos );
                end;
              '&' :
                begin
                  Include( tokenFlags, tfLongType );
                  inc( buffPos );
                end;
              '@' :
                begin
                  Include( tokenFlags, tfCurrencyType );
                  inc( buffPos );
                end;
              '#' :
                begin
                  Include( tokenFlags, tfDoubleType );
                  inc( buffPos );
                end;
              '$' :
                begin
                  Include( tokenFlags, tfStringType );
                  inc( buffPos );
                end;
              '!' :
                if ( not ( buffPos[1] in csIdentStart ) ) and ( buffPos[1] <> '[' )  // not bang operator?
                  then
                    begin
                      Include( tokenFlags, tfSingleType );
                      inc( buffPos );
                    end;
            end;
          end;
      end;

    procedure LexComment;
      begin
        with fBuffer, token do
          begin
            tokenKind := tkComment;
            p := buffPos;
            while not ( buffPos[0] in [#13, #10] ) do
              inc( buffPos );
            tokenText := StrLPas( p );
          end;
      end;

    begin
      assert( fBuffer <> nil );
      if fNextToken.tokenKind <> _tkUnknown
        then
          begin
            result := fNextToken;
            fNextToken.tokenKind := _tkUnknown;
          end
        else
          with fBuffer, token do
            begin
              tokenKind  := _tkUnknown;
              tokenText  := '';
              tokenFlags := [];
              if fBOL
                then
                  begin
                    Include( tokenFlags, tfBOL );
                    buffLineStart := buffPos;
                    inc( buffLineNumber );
                  end;
              if buffPos < buffEnd
                then
                  begin
                    repeat
                      if buffPos[0] in csWhitespaces
                        then
                          begin
                            Include( tokenFlags, tfPrevSpace );
                            repeat
                              inc( buffPos );
                            until not ( buffPos[0] in csWhitespaces );
                          end;

                      if buffPos[0] = '_'
                        then
                          begin
                            p := buffPos;
                            repeat
                              inc( p );
                            until not ( p[0] in csWhitespaces );
                            if ( p[0] = #13 ) or ( p[0] = #10)
                              then
                                begin
                                  if ( p[0] = #13 ) and ( p[1] = #10 )
                                    then inc( p );
                                  buffPos := p + 1;
                                end
                              else break;
                          end
                        else break;
                    until false;
                    
                    c := buffPos[0];
                    inc( buffPos );
                    case c of
                      'a'..'z', 'A'..'Z', '_' :
                        begin
                          LexIdent;
                          if tokenKind = tkRem
                            then LexComment;
                        end;
                      '0'..'9' : LexNumber;
                      '#' : LexDateOrHash;
                      '[' :
                        if buffPos[0] in csIdentStart
                          then LexEscapedIdent;
                      '"' : LexString;
                      '(' : tokenKind := tkOParen;
                      ')' : tokenKind := tkCParen;
                      '!' : tokenKind := tkBang;
                      ',' : tokenKind := tkComma;
                      ';' : tokenKind := tkSemiColon;
                      '+' : tokenKind := tkPlus;
                      '-' : tokenKind := tkMinus;
                      '*' : tokenKind := tkMult;
                      '^' : tokenKind := tkPower;
                      '/' : tokenKind := tkFloatDiv;
                      '\' : tokenKind := tkIntDiv;
                      '=' : tokenKind := tkEq;
                      ':' :
                        if buffPos[0] = '='
                          then
                            begin
                              tokenKind := tkColonEq;
                              inc( buffPos );
                            end
                          else tokenKind := tkColon;
                      '>' :
                        if buffPos[0] = '='
                          then
                            begin
                              tokenKind := tkGreaterEq;
                              inc( buffPos );
                            end
                          else tokenKind := tkGreater;
                      '.' :
                        begin
                          tokenKind := tkDot;
                          if buffPos[0] in csDigits
                            then LexNumber;
                        end;
                      '&' :
                        begin
                          tokenKind := tkConcat;
                          if buffPos[0] in ['h', 'H', 'o', 'O']
                            then LexNumber;
                        end;
                      '<' :
                        begin
                          tokenKind := tkLess;
                          case buffPos[0] of
                            '>' :
                              begin
                                token.tokenKind := tkNotEq;
                                inc( buffPos );
                              end;
                            '=' :
                              begin
                                token.tokenKind := tkLessEq;
                                inc( buffPos );
                              end;
                          end;
                        end;
                      '''' : LexComment;
                      #13,
                      #10 :
                        begin
                          inc( fLineNum );
                          if ( c = #13 ) and ( buffPos[0] = #10 )
                            then inc( buffPos );
                          fCurrLine := buffPos;
                          tokenKind := tkEOL;
                        end;
                    end;
                    fBOL := tokenKind = tkEOL;
                  end
                else tokenKind := _tkEOF;
            end;
    end;

  procedure TvbLexer.Fail( const msg : string );
    begin
      with fBuffer do
        raise EvbError.Create( msg, '<filaname>', buffLineNumber, buffColumnNumber );
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

  procedure CreateKeywordHashTable;
    var
      i : TvbTokenKind;
      p : PvbToken;
    begin
      Traits := nil;
      vbKeywordHashTable := nil;
      try
        Traits := TCaseInsensitiveTraits.Create;
        vbKeywordHashTable := THashList.Create( Traits, length( vbKeywords ) ) ;
        for i := low( vbKeywords ) to high( vbKeywords ) do
          begin
            p := @vbKeywords[i];
            vbKeywordHashTable.Add( vbKeywords[i].tokenText, p );
          end;
      except
        vbKeywordHashTable.Free;
        Traits.Free;
        raise;
      end;
    end;

  procedure DestroyKeywordHashTable;
    begin
      FreeAndNil( Traits );
      FreeAndNil( vbKeywordHashTable );
    end;

  procedure TvbLexer.GetPosition( var line, column : cardinal );
    begin
      line := fLineNum;
      column := cardinal(fBuffer.buffPos) - cardinal(fCurrLine);
    end;

initialization
  DecimalSeparator := '.';
  CreateKeywordHashTable;

finalization
  DestroyKeywordHashTable;

end.










