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

unit vbTokens;

interface

  type

    TvbTokenKind = (
      _tkUnknown,
      _tkEOF,
      tkEOL,

      tkDateLit,
      tkFloatLit,
      tkIntLit,
      tkStringLit,
      tkIdent,
      tkComment,

      // Concatenated block-termination tokens
      //
      _tkEndEnum,
      _tkEndFunction,
      _tkEndIf,
      _tkEndProperty,
      _tkEndSelect,
      _tkEndSub,
      _tkEndType,
      _tkEndWith,

      // Concatenated block-exit tokens
      //
      _tkExitDo,
      _tkExitFor,
      _tkExitFunction,
      _tkExitProperty,
      _tkExitSub,

      // Operators and separators
      //
      tkBang,
      tkColon,
      tkColonEq,
      tkComma,
      tkSemiColon,
      tkConcat,
      tkCParen,
      tkDot,
      tkEq,
      tkFloatDiv,
      tkGreater,
      tkGreaterEq,
      tkHash,
      tkIntDiv,
      tkLess,
      tkLessEq,
      tkMinus,
      tkMult,
      tkNotEq,
      tkOParen,
      tkPlus,
      tkPower,

      // Keywords
      //
      tkAbs,
      tkAddressOf,
      tkAnd,
      tkAny,
      tkArray,
      tkAs,
      tkAttribute,
      tkBoolean,
      tkByRef,
      tkByte,
      tkByVal,
      tkCall,
      tkCase,
      tkCBool,
      tkCByte,
      tkCDate,
      tkCDbl,
      tkCDec,
      tkCInt,
      tkCircle,
      tkCLng,
      tkClose,
      tkConst,
      tkCSng,
      tkCStr,
      tkCurrency,
      tkCVar,
      tkCVErr,
      tkDate,
      tkDebug,
      tkDecimal,
      tkDeclare,
      tkDefBool,
      tkDefByte,
      tkDefCur,
      tkDefDate,
      tkDefDbl,
      tkDefDec,
      tkDefInt,
      tkDefLng,
      tkDefObj,
      tkDefSng,
      tkDefStr,
      tkDefVar,
      tkDim,
      tkDo,
      tkDoEvents,
      tkDouble,
      tkEach,
      tkElse,
      tkElseIf,
      tkEnd,
      tkEnum,
      tkEqv,
      tkErase,
      tkEvent,
      tkExit,
      tkFalse,
      tkFix,
      tkFor,
      tkFriend,
      tkFunction,
      tkGet,
      tkGlobal,  // for backward compatibility with vb3
      tkGoSub,
      tkGoTo,
      tkIf,
      tkImp,
      tkImplements,
      tkIn,
      tkInput,
      tkInputB,
      tkInt,
      tkInteger,
      tkIs,
      tkLBound,
      tkLen,
      tkLenB,
      tkLet,
      tkLike,
      tkLocal,
      tkLock,
      tkLong,
      tkLoop,
      tkLSet,
      tkMe,
      tkMod,
      tkNew,
      tkNext,
      tkNot,
      tkNothing,
      tkObject,
      tkOn,
      tkOpen,
      tkOption,
      tkOptional,
      tkOr,
      tkParamArray,
      tkPreserve,
      tkPrint,
      tkPrivate,
      tkProperty,
      tkPSet,
      tkPublic,
      tkPut,
      tkRaiseEvent,
      tkReDim,
      tkRem,
      tkResume,
      tkReturn,
      tkRSet,
      tkSeek,
      tkSelect,
      tkSet,
      tkSgn,
      tkShared,
      tkSingle,
      tkSpc,
      tkStatic,
      tkStop,
      tkString,
      tkSub,
      tkTab,
      tkThen,
      tkTo,
      tkTrue,
      tkType,
      tkTypeOf,
      tkUBound,
      tkUnlock,
      tkUntil,
      tkVariant,
      tkWend,
      tkWhile,
      tkWith,
      tkWithEvents,
      tkWrite,
      tkXor
    );

    TvbTokenFlag = (
      tfEndBlock,
      tfExitBlock,
      tfBOL,
      tfCurrencyType,
      tfIntegerType,
      tfDoubleType,
      tfSingleType,
      tfStringType,
      tfLongType,
      tfBuiltInType,
      tfEscapedIdent,
      tfDec,
      tfHex,
      tfOctal,
      tfPrevSpace );

    TvbTokenFlags = set of TvbTokenFlag;
    TvbTokenKinds =  set of TvbTokenKind;

    PvbToken = ^TvbToken;
    TvbToken =
      record
        tokenKind  : TvbTokenKind;
        tokenText  : string;
        tokenFlags : TvbTokenFlags;
        case TvbTokenKInd of
          tkDateLit  : ( tokenDate  : TDateTime );
          tkFloatLit : ( tokenFloat : double );
          tkIntLit   : ( tokenInt   : integer );
      end;

  function GetTokenText( t : TvbTokenKind ) : string;

implementation

  const
    vbTokenTexts : array[TvbTokenKind] of string =
      (
        'UNKNOWN',
        'end-of-file',
        'end-of-line',
        'date/time constant',
        'floating constant',
        'integer constant',
        'string constant',
        'identifier',
        'comment',
        'End Enum',
        'End Function',
        'End If',
        'End Property',
        'End Select',
        'End Sub',
        'End Sub',
        'End With',
        'Exit Do',
        'Exit For',
        'Exit Function',
        'Exit Property',
        'Exit Sub',
        '!',
        ':',
        ':=',
        ',',
        ',',
        '&',
        ')',
        '.',
        '=',
        '/',
        '>',
        '>=',
        '#',
        'div',
        '<',
        '<=',
        '-',
        '*',
        '<>',
        '(',
        '+',
        '^',
        'Abs',
        'Address Of',
        'And',
        'Any',
        'Array',
        'As',
        'Attribute',
        'Boolean',
        'ByRef',
        'Byte',
        'ByVal',
        'Call',
        'Case',
        'CBool',
        'CByte',
        'CDate',
        'CDbl',
        'CDec',
        'CInt',
        'Circle',
        'CLng',
        'Close',
        'Const',
        'CSng',
        'CStr',
        'Currency',
        'CVar',
        'CVErr',
        'Date',
        'Debug',
        'Decimal',
        'Declare',
        'DefBool',
        'DefByte',
        'DefCur',
        'DefDate',
        'DefDbl',
        'DefDec',
        'DefInt',
        'DefLng',
        'DefObj',
        'DefSng',
        'DefStr',
        'DefVar',
        'Dim',
        'Do',
        'DoEvents',
        'Double',
        'Each',
        'Else',
        'Else If',
        'End',
        'Enum',
        'Eqv',
        'Erase',
        'Event',
        'Exit',
        'False',
        'Fix',
        'For',
        'Friend',
        'Function',
        'Get',
        'Global',
        'GoSub',
        'GoTo',
        'If',
        'Imp',
        'Implements',
        'In',
        'Input',
        'InputB',
        'Int',
        'Integer',
        'Is',
        'LBound',
        'Len',
        'LenB',
        'Let',
        'Like',
        'Loca',
        'Lock',
        'Long',
        'Loop',
        'LSet',
        'Me',
        'Mod',
        'New',
        'Next',
        'Not',
        'Nothing',
        'Object',
        'On',
        'Open',
        'Option',
        'Optional',
        'Or',
        'ParamArray',
        'Preserve',
        'Print',
        'Private',
        'Property',
        'PSet',
        'Public',
        'Put',
        'RaiseEvent',
        'ReDim',
        'Rem',
        'Resume',
        'Return',
        'RSet',
        'Seek',
        'Select',
        'Set',
        'Sgn',
        'Shared',
        'Single',
        'Spc',
        'Static',
        'Stop',
        'String',
        'Sub',
        'Tab',
        'Then',
        'To',
        'True',
        'Type',
        'TypeOf',
        'UBound',
        'Unlock',
        'Until',
        'Variant',
        'Wend',
        'While',
        'With',
        'WithEvents',
        'Write',
        'Xor'
      );

  function GetTokenText( t : TvbTokenKind ) : string;
    begin
      result := vbTokenTexts[t];
//        _tkEOF      : result := 'end-of-file';
//        tkEOL       : result := 'end-of-line';
//        tkDateLit   : result := 'date-literal';
//        tkFloatLit  : result := 'floating-literal';
//        tkIntLit    : result := 'integer-literal';
//        tkStringLit : result := 'string-literal';
//        tkIdent     : result := 'identifier';
//        _tkEndEnum  : result := 'End Enum';
//        _tkEndFunction  : result := 'End Function';
//        _tkEndIf        : result := 'End If';
//        _tkEndProperty  : result := 'End Property';
//        _tkEndSelect    : result := 'End Select';
//        _tkEndSub       : result := 'End Sub';
//        _tkEndType      : result := 'End Sub';
//        _tkEndWith : result := 'End With';
//        _tkExitDo : result := 'Exit Do';
//        _tkExitFor : result := 'Exit For';
//        _tkExitFunction : result := 'Exit Function';
//        _tkExitProperty : result := 'Exit Property';
//        _tkExitSub : result := 'Exit Sub';
//        tkBang : result := '!';
//        tkColon : result := ':';
//        tkColonEq : result := ':=';
//        tkComma : result := ',';
//        tkSemiColon : result := ';';
//        tkConcat : result := '&';
//        tkCParen : result := ')';
//        tkDot : result := '.';
//        tkEq : result := '=';
//        tkFloatDiv : result := '/';
//        tkGreater : result := '>';
//        tkGreaterEq : result := '>=';
//        tkHash : result := '#';
//        tkIntDiv : result := 'div';
//        tkLess : result := '<';
//        tkLessEq : result := '<=';
//        tkMinus : result := '-';
//        tkMult : result := '*';
//        tkNotEq : result := '<>';
//        tkOParen : result := '(';
//        tkPlus : result := '+';
//        tkPower : result := '^';
//        tkAbs : result := 'Abs';
//        tkAddressOf : result := 'Address Of';
//        tkAnd : result := 'And';
//        tkAny : result := 'Any';
//        tkArray : result := 'Array';
//        tkAs : result := 'As';
//        tkAttribute : result := 'Attribute';
//        tkBoolean : result := 'Boolean';
//        tkByRef : result := 'ByRef';
//        tkByte : result := 'Byte';
//        tkByVal : result := 'ByVal';
//        tkCall : result := 'Call';
//        tkCase : result := 'Case';
//        tkCBool : result := 'CBool';
//        tkCByte : result := 'CByte';
//        tkCDate : result := 'CDate';
//        tkCDbl : result := 'CDbl';
//        tkCDec : result := 'CDec';
//        tkCInt : result := 'CInt';
//        tkCircle : result := 'Circle';
//        tkCLng : result := 'CLng';
//        tkClose : result := 'Close';
//        tkConst : result := 'Const';
//        tkCShort : result := 'CShort';
//        tkCSng : result := 'CSng';
//        tkCStr : result := 'CStr';
//        tkCurrency : result := 'Currency';
//        tkCVar : result := 'CVar';
//        tkCVErr : result := 'CVErr';
//        tkDate : result := 'Date';
//        tkDebug : result := 'Debug';
//        tkDecimal : result := 'Decimal';
//        tkDeclare : result := 'Declare';
//        tkDefBool : result := 'DefBool';
//        tkDefByte : result := 'DefByte';
//        tkDefCur : result := 'DefCur';
//        tkDefDate : result := 'DefDate';
//        tkDefDbl : result := 'DefDbl';
//        tkDefDec : result := 'DefDec';
//        tkDefInt : result := 'DefInt';
//        tkDefLng : result := 'DefLng';
//        tkDefObj : result := 'DefObj';
//        tkDefSng : result := 'DefSng';
//        tkDefStr : result := 'DefStr';
//        tkDefVar : result := 'DefVar';
//        tkDim : result := 'Dim';
//        tkDo : result := 'Do';
//        tkDoEvents : result := 'DoEvents';
//        tkDouble : result := 'Double';
//        tkEach : result := 'Each';
//        tkElse : result := 'Else';
//        tkElseIf : result := 'Else If';
//        tkEnd : result := 'End';
//        tkEnum : result := 'Enum';
//        tkEqv : result := 'Eqv';
//        tkErase : result := 'Erase';
//        tkEvent : result := 'Event';
//        tkExit : result := 'Exit';
//        tkFalse : result := 'False';
//        tkFix : result := 'Fix';
//        tkFor : result := 'For';
//        tkFriend : result := 'Friend';
//        tkFunction : result := 'Function';
//        tkGet : result := 'Get';
//        tkGlobal : result := 'Global';
//        tkGoSub : result := 'GoSub';
//        tkGoTo : result := 'GoTo';
//        tkIf : result := 'If';
//        tkImp : result := 'Imp';
//        tkImplements : result := 'Implements';
//        tkIn : result := 'In';
//        tkInput : result := 'Input';
//        tkInputB : result := 'InputB';
//        tkInt : result := 'Int';
//        tkInteger : result := 'Integer';
//        tkIs : result := 'Is';
//        tkLBound : result := 'LBound';
//        tkLen : result := 'Len';
//        tkLenB : result := 'LenB';
//        tkLet : result := 'Let';
//        tkLike : result := 'Like';
//        tkLocal : result := 'Loca';
//        tkLock : result := 'Lock';
//        tkLong : result := 'Long';
//        tkLoop : result := 'Loop';
//        tkLSet : result := 'LSet';
//        tkMe : result := 'Me';
//        tkMod : result := 'Mod';
//        tkNew : result := 'New';
//        tkNext : result := 'Next';
//        tkNot : result := 'Not';
//        tkNothing : result := 'Nothing';
//        tkObject : result := 'Object';
//        tkOn : result := 'On';
//        tkOpen : result := 'Open';
//        tkOption : result := 'Option';
//        tkOptional : result := 'Optional';
//        tkOr : result := 'Or';
//        tkParamArray : result := 'ParamArray';
//        tkPreserve : result := 'Preserve';
//        tkPrint : result := 'Print';
//        tkPrivate : result := 'Private';
//        tkProperty : result := 'Property';
//        tkPSet : result := 'PSet';
//        tkPublic : result := 'Public';
//        tkPut : result := 'Put';
//        tkRaiseEvent : result := 'RaiseEvent';
//        tkReDim : result := 'ReDim';
//        tkRem : result := 'Rem';
//        tkResume : result := 'Resume';
//        tkReturn : result := 'Return';
//        tkRSet : result := 'RSet';
//        tkSeek : result := 'Seek';
//        tkSelect : result := 'Select';
//        tkSet : result := 'Set';
//        tkSgn : result := 'Sgn';
//        tkShared : result := 'Shared';
//        tkShort : result := 'Short';
//        tkSingle : result := 'Single';
//        tkSpc : result := 'Spc';
//        tkStatic : result := 'Static';
//        tkStop : result := 'Stop';
//        tkString : result := 'String';
//        tkSub : result := 'Sub';
//        tkTab : result := 'Tab';
//        tkThen : result := 'Then';
//        tkTo : result := 'To';
//        tkTrue : result := 'True';
//        tkType : result := 'Type';
//        tkTypeOf : result := 'TypeOf';
//        tkUBound : result := 'UBound';
//        tkUnlock : result := 'Unlock';
//        tkUntil : result := 'Until';
//        tkVariant : result := 'Variant';
//        tkWend : result := 'Wend';
//        tkWhen : result := 'When';
//        tkWhile : result := 'While';
//        tkWith : result := 'With';
//        tkWithEvents : result := 'WithEvents';
//        tkWrite : result := 'Write';
//        tkXor : result := 'Xor';
    end;

end.

