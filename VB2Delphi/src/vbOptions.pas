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

unit vbOptions;

interface

  uses
    Classes;

  type

    TvbOptions =
      class
        private
        public
          // Units always used in the interface section.
          optIntfUsedUnits : TStringList;
          // Units always used in the implementation section.
          optImplUsedUnits : TStringList;

          optUseAnsiString               : boolean;
          optUseByteBool                 : boolean;
          optString1ToChar               : boolean;
          optStringNToArrayOfChar        : boolean;
          optEnumToConst                 : boolean;
          optPackedRecord                : boolean;
          optCommentsBelowFuncToTop      : boolean;
          optParamDefaultByVal           : boolean;

          constructor Create;
          procedure SetDefaults;
          procedure Load( const iniFile : string );
          procedure Save( const iniFile : string );
      end;

implementation

  uses

    SysUtils,
    IniFiles;

  const
    // Default options values
    defStrToArray = false;
    defStr1ToChar = false;
    defPackedRec  = false;
    defEnumToInt  = true;

    secGlobal = 'global';

    // Name of options keys in INI file.
    // NOTE: keep these syncrhonized with the INI file.
    keyStrType    = 'use_ansistring';
    keyBoolType   = 'use_bytebool';
    keyStrToArray = 'static_string_to_array_of_char';
    keyPackedRec  = 'packed_record';
    keyStr1ToChar = 'string_1_to_char';
    keyEnumToInt  = 'enum_to_const';

  { TvbOptions }

  constructor TvbOptions.Create;
    begin
      SetDefaults;
    end;

  procedure TvbOptions.Load( const iniFile : string );
    begin
      with TIniFile.Create( iniFile ) do
        try
          optStringNToArrayOfChar := ReadBool( secGlobal, keyStrToArray, false );
          optString1ToChar        := ReadBool( secGlobal, keyStr1ToChar, false );
          optEnumToConst          := ReadBool( secGlobal, keyEnumToInt, false );
          optPackedRecord         := ReadBool( secGlobal, keyPackedRec, false );
        finally
          Free;
        end;
    end;

  procedure TvbOptions.Save( const iniFile : string );
    begin
      with TIniFile.Create( iniFile ) do
        try
        finally
          Free;
        end;
    end;

  procedure TvbOptions.SetDefaults;
    begin
      optUseAnsiString          := false;
      optString1ToChar          := false;
      optStringNToArrayOfChar   := false;
      optUseByteBool            := false;
      optEnumToConst            := false;
      optPackedRecord           := false;
      optCommentsBelowFuncToTop := false;
      optParamDefaultByVal      := false;
    end;

end.



