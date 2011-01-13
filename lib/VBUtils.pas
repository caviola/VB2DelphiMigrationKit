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

unit VBUtils;

interface

  function VB_DemangleName( const name : string ) : string;
  function VB_GetIntrinsicTypeName( const vt : TVarType ) : string;

implementation

  uses
    Windows,
    Types,
    SysUtils;

  function VB_DemangleName( const name : string ) : string;
    const
      L = length( '_B_str_' ); // also the length of "_B_var_"
    var
      n : pchar absolute name;
    begin
      if ( n[0] = '_' ) and ( n[1] = 'B' ) and ( n[2] = '_' )
        then
          if ( n[3] = 's' ) and
             ( n[4] = 't' ) and
             ( n[5] = 'r' ) and
             ( n[6] = '_' )
            then result := copy( name, L + 1, length( name ) - L ) + '$'
            else
          if ( n[3] = 'v' ) and
             ( n[4] = 'a' ) and
             ( n[5] = 'r' ) and
             ( n[6] = '_' )
            then result := copy( name, L + 1, length( name ) - L )
            else result := name
        else result := name;
    end;

  function VB_GetIntrinsicTypeName( const vt : TVarType ) : string;
    const
      VBType : array [varEmpty..varInt64] of string = (
        '?',
        '?',
        'integer',
        'long',
        'single',
        'double',
        'currency',
        'date',
        'string',
        'object',
        '?',
        'boolean',
        'variant',
        '?',
        'decimal',
        '?',
        '?',
        'byte',
        '?',
        '?',
        '?');
    begin
      result := VBType[vt];
    end;

end.



