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

unit ArrayUtils;

interface

  uses
    Classes,
    Types;

  function StrDynArrayFromStrList( strList : TStringList ) : TStringDynArray;
  function StrDynArrayConcat( const arr1, arr2 : TStringDynArray ) : TStringDynArray;
  function StrListFromStrDynArray( const arr : TStringDynArray; CreateIfEmpty : boolean = true ) : TStringList;

implementation

  function StrDynArrayFromStrList( strList : TStringList ) : TStringDynArray;
    var
      i : integer;
    begin
      assert( strList <> nil );
      if strList <> nil
        then
          with strList do
            begin
              setlength( result, Count );
              for i := 0 to Count - 1 do
                result[i] := Strings[i];
            end
        else result := nil;
    end;

  function StrDynArrayConcat( const arr1, arr2 : TStringDynArray ) : TStringDynArray;
    var
      i : integer;
    begin
      setlength( result, length( arr1 ) + length( arr2 ) );
      for i := low( arr1 ) to high( arr1 ) do
        result[i] := arr1[i];
      for i := low( arr2 ) to high( arr2 ) do
        result[length(arr1) + i] := arr2[i];
    end;

  function StrListFromStrDynArray( const arr : TStringDynArray; CreateIfEmpty : boolean ) : TStringList;
    var
      i : integer;
    begin
      if not CreateIfEmpty and ( length( arr ) = 0 )
        then result := nil
        else
          begin
            result := TStringList.Create;
            for i := low( arr ) to high( arr ) do
              result.Add( arr[i] );  
          end;
    end;

end.
