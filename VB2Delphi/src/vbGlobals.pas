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

unit vbGlobals;

interface

  uses
    vbNodes,
    Classes;

  function GetOpenStmtFileName( fileNum : TvbExpr ) : string;

implementation

  uses
    SysUtils;

  function GetOpenStmtFileName( fileNum : TvbExpr ) : string;
    var
      intExpr  : TvbIntLit absolute fileNum;
      nameExpr : TvbNameExpr absolute fileNum;
    begin
      if fileNum.NodeKind = INT_LIT_EXPR
        then result := 'file' + IntToStr( intExpr.Value )
        else
      if fileNum.NodeKind = NAME_EXPR
        then result := nameExpr.Name
        else assert( true ); // if we reach here we are in trouble
    end;

end.











