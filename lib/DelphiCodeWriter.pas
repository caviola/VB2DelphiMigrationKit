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

unit DelphiCodeWriter;

interface

  uses
    TextWriters;

  type

    TDelphiCommentStyle  = ( csBracket, csParenAsterisk );

    TDelphiCodeWriter =
      class( TIndentedTextWriter )
        public
          procedure WriteLn( const Value : string; semicolon : boolean = false ); overload;
          procedure WriteLn( semicolon : boolean ); overload;
          procedure WriteBlockCmt( const cmt : string; aStyle : TDelphiCommentStyle = csBracket );
          procedure WriteLineCmt( const cmt : string );
          procedure WriteStrConst( const str : string );
          procedure WriteBegin( const cmt : string = '' );
          procedure WriteEnd( semicolon : boolean; const cmt : string = '' );
      end;

implementation

uses SysUtils;

  { TDelphiCodeWriter }

  procedure TDelphiCodeWriter.WriteBlockCmt( const cmt : string; aStyle : TDelphiCommentStyle );
    begin
      case aStyle of
        csBracket       : Write( '{' + cmt+ '}' );
        csParenAsterisk : Write( '(*' + cmt + '*)' );
      end;
    end;

  procedure TDelphiCodeWriter.WriteStrConst( const str : string );
    begin
      Write( QuotedStr( str ) );
    end;

  procedure TDelphiCodeWriter.WriteBegin( const cmt : string );
    begin
      Write( 'begin' );
      if cmt <> ''
        then
          begin
            Write( #32 );
            WriteLineCmt( cmt );
          end
        else WriteLn;
      Indent;
    end;

  procedure TDelphiCodeWriter.WriteEnd( semicolon : boolean; const cmt : string );
    begin
      Outdent;
      Write( 'end' );
      if semicolon
        then Write( ';' );
      if cmt <> ''
        then
          begin
            Write( #32 );
            WriteLineCmt( cmt );
          end
        else WriteLn;
    end;

  procedure TDelphiCodeWriter.WriteLn( semicolon : boolean );
    begin
      if semicolon
        then WriteLn( ';' )
        else WriteLn;
    end;

  procedure TDelphiCodeWriter.WriteLn( const Value : string; semicolon : boolean );
    begin
      Write( Value );
      WriteLn( semicolon );
    end;

  procedure TDelphiCodeWriter.WriteLineCmt( const cmt : string );
    begin
      WriteLn( '//' + cmt );
    end;

end.



