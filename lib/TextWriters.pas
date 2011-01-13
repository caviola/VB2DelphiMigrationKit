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

unit TextWriters;

interface

  uses
    Classes;

  type

    TTextWriter =
      class
        private
          fStream : TStream;
          fOwnsStream : boolean;
          fBOL : boolean;
        protected
          function GetStream : TStream; virtual;
          procedure SetStream( aStream : TStream ); virtual;
          function GetBOL : boolean; virtual;
          procedure DoWrite( const aText : string ); overload;

          // Override to directly write to the underlying stream.
          // "WriteXXX" functions are implemented in terms of DoWrite.
          //
          procedure DoWrite( const Buffer; Size : cardinal ); overload; virtual;
        public
          constructor Create; overload;
          constructor Create( aFile : string ); overload;
          constructor Create( aStream : TStream; Owned : boolean = false ); overload;
          destructor Destroy; override;
          procedure Write( const Value : string ); overload; virtual;
          procedure Write( const Value : int64 ); overload; virtual;
          procedure Write( const Value : extended ); overload; virtual;
          procedure WriteLn( const Value : string ); overload; virtual;
          procedure WriteLn( const Value : int64 ); overload; virtual;
          procedure WriteLn( const Value : extended ); overload; virtual;
          procedure WriteLn( strs : TStringList ); overload; virtual;
          procedure WriteLn; overload; virtual;
          procedure WriteFmt( const Format : string; const Args : array of const ); virtual;
          property Stream : TStream read GetStream  write SetStream;
          property BOL : boolean read GetBOL;
      end;

    TIndentedTextWriter =
      class( TTextWriter )
        private
          fTabString : string;
          fIndenting : boolean;
          fIndentDepth : integer;
          fIndentString : string;
        protected
          procedure DoWrite( const Buffer; Size : cardinal ); overload; override;
          procedure SetStream( aStream : TStream ); override;
        public
          constructor Create( aFile : string ); overload;
          constructor Create( aStream : TStream; Owned : boolean = false ); overload;  
          procedure WriteLn; override;
          procedure Indent;
          procedure Outdent;
          property TabString : string read fTabString write fTabString;
          property Indenting : boolean read fIndenting write fIndenting;
          property IndentDepth : integer read fIndentDepth;
      end;

implementation

  uses
    SysUtils,
    StrUtils;

  { TTextWriter }

  constructor TTextWriter.Create;
    begin
      fBOL := true;
      fOwnsStream := false;
    end;

  constructor TTextWriter.Create( aStream : TStream; Owned : boolean = false );
    begin
      Create;
      assert( aStream <> nil );
      fStream     := aStream;
      fOwnsStream := Owned;
    end;

  constructor TTextWriter.Create( aFile : string );
    begin
      Create( TFileStream.Create( aFile, fmCreate	or fmShareDenyWrite	), true );
    end;

  procedure TTextWriter.Write( const Value : int64 );
    begin
      Write( IntToStr( Value ) );
    end;

  procedure TTextWriter.Write( const Value : string );
    begin
      DoWrite( Value );
    end;

  procedure TTextWriter.Write( const Value : extended );
    begin
      Write( FloatToStr( Value ) );
    end;

  procedure TTextWriter.DoWrite( const Buffer; Size : cardinal );
    begin
      fStream.WriteBuffer( Buffer, Size );
      fBOL := false;
    end;

  procedure TTextWriter.DoWrite( const aText : string );
    begin
      DoWrite( aText[1], Length( aText ) );
    end;

  procedure TTextWriter.WriteFmt( const Format : string; const Args : array of const );
    begin
      Write( SysUtils.Format( Format, Args ) );
    end;

  procedure TTextWriter.WriteLn;
    begin
      Write( sLineBreak );
      fBOL := true;
    end;

  destructor TTextWriter.Destroy;
    begin
      if fOwnsStream
        then fStream.Free;
      inherited;
    end;

  procedure TTextWriter.WriteLn( const Value : string );
    begin
      Write( Value );
      WriteLn;
    end;

  procedure TTextWriter.WriteLn( const Value : int64 );
    begin
      Write( Value );
      WriteLn;
    end;

  procedure TTextWriter.WriteLn( const Value : extended );
    begin
      Write( Value );
      WriteLn;
    end;

  function TTextWriter.GetBOL : boolean;
    begin
      result := fBOL;
    end;

  function TTextWriter.GetStream : TStream;
    begin
      result := fStream;
    end;

  procedure TTextWriter.SetStream( aStream : TStream );
    begin
      fStream := aStream;
    end;

  procedure TTextWriter.WriteLn(strs: TStringList);
    var
      i : integer;
    begin
      for i := 0 to strs.Count - 1 do
        WriteLn( strs[i] );
    end;

  { TIndentedTextWriter }

  constructor TIndentedTextWriter.Create( aStream : TStream; Owned : boolean );
    begin
      inherited;
      fIndentDepth := 0;
      fTabString   := #32#32;  // indent with two whitespaces
      fIndenting   := true;    // indenting is enabled
    end;

  constructor TIndentedTextWriter.Create( aFile : string );
    begin
      Create( TFileStream.Create( aFile, fmCreate	or fmShareDenyWrite	), true );
    end;

  procedure TIndentedTextWriter.DoWrite( const Buffer; Size : cardinal );
    begin
      if fBOL and fIndenting
        then inherited DoWrite( fIndentString[1], length( fIndentString ) );
      inherited;
    end;

  procedure TIndentedTextWriter.Indent;
    begin
      inc( fIndentDepth );
      fIndentString := DupeString( fTabString, fIndentDepth );
    end;

  procedure TIndentedTextWriter.Outdent;
    begin
      assert( fIndentDepth > 0 );
      dec( fIndentDepth );
      fIndentString := DupeString( fTabString, fIndentDepth );
    end;

  procedure TIndentedTextWriter.SetStream( aStream : TStream );
    begin
      inherited;
      fIndentDepth := 0;
      fTabString   := #32#32;  // indent with two whitespaces
      fIndenting   := true;    // indenting is enabled
    end;

  procedure TIndentedTextWriter.WriteLn;
    var
      prevIndent : boolean;
    begin
      // Indenting newlines is useless so we disable indentation here to
      // prevent unnecessary chars from being written to the stream.
      //
      prevIndent := fIndenting;
      fIndenting := false;
      inherited;
      fIndenting := prevIndent;
    end;

end.
