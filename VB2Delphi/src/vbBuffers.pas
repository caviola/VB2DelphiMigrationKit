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

unit vbBuffers;

interface

  const
    vbMaxLine  = 1023;
    vbBuffSize = 1024 * 4;
    vbMaxIdent = 255;

  type

    // A text buffer that always points to the end-of-buffer mark(#13) so
    // it's always empty.
    TvbCustomTextBuffer =
      class
        private
          fCR : char;
        public
          buffPos      : pchar; // the current position
          buffLimit    : pchar;
          buffLine     : pchar;
          buffLineNo   : cardinal;
          buffBOL      : boolean;
          buffFormDef  : boolean;
          buffSkipping : boolean;
          constructor Create;
          procedure Refill; virtual;
      end;

    // A buffer whose text is backed by a string.
    TvbStringBuffer =
      class( TvbCustomTextBuffer )
        private
          fStr : string;
        public
          constructor Create( const str : string );
          destructor Destroy; override;
      end;

    // Base class for buffers that read sequentially from the underlying source
    // on a chunk-by-chunk basis.
    TvbChunkBuffer =
      class( TvbCustomTextBuffer )
        private
          buff     : pchar;
          buffBase : pchar;
        protected
          procedure ReadNextChunk( var buff; count : cardinal; var bytesRead : cardinal ); virtual;
        public
          constructor Create;
          destructor Destroy; override;
          procedure Refill; override;
      end;

    // A chunk buffer backed by a file.
    TvbFileBuffer =
      class( TvbChunkBuffer )
        private
          fPath : string;
          fHandle   : cardinal;
          procedure SetFile( const path : string );
        protected
          procedure ReadNextChunk( var buff; count : cardinal; var bytesRead : cardinal ); override;
        public
          constructor Create;
          destructor Destroy; override;
          property buffFile : string read fPath write SetFile;
      end;

implementation

  uses
    SysUtils,
    Windows;

  { TvbFileBuffer }

  constructor TvbFileBuffer.Create;
    begin
      inherited;
      fHandle := INVALID_HANDLE_VALUE;
    end;

  destructor TvbFileBuffer.Destroy;
    begin
      if fHandle <> INVALID_HANDLE_VALUE
        then CloseHandle( fHandle );
      inherited;
    end;

  procedure TvbFileBuffer.ReadNextChunk( var buff; count : cardinal; var bytesRead : cardinal );
    begin
      assert( fHandle <> INVALID_HANDLE_VALUE );
      if ReadFile( fHandle, buff, count, bytesRead, nil ) = false
        then raise Exception.CreateFmt( 'couldn''t read from "%s"', [fPath] );
    end;

  procedure TvbFileBuffer.SetFile( const path : string );
    begin
      buffPos     := buffBase;
      buffLimit   := buffBase;
      buffLine    := buffBase;
      buffLineNo  := 1;
      buffBOL     := true;
      buffFormDef := false;
      fPath := path;
      if fHandle <> INVALID_HANDLE_VALUE
        then CloseHandle( fHandle );
      fHandle := CreateFile(
        pchar( path ),
        GENERIC_READ,
        FILE_SHARE_READ,
        nil,
        OPEN_EXISTING,
        FILE_FLAG_SEQUENTIAL_SCAN,
        0 );
      if fHandle <> INVALID_HANDLE_VALUE
        then Refill
        else raise Exception.CreateFmt( 'couldn''t open "%s"', [path] );
    end;

  { TvbChunkBuffer }

  constructor TvbChunkBuffer.Create;
    begin
      inherited;
      GetMem( buff, vbMaxLine + 1 + vbBuffSize + 1 );
      buffBase   := buff + vbMaxLine + 1;
    end;

  destructor TvbChunkBuffer.Destroy;
    begin
      if buff <> nil
        then FreeMem( buff );
      inherited;
    end;

  procedure TvbChunkBuffer.ReadNextChunk( var buff; count : cardinal; var bytesRead : cardinal );
    begin
      bytesRead := 0;
    end;

  procedure TvbChunkBuffer.Refill;
    var
      bytesRead, n : cardinal;
      dest : pchar;
    begin
      if buffPos >= buffLimit
        then buffPos := buffBase
        else
          begin
            n    := buffLimit - buffPos;
            dest := buffBase - n;
            buffLine := dest - ( buffPos - buffLine );
            while buffPos < buffLimit do
              begin
                dest[0] := buffPos[0];
                inc( dest );
                inc( buffPos );
              end;
            buffPos := buffBase - n;
          end;
      ReadNextChunk( buffBase[0], vbBuffSize, bytesRead );
      buffLimit    := buffBase + bytesRead;
      buffLimit[0] := #13;
    end;

  { TvbStringBuffer }

  constructor TvbStringBuffer.Create( const str : string );
    begin
      inherited Create;
      fStr := str;
    end;

  destructor TvbStringBuffer.Destroy;
    begin
      inherited;
    end;

  { TvbCustomTextBuffer }

  constructor TvbCustomTextBuffer.Create;
    begin
      buffBOL      := true;
      fCR          := #13;
      buffPos      := @fCR;
      buffLimit    := buffPos;
      buffLine     := buffPos;
      buffLineNo   := 0;
      buffSkipping := false;
      buffFormDef  := false;
    end;

  procedure TvbCustomTextBuffer.Refill;
    begin
    end;

end.
