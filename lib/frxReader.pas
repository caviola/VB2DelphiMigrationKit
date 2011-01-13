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

// TfrxReader
// Read the content of VB FRX/CTX/DSX/DOX/PGX binary files.
// Delphi interface by Albert Almeida, Jan 2009 (caviola@gmail.com)
// Ported from VB code by Brad Martinez (http://www.mvps.org)

unit frxReader;

interface

  uses
    Classes;

  type
    TfrxItemType = (
      fitText,  // text data
      fitBLOB,  // binary data
      fitBMP,   // bitmap
      fitDIB,   // DIB
      fitGIF,   // GIF
      fitJPG,   // JPEG
      fitWMF,   // Windows Meta File
      fitEMF,   // Enhanced Meta File
      fitICO,   // icon
      fitCUR    // cursor
      );

  const
    TfrxItemExtension : array[TfrxItemType] of string = (
      'txt',
      '',
      'bmp',
      'dib',
      'gif',
      'jpg',
      'wmf',
      'emf',
      'ico',
      'cur' );

  type
    IfrxItem =
      interface
        ['{A2D34A01-6005-4F3A-AEE6-243EB8604C05}']
        procedure Save( stream : TStream );
        function GetType : TfrxItemType;
        function GetOffset : cardinal;
        function GetSize : cardinal;
        property ItemOffset : cardinal     read GetOffset;
        property ItemSize   : cardinal     read GetSize;
        property ItemType   : TfrxItemType read GetType;
      end;

    TfrxItemArray = array of IfrxItem;

    TfrxReader =
      class
        private
          hFile        : THandle;
          hFileMapping : THandle;
          fBuffer      : pchar;
          fFileSize    : cardinal;
          fItems       : TInterfaceList;
          function GetItem( i : cardinal ) : IfrxItem;
          function GetItemCount : cardinal;
          procedure FreeResources;
        public
          constructor Create;
          destructor Destroy; override;
          // Set the FRX file to read items from.
          // Returns TRUE on success, FALSE if file can't be read from.
          function SetFile( const filePath : string ) : boolean;
          procedure ReadItems;
          property Item[i : cardinal] : IfrxItem read GetItem;
          property ItemCount          : cardinal read GetItemCount;
      end;

implementation

  uses
    Windows,
    SysUtils;

  const
    // Copyright © 1997-1999 Brad Martinez, http://www.mvps.org
    //
    // ============================================================================
    // VB (and COM) recognize the following graphic files: BMP, DIB, GIF, JPG, WMF, EMF, ICO, CUR
    // each containing the following propietary image file format signatures:
    // (all IMGSIG_* and IMGTERM_* constants are user-defined)

    // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    // BMP, DIB (bitmap):
    IMGSIG_BMPDIB : WORD = $4D42;   // "BM" (//424D//) WORD @ image offset 0

    // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    // GIF:
    // First 3 bytes is "GIF", next 3 bytes is version, //87a//, //89a//, etc.
    IMGSIG_GIF  : DWORD = $464947;  // "GIF" (//4749 | 46//) masked DWORD @ image offset 0
    IMGTERM_GIF : byte = $3B;       // ";" (semicolon), WORD @ offset Len(image) - 1

    // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    // JPG:
    // SOI = Start Of Image = //FFD8//
    //   This marker must be present in any JPG file *once* at the beginning of the file.
    //    (Any JPG file starts with the sequence FFD8.)
    // EOI = End Of Image = //FFD9//
    //    Similar to EOI: any JPG file ends with FFD9.
    // APP0 = it//s the marker used to identify a JPG file which uses the JFIF specification = FFE0
    IMGSIG_JPG  : DWORD = $D8FF;  // (//FFD8//) WORD @ offset image 0, may have APP0
    IMGTERM_JPG : DWORD = $D9FF;  // (//FFD9//) WORD @ offset Len(image) - 2

    IMGSIG_ICO : DWORD = $10000;  // see above, DWORD @ image offset 0
    IMGSIG_CUR : DWORD = $20000;  // see above, DWORD @ image offset 0

    // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    // WMF, EMF:
    // first try to read the DWORD enhanced metafile signature (ENHMETAHEADER.dSignature member).
    // If that fails, try to read the DWORD METAHEADER.mtSize member
    // (it should equal FRXITEMHDR*.dwSizeImage), and check mtHeaderSize too.
    // If that fails, read the 16bit Aldus Placeable metafile header key:

    // "Q129658 SAMPLE: Reading and Wfiting Aldus Placeable Metafiles" or
    // "Q66949: INFO: Windows Metafile Functions & Aldus Placeable Metafiles"
    //typedef struct {
    //    DWORD           dwKey;   // 0x9AC6CDD7
    //    WORD              hmf;
    //    SMALL_RECT  bbox;
    //    WORD              wInch;
    //    DWORD           dwReserved;
    //    WORD              wCheckSum;
    // APMHEADER, *PAPMHEADER;  // APMFILEHEADER
    IMGSIG_WMF_APM : DWORD = $9AC6CDD7;

    FIH_Key = $746C;

  type

    PNEWHEADER = ^NEWHEADER;
    NEWHEADER =
      packed record        // was ICONDIR (ICONHEADER?)
        Reserved : WORD;   // must be 0
        ResType  : WORD;   // RES_ICON or RES_CURSOR
        ResCount : WORD;   // number of images (ICONDIRENTRYs) in the file (group)
      end;

    // Public Type ICONDIRENTRY
    //   bWidth As Byte          ' Width, in pixels, of the image
    //   bHeight As Byte         ' Height, in pixels, of the image
    //   bColorCount As Byte     ' Number of colors in image (0 if >=8bpp)
    //   bReserved As Byte       ' Reserved ( must be 0)
    //   wPlanes As Integer      ' Color Planes
    //   wBitCount As Integer    ' Bits per pixel
    //   dwBytesInRes As Long    ' How many bytes in this resource?
    //   dwImageOffset As Long   ' Where in the file is this image?
    // End Type

    FRX_HEADER_TYPE = (
      fstUnknown   = 0,
      fstItemHdr   = 1,   // FRX_ITEM_HDR
      fstItemHdrEx = 2,   // FRX_ITEM_HDR_EX
      fstItemHdrDW = 3,   // FRX_ITEM_HDR_DW
      fstItemHdrW  = 4    // FRX_ITEM_HDR_W
      );

    // ========================================================================
    // VB FRX/CTX/DSX/DOX/PGX binary file item header formats:
    // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    // TexBox.Text when Multiline = True has WORD text size value
    PFRX_ITEM_HDR_W = ^FRX_ITEM_HDR_W;
    FRX_ITEM_HDR_W =  // fihw
      packed record
        wDataSize : WORD;  // size of text
      end;

    // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    // Label.Caption and VB3 frx has DWORD image/text size value
    PFRX_ITEM_HDR_DW = ^FRX_ITEM_HDR_DW;
    FRX_ITEM_HDR_DW =  // fihdw
      packed record
        dwDataSize : DWORD; // size of image/text
      end;

    // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    // VB intrinsic control StdPictures (other blobs?) use FRXITEMHDR
    PFRX_ITEM_HDR = ^FRX_ITEM_HDR;
    FRX_ITEM_HDR =  // fih
      packed record
        dwItemSize : DWORD; // = dwDataSize + 8
        dwKey      : DWORD; // &H746C "lt" ( | 6C74 | )
        dwDataSize : DWORD; // size of image (= dwItemSize - 8)
      end;

    // frx binary when Form.Icon is deleted in designtime:
    //   0800 0000 6C74 0000 0000 0000   ....lt......  (just the FRXITEMHDR, no data)
    // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    // Comctl32.ocx, Mscomctl.ocx StdPictures (other blobs?) use FRXITEMHDREX
    PFRX_ITEM_HDR_EX = ^FRX_ITEM_HDR_EX;
    FRX_ITEM_HDR_EX =  // fihex, 28 bytes
      packed record
        dwItemSize  : DWORD;   // dwDataSize + 24
        clsid       : TGUID;   // CLSID_StdPicture, CLSID_?
        dwKey       : DWORD;   // &H746C "lt" ( | 6C74 | )
        dwDataSize  : DWORD;   // size of image (= dwItemSize - 24)
      end;

    FRX_ITEM_HDR_INFO =
      packed record
        dwHeaderType : FRX_HEADER_TYPE;
        dwSizeHdr    : DWORD;
        dwDataSize   : DWORD;
      end;

    TfrxItem =
      class( TInterfacedObject, IfrxItem )
        private
          fOffset : pointer;
          fSize   : DWORD;
          fType   : TfrxItemType;
        public
          constructor Create( itemType : TfrxItemType; buffer : pointer; size : DWORD );
          // IfrxItem
          procedure Save( stream : TStream );
          function GetType : TfrxItemType;
          function GetOffset : cardinal;
          function GetSize : cardinal;
      end;

  { TfrxReader }

  constructor TfrxReader.Create;
    begin
      fItems := TInterfaceList.Create;
      hFile := INVALID_HANDLE_VALUE;
    end;

  destructor TfrxReader.Destroy;
    begin
      FreeAndNil( fItems );
      FreeResources;
      inherited;
    end;

  procedure TfrxReader.FreeResources;
    begin
      if fBuffer <> nil
        then
          begin
            UnmapViewOfFile( fBuffer );
            fBuffer := nil;
          end;
      if hFileMapping <> 0
        then
          begin
            CloseHandle( hFileMapping );
            hFileMapping := 0;
          end;
      if hFile <> INVALID_HANDLE_VALUE
        then
          begin
            CloseHandle( hFile );
            hFile := INVALID_HANDLE_VALUE;
          end;
    end;

  function TfrxReader.GetItem( i : DWORD ) : IfrxItem;
    begin
      result := fItems[i] as IfrxItem;
    end;

  function TfrxReader.GetItemCount : DWORD;
    begin
      result := fItems.Count;
    end;

  procedure TfrxReader.ReadItems;
    var
      buffPos   : pchar;
      buffLimit : pchar;

    function ReadItemHeaderInfo : FRX_ITEM_HDR_INFO;
      begin
        with result do
          begin
            dwHeaderType := fstUnknown;
            // FRX_ITEM_HDR
            if ( buffPos + sizeof( FRX_ITEM_HDR ) <= buffLimit ) and
               ( PFRX_ITEM_HDR( buffPos ).dwItemSize - 8 = PFRX_ITEM_HDR( buffPos ).dwDataSize ) and
               ( PFRX_ITEM_HDR( buffPos ).dwKey = FIH_Key )
              then
                begin
                  dwHeaderType := fstItemHdr;
                  dwSizeHdr    := sizeof( FRX_ITEM_HDR );
                  dwDataSize   := PFRX_ITEM_HDR( buffPos ).dwDataSize;
                end
              else
            // FRX_ITEM_HDR_EX
            if ( buffPos + sizeof( FRX_ITEM_HDR_EX ) <= buffLimit ) and
               ( PFRX_ITEM_HDR_EX( buffPos ).dwItemSize - 24 = PFRX_ITEM_HDR_EX( buffPos ).dwDataSize ) and
               ( PFRX_ITEM_HDR_EX( buffPos ).dwKey = FIH_Key )
              then
                begin
                  dwHeaderType := fstItemHdrEx;
                  dwSizeHdr    := sizeof( FRX_ITEM_HDR_EX );
                  dwDataSize   := PFRX_ITEM_HDR_EX( buffPos ).dwDataSize;
                end
              else
            // FRX_ITEM_HDR_DW
            if ( buffPos + sizeof( FRX_ITEM_HDR_DW ) <= buffLimit ) and
               ( buffPos + PFRX_ITEM_HDR_DW( buffPos ).dwDataSize <= buffLimit )
              then
                begin
                  dwHeaderType := fstItemHdrDW;
                  dwSizeHdr    := sizeof( FRX_ITEM_HDR_DW );
                  dwDataSize   := PFRX_ITEM_HDR_DW( buffPos ).dwDataSize;
                end;
//              else
//            // FRX_ITEM_HDR_W
//            if ( buffPos + sizeof( FRX_ITEM_HDR_W ) <= buffLimit ) and
//               ( buffPos + PFRX_ITEM_HDR_W( buffPos ).wDataSize <= buffLimit )
//              then
//                begin
//                  dwHeaderType := fstItemHdrW;
//                  dwSizeHdr    := sizeof( FRX_ITEM_HDR_W );
//                  dwDataSize   := PFRX_ITEM_HDR_W( buffPos ).wDataSize;
//                end;
          end;
      end;

    function CheckImageFormat( size : DWORD ) : TfrxItemType;
      var
        dwImgSig  : DWORD;
        bihOffset : DWORD;
      begin
        result := fitBLOB;
        if size < 4
          then exit;
        dwImgSig := PDWORD( buffPos )^;
        // BMP, DIB: check image signature, and that BITMAPFILEHEADER.bfSize matches size
        if ( LOWORD( dwImgSig ) = IMGSIG_BMPDIB ) and ( PBITMAPFILEHEADER( buffPos ).bfSize = size )
          then result := fitBMP
          else
        // GIF: check image signature and terminator
        if ( ( dwImgSig and $00FFFFFF ) = IMGSIG_GIF ) and ( PBYTE( buffPos + size - 1 )^ = IMGTERM_GIF )
          then result := fitGIF
          else
        // JPG: check image signature and terminator
        if ( LOWORD( dwImgSig ) = IMGSIG_JPG ) and ( PWORD( buffPos + size - 2 )^ = IMGTERM_JPG )
          then result := fitJPG
          else
        // ICO, CUR: check image signature, then see inline below
        if ( dwImgSig = IMGSIG_ICO ) or ( dwImgSig = IMGSIG_CUR )
          then
            begin
              // Calculate the position of the BITMAPINFOHEADER struct
              bihOffset := sizeof( NEWHEADER ) + ( PNEWHEADER( buffPos ).ResCount * 16 );
              // Get the size of the BITMAPINFOHEADER struct from it's first biSize member
              if ( PBITMAPINFOHEADER( buffPos + bihOffset ).biSize = sizeof( BITMAPINFOHEADER ) )
                then
                  if dwImgSig = IMGSIG_ICO
                    then result := fitICO
                    else result := fitCUR; // (dwImgSig = IMGSIG_CUR)
            end
          else
        // WMF, Aldus Placeable metafile: check image signature
        if dwImgSig = IMGSIG_WMF_APM
          then result := fitWMF
          else
        // WMF, EMF
        // EMF: check image signature
        if size > sizeof( ENHMETAHEADER )
          then
            begin
              // first try to read the DWORD enhanced metafile signature
              if PENHMETAHEADER( buffPos ).dSignature = ENHMETA_SIGNATURE
                then result := fitEMF
                else
                  begin
                    // WMF, no Aldus: check METAHEADER.mtHeaderSize and mtSize members
                    if ( PMETAHEADER( buffPos ).mtHeaderSize = sizeof( METAHEADER ) ) and
                       ( PMETAHEADER( buffPos ).mtSize = size )
                      then result := fitWMF
                  end;
            end;
      end;

    var
      fit : TfrxItemType;
      len : byte;
    begin
      buffPos   := fBuffer;
      buffLimit := buffPos + fFileSize;
      while buffPos < buffLimit do
        with ReadItemHeaderInfo do
          if dwHeaderType = fstUnknown
            then
              begin
                len := byte( buffPos[0] );
                // Check for [len] text....#0A#0D
                if ( buffPos + 1 + len <= buffLimit ) and ( PWORD( buffPos + len - 1)^ = $0A0D )
                  then
                    begin
                      fItems.Add( TfrxItem.Create( fitText, buffPos + 1, len ) );
                      inc( buffPos, len );
                    end
                  else inc( buffPos );
              end
            else
              begin
                inc( buffPos, dwSizeHdr );
                if ( dwDataSize > 0 ) and ( buffPos + dwDataSize <= buffLimit )
                  then
                    begin
                      fit := CheckImageFormat( dwDataSize );
                      // Check for a BLOB with size >= 2 and last two bytes
                      // are $0D and $0A.
                      if ( fit = fitBLOB ) and
                         ( dwDataSize >= 2 ) and
                         ( PWORD( buffPos + dwDataSize - 2)^ = $0A0D ) and
                         ( dwHeaderType = fstItemHdrDW )
                        then fit := fitText;
                      fItems.Add( TfrxItem.Create( fit, buffPos, dwDataSize ) );
                      inc( buffPos, dwDataSize );
                    end;
              end;
    end;

  function TfrxReader.SetFile( const filePath : string ) : boolean;
    begin
      result := false;
      FreeResources;
      hFile := CreateFile(
        pchar( filePath ),
        GENERIC_READ,
        FILE_SHARE_READ,
        nil,
        OPEN_EXISTING,
        FILE_FLAG_SEQUENTIAL_SCAN,
        0 );        
      if hFile <> INVALID_HANDLE_VALUE
        then
          begin
            fFileSize := GetFileSize( hFile, nil );
            hFileMapping := CreateFileMapping( hFile, nil, PAGE_READONLY, 0, 0, nil );
            fBuffer := MapViewOfFile( hFileMapping, FILE_MAP_READ, 0, 0, 0 );
            if fBuffer <> nil
              then result := true;
          end;
    end;

  { TfrxItem }

  constructor TfrxItem.Create( itemType : TfrxItemType; buffer : pointer; size : DWORD );
    begin
      fType   := itemType;
      fOffset := buffer;
      fSize   := size;
    end;

  function TfrxItem.GetOffset : cardinal;
    begin
      result := 0;
    end;

  function TfrxItem.GetSize : cardinal;
    begin
      result := fSize;
    end;

  function TfrxItem.GetType : TfrxItemType;
    begin
      result := fType;
    end;

  procedure TfrxItem.Save( stream : TStream );
    begin
      assert( stream <> nil );
      stream.WriteBuffer( fOffset^, fSize );
    end;

end.
