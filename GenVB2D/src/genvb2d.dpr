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

program GenVB2D;

{$APPTYPE CONSOLE}
{$R *.res}

uses
  FastMM4,
  SysUtils,
  JclFileUtils,
  Variants,
  DateUtils,
  ActiveX,
  TLI_TLB,
  CommandLine,
  TextWriters,
  VBUtils;

  const
    sInfo  = '[Info] ';
    sError = '[Error] ';

    msgTlbInf32NotFound = 
      'This program uses the COM server TLBINF32.DLL and it couldn''t be loaded.'      + #10#13 +
      'Please make sure the DLL exists in your system and it''s correctly registered.' + #10#13 +
      'You can register it with "REGSVR32 TLBINF32.DLL"';

  type

    TGenVB2D =
      class
        private
          fIgnoreEnum   : boolean;
          fIgnoreRecord : boolean;
          fIgnoreModule : boolean;
          fIgnoreClass  : boolean;
          fOverwrite    : boolean;
          fConstValues  : boolean;
          fWriteHelpStr : boolean;
          fInputFile    : string;
          fOutputDir    : string;
          fOutput       : TIndentedTextWriter;
          TLI           : TLIApplication;
          procedure ParseCommandLineArgument( var p : integer; const arg : string );
          procedure ParseCommandLineOption( var p : integer; const opt : string );
          function TryCreateTLI : boolean;
          function WriteAttrIf( const name, value : string; cond : boolean = true ) : boolean;
          procedure Convert( const TypeLibPath, OutputMapPath : string );
          procedure ParseCommandLine;
          procedure ShowUsage;
          procedure WriteAttr( const name, value : string );
          procedure WriteMember( const typ : TypeInfo; const mem : MemberInfo; invs : InvokeKinds = 0 );
          //procedure WriteParam( const param : ParameterInfo );
          procedure WriteTypeAttr( const vti : VarTypeInfo );
        public
          constructor Create;
          function Run : integer;
      end;

  // These type libraries are automatically referenced by all VB projects:
  //    VB     - C:\Program Files\Microsoft Visual Studio 6\VB6.OLB
  //    VBA    - C:\WINDOWS\system32\msvbvm60.dll
  //    VBRUN  - C:\WINDOWS\system32\msvbvm60.dll\3
  //    stdole - C:\WINDOWS\system32\stdole2.tlb

  { TGenVB2D }

  constructor TGenVB2D.Create;
    begin
      // NOTE: These fields are FALSE by default because Delphi "zeroes" the object
      //       before calling the constructor. but i do these here for clarity.
      fIgnoreEnum   := false;
      fIgnoreRecord := false;
      fIgnoreModule := false;
      fIgnoreClass  := false;
      fOverwrite    := false;
      fConstValues  := false;
      fWriteHelpStr := false;
    end;

  procedure TGenVB2D.ParseCommandLine;
    begin
      CommandLine.ParseCommandLine( ParseCommandLineOption, ParseCommandLineArgument );
      if fOutputDir = ''
        then fOutputDir := ExtractFilePath( fInputFile ) + 'vb2d\';
    end;

  procedure TGenVB2D.ParseCommandLineArgument( var p : integer; const arg : string );
    begin
      if fInputFile = ''
        then fInputFile := arg
        else
      if fOutputDir = ''
        then fOutputDir := IncludeTrailingPathDelimiter( arg );
    end;

  procedure TGenVB2D.ParseCommandLineOption( var p : integer; const opt : string );
    begin
      if opt = 'ie'
        then fIgnoreEnum := true
        else
      if opt = 'ir'
        then fIgnoreRecord := true
        else
      if opt = 'im'
        then fIgnoreModule := true
        else
      if opt = 'ic'
        then fIgnoreClass := true
        else
      if opt = 'ow'
        then fOverwrite := true
        else
      if opt = 'cv'
        then fConstValues := true
        else
      if opt = 'ih'
        then fWriteHelpStr := true;
    end;

  function TGenVB2D.Run : integer;
    var
      sr : TSearchRec;
      OutputMapPath : string;
    begin
      result := 0;
      if ParamCount > 0
        then
          begin
            if TryCreateTLI
              then
                try
                  ParseCommandLine;

                  if FindFirst( fInputFile, faAnyFile, sr ) <> 0
                    then raise Exception.CreateFmt( 'can''t read "%s"', [fInputFile] );

                  if not JclFileUtils.ForceDirectories( fOutputDir )
                    then raise Exception.CreateFmt( 'can''t create output directory ''%s''', [fOutputDir] );

                  repeat
                    OutputMapPath := fOutputDir + sr.Name + '.vb2d';
                    if not fOverwrite and FileExists( OutputMapPath )
                      then raise Exception.CreateFmt( 'translation map "%s" already exists', [OutputMapPath] );
                    Convert( ExtractFilePath( fInputFile ) + sr.Name, OutputMapPath );
                  until FindNext( sr ) <> 0;
                  FindClose( sr );
                except
                  on e : Exception do
                    begin
                      writeln( sError, e.Message );
                      result := 1;
                    end;
                end
              else
                begin
                  writeln;
                  writeln( msgTlbInf32NotFound );
                end;
          end
        else ShowUsage;
    end;

  function TGenVB2D.TryCreateTLI : boolean;
    begin
      result := true;
      try
        TLI := CoTLIApplication.Create;
      except
        result := false;
      end;
    end;

  procedure TGenVB2D.WriteTypeAttr( const vti : VarTypeInfo );
    var
      //isArray : boolean;
      vt      : TVarType;
      n       : string;
    begin
      with fOutput, vti do
        begin
          Write( ' type="' );
          // Clear the array flags from VarType to get the base type.
          vt := VarType and not ( VT_ARRAY or VT_VECTOR );
          // Check for an intrinsic VB type.
          if vt <> VT_EMPTY
            then Write( VB_GetIntrinsicTypeName( Variants.VarType( TypedVariant ) ) )
            else
              // The type is user-defined.
              try
                n := TypeInfo.Name;
                if n[1] = '_'
                  then delete( n, 1, 1 );                  
                if not IsExternalType
                  then Write( n )
                  else Write( TypeInfo.Parent.Name + '.' + n );
              except
                Write( '?' );
              end;
          Write( '"' );
        end;
    end;

  procedure TGenVB2D.WriteAttr( const name, value : string );
    begin
      fOutput.WriteFmt( ' %s="%s"', [name, value] );
    end;

  function TGenVB2D.WriteAttrIf( const name, value : string; cond : boolean = true ) : boolean;
    begin
      if cond
        then WriteAttr( name, value );
      result := cond;
    end;

//  procedure TGenVB2D.WriteParam( const param : ParameterInfo );
//    begin
//      with fOutput, param do
//        begin
//          Write( '<param type="');
//          //WriteVarType( param.VarTypeInfo );
//          WriteLn( '"/>' );
//        end;
//    end;

  procedure TGenVB2D.WriteMember( const typ : TypeInfo; const mem : MemberInfo; invs : InvokeKinds = 0 );
    var
      n : string;
      h : string;
      argc : integer;
    begin
      with fOutput, mem do
        begin
          if fWriteHelpStr
            then
              begin
                h := HelpString[0];
                if h <> ''
                  then WriteLn( '<!-- ' + h + ' -->' );
              end;
          case DescKind of
            DESCKIND_VARDESC :
              begin
                if invs and INVOKE_CONST <> 0
                  then
                    begin
                      Write( '<const'  );
                      WriteAttr( 'name', Name );
                      // Write constants values only if they are numeric.
                      if fConstValues
                        then WriteAttrIf( 'value', Value, VarIsNumeric( Value ) or VarIsFloat( Value ) );
                      WriteLn( '/>' );
                    end
                  else
                    begin
                      Write( '<var' );
                      WriteAttr( 'name', Name );
                      WriteTypeAttr( ReturnType );
                      WriteLn( '/>' );
                    end;
              end;
            DESCKIND_FUNCDESC :
              begin
                argc := mem.Parameters.Count;
                n := VB_DemangleName( Name );
                if ( invs and INVOKE_PROPERTYGET ) <> 0
                  then
                    begin
                      Write( '<propget' );
                      WriteAttr( 'name', n );
                      WriteAttrIf( 'argc', IntToStr( argc ), argc > 0 );
                      WriteTypeAttr( ReturnType );
                      WriteLn( '/>' );
                    end;
                if invs and ( INVOKE_PROPERTYPUT or INVOKE_PROPERTYPUTREF ) <> 0
                  then
                    begin
                      Write( '<propput' );
                      WriteAttr( 'name', n );
                      WriteAttrIf( 'argc', IntToStr( argc ), argc > 0 );
                      WriteLn( '/>' );
                    end;
                if invs and INVOKE_FUNC <> 0
                  then
                    begin
                      if ( ReturnType.VarType = VT_HRESULT ) or ( ReturnType.VarType = VT_VOID )
                        then
                          begin
                            //System.writeln( ReturnType.VarType );
                            Write( '<sub' );
                            WriteAttr( 'name', n );
                            WriteAttrIf( 'argc', IntToStr( argc ), argc > 0 );
                            WriteLn( '/>' );
                          end
                        else
                          begin
                            Write( '<func' );
                            WriteAttr( 'name', n );
                            WriteAttrIf( 'argc', IntToStr( argc ), argc > 0 );
                            WriteTypeAttr( ReturnType );
                            WriteLn( '/>' );
                          end;
                    end;
                if invs and INVOKE_EVENTFUNC <> 0
                  then
                    begin
                      Write( '<event' );
                      WriteAttr( 'name', n );
                      WriteLn( '/>' );
                    end;
  //                  with Parameters do
  //                    if Count > 0
  //                      then
  //                        begin
  ////                          Indent;
  ////                          for i := 1 to Count do
  ////                            WriteParam( Parameters.Item[i] );
  ////                          Outdent;
  //                        end
              end;
          end;
        end;

    end;

  procedure TGenVB2D.Convert( const TypeLibPath, OutputMapPath : string );
    var
      elem, h  : string;
      members,
      types,
      sr       : SearchResults;
      msi, tsi : SearchItem;
      ti       : TypeInfo;
      i, j     : integer;
      tlb      : TypeLibInfo;
    begin
      writeln( sInfo, 'generating "', OutputMapPath, '"' );
      sr := nil;
      fOutput := nil;
      try
        TLI.ResolveAliases := true;
        tlb := TLI.TypeLibInfoFromFile( TypeLibPath );
        fOutput := TIndentedTextWriter.Create( OutputMapPath );
        with fOutput, tlb do
          begin
            WriteLn('<?xml version="1.0" encoding="iso-8859-1" ?>');
            WriteFmt( '<!-- Generated by GenVB2D from %s on %s -->', [TypeLibPath, DateTimeToStr( Now )]);
            WriteLn;
            if fWriteHelpStr
              then
                begin
                  h := HelpString[0];
                  if h <> ''
                    then WriteLn( '<!-- ' + h + ' -->' );
                end;
            Write( '<library' );
            WriteAttr( 'name', Name );
            WriteAttr( 'majorver', IntToStr( MajorVersion ) );
            WriteAttr( 'minorver', IntToStr( MinorVersion ) );
            WriteLn( '>');
            Indent;
            types := GetTypes( sr, tliStDefault, true );
            for i := 1 to types.Count do
              begin
                tsi := types.Item[i];
                ti := tlb.TypeInfos.IndexedItem[tsi.TypeInfoNumber];
                case ti.TypeKind of
                  TKIND_ENUM :
                    if fIgnoreEnum
                      then continue
                      else elem := 'enum';
                  TKIND_MODULE :
                    if fIgnoreModule
                      then continue
                      else elem := 'module';
                  TKIND_INTERFACE,
                  TKIND_DISPATCH,
                  TKIND_COCLASS :
                    if fIgnoreClass
                      then continue
                      else elem := 'class';
                  TKIND_RECORD :
                    if fIgnoreRecord
                      then continue
                      else elem := 'record';
                  else continue;
                end;

                if fWriteHelpStr
                  then
                    begin
                      h := ti.HelpString[0];
                      if h <> ''
                        then WriteLn( '<!-- ' + h + ' -->' );
                    end;
                    
                Write( '<' + elem );
                WriteAttr( 'name', ti.Name );
                WriteLn( '>' );
                Indent;

                // Get all members of the type info.
                members := tlb.GetMembers( tsi.SearchData, true );
                for j := 1 to members.Count do
                  begin
                    msi := members.Item[j];
                    WriteMember( ti, tlb.GetMemberInfo( tsi.SearchData, msi.InvokeKinds, msi.MemberId, '' ), msi.InvokeKinds );
                  end;

                Outdent;
                WriteLn( '</' + elem + '>' );
              end;
            Outdent;
            WriteLn( '</library>' );
          end;
      finally
        fOutput.Free;
      end;
    end;

  procedure TGenVB2D.ShowUsage;
    begin
      writeln( 'GenVB2D - Generates skeleton conversion maps files. v' + VersionFixedFileInfoString( ParamStr( 0 ), vfFull, '?.?.?.?' ) );
      writeln( 'Copyright(c) 2008-2009 Albert Almeida (caviola@gmail.com)' );
      writeln( 'All rights reserved' );
      writeln;
      writeln( 'Usage:' );
      writeln( '   GenVB2D typelibrary [outputdir] [options]' );
      writeln;
      writeln( '   typelibrary   Specifies the type library.' );
      writeln( '   outputdir     Destination directory for generated conversion map(s).' );
      writeln( '                 Defaults to "VB2D" under the directory of <typelibrary>.' );
      writeln( '   -ie           Ignore enumerated types.' );
      writeln( '   -ir           Ignore records.' );
      writeln( '   -im           Ignore modules(static function and/or constants groups).' );
      writeln( '   -ic           Ignore classes.' );
      writeln( '   -cv           Include enumerated constants values.' );
      writeln( '   -ow           Overwrite existing map files in <outpurdir>.' );
      writeln( '   -ih           Include the help string.' );
    end;

  begin
    CoInitialize( nil );
    writeln;
    with TGenVB2D.Create do
      try
        ExitCode := Run;
      finally
        Free;
      end;
  end.

