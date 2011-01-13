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

program VB2Delphi;

{$APPTYPE CONSOLE}
{$R *.res}

  uses
  FastMM4,
  Variants,
  Classes,
  SysUtils,
  StrUtils,
  JclFileUtils,
  RegExpr,
  CommandLine,
  vbLexer in 'vbLexer.pas',
  vbTokens in 'vbTokens.pas',
  vbError in 'vbError.pas',
  vbParser in 'vbParser.pas',
  vbNodes in 'vbNodes.pas',
  vbDelphiConverter in 'vbDelphiConverter.pas',
  vbGlobals in 'vbGlobals.pas',
  vbOptions in 'vbOptions.pas',
  vbMap in 'vbMap.pas',
  vbClasses in 'vbClasses.pas',
  vbResolver in 'vbResolver.pas',
  vbConsts in 'vbConsts.pas',
  vbBuffers in 'vbBuffers.pas';

  type
    StdLibraries = ( sl_VB, sl_VBA, sl_VBRUN );

  const
    sError  = '[ERROR] ';
    sInfo   = '[Info] ';
    sOption = '[Option] ';
    sWarn   = '[Warning] ';

    DefaultOutputSubDir = 'vb2delphi\';
    DefaultMapsSubDir   = 'maps\';
    MapFileExt          = '.vb2d';

    StdLibraryMaps : array[StdLibraries] of string = ('VB', 'VBA', 'VBRUN');
    
  type

    TVB2Delphi =
      class
        private
          fStdMaps          : array[StdLibraries] of TvbStdModuleDef;
          fConverter        : TvbDelphiConverter;
          fInputFile        : string;
          fInputDir         : string;
          fOptions          : TvbOptions;
          fOutputDir        : string;
          fParser           : TvbParser;
          fSourceCode       : TvbFileBuffer;
          fMapLoader        : TvbMapLoader;
          fResourceProvider : IvbResourceProvider;
          fMonitor          : IvbProgressMonitor;
          fGlobalClass      : TvbClassDef;
          fControlLevel     : integer;
          function CreateDefaultProject : TvbProject;
          procedure ParseCommandLine;
          procedure ParseCommandLineArgument( var p : integer; const arg : string );
          procedure ParseCommandLineOption( var p : integer; const opt : string );
          function ParseModule( const Name, RelPath, AbsPath : string ) : TvbModuleDef;
          procedure ConvertFile( const Path, fOutputDir : string );
          procedure ConvertModule( const Path, fOutputDir : string );
          procedure ConvertProject( const Path, fOutputDir : string );
          procedure ConvertProjectGroup( const Path, fOutputDir : string );
          procedure LoadStandardMaps;
          procedure ShowUsage;

          procedure FormDefinitionProblemHandler( line, column : cardinal; const msg : string );
          procedure ParseModuleBeginHandler( const name, relativePath : string );
          procedure ParseModuleEndHandler;
          procedure ParseModuleFailHandler;
          procedure ParseProjectBeginHandler( const name, relativePath : string );
          procedure ParseProjectEndHandler;
          procedure ParseProjectFailHandler;
          procedure ParseFormControlBeginHandler( control : TvbControl );
          procedure ParseFormControlEndHandler;
          procedure ParseFormDefinitionBeginHandler( form : TvbFormDef );
          procedure ParseFormDefinitionEndHandler;
        public
          constructor Create;
          destructor Destroy; override;
          function Run : integer;
          procedure Log( const msg : string; line : cardinal; col : cardinal; logType : TvbLogType = ltInfo ); overload;
          procedure Log( const msg : string; logType : TvbLogType = ltInfo ); overload;
      end;

  { TVB2Delphi }

  procedure TVB2Delphi.FormDefinitionProblemHandler( line, column : cardinal; const msg : string );
    begin
      writeln( sWarn, Format( '' , [msg] ) );
    end;

  procedure TVB2Delphi.ParseModuleBeginHandler( const name, relativePath : string );
    begin
      writeln( sInfo, Format( 'reading module %s in %s', [name, relativePath] ) );
    end;

  procedure TVB2Delphi.ParseModuleEndHandler;
    begin
    end;

  procedure TVB2Delphi.ParseModuleFailHandler;
    begin
      //writeln( sError, 'parsing module failed!' );
    end;

  procedure TVB2Delphi.ParseProjectBeginHandler( const name, relativePath : string );
    begin
      writeln( sInfo, Format( 'reading project %s', [relativePath] ) );
    end;

  procedure TVB2Delphi.ParseProjectEndHandler;
    begin
    end;

  procedure TVB2Delphi.ParseProjectFailHandler;
    begin
      //writeln( sError, 'parsing module failed!' );
    end;

  procedure TVB2Delphi.ParseFormControlBeginHandler( control : TvbControl );
    begin
      writeln( sInfo, Format( DupeString( #32, fControlLevel * 2) + '%s %s', [control.VarType.Name, control.Name]) );
      inc( fControlLevel );
    end;

  procedure TVB2Delphi.ParseFormControlEndHandler;
    begin
      dec( fControlLevel );
    end;

  procedure TVB2Delphi.ParseFormDefinitionBeginHandler( form : TvbFormDef );
    begin
      fControlLevel := 0;  
    end;

  procedure TVB2Delphi.ParseFormDefinitionEndHandler;
    begin
      writeln( sInfo, DupeString( '=', 70 ) );
    end;

  procedure TVB2Delphi.Log( const msg : string; line : cardinal; col : cardinal; logType : TvbLogType = ltInfo );
    begin
      if fMonitor <> nil
        then fMonitor.Log( msg, line, col, logType );
    end;

  procedure TVB2Delphi.Log( const msg : string; logType : TvbLogType = ltInfo );
    begin
      if fMonitor <> nil
        then fMonitor.Log( msg, logType );
    end;

   procedure TVB2Delphi.LoadStandardMaps;
    var
      ctrl : TvbDef;
      i    : StdLibraries;
      VB   : TvbStdModuleDef;
    begin
      Log( 'loading VB standard library maps' );

      // load the maps for the standard library.
      for i := low( StdLibraryMaps ) to high( StdLibraryMaps ) do
        if not fMapLoader.Load( StdLibraryMaps[i], fStdMaps[i] )
          then Log( Format('map for standard library "%s" not found', [StdLibraryMaps[i]] ), ltWarn );

      //    (2) search for a class named "Form" and get it's scope.
      //        the "Form" scope will be the base scope for all user-defined forms.
      VB := fStdMaps[sl_VB];
      if VB <> nil
        then
          begin
            assert( VB.Scope <> nil );
            // search for the "Global" class.
            VB.Scope.Lookup( TYPE_BINDING, 'Global', ctrl, false );
            if ( ctrl <> nil ) and ( ctrl.NodeKind = CLASS_DEF )
              then fGlobalClass := TvbClassDef( ctrl );
          end;
    end;

  function TVB2Delphi.ParseModule( const Name, RelPath, AbsPath : string ) : TvbModuleDef;
    var
      ext : string;
    begin
      assert( Name <> '' );
      assert( RelPath <> '' );
      assert( AbsPath <> '' );
      fSourceCode.buffFile := AbsPath;
      ext := lowercase( ExtractFileExt( AbsPath ) );
      if ext = '.bas'
        then result := fParser.ParseStdModule( Name, RelPath, fSourceCode )
        else
      if ext = '.cls'
        then result := fParser.ParseClassModule( Name, RelPath, fSourceCode )
        else
      if ext = '.frm'
        then result := fParser.ParseFormModule( Name, RelPath, fSourceCode )
        else raise EvbError.CreateFmt( 'unsupported file type "%s"', [AbsPath] );
      result.RelativePath := RelPath;
    end;

  function TVB2Delphi.CreateDefaultProject : TvbProject;
    var
      i : StdLibraries;
    begin
      result := TvbProject.Create;
      try
        assert( result.Scope <> nil );
        result.Name := '_PROJECT';

        // import the maps for the VB standard library into the project.
        for i := low( fStdMaps ) to high( fStdMaps ) do
          if fStdMaps[i] <> nil
            then result.AddImportedModule( fStdMaps[i] );

        // if we've got the "Global" class, declare it's members in the
        // project/global scope.
        if fGlobalClass <> nil
          then result.Scope.BindMembers( fGlobalClass );
      except;
        result.Free;
        raise;
      end;
    end;

  procedure TVB2Delphi.ConvertModule( const Path, fOutputDir : string );
    var
      Module         : TvbModuleDef;
      ModuleName     : string;
      OutputFileName : string;
      OutputStream   : TStream;
      Project        : TvbProject;
    begin
      OutputStream := nil;
      Project := CreateDefaultProject;
      try
        ModuleName     := ChangeFileExt( ExtractFileName( Path ), '' );
        OutputFileName := fOutputDir + ModuleName + '.pas';
        OutputStream   := TFileStream.Create( OutputFileName, fmCreate );
        Module         := Project.AddModule( ParseModule( ModuleName, ExtractFileName( Path ), Path ) );
        ResolveProject( Project, fOptions );
        case Module.NodeKind of
          STD_MODULE_DEF   : fConverter.ConvertStdModule( TvbStdModuleDef( Module ), OutputStream );
          CLASS_MODULE_DEF : fConverter.ConvertClassModule( TvbClassModuleDef( Module ), OutputStream );
          FORM_MODULE_DEF  : fConverter.ConvertFormModule( TvbFormModuleDef( Module ), OutputStream );
        end;        
      finally
        OutputStream.Free;
        Project.Free;
      end;
    end;

  procedure TVB2Delphi.ConvertProject( const Path, fOutputDir : string );
    var
      i        : StdLibraries;
      project  : TvbProject;
      settings : TStringList;
    begin
      project := nil;
      settings := TStringList.Create;
      try
        settings.LoadFromFile( Path );
        project := fParser.ParseProject( settings, ExtractFileName( Path ) );
        assert( project.Scope <> nil );

        // import the maps for the VB standard library into the project.
        for i := low( fStdMaps ) to high( fStdMaps ) do
          if fStdMaps[i] <> nil
            then project.AddImportedModule( fStdMaps[i] );

        // if we've got the "Global" class, declare it's members in the
        // project/global scope.
        if fGlobalClass <> nil
          then project.Scope.BindMembers( fGlobalClass );

        ResolveProject( project, fOptions );
        fConverter.ConvertProject( project, );
      finally
        FreeAndNil( settings );
        FreeAndNil( project );
      end;
    end;

  procedure TVB2Delphi.ConvertProjectGroup( const Path, fOutputDir : string );
    begin
    end;

  procedure TVB2Delphi.ConvertFile( const Path, fOutputDir : string );
    var
      ext : string;      
    begin
      //fMonitor.Log( Format( 'converting %s', [Path] ) );
      ext := lowercase( ExtractFileExt( Path ) );
      if ( ext = '.bas' ) or
         ( ext = '.cls' ) or
         ( ext = '.frm' )
        then ConvertModule( Path, fOutputDir )
        else
      if ext = '.vbp'
        then ConvertProject( Path, fOutputDir )
        else
      if ext = '.vbg'
        then ConvertProjectGroup( Path, fOutputDir )
        else assert( false );
    end;

  procedure TVB2Delphi.ParseCommandLineOption( var p : integer; const opt : string );
    begin
      with fOptions do
        if opt = 'as'
          then optUseAnsiString := true
          else
        if opt = 'bb'
          then optUseByteBool := true
          else
        if opt = 'ic'
          then optEnumToConst := true
          else
        if opt = 's1c'
          then optString1ToChar := true
          else
        if opt = 'sna'
          then optStringNToArrayOfChar := true
          else
        if opt = 'bct'
          then optCommentsBelowFuncToTop := true
          else
        if opt = 'byval'
          then optParamDefaultByVal := true
          else
        if opt = 'MD'
          then
            begin
              inc( p );
              if ParamStr( p ) = ''
                then raise EvbError.Create( 'where to read maps from?' );                
              fMapLoader.AddSearchDir( ParamStr( p ) );
            end
          else raise EvbError.CreateFmt( 'invalid command line option "%s"', [opt] );
    end;

  procedure TVB2Delphi.ParseCommandLineArgument( var p : integer; const arg : string );
    var
      ext : string;
    begin
      if fInputFile = ''
        then
          begin
            ext := lowercase( ExtractFileExt( arg ) );
            if ( ext <> '.bas' ) and
               ( ext <> '.cls' ) and
               ( ext <> '.frm' ) and
               ( ext <> '.vbp' )
              then raise EvbError.Create( 'only .BAS, .CLS, .FRM or .VBP files supported' );
            fInputFile := arg;
          end
        else
      if fOutputDir = ''
        then fOutputDir := IncludeTrailingPathDelimiter( arg )
        else raise EvbError.CreateFmt( 'don''t know what to do with argument "%s"', [arg] );
    end;

  procedure TVB2Delphi.ParseCommandLine;
    begin
      CommandLine.ParseCommandLine( ParseCommandLineOption, ParseCommandLineArgument );
      if fInputFile = ''
        then raise EvbError.Create( 'no input file provided' );
      fInputDir := ExtractFilePath( fInputFile );
      if fOutputDir = ''
        then fOutputDir := ExtractFilePath( fInputFile ) + DefaultOutputSubDir
        else fOutputDir := IncludeTrailingPathDelimiter( fOutputDir );

      Log( Format( 'output directory: %s', [fOutputDir] ) );

      fResourceProvider.InputDir := fInputDir;
      fResourceProvider.OutputDir := fOutputDir;
    end;

  constructor TVB2Delphi.Create;
    begin
      fOptions          := TvbOptions.Create;
      fSourceCode       := TvbFileBuffer.Create;
      fMapLoader        := TvbMapLoader.Create;
      fMonitor          := TvbConsoleProgressMonitor.Create;
      fResourceProvider := TvbFileBasedResourceProvider.Create( '', '', fMapLoader );

      fParser := TvbParser.Create( fOptions, fResourceProvider, fMonitor );
      with fParser do
        begin
          OnFormDefinitionProblem    := FormDefinitionProblemHandler;
          OnParseModuleBegin         := ParseModuleBeginHandler;
          OnParseModuleEnd           := ParseModuleEndHandler;
          OnParseModuleFail          := ParseModuleFailHandler;
          OnParseProjectBegin        := ParseProjectBeginHandler;
          OnParseProjectEnd          := ParseProjectEndHandler;
          OnParseProjectFail         := ParseProjectFailHandler;
          OnParseFormControlBegin    := ParseFormControlBeginHandler;
          OnParseFormControlEnd      := ParseFormControlEndHandler;
          OnParseFormDefinitionBegin := ParseFormDefinitionBeginHandler;
          OnParseFormDefinitionEnd   := ParseFormDefinitionEndHandler;
        end;

      fConverter := TvbDelphiConverter.Create( fOptions, fResourceProvider, fMonitor );
      fMapLoader.AddSearchDir( ExtractFilePath( ParamStr( 0 ) ) + DefaultMapsSubDir );
    end;

  destructor TVB2Delphi.Destroy;
    begin
      FreeAndNil( fParser );
      FreeAndNil( fConverter );
      FreeAndNil( fOptions );
      FreeAndNil( fSourceCode );
      FreeAndNil( fMapLoader );
      inherited;
    end;

  function TVB2Delphi.Run : integer;
    var
      sr : TSearchRec;
    begin
      result := 0;
      if ParamCount > 0
        then
          try
            ParseCommandLine;
            if FindFirst( fInputFile, faAnyFile, sr ) <> 0
              then raise EvbError.CreateFmt( 'can''t read input file "%s"', [fInputFile] );
            if not JclFileUtils.ForceDirectories( fOutputDir )
              then raise EvbError.CreateFmt( 'can''t create output directory ''%s''', [fOutputDir] );
            LoadStandardMaps;
            repeat
              ConvertFile( fInputDir + sr.Name, fOutputDir );
            until FindNext( sr ) <> 0;
            FindClose( sr );
            Log('TRANSLATION SUCCESSFUL!');
          except
            on e : EvbSourceError do
              begin
                writeln( sError, fInputDir + sr.Name, '(', e.Line, ':', e.Column, '): ', e.Message );
                result := 1;
              end;
            on e : Exception do
              writeln( sError, e.Message );
          end
        else ShowUsage;
    end;

  procedure TVB2Delphi.ShowUsage;
    begin
      writeln( 'VB2Delphi - VB to Delphi source file converter. v', VersionFixedFileInfoString( ParamStr( 0 ), vfFull, '?.?.?.?' ) );
      writeln( 'Copyright(c) 2008-2009 Albert Almeida (caviola@gmail.com)' );
      writeln( 'All rights reserved' );
      writeln;
      writeln( 'Usage:' );
      writeln( '   vb2delphi file [outputdir] [options]' );
      writeln;
      writeln( '   file         Specifies the file(s) to convert. Wildcards allowed.' );
      writeln( '   outputdir    Destination directory for generated Delphi files.' );
      writeln( '                Defaults to "vb2delphi" under the directory of <file>.' );
      writeln( '   -as          Translate "String" to "AnsiString".' );
      writeln( '                Default is "WideString".' );
      writeln( '   -bb          Translate "Boolean" to "ByteBool".' );
      writeln( '                Default is "WordBool".' );
      writeln( '   -ic          Translate enumerated types to integer constants.' );
      writeln( '                Only when all constants have explicit value.' );
      writeln( '   -s1c         Translate "String(1)" to "char".' );
      writeln( '   -sna         Translate "String(n)" to "array[1..N] of char".' );
      writeln( '   -bct         Moves the comments below a function prototype to the top.' );
      writeln( '   -byval       Declare parameters "by value" by default.' );
      //writeln( '   -M mapfile   Loads the specified map if it isn''t already loaded.' );
      //writeln( '   -Cname=value   Defines a conditional compilation constant.' );
      writeln( '   -MD mapdir   Location of maps that are automatically loaded. These are' );
      writeln( '                the maps for the VB standard libraries VB, VB6 and VBRUN' );
      writeln( '                and for the libraries referenced in a project file.' );
      writeln( '                By default all maps are searched for in the "maps" subdirectory' );
      writeln( '                in the inslatation directory.' );
    end;

  begin
    writeln;
    with TVB2Delphi.Create do
      try
        ExitCode := Run;
      finally
        Free;
      end;

    //readln;

  end.


