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

unit vbClasses;

interface

  uses
    Classes,
    SysUtils,
    frxReader,
    vbBuffers,
    vbNodes,
    vbMap;

  type

    EvbError = class( Exception );

    // This represents an error in the source code.
    // This can either be a lexer error or a parser syntax error.
    EvbSourceError =
      class( EvbError )
        private
          fLine   : cardinal;
          fColumn : cardinal;
        public
          constructor Create( const msg : string; line : cardinal = 0; col : cardinal = 0 );
          constructor CreateFmt( const fmt : string; args : array of const; line : cardinal = 0; col : cardinal = 0 );
          property Line   : cardinal read fLine   write fLine;
          property Column : cardinal read fColumn write fColumn;
      end;

    TvbLogType = (
      ltInfo,
      ltWarn,
      ltError
    );

    IvbProgressMonitor =
      interface
        ['{23FCDF58-885B-4A5F-BFC8-4532F048DC23}']
        procedure Log( const msg : string; logType : TvbLogType = ltInfo ); overload;
        procedure Log( const msg : string; line, col : cardinal; logType : TvbLogType = ltInfo ); overload;
        procedure ConvertModuleBegin( module : TvbModuleDef );
        procedure ConvertModuleEnd( module : TvbModuleDef );
        procedure ConvertProjectBegin( project : TvbProject );
        procedure ConvertProjectEnd( project : TvbProject );
      end;

    TvbConsoleProgressMonitor =
      class( TInterfacedObject, IvbProgressMonitor )
        public
          // IvbProgressMonitor
          procedure Log( const msg : string; logType : TvbLogType = ltInfo ); overload;
          procedure Log( const msg : string; line, col : cardinal; logType : TvbLogType = ltInfo ); overload;
          procedure ConvertModuleBegin( module : TvbModuleDef );
          procedure ConvertModuleEnd( module : TvbModuleDef );
          procedure ConvertProjectBegin( project : TvbProject );
          procedure ConvertProjectEnd( project : TvbProject );
      end;

    TvbPropertyMap =
      class
        private
          fMap : TStringList;
          function GetProp( i : integer ) : TvbPropertyDef;
          function GetPropCount : integer;
        public
          constructor Create;
          destructor Destroy; override;
          function AddPropFunc( func : TvbAbstractFuncDef ) : TvbAbstractFuncDef;
          procedure Clear;
          property Prop[i : integer] : TvbPropertyDef read GetProp; default;
          property PropCount         : integer        read GetPropCount;
      end;

    TvbModuleTypeMap =
      class
        private
          fMap : TStringList;
        public
          constructor Create;
          destructor Destroy; override;
          // Adds the definition for a new type.
          // This method also creates an instance of IvbType that represents
          // the inserted type.
          procedure DefineType( typeDef : TvbTypeDef );
          function GetType( const typeName : string ) : IvbType;
          procedure Clear;
      end;

    TvbStatistics =
      class
      end;

    IvbResourceProvider =
      interface
        ['{733CA269-F4A5-4360-9A91-B03DBAFA0E53}']
        function GetSourceCode( const path : string ) : TvbCustomTextBuffer;
        function CreateOutputStream( const path : string ) : TStream;
        function LoadTranslationMap( const name : string ) : TvbStdModuleDef;
        //function GetFormItemArray( const frxName : string ) : TfrxItemArray;
        function GetInputDir : string;
        procedure SetInputDir( const dir : string );
        procedure SetOutputDir( const dir : string );
        function GetOutputDir : string;
        property InputDir  : string read GetInputDir  write SetInputDir;
        property OutputDir : string read GetOutputDir write SetOutputDir;
      end;

    TvbFileBasedResourceProvider =
      class( TInterfacedObject, IvbResourceProvider )
        private
          fSource : TvbFileBuffer;
          fInputDir : string;
          fOutputDir : string;
          fMapLoader : TvbMapLoader;
          //fFRXReader : TfrxReader;
        public
          constructor Create( const inputDir, outputDir : string; mapLoader : TvbMapLoader );
          destructor Destroy; override;
          // IvbResourceProvider
          function CreateOutputStream( const path : string ) : TStream;
          function GetSourceCode( const path : string ) : TvbCustomTextBuffer;
          function LoadTranslationMap( const name: string ) : TvbStdModuleDef;
          function GetInputDir : string;
          function GetOutputDir : string;
          procedure SetInputDir( const dir : string );
          procedure SetOutputDir( const dir : string );
          //function GetFormItemArray(const frxName : string): TfrxItemArray;
      end;

  function CreateBufferFromFile( const path : string ) : TvbChunkBuffer;

  function CreateDynArrayOf( const baseType : IvbType ) : IvbType;
  function CreateType( const baseType : IvbType; isArray : boolean; arrayBounds : TvbArrayBoundList = nil ) : IvbType;

implementation

  uses
    JclFileUtils,
    CollWrappers;

  const
    LogTypeMsg : array[TvbLogType] of string = ( '[Info] ', '[Warning] ', '[Error] ' );

  function CreateBufferFromFile( const path : string ) : TvbChunkBuffer;
    begin
      result := nil;
    end;

  function CreateDynArrayOf( const baseType : IvbType ) : IvbType;
    begin
      with baseType do
        case Code of
          STRING_TYPE   : result := vbStringDynArrayType;
          INTEGER_TYPE  : result := vbIntegerDynArrayType;
          BOOLEAN_TYPE  : result := vbBooleanDynArrayType;
          LONG_TYPE     : result := vbLongDynArrayType;
          SINGLE_TYPE   : result := vbSingleDynArrayType;
          DOUBLE_TYPE   : result := vbDoubleDynArrayType;
          BYTE_TYPE     : result := vbByteDynArrayType;
          CURRENCY_TYPE : result := vbCurrencyDynArrayType;
          DATE_TYPE     : result := vbDateDynArrayType;
          VARIANT_TYPE  : result := vbVariantDynArrayType;
          OBJECT_TYPE   : result := vbObjectDynArrayType;
          else
            begin
              // Here we create a dynamic array of an UDT.
              if baseType.TypeDef <> nil
                then result := TvbType.CreateArray( baseType.TypeDef )
                else result := TvbType.CreateArray( baseType.Name );
            end;
        end;
    end;

  function CreateType( const baseType : IvbType; isArray : boolean; arrayBounds : TvbArrayBoundList ) : IvbType;
    begin
      if not isArray
        then result := baseType
        else
          // Check if we are building an dynamic array.
          if arrayBounds = nil
            then result := CreateDynArrayOf( baseType )
            else result := TvbType.CreateVector( baseType, arrayBounds );
    end;

  { EvbSourceError }

  constructor EvbSourceError.Create( const msg : string; line, col : cardinal );
    begin
      inherited Create( msg );
      fLine   := line;
      fColumn := col;
    end;

  constructor EvbSourceError.CreateFmt( const fmt : string; args : array of const; line, col : cardinal );
    begin
      Create( Format( fmt, args ), line, col );
    end;

  { TvbPropertyMap }

  function TvbPropertyMap.AddPropFunc( func : TvbAbstractFuncDef ) : TvbAbstractFuncDef;
    var
      theProperty : TvbPropertyDef;
      i : integer;
    begin
      result := func;
      assert( func.Name <> '' );
      with fMap do
        // search for an already existent property.
        if Find( func.Name, i )
          then theProperty := TvbPropertyDef( Objects[i] )
          else
            begin
              // create a new property and insert it in our internal map.
              theProperty :=  TvbPropertyDef.Create( func.Name );
              AddObject( func.Name, theProperty );
            end;
            
      if dfPropGet in func.NodeFlags
        then
          begin
            assert( theProperty.GetFn = nil );
            theProperty.GetFn := func;
          end
        else
      if dfPropSet in func.NodeFlags
        then
          begin
            assert( theProperty.SetFn = nil );
            theProperty.SetFn := func;
          end
        else
          begin
            assert( theProperty.LetFn = nil );
            theProperty.LetFn := func;
          end;
    end;

  procedure TvbPropertyMap.Clear;
    begin
      fMap.Clear;
    end;

  constructor TvbPropertyMap.Create;
    begin
      fMap        := TStringList.Create;
      fMap.Sorted := true;
    end;

  destructor TvbPropertyMap.Destroy;
    begin
      FreeAndNil( fMap );
      inherited;
    end;

  function TvbPropertyMap.GetProp( i : integer ) : TvbPropertyDef;
    begin
      result := TvbPropertyDef( fMap.Objects[i] );
    end;

  function TvbPropertyMap.GetPropCount : integer;
    begin
      result := fMap.Count;
    end;

  { TvbModuleTypeMap }

  procedure TvbModuleTypeMap.DefineType( typeDef : TvbTypeDef );
    var
      tr : IvbType;
      n  : string;
      i  : integer;
    begin
      n := typeDef.Name;
      assert( n <> '' );
      // check if there is a reference to the given type.  
      if fMap.Find( n, i )
        then
          begin
            // set the TypeDef property in the corresponding IvbType instance.
            tr := TInterfaceWrapper( fMap.Objects[i] ).Value as IvbType;
            assert( tr.TypeDef = nil );
            tr.TypeDef := typeDef;
          end
        else
          begin
            // there are no references to the new type yet.
            // create a new type reference for this new type and add it to
            // the map.
            fMap.AddObject( n, TInterfaceWrapper.Create( TvbType.CreateUDT( typeDef ) ) );
          end;
    end;

  procedure TvbModuleTypeMap.Clear;
    var
      i : integer;
    begin
      for i := 0 to fMap.Count - 1 do
        fMap.Objects[i].Free;
      fMap.Clear;
    end;

  constructor TvbModuleTypeMap.Create;
    begin
      fMap := TStringList.Create;
      fMap.CaseSensitive := false;
      fMap.Sorted := true;
    end;

  destructor TvbModuleTypeMap.Destroy;
    begin
      Clear;
      FreeAndNil( fMap );
      inherited;
    end;

  function TvbModuleTypeMap.GetType( const typeName : string ) : IvbType;
    var
      i : integer;
    begin
      assert( typeName <> '' );
      // if we already created a reference to the given type return it to the caller.
      if fMap.Find( typeName, i )
        then result := TInterfaceWrapper( fMap.Objects[i] ).Value as IvbType
        else
          begin
            // there was no reference so create a new one.
            result := TvbType.CreateUDT( typeName );
            fMap.AddObject( typeName, TInterfaceWrapper.Create( result ) );
          end;
    end;

  { TvbFileBasedResourceProvider }

  constructor TvbFileBasedResourceProvider.Create( const inputDir, outputDir : string; mapLoader : TvbMapLoader );
    begin
      assert( mapLoader <> nil );
      SetInputDir( inputDir );
      SetOutputDir( outputDir );
      fSource    := TvbFileBuffer.Create;
      fMapLoader := mapLoader;
      //fFRXReader := TfrxReader.Create;
    end;

  function TvbFileBasedResourceProvider.CreateOutputStream( const path : string ) : TStream;
    var
      filePath : string;
    begin
      filePath := fOutputDir + path;
      if not JclFileUtils.ForceDirectories( ExtractFilePath( filePath ) )
        then raise Exception.CreateFmt( 'couldn''t create "%s"', [filePath] );
      result := TFileStream.Create( filePath, fmCreate );
    end;

  destructor TvbFileBasedResourceProvider.Destroy;
    begin
      FreeAndNil( fSource );
      //FreeAndNil( fFRXReader );
      inherited;
    end;

//  function TvbFileBasedResourceProvider.GetFormItemArray( const frxName : String ) : TfrxItemArray;
//    var
//      i       : integer;
//      theItem : IfrxItem;
//      j       : integer;
//    begin
//      result := nil;
//      with fFRXReader do
//        if SetFile( fInputDir + frxName )
//          then
//            begin
//              ReadItems;
//              if ItemCount > 0
//                then
//                  begin
//                    j := 0;
//                    setlength( result, ItemCount );
//                    for i := low( result ) to high( result ) do
//                      begin
//                        theItem := Item[i];
//                        if not ( theItem.ItemType in [fitText, fitBLOB] )
//                          then
//                            begin
//                              result[j] := theItem;
//                              inc( j );
//                            end;
//                      end;
//                    if j < length( result )
//                      then setlength( result, j );                      
//                  end;
//            end;
//    end;

  function TvbFileBasedResourceProvider.GetInputDir : string;
    begin
      result := fInputDir;
    end;

  function TvbFileBasedResourceProvider.GetOutputDir : string;
    begin
      result := fOutputDir;
    end;

  function TvbFileBasedResourceProvider.GetSourceCode( const path : string ) : TvbCustomTextBuffer;
    var
      filePath : string;
    begin
      result := nil;
      filePath := fInputDir + path;
      if FileExists( filePath )
        then
          begin
            fSource.buffFile := filePath;
            result := fSource;
          end;
    end;

  function TvbFileBasedResourceProvider.LoadTranslationMap( const name : string ) : TvbStdModuleDef;
    begin
      fMapLoader.Load( name, result );
    end;

  procedure TvbFileBasedResourceProvider.SetInputDir( const dir : string );
    begin
      fInputDir := IncludeTrailingPathDelimiter( dir );
    end;

  procedure TvbFileBasedResourceProvider.SetOutputDir( const dir : string );
    begin
      fOutputDir := IncludeTrailingPathDelimiter( dir );
    end;

  { TvbConsoleProgressMonitor }

  procedure TvbConsoleProgressMonitor.Log( const msg : string; logType : TvbLogType );
    begin
      writeln( LogTypeMsg[logType], msg );
    end;

  procedure TvbConsoleProgressMonitor.ConvertModuleBegin( module : TvbModuleDef );
    begin
      Log( Format( 'converting %s', [module.Name] ) );
    end;

  procedure TvbConsoleProgressMonitor.ConvertModuleEnd( module : TvbModuleDef );
    begin

    end;

  procedure TvbConsoleProgressMonitor.ConvertProjectBegin( project : TvbProject );
    begin
      Log( Format( 'converting %s', [project.Name] ) );
    end;

  procedure TvbConsoleProgressMonitor.ConvertProjectEnd( project : TvbProject );
    begin

    end;

  procedure TvbConsoleProgressMonitor.Log( const msg : string; line, col : cardinal; logType : TvbLogType );
    begin
      writeln( LogTypeMsg[logType], msg );
    end;

end.

























