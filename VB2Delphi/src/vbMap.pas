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

unit vbMap;

interface

  uses
    vbNodes,
    Classes;

  const
    MapFileExt = '.vb2d';
    
  type

    // This is fired the first time a map is successfully loaded.
    TvbMapLoadedEvent = procedure ( const mapName, loadedFromDir : string ) of object;
    // This is fired whenever the requested map couldn't be found in the
    // available directories.
    TvbMapNotFoundEvent = procedure ( const mapName : string ) of object;

    TvbMapLoader =
      class
        private
          fLoadedMaps    : TStringList;
          fOnMapLoaded   : TvbMapLoadedEvent;
          fOnMapNotFound : TvbMapNotFoundEvent;
          fSearchDirs    : TStringList;
        public
          constructor Create;
          function AddSearchDir( const dir : string ) : boolean;
          destructor Destroy; override;
          // The TvbStdModuleDef objects returned by this function are OWNED by
          // this class, so no one else should Free them.
          function Load( const map : string; var module : TvbStdModuleDef ) : boolean;
          property OnMapLoaded   : TvbMapLoadedEvent   read fOnMapLoaded   write fOnMapLoaded;
          property OnMapNotFound : TvbMapNotFoundEvent read fOnMapNotFound write fOnMapNotFound;
      end;

implementation

  uses
    SysUtils,
    LibXmlParser,
    RegExpr,
    StrUtils,
    vbClasses;

  const
    // These are the recognized/allowed XML elements in a translation map file.
    elemClass   = 'class';
    elemConst   = 'const';    // eg: <const name="vbCR" map="#13"/>
    elemEnum    = 'enum';     // a containter for <const.../> elements
    elemFunc    = 'func';     
    elemLibrary = 'library';  // the root node of the translation map
    elemModule  = 'module';    
    elemPropGet = 'propget';
    elemPropPut = 'propput';
    elemRecord  = 'record';
    elemSub     = 'sub';
    elemVar     = 'var';

  type
    TLoaderState = (
      smLibrary,
      smModule,
      smRecord,
      smEnum,
      smParseDef,
      smClass,
      smDone );

  function LoadTranslationMap( const Path : string ) : TvbStdModuleDef;
    var
      typeMap : TvbModuleTypeMap;

    function GetTypeRef( const N : string ) : IvbType;
      begin
        if N = ''
          then result := _vbUnknownType
          else
        if N = 'boolean'
          then result := vbBooleanType
          else
        if N = 'byte'
          then result := vbByteType
          else
        if N = 'currency'
          then result := vbCurrencyType
          else
        if N = 'date'
          then result := vbDateType
          else
        if N = 'double'
          then result := vbDoubleType
          else
        if N = 'integer'
          then result := vbIntegerType
          else
        if N = 'long'
          then result := vbLongType
          else
        if N = 'object'
          then result := vbObjectType
          else
        if N = 'single'
          then result := vbSingleType
          else
        if N = 'string'
          then result := vbStringType
          else
        if N = 'variant'
          then result := vbVariantType
          else result := typeMap.GetType( N );
      end;

    var
      baseClassAttr : string;
      baseClassType : IvbType;
      currentClass  : TvbClassDef;
      currentEnum   : TvbEnumDef;
      //currentMethod : TvbFuncDef;
      currentModule : TvbLibModuleDef;
      currentRecord : TvbRecordDef;
      flags         : TvbNodeFlags;
      fullMapAttr   : string;
      mapAttr       : string;
      nameAttr      : string;
      nameMapAttr   : string;
      argcAttr      : string;
      argc          : integer;
      propMap       : TvbPropertyMap;
      state         : TLoaderState;
      typeAttr      : string;
      i             : integer;
    begin
      currentEnum   := nil;
      currentModule := nil;
      currentRecord := nil;
      result        := nil;
      propMap       := nil;
      state         := smLibrary;
      typeMap       := nil;
      
      with TXmlParser.Create do
        try
          typeMap := TvbModuleTypeMap.Create;
          propMap := TvbPropertyMap.Create;
          Normalize := true;
          LoadFromFile( Path );
          StartScan;
          while Scan do
            begin
              if CurPartType in [ptXmlProlog, ptComment, ptPI, ptDtdc, ptCData, ptContent]
                then continue;

              flags := [dfPublic, dfExternal];

              nameAttr    := CurAttr.Value( 'name' );
              typeAttr    := CurAttr.Value( 'type' );
              mapAttr     := CurAttr.Value( 'map' );
              nameMapAttr := CurAttr.Value( 'nameMap' );
              fullMapAttr := CurAttr.Value( 'fullMap' );

              if not TryStrToInt( CurAttr.Value( 'argc' ), argc )
                then argc := 0;

              if nameMapAttr <> ''
                then
                  begin
                    mapAttr := nameMapAttr;
                    Include( flags, map_OnlyName );
                  end
                else
              if fullMapAttr <> ''
                then
                  begin
                    mapAttr := fullMapAttr;
                    Include( flags, map_Full );
                  end
                else
              if mapAttr <> ''
                then Include( flags, map_Default );

              case state of
                smLibrary :
                  begin
                    // if not <library name="..." ...[/]> we have an error.
                    if ( CurPartType = ptEndTag ) or ( CurName <> elemLibrary ) or ( nameAttr = '' )
                      then break;
                    result := TvbStdModuleDef.Create( nameAttr, nameAttr );
                    // if this is <library ... /> we are done.
                    if CurPartType = ptEmptyTag
                      then
                        begin
                          state := smDone;
                          break;
                        end
                      else state := smParseDef;
                  end;
                smParseDef :
                  begin
                    // if this is </library> we are done.
                    if ( CurPartType = ptEndTag ) and ( CurName = elemLibrary )
                      then
                        begin
                          state := smDone;
                          break;
                        end;

                    // if not <element name="..."....[/]> we have an error.
                    if ( CurPartType = ptEndTag ) or ( nameAttr = '' )
                      then break;
                      
                    if CurName = elemEnum
                      then
                        begin
                          currentEnum := result.AddEnum( TvbEnumDef.Create( flags, nameAttr, mapAttr ) );
                          typeMap.DefineType( currentEnum );
                          state := smEnum;
                        end
                      else
                    if CurName = elemRecord
                      then
                        begin
                          currentRecord := result.AddRecord( TvbRecordDef.Create( flags, nameAttr, mapAttr ) );
                          typeMap.DefineType( currentRecord );
                          state := smRecord;
                        end
                      else
                    if CurName = elemModule
                      then
                        begin
                          currentModule := result.AddLibModule( TvbLibModuleDef.Create( flags, nameAttr, mapAttr ) );
                          state := smModule;
                        end
                      else
                    if CurName = elemClass
                      then
                        begin
                          // check for the "baseClass" attribute and if provided
                          // get a type reference for the specified class.
                          baseClassAttr := CurAttr.Value( 'baseClass' );
                          if baseClassAttr <> ''
                            then baseClassType := typeMap.GetType( baseClassAttr )
                            else baseClassType := nil;
                          currentClass := result.AddClass( TvbClassDef.Create( flags, nameAttr, mapAttr, baseClassType ) );
                          typeMap.DefineType( currentClass );
                          state := smClass;
                        end
                      else break; // unknown element!
                    // if this is <element/> restart parsing definitions.
                    if CurPartType = ptEmptyTag
                      then state := smParseDef;
                  end;
                smEnum :
                  begin
                    // if this is </enum> restart parsing definitions.
                    if ( CurPartType = ptEndTag ) and ( CurName = elemEnum )
                      then state := smParseDef
                      else
                        begin
                          if ( CurPartType <> ptEmptyTag ) or ( CurName <> elemConst ) or ( nameAttr = '' )
                            then break; // error!
                          currentEnum.AddEnum( TvbConstDef.Create( flags, nameAttr, vbLongType, nil, mapAttr ) );
                        end;
                  end;
                smRecord :
                  begin
                    // if this is </record> restart parsing definitions.
                    if ( CurPartType = ptEndTag ) and ( CurName = elemRecord )
                      then state := smParseDef
                      else
                        begin
                          // if not <var.../> we have an error.
                          if ( CurPartType <> ptEmptyTag ) or ( CurName <> elemVar ) or ( nameAttr = '' )
                            then break;
                          currentRecord.AddField( TvbVarDef.Create( flags, nameAttr, GetTypeRef( typeAttr ), mapAttr ) );
                        end;
                  end;
                smModule :
                  begin
                    // if this is </module> restart parsing definitions.
                    if ( CurPartType = ptEndTag ) and ( CurName = elemModule )
                      then
                        begin
                          state := smParseDef;
                          for i := 0 to propMap.PropCount - 1 do
                            currentModule.AddProperty( propMap[i] );
                          propMap.Clear;
                        end
                      else
                        begin
                          if map_Default in flags
                            then
                              begin
                                Exclude( flags, map_Default );
                                Include( flags, map_Full );
                                //Include( flags, mapNoQualifier );
                              end;
                          if ( CurPartType <> ptEmptyTag ) or ( nameAttr = '' )
                            then break;
                          if CurName = elemConst
                            then currentModule.AddConstant( TvbConstDef.Create( flags, nameAttr, nil, nil, mapAttr ) )
                            else
                          if CurName = elemFunc
                            then currentModule.AddMethod( TvbExternalFuncDef.Create( flags, nameAttr, mapAttr, argc, GetTypeRef( typeAttr ) ) )
                            else
                          if CurName = elemSub
                            then currentModule.AddMethod( TvbExternalFuncDef.Create( flags, nameAttr, mapAttr, argc ) )
                            else
                          if CurName = elemPropGet
                            then propMap.AddPropFunc( TvbExternalFuncDef.Create( flags + [dfPropGet, dfStatic], nameAttr, mapAttr, argc, GetTypeRef( typeAttr ) ) )
                            else
                          if CurName = elemPropPut
                            then propMap.AddPropFunc( TvbExternalFuncDef.Create( flags + [dfPropSet, dfStatic], nameAttr, mapAttr, argc ) )
                            else
                          if CurName = elemVar
                            then currentClass.AddVar( TvbVarDef.Create( flags, nameAttr, GetTypeRef( typeAttr ), mapAttr ) )
                            else continue; // ignore all other elements
                        end;
                  end;
                smClass :
                  begin
                    // if this is </class> copy all properties to current class
                    // and restart parsing element definitions.
                    if ( CurPartType = ptEndTag ) and ( CurName = elemClass )
                      then
                        begin
                          state := smParseDef;
                          for i := 0 to propMap.PropCount - 1 do
                            currentClass.AddProperty( propMap[i] );
                          propMap.Clear;
                        end
                      else
                        begin
                          if ( CurPartType <> ptEmptyTag ) or ( nameAttr = '' )
                            then break;
                          if CurName = elemFunc
                            then currentClass.AddMethod( TvbExternalFuncDef.Create( flags, nameAttr, mapAttr, argc, GetTypeRef( typeAttr ) ) )
                            else
                          if CurName = elemSub
                            then currentClass.AddMethod( TvbExternalFuncDef.Create( flags, nameAttr, mapAttr, argc ) )
                            else
                          if CurName = elemPropGet
                            then propMap.AddPropFunc( TvbExternalFuncDef.Create( flags + [dfPropGet, dfStatic], nameAttr, mapAttr, argc, GetTypeRef( typeAttr ) ) )
                            else
                          if CurName = elemPropPut
                            then propMap.AddPropFunc( TvbExternalFuncDef.Create( flags + [dfPropSet, dfStatic], nameAttr, mapAttr, argc ) )
                            else
                          if CurName = elemVar
                            then currentClass.AddVar( TvbVarDef.Create( flags, nameAttr, GetTypeRef( typeAttr ), mapAttr ) )
                            else continue; // ignore all other elements
                        end;
                  end;
              end;
            end;
          if state <> smDone
            then FreeAndNil( result );            
        finally
          Free;
          FreeAndNil( propMap );
          FreeAndNil( typeMap );
        end;
    end;

  { TvbMapLoader }

  function TvbMapLoader.AddSearchDir( const dir : string ) : boolean;
    begin
      result := true;
      fSearchDirs.Add( IncludeTrailingPathDelimiter( dir ) );
    end;

  constructor TvbMapLoader.Create;
    begin
      fLoadedMaps := TStringList.Create;
      fLoadedMaps.CaseSensitive := false;
      fLoadedMaps.Duplicates := dupError;
      fLoadedMaps.Sorted := true;
      fSearchDirs := TStringList.Create;
      fSearchDirs.CaseSensitive := false;
      fSearchDirs.Duplicates := dupIgnore;
    end;

  destructor TvbMapLoader.Destroy;
    var
      i : integer;
    begin
      FreeAndNil( fSearchDirs );
      for i := 0 to fLoadedMaps.Count - 1 do
        fLoadedMaps.Objects[i].Free;
      FreeAndNil( fLoadedMaps );
      inherited;
    end;

  function TvbMapLoader.Load( const map : string; var module : TvbStdModuleDef ) : boolean;
    var
      i : integer;
      mapPath : string;
      dir : string;
    begin
      module := nil;
      result := false;

      // search for the map in our already-loaded list.
      if fLoadedMaps.Find( map, i )
        then
          begin
            module := TvbStdModuleDef( fLoadedMaps.Objects[i] );
            result := true;
          end
        else    // not loaded yet, search for it in the directory list
          with fSearchDirs do
            for i := Count - 1 downto 0 do
              begin
                dir := Strings[i];
                if not AnsiEndsText( MapFileExt, map )
                  then mapPath := dir + map + MapFileExt
                  else mapPath := dir + map;
                if FileExists( mapPath )
                  then
                    begin
                      module := LoadTranslationMap( mapPath );
                      if module <> nil
                        then
                          begin
                            assert( module.Name <> '' );
                            fLoadedMaps.AddObject( map, module );
                            if assigned( fOnMapLoaded )
                              then fOnMapLoaded( map, dir );
                            result := true;
                            break;
                          end;
                    end;
              end;
    end;

end.












