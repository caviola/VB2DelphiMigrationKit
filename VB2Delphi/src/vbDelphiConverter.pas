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

unit vbDelphiConverter;

interface

  uses
    Classes,
    Types,
    vbNodes,
    vbOptions,
    vbClasses,
    DelphiCodeWriter;

  type

    TvbConverterStatus = set of (
      csInInterface, // at the interface section
      csInImplem,    // at the implemtation section
      csInClass,     // writing code for a class module
      csInRecord     // writing the members of a record
      );

    TvbDelphiConverter =
      class
        private
          fOpts             : TvbOptions;
          fStatus           : TvbConverterStatus;
          fModule           : TvbModuleDef;
          fClassDef         : TvbClassDef;
          fFunction         : TvbAbstractFuncDef; // function currently being converted
          fWriter           : TDelphiCodeWriter;
          fResourceProvider : IvbResourceProvider;
          fMonitor          : IvbProgressMonitor;
          procedure WriteStmtBlock( stmts              : TvbStmtBlock;
                                    const downRightCmt : string  = '';
                                    semicolon          : boolean = true );
          function NeedBlock( stmts : TvbStmtBlock; const downRightCmt : string ) : boolean;
          procedure WriteArgList( args : TvbArgList; params : TvbParamList );
          procedure WriteArrayExpr( e : TvbArrayExpr );
          procedure WriteAssignStmt( s : TvbAssignStmt; semicolon : boolean = true );
          procedure WriteBaseType( const baseType : IvbType );
          procedure WriteBinaryExpr( e : TvbBinaryExpr );
          procedure WriteCallOrIndexerExpr( e : TvbCallOrIndexerExpr );
          procedure WriteCallStmt( s : TvbCallStmt; semicolon : boolean = true );
          procedure WriteCastExpr( e : TvbCastExpr );
          procedure WriteClassFields;
          procedure WriteClassFunctions( f : TvbNodeFlags );
          procedure WriteClassProperties;
          procedure WriteCloseStmt( s : TvbCloseStmt; semicolon : boolean = true );
          procedure WriteCmt( const cmt : string ); overload;
          procedure WriteCmt( const cmt : TStringDynArray ); overload;
          procedure WriteConst( cst : TvbConstDef );
          procedure WriteConstructorAndDestructor( fnConstructor, fnDestructor : TvbAbstractFuncDef );
          procedure WriteConstSection( f : TvbNodeFlags );
          procedure WriteDateStmt( s : TvbDateStmt; semicolon : boolean = true );
          procedure WriteDebugAssertStmt( s : TvbDebugAssertStmt; semicolon : boolean = true );
          procedure WriteDebugPrintStmt( s : TvbDebugPrintStmt; semicolon : boolean = true );
          procedure WriteDisclaimerComment;
          procedure WriteDllFuncSection( f : TvbNodeFlags );
          procedure WriteDoLoopStmt( s : TvbDoLoopStmt; semicolon : boolean = true );
          procedure WriteEndStmt( s : TvbEndStmt; semicolon : boolean = true );
          procedure WriteEraseStmt( s : TvbEraseStmt; semicolon : boolean = true );
          procedure WriteErrorStmt( s : TvbErrorStmt; semicolon : boolean = true );
          procedure WriteEventTypes;
          procedure WriteExitStmt( stmt : TvbStmt; e : string; semicolon : boolean );
          procedure WriteExpr( e : TvbExpr; WithParens : boolean = false );
          procedure WriteForeachStmt( s : TvbForeachStmt; semicolon : boolean = true );
          procedure WriteForStmt( s : TvbForStmt; semicolon : boolean = true );
          procedure WriteFunc( const func : string; e : TvbExpr );
          procedure WriteFunction( f : TvbAbstractFuncDef );
          procedure WriteFunctionBody( f : TvbAbstractFuncDef );
          procedure WriteFunctionParams( params : TvbParamList; WithDefaultValues : boolean );
          procedure WriteFunctionSection( f : TvbNodeFlags );
          procedure WriteGetStmt( s : TvbGetStmt; semicolon : boolean = true );
          procedure WriteGotoOrGosubStmt( s : TvbGotoOrGosubStmt; semicolon : boolean = true );
          procedure WriteIfStmt( s : TvbIfStmt; semicolon : boolean = true );
          procedure WriteInputExpr( e : TvbInputExpr );
          procedure WriteInputStmt( s : TvbInputStmt; semicolon : boolean = true );
          procedure WriteLabels( lab : TvbLabelDef );
          procedure WriteLineInputStmt( s : TvbLineInputStmt; semicolon : boolean = true );
          procedure WriteLockStmt( s : TvbLockStmt; semicolon : boolean = true );
          procedure WriteMidAssignStmt( s : TvbMidAssignStmt; semicolon : boolean = true );
          procedure WriteMidExpr( e : TvbMidExpr );
          procedure WriteName( n : TvbName );
          procedure WriteNameStmt( s : TvbNameStmt; semicolon : boolean = true );
          procedure WriteOnErrorStmt( s : TvbOnErrorStmt; semicolon : boolean = true );
          procedure WriteOnGotoOrGosubStmt( s : TvbOnGotoOrGosubStmt; semicolon : boolean = true );
          procedure WriteOpenStmt( s : TvbOpenStmt; semicolon : boolean = true );
          procedure WriteParam( p : TvbParamDef; WithDefaultValue : boolean = false );
          procedure WriteParamList( params : TvbParamList; WithDefaultValues : boolean );
          procedure WritePrintStmt( s : TvbPrintStmt; semicolon : boolean = true );
          procedure WritePutStmt( s : TvbPutStmt; semicolon : boolean = true );
          procedure WriteRaiseEventFunctions;
          procedure WriteRaiseEventStmt( s : TvbRaiseEventStmt; semicolon : boolean = true );
          procedure WriteReDimStmt( s : TvbReDimStmt; semicolon : boolean = true );
          procedure WriteResumeStmt( s : TvbResumeStmt; semicolon : boolean = true );
          procedure WriteReturnStmt( s : TvbReturnStmt; semicolon : boolean = true );
          procedure WriteSeekStmt( s : TvbSeekStmt; semicolon : boolean = true );
          procedure WriteSelectCaseStmt( s : TvbSelectCaseStmt; semicolon : boolean = true );
          procedure WriteStmt( stmt : TvbStmt; semicolon : boolean = true );
          procedure WriteStopStmt( s : TvbStopStmt; semicolon : boolean = true );
          procedure WriteTimeStmt( s : TvbTimeStmt; semicolon : boolean = true );

          procedure WriteCircleStmt( s : TvbCircleStmt; semicolon : boolean = true );
          procedure WriteLineStmt( s : TvbLineStmt; semicolon : boolean = true );
          procedure WritePSetStmt( s : TvbPSetStmt; semicolon : boolean = true );

          procedure WriteType( const typ : IvbType );
          procedure WriteTypeDefaultValue( typ : cardinal );
          procedure WriteTypeSection( f : TvbNodeFlags );
          procedure WriteUnaryExpr( e : TvbUnaryExpr );
          procedure WriteUnlockStmt( s : TvbUnlockStmt; semicolon : boolean = true );
          procedure WriteVar( v : TvbVarDef );
          procedure WriteVarSection( f : TvbNodeFlags );
          procedure WriteWidthStmt( s : TvbWidthStmt; semicolon : boolean = true );
          procedure WriteWithStmt( s : TvbWithStmt; semicolon : boolean = true );
          procedure WriteWriteStmt( s : TvbWriteStmt; semicolon : boolean = true );
          procedure WriteMemberAccessExpr( mae : TvbMemberAccessExpr );
          procedure WriteMap( const map   : string;
                              args        : TvbArgList;
                              qualifier   : TvbExpr = nil;
                              withExpr    : TvbExpr = nil;
                              value       : TvbExpr = nil;
                              writeAltMap : boolean = true );
          procedure GenerateDFM( form : TvbFormModuleDef );
          procedure WriteFormControls( form : TvbFormDef );
        public
          constructor Create( opts : TvbOptions; const resourceProvider : IvbResourceProvider; monitor : IvbProgressMonitor );
          destructor Destroy; override;
          procedure ConvertClassModule( cls : TvbClassModuleDef; OutputStream : TStream );
          procedure ConvertProject( prj : TvbProject );
          procedure ConvertStdModule( m : TvbStdModuleDef; OutputStream : TStream );
          procedure ConvertFormModule( form : TvbFormModuleDef; OutputStream : TStream );
      end;

implementation

  uses
    Math,
    vbGlobals,
    vbConsts,
    SysUtils,
    TextWriters,
    frxReader, StrUtils;

  { TvbDelphiConverter }

  procedure TvbDelphiConverter.ConvertClassModule( cls : TvbClassModuleDef; OutputStream : TStream );
    begin
      assert( cls <> nil );
      assert( OutputStream <> nil );

      if fMonitor <> nil
        then fMonitor.ConvertModuleBegin( cls );

      fModule := cls;
      fClassDef := cls.ClassDef;
      assert( fClassDef <> nil );
      with fWriter, cls do
        try
          Stream := OutputStream;
          WriteLn( 'unit ' + fModule.Name + ';' );
          WriteLn;
          WriteDisclaimerComment;
          Include( fStatus, csInInterface );
          WriteLn( 'interface' );
          WriteLn;
          Indent;
          WriteConstSection( [dfPublic] );
          WriteTypeSection( [dfPublic] );

          // Write event types and class declaration
          WriteLn( 'type' );
          Indent;
          WriteEventTypes;
          WriteLn( fClassDef.Name1 + ' =' );
          Indent;
          WriteLn( 'class' );
          Indent;
          WriteLn( 'private' );
          Indent;
          WriteClassFields;
          WriteRaiseEventFunctions;
          WriteClassFunctions( [dfPrivate] );
          Outdent;
          WriteLn( 'public' );
          Indent;
          WriteConstructorAndDestructor( fClassDef.ClassInitialize, fClassDef.ClassTerminate );
          WriteClassFunctions( [dfPublic] );
          WriteClassProperties;
          Outdent;
          Outdent;
          WriteLn( 'end;' );
          Outdent;
          Outdent;
          WriteLn;
          WriteDllFuncSection( [dfPublic] );
          if length( DownCmts ) > 0
            then
              begin
                WriteCmt( DownCmts );
                WriteLn;
              end;
          Outdent;
          Exclude( fStatus, csInInterface );
          Include( fStatus, csInImplem );
          WriteLn( 'implementation' );
          WriteLn;
          Indent;
          WriteConstSection( [dfPrivate] );
          WriteTypeSection( [dfPrivate] );
          WriteDllFuncSection( [dfPrivate] );
          WriteConstructorAndDestructor( fClassDef.ClassInitialize, fClassDef.ClassTerminate );
          WriteClassFunctions( [dfPublic, dfPrivate] );
          WriteRaiseEventFunctions;
          Exclude( fStatus, csInImplem );
          Outdent;
          WriteLn( 'end.' );

          if fMonitor <> nil
            then fMonitor.ConvertModuleEnd( cls );
        finally
          //
        end;
    end;

  procedure TvbDelphiConverter.ConvertStdModule( m : TvbStdModuleDef; OutputStream : TStream );
    begin
      assert( m <> nil );
      assert( OutputStream <> nil );

      if fMonitor <> nil
        then fMonitor.ConvertModuleBegin( m );

      fModule := m;
      with fWriter do
        try
          Stream := OutputStream;
          WriteLn( 'unit ' + fModule.Name + ';' );
          WriteLn;
          WriteDisclaimerComment;
          Include( fStatus, csInInterface );
          WriteLn( 'interface' );
          WriteLn;
          Indent;
          WriteConstSection( [dfPublic] );
          WriteTypeSection( [dfPublic] );
          WriteVarSection( [dfPublic ] );
          WriteFunctionSection( [dfPublic] );
          WriteDllFuncSection( [dfPublic] );
          if length( m.DownCmts ) > 0
            then
              begin
                WriteCmt( m.DownCmts );
                WriteLn;
              end;
          Exclude( fStatus, csInInterface );
          Include( fStatus, csInImplem );
          Outdent;
          WriteLn( 'implementation' );
          WriteLn;
          Indent;
          WriteConstSection( [dfPrivate] );
          WriteTypeSection( [dfPrivate] );
          WriteVarSection( [dfPrivate ] );
          WriteDllFuncSection( [dfPrivate] );
          WriteFunctionSection( [dfPublic, dfPrivate] );
          Exclude( fStatus, csInImplem );
          Outdent;
          WriteLn( 'end.' );

          if fMonitor <> nil
            then fMonitor.ConvertModuleEnd( m );

        finally
          //
        end;
    end;

  constructor TvbDelphiConverter.Create( opts : TvbOptions; const resourceProvider : IvbResourceProvider; monitor : IvbProgressMonitor );
    begin
      fOpts := opts;
      assert( resourceProvider <> nil );
      fResourceProvider := resourceProvider;
      fWriter := TDelphiCodeWriter.Create;
      fMonitor := monitor;
    end;

  procedure TvbDelphiConverter.WriteAssignStmt( s : TvbAssignStmt; semicolon : boolean );
    var
      args   : TvbArgList;
      params : TvbParamList;
      prop   : TvbExpr;
    begin
      with fWriter, s do
        begin
          assert( Value <> nil );
          assert( LHS <> nil );

          WriteCmt( TopCmts );

          // Check if the LHS is a property setter that we have to convert
          // to a function call.
          if ( LHS.Def <> nil ) and ( mapAsFunctionCall in LHS.Def.NodeFlags  )
            then
              begin
                // Get the property's parameter definition list.
                assert( LHS.Def is TvbFuncDef );
                params := TvbFuncDef( LHS.Def ).Params;
                // If the LHS has arguments(a CALL_OR_INDEXER_EXPR expr), also
                // get those arguments.
                if LHS.NodeKind = CALL_OR_INDEXER_EXPR
                  then
                    begin
                      args := TvbCallOrIndexerExpr( LHS ).Args;
                      prop := TvbCallOrIndexerExpr( LHS ).FuncOrVar;
                    end
                  else
                    begin
                      args := nil;
                      prop := LHS;
                    end;
                // Now write a function call adding as last parameter the
                // value of this assignment.
                WriteExpr( prop );
                Write( '( ');
                if args <> nil
                  then
                    begin
                      WriteArgList( args, params );
                      Write( ', ' );
                    end;
                WriteExpr( Value );
                Write( ' )');
              end
            else
              begin
                WriteExpr( LHS );
                write( ' := ' );
                WriteExpr( Value );
              end;
              
          if semicolon
            then Write( ';' );
          WriteCmt( TopRightCmt );
        end;
    end;

  procedure TvbDelphiConverter.WriteGotoOrGosubStmt( s : TvbGotoOrGosubStmt; semicolon : boolean );
    begin
      WriteCmt( s.TopCmts );
      with fWriter, s do
        case GotoOrGosub of
          ggGoTo :
            begin
              Write( 'goto ' + LabelName.Name );
              if semicolon
                then Write( ';' );
              WriteCmt( TopRightCmt );
            end;
          //UNIMPLEMENTED
          ggGoSub : WriteLineCmt( 'GOSUB' );
        end;
    end;

  procedure TvbDelphiConverter.WriteClassFunctions( f : TvbNodeFlags );
    begin
      Include( fStatus, csInClass );
      WriteFunctionSection( f );
      Exclude( fStatus, csInClass );
    end;

  procedure TvbDelphiConverter.WriteCloseStmt( s : TvbCloseStmt; semicolon : boolean );
    begin
      with fWriter, s do
        begin
          WriteCmt( TopCmts );
          if FileNum <> nil
            then Write( 'CloseFile( ' + GetOpenStmtFileName( FileNum ) + ' )' )
            else Write( 'CloseAllFiles{UNSUPPORTED}' );
          if semicolon
            then Write( ';' );
          WriteCmt( TopRightCmt );
        end;
    end;

  procedure TvbDelphiConverter.WriteOnGotoOrGosubStmt( s : TvbOnGotoOrGosubStmt; semicolon : boolean );
    begin
      with fWriter, s do
        begin
          WriteCmt( TopCmts );
          //UNIMPLEMENTED
          WriteLineCmt( 'ON GOSUB/GOTO' );
        end;
    end;

  procedure TvbDelphiConverter.WriteConst( cst : TvbConstDef );
    begin
      with fWriter do
        begin
          WriteCmt( cst.TopCmts );
          Write( cst.Name );
          if cst.ConstType <> nil
            then
              begin
                Write( ' : ' );
                WriteType( cst.ConstType );
              end;
          Write( ' = ' );
          WriteExpr( cst.Value );
          Write( '; ' );
          WriteCmt( cst.TopRightCmt );
        end;
    end;

  procedure TvbDelphiConverter.WriteConstSection( f : TvbNodeFlags );
    var
      i, j : integer;
      cst  : TvbConstDef;
      ec   : TvbConstDef;
      enum : TvbEnumDef;
      any  : boolean; // true if we wrote any const already
    begin
      any := false;
      with fWriter do
        begin
          for i := 0 to fModule.ConstCount - 1 do
            begin
              cst := fModule.Consts[i];
              if cst.NodeFlags * f <> []
                then
                  begin
                    // Write the section header if we haven't written it
                    // already.
                    if not any
                      then
                        begin
                          WriteLn( 'const' );
                          Indent;
                          any := true;
                        end;
                    WriteConst( cst );
                  end;
            end;
          if any
            then
              begin
                Outdent;
                WriteLn;
              end;

          // Write the enumerated constants.
          if fOpts.optEnumToConst
            then
              for i := 0 to fModule.EnumCount - 1 do
                begin
                  enum := fModule.Enum[i];
                  with enum do
                    if ( NodeFlags * f <> [] ) and
                       ( EnumCount > 0 ) and
                       not ( enum_MissingValues in NodeFlags )
                      then
                        begin
                          WriteLn( 'const' );
                          Indent;
                          WriteLineCmt( ' ' + Name + ' constants' );
                          for j := 0 to EnumCount - 1 do
                            begin
                              assert( Enum[j].NodeKind = CONST_DEF );
                              ec := TvbConstDef( Enum[j] );
                              WriteCmt( ec.TopCmts );
                              Write( ec.Name );
                              Write( ' = ' );
                              WriteExpr( ec.Value );
                              Write( '; ' );
                              WriteCmt( ec.TopRightCmt );
                            end;
                          WriteCmt( DownCmts );
                          Outdent;
                          WriteLn;
                        end;
                end;
        end;
    end;

  procedure TvbDelphiConverter.WriteDateStmt( s : TvbDateStmt; semicolon : boolean );
    begin
      with fWriter, s do
        begin
          WriteCmt( TopCmts );
          Write( 'SetDate( ' );
          WriteExpr( NewDate );
          Write( ' )' );
          if semicolon
            then Write( ';' );
          WriteCmt( TopRightCmt );
        end;
    end;

  procedure TvbDelphiConverter.WriteDebugAssertStmt( s : TvbDebugAssertStmt; semicolon : boolean );
    begin
      with fWriter, s do
        begin
          WriteCmt( TopCmts );
          Write( 'assert( ' );
          WriteExpr( BoolExpr );
          Write( ' )' );
          if semicolon
            then Write( ';' );
          WriteCmt( TopRightCmt );
        end;
    end;

  procedure TvbDelphiConverter.WriteDebugPrintStmt( s : TvbDebugPrintStmt; semicolon : boolean );
    begin
      WriteCmt( s.TopCmts );
      fWriter.WriteLineCmt( 'Debug.Print' );
    end;

  procedure TvbDelphiConverter.WriteDisclaimerComment;
    begin
      with fWriter do
        begin
          WriteLineCmt( ' DISCLAIMER!' );
          WriteLineCmt( ' ==================================================================' );
          WriteLineCmt( ' This source code has been automatically generated.' );
          WriteLineCmt( ' It still requires extensive manual fixing, debugging and testing.' );
          WriteLineCmt( ' IT IS NOT GUARANTEED TO COMPILE' );
          WriteLineCmt( ' AND MUCH LESS TO BEHAVE EXACTLY LIKE THE ORIGINAL SOURCE CODE.' );
          WriteLineCmt( ' Use it at your sole risk and discretion.' );
          WriteLineCmt( '' );
          WriteLn;
        end;
    end;

  procedure TvbDelphiConverter.WriteDllFuncSection( f : TvbNodeFlags );
    var
      i       : integer;
      func    : TvbDllFuncDef;
      any,
      isProc  : boolean;
    begin
      any := false;
      with fWriter do
        begin
          for i := 0 to fModule.DllFuncCount - 1 do
            begin
              func := fModule.DllFuncs[i];
              with func do
                if NodeFlags * f <> []
                  then
                    begin
                      any := true;
                      isProc := ReturnType = nil;
                      WriteCmt( TopCmts );
                      if isProc
                        then Write( 'procedure ' )
                        else Write( 'function ' );
                      Write( Name );
                      WriteFunctionParams( Params, true );
                      if not isProc
                        then
                          begin
                            Write( ' : ' );
                            WriteType( ReturnType );
                          end;
                      Write( '; stdcall; external ' + '''' + DllName + '''' );
                      if Alias <> ''
                        then Write( ' name ' + '''' + Alias + ''';' )
                        else Write( ';' );
                      WriteCmt( TopRightCmt );
                    end;
            end;
          if any
            then WriteLn;
        end;
    end;

  procedure TvbDelphiConverter.WriteEraseStmt( s : TvbEraseStmt; semicolon : boolean );
    begin
      with fWriter, s do
        begin
          WriteCmt( TopCmts );
          Write( 'setlength( ' );
          WriteExpr( ArrayVar );
          Write( ', 0 )' );
          if semicolon
            then Write( ';' );
          WriteCmt( TopRightCmt );
        end;
    end;

  procedure TvbDelphiConverter.WriteErrorStmt( s : TvbErrorStmt; semicolon : boolean );
    begin
      with fWriter, s do
        begin
          WriteCmt( TopCmts );
          Write( 'RunError( ' );
          WriteExpr( ErrorNumber );
          Write( ' )' );
          if semicolon
            then Write( ';' );
          WriteCmt( TopRightCmt );
        end;
    end;

  procedure TvbDelphiConverter.WriteExitStmt( stmt : TvbStmt; e : string; semicolon : boolean );
    begin
      with fWriter, stmt do
        begin
          WriteCmt( TopCmts );
          Write( e );
          if semicolon
            then Write( ';' );
          WriteCmt( TopRightCmt );
        end;
    end;

  procedure TvbDelphiConverter.WriteArgList( args : TvbArgList; params : TvbParamList );
    var
      arg           : TvbExpr;
      argCount      : integer;
      i             : integer;
      lastArgIdx    : integer;
      namedArgCount : integer;
      param         : TvbParamDef;
      paramArrayIdx : integer;
      paramCount    : integer;
      usedArgCount  : integer;

      procedure WriteArg( arg : TvbExpr; param : TvbParamDef = nil );
        begin
          with fWriter do
            // if the argument has been omitted write the corresponding
            // parameter's default value or Unassigned.
            if arg = nil
              then
                begin
                  if ( param <> nil ) and ( param.DefaultValue <> nil )
                    then WriteExpr( param.DefaultValue )
                    else Write( 'Unassigned' );
                end
              else
            // if we don't have the parameter definition simply write the argument.
            if param = nil
              then WriteExpr( arg )
              else
                begin
                  if _pfPChar in param.NodeFlags
                    then
                      if arg.NodeKind <> STRING_LIT_EXPR
                        then WriteFunc( 'pchar', arg )
                        else WriteExpr( arg )
                    else
                  // If the corresponding parameter is 'ByVal..As Any'
                  // translate the argument as a POINTER typecast.
                  // We do this except for ADDRESSOF_EXPR expressions because
                  // they are translated to '@' which returns a pointer.
                  if _pfAnyByVal in param.NodeFlags
                    then
                      if IsStringType( arg.NodeType )
                        then WriteFunc( 'pchar', arg )
                        else
                      if arg.NodeKind <> ADDRESSOF_EXPR
                        then WriteFunc( 'pointer', arg )
                        else WriteExpr( arg )
                    else WriteExpr( arg );
                end;
        end;

    begin
      assert( args <> nil );

      argCount   := args.ArgCount;
      lastArgIdx := argCount - 1; 

      with fWriter do
        // Handle the case when no parameter definitions are available.
        // Simply write all arguments(positional and named) in the same
        // order that they are provided.
        if params = nil
          then
            for i := 0 to lastArgIdx do
              begin
                WriteArg( args[i] );
                if i < lastArgIdx
                  then Write( ', ' );
              end
          else
            begin
              paramCount    := params.ParamCount;
              paramArrayIdx := params.ParamArrayIndex;

              if args.FirstNamedArgIndex <> -1 // do we have named arguments?
                then
                  begin
                    // First write all positional arguments.
                    for i := 0 to args.FirstNamedArgIndex - 1 do
                      begin
                        WriteArg( args[i], params[i] );
                        // We write the comma without further checks because we
                        // know the user provided at least one named argument. 
                        Write( ', ' );
                      end;

                    i             := args.FirstNamedArgIndex;
                    usedArgCount  := 0;
                    namedArgCount := args.ArgCount - args.FirstNamedArgIndex;

                    // Now we'll write the named arguments provided.
                    // Note that since args.FirstNamedArgIndex <> -1 there should
                    // be at least one named argument to be written by this
                    // loop, although this assumption might be wrong if the
                    // the name of the provided argument doesn't exisit in
                    // the parameter list, in which case the source VB code is
                    // incorrect.
                    while i < paramCount do
                      begin
                        // Take the next parameter from the list.
                        param := params[i];

                        // Check if an argument with the same name of the
                        // parameter has been supplied.
                        arg := args.NamedArg[param.Name];
                        if arg <> nil
                          then
                            begin
                              // The argument is present. Write its value
                              // and increment the used named argument count.
                              WriteArg( arg, param );
                              inc( usedArgCount );
                            end
                          else
                            // The argument has been omitted. Write the
                            // corresponding parameter's default value.
                            WriteArg( param.DefaultValue );

                        if usedArgCount < namedArgCount
                          then Write( ', ' )
                          else break;

                        inc( i );
                      end;
                  end
                else
              if paramArrayIdx >= 0 // do we have a ParamArray?
                then
                  begin
                    // Starting with the first argument, write each one until
                    // we either (1) have written them all or (2) reach the
                    // position of the ParamArray argument.
                    for i := 0 to min( argCount, paramArrayIdx ) - 1 do
                      begin
                        WriteArg( args[i], params[i] );
                        Write( ', ' );
                      end;

                    // If we wrote all arguments this means the user omitted the
                    // ParamArray and so we write an empty variant open array.
                    if i = argCount
                      then Write( '[]' )
                      else                     
                        // We reached the ParamArray's position and so we'll
                        // write all remaining arguments enclosed in [].
                        begin
                          Write( '[' );
                          for i := i to lastArgIdx do
                            begin
                              WriteArg( args[i] );
                              if i < lastArgIdx
                                then Write( ', ' );
                            end;
                          Write( ']' );
                        end;
                  end
                else
                  // Handle the case when there is no ParamArray.
                  // Simply write all arguments from first to last.
                  for i := 0 to lastArgIdx do
                    begin
                      WriteArg( args[i], params[i] );
                      if i < lastArgIdx
                        then Write( ', ' );
                    end;
            end;
    end;

  function GetFunctionParams( const func : TvbDef ) : TvbParamList;
    begin
      result := nil;
      if func <> nil
        then
          case func.NodeKind of
            FUNC_DEF,
            DLL_FUNC_DEF : result := TvbFuncDef( func ).Params;
            PROPERTY_DEF : result := TvbPropertyDef( func ).Params;
          end;
    end;

  procedure TvbDelphiConverter.WriteFunc( const func : string; e : TvbExpr );
    begin
      fWriter.Write( func );
      WriteExpr( e, true );
    end;

  procedure TvbDelphiConverter.WriteExpr( e : TvbExpr; WithParens : boolean );
    var
      dae     : TvbDictAccessExpr absolute e;
      inte    : TvbIntLit absolute e;
      date    : TvbDateLit absolute e;
      addre   : TvbAddressOfExpr absolute e;
      typeof  : TvbTypeOfExpr absolute e;
      strExpr : TvbStringExpr absolute e;
      ne      : TvbNameExpr absolute e;
    begin
      assert( e <> nil );
      with fWriter do
        begin
          if e.Parenthesized or WithParens
            then Write( '( ' );
          case e.NodeKind of
            BINARY_EXPR : WriteBinaryExpr( TvbBinaryExpr( e ) );
            UNARY_EXPR : WriteUnaryExpr( TvbUnaryExpr( e ) );
            TYPEOF_EXPR :
              begin
                WriteExpr( typeof.ObjectExpr );
                Write( ' is ' );
                WriteName( typeof.Name );
              end;
            FUNC_RESULT_EXPR : Write( 'result' );
            NAME_EXPR :
              begin
                if ne.Def <> nil
                  then WriteMap( ne.Def.Name1, nil, nil )
                  else fWriter.Write( ne.Name );
              end;
            MID_EXPR : WriteMidExpr( TvbMidExpr( e ) );
            INT_EXPR : WriteFunc( 'Floor', e.Op1 );
            FIX_EXPR : WriteFunc( 'Ceil', e.Op1 );
            ABS_EXPR : WriteFunc( 'Abs', e.Op1 );
            SGN_EXPR : WriteFunc( 'Sign', e.Op1 );
            STRING_EXPR :
              with strExpr do
                begin
                  Write( 'StringOfChar( ' );
                  WriteExpr( Char );
                  Write( ', ' );
                  WriteExpr( Count );
                  Write( ' )' );
                end;
            ARRAY_EXPR : WriteArrayExpr( TvbArrayExpr( e ) );
            LEN_EXPR :
              begin
                assert( e.Op1 <> nil );
                if ( e.Op1.NodeType = vbVariantType ) or IsStringType( e.Op1.NodeType ) or ( e.Op1.NodeType = nil )
                  then WriteFunc( 'length', e.Op1 )
                  else WriteFunc( 'sizeof', e.Op1 );
              end;
            LENB_EXPR :
              begin
                assert( e.Op1 <> nil );
                if ( e.Op1.NodeType = vbVariantType ) or IsStringType( e.Op1.NodeType ) or ( e.Op1.NodeType = nil )
                  then
                    begin
                      WriteFunc( 'length', e.Op1 );
                      if not fOpts.optUseAnsiString
                        then Write( ' * 2' );
                    end
                  else WriteFunc( 'sizeof', e.Op1 );
              end;
            DATE_EXPR : Write( 'variant( Date )' );
            CALL_OR_INDEXER_EXPR : WriteCallOrIndexerExpr( TvbCallOrIndexerExpr( e ) );
            DICT_ACCESS_EXPR :
              begin
                //TODO
                if dae.Target = nil
                  then Write('DEFAULTPROPERTY')
                  else WriteExpr( dae.Target );
                Write( '[''' + dae.Name + ''']' );
              end;
            MEMBER_ACCESS_EXPR : WriteMemberAccessExpr( TvbMemberAccessExpr( e ) );
            ME_EXPR : Write( 'self' );
            INT_LIT_EXPR :
              if int_Hex in e.NodeFlags
                then Write( '$' + IntToHex( inte.Value, 8 ) )
                else Write( inte.Value );
            STRING_LIT_EXPR : WriteStrConst( TvbStringLit( e ).Value );
            DATE_LIT_EXPR :
              begin
                Write( date.Value );
                WriteBlockCmt( DateTimeToStr( date.Value ) );
              end;
            FLOAT_LIT_EXPR : Write( TvbFloatLit( e ).Value );
            NOTHING_EXPR   : Write( 'nil' );
            BOOL_LIT_EXPR :
              if TvbBoolLit( e ).Value
                then Write( 'true' )
                else Write( 'false' );
            ADDRESSOF_EXPR :
              begin
                assert( addre.FuncName <> nil );
                Write( '@' );
                WriteName( addre.FuncName );
              end;
            LBOUND_EXPR :
              begin
                Write( 'low( ' );
                WriteExpr( e.Op1 );
                if e.Op2 <> nil
                  then
                    begin
                      Write( '[' );
                      WriteExpr( e.Op2 );
                      Write( ']' );
                    end;
                Write( ' )' );
              end;
            UBOUND_EXPR :
              begin
                Write( 'high( ' );
                WriteExpr( e.Op1 );
                if e.Op2 <> nil
                  then
                    begin
                      Write( '[' );
                      WriteExpr( e.Op2 );
                      Write( ']' );
                    end;
                Write( ' )' );
              end;
            CAST_EXPR  : WriteCastExpr( TvbCastExpr( e ) );
            INPUT_EXPR : WriteInputExpr( TvbInputExpr( e ) );
            SEEK_EXPR :
              begin
                assert( TvbSeekExpr( e ).FileNum <> nil );
                Write( 'FilePos( ' + GetOpenStmtFileName( TvbSeekExpr( e ).FileNum ) + ' )' );
              end;
            DOEVENTS_EXPR : Write('Application.ProcessMessages');
            else WriteBlockCmt( 'UNIMPLEMENTED', csBracket );
          end;
          if e.Parenthesized or WithParens
            then Write( ' )' );
        end;
    end;

  procedure TvbDelphiConverter.WriteForeachStmt( s : TvbForeachStmt; semicolon : boolean );
    begin
      with fWriter, s do
        begin
          WriteCmt( TopCmts );
          //UNIMPLEMENTED
          WriteLineCmt( 'FOREACH' );
        end;
    end;

  procedure TvbDelphiConverter.WriteForStmt( s : TvbForStmt; semicolon : boolean );
    var
      stmt : TvbStmt;
    begin
      with fWriter, s do
        begin
          WriteCmt( TopCmts );
          if for_StepN in NodeFlags
            then
              begin
                // Write expression to initialize counter.
                WriteExpr( Counter );
                Write( ' := ' );
                WriteExpr( InitValue );
                WriteLn( ';' );

                Write( 'while ' );
                WriteExpr( Counter );
                if for_StepMinusN in NodeFlags
                  then Write( ' >= ' )
                  else Write( ' <= ' );
                WriteExpr( FinalValue );
                Write( ' do' );
                WriteCmt( TopRightCmt );                

                Indent;
                WriteBegin;
                stmt := StmtBlock.FirstStmt;
                while stmt <> nil do
                  begin
                    WriteStmt( stmt );
                    stmt := stmt.NextStmt;
                  end;
                WriteCmt( StmtBlock.DownCmts );

                // Write the expression to increment the Counter by Step.
                Write( 'inc( ' );
                WriteExpr( Counter );
                Write( ', ' );
                WriteExpr( StepValue );
                WriteLn( ' );' );

                WriteEnd( semicolon, DownRightCmt );
                Outdent;
              end
            else
              begin
                Write( 'for ');
                WriteExpr( Counter );
                Write( ' := ' );
                WriteExpr( InitValue );
                if for_StepMinus1 in NodeFlags
                  then Write( ' downto ' )
                  else Write( ' to ' );
                WriteExpr( FinalValue );
                Write( ' do' );
                if NeedBlock( StmtBlock, DownRightCmt )
                  then
                    begin
                      WriteCmt( TopRightCmt );
                      WriteStmtBlock( StmtBlock, DownRightCmt, semicolon );
                    end
                  else
                    begin
                      Write( ';' );
                      WriteCmt( TopRightCmt );
                    end;
              end;
        end;
    end;

  procedure TvbDelphiConverter.WriteFunction( f : TvbAbstractFuncDef );
    var
      isProc, inImpl, inIntf, isClass, withCmts : boolean;
    begin
      isProc  := f.ReturnType = nil;
      inImpl  := csInImplem in fStatus;
      inIntf  := not inImpl;
      isClass := csInClass in fStatus;

      withCmts :=
        // If this is a class, write comments only in the implementation section.
        ( isClass and inImpl ) or
        // If not a class, write comments in the interface for public functions
        // and in the implementation for private functions.
        ( not isClass and ( ( dfPrivate in f.NodeFlags ) or inIntf ) );

      with fWriter, f do
        begin
          if withCmts
            then WriteCmt( f.TopCmts );
            
          if isProc
            then Write( 'procedure ' )
            else Write( 'function ' );

          // In the implementation section prefix methods with the class name.
          if inImpl and isClass
            then Write( fClassDef.Name1 + '.' );

          Write( Name1 );
          WriteFunctionParams( Params, inIntf or ( dfPrivate in f.NodeFlags ) );

          // Write function return type.
          if not isProc
            then
              begin
                Write( ' : ');
                WriteType( ReturnType );
              end;
          Write( ';' );
          if withCmts
            then WriteCmt( TopRightCmt )
            else WriteLn;

          // If we are in the implementation section write local consts, vars
          // and statement list.
          if inImpl
            then
              begin
                WriteFunctionBody( f );
                WriteLn;
              end;
        end;
    end;

  procedure TvbDelphiConverter.WriteFunctionSection( f : TvbNodeFlags );
    var
      i    : integer;
      func : TvbAbstractFuncDef;
      p    : TvbPropertyDef;
      any  : boolean;
    begin
      any := false;
      with fWriter, fModule do
        begin
          for i := 0 to MethodCount - 1 do
            begin
              func := Method[i];
              if func.NodeFlags * f <> []
                then
                  begin
                    WriteFunction( func );
                    any := true;
                  end;
            end;

          for i := 0 to PropCount - 1 do
            begin
              p := Prop[i];
              if ( p.GetFn <> nil ) and
                 ( p.GetFn.NodeFlags * f <> [] )
                then WriteFunction( p.GetFn );
              if ( p.SetFn <> nil ) and
                 ( p.SetFn.NodeFlags * f <> [] )
                then WriteFunction( p.SetFn )
                else
              if ( p.LetFn <> nil ) and
                 ( p.LetFn.NodeFlags * f <> [] )
                then WriteFunction( p.LetFn );
            end;

          if ( csInInterface in fStatus ) and ( not ( csInClass in fStatus ) ) and any
            then WriteLn;
        end;
    end;

  procedure TvbDelphiConverter.WriteGetStmt( s : TvbGetStmt; semicolon : boolean );
    var
      fn : string;
    begin
      with fWriter, s do
        begin
          assert( FileNum <> nil );
          assert( VarName <> nil );
          fn := GetOpenStmtFileName( FileNum );
          WriteCmt( TopCmts );
          if RecNum <> nil
            then
              begin
                Write( 'Seek( ' + fn + ', ' );
                WriteExpr( RecNum );
                WriteLn( ' );' );
              end;
          Write( 'Read( ' + fn + ', ' );
          WriteExpr( VarName );
          Write( ' );' );
          WriteCmt( TopRightCmt );
        end;
    end;

  procedure TvbDelphiConverter.WriteIfStmt( s : TvbIfStmt; semicolon : boolean );
    var
      i              : integer;
      elseNeedsBlock : boolean;

    // Write an if/then statemet without the "else" part.
    procedure WriteIfThen( ifs : TvbIfStmt; semicolon : boolean = true );
      begin
        with fWriter, ifs do
          begin
            WriteCmt( TopCmts );
            Write( 'if ' );
            WriteExpr( Expr );
            WriteCmt( TopRightCmt );
            Indent;
            Write( 'then' );
            if NeedBlock( ThenBlock, DownRightCmt )
              then
                begin
                  WriteLn;
                  WriteStmtBlock(  ThenBlock, DownRightCmt, semicolon );
                end
              else WriteLn( semicolon );
            Outdent;
          end;
      end;

    begin
      with fWriter, s do
        begin
          // This checks whether we have to enclose the else-part statements
          // between 'begin..end'.
          elseNeedsBlock := ( ElseBlock <> nil ) and
            NeedBlock( ElseBlock, ElseBlock.DownRightCmt );

          WriteIfThen( s, ( ElseIfCount = 0 ) and not elseNeedsBlock );
          
          for i := 0 to ElseIfCount - 1 do
            begin
              Indent;
              WriteLn( 'else' );
              Outdent;
              WriteIfThen( ElseIf[i],
                ( i = ElseIfCount - 1 ) and not elseNeedsBlock and semicolon );
            end;
          if NeedBlock( ElseBlock, DownRightCmt )
            then
              begin
                Indent;
                WriteLn( 'else' );
                WriteStmtBlock( ElseBlock, DownRightCmt, semicolon );
                Outdent;
              end;
        end;
    end;

  procedure TvbDelphiConverter.WriteInputStmt( s : TvbInputStmt; semicolon : boolean );
    var
      i : integer;
    begin
      with fWriter, s do
        begin
          WriteCmt( TopCmts );
          Write( 'Read( ' + GetOpenStmtFileName( FileNum ) );
          for i := 0 to VarCount - 1 do
            begin
              Write( ', ' );
              WriteExpr( Vars[i] );
            end;
          Write( ' )' );
          if semicolon
            then Write( ';' );
          WriteCmt( TopRightCmt );
        end;
    end;

  procedure TvbDelphiConverter.WriteLineInputStmt( s : TvbLineInputStmt; semicolon : boolean );
    begin
      with fWriter, s do
        begin
          WriteCmt( TopCmts );
          Write( 'ReadLn( ' + GetOpenStmtFileName( FileNum ) );
          Write( ', ' );
          WriteExpr( VarName );
          Write( ' )' );
          if semicolon
            then Write( ';' );
          WriteCmt( TopRightCmt );
        end;
    end;

  procedure TvbDelphiConverter.WriteLockStmt( s : TvbLockStmt; semicolon : boolean );
    begin
      with fWriter, s do
        begin
          WriteCmt( TopCmts );
          //UNIMPLEMENTED
          WriteLineCmt( 'LOCK' );
        end;
    end;

  procedure TvbDelphiConverter.WriteDoLoopStmt( s : TvbDoLoopStmt; semicolon : boolean );
    var
      stmt : TvbStmt;
    begin
      WriteCmt( s.TopCmts );
      with fWriter, s do
        if TestAtBegin
          then
            begin
              Write( 'while ' );
              if DoLoop = doUntil
                then Write( 'not ( ' );
              WriteExpr( Condition );
              if DoLoop = doUntil
                then Write( ' )' );
              Write( ' do');
              if NeedBlock( StmtBlock, DownRightCmt )
                then
                  begin
                    WriteCmt( TopRightCmt );
                    WriteStmtBlock( StmtBlock, DownRightCmt, semicolon );
                  end
                else
                  begin
                    Write( ';' );
                    WriteCmt( TopRightCmt );
                  end;
            end
          else
            begin
              Write( 'repeat' );
              WriteCmt( TopRightCmt );
              stmt := StmtBlock.FirstStmt;
              if stmt <> nil
                then
                  begin
                    Indent;
                    repeat
                      assert( stmt is TvbStmt );
                      WriteStmt( stmt );
                      stmt := stmt.NextStmt;
                    until stmt = nil;
                    WriteCmt( StmtBlock.DownCmts );
                    Outdent;
                  end;
              Write( 'until ' );
              if DoLoop = doWhile
                then Write( 'not ( ' );
              WriteExpr( Condition );
              if DoLoop = doWhile
                then Write( ' )' );
              if semicolon
                then Write( ';' );
              WriteCmt( DownRightCmt );
            end;
    end;

  procedure TvbDelphiConverter.WriteCallStmt( s : TvbCallStmt; semicolon : boolean );
    begin
      WriteCmt( s.TopCmts );
      WriteExpr( s.Method );
      if semicolon
        then fWriter.Write( ';' );
      WriteCmt( s.TopRightCmt );
    end;

  procedure TvbDelphiConverter.WriteMidAssignStmt( s : TvbMidAssignStmt; semicolon : boolean );
    begin
      with fWriter, s do
        begin
          WriteCmt( TopCmts );
          WriteExpr( VarName );
          Write( ' := ' );
          Write( 'StuffString( ' );
          WriteExpr( VarName );
          Write( ', ' );
          WriteExpr( StartPos );
          Write( ', ' );
          if Length = nil
            then
              begin
                Write( 'length( ' );
                WriteExpr( Replacement );
                Write( ' )' );
              end
            else WriteExpr( Length );
          Write( ', ' );
          WriteExpr( Replacement );
          Write( ' )' );
          if semicolon
            then fWriter.Write( ';' );
          WriteCmt( s.TopRightCmt );
        end;
    end;

  procedure TvbDelphiConverter.WriteName( N : TvbName );
    var
      QN : TvbQualifiedName absolute N;
    begin
      with fWriter do
        case N.NodeKind of
          SIMPLE_NAME :
            // If we know the declaration that this name refers to
            // write the mapping for that declaration.
            if N.Def <> nil
              then Write( N.Def.Name1 )
              else Write( N.Name );
          QUALIFIED_NAME :
            begin
              assert( QN.Qualifier <> nil );
              if ( QN.Def = nil ) or not ( map_Full in QN.Def.NodeFlags )
                then
                  begin
                    WriteName( QN.Qualifier );
                    Write( '.' );
                  end;
              if QN.Def <> nil
                then Write( QN.Def.Name1 )
                else Write( QN.Name );
            end;
        end;
    end;

  procedure TvbDelphiConverter.WriteNameStmt( s : TvbNameStmt; semicolon : boolean );
    begin
      with fWriter, s do
        begin
          WriteCmt( TopCmts );
          Write( 'RenameFile( ' );
          WriteExpr( OldPath );
          Write( ', ' );
          WriteExpr( NewPath );
          Write( ' )' );
          if semicolon
            then Write( ';' );
          WriteCmt( TopRightCmt );
        end;
    end;

  procedure TvbDelphiConverter.WriteOnErrorStmt( s : TvbOnErrorStmt; semicolon : boolean );
    begin
      WriteCmt( s.TopCmts );
      fWriter.WriteLineCmt( 'ON ERROR' );
    end;

  procedure TvbDelphiConverter.WriteOpenStmt( s : TvbOpenStmt; semicolon : boolean );
    var
      fname : string;
    begin
      with fWriter, s do
        begin
          WriteCmt( TopCmts );
          fname := GetOpenStmtFileName( FileNum );
          Write( 'Assign( ' + fname + ', ' );
          WriteExpr( FilePath );
          Write( ' ); ' );
          WriteCmt( TopRightCmt );
          case Mode of
            omRandom,
            omBinary,
            omInput  : WriteLn( 'Reset( ' + fname + ' );' );
            omAppend : WriteLn( 'Append( ' + fname + ' );' );
            omOutput : WriteLn( 'Rewrite( ' + fname + ' );' );
          end;
        end;
    end;

  procedure TvbDelphiConverter.WritePrintStmt( s : TvbPrintStmt; semicolon : boolean );
    var
      i : integer;
      e : TvbPrintStmtExpr;
    begin
      with fWriter, s do
        begin
          WriteCmt( TopCmts );
          if LineBreak
            then Write( 'WriteLn( ' )
            else Write( 'Write( ' );
          Write( GetOpenStmtFileName( FileNum ) );
          for i := 0 to ExprCount - 1 do
            begin
              e := Expr[i];
              assert( e <> nil );
              case e.ArgKind of
                pakSpc :
                  begin
                    Write( ', StringOfChar( #32, ' );
                    WriteExpr( e.Expr );
                    Write( ' )' );
                  end;
                pakTab :
                  begin
                    //REVIEW: in VB what is actually written IS NOT a tab char.
                    Write( ', #9' );
                  end;
                else
                  begin
                    Write( ', ' );
                    WriteExpr( e.Expr );
                  end;
              end;
            end;
          Write( ' )' );
          if semicolon
            then Write( ';' );
          WriteCmt( TopRightCmt );
        end;
    end;

  // Declares the class private fields.
  // These are the class module's private variables and we also declare
  // a private field to hold the value of public variables which
  // are converted to properties.
  procedure TvbDelphiConverter.WriteClassFields;
    var
      i  : integer;
      v  : TvbVarDef;
      ev : TvbEventDef;
    begin
      with fWriter, fModule do
        begin
          // Write private variables and fields to store the value
          // of public variables which are declared as properties.
          for i := 0 to VarCount - 1 do
            begin
              v := Vars[i];
              if dfPrivate in v.NodeFlags
                then WriteVar( v )
                else
              if dfPublic in v.NodeFlags
                then
                  begin
                    Write( Format( fmtDelphiClassField, [v.Name] ) );
                    Write( ' : ' );
                    WriteType( v.VarType );
                    WriteLn( ';' );
                  end;
            end;

          // Declare fields to store the user-provided event handlers.
          for i := 0 to EventCount - 1 do
            begin
              ev := Event[i];
              Write( Format( fmtDelphiEventHandlerField, [ev.Name1] ) + ' : ' );
              WriteLn( Format( fmtDelphiEventProcedureName, [ev.Name1] ), true );
            end;
        end;
    end;

  procedure TvbDelphiConverter.WritePutStmt( s : TvbPutStmt; semicolon : boolean );
    var
      fn : string;
    begin
      with fWriter, s do
        begin
          assert( FileNum <> nil );
          assert( VarName <> nil );
          fn := GetOpenStmtFileName( FileNum );
          WriteCmt( TopCmts );
          if RecNum <> nil
            then
              begin
                Write( 'Seek( ' + fn + ', ' );
                WriteExpr( RecNum );
                WriteLn( ' );' );
              end;
          Write( 'Write( ' + fn + ', ' );
          WriteExpr( VarName );
          Write( ' );' );
          WriteCmt( TopRightCmt );
        end;
    end;
    
  procedure TvbDelphiConverter.WriteRaiseEventStmt( s : TvbRaiseEventStmt; semicolon : boolean );
    begin
      with fWriter, s do
        begin
          WriteCmt( TopCmts );
          Write( Format( fmtDelphiRaiseEventFunction, [EventName] ) );
          if ( EventArgs <> nil ) and ( EventArgs.ArgCount > 0 )
            then
              begin
                Write( '( ' );
                WriteArgList( EventArgs, nil );
                Write( ' )' );
              end;
          if semicolon
            then Write( ';' );
          WriteCmt( TopRightCmt );
        end;
    end;

  procedure TvbDelphiConverter.WriteReDimStmt( s : TvbReDimStmt; semicolon : boolean );
    var
      i     : integer;
      bound : TvbArrayBounds;
      n     : integer;
    begin
      with fWriter, s do
        begin
          WriteCmt( TopCmts );

          // We convert ReDim to SetLength.
          // This works well except when:
          //    (1) ReDim changes the number of dimensions. In this case
          //        Delphi complains with a syntax error.
          //    (2) ReDim specifies a lower bound other than 0. This results
          //        in unexpected runtime behavior because in Delphi the
          //        lower bound of dynamic arrays is always 0.
          Write( 'setlength( ' );
          Write( ArrayVar.Name + ', ' );
          n := ArrayVar.VarType.ArrayBoundCount - 1;
          for i := 0 to n do
            begin
              bound := ArrayVar.VarType.ArrayBounds[i];
              assert( bound <> nil );
              with bound do
                begin
                  assert( Upper <> nil );
                  // Now we determine the element count which is what is passed
                  // to SetLength.
                  if Lower = nil
                    then WriteExpr( Upper )
                    else
                      begin
                        // Otherwise the element count is "upper-lower+1" if
                        // Lower<>1.
                        WriteExpr( Upper );
                        if not IsIntegerEqTo( Lower, 1 )
                          then
                            begin
                              // write Lower is it's not zero.
                              if not IsIntegerEqTo( Lower, 0 )
                                then
                                  begin
                                    Write( ' - ' );
                                    WriteExpr( Lower );
                                  end;
                              Write( ' + 1' );
                            end;
                      end;
                end;
              if i < n
                then Write( ', ' );
            end;
          Write( ' )' );
          if semicolon
            then Write( ';' );
          WriteCmt( TopRightCmt );
        end;
    end;

  procedure TvbDelphiConverter.WriteResumeStmt( s : TvbResumeStmt; semicolon : boolean );
    begin
      with fWriter, s do
        begin
          WriteCmt( TopCmts );
          //UNIMPLEMENTED
          WriteLineCmt( 'RESUME' );
        end;
    end;

  procedure TvbDelphiConverter.WriteReturnStmt( s : TvbReturnStmt; semicolon : boolean );
    begin
      with fWriter, s do
        begin
          WriteCmt( TopCmts );
          //UNIMPLEMENTED
          WriteLineCmt( 'RETURN' );
        end;
    end;

  procedure TvbDelphiConverter.WriteSeekStmt( s : TvbSeekStmt; semicolon : boolean );
    begin
      with fWriter, s do
        begin
          WriteCmt( TopCmts );
          Write( 'seek( ' + GetOpenStmtFileName( FileNum ) + ', ' );
          WriteExpr( Pos );
          Write( ' )' );
          if semicolon
            then Write( ';' );
          WriteCmt( TopRightCmt );
        end;
    end;

  procedure TvbDelphiConverter.WriteStmt( stmt : TvbStmt; semicolon : boolean );
    begin
      // Now we write all labels in front of this statement.
      // This step is common to all kinds of statements.
      WriteLabels( stmt.FrontLabel );
      case stmt.NodeKind of
        ASSIGN_STMT          : WriteAssignStmt( TvbAssignStmt( stmt ), semicolon );
        LABEL_DEF_STMT       : ; //WriteLabelDefStmt( TvbLabelDefStmt( stmt ), semicolon );
        IF_STMT              : WriteIfStmt( TvbIfStmt( stmt ), semicolon );
        DO_LOOP_STMT         : WriteDoLoopStmt( TvbDoLoopStmt( stmt ), semicolon );
        FOR_STMT             : WriteForStmt( TvbForStmt( stmt ), semicolon );
        FOREACH_STMT         : WriteForeachStmt( TvbForeachStmt( stmt ), semicolon );
        WITH_STMT            : WriteWithStmt( TvbWithStmt( stmt ), semicolon );
        GOTO_OR_GOSUB_STMT   : WriteGotoOrGosubStmt( TvbGotoOrGosubStmt( stmt ), semicolon );
        RESUME_STMT          : WriteResumeStmt( TvbResumeStmt( stmt ), semicolon );
        EXIT_STMT            : WriteExitStmt( stmt, 'exit', semicolon );
        EXIT_LOOP_STMT       : WriteExitStmt( stmt, 'break', semicolon );
        RETURN_STMT          : WriteReturnStmt( TvbReturnStmt( stmt ), semicolon );
        ON_ERROR_STMT        : WriteOnErrorStmt( TvbOnErrorStmt( stmt ), semicolon );
        ONGOTO_OR_GOSUB_STMT : WriteOnGotoOrGosubStmt( TvbOnGotoOrGosubStmt( stmt ), semicolon );
        CALL_STMT            : WriteCallStmt( TvbCallStmt( stmt ), semicolon );
        RAISE_EVENT_STMT     : WriteRaiseEventStmt( TvbRaiseEventStmt( stmt ), semicolon );
        MID_ASSIGN_STMT      : WriteMidAssignStmt( TvbMidAssignStmt( stmt ), semicolon );
        OPEN_STMT            : WriteOpenStmt( TvbOpenStmt( stmt ), semicolon );
        REDIM_STMT           : WriteReDimStmt( TvbReDimStmt( stmt ), semicolon );
        ERASE_STMT           : WriteEraseStmt( TvbEraseStmt( stmt ), semicolon );
        INPUT_STMT           : WriteInputStmt( TvbInputStmt( stmt ), semicolon );
        NAME_STMT            : WriteNameStmt( TvbNameStmt( stmt ), semicolon );
        DATE_STMT            : WriteDateStmt( TvbDateStmt( stmt ), semicolon );
        STOP_STMT            : WriteStopStmt( TvbStopStmt( stmt ), semicolon );
        TIME_STMT            : WriteTimeStmt( TvbTimeStmt( stmt ), semicolon );
        ERROR_STMT           : WriteErrorStmt( TvbErrorStmt( stmt ), semicolon );
        WRITE_STMT           : WriteWriteStmt( TvbWriteStmt( stmt ), semicolon );
        WIDTH_STMT           : WriteWidthStmt( TvbWidthStmt( stmt ), semicolon );
        SEEK_STMT            : WriteSeekStmt( TvbSeekStmt( stmt ), semicolon );
        PUT_STMT             : WritePutStmt( TvbPutStmt( stmt ), semicolon );
        UNLOCK_STMT          : WriteUnlockStmt( TvbUnlockStmt( stmt ), semicolon );
        LOCK_STMT            : WriteLockStmt( TvbLockStmt( stmt ), semicolon );
        LINE_INPUT_STMT      : WriteLineInputStmt( TvbLineInputStmt( stmt ), semicolon );
        GET_STMT             : WriteGetStmt( TvbGetStmt( stmt ), semicolon );
        CLOSE_STMT           : WriteCloseStmt( TvbCloseStmt( stmt ), semicolon );
        PRINT_STMT           : WritePrintStmt( TvbPrintStmt( stmt ), semicolon );
        END_STMT             : WriteEndStmt( TvbEndStmt( stmt ), semicolon );
        SELECT_CASE_STMT     : WriteSelectCaseStmt( TvbSelectCaseStmt( stmt ), semicolon );
        DEBUG_ASSERT_STMT    : WriteDebugAssertStmt( TvbDebugAssertStmt( stmt ), semicolon );
        DEBUG_PRINT_STMT     : WriteDebugPrintStmt( TvbDebugPrintStmt( stmt ), semicolon );
        CIRCLE_STMT          : WriteCircleStmt( TvbCircleStmt( stmt ), semicolon );
        LINE_STMT            : WriteLineStmt( TvbLineStmt( stmt ), semicolon );
        PSET_STMT            : WritePSetStmt( TvbPSetStmt( stmt ), semicolon );
        else fWriter.WriteLineCmt( ' UNIMPLEMENTED' );
      end;
    end;

  procedure TvbDelphiConverter.WriteStopStmt( s : TvbStopStmt; semicolon : boolean );
    begin
      WriteCmt( s.TopCmts );
      fWriter.Write( 'raise Exception.Create(''Stop'')' );
      if semicolon
        then fWriter.Write( ';' );
      WriteCmt( s.TopRightCmt );
    end;

  procedure TvbDelphiConverter.WriteTimeStmt( s : TvbTimeStmt; semicolon : boolean );
    begin
      with fWriter, s do
        begin
          WriteCmt( TopCmts );
          Write( 'SetTime( ' );
          WriteExpr( NewTime );
          Write( ' )' );
          if semicolon
            then Write( ';' );
          WriteCmt( TopRightCmt );
        end;
    end;

  procedure TvbDelphiConverter.WriteBaseType( const baseType : IvbType );
    var
      bt : cardinal;
    begin
      bt := baseType.Code and not ARRAY_TYPE; // clear the array flag
      with fWriter, baseType do
        case bt of
          INTEGER_TYPE : Write( 'smallint' );
          LONG_TYPE    : Write( 'longint' );
          BOOLEAN_TYPE :
            // Inside record declarations we'll use WORDBOOL.
            if ( csInRecord in fStatus ) and not fOpts.optUseByteBool
              then Write( 'wordbool' )
              else Write( 'boolean' );
              
          STRING_TYPE :
            // If this is a fixed-length string check (1) if we are inside
            // a record and (2) have to convert it to an array of chars.
            // If not (1) or not (2) we'll declare "string[N]".
            if StringLen <> nil
              then
                if ( csInRecord in fStatus ) and ( fOpts.optStringNToArrayOfChar )
                  then
                    begin
                      Write( 'array[1..' );
                      WriteExpr( StringLen );
                      Write( '] of char' );
                    end
                  else
                    begin
                      // I use 'string[N]' in spite of fOpts.optUseAnsiString
                      // becuase Delphi doesn't support 'widestring[N]'.
                      Write( 'string[' );
                      WriteExpr( StringLen );
                      Write( ']' );
                    end
              else
                if fOpts.optUseAnsiString
                  then Write( 'string' )
                  else Write( 'widestring' );

          BYTE_TYPE     : Write( 'byte' );
          SINGLE_TYPE   : Write( 'single' );
          DATE_TYPE     : Write( 'TDateTime' );
          DOUBLE_TYPE   : Write( 'double' );
          CURRENCY_TYPE : Write( 'currency' );
          OBJECT_TYPE   : Write( 'olevariant' );
          VARIANT_TYPE  : Write( 'olevariant' );

          // We should never get a reference to the type Any outside
          // Declare statements. But since it means that the parameter
          // gets the address of the argument we translate it as a pointer
          // just in case.
          ANY_TYPE : Write( 'pointer' );
          else Write( Name1 );
        end;
    end;

  procedure TvbDelphiConverter.WriteType( const typ : IvbType );
    var
      i, n     : integer;
      b        : TvbArrayBounds;
      baseType : cardinal;
    begin
      assert( typ <> nil );
      with fWriter, typ do
        begin
          baseType := Code and not ARRAY_TYPE;
          if Code and ARRAY_TYPE <> 0
            then
              begin
                if ArrayBoundCount > 0  // an static array(a.k.a. vector)?
                  then
                    begin
                      Write( 'array[');
                      n := ArrayBoundCount - 1;
                      for i := 0 to n do
                        begin
                          b := ArrayBounds[i];
                          if b.Lower = nil
                            then Write( fModule.OptionBase )
                            else WriteExpr( b.Lower );
                          Write( '..' );
                          assert( b.Upper <> nil );
                          WriteExpr( b.Upper );
                          if i < n
                            then Write( ', ' );
                        end;
                      Write( '] of ');
                      WriteBaseType( typ );
                    end
                  else
                    case baseType of
                      STRING_TYPE  : Write( 'TStringDynArray' );
                      INTEGER_TYPE : Write( 'TSmallIntDynArray' );
                      LONG_TYPE    : Write( 'TIntegerDynArray' );
                      BOOLEAN_TYPE : Write( 'TBooleanDynArray' );
                      DATE_TYPE,
                      DOUBLE_TYPE  : Write( 'TDoubleDynArray' );
                      SINGLE_TYPE  : Write( 'TSingleDynArray' );
                      BYTE_TYPE    : Write( 'TByteDynArray' );
                      else
                        begin
                          Write( 'array of ');
                          WriteBaseType( typ );
                        end;
                    end;
              end
            else WriteBaseType( typ );
        end;
    end;

  procedure TvbDelphiConverter.WriteTypeSection( f : TvbNodeFlags );
    var
      i, j, n  : integer;
      any      : boolean;
      enum     : TvbEnumDef;
      cst      : TvbConstDef;
      rec      : TvbRecordDef;
    begin
      any := false;
      with fWriter do
        begin
          for i := 0 to fModule.EnumCount - 1 do
            begin
              enum := fModule.Enum[i];
              with enum do
                if NodeFlags * f <> []
                  then
                    begin
                      // Write the section header if we haven't written it
                      // already.
                      if not any
                        then
                          begin
                            WriteLn( 'type' );
                            Indent;
                            any := true;
                          end;

                      WriteCmt( TopCmts );
                      if ( fOpts.optEnumToConst ) and not ( enum_MissingValues in NodeFlags )
                        then
                          begin
                            Write( Name + ' = longint; ' );
                            WriteCmt( TopRightCmt );
                          end
                        else
                          begin
                            Write( Name + ' = (' );
                            WriteCmt( TopRightCmt );
                            Indent;
                            n := EnumCount - 1;
                            for j := 0 to n do
                              begin
                                cst := Enum[j];
                                Write( cst.Name );
                                if cst.Value <> nil
                                  then
                                    begin
                                      Write( ' = ' );
                                      WriteExpr( cst.Value );
                                    end;
                                if j < n
                                  then Write( ', ' );
                                WriteCmt( cst.TopRightCmt );
                              end;
                            WriteCmt( DownCmts );
                            Write( ');' );
                            WriteCmt( DownRightCmt );
                            Outdent;
                          end;
                    end;
            end;

          if any
            then WriteLn;

          for i := 0 to fModule.RecordCount - 1 do
            begin
              rec := fModule.Records[i];
              with rec do
                if NodeFlags * f <> []
                  then
                    begin
                      if not any
                        then
                          begin
                            WriteLn( 'type' );
                            Indent;
                            any := true;
                          end;
                      WriteCmt( TopCmts );
                      Write( Name + ' = ' );
                      WriteCmt( TopRightCmt );
                      Indent;
                      WriteLn( 'packed record' );
                      Indent;
                      Include( fStatus, csInRecord );
                      for j := 0 to FieldCount - 1 do
                        WriteVar( Field[j] );
                      Exclude( fStatus, csInRecord );
                      WriteCmt( DownCmts );
                      WriteEnd( true, DownRightCmt );
                      Outdent;
                      WriteLn;
                    end;
            end;

          if any
            then Outdent;
        end;
    end;

  procedure TvbDelphiConverter.WriteUnlockStmt( s : TvbUnlockStmt; semicolon : boolean );
    begin
      with fWriter, s do
        begin
          WriteCmt( TopCmts );
          //UNIMPLEMENTED
          WriteLineCmt( 'UNLOCK' );
        end;
    end;

  procedure TvbDelphiConverter.WriteVar( v : TvbVarDef );
    var
      i : integer;
    begin
      with fWriter, v do
        begin
          WriteCmt( TopCmts );
          Write( Name1 );
          assert( VarType <> nil );
          Write( ' : ' );

          // The nfReDimArray flag indicates this array has been declared
          // with a ReDim statement. The Dephi dynamic array that we'll
          // declare now will have the same number of dimensions
          // specified in the ReDim. If later another ReDim statement tries
          // to change the number of dimensions of this same array
          // Delphi will complain.
          if nfReDim_DeclaredLocalVar in NodeFlags
            then
              with VarType do
                begin
                  assert( ArrayBoundCount > 0 );
                  for i := 0 to ArrayBoundCount - 1 do
                    Write( 'array of ' );
                  WriteBaseType( VarType );
                end
            else WriteType( VarType );
          Write( '; ' );
          WriteCmt( TopRightCmt );
        end;
    end;

  procedure TvbDelphiConverter.WriteVarSection( f : TvbNodeFlags );
    var
      i   : integer;
      v   : TvbVarDef;
      any : boolean;
    begin
      any := false;
      with fWriter do
        begin
          for i := 0 to fModule.VarCount - 1 do
            begin
              v := fModule.Vars[i];
              if v.NodeFlags * f <> []
                then
                  begin
                    // Write the section header if we haven't written it
                    // already.
                    if not any
                      then
                        begin
                          WriteLn( 'var' );
                          Indent;
                          any := true;
                        end;
                    WriteVar( v );
                  end;
            end;
          if any
            then
              begin
                Outdent;
                WriteLn;
              end;
        end;
    end;

  procedure TvbDelphiConverter.WriteWidthStmt( s : TvbWidthStmt; semicolon : boolean );
    begin
      with fWriter, s do
        begin
          WriteCmt( TopCmts );
          //UNIMPLEMENTED
          WriteLineCmt( 'WIDTH' );
        end;
    end;

  procedure TvbDelphiConverter.WriteWithStmt( s : TvbWithStmt; semicolon : boolean );
    begin
      with fWriter, s do
        begin
          WriteCmt( TopCmts );
          Write( 'with ' );
          WriteExpr( ObjOrRec );
          Write( ' do' );
          if NeedBlock( StmtBlock, DownRightCmt )
            then
              begin
                WriteCmt( TopRightCmt );
                WriteStmtBlock( StmtBlock, DownRightCmt, semicolon );
              end
            else
              begin
                Write( ';' );
                WriteCmt( TopRightCmt );
              end;
        end;
    end;

  procedure TvbDelphiConverter.WriteWriteStmt( s : TvbWriteStmt; semicolon : boolean );
    var
      i : integer;
    begin
      with fWriter, s do
        begin
          WriteCmt( TopCmts );
          Write( 'WriteLn( ' + GetOpenStmtFileName( FileNum ) );
          for i := 0 to ValueCount - 1 do
            begin
              Write( ', ' );
              WriteExpr( Value[i] );
            end;
          Write( ' )' );
          if semicolon
            then Write( ';' );
          WriteCmt( TopRightCmt );
        end;
    end;

  // Defines a property for each public variable in the class module.
  procedure TvbDelphiConverter.WriteClassProperties;
    var
      i, j, n   : integer;
      v         : TvbVarDef;
      ev        : TvbEventDef;
      fieldName : string;
    begin
      with fWriter, fModule do
        begin
          // Write public variables as properties.
          // In the private section of the class we already defined
          // fields to store these properties values.
          for i := 0 to VarCount - 1 do
            begin
              v := Vars[i];
              if dfPublic in v.NodeFlags
                then
                  begin
                    WriteCmt( v.TopCmts );
                    Write( 'property ' + v.Name1 + ' : ' );
                    WriteType( v.VarType );
                    fieldName := Format( fmtDelphiClassField, [v.Name] );
                    Write( ' read ' + fieldName );
                    Write( ' write ' + fieldName );
                    Write( '; ' );
                    WriteCmt( v.TopRightCmt );
                  end;
            end;

          // Now write the property declarations.
          for i := 0 to PropCount - 1 do
            with Prop[i] do
              if true {dfPublic in NodeFlags}
                then
                  begin
                    Write( 'property ' + Name1 );
                    if ParamCount > 0
                      then
                        begin
                          Write( '[' );
                          n := ParamCount - 1;
                          for j := 0 to n do
                            begin
                              WriteParam( Param[j], false );
                              if j < n
                                then Write( '; ' );
                            end;
                          Write( ']' );
                        end;
                    Write( ' : ' );
                    WriteType( PropType );
                    if GetFn <> nil
                      then Write( ' read ' + GetFn.Name1 );
                    if SetFn <> nil
                      then Write( ' write ' + SetFn.Name1 )
                      else
                    if LetFn <> nil
                      then Write( ' write ' + LetFn.Name1 );
                    WriteLn( ';' );
                  end;

          // Write the properties to set/get the event handlers.
          for i := 0 to EventCount - 1 do
            begin
              ev := Event[i];
              Write( 'property ' + Format( fmtDelphiEventProperty, [ev.Name] ) + ' : ' );
              Write( Format( fmtDelphiEventProcedureName, [ev.Name] ) );
              fieldName := Format( fmtDelphiEventHandlerField, [ev.Name] );
              Write( ' read ' +  fieldName );
              WriteLn( ' write ' +  fieldName, true );
            end;
        end;
    end;

  procedure TvbDelphiConverter.WriteParam( p : TvbParamDef; WithDefaultValue : boolean );
    begin
      with fWriter, p do
        if _pfAnyByRef in NodeFlags
          then Write( 'var ' + Name )
          else
        if ( _pfAnyByVal in NodeFlags ) or ( _pfPointerType in NodeFlags )
          then Write( Name + ' : pointer' )
          else
            begin
              if pfByRef in NodeFlags
                then Write( 'var ' );      
              Write( Name );
              Write( ' : ' );
              if ParamType = vbVariantArrayType
                then Write( 'array of const' )
                else
              if _pfPChar in NodeFlags
                then Write( 'pchar' )
                else
              if _pfWordBool in NodeFlags
                then Write( 'wordbool' )
                else
                  begin
                    WriteType( ParamType );
                    if WithDefaultValue
                      then
                        begin
                          if ( pfByRef in NodeFlags ) and
                             ( pfOptional in NodeFlags )
                            then Write( '{' );

                          if DefaultValue <> nil
                            then
                              begin
                                Write( ' = ' );
                                WriteExpr( DefaultValue );
                              end
                            else
                          if ( pfOptional in NodeFlags ) and
                             ( ParamType.Code <> VARIANT_TYPE )
                            then
                              begin
                                Write( ' = ' );
                                WriteTypeDefaultValue( ParamType.Code );
                              end;

                          if ( pfByRef in NodeFlags ) and
                             ( pfOptional in NodeFlags )
                            then Write( '}' );
                        end;
                  end;
            end;
    end;

  // Writes a type for each event declared in the class.
  // For example, if we have this VB event:
  //
  //    Event FieldModified(ByVal FieldName As string)
  //
  // this function generates an event type like:
  //
  //    TFieldModifiedEvent = procedure( FieldName : string );
  procedure TvbDelphiConverter.WriteEventTypes;
    var
      i  : integer;
      ev : TvbEventDef;
    begin
      if fModule.EventCount > 0
        then
          with fWriter do
            begin
              for i := 0 to fModule.EventCount - 1 do
                begin
                  ev := fModule.Event[i];
                  Write( Format( fmtDelphiEventProcedureName, [ev.Name] ) + ' = procedure' );
                  WriteFunctionParams( ev.Params, false );
                  WriteLn( ' of object;' );
                end;
              WriteLn;
            end;
    end;

  // Write the given parameter list.
  // params CANNOT be nil.
  procedure TvbDelphiConverter.WriteParamList( params            : TvbParamList;
                                               WithDefaultValues : boolean );
    var
      i, n : integer;
    begin
      assert( params <> nil );
      with fWriter, params do
        begin
          n := ParamCount - 1;
          for i := 0 to n do
            begin
              WriteParam( Params[i], WithDefaultValues );
              if i < n
                then Write( '; ' );
            end;
        end;
    end;

  procedure TvbDelphiConverter.WriteFunctionParams( params            : TvbParamList;
                                                    WithDefaultValues : boolean );
    begin
      if ( params <> nil ) and ( params.ParamCount > 0 )
        then
          with fWriter do
            begin
              Write( '( ' );
              WriteParamList( params, WithDefaultValues );
              Write( ' )' );
            end;
    end;

  // This function writes the event raising procedures, one for each event
  // declared in the class. Assuming that in a VB class module we have the
  // following event declaration:
  //
  //    Event UserLoggedOn(ByVal UserName As String)
  //
  // the generated event-firing procedure would look like this:
  //
  //    procedure TClass.DoUserLoggedOn( UserName : string );
  //      begin
  //        if fOnUserLoggedOn <> nil
  //          then fOnUserLoggedOn( UserName );
  //      end;
  procedure TvbDelphiConverter.WriteRaiseEventFunctions;
    var
      i, j, n : integer;
      ev      : TvbEventDef;
    begin
      with fWriter, fModule do
        for i := 0 to EventCount - 1 do
          begin
            ev := Event[i];
            Write( 'procedure ' );
            if csInImplem in fStatus
              then Write( Name1 + '.' );
            Write( Format( fmtDelphiRaiseEventFunction, [ev.Name] ) );
            WriteFunctionParams( ev.Params, csInInterface in fStatus );
            WriteLn( ';' );
            if csInImplem in fStatus
              then
                begin
                  Indent;
                  WriteBegin;
                  WriteLn( 'if ' + Format( fmtDelphiEventHandlerField, [ev.Name] ) + ' <> nil' );
                  Indent;
                  Write( 'then ' + Format( fmtDelphiEventHandlerField, [ev.Name] ) );
                  with ev do
                    if ( Params <> nil ) and ( Params.ParamCount > 0 )
                      then
                        begin
                          Write( '( ' );
                          n := Params.ParamCount - 1;
                          for j := 0 to n do
                            begin
                              Write( Params[j].Name );
                              if j < n
                                then Write( ', ' );
                            end;
                          Write( ' )' );
                        end;
                  WriteLn( ';' );
                  Outdent;
                  WriteEnd( true );
                  Outdent;
                  WriteLn;
                end;
          end;
    end;

  // This function writes the constuctor and destructor of a class.
  // In the class declaration the code looks like this:
  //
  //    constructor Create;
  //    destructor Destroy; override;
  //
  // In the implementation section the code would be:
  //
  //    constructor TClass.Create;
  //      begin
  //        stmts;
  //      end;
  //
  //    destructor TClass.Destroy;
  //      begin
  //        stmts;
  //      end;
  procedure TvbDelphiConverter.WriteConstructorAndDestructor( fnConstructor, fnDestructor : TvbAbstractFuncDef );
    begin
      with fWriter do
        begin
          if fnConstructor <> nil
            then
              if csInInterface in fStatus
                then
                  begin
                    WriteCmt( fnConstructor.TopCmts );
                    WriteLn( 'constructor Create;' )
                  end
                else
                  begin
                    WriteLn( 'constructor ' + fClassDef.Name1 + '.Create;' );
                    WriteFunctionBody( fnConstructor );
                    WriteLn;
                  end;
          if fnDestructor <> nil
            then
              if csInInterface in fStatus
                then
                  begin
                    WriteCmt( fnDestructor.TopCmts );
                    WriteLn( 'destructor Destroy; override;' )
                  end
                else
                  begin
                    WriteLn( 'destructor ' + fClassDef.Name1 + '.Destroy;' );
                    WriteFunctionBody( fnDestructor );
                    WriteLn;
                  end;
        end;
    end;

  // This function writes the body of a function/procedure.
  // The order of the sections is:
  //
  //    const
  //      const1 = value1;
  //      const2 = value2;
  //      ....
  //      constN = valueN;
  //    var
  //      var1 : type1;
  //      var2 : type2;
  //      ...
  //      varN : typeN;
  //    label
  //      label1,
  //      label2,
  //      ...
  //      labelN;
  //    begin
  //      stmt1;
  //      stmt2;
  //      ...
  //      stmtN;
  //    end;
  procedure TvbDelphiConverter.WriteFunctionBody( f : TvbAbstractFuncDef );
    var
      i, n : integer;
      v    : TvbVarDef;
    begin
      assert( f <> nil );
      fFunction := f;
      //writeln(f.Name);
      with fWriter, f do
        begin
          if ( ConstCount > 0 ) or ( StaticVarCount > 0 )
            then
              begin
                Indent;
                WriteLn( 'const' );
                Indent;
                for i := 0 to ConstCount - 1 do
                  WriteConst( Consts[i] );

                // We simulate VB6 static variables with Delphi's writeable
                // typed constants that also preserve their values between
                // procedure calls.
                for i := 0 to VarCount - 1 do
                  begin
                    v := Vars[i];
                    if dfStatic in v.NodeFlags
                      then
                        begin
                          WriteLn( '{$WRITEABLECONST ON}' );
                          Write( v.Name );
                          Write( ' : ' );
                          WriteType( v.VarType );

                          // Since we are declaring a constant, we must give it
                          // an initial value which we assign if the type is
                          // primitive. Otherwise we left the constant
                          // uninitialized and Delphi will complain!
                          if v.VarType.Code and ARRAY_TYPE = 0
                            then
                              begin
                                Write( ' = ' );
                                WriteTypeDefaultValue( v.VarType.Code );
                              end;
                          WriteLn( ';' );
                          WriteLn( '{$WRITEABLECONST OFF}' );
                        end;
                  end;
                Outdent;
                Outdent;
              end;

          // Check if we have regular vars to declare because static
          // vars were earlier declared in the constant section as
          // writteable typed constants.
          if ( VarCount - StaticVarCount ) > 0
            then
              begin
                Indent;
                WriteLn( 'var' );
                Indent;
                for i := 0 to VarCount - 1 do
                  if not ( dfStatic in Vars[i].NodeFlags )
                    then WriteVar( Vars[i] );
                Outdent;
                Outdent;
              end;
          if LabelCount > 0
            then
              begin
                Indent;
                WriteLn( 'label' );
                Indent;
                n := LabelCount - 1 ;
                for i := 0 to n do
                  begin
                    Write( Labels[i] );
                    if i < n
                      then WriteLn( ',' )
                      else WriteLn( ';' );
                  end;
                Outdent;
                Outdent;
              end;
          WriteStmtBlock( StmtBlock, DownRightCmt, true );
        end;
      fFunction := nil;
    end;

  procedure TvbDelphiConverter.WriteEndStmt( s : TvbEndStmt; semicolon : boolean );
    begin
      WriteCmt( s.TopCmts );
      fWriter.Write( 'Application.Terminate' );
      if semicolon
        then fWriter.Write( ';' );
      WriteCmt( s.TopRightCmt );
    end;

  procedure TvbDelphiConverter.WriteTypeDefaultValue( typ : cardinal );
    begin
      with fWriter do
        case typ of
          BOOLEAN_TYPE : Write( 'false' );
          DOUBLE_TYPE,
          INTEGER_TYPE,
          LONG_TYPE,
          BYTE_TYPE,
          SINGLE_TYPE,
          DATE_TYPE,
          CURRENCY_TYPE : Write( '0' );
          OBJECT_TYPE : Write( 'nil' );
          STRING_TYPE : Write( '''''' );
          else assert( true );
        end;
    end;

  procedure TvbDelphiConverter.WriteCmt( const cmt : TStringDynArray );
    var
      i : integer;
    begin
      for i := low( cmt ) to high( cmt ) do
        fWriter.WriteLineCmt( cmt[i] );
    end;

  procedure TvbDelphiConverter.WriteCmt( const cmt : string );
    begin
      with fWriter do
        if cmt <> ''
          then
            begin
              Write( #32 );
              WriteLineCmt( cmt );
            end
          else WriteLn;
    end;

  // Writes a list of statements enclosed between begin..end.
  // At the bottom of the block writes the DownCmts of stmt.
  // After the closing "end" downRightCmt is written.
  procedure TvbDelphiConverter.WriteStmtBlock( stmts              : TvbStmtBlock;
                                               const downRightCmt : string;
                                               semicolon          : boolean );
    var
      stmt : TvbStmt;
    begin
      assert( stmts <> nil );
      with fWriter do
        begin
          Indent;
          WriteBegin;
          stmt := stmts.FirstStmt;
          while stmt <> nil do
            begin
              WriteStmt( stmt );
              stmt := stmt.NextStmt;
            end;
          WriteCmt( stmts.DownCmts );
          WriteEnd( semicolon, downRightCmt );
          Outdent;
        end;
    end;

  function TvbDelphiConverter.NeedBlock( stmts              : TvbStmtBlock;
                                         const downRightCmt : string ) : boolean;
    begin
      result := ( stmts <> nil ) and
                ( ( stmts.FirstStmt <> nil  ) or // at least one statement?
                  ( length( stmts.DownCmts ) > 0 ) or // or with comments at the bottom?
                  ( downRightCmt <> '' ) ); // or a comment at the bottom right?
    end;

  procedure TvbDelphiConverter.ConvertProject( prj : TvbProject );
    var
      i           : integer;
      moduleDef   : TvbModuleDef;
      projectName : string;
      output      : TStream;
    begin
      assert( fResourceProvider <> nil );
      assert( prj <> nil );
      assert( prj.Name <> '' );

      projectName := prj.Name;

      // generate the .DPR file.
      output := fResourceProvider.CreateOutputStream( projectName + '\' + projectName + '.dpr' );
      try
        fWriter.Stream := output;
        with fWriter do
          begin
            WriteLn( 'program ' + projectName, true );
            WriteLn;
            Indent;
            WriteLn( 'uses' );
            Indent;
            for i := 0 to prj.ModuleCount - 1 do
              begin
                moduleDef := prj.Module[i];
                Write( moduleDef.Name1 + ' in ''' + moduleDef.Name1 + '.pas''');
                if i < prj.ModuleCount - 1
                  then WriteLn( ',' )
                  else WriteLn( true );                  
              end;
            Outdent;
            Outdent;
            WriteLn;
            WriteLn( '{$R *.RES}' );
            WriteLn;
            WriteLn( 'begin' );
            Indent;
            WriteLn( 'Application.Initialize;' );
            //TODO: foreach form in the module.
            //Application.CreateForm(TForm1, Form1);
            WriteLn( 'Application.Run;' );
            Outdent;
            WriteLn( 'end.' );
          end;
      finally
        output.Free;
      end;

      // generate the code for each module.
      with prj do
        for i := 0 to ModuleCount - 1 do
          begin
            moduleDef := Module[i];

            assert( moduleDef <> nil );
            assert( moduleDef.Name1 <> '' );
            
            output := fResourceProvider.CreateOutputStream( projectName + '\' + moduleDef.Name1 + '.pas' );
            try
              assert( output <> nil );
              if moduleDef.NodeKind = STD_MODULE_DEF
                then ConvertStdModule( TvbStdModuleDef( moduleDef ), output )
                else
              if moduleDef.NodeKind = CLASS_MODULE_DEF
                then ConvertClassModule( TvbClassModuleDef( moduleDef ), output )
                else
              if moduleDef.NodeKind =  FORM_MODULE_DEF
                then ConvertFormModule( TvbFormModuleDef( moduleDef ), output )
                else assert( false );
            finally
              FreeAndNil( output );
            end;
          end;
    end;

  procedure TvbDelphiConverter.WriteLabels( lab : TvbLabelDef );
    begin
      while lab <> nil do
        with fWriter, lab do
          begin
            //WriteLn;
            WriteCmt( TopCmts );
            Write( Name + ':' );
            WriteCmt( TopRightCmt );
            lab := lab.NextLabel;
          end;
    end;

  procedure TvbDelphiConverter.WriteCastExpr( e : TvbCastExpr );
    begin
      with fWriter, e do
        begin
          case CastTo  of
            BOOLEAN_CAST :
              if ( Op1.NodeType = vbByteType ) or
                 ( Op1.NodeType = vbIntegerType ) or
                 ( Op1.NodeType = vbLongType ) or
                 ( Op1.NodeType = vbVariantType )
                then WriteFunc( 'boolean', Op1 )
                else
              if Op1.NodeType = vbStringType
                then WriteFunc( 'StrToBool', Op1 )
                else
                  begin
                    Write( 'VarAsType( ' );
                    WriteExpr( Op1 );
                    Write( ', varBoolean )' );
                  end;
            BYTE_CAST :
              if IsStringType( Op1.NodeType )
                then WriteFunc( 'StrToInt', Op1 )
                else
              if IsIntegralType( Op1.NodeType ) or
                 ( Op1.NodeType = vbVariantType )
                then WriteFunc( 'byte', Op1 )
                else
                  begin
                    Write( 'VarAsType( ' );
                    WriteExpr( Op1 );
                    Write( ', varByte )' );
                  end;
            CURRENCY_CAST :
              if IsFloatingType( Op1.NodeType )
                then WriteFunc( 'FloatToCurr', Op1 )
                else
                  begin
                    Write( 'VarAsType( ' );
                    WriteExpr( Op1 );
                    Write( ', varCurrency )' );
                  end;
            DATE_CAST : WriteFunc( 'VarToDateTime', Op1 );
            DOUBLE_CAST :
              if IsStringType( Op1.NodeType )
                then WriteFunc( 'StrToFloat', Op1 )
                else
              if Op1.NodeType = vbVariantType
                then WriteFunc( 'double', Op1 )
                else
                  begin
                    Write( 'double( variant( ' );
                    WriteExpr( Op1 );
                    Write( ' ) )' );
                  end;
            DECIMAL_CAST :
              begin // UNSUPPORTED
              end;
            INTEGER_CAST :
              if IsStringType( Op1.NodeType )
                then WriteFunc( 'StrToInt', Op1 )
                else
              if IsIntegralType( Op1.NodeType ) or
                 ( Op1.NodeType = vbVariantType )
                then WriteFunc( 'smallint', Op1 )
                else
                  begin
                    Write( 'VarAsType( ' );
                    WriteExpr( Op1 );       
                    Write( ', varSmallint )' );
                  end;
            LONG_CAST :
              if IsStringType( Op1.NodeType )
                then WriteFunc( 'StrToInt', Op1 )
                else
              if IsIntegralType( Op1.NodeType ) or
                 ( Op1.NodeType = vbVariantType )
                then WriteFunc( 'integer', Op1 )
                else
                  begin
                    Write( 'VarAsType( ' );
                    WriteExpr( Op1 );
                    Write( ', varInteger )' );
                  end;
            SINGLE_CAST :
              if IsStringType( Op1.NodeType )
                then WriteFunc( 'StrToFloat', Op1 )
                else
              if Op1.NodeType = vbVariantType
                then WriteFunc( 'single', Op1 )
                else
                  begin
                    Write( 'VarAsType( ' );
                    WriteExpr( Op1 );
                    Write( ', varSingle )' );
                  end;
            STRING_CAST :
              // STRING to STRING => no typecast at all
              // Just enclose the expression in parenthesis.
              if IsStringType( Op1.NodeType )
                then WriteExpr( Op1, true )
                else
              if IsIntegralType( Op1.NodeType )
                then WriteFunc( 'IntToStr', Op1 )
                else
              if Op1.NodeType = vbDateType
                then WriteFunc( 'DateToStr', Op1 )
                else
              if IsFloatingType( Op1.NodeType )
                then WriteFunc( 'FloatToStr', Op1 )
                else
              if Op1.NodeType = vbCurrencyType
                then WriteFunc( 'CurrToStr', Op1 )
                else
              if Op1.NodeType = vbVariantType
                then WriteFunc( 'string', Op1 )
                else
                  begin
                    Write( 'VarAsType( ' );
                    WriteExpr( Op1 );
                    Write( ', varOleStr )' );
                  end;
            VARIANT_CAST :
              if Op1.NodeType <> vbVariantType
                then WriteFunc( 'variant', Op1 )
                else WriteExpr( Op1, true );
            VDATE_CAST :
              begin
                Write( 'VarAsType( ' );
                WriteExpr( Op1 );
                Write( ', varDate )' );
              end;
            VERROR_CAST :
              begin
                Write( 'VarAsType( ' );
                WriteExpr( Op1 );
                Write( ', varError )' );
              end;
          end;
        end;
    end;

  procedure TvbDelphiConverter.WriteMidExpr( e : TvbMidExpr );
    begin
      with fWriter, e do
        if Len <> nil
          then
            // Convert Mid$(str, start, 1) to "str[start]" when
            // "str" is not an string literal.
            if IsIntegerEqTo( Len, 1 ) and ( Str.NodeKind <> STRING_LIT_EXPR )
              then
                begin
                  WriteExpr( Str );
                  Write( '[' );
                  WriteExpr( Start );
                  Write( ']' );
                end
              else
                begin
                  // Convert Mid$(str,start,len) to Copy(str,start,len)
                  Write( 'copy( ' );
                  WriteExpr( Str );
                  Write( ', ' );
                  WriteExpr( Start );
                  Write( ', ' );
                  WriteExpr( Len );
                  Write( ' )' );
                end
        else
          begin
            // Convert Mid$(str,start) to RightStr(str,start,length(str));
            Write( 'RightStr( ' );
            WriteExpr( Str );
            Write( ', ' );
            WriteExpr( Start );
            Write( ', ' );
            Write( 'length( ' );
            WriteExpr( Str );
            Write( ' ) )' );
          end;
    end;

  procedure TvbDelphiConverter.WriteArrayExpr( e : TvbArrayExpr );
    var
      i, n : integer;
    begin
      with fWriter, e do
        begin
          Write( 'VarArrayOf( [' );
          n := ValueCount - 1;
          for i := 0 to n do
            begin
              WriteExpr( Values[i] );
              if i < n
                then Write( ', ' );
            end;
          Write( '] )' );
        end;
    end;

  procedure TvbDelphiConverter.WriteCallOrIndexerExpr( e : TvbCallOrIndexerExpr );

    procedure WriteArgs( withBrackets : boolean; args : TvbArgList; params : TvbParamList = nil );
      begin
        with fWriter do
          begin
            if withBrackets
              then Write( '[' )
              else Write( '( ' );
            WriteArgList( args, params );
            if withBrackets
              then Write( ']' )
              else Write( ' )' );
          end;
      end;

    var
      def        : TvbDef;
      target     : TvbExpr;
      funcOrVar  : TvbExpr;
      withExpr   : TvbExpr;
      mapDefault : boolean;
      mapFull     : boolean;
      args       : TvbArgList;
      params     : TvbParamList;
    begin
      with fWriter do
        begin
          args      := e.Args;
          def       := e.Def;
          funcOrVar := e.FuncOrVar;
          assert( funcOrVar <> nil );

          // if we've got an external definition check which mapping we have
          // to apply.
          if ( def <> nil ) and ( dfExternal in def.NodeFlags )
            then
              begin
                mapDefault := map_Default in def.NodeFlags;
                mapFull    := map_Full in def.NodeFlags;

                // get the target/qualifier if applicable.
                if funcOrVar.NodeKind = MEMBER_ACCESS_EXPR
                  then
                    begin
                      target   := TvbMemberAccessExpr( funcOrVar ).Target;
                      withExpr := TvbMemberAccessExpr( funcOrVar ).WithExpr;
                      {$IFDEF DEBUG}
                      if target = nil
                        then assert( withExpr <> nil);
                      {$ENDIF}
                    end
                  else
                    begin
                      target := nil;
                      withExpr := nil;
                    end;

                // if the user didn't request "full" mapping write the qualifier
                // if it's present.
                if not mapFull and ( target <> nil )
                  then
                    begin
                      WriteExpr( target );
                      Write( '.' );
                    end;

                if mapDefault or mapFull
                  then WriteMap( def.Name1, args, target, withExpr )
                  else
                    begin
                      Write( def.Name1 );
                      if args <> nil
                        then
                          begin
                            if def.NodeKind in [FUNC_DEF, DLL_FUNC_DEF]
                              then params := TvbFuncDef( def ).Params
                              else params := nil;                              
                            WriteArgs( expr_isNotMethod in e.NodeFlags, args, params );
                          end;
                    end;
              end
            else
              begin
                WriteExpr( funcOrVar );
                if args <> nil
                  then WriteArgs( expr_isNotMethod in e.NodeFlags, args );
              end;
        end;
    end;

  procedure TvbDelphiConverter.WriteUnaryExpr( e : TvbUnaryExpr );
    begin
      assert( e.Op1 <> nil );
      with fWriter, e do
        case Oper of
          NOT_OP :
            begin
              Write( 'not ' );
              WriteExpr( Expr );
            end;
          NEW_OP :
            begin
              //TODO: create the object depending on type
              //CreateOleObject, Type.Create, etc.
              Write( 'new ' );
            end;
          SUB_OP :
            begin
              Write( '-' );
              WriteExpr( Expr );
            end;
          ADD_OP :
            begin
              Write( '+' );
              WriteExpr( Expr );
            end;
        end;
    end;

  procedure TvbDelphiConverter.WriteBinaryExpr( e : TvbBinaryExpr );
    begin
      assert( e.LeftExpr <> nil );
      assert( e.RightExpr <> nil );
      with fWriter, e do
        case Oper of
          EQV_OP :
            begin
              // a EQV b = not (a xor b)
              Write( 'not ( ' );
              WriteExpr( LeftExpr );
              Write( ' xor ' );
              WriteExpr( RightExpr );
              Write( ' )');
            end;
          IMP_OP :
            begin
              // a IMP b = (not a) or b
              Write( '( not ' );
              WriteExpr( LeftExpr);
              Write( ' )' );
              Write( ' or ' );
              WriteExpr( RightExpr );
            end;
          POWER_OP :
            // a^2 = sqr(a)
            // a^b = Power( a, b )
            if IsIntegerEqTo( RightExpr, 2 )
              then WriteFunc( 'sqr', LeftExpr )
              else
                begin
                  Write( 'Power( ' );
                  WriteExpr( LeftExpr );
                  Write( ', ' );
                  WriteExpr( RightExpr );
                  Write( ' )' );
                end;
          else
            begin
              WriteExpr( LeftExpr );
              case Oper of
                ADD_OP,
                STR_CONCAT_OP : Write( ' + ' );
                MULT_OP       : Write( ' * ' );
                SUB_OP        : Write( ' - ' );
                DIV_OP        : Write( ' / ' );
                INT_DIV_OP    : Write( ' div ' );
                IS_OP,
                EQ_OP         : Write( ' = ' );
                AND_OP        : Write( ' and ' );
                OR_OP         : Write( ' or ' );
                XOR_OP        : Write( ' xor ' );
                LESS_OP       : Write( ' < ' );
                LESS_EQ_OP    : Write( ' <= ' );
                GREATER_OP    : Write( ' > ' );
                GREATER_EQ_OP : Write( ' >= ' );
                NOT_EQ_OP     : Write( ' <> ' );
                MOD_OP        : Write( ' mod ' );
                LIKE_OP       : Write( ' like ' ); //TODO: LIKE operator
              end;
              WriteExpr( RightExpr );
            end;
        end;
    end;

  procedure TvbDelphiConverter.WriteInputExpr( e : TvbInputExpr );
    begin
      assert( e.CharCount <> nil );
      assert( e.FileNum <> nil );
      with fWriter, e do
        begin
          //TODO:Input expression
          Write( 'Input( ' );
          WriteExpr( CharCount );
          Write( ', ' );
          WriteExpr( FileNum );
          Write( ' )' );
        end;
    end;

  procedure TvbDelphiConverter.WriteSelectCaseStmt( s : TvbSelectCaseStmt; semicolon : boolean );
    const
      TempVar = 'tmp';

    procedure WriteCaseStmt_AsCase( stmt : TvbCaseStmt );
      var
        j : integer;
        c : TvbCaseClause;
      begin
        with fWriter, stmt do
          begin
            Indent;
            // Write the case expressions separated by commas.
            for j := 0 to ClauseCount - 1 do
              begin
                c := Clause[j];
                assert( c.Expr <> nil );
                case c.NodeKind of
                  CASE_CLAUSE,
                  REL_CASE_CLAUSE : WriteExpr( c.Expr );
                  RANGE_CASE_CLAUSE :
                    begin
                      WriteExpr( c.Expr );
                      Write( '..' );
                      WriteExpr( TvbRangeCaseClause( c ).UpperExpr );
                    end;
                  else assert( true ); // internal error!
                end;
                if j < ClauseCount - 1
                  then WriteLn( ', ' );
              end;
            Write( ' :' );
            WriteCmt( TopRightCmt );
            WriteStmtBlock( StmtBlock );
            Outdent;
          end;
      end;

    procedure WriteCaseStmt_AsIf( stmt : TvbCaseStmt; semicolon : boolean );
      var
        j : integer;
        C : TvbCaseClause;
      begin
        with fWriter, stmt do
          begin
            Write( 'if ' );
            for j := 0 to ClauseCount - 1 do
              begin
                Write( '( ' );
                C := Clause[j];
                assert( C.Expr <> nil );
                case C.NodeKind of
                  CASE_CLAUSE : // tmp = expr
                    begin
                      Write( TempVar + ' = ' );
                      WriteExpr( C.Expr );
                    end;
                  REL_CASE_CLAUSE :  // tmp [REL_OP] expr, eg: tmp >= 3
                    begin
                      Write( TempVar );
                      case TvbRelationalCaseClause( C ).Operator of
                        EQ_OP         : Write( ' = ' );
                        GREATER_EQ_OP : Write( ' >= ' );
                        GREATER_OP    : Write( ' > ' );
                        LESS_EQ_OP    : Write( ' <= ' );
                        LESS_OP       : Write( ' < ' );
                        NOT_EQ_OP     : Write( ' <> ' );
                        else assert( true );
                      end;
                      WriteExpr( C.Expr );
                    end;
                  RANGE_CASE_CLAUSE : // ( tmp >= expr1 ) and ( tmp <= expr2 )
                    begin
                      Write( '( ' );
                      Write( TempVar + ' => ' );
                      WriteExpr( C.Expr );
                      Write( ' ) and ( ' );
                      Write( TempVar + ' <= ' );
                      WriteExpr( TvbRangeCaseClause( C ).UpperExpr );
                      Write( ' )' );
                    end;
                  else assert( true ); // internal error!
                end;
                Write( ' )' );
                if j < ClauseCount - 1
                  then Write( ' or ' );
              end;
            WriteCmt( TopRightCmt );
            Indent;
            WriteLn( 'then' );
            WriteStmtBlock( StmtBlock, DownRightCmt, semicolon );
            Outdent;
          end;
      end;

    var
      i, cc : integer;
    begin
      with fWriter, s do
        begin
          if not ( select_AsIfElse in NodeFlags )
            then
              begin
                // Write "case expr of"
                Write( 'case ' );
                assert( Expr <> nil );
                WriteExpr( Expr );
                Write( ' of' );
                WriteCmt( TopRightCmt );
                for i := 0 to CaseCount - 1 do
                  WriteCaseStmt_AsCase( Cases[i] );
                if CaseElse <> nil
                  then
                    begin
                      Indent;
                      WriteLn( 'else' );
                      WriteStmtBlock( CaseElse );
                      Outdent;
                    end;
                Write( 'end' );
                if semicolon
                  then Write( ';' );
                WriteCmt( DownRightCmt );
              end
            else
              begin
                Write( TempVar + '{DECLARE} := ' );
                WriteExpr( Expr );
                Write( ';' );
                WriteCmt( TopRightCmt );
                cc := CaseCount - 1;
                //TODO: el DownRightCmt del Select..Case cuando no hay Case..Else...
                // ahora solo se escribe cuando hay Case..Else.
                for i := 0 to cc do
                  begin
                    WriteCaseStmt_AsIf( Cases[i], ( i = cc ) and ( CaseElse = nil ) );
                    if i < cc
                      then
                        begin
                          Indent;
                          WriteLn( 'else' );
                          Outdent;
                        end;
                  end;
                if CaseElse <> nil
                  then
                    begin
                      Indent;
                      WriteLn( 'else' );
                      WriteStmtBlock( CaseElse, DownRightCmt, semicolon );
                      Outdent;
                    end;
              end;
        end;
    end;

//    var
//      i               : integer;
//      frxName         : string;
//      itemCountByType : array[TfrxItemType] of integer;
//      items           : TfrxItemArray;
//      stream          : TStream;
//    begin
//      fillchar( itemCountByType, sizeof( itemCountByType ), 0 );
//      frxName := ChangeFileExt( frm.RelativePath, '.frx' );
//      items := fResourceProvider.GetFormItemArray( frxName );
//      for i := low( items ) to high( items ) do
//        begin
//          inc( itemCountByType[items[i].ItemType] );
//          stream := fResourceProvider.CreateOutputStream(
//            frm.Name +
//            '_' +
//            TfrxItemExtension[items[i].ItemType] +
//            IntToStr( itemCountByType[items[i].ItemType] ) +
//            '.' +
//            TfrxItemExtension[items[i].ItemType] );
//          try
//            items[i].Save( stream );
//          finally
//            FreeAndNil( stream );
//          end;
//        end;
//    end;
  procedure TvbDelphiConverter.ConvertFormModule( form : TvbFormModuleDef; OutputStream : TStream );
    begin
      assert( form <> nil );
      assert( OutputStream <> nil );

      if fMonitor <> nil
        then fMonitor.ConvertModuleBegin( form );

      GenerateDFM( form );

      fModule := form;
      fClassDef := form.ClassDef;
      assert( fClassDef <> nil );
      with fWriter, form do
        try
          Stream := OutputStream;
          WriteLn( 'unit ' + fModule.Name1 + ';' );
          WriteLn;

          WriteDisclaimerComment;
          Include( fStatus, csInInterface );
          WriteLn( 'interface' );
          WriteLn;
          Indent;
          WriteLn( 'uses' );
          Indent;
          WriteLn( 'Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,');
          WriteLn( 'Dialogs;');
          //TODO: write refereced units in project...
          WriteLn;
          Outdent;
          WriteConstSection( [dfPublic] );
          WriteTypeSection( [dfPublic] );

          // Write event types and class declaration
          WriteLn( 'type' );
          Indent;
          WriteEventTypes;
          WriteLn( fClassDef.Name1 + ' =' );
          Indent;
          WriteLn( 'class( TForm )' );
          Indent;
          WriteFormControls( form.FormDef );
          WriteLn( 'private' );
          Indent;
          WriteClassFields;
          WriteRaiseEventFunctions;
          WriteClassFunctions( [dfPrivate] );
          Outdent;
          WriteLn( 'public' );
          Indent;
          //WriteConstructorAndDestructor( ClassInitialize, ClassTerminate );
          WriteClassFunctions( [dfPublic] );
          WriteClassProperties;
          Outdent;
          Outdent;
          WriteLn( 'end;' );
          Outdent;
          Outdent;
          WriteLn;
          WriteDllFuncSection( [dfPublic] );
          if length( DownCmts ) > 0
            then
              begin
                WriteCmt( DownCmts );
                WriteLn;
              end;
          Exclude( fStatus, csInInterface );
          Include( fStatus, csInImplem );
          Outdent;
          WriteLn( 'implementation' );
          WriteLn;
          Indent;
          WriteLn( '{$R *.DFM}' );
          WriteLn;
          WriteConstSection( [dfPrivate] );
          WriteTypeSection( [dfPrivate] );
          WriteDllFuncSection( [dfPrivate] );
          //WriteConstructorAndDestructor( ClassInitialize, ClassTerminate );
          WriteClassFunctions( [dfPublic, dfPrivate] );
          WriteRaiseEventFunctions;
          Exclude( fStatus, csInImplem );
          Outdent;
          WriteLn( 'end.' );

          if fMonitor <> nil
            then fMonitor.ConvertModuleEnd( form );
        finally
          //
        end;
    end;

  destructor TvbDelphiConverter.Destroy;
    begin
      FreeAndNil( fWriter );
      inherited;
    end;

  procedure TvbDelphiConverter.GenerateDFM( form : TvbFormModuleDef );
    begin
    end;

  procedure TvbDelphiConverter.WriteMap( const map   : string;
                                         args        : TvbArgList;
                                         qualifier   : TvbExpr;
                                         withExpr    : TvbExpr;
                                         value       : TvbExpr;
                                         writeAltMap : boolean );
    var
      altMapPos   : pchar;
      arg         : TvbExpr;
      ch          : pchar;
      firstArgPos : integer;
      i           : integer;
      lastArgPos  : integer;

    function GetArgumentAt( p : cardinal ) : TvbExpr;
      begin
        if args <> nil
          then result := args[p - 1]
          else result := nil;
      end;

    begin
      ch := pchar( map );
      with fWriter do
        repeat
          // get the position of the next command.
          i := pos( '%', ch );
          if i <> 0
            then
              begin
                // write all character up to the command position.
                Write( copy( ch, 0, i - 1 ) );
                inc( ch, i );
                case ch[0] of
                  '0'..'9' :
                    begin
                      firstArgPos := 0;
                      // parse the first argument index.
                      repeat
                        firstArgPos := firstArgPos * 10 + ( ord( ch[0] ) - ord( '0' ) );
                        inc( ch );
                      until not ( ch[0] in ['0'..'9'] );

                      // check for '..%' and parse the second argument position.
                      if ( ch[0] = '.' ) and
                         ( ch[1] = '.' ) and
                         ( ch[2] = '%' )
                        then
                          begin
                            lastArgPos := 0;
                            inc( ch, 2 );
                            while ch[0] in ['0'..'9'] do
                              begin
                                lastArgPos := lastArgPos * 10 + ( ord( ch[0] ) - ord( '0' ) );
                                inc( ch );
                              end;
                          end
                        else lastArgPos := 0;

                      // get the alternative map if it's been provided.
                      if ch[0] = '{'
                        then
                          begin
                            inc( ch );
                            altMapPos := ch;
                            repeat
                              case ch[0] of
                                '}' :
                                  begin
                                    inc( ch );
                                    break;
                                  end;
                                #0 : break;
                              end;
                              inc( ch );
                            until false;
                          end
                        else altMapPos := nil;

//                      for i := firstArgPos to lastArgPos do
//                        begin
//                          arg := GetArgumentAt( firstArgPos );
//                          if arg = nil
//                            then Write( 'Unassigned' )
//                            else WriteExpr( arg );
//                          if i < lastArgPos - 1
//                            then Write( ', ' );
//                        end;

                      // get the argument at the specified position.
                      arg := GetArgumentAt( firstArgPos );
                      // write the argument if provided, otherwise
                      // write alternative map.
                      if arg = nil
                        then
                          // since the alternative map itself may contain mapping
                          // commands (e.g. reference to other arguments), 
                          // to avoid cross references and thus infinite recursion,
                          // we don't write the alternative maps for
                          // arguments referenced inside another alternative map.
                          if ( altMapPos <> nil ) and writeAltMap
                            then WriteMap( copy( altMapPos, 0, ch - altMapPos - 1 ), args, qualifier, withExpr, value, false )
                            else Write( 'Unassigned' )
                        else WriteExpr( arg );
                    end;
                  'Q' :
                    begin
                      inc( ch );
                      // if we've got the qualifier, write it.
                      // otherwise skip the member access operator '.' if it's
                      // been provided. this way if the user indended %Q.handle
                      // and %Q has not been provided, only "handle" is written.
                      if qualifier <> nil
                        then WriteExpr( qualifier )
                        else
                          if ch[0] = '.'
                            then inc( ch );
                    end;
                  'W' :
                    begin
                      inc( ch );
                      // write the qualifier or the "With" expression.
                      if qualifier <> nil
                        then WriteExpr( qualifier )
                        else
                      if withExpr <> nil
                        then WriteExpr( withExpr );
                    end;
                  'V' :
                    begin
                      inc( ch );
                      if value <> nil
                        then WriteExpr( value )
                        else Write( 'Unassigned' );
                    end;
                  '%' :        
                    begin
                      Write( '%' );
                      inc( ch );
                    end;
                end;
              end
            else
              begin
                Write( ch );  // write from 'ch' to end of string.
                break;
              end;
        until false;
    end;

  procedure TvbDelphiConverter.WriteMemberAccessExpr( mae : TvbMemberAccessExpr );
    var
      def      : TvbDef;
      target   : TvbExpr;
      withExpr : TvbExpr;
    begin
      assert( mae <> nil );
      
      def      := mae.Def;
      target   := mae.Target;
      withExpr := mae.WithExpr;
      
      // do a "full" mapping if the user requested it.
      if ( def <> nil ) and ( map_Full in def.NodeFlags )
        then WriteMap( def.Name1, nil, target, withExpr )
        else
          // otherwise we'll do a "name/default" mapping or
          // no mapping at all if we have no definition.
          with fWriter do
            begin
              if target <> nil
                then
                  begin
                    WriteExpr( target );
                    Write( '.' );
                  end;
              if def <> nil
                then WriteMap( def.Name1, nil, target, withExpr )
                else Write( mae.Member );
            end;
    end;

  procedure TvbDelphiConverter.WriteFormControls( form : TvbFormDef );
    var
      i  : integer;
      ctl : TvbControl;
      indexProp : TvbControlProperty;
      name : string;
    begin
      assert( form <> nil );
      for i := 0 to form.ControlCount - 1 do
        begin
          ctl := form.Control[i];
          assert( ctl <> nil );
          name := ctl.Name;
          indexProp := ctl.PropByName['Index'];
          if indexProp <> nil
            then name := name + '_' + IntToStr( indexProp.Value );            
          fWriter.Write( name + ' : ' );
          WriteType( ctl.VarType );
          fWriter.WriteLn( ';' );
        end;
    end;

  procedure TvbDelphiConverter.WriteCircleStmt( s : TvbCircleStmt; semicolon : boolean );
    begin
      with fWriter, s do
        begin
          if ObjectRef <> nil
            then
              begin
                WriteExpr( ObjectRef );
                Write( '.' );
              end;
          Write( 'Canvas.Arc({TODO})' );
          WriteLn( semicolon );
        end;
    end;

  procedure TvbDelphiConverter.WriteLineStmt( s : TvbLineStmt; semicolon : boolean );
    begin
      with fWriter, s do
        begin
          if ObjectRef <> nil
            then
              begin
                WriteExpr( ObjectRef );
                Write( '.' );
              end;
          Write( 'Canvas.Rectangle({TODO})' );
          WriteLn( semicolon );
        end;
    end;

  procedure TvbDelphiConverter.WritePSetStmt( s : TvbPSetStmt; semicolon : boolean );
    begin
      with fWriter, s do
        begin
          if ObjectRef <> nil
            then
              begin
                WriteExpr( ObjectRef );
                Write( '.' );
              end;
          Write( 'Canvas.Pixels[' );
          WriteExpr( X );
          Write( ', ' );
          WriteExpr( Y );
          Write( '] := ' );
          WriteExpr( Color );
          WriteLn( semicolon );
        end;
    end;

end.



