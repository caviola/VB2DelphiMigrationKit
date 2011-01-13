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

unit vbResolver;

interface

  uses
    vbOptions,
    vbNodes;

  procedure ResolveProject( P : TvbProject; Options : TvbOptions );

  function GetCommonIntegralType( const T1, T2 : IvbType ) : IvbType;
  function GetCommonArithmeticType( const T1, T2 : IvbType ) : IvbType;

  function GetBinaryExpressionType( op : TvbOperator; const T1, T2 : IvbType ): IvbType;
  function GetUnaryExpressionType( op : TvbOperator; const t : IvbType ): IvbType;


implementation

  uses
    StrUtils,
    SysUtils,
    Variants,
    vbConsts, Classes;

  type

    TvbPrefixNodeVisitor =
      class( TvbNodeVisitor )
        private
          fCurrentModule : TvbModuleDef;
        public
          //procedure OnTypeRef( Node : Tvbtype; Context : TObject = nil ); reintroduce; virtual; 
          procedure OnAbsExpr( Node : TvbAbsExpr; Context : TObject = nil ); override;
          procedure OnAddressOfExpr( Node : TvbAddressOfExpr; Context : TObject = nil ); override;
          procedure OnArrayExpr( Node : TvbArrayExpr; Context : TObject = nil ); override;
          procedure OnAssignStmt( Node : TvbAssignStmt; Context : TObject = nil ); override;
          procedure OnBinaryExpr( Node : TvbBinaryExpr; Context : TObject = nil ); override;
          procedure OnCallOrIndexerExpr( Node : TvbCallOrIndexerExpr; Context : TObject = nil ); override;
          procedure OnCallStmt( Node : TvbCallStmt; Context : TObject = nil ); override;
          procedure OnCastExpr( Node : TvbCastExpr; Context : TObject = nil ); override;
          procedure OnCircleStmt( Node : TvbCircleStmt; Context : TObject = nil ); override;
          procedure OnClassDef( Node : TvbClassDef; Context : TObject = nil ); override;
          procedure OnClassModuleDef( Node : TvbClassModuleDef; Context : TObject = nil ); override;
          procedure OnCloseStmt( Node : TvbCloseStmt; Context : TObject = nil ); override;
          procedure OnConstDef( Node : TvbConstDef; Context : TObject = nil ); override;
          procedure OnDateLitExpr( Node : TvbDateLit; Context : TObject = nil ); override;
          procedure OnDateStmt( Node : TvbDateStmt; Context : TObject = nil ); override;
          procedure OnDebugAssertStmt( Node : TvbDebugAssertStmt; Context : TObject = nil ); override;
          procedure OnDebugPrintStmt( Node : TvbDebugPrintStmt; Context : TObject = nil ); override;
          procedure OnDictAccessExpr( Node : TvbDictAccessExpr; Context : TObject = nil ); override;
          procedure OnDllFuncDef( Node : TvbDllFuncDef; Context : TObject = nil ); override;
          procedure OnDoEventsExpr( Node : TvbDoEventsExpr; Context : TObject = nil ); override;
          procedure OnDoLoopStmt( Node : TvbDoLoopStmt; Context : TObject = nil ); override;
          procedure OnEndStmt( Node : TvbEndStmt; Context : TObject = nil ); override;
          procedure OnEnumDef( Node : TvbEnumDef; Context : TObject = nil ); override;
          procedure OnEraseStmt( Node : TvbEraseStmt; Context : TObject = nil ); override;
          procedure OnErrorStmt( Node : TvbErrorStmt; Context : TObject = nil ); override; 
          procedure OnEventDef( Node : TvbEventDef; Context : TObject = nil ); override; 
          procedure OnExitStmt( Node : TvbExitStmt; Context : TObject = nil ); override;
          procedure OnFixExpr( Node : TvbFixExpr; Context : TObject = nil ); override; 
          procedure OnFloatLitExpr( Node : TvbFloatLit; Context : TObject = nil ); override; 
          procedure OnForeachStmt( Node : TvbForeachStmt; Context : TObject = nil ); override; 
          procedure OnForStmt( Node : TvbForStmt; Context : TObject = nil ); override;
          procedure OnFormDef( Node : TvbFormDef; Context : TObject = nil ); override; 
          procedure OnFormModuleDef( Node : TvbFormModuleDef; Context : TObject = nil ); override;
          procedure OnFunctionDef( Node : TvbFuncDef; Context : TObject = nil ); override;
          procedure OnFuncResultExpr( Node : TvbFuncResultExpr; Context : TObject = nil ); override;
          procedure OnGetStmt( Node : TvbGetStmt; Context : TObject = nil ); override;
          procedure OnGotoOrGosubStmt( Node : TvbGotoOrGosubStmt; Context : TObject = nil ); override;
          procedure OnIfStmt( Node : TvbIfStmt; Context : TObject = nil ); override;
          procedure OnInputExpr( Node : TvbInputExpr; Context : TObject = nil ); override;
          procedure OnInputStmt( Node : TvbInputStmt; Context : TObject = nil ); override;
          procedure OnLabelDef( Node : TvbLabelDef; Context : TObject = nil ); override;
          procedure OnLabelDefStmt( Node : TvbLabelDefStmt; Context : TObject = nil ); override;
          //procedure OnLabelRef( Node : tvblabel; Context : TObject = nil ); override;
          procedure OnLBoundExpr( Node : TvbLBoundExpr; Context : TObject = nil ); override;
          procedure OnLenBExpr( Node : TvbLenBExpr; Context : TObject = nil ); override;
          procedure OnLenExpr( Node : TvbLenExpr; Context : TObject = nil ); override;
          procedure OnLibModuleDef( Node : TvbLibModuleDef; Context : TObject = nil ); override;
          procedure OnLineStmt( Node : TvbLineStmt; Context : TObject = nil ); override;
          procedure OnLineInputStmt( Node : TvbLineInputStmt; Context : TObject = nil ); override;
          procedure OnLockStmt( Node : TvbLockStmt; Context : TObject = nil ); override;
          procedure OnMeExpr( Node : TvbMeExpr; Context : TObject = nil ); override;
          procedure OnMemberAccessExpr( Node : TvbMemberAccessExpr; Context : TObject = nil ); override;
          procedure OnMidAssignStmt( Node : TvbMidAssignStmt; Context : TObject = nil ); override;
          procedure OnMidExpr( Node : TvbMidExpr; Context : TObject = nil ); override;
          procedure OnNameExpr( Node : TvbNameExpr; Context : TObject = nil ); override;
          procedure OnNameStmt( Node : TvbNameStmt; Context : TObject = nil ); override;
          procedure OnNewExpr( Node : TvbNewExpr; Context : TObject = nil ); override;
          procedure OnNothingExpr( Node : TvbNothingLit; Context : TObject = nil ); override;
          procedure OnOnErrorStmt( Node : TvbOnErrorStmt; Context : TObject = nil ); override;
          procedure OnOnGotoOrGosubStmt( Node : TvbOnGotoOrGosubStmt; Context : TObject = nil ); override;
          procedure OnOpenStmt( Node : TvbOpenStmt; Context : TObject = nil ); override;
          procedure OnParamDef( Node : TvbParamDef; Context : TObject = nil ); override;
          procedure OnPrintArgExpr( Node : TvbPrintStmtExpr; Context : TObject = nil ); override;
          procedure OnPrintStmt( Node : TvbPrintStmt; Context : TObject = nil ); override;
          procedure OnProjectDef( Node : TvbProject; Context : TObject = nil ); override;
          procedure OnPropertyDef( Node : TvbPropertyDef; Context : TObject = nil ); override;
          procedure OnPSetStmt( Node : TvbPSetStmt; Context : TObject = nil ); override;
          procedure OnPutStmt( Node : TvbPutStmt; Context : TObject = nil ); override;
          procedure OnQualifiedName( Node : TvbQualifiedName; Context : TObject = nil ); override;
          procedure OnRaiseEventStmt( Node : TvbRaiseEventStmt; Context : TObject = nil ); override;
          procedure OnRecordDef( Node : TvbRecordDef; Context : TObject = nil ); override;
          procedure OnReDimStmt( Node : TvbReDimStmt; Context : TObject = nil ); override; 
          procedure OnResumeStmt( Node : TvbResumeStmt; Context : TObject = nil ); override; 
          procedure OnReturnStmt( Node : TvbReturnStmt; Context : TObject = nil ); override;
          procedure OnSelectCaseStmt( Node : TvbSelectCaseStmt; Context : TObject = nil ); override;
          procedure OnCaseStmt( Node : TvbCaseStmt; Context : TObject = nil ); override;
          procedure OnRelationalCaseClause( Node : TvbRelationalCaseClause; Context : TObject = nil ); override;
          procedure OnRangeCaseClause( Node : TvbRangeCaseClause; Context : TObject = nil ); override;
          procedure OnCaseClause( Node : TvbCaseClause; Context : TObject = nil ); override;
          procedure OnSeekExpr( Node : TvbSeekExpr; Context : TObject = nil ); override;
          procedure OnSeekStmt( Node : TvbSeekStmt; Context : TObject = nil ); override; 
          procedure OnSgnExpr( Node : TvbSgnExpr; Context : TObject = nil ); override; 
          procedure OnSimpleName( Node : TvbSimpleName; Context : TObject = nil ); override; 
          procedure OnStdModuleDef( Node : TvbStdModuleDef; Context : TObject = nil ); override;
          procedure OnStmtBlock( Node : TvbStmtBlock; Context : TObject = nil ); override; 
          procedure OnStopStmt( Node : TvbStopStmt; Context : TObject = nil ); override; 
          procedure OnStringExpr( Node : TvbStringExpr; Context : TObject = nil ); override; 
          procedure OnTimeStmt( Node : TvbTimeStmt; Context : TObject = nil ); override;
          procedure OnIntExpr( Node : TvbIntExpr; Context : TObject = nil ); override;
          procedure OnTypeOfExpr( Node : TvbTypeOfExpr; Context : TObject = nil ); override; 
          procedure OnUboundExpr( Node : TvbUBoundExpr; Context : TObject = nil ); override; 
          procedure OnUnaryExpr( Node : TvbUnaryExpr; Context : TObject = nil ); override; 
          procedure OnUnlockStmt( Node : TvbUnlockStmt; Context : TObject = nil ); override; 
          procedure OnVarDef( Node : TvbVarDef; Context : TObject = nil ); override;
          procedure OnWidthStmt( Node : TvbWidthStmt; Context : TObject = nil ); override;
          procedure OnWithStmt( Node : TvbWithStmt; Context : TObject = nil ); override;
          procedure OnWriteStmt( Node : TvbWriteStmt; Context : TObject = nil ); override;
          procedure OnDateExpr( Node : TvbDateExpr; Context : TObject = nil ); override;
      end;

    TvbPass1_ResolveTypeRefs =
      class( TvbPrefixNodeVisitor )
        private
          fGlobalScope : TvbScope;
          fOptions : TvbOptions;
        public
          constructor Create( Options : TvbOptions );
          procedure OnClassDef( Node : TvbClassDef; Context : TObject = nil ); override;
          procedure OnFormDef( Node : TvbFormDef; Context : TObject = nil ); override;
          procedure OnClassModuleDef( Node : TvbClassModuleDef; Context : TObject = nil ); override;
          procedure OnConstDef( Node : TvbConstDef; Context : TObject = nil ); override;
          procedure OnDllFuncDef( Node : TvbDllFuncDef; Context : TObject = nil ); override;
          procedure OnEnumDef( Node : TvbEnumDef; Context : TObject = nil ); override;
          procedure OnEventDef( Node : TvbEventDef; Context : TObject = nil ); override;
          procedure OnFunctionDef( Node : TvbFuncDef; Context : TObject = nil ); override;
          procedure OnFuncResultExpr( Node : TvbFuncResultExpr; Context : TObject = nil ); override;
          procedure OnLibModuleDef( Node : TvbLibModuleDef; Context : TObject = nil ); override;
          procedure OnParamDef( Node : TvbParamDef; Context : TObject = nil ); override;
          procedure OnProjectDef( Node : TvbProject; Context : TObject = nil ); override; 
          procedure OnPropertyDef( Node : TvbPropertyDef; Context : TObject = nil ); override;
          procedure OnQualifiedName( Node : TvbQualifiedName; Context : TObject = nil ); override;
          procedure OnRecordDef( Node : TvbRecordDef; Context : TObject = nil ); override;
          procedure OnSimpleName( Node : TvbSimpleName; Context : TObject = nil ); override;
          procedure OnStdModuleDef( Node : TvbStdModuleDef; Context : TObject = nil ); override;
          procedure OnTypeOfExpr( Node : TvbTypeOfExpr; Context : TObject = nil ); override;
          procedure OnVarDef( Node : TvbVarDef; Context : TObject = nil ); override;
      end;

    TvbPass2_ResolveExprs =
      class( TvbPrefixNodeVisitor )
        private
          fGlobalScope       : TvbScope;
          fCurrentSelectCase : TvbSelectCaseStmt;
        public
          //procedure OnTypeRef( Node : Tvbtype; Context : TObject = nil ); reintroduce; virtual; 
          procedure OnAbsExpr( Node : TvbAbsExpr; Context : TObject = nil ); override;
          procedure OnAddressOfExpr( Node : TvbAddressOfExpr; Context : TObject = nil ); override;
          procedure OnArrayExpr( Node : TvbArrayExpr; Context : TObject = nil ); override;
          procedure OnAssignStmt( Node : TvbAssignStmt; Context : TObject ); override;
          procedure OnBinaryExpr( Node : TvbBinaryExpr; Context : TObject = nil ); override;
          procedure OnCallOrIndexerExpr( Node : TvbCallOrIndexerExpr; Context : TObject = nil ); override;
          procedure OnCastExpr( Node : TvbCastExpr; Context : TObject = nil ); override;
          procedure OnCircleStmt( Node : TvbCircleStmt; Context : TObject = nil ); override;
          procedure OnConstDef( Node : TvbConstDef; Context : TObject = nil ); override;
          procedure OnDateLitExpr( Node : TvbDateLit; Context : TObject = nil ); override;
          procedure OnDictAccessExpr( Node : TvbDictAccessExpr; Context : TObject = nil ); override;
          procedure OnFixExpr( Node : TvbFixExpr; Context : TObject = nil ); override;
          procedure OnFloatLitExpr( Node : TvbFloatLit; Context : TObject = nil ); override;
          procedure OnForStmt( Node : TvbForStmt; Context : TObject = nil ); override;
          procedure OnFuncResultExpr( Node : TvbFuncResultExpr; Context : TObject = nil ); override;
          procedure OnInputExpr( Node : TvbInputExpr; Context : TObject = nil ); override;
          procedure OnLabelDef( Node : TvbLabelDef; Context : TObject = nil ); override;
          procedure OnLabelDefStmt( Node : TvbLabelDefStmt; Context : TObject = nil ); override;
          //procedure OnLabelRef( Node : tvblabel; Context : TObject = nil ); override;
          procedure OnLBoundExpr( Node : TvbLBoundExpr; Context : TObject = nil ); override;
          procedure OnLenBExpr( Node : TvbLenBExpr; Context : TObject = nil ); override;
          procedure OnLenExpr( Node : TvbLenExpr; Context : TObject = nil ); override;
          procedure OnLineStmt( Node : TvbLineStmt; Context : TObject = nil ); override;
          procedure OnMeExpr( Node : TvbMeExpr; Context : TObject = nil ); override;
          procedure OnMemberAccessExpr( Node : TvbMemberAccessExpr; Context : TObject = nil ); override;
          procedure OnMidExpr( Node : TvbMidExpr; Context : TObject = nil ); override;
          procedure OnNameExpr( Node : TvbNameExpr; Context : TObject = nil ); override;
          procedure OnNewExpr( Node : TvbNewExpr; Context : TObject = nil ); override;
          procedure OnNothingExpr( Node : TvbNothingLit; Context : TObject = nil ); override;
          procedure OnParamDef( Node : TvbParamDef; Context : TObject = nil ); override;
          procedure OnPrintArgExpr( Node : TvbPrintStmtExpr; Context : TObject = nil ); override;
          procedure OnProjectDef( Node : TvbProject; Context : TObject ); override;
          procedure OnPSetStmt( Node : TvbPSetStmt; Context : TObject = nil ); override;
          procedure OnQualifiedName( Node : TvbQualifiedName; Context : TObject = nil ); override;
          procedure OnResumeStmt( Node : TvbResumeStmt; Context : TObject = nil ); override;
          procedure OnSelectCaseStmt( Node : TvbSelectCaseStmt; Context : TObject = nil ); override;
          procedure OnCaseStmt( Node : TvbCaseStmt; Context : TObject = nil ); override;
          procedure OnRelationalCaseClause( Node : TvbRelationalCaseClause; Context : TObject = nil ); override;
          procedure OnRangeCaseClause( Node : TvbRangeCaseClause; Context : TObject = nil ); override;
          procedure OnCaseClause( Node : TvbCaseClause; Context : TObject = nil ); override;
          procedure OnSeekExpr( Node : TvbSeekExpr; Context : TObject = nil ); override;
          procedure OnSgnExpr( Node : TvbSgnExpr; Context : TObject = nil ); override;
          procedure OnSimpleName( Node : TvbSimpleName; Context : TObject = nil ); override;
          procedure OnStmtBlock( Node : TvbStmtBlock; Context : TObject = nil ); override;
          procedure OnStringExpr( Node : TvbStringExpr; Context : TObject = nil ); override;
          procedure OnIntegerLitExpr( Node : TvbIntLit; Context : TObject = nil ); override;
          procedure OnIntExpr( Node : TvbIntExpr; Context : TObject = nil ); override;
          procedure OnTypeOfExpr( Node : TvbTypeOfExpr; Context : TObject = nil ); override; 
          procedure OnUboundExpr( Node : TvbUBoundExpr; Context : TObject = nil ); override; 
          procedure OnUnaryExpr( Node : TvbUnaryExpr; Context : TObject = nil ); override; 
          procedure OnWithStmt( Node : TvbWithStmt; Context : TObject = nil ); override;
          procedure OnDateExpr( Node : TvbDateExpr; Context : TObject = nil ); override;
      end;
    
  procedure ResolveProject( P : TvbProject; Options : TvbOptions );
    var
      Pass1 : TvbPass1_ResolveTypeRefs;
      Pass2 : TvbPass2_ResolveExprs;
    begin
      assert( P <> nil );
      Pass1 := nil;
      Pass2 := nil;
      try
        Pass1 := TvbPass1_ResolveTypeRefs.Create( Options );
        Pass2 := TvbPass2_ResolveExprs.Create;
        P.Accept( Pass1 );
        P.Accept( Pass2 );
      finally
        Pass1.Free;
        Pass2.Free;
      end;
    end;

  procedure ResolveTypeName( T : IvbType; CurrentScope : TvbScope );
    var
      def             : TvbDef;
      SearchInParents : boolean;
      i, start        : integer;
      typeName, n1    : string;
    begin
      if ( T.TypeDef <> nil ) or ( T.Name = '' )
        then exit; // type already resolved or not an UDT type

      SearchInParents := true;
      typeName := T.Name;
      start := 1;
      n1 := '';  // mapped name
      
      repeat
        assert( CurrentScope <> nil );
        // Search for a '.' in the type name.
        // If there is one, the type is qualified by the project name or
        // a module name or both.
        i := PosEx( '.', typeName, start );
        if i <> 0 
          then
            begin
              // Search for the qualifier name in the current scope and if found
              // set the current scope to the qualifier's scope.
              if CurrentScope.Lookup( PROJECT_OR_MODULE_BINDING, copy( typeName, start, i - 1 ), def, SearchInParents )
                then
                  begin
                    if map_Full in def.NodeFlags
                      then n1 := def.Name1 + '.'
                      else n1 := n1 + def.Name1 + '.';
                    // Get the qualifier's scope.
                    case def.NodeKind of
                      STD_MODULE_DEF,
                      CLASS_MODULE_DEF : CurrentScope := TvbModuleDef( def ).Scope;
                      PROJECT_DEF      : CurrentScope := TvbProject( def ).Scope;
                    end;
                    SearchInParents := false;
                  end
                else break;
            end
          else
            begin
              // At this point there are no more '.' in the type name.
              // So now we reached the actual name of the type.
              // Search the type name in the current scope, which should be
              // the qualifier's scope or the passed scope if this is an
              // unqualified type.
              if CurrentScope.Lookup( TYPE_BINDING, copy( typeName, start, length( typeName ) ), def, SearchInParents )
                then
                  begin
                    T.TypeDef := TvbTypeDef( def );
                    if map_Full in def.NodeFlags
                      then n1 := def.Name1
                      else n1 := n1 + def.Name1;
                    T.Name1 := n1;
                  end;
              break;
            end;
        start := i + 1;
      until i = 0;
    end;

  procedure ResolveTypeExprs( const T : IvbType; Visitor : TvbNodeVisitor; Context : TObject );
    var
      i      : integer;
      Bounds : TvbArrayBounds;
    begin
      // Resolve the array bounds if any.
      for i := 0 to T.ArrayBoundCount - 1 do
        begin
          Bounds := T.ArrayBounds[i];
          if Bounds.Lower <> nil
            then Bounds.Lower.Accept( Visitor, Context );
          assert( Bounds.Upper <> nil );
          Bounds.Upper.Accept( Visitor, Context );
        end;
      // Resolve the string size.
      if T.StringLen  <> nil
        then T.StringLen.Accept( Visitor, Context );        
    end;

  // Returns the common integral type of two types.
  // Assumes the 2 types are integral.
  function GetCommonIntegralType( const T1, T2 : IvbType ) : IvbType;
    begin
      if ( T1 = vbLongType ) or ( T2 = vbLongType )
        then result := vbLongType
        else
      if ( ( T1 = vbBooleanType ) or ( T1 = vbIntegerType ) ) or
         ( ( T2 = vbBooleanType ) or ( T2 = vbIntegerType ) )
        then result := vbIntegerType
        else
      if ( T1 = vbByteType ) or ( T2 = vbByteType )
        then result := vbByteType
        else result := nil;
    end;

  function GetCommonArithmeticType( const T1, T2 : IvbType ) : IvbType;
    begin
      // Note that VB6 prefers CURRENCY to DOUBLE in addition and
      // substraction expressions.
      if ( T1 = vbCurrencyType ) or ( T2 = vbCurrencyType )
        then result := vbCurrencyType
        else
      // Remember that when performing simple arithmetic VB6 treats
      // STRING as DOUBLE.
      if ( ( T1 = vbStringType ) or ( T1 = vbDoubleType ) ) or
         ( ( T2 = vbStringType ) or ( T2 = vbDoubleType ) )
        then result := vbDoubleType
        else
      if ( ( T1 = vbSingleType ) and ( T2 = vbLongType ) ) or
         ( ( T2 = vbSingleType ) and ( T1 = vbLongType ) )
        then result := vbDoubleType
        else
      if ( T1 = vbSingleType ) or ( T2 = vbSingleType )
        then result := vbSingleType
        else result := GetCommonIntegralType( T1, T2 );
    end;

  // Given a binary operator and the type its two operands determines the
  // data type returned by the expression.
  // Returns NIL is the data type cannot be determined.
  function GetBinaryExpressionType( op : TvbOperator; const T1, T2 : IvbType ) : IvbType;
    begin
      result := nil;

      // If at least one type is VARIANT, the expression type is VARIANT.
      if ( T1 = vbVariantType ) or ( T2 = vbVariantType )
        then
          begin
            result := vbVariantType;
            exit;
          end;
      if ( T1 = nil ) or ( T2 = nil )
        then exit;

      case op of
        ADD_OP :
          // If both operands are STRING, result is STRING(concatenation).
          // Note that because we need to take into account both dynamic
          // and fixed-length string, we can't check for equality with
          // vbStringType.
          if ( T1.Code = STRING_TYPE ) and ( T2.Code = STRING_TYPE )
            then result := vbStringType
            else
          // If one operator is DATE, result is DATE.
          if ( T1 = vbDateType ) or ( T2 = vbDateType )
            then result := vbDateType
            else result := GetCommonArithmeticType( T1, T2 );          

        DIV_OP :
          // If both operands are SINGLE, result is also SINGLE.
          // Otherwise the result is always DOUBLE.
          if ( T1 = vbSingleType ) and ( T2 = vbSingleType )
            then result := vbSingleType
            else result := vbDoubleType;

        IS_OP,
        EQ_OP,
        GREATER_EQ_OP,
        GREATER_OP,
        NOT_EQ_OP,
        LESS_EQ_OP,
        LESS_OP,
        LIKE_OP : result := vbBooleanType;
        MOD_OP,
        INT_DIV_OP :
          // If any operand is of floating type, VB6 coerces it to LONG
          // and the result type is LONG.
          if IsFloatingType( T1 ) or IsFloatingType( T2 )
            then result := vbLongType
            else result := GetCommonIntegralType( T1, T2 );
        MULT_OP :
          // Note that VB6 prefers DOUBLE to CURRENCY when multiplying.
          // When adding or substracting, the inverse is prefered.
          if ( ( T1 = vbDateType ) or    // treated as DOUBLE
               ( T1 = vbStringType ) or  // treated as DOUBLE
               ( T1 = vbDoubleType ) ) or
             ( ( T2 = vbDateType ) or
               ( T2 = vbStringType ) or
               ( T2 = vbDoubleType ) )
            then result := vbDoubleType
            else
          if ( T1 = vbCurrencyType ) or ( T2 = vbCurrencyType )
            then result := vbCurrencyType
            else
          if ( ( T1 = vbSingleType ) and ( T2 = vbCurrencyType ) ) or
             ( ( T2 = vbSingleType ) and ( T1 = vbCurrencyType ) )
            then result := vbDoubleType
            else
          if ( ( T1 = vbSingleType ) and ( T2 = vbLongType ) ) or
             ( ( T2 = vbSingleType ) and ( T1 = vbLongType ) )
            then result := vbDoubleType
            else
          if ( T1 = vbSingleType ) or ( T2 = vbSingleType )
            then result := vbSingleType
            else result := GetCommonIntegralType( T1, T2 );

        EQV_OP,
        IMP_OP,
        AND_OP,
        OR_OP,
        XOR_OP :
          // With BOOLEAN operands they operate as logical operators
          // and the result type is BOOLEAN.
          // With other operand types(numeric) they perform bitwise
          // operations.
          if ( T1 = vbBooleanType ) and ( T2 = vbBooleanType )
            then result := vbBooleanType
            else result := GetCommonIntegralType( T1, T2 );
        POWER_OP :
          // Exponentiation always returns DOUBLE.
          result := vbDoubleType;
        STR_CONCAT_OP :
          // STRING concatenation always returns STRING.
          result := vbStringType;
        SUB_OP :
          // If one operand is DATE, result is DATE.
          if ( T1 = vbDateType ) or ( T2 = vbDateType )
            then result := vbDateType
            else result := GetCommonArithmeticType( T1, T2 );
      end;
    end;

  function GetUnaryExpressionType( op : TvbOperator; const T : IvbType ): IvbType;
    begin
      result := nil;
      if T = vbVariantType
        then
          begin
            result := vbVariantType;
            exit;
          end;
      if T = nil then exit;
      case op of
        NOT_OP :
          if T = vbBooleanType
            then result := vbBooleanType
            else
          if IsFloatingType( T )
            then result := vbLongType
            else result := T;
        SUB_OP :
          if ( T = vbByteType ) or ( T = vbBooleanType )
            then result := vbIntegerType
            else result := T;
        ADD_OP : result := T;
      end;
    end;

  { TvbPrefixNodeVisitor }

  procedure TvbPrefixNodeVisitor.OnAbsExpr( Node : TvbAbsExpr; Context : TObject );
    begin
      with Node do
        begin
          assert( Number <> nil );
          Number.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnAddressOfExpr( Node : TvbAddressOfExpr; Context : TObject );
    begin
      with Node do
        begin
          assert( FuncName <> nil );
          FuncName.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnArrayExpr( Node : TvbArrayExpr; Context : TObject );
    var
      i : integer;
    begin
      with Node do
        begin
          for i := 0 to ValueCount - 1 do
            Values[i].Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnAssignStmt( Node : TvbAssignStmt; Context : TObject );
    begin
      with Node do
        begin
          assert( Value <> nil );
          assert( LHS <> nil );
          Value.Accept( self, Context );
          LHS.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnBinaryExpr( Node : TvbBinaryExpr; Context : TObject );
    begin
      with Node do
        begin
          assert( RightExpr <> nil );
          assert( LeftExpr <> nil );
          RightExpr.Accept( self, Context );
          LeftExpr.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnCallOrIndexerExpr( Node : TvbCallOrIndexerExpr; Context : TObject );
    begin
      with Node do
        begin
          assert( FuncOrVar <> nil );
          FuncOrVar.Accept( self, Context );
          if Args <> nil
            then Args.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnCallStmt( Node : TvbCallStmt; Context : TObject );
    begin
      with Node do
        begin
          assert( Method <> nil );
          Method.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnCastExpr( Node : TvbCastExpr; Context : TObject );
    begin
      with Node do
        begin
          assert( Op1 <> nil );
          Op1.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnCircleStmt( Node : TvbCircleStmt; Context : TObject );
    begin
      with Node do
        begin
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnClassDef( Node : TvbClassDef; Context : TObject );
    var
      i : integer;
    begin                
      with Node do
        begin
          for i := 0 to EventCount - 1 do
            Event[i].Accept( self, Context );
          for i := 0 to MethodCount - 1 do
            Method[i].Accept( self, Context );
          for i := 0 to VarCount - 1 do
            Vars[i].Accept( self, Context );
          for i := 0 to PropCount - 1 do
            Prop[i].Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnClassModuleDef( Node : TvbClassModuleDef; Context : TObject );
    begin
      fCurrentModule := Node;
      with Node do
        begin
          assert( ClassDef <> nil );
          ClassDef.Accept( self, Scope );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnCloseStmt( Node : TvbCloseStmt; Context : TObject );
    begin
      with Node do
        begin
          if FileNum <> nil
            then FileNum.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnConstDef( Node : TvbConstDef; Context : TObject );
    begin
      with Node do
        begin
          if Value <> nil
            then Value.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnDateExpr( Node : TvbDateExpr; Context : TObject );
    begin
      with Node do
        begin          
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnDateLitExpr( Node : TvbDateLit; Context : TObject );
    begin
      with Node do
        begin
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnDateStmt( Node : TvbDateStmt; Context : TObject );
    begin
      with Node do
        begin
          assert( NewDate <> nil );
          NewDate.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnDebugAssertStmt( Node : TvbDebugAssertStmt; Context : TObject );
    begin
      with Node do
        begin
          assert( BoolExpr <> nil );
          BoolExpr.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnDebugPrintStmt( Node : TvbDebugPrintStmt; Context : TObject );
    begin
      with Node do
        begin
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnDictAccessExpr( Node : TvbDictAccessExpr; Context : TObject );
    begin
      with Node do
        begin
          // Target may be NIL when this expression is used inside a
          // WITH statement and refers to the WITTH's expression.
          if Target <> nil
            then Target.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnDllFuncDef( Node : TvbDllFuncDef; Context : TObject );
    var
      i : integer;
    begin
      with Node do
        for i := 0 to ParamCount - 1 do
          Param[i].Accept( self, Context );
    end;

  procedure TvbPrefixNodeVisitor.OnDoEventsExpr( Node : TvbDoEventsExpr; Context : TObject );
    begin
      with Node do
        begin          
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnDoLoopStmt( Node : TvbDoLoopStmt; Context : TObject );
    begin
      with Node do
        begin
          assert( Condition <> nil );
          Condition.Accept( self, Context );
          if StmtBlock <> nil
            then StmtBlock.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnEndStmt( Node : TvbEndStmt; Context : TObject );
    begin
    end;

  procedure TvbPrefixNodeVisitor.OnEnumDef( Node : TvbEnumDef; Context : TObject );
    var
      i : integer;
    begin
      with Node do
        for i := 0 to EnumCount - 1 do
          Enum[i].Accept( self, Context );
    end;

  procedure TvbPrefixNodeVisitor.OnEraseStmt( Node : TvbEraseStmt; Context : TObject );
    begin
      with Node do
        begin
          assert( ArrayVar <> nil );
          ArrayVar.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnErrorStmt( Node : TvbErrorStmt; Context : TObject );
    begin
      with Node do
        begin
          assert( ErrorNumber <> nil );
          ErrorNumber.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnEventDef( Node : TvbEventDef; Context : TObject );
    var
      i : integer;
    begin
      with Node do
        for i := 0 to ParamCount - 1 do
          Param[i].Accept( self, Context );
    end;

  procedure TvbPrefixNodeVisitor.OnExitStmt( Node : TvbExitStmt; Context : TObject );
    begin
      with Node do
        begin
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnFixExpr( Node : TvbFixExpr; Context : TObject );
    begin
      with Node do
        begin
          assert( Number <> nil );
          Number.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnFloatLitExpr( Node : TvbFloatLit; Context : TObject );
    begin
      with Node do
        begin
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnForeachStmt( Node : TvbForeachStmt; Context : TObject );
    begin
      with Node do
        begin
          Element.Accept( self, Context );
          Group.Accept( self, Context );
          if StmtBlock <> nil
            then StmtBlock.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnForStmt( Node : TvbForStmt; Context : TObject );
    begin
      with Node do
        begin
          Counter.Accept( self, Context );
          InitValue.Accept( self, Context );
          FinalValue.Accept( self, Context );
          if StepValue <> nil
            then StepValue.Accept( self, Context );
          if StmtBlock <> nil
            then StmtBlock.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnFunctionDef( Node : TvbFuncDef; Context : TObject );
    var
      i : integer;
      funcScope : TvbScope;
    begin
      with Node do
        begin
          // Note that the parameters are resolved in the parent module's scope,
          // so they won't see any name declared inside the function.
          for i := 0 to ParamCount - 1 do
            Param[i].Accept( self, Context );
          funcScope := Scope; // the function scope
          if funcScope <> nil
            then
              begin
                // Instead, local constants, variables and of course statements are
                // resolved in the function's local scope.
                for i := 0 to ConstCount - 1 do
                  Consts[i].Accept( self, funcScope );
                for i := 0 to VarCount - 1 do
                  Vars[i].Accept( self, funcScope );
                if StmtBlock <> nil
                  then StmtBlock.Accept( self, funcScope );
              end;
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnFuncResultExpr( Node : TvbFuncResultExpr; Context : TObject );
    begin
      with Node do
        begin
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnGetStmt( Node : TvbGetStmt; Context : TObject );
    begin
      with Node do
        begin
          assert( FileNum <> nil );
          FileNum.Accept( self, Context );
          if RecNum <> nil
            then RecNum.Accept( self, Context );
          assert( VarName <> nil );
          VarName.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnGotoOrGosubStmt( Node : TvbGotoOrGosubStmt; Context : TObject );
    begin
      with Node do
        begin
          assert( LabelName <> nil );
          LabelName.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnIfStmt( Node : TvbIfStmt; Context : TObject );
    var
      i : integer;
    begin
      with Node do
        begin
          Expr.Accept( self, Context );
          if ThenBlock <> nil
            then ThenBlock.Accept( self, Context );
          for i := 0 to ElseIfCount - 1 do
            ElseIf[i].Accept( self, Context );
          if ElseBlock <> nil
            then ElseBlock.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnInputExpr( Node : TvbInputExpr; Context : TObject );
    begin
      with Node do
        begin
          assert( FileNum <> nil );
          assert( CharCount <> nil );
          FileNum.Accept( self, Context );
          CharCount.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnInputStmt( Node : TvbInputStmt; Context : TObject );
    var
      i : integer;
    begin
      with Node do
        begin
          assert( FileNum <> nil );
          FileNum.Accept( self, Context );
          for i := 0 to VarCount - 1 do
            Vars[i].Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnLabelDef( Node : TvbLabelDef; Context : TObject );
    begin
      with Node do
        begin
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnLabelDefStmt( Node : TvbLabelDefStmt; Context : TObject );
    begin
      with Node do
        begin
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnLBoundExpr( Node : TvbLBoundExpr; Context : TObject );
    begin
      with Node do
        begin
          assert( ArrayVar <> nil );
          ArrayVar.Accept( self, Context );
          if Dim <> nil
            then Dim.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnLenBExpr( Node : TvbLenBExpr; Context : TObject );
    begin
      with Node do
        begin
          assert( StrOrVar <> nil );
          StrOrVar.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnLenExpr( Node : TvbLenExpr; Context : TObject );
    begin
      with Node do
        begin
          assert( StrOrVar <> nil );
          StrOrVar.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnLibModuleDef( Node : TvbLibModuleDef; Context : TObject );
    var
      i : integer;
    begin
      with Node do
        begin
          for i := 0 to ConstCount - 1 do
            Consts[i].Accept( self, Context );
          for i := 0 to FunctionCount - 1 do
            Functions[i].Accept( self, Context );
          for i := 0 to VarCount - 1 do
            Vars[i].Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnLineStmt( Node : TvbLineStmt; Context : TObject );
    begin
      with Node do
        begin
          if Color <> nil
            then Color.Accept( self, Context );
            //TODO            
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnLineInputStmt( Node : TvbLineInputStmt; Context : TObject );
    begin
      with Node do
        begin
          assert( FileNum <> nil );
          assert( VarName <> nil );
          FileNum.Accept( self, Context );
          VarName.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnLockStmt( Node : TvbLockStmt; Context : TObject );
    begin
      with Node do
        begin
          assert( FileNum <> nil );
          FileNum.Accept( self, Context );
          if StartRec <> nil
            then StartRec.Accept( self, Context );
          if EndRec <> nil
            then EndRec.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnMeExpr( Node : TvbMeExpr; Context : TObject );
    begin
    end;

  procedure TvbPrefixNodeVisitor.OnMemberAccessExpr( Node : TvbMemberAccessExpr; Context : TObject );
    begin
      with Node do
        begin
          if Target <> nil
            then Target.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnMidAssignStmt( Node : TvbMidAssignStmt; Context : TObject );
    begin
      with Node do
        begin
          assert( StartPos <> nil );
          assert( Replacement <> nil );
          assert( VarName <> nil );
          StartPos.Accept( self, Context );
          Replacement.Accept( self, Context );
          VarName.Accept( self, Context );
          if Length <> nil
            then Length.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnMidExpr( Node : TvbMidExpr; Context : TObject );
    begin
      with Node do
        begin
          assert( Str <> nil );
          assert( Start <> nil );
          Str.Accept( self, Context );
          Start.Accept( self, Context );
          if Len <> nil
            then Len.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnNameExpr( Node : TvbNameExpr; Context : TObject );
    begin
    end;

  procedure TvbPrefixNodeVisitor.OnNameStmt( Node : TvbNameStmt; Context : TObject );
    begin
      with Node do
        begin
          assert( NewPath <> nil );
          assert( OldPath <> nil );
          NewPath.Accept( self, Context );
          OldPath.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnNewExpr( Node : TvbNewExpr; Context : TObject );
    begin
      with Node do
        begin
          assert( TypeExpr <> nil );
          TypeExpr.Accept( self, Context );  
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnNothingExpr( Node : TvbNothingLit; Context : TObject );
    begin
    end;

  procedure TvbPrefixNodeVisitor.OnOnErrorStmt( Node : TvbOnErrorStmt; Context : TObject );
    begin
      with Node do
        begin
          if LabelName <> nil
            then LabelName.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnOnGotoOrGosubStmt( Node : TvbOnGotoOrGosubStmt; Context : TObject );
    var
      i : integer;
    begin
      with Node do
        begin
          Expr.Accept( self, Context );
          for i := 0 to LabelCount - 1 do
            LabelName[i].Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnOpenStmt( Node : TvbOpenStmt; Context : TObject );
    begin
      with Node do
        begin
          FileNum.Accept( self, Context );
          FilePath.Accept( self, Context );
          if RecLen <> nil
            then RecLen.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnParamDef( Node : TvbParamDef; Context : TObject );
    begin
      with Node do
        begin
          if DefaultValue <> nil
            then DefaultValue.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnPrintArgExpr( Node : TvbPrintStmtExpr; Context : TObject );
    begin
      with Node do
        begin
          if Expr <> nil
            then Expr.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnPrintStmt( Node : TvbPrintStmt; Context : TObject );
    begin
      with Node do
        begin
          assert( FileNum <> nil );
          FileNum.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnProjectDef( Node : TvbProject; Context : TObject );
    var
      i : integer;
    begin
      with Node do
        begin
          for i := 0 to ImportedModuleCount - 1 do
            ImportedModule[i].Accept( self, Context );
          for i := 0 to ModuleCount - 1 do
            Module[i].Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnPropertyDef( Node : TvbPropertyDef; Context : TObject );
    begin
      with Node do
        begin
          if GetFn <> nil
            then GetFn.Accept( self, Context );
          if SetFn <> nil
            then SetFn.Accept( self, Context );
          if LetFn <> nil
            then LetFn.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnPSetStmt( Node : TvbPSetStmt; Context : TObject );
    begin
      with Node do
        begin
          if Color <> nil
            then Color.Accept( self, Context );
          Y.Accept( self, Context );
          X.Accept( self, Context );
          if ObjectRef <> nil
            then ObjectRef.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnPutStmt( Node : TvbPutStmt; Context : TObject );
    begin
      with Node do
        begin
          VarName.Accept( self, Context );
          FileNum.Accept( self, Context );
          if RecNum <> nil
            then RecNum.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnQualifiedName( Node : TvbQualifiedName; Context : TObject );
    begin
      with Node do
        begin
          Qualifier.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnRaiseEventStmt( Node : TvbRaiseEventStmt; Context : TObject );
    begin
      with Node do
        begin
          if EventArgs <> nil
            then EventArgs.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnRecordDef( Node : TvbRecordDef; Context : TObject );
    var
      i : integer;
    begin
      with Node do
        for i := 0 to FieldCount - 1 do
          Field[i].Accept( self, Context );
    end;

  procedure TvbPrefixNodeVisitor.OnReDimStmt( Node : TvbReDimStmt; Context : TObject );
    begin
      with Node do
        begin
          ArrayVar.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnResumeStmt( Node : TvbResumeStmt; Context : TObject );
    begin
      with Node do
        begin
          if LabelName <> nil
            then LabelName.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnReturnStmt( Node : TvbReturnStmt; Context : TObject );
    begin
    end;

  procedure TvbPrefixNodeVisitor.OnSeekExpr( Node : TvbSeekExpr; Context : TObject );
    begin
      with Node do
        begin
          FileNum.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnSeekStmt( Node : TvbSeekStmt; Context : TObject );
    begin
      with Node do
        begin
          FileNum.Accept( self, Context );
          Pos.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnSgnExpr( Node : TvbSgnExpr; Context : TObject );
    begin
      with Node do
        begin
          assert( Number <> nil );
          Number.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnSimpleName( Node : TvbSimpleName; Context : TObject );
    begin
    end;

  procedure TvbPrefixNodeVisitor.OnStdModuleDef( Node : TvbStdModuleDef; Context : TObject );
    var
      i : integer;
      ModuleScope : TvbScope;
    begin
      fCurrentModule := Node;
      with Node do
        begin
          ModuleScope := Scope;
          assert( ModuleScope <> nil );
          for i := 0 to ConstCount - 1 do
            Consts[i].Accept( self, ModuleScope );
          for i := 0 to EnumCount - 1 do
            Enum[i].Accept( self, ModuleScope );
          for i := 0 to RecordCount - 1 do
            Records[i].Accept( self, ModuleScope );
          for i := 0 to ClassCount - 1 do
            Classes[i].Accept( self, ModuleScope );
          for i := 0 to DllFuncCount - 1 do
            DllFuncs[i].Accept( self, ModuleScope );
          for i := 0 to LibModuleCount - 1 do
            LibModule[i].Accept( self, ModuleScope );
          for i := 0 to MethodCount - 1 do
            Method[i].Accept( self, ModuleScope );
          for i := 0 to PropCount - 1 do
            Prop[i].Accept( self, ModuleScope );
          for i := 0 to VarCount - 1 do
            Vars[i].Accept( self, ModuleScope );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnStmtBlock( Node : TvbStmtBlock; Context : TObject );
    var
      stmt : TvbStmt;
    begin
      with Node do
        begin
          stmt := FirstStmt;
          while stmt <> nil do
            begin
              stmt.Accept( self, Context );
              stmt := stmt.NextStmt;
            end;    
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnStopStmt( Node : TvbStopStmt; Context : TObject );
    begin
    end;

  procedure TvbPrefixNodeVisitor.OnStringExpr( Node : TvbStringExpr; Context : TObject );
    begin
      with Node do
        begin
          Char.Accept( self, Context );
          Count.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnTimeStmt( Node : TvbTimeStmt; Context : TObject );
    begin
      with Node do
        begin
          NewTime.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnIntExpr( Node : TvbIntExpr; Context : TObject );
    begin
      with Node do
        begin
          assert( Number <> nil );
          Number.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnTypeOfExpr( Node : TvbTypeOfExpr; Context : TObject );
    begin
      with Node do
        begin
          Name.Accept( self, Context );                          
          ObjectExpr.Accept( self, Context );  
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnUboundExpr( Node : TvbUBoundExpr; Context : TObject );
    begin
      with Node do
        begin
          ArrayVar.Accept( self, Context );
          if Dim <> nil
            then Dim.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnUnaryExpr( Node : TvbUnaryExpr; Context : TObject );
    begin
      with Node do
        begin
          Expr.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnUnlockStmt( Node : TvbUnlockStmt; Context : TObject );
    begin
      with Node do
        begin
          FileNum.Accept( self, Context );
          if EndRec <> nil
            then EndRec.Accept( self, Context );
          if StartRec <> nil
            then StartRec.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnVarDef( Node : TvbVarDef; Context : TObject );
    begin
    end;

  procedure TvbPrefixNodeVisitor.OnWidthStmt( Node : TvbWidthStmt; Context : TObject );
    begin
      with Node do
        begin
          FileNum.Accept( self, Context );
          Width.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnWithStmt( Node : TvbWithStmt; Context : TObject );
    begin
      with Node do
        begin
          ObjOrRec.Accept( self, Context );
          if StmtBlock <> nil
            then StmtBlock.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnWriteStmt( Node : TvbWriteStmt; Context : TObject );
    var
      i : integer;
    begin
      with Node do
        begin
          FileNum.Accept( self, Context );
          for i := 0 to ValueCount - 1 do
            Value[i].Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnCaseClause( Node : TvbCaseClause; Context : TObject );
    begin
      with Node do
        begin
          if Expr <> nil
            then Expr.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnCaseStmt( Node : TvbCaseStmt; Context : TObject );
    var
      i : integer;
    begin
      with Node do
        begin
          for i := 0 to ClauseCount - 1 do
            Clause[i].Accept( self, Context );
          if StmtBlock <> nil
            then StmtBlock.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnRangeCaseClause( Node : TvbRangeCaseClause; Context : TObject );
    begin
      with Node do
        begin
          if Expr <> nil
            then Expr.Accept( self, Context );
          if UpperExpr <> nil
            then UpperExpr.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnRelationalCaseClause( Node : TvbRelationalCaseClause; Context : TObject );
    begin
      with Node do
        begin
          if Expr <> nil
            then Expr.Accept( self, Context );
        end;
    end;

  procedure TvbPrefixNodeVisitor.OnSelectCaseStmt( Node : TvbSelectCaseStmt; Context : TObject );
    var
      i : integer;
    begin
      with Node do
        begin
          assert( Expr <> nil );
          Expr.Accept( self, Context );
          for i := 0 to CaseCount - 1 do
            Cases[i].Accept( self, Context );
          if CaseElse <> nil
            then CaseElse.Accept( self, Context );
        end;    
    end;

  procedure TvbPrefixNodeVisitor.OnFormModuleDef( Node : TvbFormModuleDef; Context : TObject );
    begin
      //writeln( Node.name );
      OnClassModuleDef( Node, Context );
    end;

  procedure TvbPrefixNodeVisitor.OnFormDef( Node : TvbFormDef; Context : TObject );
    var
      i : integer;
    begin
      OnClassDef( Node, Context );
      for i := 0 to Node.ControlCount - 1 do
        Node.Control[i].Accept( self, Context );
    end;

  { TvbPass1_ResolveTypeRefs }

  procedure TvbPass1_ResolveTypeRefs.OnClassDef( Node : TvbClassDef; Context : TObject );
    begin
      assert( ( Context <> nil ) and ( Context is TvbScope ) );
      with Node do
        begin
          // make the class visible in the global scope only if it's public.
          if ( fGlobalScope <> nil ) and ( dfPublic in NodeFlags )
            then fGlobalScope.Bind( Node, TYPE_BINDING );

          // if the class it not external, format it's name 'a la Delphi".
          if not ( dfExternal in NodeFlags )
            then Name1 := Format( fmtDelphiClassType, [Name] );

          // if we've got a base class resolve it's type name.
          if BaseClass <> nil
            then
              begin
                ResolveTypeName( BaseClass, TvbScope( Context ) );
                // if we really have a class as base type, set the its scope
                // as the scope of this class.
                if ( BaseClass.TypeDef <> nil ) and
                   ( BaseClass.TypeDef.NodeKind = CLASS_DEF )
                  then Scope.BaseScope := TvbClassDef( BaseClass.TypeDef ).Scope;
              end;
        end;
      inherited;
    end;

  procedure TvbPass1_ResolveTypeRefs.OnClassModuleDef( Node : TvbClassModuleDef; Context : TObject );
    begin
      inherited;
    end;

  procedure TvbPass1_ResolveTypeRefs.OnConstDef( Node : TvbConstDef; Context : TObject );
    begin
      inherited;
      with Node do
        begin
          if ConstType <> nil
            then ResolveTypeName( ConstType, TvbScope( Context ) );
        end;
    end;

  procedure TvbPass1_ResolveTypeRefs.OnDllFuncDef( Node : TvbDllFuncDef; Context : TObject );
    begin
      inherited;
      with Node do
        begin
          if ReturnType <> nil
            then ResolveTypeName( ReturnType, TvbScope( Context ) );
        end;
    end;

  procedure TvbPass1_ResolveTypeRefs.OnEnumDef( Node : TvbEnumDef; Context : TObject );
    var
      i               : integer;
      ContainingScope : TvbScope;
      cst             : TvbConstDef;
      isPublic        : boolean;
    begin
      ContainingScope := TvbScope( Context );
      assert( ContainingScope <> nil );
      if fGlobalScope <> nil
        then
          begin
            isPublic := dfPublic in Node.NodeFlags;
            if isPublic
              then fGlobalScope.Bind( Node, TYPE_BINDING );
            with Node do
              for i := 0 to EnumCount - 1 do
                begin
                  cst := Enum[i];
                  // if there is at least one constant without value flag
                  // this enum type as having missing values.
                  // this flag prevents the enum from being converted to single
                  // integer constants.
                  if cst.Value = nil
                    then NodeFlags := NodeFlags + [enum_MissingValues];

                  // Enumerated constanst are always visible in their parent
                  // module's scope, whether they are public or not.
                  ContainingScope.Bind( cst, OBJECT_BINDING );
                  // However, they are visible in the global scope only if
                  // the enum type is public.
                  if isPublic
                    then fGlobalScope.Bind( cst, OBJECT_BINDING );
                end;
          end;
      with Node do
        if not ( dfExternal in NodeFlags )
          then Name1 := Format( fmtDelphiEnumType, [Name] );
      inherited;
    end;

  procedure TvbPass1_ResolveTypeRefs.OnEventDef( Node : TvbEventDef; Context : TObject );
    begin
      inherited;

    end;

  procedure TvbPass1_ResolveTypeRefs.OnFunctionDef( Node : TvbFuncDef; Context : TObject );
    begin
      inherited;
      with Node do
        begin
          if ReturnType <> nil
            then ResolveTypeName( ReturnType, TvbScope( Context ) );
        end;
    end;

  procedure TvbPass1_ResolveTypeRefs.OnFuncResultExpr( Node : TvbFuncResultExpr; Context : TObject );
    begin
      inherited;
      with Node do
        begin
        end;
    end;

  procedure TvbPass1_ResolveTypeRefs.OnLibModuleDef( Node : TvbLibModuleDef; Context : TObject );
    var
      i           : integer;
      ParentScope : TvbScope;
      cst         : TvbConstDef;
      func        : TvbAbstractFuncDef;
      p           : TvbPropertyDef;
    begin
      ParentScope := TvbScope( Context ); // scope of containing module
      assert( ParentScope <> nil );
      if fGlobalScope <> nil
        then
          with Node do
            begin
              // A library module comes from an imported library and they
              // are always public and so we make them visible in the
              // global scope.
              fGlobalScope.Bind( Node, TYPE_BINDING );
              // Declare constants in the parent module and the global scope.
              for i := 0 to ConstCount - 1 do
                begin
                  cst := Consts[i];
                  ParentScope.Bind( cst, OBJECT_BINDING );
                  fGlobalScope.Bind( cst, OBJECT_BINDING );
                end;
              // Declare functions in the parent module and the global scope.
              for i := 0 to FunctionCount - 1 do
                begin
                  func := Functions[i];
                  ParentScope.Bind( func, OBJECT_BINDING );
                  fGlobalScope.Bind( func, OBJECT_BINDING );
                end;
              // Declare properties in the parent module and global scope;
              for i := 0 to PropCount - 1 do
                begin
                  p := Prop[i];
                  ParentScope.Bind( p, OBJECT_BINDING );
                  fGlobalScope.Bind( p, OBJECT_BINDING );
                end;
            end;
      inherited;
    end;

  procedure TvbPass1_ResolveTypeRefs.OnParamDef( Node : TvbParamDef; Context : TObject );
    begin
      inherited;
      with Node do
        begin
          if ParamType <> nil
            then ResolveTypeName( ParamType, TvbScope( Context ) );
        end;
    end;

  procedure TvbPass1_ResolveTypeRefs.OnProjectDef( Node : TvbProject; Context : TObject );
    begin
      fGlobalScope := Node.Scope;
      assert( fGlobalScope <> nil );
      inherited OnProjectDef( Node, fGlobalScope );
      with Node do
        begin
        end;
    end;

  procedure TvbPass1_ResolveTypeRefs.OnPropertyDef( Node : TvbPropertyDef; Context : TObject );
    begin
      inherited;
      with Node do
        begin
          if ( SetFn <> nil ) and not ( dfExternal in SetFn.NodeFlags )
            then SetFn.Name1 := Format( fmtDelphiPropertySetter, [Name] );
          if ( GetFn <> nil ) and not ( dfExternal in GetFn.NodeFlags )
            then GetFn.Name1 := Format( fmtDelphiPropertyGetter, [Name] );
          if ( LetFn <> nil ) and not ( dfExternal in LetFn.NodeFlags )
            then LetFn.Name1 := Format( fmtDelphiPropertyLetter, [Name] );
        end;
    end;

  procedure TvbPass1_ResolveTypeRefs.OnQualifiedName( Node : TvbQualifiedName; Context : TObject );
    begin
      inherited;
      with Node do
        begin
        end;
    end;

  procedure TvbPass1_ResolveTypeRefs.OnRecordDef( Node : TvbRecordDef; Context : TObject );
    begin
      // Make the record visible in the global scope only if it's public.
      if ( fGlobalScope <> nil ) and ( dfPublic in Node.NodeFlags )
        then fGlobalScope.Bind( Node, TYPE_BINDING );
      inherited;
      with Node do
        if not ( dfExternal in NodeFlags )
          then Name1 := Format( fmtDelphiRecordType, [Name] );
    end;         

  procedure TvbPass1_ResolveTypeRefs.OnSimpleName( Node : TvbSimpleName; Context : TObject );
    begin
      inherited;
      with Node do
        begin
        end;
    end;

  procedure TvbPass1_ResolveTypeRefs.OnStdModuleDef( Node : TvbStdModuleDef; Context : TObject );
    var
      i : integer;
      p : TvbPropertyDef;
    begin
      inherited;
      with Node do
        for i := 0 to PropCount - 1 do
          begin
            p := Prop[i];
            if p.SetFn <> nil
              then
                begin
                  //TODO: por que static?
                  p.SetFn.NodeFlags := p.SetFn.NodeFlags + [dfStatic];
                  if not ( dfExternal in p.SetFn.NodeFlags )
                    then p.SetFn.NodeFlags := p.SetFn.NodeFlags + [mapAsFunctionCall];
                end;
            if p.LetFn <> nil
              then
                begin
                  p.LetFn.NodeFlags := p.LetFn.NodeFlags + [dfStatic];
                  if not ( dfExternal in p.LetFn.NodeFlags )
                    then p.LetFn.NodeFlags := p.LetFn.NodeFlags + [mapAsFunctionCall];
                end;
            if p.GetFn <> nil
              then p.GetFn.NodeFlags := p.GetFn.NodeFlags + [dfStatic];
            if fGlobalScope <> nil
              then fGlobalScope.Bind( p, OBJECT_BINDING );
          end;
    end;

  procedure TvbPass1_ResolveTypeRefs.OnTypeOfExpr( Node : TvbTypeOfExpr; Context : TObject );
    begin
      inherited;
      with Node do
        begin
          Name.Accept( self, Context );
        end;
    end;

  procedure TvbPass1_ResolveTypeRefs.OnVarDef( Node : TvbVarDef; Context : TObject );
    begin
      inherited;
      with Node do
        begin
          if VarType <> nil
            then ResolveTypeName( VarType, TvbScope( Context ) );
        end;
    end;

  constructor TvbPass1_ResolveTypeRefs.Create( Options : TvbOptions );
    begin
      fOptions := Options;
    end;

  procedure TvbPass1_ResolveTypeRefs.OnFormDef( Node : TvbFormDef; Context : TObject );
    begin
      assert( Context is TvbScope );
      // a form inherits from the class of root control in the form definition
      // which should be "VB.Form" or "VB.MDIForm".
      Node.BaseClass := Node.RootControl.VarType;
      inherited;
    end;

  { TvbPass2_ResolveExprs }

  procedure TvbPass2_ResolveExprs.OnAbsExpr( Node : TvbAbsExpr; Context : TObject );
    begin
      inherited;
      with Node do
        begin
          if expr_Const in Number.NodeFlags
            then NodeFlags := NodeFlags + [expr_Const];            
          NodeType := vbLongType;
        end;
    end;

  procedure TvbPass2_ResolveExprs.OnAddressOfExpr( Node : TvbAddressOfExpr; Context : TObject );
    begin
      inherited;
      with Node do
        begin
          NodeFlags := NodeFlags + [expr_Const];
          NodeType  := vbLongType;
        end;
    end;

  procedure TvbPass2_ResolveExprs.OnArrayExpr( Node : TvbArrayExpr; Context : TObject );
    begin
      inherited;
      Node.NodeType := vbVariantType;
    end;

  procedure TvbPass2_ResolveExprs.OnAssignStmt( Node : TvbAssignStmt; Context : TObject );
    begin
      inherited;
    end;

  procedure TvbPass2_ResolveExprs.OnBinaryExpr( Node : TvbBinaryExpr; Context : TObject );
    begin
      inherited;
      with Node do
        begin
          if ( expr_Const in LeftExpr.NodeFlags ) and
             ( expr_Const in RightExpr.NodeFlags )
            then NodeFlags := NodeFlags + [expr_Const];
            
          NodeType := GetBinaryExpressionType( Oper, LeftExpr.NodeType, RightExpr.NodeType );
//          if ( Oper = STR_CONCAT_OP ) and
//             ( LeftExpr.NodeType <> nil ) and
//             ( RightExpr.NodeType <> nil )
//            then
//              begin
//                if LeftExpr.NodeType.Code <> STRING_TYPE
//                  then LeftExpr.NodeFlags := LeftExpr.NodeFlags + [];
//                if RightExpr.NodeType.Code <> STRING_TYPE
//                  then RightExpr.NodeFlags := RightExpr.NodeFlags + [];
//              end;
        end;
    end;

  procedure TvbPass2_ResolveExprs.OnCallOrIndexerExpr( Node : TvbCallOrIndexerExpr; Context : TObject );

    procedure FlagPointerParams( dllFunc : TvbDllFuncDef );
      var
        arg        : TvbExpr;
        i          : integer;
        p          : TvbParamDef;
        paramCount : integer;
        params     : TvbParamList;
      begin
        params := dllFunc.Params;
        if ( params = nil ) or ( Node.Args = nil ) then exit;
        paramCount := params.ParamCount;

        // Search the argument lists for ADDRESSOF_EXPRs.
        // For each one found, flag the corresponding parameter in the
        // dll-function definition as being of POINTER type.
        for i := 0 to Node.Args.ArgCount - 1 do
          begin
            arg := Node.Args[i];
            if ( arg <> nil ) and
               ( arg.NodeKind = ADDRESSOF_EXPR ) and
               ( i < paramCount )
              then
                begin
                  p := params[i];
                  p.NodeFlags := p.NodeFlags + [_pfPointerType];
                end;
          end;
      end;

    var
      d : TvbDef;
    begin
      // The inherited method resolves the arguemnts and the target expression.
      inherited;
      with Node do
        begin
          // this is the definition that the target refers to.
          d := FuncOrVar.Def;
          
          if FuncOrVar.NodeType <> nil
            then
              begin
                // Keep in mind that if FuncOrVar is an array variable
                // this TvbCallOrIndexerExpr expression is accessing an element
                // which base type is FuncOrVar.NodeType without the array flag.
                NodeType := FuncOrVar.NodeType;
              end;
          // Store the definition that FuncOrVar refers to in the
          // TvbCallOrIndexerExpr node.
          if d <> nil
            then
              begin
                Def := d;
                // if the target definition is a VAR, PARAM or PROPERTY
                // flag this expression as "not-being-a-method".
                // i use this flag to later decide whether to enclose
                // the arguemnts in [] or ().
                if ( d.NodeKind in [VAR_DEF, PARAM_DEF] ) or
                   ( ( d.NodeKind = FUNC_DEF ) and
                     ( d.NodeFlags * [dfPropLet, dfPropSet, dfPropGet] <> [] ) )
                  then NodeFlags := NodeFlags + [expr_isNotMethod];
              end;

          if ( Def <> nil ) and ( Def.NodeKind = DLL_FUNC_DEF )
            then FlagPointerParams( TvbDllFuncDef( Def ) );
        end;
    end;

  procedure TvbPass2_ResolveExprs.OnCaseClause( Node : TvbCaseClause; Context : TObject );
    begin
      inherited;
      assert( fCurrentSelectCase <> nil );
      with Node do
        begin
          if not ( expr_Const in Expr.NodeFlags ) or
             not IsIntegralType( Expr.NodeType )
            then fCurrentSelectCase.NodeFlags := fCurrentSelectCase.NodeFlags + [select_AsIfElse]; 
        end;
    end;

  procedure TvbPass2_ResolveExprs.OnCaseStmt( Node : TvbCaseStmt; Context : TObject );
    begin
      inherited;
    end;

  procedure TvbPass2_ResolveExprs.OnCastExpr( Node : TvbCastExpr; Context : TObject );
    begin
      inherited;
      with Node do
        begin
          if expr_Const in Op1.NodeFlags
            then  NodeFlags := NodeFlags + [expr_Const];
          case CastTo of
            BOOLEAN_CAST  : NodeType := vbBooleanType;
            BYTE_CAST     : NodeType := vbByteType;
            CURRENCY_CAST : NodeType := vbCurrencyType;
            DATE_CAST     : NodeType := vbDateType;
            DOUBLE_CAST   : NodeType := vbDoubleType;
            INTEGER_CAST  : NodeType := vbIntegerType;
            LONG_CAST     : NodeType := vbLongType;
            SINGLE_CAST   : NodeType := vbSingleType;
            STRING_CAST   : NodeType := vbStringType;
            else NodeType := vbVariantType;
          end;
        end;
    end;

  procedure TvbPass2_ResolveExprs.OnCircleStmt( Node : TvbCircleStmt; Context : TObject );
    begin
      inherited;
    end;

  procedure TvbPass2_ResolveExprs.OnConstDef( Node : TvbConstDef; Context : TObject );
    begin
      inherited;
      with Node do
        begin
          if Value <> nil
            then Value.Accept( self, Context );
          // If the constant has no type, get the type from the value(if any).
          if ( ConstType = nil ) and ( Value <> nil )
            then ConstType := Value.NodeType;
        end;
    end;

  procedure TvbPass2_ResolveExprs.OnDateExpr( Node : TvbDateExpr; Context : TObject );
    begin
      inherited;
      Node.NodeType := vbVariantType;
    end;

  procedure TvbPass2_ResolveExprs.OnDateLitExpr( Node : TvbDateLit; Context : TObject );
    begin
      inherited;
      with Node do
        begin
          NodeFlags := Node.NodeFlags + [expr_Const];
          NodeType := vbDateType;
        end;
    end;

  procedure TvbPass2_ResolveExprs.OnDictAccessExpr( Node : TvbDictAccessExpr; Context : TObject );
    begin
      inherited;

    end;

  procedure TvbPass2_ResolveExprs.OnFixExpr( Node : TvbFixExpr; Context : TObject );
    begin
      inherited;
      with Node do
        begin
          if expr_Const in Number.NodeFlags
            then NodeFlags := NodeFlags + [expr_Const];
          NodeType := vbIntegerType;
        end;
    end;

  procedure TvbPass2_ResolveExprs.OnFloatLitExpr( Node : TvbFloatLit; Context : TObject );
    begin
      inherited;
      with Node do
        begin
          NodeFlags := NodeFlags + [expr_Const];
          if sfx_Single in NodeFlags
            then NodeType := vbSingleType
            else NodeType := vbDoubleType;
        end;
    end;

  procedure TvbPass2_ResolveExprs.OnForStmt( Node : TvbForStmt; Context : TObject );
    begin
      inherited;
      with Node do
        if StepValue <> nil
          then
            // Check if Step is negative.
            if ( StepValue.NodeKind = UNARY_EXPR ) and
               ( TvbUnaryExpr( StepValue ).Oper = SUB_OP )
              then
                // If Step is -1 we'll convert it to FOR..DOWNTO,
                // otherwise to a WHILE..DO.
                if IsIntegerEqTo( TvbUnaryExpr( StepValue ).Expr, 1 )
                  then NodeFlags := NodeFlags + [for_StepMinus1]
                  else NodeFlags := NodeFlags + [for_StepN, for_StepMinusN]
              else
            // If Step is positive 1 we'll convert to a regular FOR..TO,
            // otherwise to a WHILE..DO.
            if not IsIntegerEqTo( StepValue, 1 )
              then NodeFlags := NodeFlags + [for_StepN];
    end;

  procedure TvbPass2_ResolveExprs.OnFuncResultExpr( Node : TvbFuncResultExpr; Context : TObject );
    begin
      inherited;

    end;

  procedure TvbPass2_ResolveExprs.OnInputExpr( Node : TvbInputExpr; Context : TObject );
    begin
      inherited;
      Node.NodeType := vbStringType;
    end;

  procedure TvbPass2_ResolveExprs.OnIntegerLitExpr( Node : TvbIntLit; Context : TObject );
    begin
      inherited;
      with Node do
        begin
          NodeFlags := NodeFlags + [expr_Const];
          if sfx_Int in NodeFlags
            then NodeType := vbIntegerType
            else
          if sfx_Long in NodeFlags
            then NodeType := vbLongType
            else
              if ( Value > vbMaxInteger ) or ( Value < vbMinInteger )
                then NodeType := vbLongType
                else NodeType := vbIntegerType;
        end;
    end;

  procedure TvbPass2_ResolveExprs.OnIntExpr( Node : TvbIntExpr; Context : TObject );
    begin
      inherited;
      with Node do
        begin
          if expr_Const in Number.NodeFlags
            then NodeFlags := NodeFlags + [expr_Const];
          NodeType := vbIntegerType;
        end;
    end;

  procedure TvbPass2_ResolveExprs.OnLabelDef( Node : TvbLabelDef; Context : TObject );
    begin
      inherited;

    end;

  procedure TvbPass2_ResolveExprs.OnLabelDefStmt( Node : TvbLabelDefStmt; Context : TObject );
    begin
      inherited;

    end;

  procedure TvbPass2_ResolveExprs.OnLBoundExpr( Node : TvbLBoundExpr; Context : TObject );
    begin
      inherited;
      with Node do
        begin
          if IsStaticArrayType( ArrayVar.NodeType )
            then NodeFlags := NodeFlags + [expr_Const];
          NodeType := vbLongType;
        end;
    end;

  procedure TvbPass2_ResolveExprs.OnLenBExpr( Node : TvbLenBExpr; Context : TObject );
    begin
      inherited;
      with Node do
        begin
          if expr_Const in StrOrVar.NodeFlags
            then NodeFlags := NodeFlags + [expr_Const];
          NodeType := vbLongType;
        end;
    end;

  procedure TvbPass2_ResolveExprs.OnLenExpr( Node : TvbLenExpr; Context : TObject );
    begin
      inherited;
      with Node do
        begin
          if expr_Const in StrOrVar.NodeFlags
            then NodeFlags := NodeFlags + [expr_Const];
          NodeType := vbLongType;
        end;
    end;

  procedure TvbPass2_ResolveExprs.OnLineStmt( Node : TvbLineStmt; Context : TObject );
    begin
      inherited;
    end;

  procedure TvbPass2_ResolveExprs.OnMeExpr( Node : TvbMeExpr; Context : TObject );
    begin
      inherited;

    end;

  procedure TvbPass2_ResolveExprs.OnMemberAccessExpr( Node : TvbMemberAccessExpr; Context : TObject );

    function GetContainerScope : TvbScope;
      var
        Target     : TvbExpr;
        TargetType : IvbType;
      begin
        result := nil;
        // Remember that if our Target property is NIL this expression
        // is refering to the expression of the most-recent With statement.
        if Node.Target = nil
          then Target := Node.WithExpr
          else Target := Node.Target;

        assert( Target <> nil );

        TargetType := Target.NodeType;

        if TargetType <> nil
          then
            begin
              if TargetType.TypeDef <> nil
                then
                  case TargetType.TypeDef.NodeKind of
                    CLASS_DEF,
                    RECORD_DEF,
                    ENUM_DEF  : result := TvbTypeDef( TargetType.TypeDef ).Scope;
                  end;
            end
          else
        if Target.Def <> nil
          then
            begin
              case Target.Def.NodeKind of
                CLASS_MODULE_DEF,
                STD_MODULE_DEF     : result := TvbModuleDef( Target.Def ).Scope;
                PROJECT_DEF        : result := TvbProject( Target.Def ).Scope;
                LIB_MODULE_DEF     : result := TvbLibModuleDef( Target.Def ).Scope;
              end;
            end;
      end;

    var
      containerScope : TvbScope;
      FoundDef    : TvbDef;
      flags : TvbLookupObjectFlags;
    begin
      inherited;
      containerScope := GetContainerScope;
      if containerScope = nil
        then exit;

      with Node do
        begin
          flags := [];
          if IsRValue
            then Include( flags, lofMustBeRValue );
          if IsUsedWithArgs
            then Include( flags, lofMustTakeArgs );
          if IsUsedWithEmptyParens
            then Include( flags, lofMustBeFunc );

          // First search the container scope for an object matching the criteria.
          if containerScope.LookupObject( Member, FoundDef, flags, false )
            then
              begin
                NodeType := FoundDef.NodeType;
                Def := FoundDef;
                // If we found a constant flag the expression node as
                // being constant.
                if FoundDef.NodeKind = CONST_DEF
                  then NodeFlags := NodeFlags + [expr_Const];
              end
            else
          // We didn't find an object, so now search for a module or type
          // definition since they may be used as qualifiers.
          if containerScope.Lookup( PROJECT_OR_MODULE_BINDING, Member, FoundDef, false ) or
             containerScope.Lookup( TYPE_BINDING, Member, FoundDef, false )
            then Def := FoundDef;
        end;
    end;

  procedure TvbPass2_ResolveExprs.OnMidExpr( Node : TvbMidExpr; Context : TObject );
    begin
      inherited;
      Node.NodeType := vbStringType;
    end;

  procedure TvbPass2_ResolveExprs.OnNameExpr( Node : TvbNameExpr; Context : TObject );
    var
      FoundDef     : TvbDef;
      CurrentScope : TvbScope absolute Context;
      flags        : TvbLookupObjectFlags;
    begin
      assert( ( Context <> nil ) and ( Context is TvbScope ) );
      flags := [];
      with Node do
        begin
          if IsRValue
            then Include( flags, lofMustBeRValue );
          if IsUsedWithArgs
            then Include( flags, lofMustTakeArgs );
          if IsUsedWithEmptyParens
            then Include( flags, lofMustBeFunc );
            
          // First search for an object definition
          // (var,const, parameter,function or property).
          if CurrentScope.LookupObject( Name, FoundDef, flags, true )
            then
              begin
                NodeType := FoundDef.NodeType;
                Def := FoundDef;
                if FoundDef.NodeKind = CONST_DEF
                  then NodeFlags := NodeFlags + [expr_Const];
              end
            else
          // Now search FIRST for a module/project, LATER for a type definition.
          // This order is important because in a class module the class type
          // and the containing module have the same name and in the context
          // of an expression we are interested in finding the module's
          // declaration which is typically used as a qualifier.
          if CurrentScope.Lookup( PROJECT_OR_MODULE_BINDING, Name, FoundDef, true ) or
             CurrentScope.Lookup( TYPE_BINDING, Name, FoundDef, true )
            then Def := FoundDef;

        end;
    end;

  procedure TvbPass2_ResolveExprs.OnNewExpr( Node : TvbNewExpr; Context : TObject );
    begin
      inherited;

    end;

  procedure TvbPass2_ResolveExprs.OnNothingExpr( Node : TvbNothingLit; Context : TObject );
    begin
      inherited;
      with Node do
        begin
          NodeFlags := NodeFlags + [expr_Const];
          NodeType := vbLongType;
        end;
    end;

  procedure TvbPass2_ResolveExprs.OnParamDef( Node : TvbParamDef; Context : TObject );
    begin
      inherited;
      if Node.DefaultValue <> nil
        then Node.DefaultValue.Accept( self, Context );
    end;

  procedure TvbPass2_ResolveExprs.OnPrintArgExpr( Node : TvbPrintStmtExpr; Context : TObject );
    begin
      inherited;
      if Node.Expr <> nil
        then Node.Expr.Accept( self, Context );
    end;

  procedure TvbPass2_ResolveExprs.OnProjectDef( Node : TvbProject; Context : TObject );
    begin
      fGlobalScope := Node.Scope;
      assert( fGlobalScope <> nil );
      inherited OnProjectDef( Node, fGlobalScope );
      with Node do
        begin
        end;
    end;

  procedure TvbPass2_ResolveExprs.OnPSetStmt( Node : TvbPSetStmt; Context : TObject );
    begin
      inherited;
    end;

  procedure TvbPass2_ResolveExprs.OnQualifiedName( Node : TvbQualifiedName; Context : TObject );
    begin
      inherited;

    end;

  procedure TvbPass2_ResolveExprs.OnRangeCaseClause( Node : TvbRangeCaseClause; Context : TObject );
    begin
      inherited;
      assert( fCurrentSelectCase <> nil );
      with Node do
        begin
          if not ( expr_Const in Expr.NodeFlags ) or
             not ( expr_Const in UpperExpr.NodeFlags ) or
             not IsIntegralType( Expr.NodeType ) or
             not IsIntegralType( UpperExpr.NodeType )
            then fCurrentSelectCase.NodeFlags := fCurrentSelectCase.NodeFlags + [select_AsIfElse];
        end;
    end;

  procedure TvbPass2_ResolveExprs.OnRelationalCaseClause( Node : TvbRelationalCaseClause; Context : TObject );
    begin
      inherited;
      assert( fCurrentSelectCase <> nil );
      with Node do
        begin
          if ( Operator <> EQ_OP ) or
             not ( expr_Const in Expr.NodeFlags ) or
             not IsIntegralType( Expr.NodeType )
            then fCurrentSelectCase.NodeFlags := fCurrentSelectCase.NodeFlags + [select_AsIfElse];
        end;
    end;

  procedure TvbPass2_ResolveExprs.OnResumeStmt( Node : TvbResumeStmt; Context : TObject );
    begin
      inherited;
      with Node do
        begin

        end;
    end;

  procedure TvbPass2_ResolveExprs.OnSeekExpr( Node : TvbSeekExpr; Context : TObject );
    begin
      inherited;
      Node.NodeType := vbLongType;
    end;

  procedure TvbPass2_ResolveExprs.OnSelectCaseStmt( Node : TvbSelectCaseStmt; Context : TObject );
    begin
      fCurrentSelectCase := Node;
      inherited;
      with Node do
        begin
          // Among other things, if the expression hasn't integral type
          // we CANNOT convert it to a Delphi's CASE OF statement.
          if not IsIntegralType( Expr.NodeType )
            then NodeFlags := NodeFlags + [select_AsIfElse];
        end;
    end;

  procedure TvbPass2_ResolveExprs.OnSgnExpr( Node : TvbSgnExpr; Context : TObject );
    begin
      inherited;
      with Node do
        begin
          if expr_Const in Number.NodeFlags
            then NodeFlags := NodeFlags + [expr_Const];
          NodeType := vbVariantType;
        end;
    end;

  procedure TvbPass2_ResolveExprs.OnSimpleName( Node : TvbSimpleName; Context : TObject );
    begin
      inherited;

    end;

  procedure TvbPass2_ResolveExprs.OnStmtBlock( Node : TvbStmtBlock; Context : TObject );
    begin
      inherited;

    end;

  procedure TvbPass2_ResolveExprs.OnStringExpr( Node : TvbStringExpr; Context : TObject );
    begin
      inherited;
      Node.NodeType := vbStringType;
    end;

  procedure TvbPass2_ResolveExprs.OnTypeOfExpr( Node : TvbTypeOfExpr; Context : TObject );
    begin
      inherited;
      with Node do
        begin
          NodeFlags := NodeFlags + [expr_Const];
          NodeType := vbBooleanType;
        end;
    end;

  procedure TvbPass2_ResolveExprs.OnUboundExpr( Node : TvbUBoundExpr; Context : TObject );
    begin
      inherited;
      with Node do
        begin
          if IsStaticArrayType( ArrayVar.NodeType )
            then NodeFlags := NodeFlags + [expr_Const];
          NodeType := vbLongType;
        end;      
    end;

  procedure TvbPass2_ResolveExprs.OnUnaryExpr( Node : TvbUnaryExpr; Context : TObject );
    begin
      inherited;
      with Node do
        begin
          if expr_Const in Expr.NodeFlags
            then NodeFlags := NodeFlags + [expr_Const];
          NodeType := GetUnaryExpressionType( Oper, Expr.NodeType );
        end;      
    end;

  procedure TvbPass2_ResolveExprs.OnWithStmt( Node : TvbWithStmt; Context : TObject );
    begin
      inherited;

    end;

end.























