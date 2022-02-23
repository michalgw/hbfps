/************************************************************************

hbfps

Copyright (C) 2022  GM Systems Michal Gawrycki (gmsystems.pl)

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program. If not, see <http://www.gnu.org/licenses/>.

************************************************************************/

#include "hbdyn.ch"
#include "hbclass.ch"

STATIC p_hbfpslib := NIL

STATIC FUNCTION hbfpsLoadLib()

   IF p_hbfpslib <> NIL
      RETURN .T.
   ENDIF

   p_hbfpslib := hb_libLoad( "hbfps.dll" )

   RETURN p_hbfpslib <> NIL

CREATE CLASS TsWorkbook

   PROTECTED:
   VAR pClass

   EXPORTED:
   METHOD New() CONSTRUCTOR
   DESTRUCTOR Free()
   METHOD ReadFromFile( cFileName )
   METHOD WriteToFile( cFileName )
   METHOD AddWorksheet( cName )

ENDCLASS

CREATE CLASS TsWorksheet

   PROTECTED:
   VAR pClass

   EXPORTED:
   METHOD New( pClass ) CONSTRUCTOR
   METHOD WriteRowHeight( nRow, nHeight, nUnits )
   METHOD WriteColWidth( nCol, nWidth, nUnits)
   METHOD WriteCurrency( nRow, nCol, nValue)
   METHOD WriteNumber( nRow, nCol, nValue )
   METHOD WriteText( nRow, nCol, cText )
   METHOD WriteDate( nRow, nCol, dDate )

ENDCLASS

METHOD New() CLASS TsWorkbook

   IF ! hbfpsLoadLib()
      RETURN NIL
   ENDIF

   IF ( ::pClass := hb_DynCall( { "fps_Workbook_Create", p_hbfpslib, ;
      hb_bitOr( HB_DYN_CTYPE_VOID_PTR, HB_DYN_CALLCONV_CDECL ) } ) ) == NIL
      RETURN NIL
   ENDIF

   RETURN Self

DESTRUCTOR Free() CLASS TsWorkbook

   hb_DynCall( { "fps_Workbook_Free", p_hbfpslib, ;
      hb_bitOr( HB_DYN_CTYPE_VOID, HB_DYN_CALLCONV_CDECL ), ;
      HB_DYN_CTYPE_VOID_PTR }, ::pClass )

   RETURN

METHOD ReadFromFile( cFileName ) CLASS TsWorkbook

   LOCAL nRes

   nRes := hb_DynCall( { "fps_Workbook_ReadFromFile", p_hbfpslib, ;
      hb_bitOr( HB_DYN_CTYPE_INT, HB_DYN_CALLCONV_CDECL ), ;
      HB_DYN_CTYPE_VOID_PTR, hb_bitOr( HB_DYN_CTYPE_CHAR_PTR, HB_DYN_ENC_UTF8 ) }, ;
      ::pClass, cFileName )

   RETURN nRes

METHOD WriteToFile( cFileName ) CLASS TsWorkbook

   LOCAL nRes

   nRes := hb_DynCall( { "fps_Workbook_WriteToFile", p_hbfpslib, ;
      hb_bitOr( HB_DYN_CTYPE_INT, HB_DYN_CALLCONV_CDECL ), ;
      HB_DYN_CTYPE_VOID_PTR, hb_bitOr( HB_DYN_CTYPE_CHAR_PTR, HB_DYN_ENC_UTF8 ) }, ;
      ::pClass, cFileName )

   RETURN nRes

METHOD AddWorksheet( cName ) CLASS TsWorkbook

   LOCAL pRes

   pRes := hb_DynCall( { "fps_Workbook_AddWorksheet", p_hbfpslib, ;
      hb_bitOr( HB_DYN_CTYPE_VOID_PTR, HB_DYN_CALLCONV_CDECL ), ;
      HB_DYN_CTYPE_VOID_PTR, hb_bitOr( HB_DYN_CTYPE_CHAR_PTR, HB_DYN_ENC_UTF8 ) }, ;
      ::pClass, cName )

   RETURN TsWorksheet():New( pRes )

////////////////////////////////////////////

METHOD New( pClass ) CLASS TsWorksheet

   ::pClass := pClass

   RETURN Self

METHOD WriteRowHeight( nRow, nHeight, nUnits ) CLASS TsWorksheet

   LOCAL nRes

   nRes := hb_DynCall( { "fps_Worksheet_WriteRowHeight", p_hbfpslib, ;
      hb_bitOr( HB_DYN_CTYPE_VOID, HB_DYN_CALLCONV_CDECL ), ;
      HB_DYN_CTYPE_VOID_PTR, HB_DYN_CTYPE_INT_UNSIGNED, HB_DYN_CTYPE_DOUBLE, ;
      HB_DYN_CTYPE_INT }, ::pClass, nRow, nHeight, nUnits )

   RETURN nRes

METHOD WriteColWidth( nCol, nWidth, nUnits) CLASS TsWorksheet

   LOCAL nRes

   nRes := hb_DynCall( { "fps_Worksheet_WriteColWidth", p_hbfpslib, ;
      hb_bitOr( HB_DYN_CTYPE_VOID, HB_DYN_CALLCONV_CDECL ), ;
      HB_DYN_CTYPE_VOID_PTR, HB_DYN_CTYPE_INT_UNSIGNED, HB_DYN_CTYPE_DOUBLE, ;
      HB_DYN_CTYPE_INT }, ::pClass, nCol, nWidth, nUnits )

   RETURN nRes

METHOD WriteCurrency( nRow, nCol, nValue) CLASS TsWorksheet

   LOCAL nRes

   nRes := hb_DynCall( { "fps_Worksheet_WriteCurrency", p_hbfpslib, ;
      hb_bitOr( HB_DYN_CTYPE_VOID, HB_DYN_CALLCONV_CDECL ), ;
      HB_DYN_CTYPE_VOID_PTR, HB_DYN_CTYPE_INT_UNSIGNED, HB_DYN_CTYPE_INT_UNSIGNED, ;
      HB_DYN_CTYPE_DOUBLE }, ::pClass, nRow, nCol, nValue )

   RETURN nRes

METHOD WriteNumber( nRow, nCol, nValue ) CLASS TsWorksheet

   LOCAL nRes

   nRes := hb_DynCall( { "fps_Worksheet_WriteNumber", p_hbfpslib, ;
      hb_bitOr( HB_DYN_CTYPE_VOID, HB_DYN_CALLCONV_CDECL ), ;
      HB_DYN_CTYPE_VOID_PTR, HB_DYN_CTYPE_INT_UNSIGNED, HB_DYN_CTYPE_INT_UNSIGNED, ;
      HB_DYN_CTYPE_DOUBLE }, ::pClass, nRow, nCol, nValue )

   RETURN nRes

METHOD WriteText( nRow, nCol, cText ) CLASS TsWorksheet

   LOCAL nRes

   nRes := hb_DynCall( { "fps_Worksheet_WriteText", p_hbfpslib, ;
      hb_bitOr( HB_DYN_CTYPE_VOID, HB_DYN_CALLCONV_CDECL ), ;
      HB_DYN_CTYPE_VOID_PTR, HB_DYN_CTYPE_INT_UNSIGNED, HB_DYN_CTYPE_INT_UNSIGNED, ;
      hb_bitOr( HB_DYN_CTYPE_CHAR_PTR, HB_DYN_ENC_UTF8 ) }, ::pClass, nRow, nCol, cText )

   RETURN nRes

METHOD WriteDate( nRow, nCol, dDate ) CLASS TsWorksheet

   LOCAL nRes

   nRes := hb_DynCall( { "fps_Worksheet_WriteDate", p_hbfpslib, ;
      hb_bitOr( HB_DYN_CTYPE_VOID, HB_DYN_CALLCONV_CDECL ), ;
      HB_DYN_CTYPE_VOID_PTR, HB_DYN_CTYPE_INT_UNSIGNED, HB_DYN_CTYPE_INT_UNSIGNED, ;
      HB_DYN_CTYPE_INT, HB_DYN_CTYPE_INT, HB_DYN_CTYPE_INT }, ::pClass, nRow, nCol, ;
      Year( dDate ), Month( dDate ), Day( dDate ) )

   RETURN nRes
