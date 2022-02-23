{************************************************************************

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

************************************************************************}

unit fpsintf;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils;

function fps_Workbook_Create: Pointer; cdecl;
procedure fps_Workbook_Free(ASelf: Pointer); cdecl;
function fps_Workbook_ReadFromFile(ASelf: Pointer; AFileName: PChar): LongInt; cdecl;
function fps_Workbook_WriteToFile(ASelf: Pointer; AFileName: PChar): LongInt; cdecl;
function fps_Workbook_AddWorksheet(ASelf: Pointer; AName: PChar): Pointer; cdecl;

function fps_Worksheet_WriteRowHeight(ASelf: Pointer; ARowIndex: Cardinal; AHeight: Double; AUnits: LongInt): LongInt; cdecl;
function fps_Worksheet_WriteColWidth(ASelf: Pointer; AColIndex: Cardinal; AWidth: Double; AUnits: LongInt): LongInt; cdecl;
function fps_Worksheet_WriteCurrency(ASelf: Pointer; ARow, ACol: LongInt; AValue: Double): LongInt; cdecl;
function fps_Worksheet_WriteNumber(ASelf: Pointer; ARow, ACol: LongInt; AValue: Double): LongInt; cdecl;
function fps_Worksheet_WriteDate(ASelf: Pointer; ARow, ACol: LongInt; AYear, AMonth, ADay: LongInt): LongInt; cdecl;
function fps_Worksheet_WriteText(ASelf: Pointer; ARow, ACol: LongInt; AText: PChar): LongInt; cdecl;

implementation

uses
  fpSpreadsheet, fpsTypes, {%H-}fpsallformats;

function CheckClass(const AClassPtr: Pointer; const AClass: TClass): Boolean;  inline;
begin
  Result := (AClassPtr <> nil) and (TObject(AClassPtr) is AClass);
end;

function fps_Workbook_Create: Pointer; cdecl;
begin
  try
    Result := TsWorkbook.Create;
  except
    Result := nil;
  end;
end;

procedure fps_Workbook_Free(ASelf: Pointer); cdecl;
begin
  try
    if CheckClass(ASelf, TObject) then
      TObject(ASelf).Free;
  except
  end;
end;

function fps_Workbook_ReadFromFile(ASelf: Pointer; AFileName: PChar): LongInt;
  cdecl;
begin
  try
    if CheckClass(ASelf, TsWorkbook) then
    begin
      TsWorkbook(ASelf).ReadFromFile(AFileName);
      Result := 0;
    end
    else
      Result := -1;
  except
    Result := -2;
  end;
end;

function fps_Workbook_WriteToFile(ASelf: Pointer; AFileName: PChar): LongInt;
  cdecl;
begin
  try
    if CheckClass(ASelf, TsWorkbook) then
    begin
      TsWorkbook(ASelf).WriteToFile(AFileName, True);
      Result := 0;
    end
    else
      Result := -1;
  except
    Result := -2;
  end;
end;

function fps_Workbook_AddWorksheet(ASelf: Pointer; AName: PChar): Pointer;
  cdecl;
begin
  Result := nil;
  try
    if CheckClass(ASelf, TsWorkbook) then
      Result := TsWorkbook(ASelf).AddWorksheet(AName);
  except
  end;
end;

function fps_Worksheet_WriteRowHeight(ASelf: Pointer; ARowIndex: Cardinal;
  AHeight: Double; AUnits: LongInt): LongInt; cdecl;
begin
  try
    if CheckClass(ASelf, TsWorksheet) then
      TsWorksheet(ASelf).WriteRowHeight(ARowIndex, AHeight, TsSizeUnits(AUnits))
    else
      Result := -1;
  except
    Result := -2;
  end;
end;

function fps_Worksheet_WriteColWidth(ASelf: Pointer; AColIndex: Cardinal;
  AWidth: Double; AUnits: LongInt): LongInt; cdecl;
begin
  try
    if CheckClass(ASelf, TsWorksheet) then
      TsWorksheet(ASelf).WriteColWidth(AColIndex, AWidth, TsSizeUnits(AUnits))
    else
      Result := -1;
  except
    Result := -2;
  end;
end;

function fps_Worksheet_WriteCurrency(ASelf: Pointer; ARow, ACol: LongInt;
  AValue: Double): LongInt; cdecl;
begin
  try
    if CheckClass(ASelf, TsWorksheet) then
      if TsWorksheet(ASelf).WriteCurrency(ARow, ACol, AValue) <> nil then
        Result := 0
      else
        Result := -3
    else
      Result := -1;
  except
    Result := -2;
  end;
end;

function fps_Worksheet_WriteNumber(ASelf: Pointer; ARow, ACol: LongInt;
  AValue: Double): LongInt; cdecl;
begin
  try
    if CheckClass(ASelf, TsWorksheet) then
      if TsWorksheet(ASelf).WriteNumber(ARow, ACol, AValue) <> nil then
        Result := 0
      else
        Result := -3
    else
      Result := -1;
  except
    Result := -2;
  end;
end;

function fps_Worksheet_WriteDate(ASelf: Pointer; ARow, ACol: LongInt; AYear,
  AMonth, ADay: LongInt): LongInt; cdecl;
begin
  try
    if CheckClass(ASelf, TsWorksheet) then
      if TsWorksheet(ASelf).WriteDateTime(ARow, ACol, EncodeDate(AYear, AMonth, ADay)) <> nil then
        Result := 0
      else
        Result := -3
    else
      Result := -1;
  except
    Result := -2;
  end;
end;

function fps_Worksheet_WriteText(ASelf: Pointer; ARow, ACol: LongInt;
  AText: PChar): LongInt; cdecl;
begin
  try
    if CheckClass(ASelf, TsWorksheet) then
      if TsWorksheet(ASelf).WriteText(ARow, ACol, AText) <> nil then
        Result := 0
      else
        Result := -3
    else
      Result := -1;
  except
    Result := -2;
  end;
end;

end.

