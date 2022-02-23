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

library hbfps;

{$mode objfpc}{$H+}

uses
  LazUTF8, Classes, fpsintf
  { you can add units after this };

exports
  fps_Workbook_Create,
  fps_Workbook_Free,
  fps_Workbook_ReadFromFile,
  fps_Workbook_WriteToFile,
  fps_Workbook_AddWorksheet,

  fps_Worksheet_WriteRowHeight,
  fps_Worksheet_WriteColWidth,
  fps_Worksheet_WriteCurrency,
  fps_Worksheet_WriteNumber,
  fps_Worksheet_WriteDate,
  fps_Worksheet_WriteText;

begin
end.

