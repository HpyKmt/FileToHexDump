# FileToHexDump
Convert a file to hex dump text files.

## Situation
- Cannot install modern technology or 3rd party tools.
- Only VBScript and Batch are runnable.

## Objective
1. Want to view binary file structure in hex dump text file.
2. Want split the hex dump
   1. Want to specify output column count
   2. Want to specify output row count per file.
3. Want to output each file as a chuck is read so that analysis can occur while parsing is ongoing.

## How to use

```text
+---------------------------------------------------+
|                                                   |
| C:\_temp\FileToHexDump>run                        |
| ---Get Input File Path---                         |
| Source File: C:\_temp\MyTempFile.dat              | Set source file
| FpIn=C:\_temp\MyTempFile.dat                      |
| ---Get Output Folder---                           |
| Output Folder: C:\_temp\output                    | Set output folder
| DirOut=C:\_temp\output                            |
| ---Row Count---                                   |
| Row Count: 8                                      | Set how many rows per file
| RowCount=8                                        |
| ---Column Count---                                |
| Column Count: 16                                  | Set how many columns per line
| ColumnCount=16                                    |
| ---Run---                                         |
| Delete Files in argDir=C:\_temp\output            |
| Deleting: 00000000.txt                            |
| TypeName(arrBytes)=Byte() VarType(arrBytes)=8209  |
| UBound(arrBytes)=127                              |
| WriteTextASCII strFp=C:\_temp\output\00000000.txt |
| TypeName(arrBytes)=Byte() VarType(arrBytes)=8209  |
| UBound(arrBytes)=127                              |
| WriteTextASCII strFp=C:\_temp\output\00000080.txt |
| TypeName(arrBytes)=Byte() VarType(arrBytes)=8209  |
| UBound(arrBytes)=127                              |
| WriteTextASCII strFp=C:\_temp\output\00000100.txt |
| TypeName(arrBytes)=Byte() VarType(arrBytes)=8209  |
| UBound(arrBytes)=45                               |
| WriteTextASCII strFp=C:\_temp\output\00000180.txt |
| TypeName(arrBytes)=Null VarType(arrBytes)=1       |
| No more data                                      |
|                                                   |
|                                                   |
|                                                   |
| C:\_temp\FileToHexDump>                           |
+---------------------------------------------------+

+----------------------------------------------------------+
| C:\_temp\output>dir /b                                   | Several files are exported.
| 00000000.txt                                             | Each file is named by byte starting position.
| 00000080.txt                                             |
| 00000100.txt                                             |
| 00000180.txt                                             |
|                                                          |
| C:\_temp\output>type 00000000.txt                        | Output is columns=16 x rows=8
| -------- 00 01 02 03 04 05 06 07 08 09 0A 0B 0C 0D 0E 0F |
| 00000000 31 32 33 34 35 36 37 38 39 30 51 41 5A 57 53 58 |
| 00000010 45 44 43 52 46 56 0D 0A 31 32 33 34 35 36 37 38 |
| 00000020 39 30 51 41 5A 57 53 58 45 44 43 52 46 56 0D 0A |
| 00000030 31 32 33 34 35 36 37 38 39 30 51 41 5A 57 53 58 |
| 00000040 45 44 43 52 46 56 0D 0A 31 32 33 34 35 36 37 38 |
| 00000050 39 30 51 41 5A 57 53 58 45 44 43 52 46 56 0D 0A |
| 00000060 31 32 33 34 35 36 37 38 39 30 51 41 5A 57 53 58 |
| 00000070 45 44 43 52 46 56 0D 0A 31 32 33 34 35 36 37 38 |
|                                                          |
| C:\_temp\output>                                         |
+----------------------------------------------------------+


+---------------------------------------------------+
| C:\_temp\FileToHexDump>run                        |
| ---Get Input File Path---                         |
| Source File: C:\_temp\MyTempFile.dat              |
| FpIn=C:\_temp\MyTempFile.dat                      |
| ---Get Output Folder---                           |
| Output Folder: C:\_temp\output                    |
| DirOut=C:\_temp\output                            |
| ---Row Count---                                   |
| Row Count: 8                                      | Row = 8
| RowCount=8                                        |
| ---Column Count---                                |
| Column Count: 32                                  | Column = 32
| ColumnCount=32                                    |
| ---Run---                                         |
| Delete Files in argDir=C:\_temp\output            |
| Deleting: 00000000.txt                            |
| Deleting: 00000080.txt                            |
| Deleting: 00000100.txt                            |
| Deleting: 00000180.txt                            |
| TypeName(arrBytes)=Byte() VarType(arrBytes)=8209  |
| UBound(arrBytes)=255                              |
| WriteTextASCII strFp=C:\_temp\output\00000000.txt |
| TypeName(arrBytes)=Byte() VarType(arrBytes)=8209  |
| UBound(arrBytes)=173                              |
| WriteTextASCII strFp=C:\_temp\output\00000100.txt |
| TypeName(arrBytes)=Null VarType(arrBytes)=1       |
| No more data                                      |
|                                                   |
|                                                   |
|                                                   |
| C:\_temp\FileToHexDump>                           |
+---------------------------------------------------+

Different look now...

+----------------------------------------------------------------------------------------------------------+
| C:\_temp\output>dir /b                                                                                   |
| 00000000.txt                                                                                             |
| 00000100.txt                                                                                             |
|                                                                                                          |
| C:\_temp\output>type 00000100.txt                                                                        |
| -------- 00 01 02 03 04 05 06 07 08 09 0A 0B 0C 0D 0E 0F 10 11 12 13 14 15 16 17 18 19 1A 1B 1C 1D 1E 1F |
| 00000100 45 44 43 52 46 56 0D 0A 31 32 33 34 35 36 37 38 39 30 51 41 5A 57 53 58 45 44 43 52 46 56 0D 0A |
| 00000120 31 32 33 34 35 36 37 38 39 30 51 41 5A 57 53 58 45 44 43 52 46 56 0D 0A 31 32 33 34 35 36 37 38 |
| 00000140 39 30 51 41 5A 57 53 58 45 44 43 52 46 56 0D 0A 31 32 33 34 35 36 37 38 39 30 51 41 5A 57 53 58 |
| 00000160 45 44 43 52 46 56 0D 0A 31 32 33 34 35 36 37 38 39 30 51 41 5A 57 53 58 45 44 43 52 46 56 0D 0A |
| 00000180 31 32 33 34 35 36 37 38 39 30 51 41 5A 57 53 58 45 44 43 52 46 56 0D 0A 31 32 33 34 35 36 37 38 |
| 000001A0 39 30 51 41 5A 57 53 58 45 44 43 52 46 56                                                       |
|                                                                                                          |
| C:\_temp\output>                                                                                         |
+----------------------------------------------------------------------------------------------------------+
```
