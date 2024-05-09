Attribute VB_Name = "ModZlib"


Option Explicit
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, FreeBytesToCaller As Currency, BytesTotal As Currency, FreeBytesTotal As Currency) As Long
 
Public Type FILEHEADER
    lngFileSize As Long
    intNumFiles As Integer
End Type
 
Public Type INFOHEADER
    lngFileStart As Long
    lngFileSize As Long
    strFileName As String * 16
    lngFileSizeUncompressed As Long
End Type
 
Public Temp_Windows_Directory As String
 
Public Enum resource_file_type
    graphics
    Interface
End Enum
 
Private Const GRAPHIC_PATH As String = "\BMP\"
Private Const GRAPHIC_PNG_PATH As String = "\PNG\"
Private Const MIDI_PATH As String = "\Midi\"
Private Const MP3_PATH As String = "\Mp3\"
Private Const WAV_PATH As String = "\Wav\"
Private Const MAP_PATH As String = "\Mapas\"
Private Const INTERFACE_PATH As String = "\Interface\"
Private Const SCRIPT_PATH As String = "\Init\"
Private Const PATCH_PATH As String = "\Patches\"
Private Const OUTPUT_PATH As String = "\Output\"
 
Private Declare Function Compress Lib "zlib.dll" Alias "compress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function UnCompress Lib "zlib.dll" Alias "uncompress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
 
Private Const MAX_LENGTH = 512
 
Public Sub Compress_Data(ByRef Data() As Byte)
    Dim Dimensions As Long
    Dim DimBuffer As Long
    Dim BufTemp() As Byte
    Dim BufTemp2() As Byte
    Dim LoopC As Long
   
    Dimensions = UBound(Data)
    DimBuffer = Dimensions * 1.06
   
    ReDim BufTemp(DimBuffer)
    Compress BufTemp(0), DimBuffer, Data(0), Dimensions
    Erase Data
   
    ReDim Preserve BufTemp(DimBuffer - 1)
    Data = BufTemp
    Erase BufTemp
   
    Data(0) = Data(0) Xor 166
End Sub
 
Public Sub Decompress_Data(ByRef Data() As Byte, ByVal OrigSize As Long)
    Dim BufTemp() As Byte
   
    ReDim BufTemp(OrigSize - 1)
    Data(0) = Data(0) Xor 166
    UnCompress BufTemp(0), OrigSize, Data(0), UBound(Data) + 1
   
    ReDim Data(OrigSize - 1)
    Data = BufTemp
    Erase BufTemp
End Sub
Public Function Extract_File(ByVal file_type As resource_file_type, ByVal resource_path As String, ByVal File_Name As String, ByVal OutputFilePath As String) As Boolean
   
    Dim LoopC As Long
    Dim SourceFilePath As String
    Dim SourceData() As Byte
    Dim InfoHead As INFOHEADER
    Dim handle As Integer
   
On Local Error GoTo ErrHandler
   
    Select Case file_type
 
        Case graphics
                SourceFilePath = resource_path & "\Graficos.LAO"
           
        Case Interface
                SourceFilePath = resource_path & "\Interface.LAO"
       
        Case Else
            Exit Function
    End Select
   
    InfoHead = File_Find(SourceFilePath, File_Name)
   
    If InfoHead.strFileName = "" Or InfoHead.lngFileSize = 0 Then Exit Function
 
    handle = FreeFile
    Open SourceFilePath For Binary Access Read Lock Write As handle
   
    If InfoHead.lngFileSizeUncompressed > General_Drive_Get_Free_Bytes(Left$(App.path, 3)) Then
        Close handle
        MsgBox "There is not enough drive space to extract the compressed file.", , "Error"
        Exit Function
    End If
   
   
    ReDim SourceData(InfoHead.lngFileSize - 1)
   
    Get handle, InfoHead.lngFileStart, SourceData
        Decompress_Data SourceData, InfoHead.lngFileSizeUncompressed
    Close handle
   
    handle = FreeFile
    Open OutputFilePath & InfoHead.strFileName For Binary As handle
        Put handle, 1, SourceData
    Close handle
   
    Erase SourceData
       
    Extract_File = True
Exit Function
 
ErrHandler:
    Close handle
    Erase SourceData
End Function
 
Public Function Extract_File_Memory(ByVal file_type As resource_file_type, ByVal resource_path As String, ByVal File_Name As String, ByRef SourceData() As Byte) As Boolean
 
    ' Parra was here (;
    Dim LoopC As Long
    Dim SourceFilePath As String
    Dim InfoHead As INFOHEADER
    Dim handle As Integer
   
On Local Error GoTo ErrHandler
   
    Select Case file_type
 
        Case graphics
                SourceFilePath = resource_path & "\Graficos.LAO"
           
        Case Interface
                SourceFilePath = resource_path & "\Interface.LAO"
       
        Case Else
            Exit Function
    End Select
   
    InfoHead = File_Find(SourceFilePath, File_Name)
   
    If InfoHead.strFileName = "" Or InfoHead.lngFileSize = 0 Then Exit Function
 
    handle = FreeFile
    Open SourceFilePath For Binary Access Read Lock Write As handle
   
    If InfoHead.lngFileSizeUncompressed > General_Drive_Get_Free_Bytes(Left$(App.path, 3)) Then
        Close handle
        MsgBox "There is not enough drive space to extract the compressed file.", , "Error"
        Exit Function
    End If
   
   
    ReDim SourceData(InfoHead.lngFileSize - 1)
   
    Get handle, InfoHead.lngFileStart, SourceData
        Decompress_Data SourceData, InfoHead.lngFileSizeUncompressed
    Close handle
       
    Extract_File_Memory = True
Exit Function
 
ErrHandler:
    Close handle
    Erase SourceData
End Function
 
Public Sub Delete_File(ByVal file_path As String)
    Dim handle As Integer
    Dim Data() As Byte
   
    On Error GoTo Error_Handler
   
    handle = FreeFile
    Open file_path For Binary Access Write Lock Read As handle
   
    ReDim Data(LOF(handle) - 1)
    Put handle, 1, Data
   
    Close handle
   
    Kill file_path
   
    Exit Sub
   
Error_Handler:
    Kill file_path
       
End Sub
 
Public Function File_Find(ByVal resource_file_path As String, ByVal File_Name As String) As INFOHEADER
 
On Error GoTo ErrHandler
 
    Dim Max As Integer
    Dim Min As Integer
    Dim mid As Integer
    Dim file_handler As Integer
   
    Dim file_head As FILEHEADER
    Dim info_head As INFOHEADER
   
    If Len(File_Name) < Len(info_head.strFileName) Then _
        File_Name = File_Name & Space$(Len(info_head.strFileName) - Len(File_Name))
   
    file_handler = FreeFile
    Open resource_file_path For Binary Access Read Lock Write As file_handler
   
    Get file_handler, 1, file_head
   
    Min = 1
    Max = file_head.intNumFiles
   
    Do While Min <= Max
        mid = (Min + Max) / 2
       
        Get file_handler, CLng(Len(file_head) + CLng(Len(info_head)) * CLng((mid - 1)) + 1), info_head
               
        If File_Name < info_head.strFileName Then
            If Max = mid Then
                Max = Max - 1
            Else
                Max = mid
            End If
        ElseIf File_Name > info_head.strFileName Then
            If Min = mid Then
                Min = Min + 1
            Else
                Min = mid
            End If
        Else
            File_Find = info_head
           
            Close file_handler
            Exit Function
        End If
    Loop
   
ErrHandler:
    Close file_handler
    File_Find.strFileName = ""
    File_Find.lngFileSize = 0
End Function
 
Public Function General_Get_Temp_Dir() As String
   Dim s As String
   Dim c As Long
   s = Space$(MAX_LENGTH)
   c = GetTempPath(MAX_LENGTH, s)
   If c > 0 Then
       If c > Len(s) Then
           s = Space$(c + 1)
           c = GetTempPath(MAX_LENGTH, s)
       End If
   End If
   General_Get_Temp_Dir = IIf(c > 0, Left$(s, c), "")
End Function
Public Sub General_Quick_Sort(ByRef SortArray As Variant, ByVal first As Long, ByVal last As Long)
    Dim Low As Long, High As Long
    Dim temp As Variant
    Dim List_Separator As Variant
   
    Low = first
    High = last
    List_Separator = SortArray((first + last) / 2)
    Do While (Low <= High)
        Do While SortArray(Low) < List_Separator
            Low = Low + 1
        Loop
        Do While SortArray(High) > List_Separator
            High = High - 1
        Loop
        If Low <= High Then
            temp = SortArray(Low)
            SortArray(Low) = SortArray(High)
            SortArray(High) = temp
            Low = Low + 1
            High = High - 1
        End If
    Loop
    If first < High Then General_Quick_Sort SortArray, first, High
    If Low < last Then General_Quick_Sort SortArray, Low, last
End Sub
 
Public Function General_Drive_Get_Free_Bytes(ByVal DriveName As String) As Currency
    Dim RetVal As Long
    Dim FB As Currency
    Dim BT As Currency
    Dim FBT As Currency
   
    RetVal = GetDiskFreeSpace(Left(DriveName, 2), FB, BT, FBT)
   
    General_Drive_Get_Free_Bytes = FB * 10000
End Function
 
Public Function Get_Extract(ByVal file_type As resource_file_type, ByVal File_Name As String) As String
    Extract_File file_type, App.path & "\Graficos", LCase$(File_Name), App.path & "\amdInData\"
    Get_Extract = App.path & "\amdInData\" & LCase$(File_Name)
End Function
Public Function Get_Interface(ByVal file_type As resource_file_type, ByVal File_Name As String) As String
    Extract_File file_type, App.path & "\Interface", LCase$(File_Name), App.path & "\Interface\"
    Get_Interface = App.path & "\Interface\" & LCase$(File_Name)
End Function



