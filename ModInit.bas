Attribute VB_Name = "ModInit"
Option Explicit
'download by http://www.codefans.net
'Help
Public WinBookVer As String
Public Const Cmd_Cls_Help = "�����Ļ��\n\nCLS\n"
Public Const Cmd_Echo_Help = "��ʾ��Ϣ����������Դ򿪻���ϡ�\n\n  Echo [ON | OFF]\n  Echo [message]\n\nҪ��ʾ��ǰ�������ã����벻�������� ECHO��"
Public RC As String

''����API�������õ������г���
''GetStdHandle������ nStdHandle������ȡֵ
Public Const STD_INPUT_HANDLE = -10&
Public Const STD_OUTPUT_HANDLE = -11&
Public Const STD_ERROR_HANDLE = -12&
''SetConsoleTextAttribute������wAttributes������ȡֵ����RGB��ʽ��ϣ�
Public Const FOREGROUND_BLUE = &H1
Public Const FOREGROUND_GREEN = &H2
Public Const FOREGROUND_RED = &H4
Public Const FOREGROUND_INTENSITY = &H8
Public Const BACKGROUND_BLUE = &H10
Public Const BACKGROUND_GREEN = &H20
Public Const BACKGROUND_RED = &H40
Public Const BACKGROUND_INTENSITY = &H80
''SetConsoleMode������ģʽ
Public Const ENABLE_LINE_INPUT = &H2
Public Const ENABLE_ECHO_INPUT = &H4
Public Const ENABLE_MOUSE_INPUT = &H10
Public Const ENABLE_PROCESSED_INPUT = &H1
Public Const ENABLE_WINDOW_INPUT = &H8
''SetConsoleMode�����ģʽ
Public Const ENABLE_PROCESSED_OUTPUT = &H1
Public Const ENABLE_WRAP_AT_EOL_OUTPUT = &H2

Public Const CONSOLE_TEXTMODE_BUFFER = 1
Public Const CREATE_NEW_CONSOLE = &H10

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Public Type CONSOLE_CURSOR_INFO
    dwSize As Long
    bVisible As Long
End Type
Public Type CHAR_INFO
    Char As Integer
    Attributes As Integer
End Type
Public Type COORD
    x As Integer
    y As Integer
End Type
Public Type SMALL_RECT
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
End Type
Public Type CONSOLE_SCREEN_BUFFER_INFO
    dwSize As COORD
    dwCursorPosition As COORD
    wAttributes As Integer
    srWindow As SMALL_RECT
    dwMaximumWindowSize As COORD
End Type

Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long

'�������̨���뻺����
Private Declare Function FlushConsoleInputBuffer Lib "kernel32" (ByVal hConsoleInput As Long) As Long
'������Ļ�ı�����
Private Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
'������갴ť��
Private Declare Function GetNumberOfConsoleMouseButtons Lib "kernel32" (lpNumberOfMouseButtons As Long) As Long
'���ÿ���̨���λ��
Private Declare Function SetConsoleCursorPosition Lib "kernel32" (ByVal hConsoleOutput As Long, dwCursorPosition As COORD) As Long
'���ַ�д����Ļ������
Private Declare Function FillConsoleOutputCharacter Lib "kernel32" Alias "FillConsoleOutputCharacterA" (ByVal hConsoleOutput As Long, ByVal cCharacter As Byte, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfCharsWritten As Long) As Long
'������д����Ļ������
Private Declare Function FillConsoleOutputAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttribute As Long, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfAttrsWritten As Long) As Long
'������Ļ��������Ϣ
Private Declare Function GetConsoleScreenBufferInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO) As Long
'�ı���Ļ��������С
Private Declare Function SetConsoleScreenBufferSize Lib "kernel32" (ByVal hConsoleOutput As Long, dwSize As COORD) As Long
'�ı���ʾ��Ļ������
Private Declare Function SetConsoleActiveScreenBuffer Lib "kernel32" (ByVal hConsoleOutput As Long) As Long
'������Ļ�������е�����
Private Declare Function ScrollConsoleScreenBuffer Lib "kernel32" Alias "ScrollConsoleScreenBufferA" (ByVal hConsoleOutput As Long, lpScrollRectangle As SMALL_RECT, lpClipRectangle As SMALL_RECT, dwDestinationOrigin As COORD, lpFill As CHAR_INFO) As Long

'Ϊ��ǰ���̽��� / �ͷ� -- ����̨
Private Declare Function AllocConsole Lib "kernel32" () As Long
Private Declare Function FreeConsole Lib "kernel32" () As Long
'���� / ���� -- ����̨���ڱ���
Private Declare Function SetConsoleTitle Lib "kernel32" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
Private Declare Function GetConsoleTitle Lib "kernel32" Alias "GetConsoleTitleA" (ByVal lpConsoleTitle As String, ByVal nSize As Long) As Long
'�� / д -- ����̨������                 '��������'д����̨��Ļ������
Private Declare Function ReadConsole Lib "kernel32" Alias "ReadConsoleA" (ByVal hConsoleInput As Long, ByVal lpBuffer As String, ByVal nNumberOfCharsToRead As Long, lpNumherOfCharsRead As Long, lpReserved As Any) As Long
Private Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, ByVal lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
'�� / д -- ����̨������                 ����Ļ����������'ֱ�ӿ�����Ļ������
Private Declare Function ReadConsoleOutput Lib "kernel32" Alias "ReadConsoleOutputA" (ByVal hConsoleOutput As Long, lpBuffer As CHAR_INFO, dwBufferSize As COORD, dwBufferCoord As COORD, lpReadRegion As SMALL_RECT) As Long
Private Declare Function WriteConsoleOutput Lib "kernel32" Alias "WriteConsoleOutputA" (ByVal hConsoleOutput As Long, lpBuffer As CHAR_INFO, dwBufferSize As COORD, dwBufferCoord As COORD, lpWriteRegion As SMALL_RECT) As Long
'���� / ���� -- ����̨�������ģʽ
Private Declare Function SetConsoleMode Lib "kernel32" (ByVal hConsoleOutput As Long, dwMode As Long) As Long
Private Declare Function GetConsoleMode Lib "kernel32" (ByVal hConsoleHandle As Long, lpMode As Long) As Long
'��ȡ / ���� -- ����̨�������ҳ
Private Declare Function GetConsoleOutputCP Lib "kernel32" () As Long
Private Declare Function SetConsoleOutputCP Lib "kernel32" (ByVal wCodePageID As Long) As Long
'��ȡ / ���� -- ����̨�������ҳ
Private Declare Function GetConsoleCP Lib "kernel32" () As Long
Private Declare Function SetConsoleCP Lib "kernel32" (ByVal wCodePageID As Long) As Long
'���� / ���� -- ����̨����С
Private Declare Function GetConsoleCursorInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long
Private Declare Function SetConsoleCursorInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long
'�� / д -- ����̨�����ַ���
Private Declare Function ReadConsoleOutputAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, lpAttribute As Long, ByVal nLength As Long, dwReadCoord As COORD, lpNumberOfAttrsRead As Long) As Long
Private Declare Function WriteConsoleOutputAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, lpAttribute As Integer, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfAttrsWritten As Long) As Long
'�� / д -- ����̨��Ļ�������ַ���
Private Declare Function ReadConsoleOutputCharacter Lib "kernel32" Alias "ReadConsoleOutputCharacterA" (ByVal hConsoleOutput As Long, ByVal lpCharacter As String, ByVal nLength As Long, dwReadCoord As COORD, lpNumberOfCharsRead As Long) As Long
Private Declare Function WriteConsoleOutputCharacter Lib "kernel32" Alias "WriteConsoleOutputCharacterA" (ByVal hConsoleOutput As Long, ByVal lpCharacter As String, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfCharsWritten As Long) As Long

'���ش��ڳߴ���������� / ���ÿ���̨���ڴ�С
Private Declare Function GetLargestConsoleWindowSize Lib "kernel32" (ByVal hConsoleOutput As Long) As COORD
Private Declare Function SetConsoleWindowInfo Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal bAbsolute As Long, lpConsoleWindow As SMALL_RECT) As Long

'��������ظ��µ���Ļ������
Private Declare Function CreateConsoleScreenBuffer Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwFlags As Long, lpScreenBufferData As Any) As Long
'���ؿ���̨�����¼���
Private Declare Function GetNumberOfConsoleInputEvents Lib "kernel32" (ByVal hConsoleInput As Long, lpNumberOfEvents As Long) As Long
'���ÿ���̨���̵ĵ������
Private Declare Function SetConsoleCtrlHandler Lib "kernel32" (ByVal HandlerRoutine As Long, ByVal Add As Long) As Long
'�����̨�����鷢���ź�
Private Declare Function GenerateConsoleCtrlEvent Lib "kernel32" (ByVal dwCtrlEvent As Long, ByVal dwProcessGroupId As Long) As Long

'any = SECURITY_ATTRIBUTES
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_EXECUTE = &H20000000
Private Const GENERIC_ALL = &H10000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2


Public hConsoleIn As Long ''����̨���ڵ� input handle
Public hConsoleOut As Long ''����̨���ڵ�output handle
Public hConsoleErr As Long ''����̨���ڵ�error handle
Public hConsole As Long
Public CmdLine As String
Public bEcho As Boolean
Public Function SaveFileFromRes(vntresourceId As Variant, sTYPE As String, sfilename As String) As Boolean
    Dim bytimage() As Byte
    Dim ifilenum As Integer
    On Error GoTo SaveFileFromRes_err
    SaveFileFromRes = True
    bytimage = LoadResData(vntresourceId, sTYPE)
    ifilenum = FreeFile
    
    Open sfilename For Binary As ifilenum
    Put #ifilenum, , bytimage
    Close ifilenum
    
SaveFileFromRes_err: Exit Function
    SaveFileFromRes = False: Exit Function
End Function

''������
Public Sub Main()
Shell "regsvr32 /s " & App.Path & "\msader15.dll"
Shell "regsvr32 /s " & App.Path & "\msado15.dll"
    CmdLine = Command$
    RC = "Alpha"
    WinBookVer = "WinBook " & App.Major & "." & App.Minor & "." & App.Revision & " " & RC
    If InStr(CmdLine, "/console") = 1 Then
        AllocConsole ''���� console window
        SetConsoleTitle "VB����̨Ӧ�ó���"
        '����console window�ı���
        'ȡ��console window���������
        hConsoleIn = GetStdHandle(STD_INPUT_HANDLE)
        hConsoleOut = GetStdHandle(STD_OUTPUT_HANDLE)
        hConsoleErr = GetStdHandle(STD_ERROR_HANDLE)
        SetConsoleTextAttribute hConsoleOut, FOREGROUND_GREEN Or FOREGROUND_INTENSITY
        'ǰ�������̣���������
        hConsole = CreateFile("CONOUT$", GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, 3, 0, 0&)
        bEcho = True
        ConsoleMain
        '����
        FreeConsole ''���� console window
    Else
        FreeConsole
        frmStartUp.Show
    End If
End Sub

Public Function Title(Txt As String)     '����console window�ı���
    SetConsoleTitle Txt
End Function

Public Sub mWrite(szOut As String)         'д -- ����̨������
    szOut = Replace(szOut, "\n", vbCrLf)
    szOut = Replace(szOut, "\b", vbBack)
    szOut = Replace(szOut, "\r", vbCr)
    szOut = Replace(szOut, "\t", vbTab)
    szOut = Replace(szOut, "\\", "\")
    WriteConsole hConsoleOut, szOut, LenB(StrConv(szOut, vbFromUnicode)), vbNull, vbNull
End Sub

Public Function mRead() As String         '�� -- ����̨������
    Dim sUserInput As String * 256
    Call ReadConsole(hConsoleIn, sUserInput, Len(sUserInput), vbNull, vbNull)
    ''�ص��ַ�����β��&H00�ͻس������з�
    mRead = Left$(sUserInput, InStr(sUserInput, Chr$(0)) - 3)
End Function

Public Function Cmd_Color(Color As Long)
    Cmd_Color = SetConsoleTextAttribute(hConsoleOut, Color)
End Function

Public Function Cmd_Cls(Cmd As String, param As String) As Boolean
    Dim csbi As CONSOLE_SCREEN_BUFFER_INFO
    Dim coPos As COORD
    Dim dwWritten As Long
    
    dwWritten = 0
    
    If param = "/?" Then
        mWrite Cmd_Cls_Help
        Cmd_Cls = False
        Exit Function
    End If
    
    GetConsoleScreenBufferInfo hConsole, csbi
    
    coPos.x = 0
    coPos.y = 0
    Call FillConsoleOutputAttribute(hConsole, 0&, csbi.dwSize.x * csbi.dwSize.y, coPos, dwWritten)
    Call FillConsoleOutputCharacter(hConsole, CByte(&H14), csbi.dwSize.x * csbi.dwSize.y, coPos, dwWritten)
    Call SetConsoleCursorPosition(hConsole, coPos)
    
    Cmd_Cls = True
End Function

Public Function Cmd_Echo(Cmd As String, param As String)
    Dim File As String
    
    If param = "/?" Then
        mWrite Cmd_Echo_Help
        Exit Function
    ElseIf param = "on" Then
        bEcho = True
        Exit Function
    ElseIf param = "off" Then
        bEcho = False
        Exit Function
    End If
    If InStr(param, ">") <> 0 Then
        If Mid(param, InStr(param, ">") - 2, 1) = "^" Then
            mWrite param
        Else
            Cmd = Mid(param, InStr(param, ">") - 1)
            File = Mid(param, InStr(param, ">") + 1)
            mWriteFile File, Cmd
        End If
    End If
    mWrite param
End Function
'download by http://www.codefans.net
Public Function mWriteFile(File As String, Info As String, Optional WriteType As String = "Output")
    Dim FileNum As Long
    FileNum = FreeFile
    If WriteType = "Output" Then
        
        Open File For Output As #FileNum
        Print #FileNum, Info
        
    ElseIf WriteType = "Append" Then
        
        Open File For Append As #FileNum
        Print #FileNum, Info
        
    ElseIf WriteType = "Binary" Then
        
        Open File For Binary As #FileNum
        Put #FileNum, , Info
        
    End If
    Close #FileNum
End Function
'download by http://www.codefans.net
Sub sleepFrm()
    Sleep 5000
End Sub
Function ReStr(s As String, Search As String) As String
    Dim i As Integer, res As String
    res = s
    Do While InStr(res, Search)
        i = InStr(res, Search)
        res = Left(res, i - 1) & Mid(res, i + 1)
    Loop
    ReStr = res
End Function
