Attribute VB_Name = "ModRun"
'download by http://www.codefans.net
'Title ���� | ���ÿ���̨������� | Function Title(Txt As String)
'Echo ����  | �������̨��Ϣ     | Sub Echo(szOut As String)
'EchoGet    | �õ�����̨������Ϣ |  Function EchoGet() As String
'���� -- ����̨�������ģʽ   Function GetMode(dwMode As Long)
'���� -- ����̨�������ģʽ   Function SetMode(lpMode As Long)
'���� -- ����̨�������ҳ     Function GetOutputCP(dwMode As Long)
'���� -- ����̨�������ҳ     Function SetOutputCP(wCodePageID As Long)
'��ȡ -- ����̨�������ҳ     Function GetCP()
'���� -- ����̨�������ҳ     Function SetCP(wCodePageID As Long)
'��ȡ -- ����̨����С       Function GetCursor(By As CONSOLE_CURSOR_INFO)
'���� -- ����̨����С       Function SetCursor(Size As Long, bHide As Long)
'���ش��ڳߴ����������     Function GetMaxConsole(ByVal XY As COORD)
'д -- ���ÿ���̨���ڴ�С     Function SetConsoleWindow(Left As Long, Right As Long, Top As Long, Bottom As Long)
'���ؿ���̨�����¼���         Function GetEvents()
'�������̨���뻺����         Sub Cls()
'������Ļ�ı�����             Sub Color(Num As Long)
'������갴ť��               Function GetMouseButton()
'������갴ť��               Function Set_Cursor(Xx As Long, Yy As Long)
'�ı���Ļ��������С           Function ScreenSize(Xx As Long, Yy As Long)
'�ı���ʾ��Ļ������           Function SetScreen()
'CmdLine    | ������������
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
Public Function ConsoleMain()
    Dim mColors As Long
    Dim sInStr  As String
    Dim Tmp(2)  As String
    Dim lInStr As String
    Dim lang As String
    If Dir(App.Path & "\Settings.ini") = "" Then
        Open "Settings.ini" For Output As #1
        Print #1, "Language=English"
        lang = "English"
    ElseIf Dir(App.Path & "\Settings.ini") <> "" Then
        Open "Settings.ini" For Input As #2
        Input #2, lang
    End If
    If lang = "English" Then GoTo English
    '-----------------------
English:
    
    Title "Win Book"
    mWrite "Win Book " & "Ver " & App.Major & "." & App.Minor & "." & App.Revision & Space(1) & RC & "\n"
    mWrite "Program starting up...\n"
    Sleep 1000
    mWrite "Done.\n"
    mWrite "Please type 'help' to see help of the commands in Win Book Console program!\n"
Englishs:
    
    mWrite ">"
    sInStr = mRead()
    If LCase(sInStr) = "exit" Then Exit Function
    lInStr = LCase(sInStr)
    Select Case LCase(sInStr)
    Case "help":

        mWrite "exit  Exit Win Book\n"
        mWrite "gui  Quit Win Book and restart Win Book Gui \n"
        mWrite "help  Display help\n"
        mWrite "sys   Display the staus of your current operating system\n"
        mWrite "***** To see the usage of a command. Please add '/?' behind the command ****\n"
        
    Case Else
        mWrite "!!!Command Not Exist\n"
    End Select
    GoTo Englishs:
    
End Function


