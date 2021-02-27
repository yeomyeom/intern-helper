Attribute VB_Name = "Module1"
Option Explicit
'��ũ�� �ٷΰ��� Ű�� �����ּ���
Const VK_CUSTOMKEY = &H51

'Ŭ�����忡 ����� �̹����� jpg�� ��ȯ�ϱ� ���� ���̺귯����
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Integer) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (PicDesc As uPicDesc, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
'64bit�� �Ѿ���鼭 olepro32.dll -> oleaut32.dll �ٲ�

Private Type GUID
'\\ Declare a UDT to store a GUID for the IPicture OLE Interface
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
'\\ Declare a UDT to store the bitmap information
Private Type uPicDesc
    Size As Long
    Type As Long
    hPic As Long
    hPal As Long
End Type
Private Const CF_BITMAP = 2
Private Const PICTYPE_BITMAP = 1
'���� ���õ� �� ��ġ
Private Type cellPos
    row As Integer
    col As Integer
End Type
'Ű���� �Է�
Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal_bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub makeImageComment()
Attribute makeImageComment.VB_ProcData.VB_Invoke_Func = "q\n14"
    'ȭ�� ��ȯ ����
    Application.ScreenUpdating = False
    'ȭ�� ĸ�� ����
    captureProcess
    Dim cell As cellPos
    '���� ���õ� �� ��ġ row col ��������
    cell = getActiveCell
    Dim tempPath As String
    tempPath = Environ("Temp")
    '���� ��ġ ����(�ӽ����Ϸ�)
    Dim fileName As String
    fileName = tempPath & "\" & cell.row & "_" & cell.col & ".jpg"
    'ĸ���� ȭ�� �������Ϸ� ����
    SavePic fileName
    '����� ���������� �� ��ġ �޸� ����
    Insert_pic_to_Comments fileName, cell.row, cell.col
    '���� �ʿ���� �������� ����
    deletePic fileName
End Sub
Private Function getActiveCell() As cellPos
    Dim cell As cellPos
    With cell
        .row = ActiveCell.row
        .col = ActiveCell.Column
    End With
    getActiveCell = cell
End Function
Private Sub captureProcess()
    Dim WShell As Object
    Set WShell = CreateObject("WScript.Shell")
    'wscript.shell�� keybd_event �����ڷ�
    'https://www.vbforums.com/showthread.php?277384-VB-Key-COnsts
    'https://chany1995.tistory.com/75
    'https://blog.daum.net/sadest/15853449
    Const KEYEVENTF_KEYUP = &H2
    Const VK_LWIN = &H5B
    Const VK_SHIFT = &H10
    Const VK_S = &H53
    '���� ����ӵ��� ���� ���� ������ ����Ű�� ���� ���¿��� win + shift + s �ϴٺ��� ���� ����
    '���� ������ ������ ����Ű�� ������ keyup �������
    Const VK_CONTROL = &H11
    
    keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0
    keybd_event VK_CUSTOMKEY, 0, KEYEVENTF_KEYUP, 0
    
    'win + shift + s
    keybd_event VK_LWIN, 0, 0, 0
    keybd_event VK_SHIFT, 0, 0, 0
    keybd_event VK_S, 0, 0, 0
    
    keybd_event VK_LWIN, 0, KEYEVENTF_KEYUP, 0
    keybd_event VK_SHIFT, 0, KEYEVENTF_KEYUP, 0
    keybd_event VK_S, 0, KEYEVENTF_KEYUP, 0
    
    'ĸó ���������� �̹��� �������� �ʰ� �ϱ� ���� �ӽ� �ڵ� ȭ���� ���� �� ����
    Sleep 500
    WShell.Popup "ĸ�� �� Ȯ��", , "Check", vbOKOnly
    Set WShell = Nothing
    'MsgBox "ĸ�� �� Ȯ��", , "Excel"
End Sub
Private Sub SavePic(FilePathName As String)
    'http://www.program1472.com/bbs/board.php?bo_table=TB_03&wr_id=25&sca=vba&sst=wr_datetime&sod=desc&sop=and&page=1
    Dim IID_IDispatch As GUID
    Dim uPicinfo As uPicDesc
    Dim IPic As IPicture
    Dim hPtr As Long
    
    '\\ Copy Range to ClipBoard
    'SourceRange.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
    
    OpenClipboard 0
    hPtr = GetClipboardData(CF_BITMAP)
    CloseClipboard
    '\\ Create the interface GUID for the picture
    With IID_IDispatch
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    '\\ Fill uPicInfo with necessary parts.
    With uPicinfo
        .Size = Len(uPicinfo) '\\ Length of structure.
        .Type = PICTYPE_BITMAP '\\ Type of Picture
        .hPic = hPtr '\\ Handle to image.
        .hPal = 0 '\\ Handle to palette (if bitmap).
    End With
    '\\ Create the Range Picture Object
    OleCreatePictureIndirect uPicinfo, IID_IDispatch, True, IPic
    '\\ Save Picture Object
    SavePicture IPic, FilePathName
End Sub
Private Sub Insert_pic_to_Comments(FilePathName As String, row As Integer, col As Integer)
    Dim oldMemo As String
    Dim picW As Single
    Dim picH As Single
    '������ �߻��ص� ���� �ڵ� ����
    '������ �޸� ������ ���α׷� ������ �ƴ� ��� ����
    On Error Resume Next
    With Cells(row, col)
        oldMemo = .Comment.Text
        On Error GoTo 0
        If Len(oldMemo) = 0 Then oldMemo = ""
        .ClearComments
        .AddComment (oldMemo)
        .Comment.Shape.Fill.UserPicture (FilePathName)
        ActiveSheet.Pictures.Insert(FilePathName).Select
        picW = Selection.Width
        picH = Selection.Height
        Selection.Delete
        .Comment.Shape.LockAspectRatio = msoFalse
        .Comment.Shape.Width = picW
        .Comment.Shape.Height = picH
        .Comment.Shape.LockAspectRatio = msoTrue
    End With
End Sub
Private Sub deletePic(FilePathName As String)
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    '���ٱ����� ���� ���� ������ �������� �� ������ ������ ��
    fs.deletefile FilePathName
    Set fs = Nothing
End Sub
