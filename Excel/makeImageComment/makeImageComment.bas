Attribute VB_Name = "Module1"
Option Explicit
'매크로 바로가기 키와 맞춰주세요
Const VK_CUSTOMKEY = &H51

'클립보드에 저장된 이미지를 jpg로 변환하기 위한 라이브러리들
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Integer) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (PicDesc As uPicDesc, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
'64bit로 넘어오면서 olepro32.dll -> oleaut32.dll 바뀜

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
'현재 선택된 셀 위치
Private Type cellPos
    row As Integer
    col As Integer
End Type
'키보드 입력
Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal_bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub makeImageComment()
Attribute makeImageComment.VB_ProcData.VB_Invoke_Func = "q\n14"
    '화면 전환 중지
    Application.ScreenUpdating = False
    '화면 캡쳐 진행
    captureProcess
    Dim cell As cellPos
    '현재 선택된 셀 위치 row col 가져오기
    cell = getActiveCell
    Dim tempPath As String
    tempPath = Environ("Temp")
    '파일 위치 지정(임시파일로)
    Dim fileName As String
    fileName = tempPath & "\" & cell.row & "_" & cell.col & ".jpg"
    '캡쳐한 화면 사진파일로 저장
    SavePic fileName
    '저장된 사진파일을 셀 위치 메모에 저장
    Insert_pic_to_Comments fileName, cell.row, cell.col
    '이제 필요없는 사진파일 삭제
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
    'wscript.shell과 keybd_event 참고자료
    'https://www.vbforums.com/showthread.php?277384-VB-Key-COnsts
    'https://chany1995.tistory.com/75
    'https://blog.daum.net/sadest/15853449
    Const KEYEVENTF_KEYUP = &H2
    Const VK_LWIN = &H5B
    Const VK_SHIFT = &H10
    Const VK_S = &H53
    '엑셀 실행속도가 워낙 빨라 지정된 단축키가 눌린 상태에서 win + shift + s 하다보니 동작 안함
    '따라서 엑셀의 지정된 단축키를 강제로 keyup 해줘야함
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
    
    '캡처 끝나기전에 이미지 저장하지 않게 하기 위해 임시 코드 화면을 가릴 수 있음
    Sleep 500
    WShell.Popup "캡쳐 후 확인", , "Check", vbOKOnly
    Set WShell = Nothing
    'MsgBox "캡쳐 후 확인", , "Excel"
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
    '오류가 발생해도 다음 코드 실행
    '기존에 메모가 없더라도 프로그램 중지가 아닌 계속 실행
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
    '접근권한이 없어 사진 파일이 안지워질 수 있으니 주의할 것
    fs.deletefile FilePathName
    Set fs = Nothing
End Sub
