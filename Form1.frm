VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Excel Open"
   ClientHeight    =   2070
   ClientLeft      =   11850
   ClientTop       =   6660
   ClientWidth     =   4725
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   4725
   Begin VB.CheckBox Check1 
      Caption         =   "0320"
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "奐偔"
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet

Dim mastPath, truePathName, sofaPath, sheetNum, fso, f, myPath$, fileName, phName
Dim sonaPath As String

  Dim tw(9) As String
  Open "path.txt" For Input Access Read As #1
  I = 0
    Do While Not EOF(1)
        I = I + 1
        Line Input #1, inputstring
        tw(I) = inputstring
    Loop
    Close #1

'    MsgBox tw(1)
'    MsgBox tw(2)
'    MsgBox tw(3)
    

'巇條彂椶宆偺敾掕
filetype = Mid(Trim(Text1.Text), 3, 4)
If filetype = "MFSY" Then
    '廂擺戙峴
    sonaPath = tw(1)

ElseIf filetype = "MMIF" Then
    '儅僗僞
    sonaPath = tw(2)
ElseIf filetype = "MFML" Then
    'FC夛寁
    sonaPath = tw(4)
ElseIf filetype = "MFSI" Then
    'FC巇擖
    sonaPath = tw(5)
ElseIf filetype = "MFFD" Then
    'FC暔棳旓
    sonaPath = tw(6)
ElseIf filetype = "MFEN" Then
    '塩嬈擔曬
    sonaPath = tw(7)
Else
    '庴敪拲
    sonaPath = tw(3)
End If


'sonaPath = "D:\vss\015-SY17BNC5\03_PJ奐敪_枾\0305_嶌惉愝寁彂乮FC夛寁乯乮廂擺戙峴乯\01_僶僢僠婡擻愝寁彂丒儊僢僙乕僕堦棗\婡擻愝寁彂(嬈柋曇 MFSY.廂擺戙峴 MFSY0700.婇嬈岦偗攝怣僨乕僞嶌惉 4.僶僢僠巇條).xlsx"
'sonaPath = "D:\vss\015-SY17BNC5\03_PJ奐敪_枾\0305_嶌惉愝寁彂乮FC夛寁乯乮廂擺戙峴乯\01_僶僢僠婡擻愝寁彂丒儊僢僙乕僕堦棗"

ph = Mid(Trim(Text1.Text), 3, 8)
filetype = Mid(Trim(Text1.Text), 3, 4)

If Check1.Value = 1 Then
   file0320
Else

    Dim Folder() As String
    
    myPath = Dir(sonaPath, vbDirectory)
    
    Dim fs2, f2, f12, s2, sf2
         
         Set fs2 = CreateObject("Scripting.FileSystemObject")
         Set f2 = fs2.GetFolder(sonaPath)
         Set sf2 = f2.SubFolders
         For Each f12 In sf2
              
        If InStr(1, f12.Name, ph) Then
          truePathName = f12.Name
          Exit For
        End If
            
         Next
         
     fullPath = sonaPath & "\" & truePathName
        
    
    
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    For Each f In fso.GetFolder(fullPath).Files
        fileName = f.Name
    
        a = InStr(1, fileName, ph)
    
        If a <> 0 Then
            phName = f.Name
        End If
    
    Next
    
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(fullPath & "\" & phName, "3", , , , , , , , True)
    
    '僔乕僩偺敾抐
    For I = 1 To xlBook.Worksheets.count
    
        sheetValue = xlBook.Worksheets(I).Cells(7, 11).Value
    
        If sheetValue = Trim(Text1.Text) Then
    
            sheetNum = I
            Exit For
        End If
    Next
    
    If sheetNum <> "" Then
        xlBook.Worksheets(sheetNum).Select
        xlBook.Worksheets(sheetNum).Cells(7, 11).Select
    End If
    
    
    xlApp.Visible = True

End If


End Sub


Private Sub file0320()
Dim count As Long
Dim newArray()
Dim arrayCount As Long
Dim strFiles(10000) As String
On Error GoTo MyerrNum


  Dim tw(9) As String
  Open "path.txt" For Input Access Read As #1
  I = 0
    Do While Not EOF(1)
        I = I + 1
        Line Input #1, inputstring
        tw(I) = inputstring
    Loop
    Close #1
    
    
    FindPath tw(9), strFiles(), count
arrayCount = 0
' 乽xlsx乿埲奜懳徾傪嶍彍偝傟偰偄傞丅
For I = 0 To count
  If strFiles(I) <> "" Then
    If Right(strFiles(I), 4) <> "xlsx" Or InStr(1, strFiles(I), "婡擻愝寁彂(嬈柋曇") = 0 Or InStr(1, strFiles(I), "儗價儏乕寢壥") <> 0 Or InStr(1, strFiles(I), "婡擻娭楢帒椏") <> 0 Then

        strFiles(I) = ""
    Else
        ReDim newArray(arrayCount)
        arrayCount = arrayCount + 1
    End If
  End If

Next

arrayCount = 0
For I = 0 To count
  If strFiles(I) <> "" Then
        newArray(arrayCount) = strFiles(I)
        arrayCount = arrayCount + 1
  End If
Next

ph = Mid(Trim(Text1.Text), 3, 8)
filetype = Mid(Trim(Text1.Text), 3, 4)

For w = 0 To arrayCount

    a = InStr(1, newArray(w), ph)

    If a <> 0 Then
               
        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Open(newArray(w), "3", , , , , , , , True)
        
        '僔乕僩偺敾抐
        For I = 1 To xlBook.Worksheets.count
        
            sheetValue = xlBook.Worksheets(I).Cells(7, 11).Value
        
            If sheetValue = Trim(Text1.Text) Then
        
                sheetNum = I
                Exit For
            End If
        Next
        
        If sheetNum <> "" Then
            xlBook.Worksheets(sheetNum).Select
            xlBook.Worksheets(sheetNum).Cells(7, 11).Select
        End If
        
        
        xlApp.Visible = True
        
        If sheetNum <> "" Then
            Exit For
        End If
        
        
    End If
    
Next

MyerrNum:
 If Err.Number = 1004 Then
    MsgBox "愝寁彂傕奐偐傟偨両"
 End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    Command1_Click
End If
End Sub

''=============================================
''名称: FindPath
''作用: 查找指定文件夹下面的所有文件和其子目录下的文件
''参数：strPath 僼傽僀儖僷僗
''      strFiles 用于存查找结果的缓冲区，String 类型的动态数组，调用时事先初始化， 如Redim strFiles(0)
''      FileCount 僼傽僀儖COUNT
''=============================================
Public Sub FindPath(ByVal strPath As String, strFiles() As String, FileCount As Long)
Dim strDirs()   As String
Dim strResult   As String
Dim FileLimit   As Long
Dim dirLimit    As Long
Dim dirCount    As Long
Dim I           As Long
    
    FileLimit = UBound(strFiles) + 1
    dirLimit = 0
    If Right$(strPath, 1) <> "\" Then strPath = strPath & "\"
    strResult = Dir(strPath, vbDirectory + vbSystem + vbReadOnly + vbHidden + vbNormal + vbArchive)
    Do While Len(strResult) > 0
        If strResult <> "." And strResult <> ".." Then
            If (GetAttr(strPath & strResult) And vbDirectory) <> vbDirectory Then
                If FileCount >= FileLimit Then
                    ReDim Preserve strFiles(FileLimit + 10)
                    FileLimit = FileLimit + 10
                End If
                strFiles(FileCount) = strPath & strResult
                FileCount = FileCount + 1
            Else
                If dirCount >= dirLimit Then
                    ReDim Preserve strDirs(dirLimit + 10)
                    dirLimit = dirLimit + 10
                End If
                strDirs(dirCount) = strPath & strResult
                dirCount = dirCount + 1
            End If
        End If
        strResult = Dir(, vbDirectory + vbSystem + vbReadOnly + vbHidden + vbNormal + vbArchive)
    Loop
    
    For I = 0 To dirCount - 1
        Call FindPath(strDirs(I), strFiles, FileCount)
    Next I
End Sub


Public Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
    Me.WindowState = 1
End If

If KeyAscii = 11 Then
    MsgBox "OOOO"
    Me.WindowState = 0
End If

End Sub


Private Sub Form_Resize()
'MsgBox "resize"
'If Me.WindowState = 0 Then
'
'End If
End Sub
