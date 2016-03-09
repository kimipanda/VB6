VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows ±âº»°ª
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim FilePath$, File$
    
    'Directory path in txt
    FilePath = ""
    If Right(FilePath, 1) <> "\" Then: FilePath = FilePath & "\"
    
    File = Dir$(FilePath & "*.txt")
    If File = vbNullString Then: MsgBox "Does not exist a txt file.", vbExclamation, vbNullString: Exit Sub

    Dim Files$()
    
    Call DirectoryFiles(Files, File)
    Call BubbleSort(Files, 1) '**** StrLen = 'String Length ****
    
    For i = 0 To UBound(Files)
        Debug.Print Files(i)
    Next i
End Sub
Public Sub DirectoryFiles(ByRef Files$(), ByVal File$)
    Dim FilesIndex%: FilesIndex = 0

    Do While (File <> vbNullString)
        ReDim Preserve Files$(FilesIndex)
        
        Files(FilesIndex) = File
        
        FilesIndex = FilesIndex + 1
        File = Dir$
    Loop
End Sub
Public Sub BubbleSort(ByRef Arrtmp$(), StrLen%) '**** StrLen = 'String Length ****
    Dim i%, tmp1$, tmp2$, tmp3$

    'Set initial value
    Dim result As Boolean: result = True
    
    Do While (result = True)
        result = False
        For i = LBound(Arrtmp) To UBound(Arrtmp) - 1
            
            tmp1 = Arrtmp(i)
            tmp2 = Arrtmp(i + 1)
            
            If (Mid(tmp1, StrLen + 1, Len(tmp1)) > Mid(tmp2, StrLen + 1, Len(tmp2))) Then
            
              tmp3 = Arrtmp(i)
              
              'Swap
              Arrtmp(i) = Arrtmp(i + 1)
              Arrtmp(i + 1) = tmp3
              result = True
            End If
        Next i
    Loop
End Sub

