VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      Height          =   3765
      Left            =   4080
      TabIndex        =   4
      Top             =   1560
      Width           =   1935
   End
   Begin VB.ListBox lstSearch 
      Height          =   1035
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Files to search for."
      Top             =   120
      Width           =   5895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close File Search:"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   6000
      Width           =   5895
   End
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   120
      MultiSelect     =   1  'Simple
      TabIndex        =   1
      ToolTipText     =   "Search result."
      Top             =   1560
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Begin Search:"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Width           =   5895
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      Caption         =   "Search results (Path)                                                         (File name)"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   4800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lCounter As Long
Private lTickCount As Long

'***********************************************************************************
'Command1_Click
'***********************************************************************************
Private Sub Command1_Click()
    
    Me.MousePointer = 11 '----------Set MousePointer
    lCounter& = 0 '-----------------Reset lCounter&
    lTickCount& = GetTickCount '----GetTickCount
    
'***********************************************************************************
'READ THIS BEFORE YOU START THIS APPLICATION.
'***********************************************************************************
'Don`t forget to set sDrive() Array, For Next Loop according to your System and sSearchFor$
'***********************************************************************************
'**********************************************************************************

    Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        
    Dim sSearchFor As String '------String to search for
        
    Dim sDrive(2) As String '-------You can use the FileSystemObject to return all the
        sDrive(0) = "c:\" '---------the Drives on your Hard drive (I`m just using an
        sDrive(1) = "d:\" '---------Array for speed).
        sDrive(2) = "e:\"
                               
    Dim x As Byte '-----------------Loop through drives
    Dim y As Byte
    
    lstSearch.Visible = False
                                            
    For x = 0 To 2 '----------------Loop through Array of drives
            
        subShowFolderList (sDrive(x))
    
        For y = 0 To lstSearch.ListCount - 1 'Loop through search list to check if File is in root of Drive
        
            lstSearch.ListIndex = (y)
                 
            sSearchFor$ = lstSearch.Text
        
            'Use FileSystemObject to see if File exists in Root of Drive
            If (fso.FileExists(sDrive(x) & sSearchFor$)) Then
                'Add to ListBox if return value is True
                List1.AddItem sDrive(x)
                List2.AddItem sSearchFor$
            End If
    
        Next y
            
    Next x
    
    lstSearch.Visible = True
    
    Me.MousePointer = 0 '-----------Reset MousePointer
                        '-----------Display Results
    MsgBox "Time Elapsed: " & Format(GetTickCount - lTickCount&, "###,000") & "ms" & "   " & "Folders searched: " & lCounter&, vbInformation + vbOKOnly, "File search by (Neil Etherington)"

End Sub

'***********************************************************************************
'User Defined Sub subShowFolderList
'***********************************************************************************
Private Sub subShowFolderList(sFolder As String)

    Dim lReturn As Long '-----------Search handle of specified Path
    Dim lNextFile As Long '---------Search handle of specified File
    Dim sPath As String '-----------Path to search
    Dim WFD As WIN32_FIND_DATA '----Set Variable WFD as Structure WIN32_FIND_DATA
    Dim sFileName As String '-------Strips the Filename of Chr$(34), used with WFD.cFileName
        
    sPath$ = (sFolder$ & "*.*") & Chr$(0)
    
    lReturn& = FindFirstFile(sPath$, WFD)
    
    Do
        'If we find a Directory then strip vbNullChar
        If (WFD.dwFileAttributes And vbDirectory) Then
            
            sFileName$ = ftnStripNullChar(WFD.cFileName)

            'Exclude "." and ".." Folders
            If sFileName$ <> "." And sFileName$ <> ".." Then

                Call subFileSearch(sFolder$ & sFileName$ & "\")
                
                'Folder counter
                lCounter& = lCounter& + 1
                
                'Loop through Sub Folders
                subShowFolderList (sFolder & sFileName$ & "\")
            
            End If
        
        End If
         
         'Find next Folder
         lNextFile& = FindNextFile(lReturn&, WFD)

    Loop Until lNextFile& = False
  
    'Close Handle
    lNextFile& = FindClose(lReturn&)

End Sub

'***********************************************************************************
'User Defined Sub subFileSearch, Checks Folder for Search list items
'***********************************************************************************
Private Sub subFileSearch(sPath As String)
    
    Dim fso As Object '-------------Use FilesystemObjetc to see if File Exists
        Set fso = CreateObject("Scripting.FilesystemObject")
    Dim x As Byte
    Dim sSearchFor() As String
        ReDim sSearchFor(lstSearch.ListCount - 1)

    'Search through search list-----------------------------------------------------
    For x = 0 To lstSearch.ListCount - 1
        
        lstSearch.ListIndex = (x)
        sSearchFor(x) = lstSearch.Text
        
        'Use FileSystemObject to see if File exists
        If (fso.FileExists(sPath$ & sSearchFor$(x))) Then
            'Add to ListBox if return value is True
            List1.AddItem sPath$
            List2.AddItem sSearchFor$(x)
        End If
    
    Next x
    '-------------------------------------------------------------------------------
                                                                                    
End Sub


'***********************************************************************************
'User Defined Function ftnStripNullChar
'***********************************************************************************
Private Function ftnStripNullChar(sInput As String)

    Dim iSearch As Integer
    
    'I`ve used "Instr" over "InstrRev" for backward compatibility
    iSearch% = InStr(1, sInput$, Chr$(0))
    
    If iSearch% > 0 Then
        ftnStripNullChar = Left(sInput$, iSearch% - 1)
    End If

End Function

Private Sub Command2_Click()
    Unload Me
    Set Form1 = Nothing
End Sub


'***********************************************************************************
'Add items to search for in lstSearch
'***********************************************************************************
Private Sub Form_Load()
    lstSearch.AddItem "info.txt"
    lstSearch.AddItem "pkzip25.exe"
    lstSearch.AddItem "winhelp.hlp"
End Sub
