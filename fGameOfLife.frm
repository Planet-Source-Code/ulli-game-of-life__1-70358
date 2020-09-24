VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fGameOfLife 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   10065
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   14025
   Icon            =   "fGameOfLife.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10065
   ScaleWidth      =   14025
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton btRestart 
      Caption         =   "Restart"
      Height          =   315
      Left            =   7530
      TabIndex        =   13
      Top             =   9525
      Width           =   840
   End
   Begin VB.CommandButton btInfo 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8640
      TabIndex        =   12
      Top             =   9525
      Width           =   330
   End
   Begin VB.CommandButton btNext 
      Caption         =   "Step"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6570
      TabIndex        =   11
      Top             =   9525
      Width           =   840
   End
   Begin VB.CommandButton btPause 
      Caption         =   "Pause"
      Height          =   315
      Left            =   5610
      TabIndex        =   10
      Top             =   9525
      Width           =   840
   End
   Begin MSComDlg.CommonDialog CDl 
      Left            =   9870
      Top             =   9420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.HScrollBar scrDelay 
      Height          =   225
      LargeChange     =   5
      Left            =   1995
      Max             =   0
      Min             =   50
      TabIndex        =   8
      Top             =   9435
      Value           =   25
      Width           =   2595
   End
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Ausgefüllt
      Height          =   9060
      Left            =   225
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   900
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   13560
   End
   Begin VB.PictureBox picView 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Ausgefüllt
      Height          =   9060
      Left            =   225
      OLEDropMode     =   1  'Manuell
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   900
      TabIndex        =   0
      Top             =   240
      Width           =   13560
   End
   Begin VB.Label lbCpS 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   3015
      TabIndex        =   17
      Top             =   9765
      Width           =   1260
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "CpS"
      Height          =   195
      Index           =   5
      Left            =   4335
      TabIndex        =   16
      Top             =   9765
      Width           =   285
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "FpS"
      Height          =   195
      Index           =   4
      Left            =   2340
      TabIndex        =   15
      Top             =   9765
      Width           =   285
   End
   Begin VB.Label lbFPS 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   2010
      TabIndex        =   14
      Top             =   9765
      Width           =   270
   End
   Begin VB.Image img 
      Height          =   630
      Left            =   195
      Picture         =   "fGameOfLife.frx":08CA
      Top             =   9345
      Width           =   675
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Speed"
      Height          =   195
      Index           =   3
      Left            =   1335
      TabIndex        =   9
      Top             =   9630
      Width           =   465
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Generation: "
      Height          =   195
      Index           =   2
      Left            =   11445
      TabIndex        =   7
      Top             =   9345
      Width           =   870
   End
   Begin VB.Label lbGen 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   13095
      TabIndex        =   6
      Top             =   9345
      Width           =   675
   End
   Begin VB.Label lbAvgAge 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   13095
      TabIndex        =   5
      Top             =   9795
      Width           =   675
   End
   Begin VB.Label lbActive 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   13095
      TabIndex        =   4
      Top             =   9570
      Width           =   675
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Average Age: "
      Height          =   195
      Index           =   1
      Left            =   11445
      TabIndex        =   3
      Top             =   9795
      Width           =   1020
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of living cells: "
      Height          =   195
      Index           =   0
      Left            =   11445
      TabIndex        =   2
      Top             =   9570
      Width           =   1590
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLoad 
         Caption         =   "Load Pattern from File..."
      End
      Begin VB.Menu mnuRandom 
         Caption         =   "Load Random Pattern"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuColors 
      Caption         =   "Colors"
      Begin VB.Menu mnuLiveColor 
         Caption         =   "Live"
      End
      Begin VB.Menu mnuDeadColor 
         Caption         =   "Dead"
      End
      Begin VB.Menu mnuBackColor 
         Caption         =   "Background"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResetColor 
         Caption         =   "Reset"
      End
   End
   Begin VB.Menu mnuUniverse 
      Caption         =   "Universe"
      Begin VB.Menu mnuSize 
         Caption         =   "90 x 60"
         Index           =   0
      End
      Begin VB.Menu mnuSize 
         Caption         =   "150 x 100"
         Index           =   1
      End
      Begin VB.Menu mnuSize 
         Caption         =   "180 x 120"
         Index           =   2
      End
      Begin VB.Menu mnuSize 
         Caption         =   "300 x 200"
         Index           =   3
      End
      Begin VB.Menu mnuSize 
         Caption         =   "450 x 300"
         Index           =   4
      End
   End
End
Attribute VB_Name = "fGameOfLife"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Type tCell
    Age(0 To 1)         As Long     'age of cell: this and next generation (alternates)
    Neighbors(0 To 1)   As Long     'number of living neighbors: this and next generation (alternates)
    TopNeighbor         As Long     'index to neighbor cell above
    LeftNeighbor        As Long     'index to neighbor cell left
    RightNeighbor       As Long     'index to neighbor cell right
    BottomNeighbor      As Long     'index to neighbor cell below
    X                   As Long     'top left corner of cell on screen
    Y                   As Long
End Type

Private Cells()         As tCell    'array of cells

Private CellsAcross     As Long
Private CellsDown       As Long
Private NumCells        As Long     'number of cells
Private CellSize        As Long     'size of cell on screen
Private Generation      As Long     'generation counter
Private LiveColor       As Long
Private DeadColor       As Long
Private Seed            As Long     'rnd seed
Private PrevTick        As Long     'for FpS
Private FPSCount        As Long
Private PerfFrq         As Currency
Private PerfCnt1        As Currency
Private PerfCnt2        As Currency
Private Paused          As Boolean
Private Desc            As String   'pattern description

Private Const TooBig As String = "Pattern is too big for this universe."

Private Sub Activate()

    If Paused Then
        btPause_Click
    End If
    DoEvents
    Sleep 600

End Sub

Private Sub btInfo_Click()

    If Desc = vbNullString Then
        Desc = "No Info available."
    End If
    MsgBox Desc, , Caption

End Sub

Private Sub btNext_Click()

    CreateNextGeneration

End Sub

Private Sub btPause_Click()

  'pause/continue execution

    Paused = Not Paused
    btNext.Enabled = Paused
    btPause.Caption = IIf(Paused, "Continue", "Pause")

End Sub

Private Sub btRestart_Click()

    Restart

End Sub

Private Sub CreateNextGeneration()

  'creates and displays the next generation

  Dim i             As Long
  Dim CurrGen       As Long
  Dim NextGen       As Long
  Dim Active        As Long
  Dim TotalAge      As Long
  Dim Color         As Long

    QueryPerformanceCounter PerfCnt1 'for cell timing
    Generation = Generation + 1
    NextGen = Generation And 1
    CurrGen = 1 - NextGen

    For i = 0 To NumCells - 1
        Cells(i).Neighbors(NextGen) = 0 'reset neighbors for next gen
    Next i

    For i = 0 To NumCells - 1

        With Cells(i)

            .Neighbors(NextGen) = .Neighbors(NextGen) + .Neighbors(CurrGen) 'neighbors for next gen

            'live and let die
            If (.Age(CurrGen) And .Neighbors(CurrGen) = 2) Or .Neighbors(CurrGen) = 3 Then 'still alive or just born
                .Age(NextGen) = .Age(CurrGen) + 1
                If .Age(NextGen) = 1 Then 'just born
                    UpdateVicinity i, NextGen, 1 'update neighbors
                    picBuffer.Line (.X, .Y)-(.X + CellSize, .Y + CellSize), LiveColor, BF 'draw live cell
                End If
                Active = Active + 1
                TotalAge = TotalAge + .Age(NextGen) - 1
              Else 'NOT (.AGE(CURRGEN)...
                .Age(NextGen) = 0
                If .Age(CurrGen) Then 'just died
                    UpdateVicinity i, NextGen, -1 'update neighbors
                    picBuffer.Line (.X, .Y)-(.X + CellSize, .Y + CellSize), DeadColor, BF 'draw dead cell
                End If
            End If

        End With 'CELLS(I)

    Next i

    'present backbuffer
    With picView
        BitBlt .hDC, 0, 0, .ScaleWidth, .ScaleHeight, picBuffer.hDC, 0, 0, vbSrcCopy
        .Refresh
    End With 'PICVIEW
    FPSCount = FPSCount + 1 'frame counter
    QueryPerformanceCounter PerfCnt2 'for cell timing

    'display stats
    lbGen = Generation
    lbActive = Active
    If Active Then
        lbAvgAge = Format$(Round(TotalAge / Active, 2), "#0.00")
      Else 'ACTIVE = FALSE/0
        lbAvgAge = "all dead"
        FPSCount = 0
        Generation = Generation - 1
    End If

End Sub

Private Sub CreateRandom()

  Dim i As Long

    Rnd -Seed
    CreateUniverse
    Caption = App.ProductName & " [Random]"
    For i = 0 To NumCells - 1
        Cells(i).Age(0) = Rnd
        UpdateVicinity i, 0, Cells(i).Age(0)
    Next i
    Desc = "Random Pattern"
    DisplayCurrent
    Activate

End Sub

Private Sub CreateUniverse()

  'creates the cells

  Dim i     As Long

    picBuffer.Cls
    picView.Cls
    Generation = 0
    NumCells = CellsAcross * CellsDown
    ReDim Cells(0 To NumCells - 1)
    CellSize = picView.ScaleWidth \ CellsAcross

    For i = 0 To NumCells - 1

        'positions and Neighbors
        With Cells(i)
            .X = (i Mod CellsAcross) * CellSize
            .Y = (i \ CellsAcross) * CellSize
            'Neighbors wrap around horizontally and vertically
            .LeftNeighbor = (i \ CellsAcross) * CellsAcross + ((i - 1 + CellsAcross) Mod CellsAcross)
            .RightNeighbor = (i \ CellsAcross) * CellsAcross + ((i + 1) Mod CellsAcross)
            .TopNeighbor = (i + NumCells - CellsAcross) Mod NumCells
            .BottomNeighbor = (i + NumCells + CellsAcross) Mod NumCells
        End With 'CELLS(I)

    Next i

    CellSize = CellSize - 2 'so that a little background remains between drawn cells
    Desc = vbNullString

End Sub

Private Sub DisplayCurrent()

  Dim i     As Long

    For i = 0 To NumCells - 1
        With Cells(i)
            If .Age(Generation And 1) Then
                picView.Line (.X, .Y)-(.X + CellSize, .Y + CellSize), LiveColor, BF
                picBuffer.Line (.X, .Y)-(.X + CellSize, .Y + CellSize), LiveColor, BF
            End If
        End With 'CELLS(I)
    Next i
    lbGen = Generation
    lbActive = vbNullString
    lbAvgAge = vbNullString

End Sub

Private Sub Form_Initialize()

    InitCommonControls

End Sub

Private Sub Form_Load()

  Dim CurrTick As Long

    Caption = App.ProductName
    QueryPerformanceFrequency PerfFrq
    Randomize Timer
    mnuResetColor_Click

    Show
    DoEvents
    mnuSize_Click 2

    'life cycle
    Do

        If Not Paused Then

            'wait a little - then breed next generation
            CurrTick = scrDelay
            If CurrTick Then
                Sleep CurrTick * 10
            End If

            CreateNextGeneration

        End If

        'timing frames and cells per second
        CurrTick = GetTickCount
        If CurrTick >= PrevTick Then
            PrevTick = CurrTick + 1000
            lbFPS = FPSCount
            If FPSCount Then
                If PerfCnt1 < PerfCnt2 Then
                    lbCpS = Format$(PerfFrq * NumCells / (PerfCnt2 - PerfCnt1), "#,0")
                End If
              Else 'FPSCOUNT = FALSE/0
                lbCpS = 0
            End If
            FPSCount = 0
        End If

    Loop While DoEvents 'until form ist gone

End Sub

Private Sub GliderIni()

  'ini with a glider

  Dim i As Long
  Dim j As Long

    i = (NumCells + CellsAcross) / 2
    For j = 0 To 1
        Cells(i).Age(j) = 1
        UpdateVicinity i, j, 1

        Cells(i + CellsAcross + 1).Age(j) = 1
        UpdateVicinity i + CellsAcross + 1, j, 1

        Cells(i + 2 * CellsAcross - 1).Age(j) = 1
        UpdateVicinity i + 2 * CellsAcross - 1, j, 1

        Cells(i + 2 * CellsAcross).Age(j) = 1
        UpdateVicinity i + 2 * CellsAcross, j, 1

        Cells(i + 2 * CellsAcross + 1).Age(j) = 1
        UpdateVicinity i + 2 * CellsAcross + 1, j, 1
    Next j
    scrDelay = 25
    Desc = "Glider"
    DisplayCurrent

End Sub

Private Sub LoadArith(Filename As String)

  'loads an initial pattern

  'file format is as follows:

  'blank lines are ignored

  'lines starting with a semicolon or an apostophe are remarks and will be ignored
  'lines starting with a hash mark or a quote are descriptive text
  'lines starting with s= define the speed 1 - 50; illegal value are ignored
  'all other lines define the x y coordinates of a pattern cell

  'lines addressing positions outside the canvas are ignored

  Dim hFile     As Long
  Dim Line      As String
  Dim i         As Long
  Dim Pos       As Long
  Dim NoFit     As Boolean

    CreateUniverse
    scrDelay = 1
    hFile = FreeFile
    Open Filename For Input As hFile
    Do Until EOF(hFile)
        Line Input #hFile, Line
        If Len(Line) Then
            Select Case Left$(Line, 1)
              Case ";", "'"
                'do nothing
              Case "#", """"
                Desc = Desc & Mid$(Line, 2) & vbCrLf
              Case "s"
                On Error Resume Next
                    scrDelay = 51 - Val(Mid$(Line, 3)) 'speed
                On Error GoTo 0
              Case Else
                i = InStr(Line, " ")
                Pos = Val(Left$(Line, i - 1)) + CellsAcross / 2 + (Val(Mid$(Line, i)) + CellsDown / 2) * CellsAcross
                If Pos >= 0 And Pos < NumCells Then
                    Cells(Pos).Age(0) = 1
                    UpdateVicinity Pos, 0, 1
                  Else 'NOT XPOS... 'NOT POS...
                    NoFit = True
                End If
            End Select
        End If
    Loop
    Close hFile
    If NoFit Then
        MsgBox TooBig, vbExclamation, Caption
    End If
    DisplayCurrent

End Sub

Private Sub LoadPattern(Filename As String)

  'loads an initial pattern

  'file format is as follows:

  'first:
  'all lines are space suppressed

  'then:
  'blank lines are ignored
  'lines starting with a semicolon or an apostophe are remarks and will be ignored
  'lines starting with a hash mark or a quote are descriptive text
  'lines starting with x= define a new horizontal position, 1-based
  'lines starting with y= define a new  vertical  position, 1-based
  'line  starting with h= defines the width  of the pattern
  'line  starting with w= defines the height of the pattern
  'line  starting with s= defines the speed 1 - 50; illegal value are ignored
  'all other lines define a pattern line where the characters o, *, + or 1 define a living cell

  'lines addressing positions outside the canvas are ignored

  'example:

  '       ;gliders
  '       ;will place two gliders in positions (100 : 30) and (100 : 60)

  '       x = 100
  '       y = 30

  '       .o
  '       ..o
  '       ooo

  '       y = 60

  '       .o
  '       o
  '       ooo

  Dim hFile     As Long
  Dim LineRead  As String
  Dim Line      As String
  Dim i         As Long
  Dim xPos      As Long
  Dim yPos      As Long
  Dim NoFit     As Boolean

    CreateUniverse
    scrDelay = 1
    hFile = FreeFile
    Open Filename For Input As hFile
    Do Until EOF(hFile)
        Line Input #hFile, LineRead
        Line = LCase$(Replace$(LineRead, " ", ""))
        If Len(Line) Then
            Select Case Left$(Line, 1)
              Case ";", "'"
                'do nothing
              Case "#", """"
                Desc = Desc & Mid$(LineRead, 2) & vbCrLf
              Case "h"
                yPos = ((CellsDown - Val(Mid$(Line, 3))) \ 2) * CellsAcross
              Case "w"
                xPos = (CellsAcross - Val(Mid$(Line, 3))) \ 2
              Case "x"
                xPos = Val(Mid$(Line, 3)) - 1
              Case "y"
                yPos = (Val(Mid$(Line, 3)) - 1) * CellsAcross
              Case "s"
                On Error Resume Next
                    scrDelay = 51 - Val(Mid$(Line, 3)) 'speed
                On Error GoTo 0
              Case Else
                If xPos + yPos >= 0 Then
                    If xPos + yPos + Len(Line) < NumCells Then
                        For i = 1 To Len(Line)
                            Cells(yPos + xPos + i - 1).Age(0) = Sgn(InStr("o*+1", Mid$(Line, i, 1)))
                            UpdateVicinity yPos + xPos + i - 1, 0, Sgn(InStr("o*+1", Mid$(Line, i, 1)))
                        Next i
                        yPos = yPos + CellsAcross
                      Else 'NOT XPOS...
                        NoFit = True
                    End If
                  Else 'NOT XPOS...
                    NoFit = True
                End If
            End Select
        End If
    Loop
    Close hFile
    If NoFit Then
        MsgBox TooBig, vbExclamation, Caption
    End If
    DisplayCurrent

End Sub

Private Sub mnuBackColor_Click()

    With CDl
        On Error Resume Next
            .ShowColor
            If Err = 0 Then
                picBuffer.BackColor = .Color
                picView.BackColor = .Color
                DisplayCurrent
            End If
        On Error GoTo 0
    End With 'CDL

End Sub

Private Sub mnuDeadColor_Click()

    With CDl
        On Error Resume Next
            .ShowColor
            If Err = 0 Then
                DeadColor = .Color
            End If
        On Error GoTo 0
    End With 'CDL

End Sub

Private Sub mnuExit_Click()

    Unload Me

End Sub

Private Sub mnuLiveColor_Click()

    With CDl
        On Error Resume Next
            .ShowColor
            If Err = 0 Then
                LiveColor = .Color
                DisplayCurrent
            End If
        On Error GoTo 0
    End With 'CDL

End Sub

Private Sub mnuLoad_Click()

    With CDl
        .InitDir = App.Path & "\Patterns"
        .DialogTitle = "Enter/Select file to load..."
        .Filename = vbNullString
        .DefaultExt = ".TXT"
        .Filter = "GoL Pattern(*.TXT)|*.TXT|Arith Notation(*.LIF)|*.LIF|All Files(*.*)|*.*"
        .Flags = cdlOFNPathMustExist Or cdlOFNLongNames
        On Error Resume Next
            .ShowOpen
            If Err = 0 Then
                Seed = 0
                Caption = App.ProductName & " [" & CDl.FileTitle & "]"
                Select Case LCase$(Right$(.FileTitle, 3))
                  Case "txt"
                    LoadPattern .Filename
                  Case "lif"
                    LoadArith .Filename
                  Case Else
                    MsgBox "Unknown File Type:" & vbCrLf & vbCrLf & .FileTitle, vbExclamation, App.ProductName
                End Select
            End If
        On Error GoTo 0
    End With 'CDL

End Sub

Private Sub mnuRandom_Click()

  'create random colony

    Seed = GetTickCount
    CreateRandom

End Sub

Private Sub mnuResetColor_Click()

    LiveColor = vbYellow
    DeadColor = &H3030&
    picBuffer.BackColor = vbBlack
    picView.BackColor = vbBlack
    DisplayCurrent

End Sub

Private Sub mnuSize_Click(Index As Integer)

  Dim i As Long

    Select Case Index
      Case 0
        CellsAcross = 90 '10
        CellsDown = 60
      Case 1
        CellsAcross = 150 '6
        CellsDown = 100
      Case 2
        CellsAcross = 180 '5
        CellsDown = 120
      Case 3
        CellsAcross = 300 '3
        CellsDown = 200
      Case 4
        CellsAcross = 450 '2
        CellsDown = 300
    End Select
    For i = 0 To 4
        mnuSize(i).Checked = (i = Index)
    Next i
    Restart

End Sub

Private Sub picView_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim hFile As Long
  Dim e As Long

    hFile = FreeFile
    With Data
        On Error Resume Next
            Open .Files(1) For Input As hFile
            e = Err
        On Error GoTo 0
        If e Then
            MsgBox "Files only, please", vbExclamation, App.ProductName
          Else 'E = FALSE/0
            Close hFile
            CDl.Filename = .Files(1)
            Caption = App.ProductName & " [" & Mid$(.Files(1), InStrRev(.Files(1), "\") + 1) & "]"
            Select Case LCase$(Right$(.Files(1), 3))
              Case "txt"
                LoadPattern .Files(1)
              Case "lif"
                LoadArith .Files(1)
              Case Else
                Caption = App.ProductName
                MsgBox "Unknown File Type:" & vbCrLf & vbCrLf & .Files(1), vbCritical, App.ProductName
            End Select
        End If
    End With 'DATA

End Sub

Private Sub Restart()

    CreateUniverse
    With CDl
        If Seed Then
            CreateRandom
          ElseIf Len(.Filename) Then 'SEED = FALSE/0
            Select Case LCase$(Right$(.Filename, 3))
              Case "txt"
                LoadPattern .Filename
              Case "lif"
                LoadArith .Filename
            End Select
          Else 'LEN(.FILENAME) = FALSE/0
            GliderIni
        End If
    End With 'CDL
    Activate

End Sub

Private Sub UpdateVicinity(ByVal Idx As Long, ByVal NextGen As Long, ByVal IncDec As Long)

  'updates the vicinity of a cell

    With Cells(Idx) 'current cell

        Cells(.TopNeighbor).Neighbors(NextGen) = Cells(.TopNeighbor).Neighbors(NextGen) + IncDec 'current cell's top neighbor(north)

        Cells(.BottomNeighbor).Neighbors(NextGen) = Cells(.BottomNeighbor).Neighbors(NextGen) + IncDec 'current cell's bottom neighbor(south)

        With Cells(.LeftNeighbor) 'current cell's left neighbor(west)
            .Neighbors(NextGen) = .Neighbors(NextGen) + IncDec
            Cells(.TopNeighbor).Neighbors(NextGen) = Cells(.TopNeighbor).Neighbors(NextGen) + IncDec 'left neighbor's top neighbor(north west)
            Cells(.BottomNeighbor).Neighbors(NextGen) = Cells(.BottomNeighbor).Neighbors(NextGen) + IncDec 'left neighbor's bottom neighbor(south west)
        End With 'CELLS(.LEFTNeighbor)

        With Cells(.RightNeighbor) 'current cell's right neighbor(east)
            .Neighbors(NextGen) = .Neighbors(NextGen) + IncDec
            Cells(.TopNeighbor).Neighbors(NextGen) = Cells(.TopNeighbor).Neighbors(NextGen) + IncDec 'right neighbor's top neighbor(north east)
            Cells(.BottomNeighbor).Neighbors(NextGen) = Cells(.BottomNeighbor).Neighbors(NextGen) + IncDec 'right neighbor's bottom neighbor(south east)
        End With 'CELLS(.RIGHTNeighbor)

    End With 'CELLS(IDX)

End Sub

':) Ulli's VB Code Formatter V2.23.17 (2008-Apr-01 19:36)  Decl: 39  Code: 619  Total: 658 Lines
':) CommentOnly: 53 (8,1%)  Commented: 66 (10%)  Filled: 515 (78,3%)  Empty: 143 (21,7%)  Max Logic Depth: 7
