VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TSP Testing"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   12645
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "show path"
      Height          =   375
      Left            =   11160
      TabIndex        =   14
      Top             =   6420
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ComPare"
      Height          =   435
      Left            =   11100
      TabIndex        =   13
      Top             =   5880
      Width           =   1515
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Text            =   "frmMain.frx":0000
      Top             =   7800
      Width           =   10935
   End
   Begin VB.CommandButton cmdClearRoutes 
      Caption         =   "Clear Routes"
      Height          =   435
      Left            =   11100
      TabIndex        =   11
      Top             =   3720
      Width           =   1515
   End
   Begin VB.TextBox txtFile 
      Height          =   345
      Left            =   11100
      TabIndex        =   8
      Text            =   "pr76.tsp"
      Top             =   1080
      Width           =   1515
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   435
      Left            =   11100
      TabIndex        =   7
      Top             =   180
      Width           =   1515
   End
   Begin VB.CommandButton cmdImProve 
      Caption         =   "ImProve"
      Height          =   435
      Left            =   11100
      TabIndex        =   6
      Top             =   5340
      Width           =   1515
   End
   Begin VB.CommandButton cmdPath3 
      Caption         =   "TSP_BEA"
      Height          =   435
      Left            =   11100
      TabIndex        =   5
      Top             =   4800
      Width           =   1515
   End
   Begin VB.TextBox txtCitys 
      Height          =   315
      Left            =   11220
      TabIndex        =   3
      Text            =   "100"
      Top             =   3240
      Width           =   1395
   End
   Begin VB.CommandButton cmdMakeCircle 
      Caption         =   "Make circle"
      Height          =   435
      Left            =   11100
      TabIndex        =   2
      Top             =   2340
      Width           =   1515
   End
   Begin VB.CommandButton cmdMake 
      Caption         =   "Make RND"
      Height          =   435
      Left            =   11100
      TabIndex        =   1
      Top             =   1800
      Width           =   1515
   End
   Begin VB.CommandButton cmdpath1 
      Caption         =   "Nearest Neighbour"
      Height          =   435
      Left            =   11100
      TabIndex        =   0
      Top             =   4380
      Width           =   1515
   End
   Begin VB.Label LabStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ready"
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   9060
      Width           =   12615
   End
   Begin VB.Label Label1 
      Caption         =   "File Name"
      Height          =   315
      Left            =   11100
      TabIndex        =   9
      Top             =   720
      Width           =   1395
   End
   Begin VB.Line Line1 
      X1              =   11040
      X2              =   11040
      Y1              =   60
      Y2              =   9120
   End
   Begin VB.Label Label2 
      Caption         =   "Citys"
      Height          =   255
      Left            =   11220
      TabIndex        =   4
      Top             =   2940
      Width           =   555
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   cmdMake_Click
End Sub

Private Sub cmdOpen_Click()
   If LoadTSP(OpenTSp(App.Path & "\" & txtFile.Text)) Then
      ShowCitys
   Else
       MsgBox "Erorr Load File"
   End If
End Sub

Private Sub cmdMake_Click()
   MakeCitys (txtCitys.Text)
   ShowCitys
End Sub

Private Sub cmdMakeCircle_Click()
   MakeRoundCitys (txtCitys.Text)
   ShowCitys
End Sub

Private Sub Check1_Click()
   GL_Showpath = IIf(Check1.Value = vbChecked, True, False)
End Sub

Private Sub cmdClearRoutes_Click()
   ShowCitys
End Sub

Private Sub cmdpath1_Click()
   LabStatus.Caption = "NearestNeighbour...": LabStatus.Refresh
   ST_Time = GetTickCount
   NearestNeighbour
   ShowRouteTour
   ShowCost
   ED_time = GetTickCount
   LabStatus.Caption = "Ready..." & "Time " & Format((ED_time - ST_Time) / 1000, "####.#0") & " s": LabStatus.Refresh
End Sub

Private Sub cmdPath3_Click()
   LabStatus.Caption = "BEA_TSP...": LabStatus.Refresh
   ST_Time = GetTickCount
   BEA_TSP
   ShowRouteTour
   ShowCost
   ED_time = GetTickCount
   LabStatus.Caption = "Ready..." & "Time " & Format((ED_time - ST_Time) / 1000, "####.#0") & " s": LabStatus.Refresh
End Sub

Private Sub cmdImProve_Click()
   LabStatus.Caption = "Improve1...": LabStatus.Refresh
   ST_Time = GetTickCount
   Improve1
   ShowRouteTour
   ShowCost
   ED_time = GetTickCount
   LabStatus.Caption = "Ready..." & "Time " & Format((ED_time - ST_Time) / 1000, "####.#0") & " s": LabStatus.Refresh
End Sub

Private Sub Command1_Click()
Dim cost1 As Single
Dim cost2 As Single
Dim Scost1 As Single
Dim Scost2 As Single
Dim i As Long
Dim C As Long

   Text1.Text = ""
   For C = 20 To 50 Step 2
      txtCitys.Text = C
      For i = 1 To 5
         MakeCitys (txtCitys.Text)
            
         NearestNeighbour
         'Improve1
         cost1 = GetRouteCost
   
         BEA_TSP
         'Improve1
         cost2 = GetRouteCost
      
         Text1.Text = Text1.Text & i & vbTab & "Cities = " & C & vbTab & "Path1 Cost = " & cost1 & vbTab & "Path3 Cost = " & cost2 & vbTab & 100 - ((cost2 / cost1) * 100) & " % Better" & vbCrLf
         Text1.SelStart = Len(Text1)
         Scost1 = Scost1 + cost1
         Scost2 = Scost2 + cost2
      Next
   Next
   Text1.Text = Text1.Text & "SUM Path1 Cost = " & Scost1 & vbTab & "Sum Path3 Cost = " & Scost2 & vbTab & 100 - ((Scost2 / Scost1) * 100) & " % Better" & vbCrLf
   Text1.SelStart = Len(Text1)
   
End Sub

