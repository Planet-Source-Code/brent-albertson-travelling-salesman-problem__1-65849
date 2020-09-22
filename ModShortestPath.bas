Attribute VB_Name = "ModShortestPath"
Option Explicit

Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Const Pi  As Double = 3.14159265358979
Public Const CRan  As Double = (Pi / 180)


Public Const Con_MinCost = 999999
Public Const Con_MaxCost = -999999

Public Type City
    x As Single
    Y As Single
    Visted As Boolean
    FromCentre As Single
End Type

Public SP_Citys() As City
Public SP_CityCount As Long

Public Type Route
    CityID As Long
End Type

Public SP_Routes() As Route
Public SP_RouteMax As Long
Public GL_Showpath As Boolean
Public CentreX As Single
Public CentreY As Single
Public CentreMassX As Single
Public CentreMassY As Single
Public frm As frmMain
Public ST_Time As Single
Public ED_time As Single

Public Sub Main()

    Set frm = New frmMain
    frm.DrawStyle = vbPixels
    frm.DrawWidth = 3
    frm.AutoRedraw = True
    frm.Width = 12800
    frm.Height = 9800
    CentreX = frm.Width / 2
    CentreY = frm.Height / 2
    frm.Show
    
End Sub

Public Function OpenTSp(sFileName As String) As String
Dim FN As Integer
Dim FileStr As String
On Error GoTo err1

   frm.LabStatus.Caption = "OpenTSp...": frm.LabStatus.Refresh
   FN = FreeFile
   Open sFileName For Binary Access Read As #FN
      FileStr = Space(LOF(FN))
      Get #FN, , FileStr
   Close #FN
   OpenTSp = FileStr
   
Exit Function
err1:
OpenTSp = Err.Description
End Function

Public Function LoadTSP(Sdata As String) As Boolean
Dim Rows() As String
Dim RowMax As Long
Dim R As Long
Dim RowStart As Long
Dim Fields() As String
Dim FieldMax As Long

   On Error GoTo err1

   frm.LabStatus.Caption = "LoadTSP...": frm.LabStatus.Refresh
   Rows = Split(Sdata, vbCrLf)
   RowMax = UBound(Rows)
   For R = 0 To RowMax
      If Rows(R) = "NODE_COORD_SECTION" Then
         RowStart = R + 1
         Exit For
      End If
   Next

   If RowStart > 0 Then
      SP_CityCount = RowMax - RowStart - 1
      ReDim SP_Citys(1 To SP_CityCount) As City
      For R = RowStart To RowMax - 2
         Fields = Split(Rows(R), " ")
         With SP_Citys(R - RowStart + 1)
         If Fields(0) = "EOF" Then Exit For
            .x = Fields(1)
            .Y = Fields(2)
         End With
      Next
   End If

   LoadTSP = True
Exit Function
err1:
   LoadTSP = False
   MsgBox Err.Description
End Function

Private Function GetDistance(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Single
    GetDistance = ((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2) ^ 0.5
End Function

Public Function MakeCitys(citys As Long) As Boolean
Dim XRnd As Single
Dim YRnd As Single
Dim C As Long

   SP_CityCount = citys
   ReDim SP_Citys(1 To SP_CityCount) As City
   For C = 1 To SP_CityCount
      XRnd = Rnd(1) * 11000
      YRnd = Rnd(1) * 7000
      With SP_Citys(C)
         .x = XRnd
         .Y = YRnd
      End With
   Next
   
   MakeCitys = True
   
End Function

Public Function MakeRoundCitys(citys As Long) As Boolean
Dim i As Long
Dim CAng As Single
Dim AngStep As Single
Dim dia As Long

   SP_CityCount = citys
   ReDim SP_Citys(1 To SP_CityCount) As City
   dia = 5000
   AngStep = 360 / SP_CityCount
   For i = 1 To SP_CityCount
      CAng = (AngStep * i) * CRan
      With SP_Citys(i)
         .x = (Cos(CAng) * dia / 2) + CentreX
         .Y = (Sin(CAng) * dia / 2) + CentreY
      End With
   Next
   
   MakeRoundCitys = True
    
End Function

'Draw The Cities First City is in Blue
Public Function ShowCitys() As Boolean
Dim C As Long

   frm.LabStatus.Caption = "ShowCitys...": frm.LabStatus.Refresh
   frm.Cls
   frm.DrawWidth = 5
   frm.ForeColor = vbBlue
   frm.PSet (SP_Citys(1).x, SP_Citys(1).Y)
   For C = 2 To SP_CityCount
      frm.ForeColor = vbBlack
      frm.CurrentX = SP_Citys(C).x
      frm.CurrentY = SP_Citys(C).Y + 5
      frm.Print C
      frm.ForeColor = vbRed
      frm.PSet (SP_Citys(C).x, SP_Citys(C).Y)
   Next
   
   ShowCitys = True
   
End Function


'Show the Cost
Public Function ShowCost() As Boolean

   frm.LabStatus.Caption = "ShowCost...": frm.LabStatus.Refresh
   frm.ForeColor = vbBlack
   frm.DrawWidth = 3
   frm.CurrentX = 11200
   frm.CurrentY = 8500
   frm.Print "TotalCost " & Format(GetRouteCost, "###,###,###.#0")
   
   ShowCost = True
   
End Function


Public Function ShowRouteTour() As Boolean
Dim R As Long
Dim CurrentX As Single
Dim CurrentY As Single
Dim TempX As Single
Dim TempY As Single
Dim TotalCost As Single

On Error GoTo err1

   frm.LabStatus.Caption = "ShowRoutes...": frm.LabStatus.Refresh
   ShowCitys
   frm.ForeColor = vbBlue
   frm.DrawWidth = 3
   CurrentX = SP_Citys(SP_Routes(1).CityID).x
   CurrentY = SP_Citys(SP_Routes(1).CityID).Y
   For R = 2 To SP_RouteMax
       With SP_Routes(R)
           TempX = SP_Citys(.CityID).x
           TempY = SP_Citys(.CityID).Y
           frm.Line (CurrentX, CurrentY)-(TempX, TempY)
           CurrentX = TempX
           CurrentY = TempY
       End With
   Next
   frm.Line (CurrentX, CurrentY)-(SP_Citys(SP_Routes(1).CityID).x, SP_Citys(SP_Routes(1).CityID).Y)
   ShowRouteTour = True
   
Exit Function
err1:
ShowRouteTour = False
End Function

'Calulate the Cost of the Path
Public Function GetRouteCost() As Single
Dim R As Long
Dim CurrentX As Single
Dim CurrentY As Single
Dim TotalCost As Single

   CurrentX = SP_Citys(SP_Routes(1).CityID).x
   CurrentY = SP_Citys(SP_Routes(1).CityID).Y
   For R = 1 To SP_CityCount
      With SP_Routes(R)
          TotalCost = TotalCost + GetDistance(CurrentX, CurrentY, SP_Citys(.CityID).x, SP_Citys(.CityID).Y)
         CurrentX = SP_Citys(.CityID).x
         CurrentY = SP_Citys(.CityID).Y
      End With
   Next
   TotalCost = TotalCost + GetDistance(CurrentX, CurrentY, SP_Citys(SP_Routes(1).CityID).x, SP_Citys(SP_Routes(1).CityID).Y)
   GetRouteCost = TotalCost
End Function

'Basic
'Start at The First City
'Calulate the distance to all other Cities Pick the Shorthess distance
Public Function NearestNeighbour() As Boolean
Dim C As Long
Dim R As Long
Dim CurrentX As Single
Dim CurrentY As Single
Dim TempDist As Single
Dim MinDist As Single
Dim TempID As Long

  
   For C = 1 To SP_CityCount: SP_Citys(C).Visted = False: Next
   ReDim SP_Routes(1 To SP_CityCount) As Route
   
   CurrentX = SP_Citys(1).x
   CurrentY = SP_Citys(1).Y
   SP_Citys(1).Visted = True
   For R = 1 To SP_CityCount
      MinDist = Con_MinCost
      TempID = 1
      For C = 1 To SP_CityCount
         If SP_Citys(C).Visted = False Then
            TempDist = GetDistance(CurrentX, CurrentY, SP_Citys(C).x, SP_Citys(C).Y)
            If TempDist < MinDist Then
               MinDist = TempDist
               TempID = C
            End If
         End If
      Next
      SP_Citys(TempID).Visted = True
      CurrentX = SP_Citys(TempID).x
      CurrentY = SP_Citys(TempID).Y
      SP_Routes(R).CityID = TempID
      SP_RouteMax = R
      If GL_Showpath Then ShowRouteTour
   Next
   
NearestNeighbour = True
Exit Function
err1:
NearestNeighbour = False
End Function


Public Function Improve1() As Boolean
Dim R As Long
Dim C As Long
Dim A_City As Long
Dim B_City As Long
Dim NewCost As Single
Dim T_COST As Single

   For R = 5 To SP_CityCount - 5
      T_COST = GetRouteCost
      For C = R - 4 To R + 4
         A_City = SP_Routes(R).CityID
         B_City = SP_Routes(C).CityID
         'Swap them
         SP_Routes(R).CityID = B_City
         SP_Routes(C).CityID = A_City
         'Test if woste put back
         NewCost = GetRouteCost
         If NewCost > T_COST Then
            SP_Routes(R).CityID = A_City
            SP_Routes(C).CityID = B_City
         Else
            T_COST = NewCost     ' was better so update cost
         End If
      Next
   Next
   Improve1 = True

Exit Function
err1:
   Improve1 = False
End Function

Public Function BEA_TSP() As Boolean
Dim C As Long
Dim R As Long
Dim SumX As Single
Dim SumY As Single
Dim CiD As Long

On Error GoTo err1

   For C = 1 To SP_CityCount: SP_Citys(C).Visted = False: Next
   ReDim SP_Routes(1 To SP_CityCount) As Route

   'Find the centre of cities
   For C = 1 To SP_CityCount
      SumX = SumX + SP_Citys(C).x
      SumY = SumY + SP_Citys(C).Y
   Next
   CentreMassX = SumX / SP_CityCount
   CentreMassY = SumY / SP_CityCount

   'Set the distance to the centre
   For C = 1 To SP_CityCount
      SP_Citys(C).FromCentre = GetDistance(CentreMassX, CentreMassY, SP_Citys(C).x, SP_Citys(C).Y)
   Next
   
   SP_RouteMax = 1
   CiD = GetFarCity
   SP_Citys(CiD).Visted = True
   SP_Routes(SP_RouteMax).CityID = CiD
   For C = 1 To SP_CityCount - 1
      CiD = GetFarCity
      InsertCity CiD
      If GL_Showpath Then ShowRouteTour
   Next
   
   BEA_TSP = True
   
Exit Function
err1:
BEA_TSP = False
End Function

Private Function GetFarCity() As Long
Dim C As Long
Dim MaxDis As Single

   MaxDis = Con_MaxCost
   For C = 1 To SP_CityCount
      If Not SP_Citys(C).Visted Then
         If SP_Citys(C).FromCentre > MaxDis Then
            MaxDis = SP_Citys(C).FromCentre
            GetFarCity = C
         End If
      End If
   Next
   
End Function

Private Function InsertCity(InsertCityID As Long) As Boolean
Dim R As Long
Dim WhereID As Long
Dim RouteNo As Long
Dim ICostMin As Single
Dim StartX As Single
Dim StartY As Single
Dim LastX As Single
Dim LastY As Single
Dim NewX As Single
Dim NewY As Single
Dim TotalCost As Single
Dim CurrentCost As Single
Dim ST As Long
Dim ED As Long
On Error GoTo err1

   'find cheapest insertion cost
   
   ICostMin = Con_MinCost
   RouteNo = SP_RouteMax + 1
   NewX = SP_Citys(InsertCityID).x
   NewY = SP_Citys(InsertCityID).Y
   
   For WhereID = 1 To SP_RouteMax
      If SP_RouteMax = 1 Then Exit For
      ST = WhereID
      ED = WhereID + 1
      If ED = (SP_RouteMax + 1) Then ED = 1

      With SP_Routes(ST)
         StartX = SP_Citys(.CityID).x
         StartY = SP_Citys(.CityID).Y
      End With
      With SP_Routes(ED)
         LastX = SP_Citys(.CityID).x
         LastY = SP_Citys(.CityID).Y
      End With
      'get the cost of the orginal path
      CurrentCost = GetDistance(StartX, StartY, LastX, LastY)
      'Get the cost of the new path with an city inserted
      TotalCost = (GetDistance(StartX, StartY, NewX, NewY) + GetDistance(NewX, NewY, LastX, LastY)) - CurrentCost

      'Get the Cost
      If TotalCost < ICostMin Then
         ICostMin = TotalCost
         RouteNo = WhereID + 1
      End If
   Next

'Insert Node at cheapest place
   For R = SP_RouteMax To RouteNo Step -1
      SP_Routes(R + 1).CityID = SP_Routes(R).CityID
   Next
   SP_RouteMax = SP_RouteMax + 1
   SP_Routes(RouteNo).CityID = InsertCityID
   SP_Citys(InsertCityID).Visted = True
   
'Still need to do some local smothing


   InsertCity = True
Exit Function
err1:
InsertCity = False
MsgBox Err.Description
End Function
