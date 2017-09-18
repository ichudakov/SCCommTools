VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufSpaceBomb 
   Caption         =   "Space Bomb"
   ClientHeight    =   3615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6765
   OleObjectBlob   =   "ufSpaceBomb-backup.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufSpaceBomb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const bIsDebug As Boolean = False

Public StartLocX As Double
Public StartLocY As Double
Public StartLocZ As Double
Public BoxSize As Double
Public MaxScale As Double 'max permitted scale
Public MaxQueueCount As Long
Public MinQueueCount As Long
Public UnitConversionCoeff As Double
Public UnitDescription As String
Public BackupAfterBoxes As Long

Private sldVoidSpace As Acad3DSolid
Private sldVoidSpaceBackup As Acad3DSolid
Private cQueue As Collection
Private cOverflow As Collection
Private cFinishedAreas As Collection
Private minDwgExtents(0 To 2) As Double
Private maxDwgExtents(0 To 2) As Double

Private bIsStopExplosion As Boolean
Private bIsRunning As Boolean
Private bIsPaused As Boolean
Private CurExpDir As Integer
Private CurAreaX As Double
Private CurAreaY As Double
Private CurAreaZ As Double
Private CurAreaD As Double

Private Enum ExpansionDir
    ed_bottom
    ed_left
    ed_front
    ed_right
    ed_back
    ed_top
End Enum

Private Enum BoxDir
    bottom_11
    bottom_12
    bottom_13
    bottom_21
    bottom_22
    bottom_23
    bottom_31
    bottom_32
    bottom_33
    middle_11
    middle_12
    middle_13
    middle_21
    middle_22
    middle_23
    middle_31
    middle_32
    middle_33
    top_11
    top_12
    top_13
    top_21
    top_22
    top_23
    top_31
    top_32
    top_33
End Enum

Private Enum InterferenceType
    it_unknown
    it_none
    it_complete
    it_partial
    it_some
    it_error
End Enum

Private Function SpaceBomb(x As Double, y As Double, z As Double) As Double
    Dim box(0 To 8) As Double
    Dim LastBackupVolume As Double
            
    lblStatus.Caption = "Running..."
    SetUndo (False)
    
    If bIsPaused Then
        bIsPaused = False
    Else
        ThisDrawing.Utility.Prompt "" & vbLf & Time$ & " - Operation Started. Unit Size: " & BoxSize & vbLf
            
        DetermineDrawingExtents
    
        Set cQueue = New Collection
        Set cOverflow = New Collection
        Set cFinishedAreas = New Collection
    
        'the first box
        box(0) = x ' the X of the cube center
        box(1) = y ' the Y of the cube center
        box(2) = z 'the Z of the cube center
        box(3) = BoxSize 'the box size
        box(4) = 0 'the current scale
        box(5) = 0 'previous scale
        box(6) = BoxDir.middle_22 'the expansion direction
        box(7) = it_unknown 'volume interference
        box(8) = it_unknown 'object interference
        
        cQueue.Add box
    
        CurExpDir = ed_bottom
        CurAreaD = (2 ^ MaxScale) * BoxSize
        SetAreaByBox (box)
        UpdateAreaText
    
        Set sldVoidSpace = Nothing
        LastBackupVolume = 0
    End If
        
    While cQueue.Count > 0 And (Not bIsStopExplosion)
        If bIsPaused Then
            lblStatus.Caption = "Paused"
            bttnPause.Enabled = True
            Exit Function
        Else
            If cQueue.Count > 0 Then
                FillVoid
                BackupVoidSpace
            End If
            If cQueue.Count = 0 Then MoveQueue
                        
            lblQueue.Caption = cQueue.Count
            lblOverflow.Caption = cOverflow.Count
            lblExpansionDir.Caption = CStr(CurExpDir)
            UpdateAreaText
            
            If Not sldVoidSpace Is Nothing Then
                lblVolDiscovered.Caption = Round(sldVoidSpace.Volume * UnitConversionCoeff, 3)
                ThisDrawing.Utility.Prompt "" & vbLf & Time$ & " - Discovered Volume: " & lblVolDiscovered.Caption & " " & UnitDescription & vbLf
            End If
        End If
        DoEvents
    Wend
        
    If sldVoidSpace Is Nothing Then
        SpaceBomb = 0
    Else
        SpaceBomb = Round(sldVoidSpace.Volume / (10 ^ 9), 3)
    End If
    
    ThisDrawing.Utility.Prompt Time$ & " - Polishing edges..." & vbLf
    DoEvents
    
    PolishEdges
            
    Set cQueue = Nothing
    Set cOverflow = Nothing
    
    bIsStopExplosion = True
    bIsRunning = False
    Set sldVoidSpace = Nothing
    If Not sldVoidSpaceBackup Is Nothing Then
        sldVoidSpaceBackup.Delete
        Set sldVoidSpaceBackup = Nothing
    End If
        
    SetUndo True
    ThisDrawing.Utility.Prompt Time$ & " - Operation Complete. Final Volume: " & SpaceBomb & " " & UnitDescription & vbLf
    
    Unload Me
End Function

Private Sub BackupVoidSpace(Optional bIsForced As Boolean = False)
    If (Not sldVoidSpaceBackup Is Nothing) And (Not bIsForced) Then
        If (sldVoidSpace.Volume - sldVoidSpaceBackup.Volume) / (BoxSize ^ 3) < BackupAfterBoxes Then
            Exit Sub
        End If
    End If
    If Not sldVoidSpaceBackup Is Nothing Then
        sldVoidSpaceBackup.Delete
        Set sldVoidSpaceBackup = Nothing
    End If
    If Not sldVoidSpace Is Nothing Then
        Set sldVoidSpaceBackup = sldVoidSpace.Copy()
    End If
End Sub

Private Sub RestoreVoidSpace()
    If Not sldVoidSpaceBackup Is Nothing Then
        If Not sldVoidSpace Is Nothing Then
            sldVoidSpace.Delete
        End If
        Set sldVoidSpace = sldVoidSpaceBackup.Copy()
    End If
End Sub

Private Sub FillVoid()
    Dim box(0 To 8) As Double
    Dim var As Variant
    Dim sld As Variant
    Dim x As Double
    Dim y As Double
    Dim z As Double
    Dim d As Double
    Dim m As Double
    Dim mprev As Double
    Dim k As Double
    Dim i As Long
    Dim edir As BoxDir
    Dim sldTmp As Acad3DSolid
    Dim sldBox As Acad3DSolid
    Dim cBoxData As Collection
    Dim cBoxSolids As Collection
       
    i = 1
    
    x = cQueue.Item(i)(0)
    y = cQueue.Item(i)(1)
    z = cQueue.Item(i)(2)
    d = cQueue.Item(i)(3)
    m = cQueue.Item(i)(4)
    mprev = cQueue.Item(i)(5)
    edir = cQueue.Item(i)(6)
        
    'attempt an expansion by the sldBox
    box(0) = cQueue(i)(0)
    box(1) = cQueue(i)(1)
    box(2) = cQueue(i)(2)
    box(3) = cQueue(i)(3)
    box(4) = cQueue(i)(4)
    box(5) = cQueue(i)(5)
    box(6) = cQueue(i)(6)
    box(7) = cQueue(i)(7)
    box(8) = cQueue(i)(8)
    
    cQueue.Remove i
    
    If CheckVolumeInterference(box) = it_complete Then
        Exit Sub
    End If
        
    Select Case CheckObjectInterference(box)
        Case it_some:
            If m = 0 Then
                'minimal size box
                Set sldBox = DrawBox(box)
                UnionWithVoidSpace sldBox
            Else
                'the box can be split
                SplitUpBox x, y, z, d, m, mprev, edir
            End If
        
        Case it_none:
            If m = 0 Then
                Set sldBox = DrawBox(box)
                UnionWithVoidSpace sldBox
            Else
                Select Case CheckBoxInterferenceWithArea(box)
                    Case it_complete
                        Set sldBox = DrawBox(box)
                        UnionWithVoidSpace sldBox
                        
                        PurgeFromQueue cQueue, box
                        PurgeFromQueue cOverflow, box
                    
                    Case it_partial
                        SplitUpBox x, y, z, d, m, mprev, edir
                        Exit Sub
                    
                    Case it_none
                        're-add the box to the queue
                        AddToQueue box
                        Exit Sub
                End Select
            End If
            
            'spawn more boxes
            mprev = m
            m = m + 1
            
            If MaxScale >= 0 Then
                If m > MaxScale Then
                    m = mprev
                End If
            End If
            k = (2 ^ mprev + 2 ^ m) / 2
            
            box(3) = d: box(4) = m: box(5) = mprev: box(7) = it_unknown: box(8) = it_unknown
            
            If bIsDebug Then
                Set cBoxData = New Collection
            Else
                Set cBoxData = Nothing
            End If
            
            'bottom tier
            box(0) = x - d * k: box(1) = y - d * k: box(2) = z - d * k: box(6) = BoxDir.bottom_11
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            box(0) = x - d * k: box(1) = y: box(2) = z - d * k: box(6) = BoxDir.bottom_12
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            box(0) = x - d * k: box(1) = y + d * k: box(2) = z - d * k: box(6) = BoxDir.bottom_13
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            
            box(0) = x: box(1) = y - d * k: box(2) = z - d * k: box(6) = BoxDir.bottom_21
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            box(0) = x: box(1) = y: box(2) = z - d * k: box(6) = BoxDir.bottom_22
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            box(0) = x: box(1) = y + d * k: box(2) = z - d * k: box(6) = BoxDir.bottom_23
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            
            box(0) = x + d * k: box(1) = y - d * k: box(2) = z - d * k: box(6) = BoxDir.bottom_31
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            box(0) = x + d * k: box(1) = y: box(2) = z - d * k: box(6) = BoxDir.bottom_32
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            box(0) = x + d * k: box(1) = y + d * k: box(2) = z - d * k: box(6) = BoxDir.bottom_33
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            
            'middle tier
            box(0) = x - d * k: box(1) = y - d * k: box(2) = z: box(6) = BoxDir.middle_11
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            box(0) = x - d * k: box(1) = y: box(2) = z: box(6) = BoxDir.middle_12
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            box(0) = x - d * k: box(1) = y + d * k: box(2) = z: box(6) = BoxDir.middle_13
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            
            box(0) = x: box(1) = y - d * k: box(2) = z: box(6) = BoxDir.middle_21
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            'skipped middle_22
            'skipped middle_22
            box(0) = x: box(1) = y + d * k: box(2) = z: box(6) = BoxDir.middle_23
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            
            box(0) = x + d * k: box(1) = y - d * k: box(2) = z: box(6) = BoxDir.middle_31
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            box(0) = x + d * k: box(1) = y: box(2) = z: box(6) = BoxDir.middle_32
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            box(0) = x + d * k: box(1) = y + d * k: box(2) = z: box(6) = BoxDir.middle_33
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            
            'top tier
            box(0) = x - d * k: box(1) = y - d * k: box(2) = z + d * k: box(6) = BoxDir.top_11
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            box(0) = x - d * k: box(1) = y: box(2) = z + d * k: box(6) = BoxDir.top_12
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            box(0) = x - d * k: box(1) = y + d * k: box(2) = z + d * k: box(6) = BoxDir.top_13
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            
            box(0) = x: box(1) = y - d * k: box(2) = z + d * k: box(6) = BoxDir.top_21
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            box(0) = x: box(1) = y: box(2) = z + d * k: box(6) = BoxDir.top_22
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            box(0) = x: box(1) = y + d * k: box(2) = z + d * k:  box(6) = BoxDir.top_23
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            
            box(0) = x + d * k: box(1) = y - d * k: box(2) = z + d * k: box(6) = BoxDir.top_31
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            box(0) = x + d * k: box(1) = y: box(2) = z + d * k: box(6) = BoxDir.top_32
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            box(0) = x + d * k: box(1) = y + d * k: box(2) = z + d * k: box(6) = BoxDir.top_33
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            
            If Not cBoxData Is Nothing Then
                Set cBoxSolids = New Collection
                For Each var In cBoxData
                    cBoxSolids.Add DrawBox(var)
                Next var
                ThisDrawing.Utility.Prompt "FillVoid()" & vbLf
                DoEvents
                For Each sld In cBoxSolids
                    sld.Delete
                Next sld
                ThisDrawing.Utility.Prompt "FillVoid()" & vbLf
                DoEvents
                Set cBoxSolids = Nothing
                Set cBoxData = Nothing
            End If
    End Select
End Sub

Private Sub UnionWithVoidSpace(sldBox As Acad3DSolid)
    Dim sldBackup As Acad3DSolid
    
    If sldVoidSpace Is Nothing Then
        Set sldVoidSpace = sldBox
    Else
        If sldBox.Volume > 0 Then
            On Error GoTo errUnionSolids
            sldVoidSpace.Boolean acUnion, sldBox
            On Error GoTo 0
        End If
    End If

    Exit Sub

errUnionSolids:
    If Not sldBox Is Nothing Then sldBox.Delete
    RestoreVoidSpace
End Sub

Private Sub MoveQueue()
    Dim cnt As Long
    Dim i As Long
    Dim bIsFound As Boolean
    Dim area(0 To 2) As Double
               
    'determine the working queue size based on the size of the overflow
    cnt = 10000 / (cOverflow.Count + 1)
    If cnt > MaxQueueCount Then
        cnt = MaxQueueCount
    End If
    If cnt < MinQueueCount Then
        cnt = MinQueueCount
    End If
        
    lblQueueSize.Caption = cnt
            
    While cQueue.Count < cnt And cOverflow.Count > 0
        DoEvents
        If bIsStopExplosion Then
            Exit Sub
        End If
        
        bIsFound = False
        For i = 1 To cOverflow.Count
            If IsBoxCenterInsideArea(cOverflow(i)) Then
                bIsFound = True
                Exit For
            End If
        Next i
                      
        If bIsFound Then
            cQueue.Add cOverflow(i)
            cOverflow.Remove (i)
        Else
            'the current area no longer has any expansion boxes
            area(0) = CurAreaX: area(1) = CurAreaY: area(2) = CurAreaY
            cFinishedAreas.Add area
            If cQueue.Count > 0 Then
                'finish whatever is left in the queue
                Exit Sub
            End If
            
            'the queue is empty; move to different area
            If Not MoveAreaInExpansionDir() Then
                If cOverflow.Count > 0 Then
                    SetAreaByBox GetClosestBoxToStart(cOverflow)
                End If
            End If
        End If
    Wend
End Sub

Private Sub PurgeOverflowQueue()
    Dim i As Long
    Dim it As InterferenceType
    Dim box As Variant
    
    Exit Sub
    
    i = 1
    While i <= cOverflow.Count
        it = CheckVolumeInterference(cOverflow(i))
        If it = it_complete Then
            cOverflow.Remove i
        Else
            If it <> cOverflow(i)(7) Then
                box = cOverflow(i)
                box(7) = it
                cOverflow.Add box, , i
                cOverflow.Remove i + 1
            End If
            i = i + 1
        End If
    Wend
End Sub

Private Function GetClosestBoxToStart(cBoxes As Collection) As Variant
    Dim box As Variant
    Dim minDistBox As Variant
    Dim dist As Double
    Dim minDist  As Double
        
    minDist = -1
    For Each box In cBoxes
        dist = (((StartLocX - box(0)) / (maxDwgExtents(0) - minDwgExtents(0)))) ^ 2 + _
            (((StartLocY - box(1)) / (maxDwgExtents(1) - minDwgExtents(1)))) ^ 2 + _
            (((StartLocZ - box(2)) / (maxDwgExtents(2) - minDwgExtents(2)))) ^ 2
        If minDist < 0 Or dist < minDist Then
            minDist = dist
            minDistBox = box
        End If
    Next box
        
    GetClosestBoxToStart = minDistBox
End Function

Private Function GetAreaByBox(box As Variant) As Variant
    Dim x As Double
    Dim y As Double
    Dim z As Double
    Dim area(0 To 2) As Double
    
    x = minDwgExtents(0)
    Do
        If box(0) >= x - CurAreaD / 2 And box(0) <= x + CurAreaD / 2 Then
            Exit Do
        End If
        x = x + CurAreaD
        If x > maxDwgExtents(0) Then
            GetAreaByBox = Nothing
            Exit Function
        End If
    Loop While True
    
    y = minDwgExtents(1)
    Do
        If box(1) >= y - CurAreaD / 2 And box(1) <= y + CurAreaD / 2 Then
            Exit Do
        End If
        y = y + CurAreaD
        If y > maxDwgExtents(1) Then
            GetAreaByBox = Nothing
            Exit Function
        End If
    Loop While True
    
    z = minDwgExtents(2)
    Do
        If box(2) >= z - CurAreaD / 2 And box(2) <= z + CurAreaD / 2 Then
            Exit Do
        End If
        z = z + CurAreaD
        If z > maxDwgExtents(2) Then
            GetAreaByBox = Nothing
            Exit Function
        End If
    Loop While True
        
    area(0) = x: area(1) = y: area(2) = z
    
    GetAreaByBox = area
End Function

Private Function SetAreaByBox(box As Variant) As Boolean
    CurAreaX = minDwgExtents(0)
    Do
        If box(0) >= CurAreaX - CurAreaD / 2 And box(0) <= CurAreaX + CurAreaD / 2 Then
            Exit Do
        End If
        CurAreaX = CurAreaX + CurAreaD
        If CurAreaX > maxDwgExtents(0) Then
            SetAreaByBox = False
            Exit Function
        End If
    Loop While True
    
    CurAreaY = minDwgExtents(1)
    Do
        If box(1) >= CurAreaY - CurAreaD / 2 And box(1) <= CurAreaY + CurAreaD / 2 Then
            Exit Do
        End If
        CurAreaY = CurAreaY + CurAreaD
        If CurAreaY > maxDwgExtents(1) Then
            SetAreaByBox = False
            Exit Function
        End If
    Loop While True
    
    CurAreaZ = minDwgExtents(2)
    Do
        If box(2) >= CurAreaZ - CurAreaD / 2 And box(2) <= CurAreaZ + CurAreaD / 2 Then
            Exit Do
        End If
        CurAreaZ = CurAreaZ + CurAreaD
        If CurAreaZ > maxDwgExtents(2) Then
            SetAreaByBox = False
            Exit Function
        End If
    Loop While True
    
    SetAreaByBox = True
End Function

Private Function MoveAreaInExpansionDir() As Boolean
    Select Case CurExpDir
        Case ed_bottom
            CurAreaZ = CurAreaZ - CurAreaD
        Case ed_left
            CurAreaX = CurAreaX - CurAreaD
        Case ed_front
            CurAreaY = CurAreaY - CurAreaD
        Case ed_right
            CurAreaX = CurAreaX + CurAreaD
        Case ed_back
            CurAreaY = CurAreaY + CurAreaD
        Case ed_top
            CurAreaZ = CurAreaZ + CurAreaD
    End Select
    
    If CurAreaX < minDwgExtents(0) Or CurAreaX > maxDwgExtents(0) Or _
        CurAreaY < minDwgExtents(1) Or CurAreaY > maxDwgExtents(1) Or _
        CurAreaZ < minDwgExtents(2) Or CurAreaZ > maxDwgExtents(2) Then
        'the area is outside drawing extents
        MoveAreaInExpansionDir = False
    Else
        MoveAreaInExpansionDir = True
    End If
        
End Function

Private Sub UpdateAreaText()
    lblCurAreaX.Caption = CStr(Round(CurAreaX, 0))
    lblCurAreaY.Caption = CStr(Round(CurAreaY, 0))
    lblCurAreaZ.Caption = CStr(Round(CurAreaZ, 0))
End Sub

Private Function DrawBox(box As Variant) As Acad3DSolid
    Dim c(0 To 2) As Double
    Dim m As Double
    Dim d As Double
    
    c(0) = box(0): c(1) = box(1): c(2) = box(2)
    d = box(3)
    m = box(4)
    
    Set DrawBox = ThisDrawing.ModelSpace.AddBox(c, d * (2 ^ m), d * (2 ^ m), d * (2 ^ m))
    If bIsDebug Then
        ThisDrawing.Utility.Prompt "DrawBox()" & vbLf
        DoEvents
    End If
End Function

Private Sub PolishEdges()
    Dim i As Long
    Dim obj As Variant
    Dim objCopy As Variant
    Dim bIsInterference As Boolean
    
    If sldVoidSpace Is Nothing Then Exit Sub
    If sldVoidSpace.Volume = 0 Then Exit Sub
    'If bIsStopExplosion Then Exit Sub
    
    lblStatus.Caption = "Polishing edges..."
    
    For i = ThisDrawing.ModelSpace.Count - 1 To 0 Step -1
        lblQueue.Caption = i
        DoEvents
        Set obj = ThisDrawing.ModelSpace(i)
        If TypeName(obj) = "IAcad3DSolid" Then
            If obj.Volume > 0 Then
                If (Not obj Is sldVoidSpace) And (Not obj Is sldVoidSpaceBackup) Then
                    bIsInterference = False
                    On Error GoTo errPolishEdgesInterference
                    sldVoidSpace.CheckInterference obj, False, bIsInterference
                    On Error GoTo 0
                    If bIsInterference Then
                        BackupVoidSpace True
                        Set objCopy = obj.Copy()
                        On Error GoTo errPolishEdgesSubstract
                        sldVoidSpace.Boolean acSubtraction, objCopy
                        On Error GoTo 0
                    End If
                End If
            Else
                obj.Delete
            End If
        End If
    Next i
        
    Exit Sub
    
errPolishEdgesInterference:
    bIsInterference = False
    Resume Next
errPolishEdgesSubstract:
    RestoreVoidSpace
    If Not objCopy Is Nothing Then objCopy.Delete
    Resume Next
End Sub

Private Sub SplitUpBox(x As Double, y As Double, z As Double, d As Double, m As Double, mprev As Double, edir As BoxDir)
    Dim box(0 To 8) As Double
    Dim k As Double
    Dim b_StartLocX As Double
    Dim b_StartLocY As Double
    Dim b_StartLocZ As Double
    Dim n As Double
    Dim cBoxData As Collection
    Dim cBoxSolids As Collection
    Dim sld As Acad3DSolid
    Dim var As Variant
    Dim w As Double
    Dim wprev As Double
                
    If m < 1 Then Exit Sub
    
    If bIsDebug Then
        Set cBoxData = New Collection
        box(0) = x
        box(1) = y
        box(2) = z
        box(3) = d
        box(4) = m
        box(5) = mprev
        box(6) = edir
        box(7) = it_unknown
        box(8) = it_unknown
        cBoxData.Add box
    End If
    
    box(3) = d
    box(4) = 0
    box(5) = 0
    box(6) = edir
    box(7) = it_unknown
    box(8) = it_unknown
    
    k = (2 ^ m - 1) / 2
    w = (2 ^ m) * d
    wprev = (2 ^ mprev) * d
    
    Select Case edir
        'bottom
        Case bottom_11:
            box(0) = x + k * d: box(1) = y + k * d: box(2) = z + k * d
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            
        Case bottom_12:
            box(0) = x + k * d
            box(2) = z + k * d
            b_StartLocY = y - wprev / 2
            For n = 0 To 2 ^ mprev - 1
                box(1) = b_StartLocY + d / 2 + d * n
                AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            Next n
        Case bottom_13:
            box(0) = x + k * d: box(1) = y - k * d: box(2) = z + k * d
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
        Case bottom_21:
            box(1) = y + k * d
            box(2) = z + k * d
            b_StartLocX = x - wprev / 2
            For n = 0 To 2 ^ mprev - 1
                box(0) = b_StartLocX + d / 2 + d * n
                AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            Next n
        Case bottom_22:
            box(2) = z + k * d
            b_StartLocX = x - wprev / 2
            b_StartLocY = y - wprev / 2
            For n = 0 To 2 ^ mprev - 1
                For k = 0 To 2 ^ mprev - 1
                    box(0) = b_StartLocX + d / 2 + d * n
                    box(1) = b_StartLocY + d / 2 + d * k
                    AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
                Next k
            Next n
        Case bottom_23:
            box(1) = y - k * d
            box(2) = z + k * d
            b_StartLocX = x - wprev / 2
            For n = 0 To 2 ^ mprev - 1
                box(0) = b_StartLocX + d / 2 + d * n
                AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            Next n
        Case bottom_31:
            box(0) = x - k * d: box(1) = y + k * d: box(2) = z + k * d
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
        Case bottom_32:
            box(0) = x - k * d
            box(2) = z + k * d
            b_StartLocY = y - wprev / 2
            For n = 0 To 2 ^ mprev - 1
                box(1) = b_StartLocY + d / 2 + d * n
                AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            Next n
        Case bottom_33:
            box(0) = x - k * d: box(1) = y - k * d: box(2) = z + k * d
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            
        'middle
        Case middle_11:
            box(0) = x + k * d: box(1) = y + k * d
            b_StartLocZ = z - wprev / 2
            For n = 0 To 2 ^ mprev - 1
                box(2) = b_StartLocZ + d / 2 + d * n
                AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            Next n
        Case middle_12:
            box(0) = x + k * d
            b_StartLocY = y - wprev / 2
            b_StartLocZ = z - wprev / 2
            For n = 0 To 2 ^ mprev - 1
                For k = 0 To 2 ^ mprev - 1
                    box(1) = b_StartLocY + d / 2 + d * n
                    box(2) = b_StartLocZ + d / 2 + d * k
                    AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
                Next k
            Next n
        Case middle_13:
            box(0) = x + k * d: box(1) = y - k * d
            b_StartLocZ = z - wprev / 2
            For n = 0 To 2 ^ mprev - 1
                box(2) = b_StartLocZ + d / 2 + d * n
                AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            Next n
        Case middle_21:
            box(1) = y + k * d
            b_StartLocX = x - wprev / 2
            b_StartLocZ = z - wprev / 2
            For n = 0 To 2 ^ mprev - 1
                For k = 0 To 2 ^ mprev - 1
                    box(0) = b_StartLocX + d / 2 + d * n
                    box(2) = b_StartLocZ + d / 2 + d * k
                    AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
                Next k
            Next n
        Case middle_22:
            'do nothing, not supposed to be here
        Case middle_23:
            box(1) = y - k * d
            b_StartLocX = x - wprev / 2
            b_StartLocZ = z - wprev / 2
            For n = 0 To 2 ^ mprev - 1
                For k = 0 To 2 ^ mprev - 1
                    box(0) = b_StartLocX + d / 2 + d * n
                    box(2) = b_StartLocZ + d / 2 + d * k
                    AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
                Next k
            Next n
        Case middle_31:
            box(0) = x - k * d: box(1) = y + k * d
            b_StartLocZ = z - wprev / 2
            For n = 0 To 2 ^ mprev - 1
                box(2) = b_StartLocZ + d / 2 + d * n
                AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            Next n
        Case middle_32:
            box(0) = x - k * d
            b_StartLocY = y - wprev / 2
            b_StartLocZ = z - wprev / 2
            For n = 0 To 2 ^ mprev - 1
                For k = 0 To 2 ^ mprev - 1
                    box(1) = b_StartLocY + d / 2 + d * n
                    box(2) = b_StartLocZ + d / 2 + d * k
                    AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
                Next k
            Next n
        Case middle_33:
            box(0) = x - k * d: box(1) = y - k * d
            b_StartLocZ = z - wprev / 2
            For n = 0 To 2 ^ mprev - 1
                box(2) = b_StartLocZ + d / 2 + d * n
                AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            Next n
        
        'top
        Case top_11:
            box(0) = x + k * d: box(1) = y + k * d: box(2) = z - k * d
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
        Case top_12:
            box(0) = x + k * d
            box(2) = z - k * d
            b_StartLocY = y - wprev / 2
            For n = 0 To 2 ^ mprev - 1
                box(1) = b_StartLocY + d / 2 + d * n
                AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            Next n
        Case top_13:
            box(0) = x + k * d: box(1) = y - k * d: box(2) = z - k * d
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
        Case top_21:
            box(1) = y + k * d
            box(2) = z - k * d
            b_StartLocX = x - wprev / 2
            For n = 0 To 2 ^ mprev - 1
                box(0) = b_StartLocX + d / 2 + d * n
                AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            Next n
        Case top_22:
            box(2) = z - k * d
            b_StartLocX = x - wprev / 2
            b_StartLocY = y - wprev / 2
            For n = 0 To 2 ^ mprev - 1
                For k = 0 To 2 ^ mprev - 1
                    box(0) = b_StartLocX + d / 2 + d * n
                    box(1) = b_StartLocY + d / 2 + d * k
                    AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
                Next k
            Next n
        Case top_23:
            box(1) = y - k * d
            box(2) = z - k * d
            b_StartLocX = x - wprev / 2
            For n = 0 To 2 ^ mprev - 1
                box(0) = b_StartLocX + d / 2 + d * n
                AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            Next n
        Case top_31:
            box(0) = x - k * d: box(1) = y + k * d: box(2) = z - k * d
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
        Case top_32:
            box(0) = x - k * d
            box(2) = z - k * d
            b_StartLocY = y - wprev / 2
            For n = 0 To 2 ^ mprev - 1
                box(1) = b_StartLocY + d / 2 + d * n
                AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
            Next n
        Case top_33:
            box(0) = x - k * d: box(1) = y - k * d: box(2) = z - k * d
            AddToQueue box: If Not cBoxData Is Nothing Then cBoxData.Add box
    End Select
    
    'If Not cBoxData Is Nothing Then
    If False Then
        Set cBoxSolids = New Collection
        For Each var In cBoxData
            cBoxSolids.Add DrawBox(var)
        Next var
        ThisDrawing.Utility.Prompt "SplitUpBox()" & vbLf
        DoEvents
        For Each sld In cBoxSolids
            sld.Delete
        Next sld
        ThisDrawing.Utility.Prompt "SplitUpBox()" & vbLf
        DoEvents
        Set cBoxSolids = Nothing
        Set cBoxData = Nothing
    End If
End Sub

Private Sub AddToQueue(box() As Double)
    Dim edir As BoxDir
    Dim bIsFound As Boolean
    Dim area As Variant
        
    If IsBoxInsideDrawingExtents(box) Then
        Select Case CheckBoxInterferenceWithArea(box)
            Case it_complete:
                'check queues for presence of identicals
                bIsFound = CheckInQueue(cQueue, box)
                If Not bIsFound Then bIsFound = CheckInQueue(cOverflow, box)
                
                'check volume interference
                If Not bIsFound Then
                    box(7) = CheckVolumeInterference(box)
                    If box(7) = it_complete Then Exit Sub
                End If
                
                If Not bIsFound Then
                    box(8) = CheckObjectInterference(box)
                    Select Case box(8)
                        Case it_none
                            If box(4) > 0 Then
                                PurgeFromQueue cQueue, box
                                PurgeFromQueue cOverflow, box
                            End If
                        Case it_some
                            edir = box(6)
                            SplitUpBox box(0), box(1), box(2), box(3), box(4), box(5), edir
                            Exit Sub
                    End Select
                                                            
                    If cOverflow.Count = 0 Then
                        cOverflow.Add box
                    Else
                        cOverflow.Add box, , 1
                    End If
                    
                    lblOverflow.Caption = cOverflow.Count
                End If
            
            Case it_partial:
                edir = box(6)
                SplitUpBox box(0), box(1), box(2), box(3), box(4), box(5), edir
                Exit Sub
                
            Case it_none:
                bIsFound = CheckInQueue(cQueue, box)
                If Not bIsFound Then bIsFound = CheckInQueue(cOverflow, box)
                
                'check volume interference
                If Not bIsFound Then
                    box(7) = CheckVolumeInterference(box)
                    If box(7) = it_complete Then Exit Sub
                End If
                
                'check if the box falls inside a finished area
                If Not bIsFound Then
                    For Each area In cFinishedAreas
                            Select Case CheckBoxInterferenceWithArea(box, False, area)
                                Case it_complete
                                    Exit Sub
                                    
                                Case it_partial
                                    edir = box(6)
                                    SplitUpBox box(0), box(1), box(2), box(3), box(4), box(5), edir
                                    Exit Sub
                                
                                Case it_none
                                    'do nothing
                            End Select
                    Next area
                End If
                
                If Not bIsFound Then
                    box(8) = CheckObjectInterference(box)
                    Select Case box(8)
                        Case it_none
                            If box(4) > 0 Then
                                PurgeFromQueue cQueue, box
                                PurgeFromQueue cOverflow, box
                            End If
                        Case it_some
                            edir = box(6)
                            SplitUpBox box(0), box(1), box(2), box(3), box(4), box(5), edir
                            Exit Sub
                    End Select
                    
                    cOverflow.Add box
                    lblOverflow.Caption = cOverflow.Count
                End If
        End Select
    Else
        'the box is outside drawing extents
        If box(4) > 0 Then
            'if the box is larger than the minimum scale, break it up
            edir = box(6)
            SplitUpBox box(0), box(1), box(2), box(3), box(4), box(5), edir
        End If
    End If
End Sub

Private Function IsBoxInsideDrawingExtents(box As Variant) As Boolean
    Dim d As Double
    
    d = box(3) * (2 ^ box(4))
    If box(0) - d / 2 < minDwgExtents(0) Or box(0) + d / 2 > maxDwgExtents(0) Or _
        box(1) - d / 2 < minDwgExtents(1) Or box(1) + d / 2 > maxDwgExtents(1) Or _
        box(2) - d / 2 < minDwgExtents(2) Or box(2) + d / 2 > maxDwgExtents(2) Then
        IsBoxInsideDrawingExtents = False
    Else
        IsBoxInsideDrawingExtents = True
    End If
End Function

Private Function IsBoxCenterInsideArea(box As Variant, Optional bIsCurrentArea As Boolean = True, Optional area As Variant = Nothing) As Boolean
    Dim d As Double
    Dim x As Double
    Dim y As Double
    Dim z As Double
    
    If bIsCurrentArea Then
        x = CurAreaX
        y = CurAreaY
        z = CurAreaZ
    Else
        x = area(0)
        y = area(1)
        z = area(2)
    End If

    If (box(0) < x - CurAreaD / 2 Or box(0) > x + CurAreaD / 2) Or _
        (box(1) < y - CurAreaD / 2 Or box(1) > y + CurAreaD / 2) Or _
        (box(2) < z - CurAreaD / 2 Or box(2) > z + CurAreaD / 2) Then
        IsBoxCenterInsideArea = False
    Else
        IsBoxCenterInsideArea = True
    End If
End Function


Private Function CheckInQueue(ByRef cMyQueue As Collection, box As Variant, Optional SkipIndex As Long = -1) As Boolean
    Dim var As Variant
    Dim bx1 As Double
    Dim by1 As Double
    Dim bz1 As Double
    Dim bx2 As Double
    Dim by2 As Double
    Dim bz2 As Double
    Dim i As Long
    
    bx1 = box(0) - box(3) * (2 ^ box(4)) / 2
    by1 = box(1) - box(3) * (2 ^ box(4)) / 2
    bz1 = box(2) - box(3) * (2 ^ box(4)) / 2
    bx2 = box(0) + box(3) * (2 ^ box(4)) / 2
    by2 = box(1) + box(3) * (2 ^ box(4)) / 2
    bz2 = box(2) + box(3) * (2 ^ box(4)) / 2
    
    CheckInQueue = False
    
    i = 1
    For Each var In cMyQueue
        If i <> SkipIndex Then
            If bIsStopExplosion Then Exit Function
            DoEvents
            If var(0) - var(3) * (2 ^ var(4)) / 2 <= bx1 And _
                var(1) - var(3) * (2 ^ var(4)) / 2 <= by1 And _
                var(2) - var(3) * (2 ^ var(4)) / 2 <= bz1 And _
                var(0) + var(3) * (2 ^ var(4)) / 2 >= bx2 And _
                var(1) + var(3) * (2 ^ var(4)) / 2 >= by2 And _
                var(2) + var(3) * (2 ^ var(4)) / 2 >= bz2 Then
                CheckInQueue = True
                Exit For
            End If
        End If
        i = i + 1
    Next var
End Function

Private Function CheckBoxInterferenceWithArea(box As Variant, Optional bIsCurrentArea As Boolean = True, Optional area As Variant = Nothing) As InterferenceType
    Dim d As Double
    Dim x As Double
    Dim y As Double
    Dim z As Double
    
    If bIsCurrentArea Then
        x = CurAreaX
        y = CurAreaY
        z = CurAreaZ
    Else
        x = area(0)
        y = area(1)
        z = area(2)
    End If
    
    d = box(3) * (2 ^ box(4))
    If (box(0) + d / 2 < x - CurAreaD / 2 Or box(0) - d / 2 > x + CurAreaD / 2) Or _
        (box(1) + d / 2 < y - CurAreaD / 2 Or box(1) - d / 2 > y + CurAreaD / 2) Or _
        (box(2) + d / 2 < z - CurAreaD / 2 Or box(2) - d / 2 > z + CurAreaD / 2) Then
        CheckBoxInterferenceWithArea = it_none
    Else
        If (box(0) - d / 2 >= x - CurAreaD / 2 Or box(0) + d / 2 <= x + CurAreaD / 2) And _
            (box(1) - d / 2 >= y - CurAreaD / 2 Or box(1) + d / 2 <= y + CurAreaD / 2) And _
            (box(2) - d / 2 >= z - CurAreaD / 2 Or box(2) + d / 2 <= z + CurAreaD / 2) Then
            CheckBoxInterferenceWithArea = it_complete
        Else
            CheckBoxInterferenceWithArea = it_partial
        End If
    End If
End Function

Private Sub PurgeFromQueue(ByRef cMyQueue As Collection, box As Variant)
    Dim var As Variant
    Dim i As Long
    Dim bx1 As Double
    Dim by1 As Double
    Dim bz1 As Double
    Dim bx2 As Double
    Dim by2 As Double
    Dim bz2 As Double
    
    bx1 = box(0) - box(3) * (2 ^ box(4)) / 2
    by1 = box(1) - box(3) * (2 ^ box(4)) / 2
    bz1 = box(2) - box(3) * (2 ^ box(4)) / 2
    bx2 = box(0) + box(3) * (2 ^ box(4)) / 2
    by2 = box(1) + box(3) * (2 ^ box(4)) / 2
    bz2 = box(2) + box(3) * (2 ^ box(4)) / 2
    
    i = 1
    While i <= cMyQueue.Count
        If bIsStopExplosion Then Exit Sub
        DoEvents
        var = cMyQueue(i)
        If var(0) - var(3) * (2 ^ var(4)) / 2 >= bx1 And _
            var(1) - var(3) * (2 ^ var(4)) / 2 >= by1 And _
            var(2) - var(3) * (2 ^ var(4)) / 2 >= bz1 And _
            var(0) + var(3) * (2 ^ var(4)) / 2 <= bx2 And _
            var(1) + var(3) * (2 ^ var(4)) / 2 <= by2 And _
            var(2) + var(3) * (2 ^ var(4)) / 2 <= bz2 Then
            cMyQueue.Remove i
        Else
            i = i + 1
        End If
    Wend
End Sub

Private Function CheckVolumeInterference(box As Variant) As InterferenceType
    Dim sldBox As Acad3DSolid
    Dim sldInt As Acad3DSolid
    Dim bIsInterference As Boolean
    
    If sldVoidSpace Is Nothing Then
        CheckVolumeInterference = it_none
        Exit Function
    End If
    
    If box(7) <> it_unknown Then
        CheckVolumeInterference = box(7)
        Exit Function
    End If
    
    Set sldBox = DrawBox(box)
    On Error GoTo errCheckVolInterferenceNoSolid
    sldVoidSpace.CheckInterference sldBox, False, bIsInterference
    On Error GoTo 0
        
    If bIsInterference Then
        On Error GoTo errCheckVolInterferenceWithSolid
        Set sldInt = sldBox.CheckInterference(sldVoidSpace, True, bIsInterference)
        On Error GoTo 0
        If sldBox.Volume = sldInt.Volume Then
            CheckVolumeInterference = it_complete
        Else
            CheckVolumeInterference = it_partial
        End If
        If Not sldInt Is Nothing Then sldInt.Delete
    Else
        CheckVolumeInterference = it_none
    End If
    
    sldBox.Delete
    Exit Function
    
errCheckVolInterferenceNoSolid:
    If Not sldBox Is Nothing Then sldBox.Delete
errCheckVolInterferenceWithSolid:
    If Not sldInt Is Nothing Then sldInt.Delete
    CheckVolumeInterference = it_error
End Function

Private Function CheckObjectInterference(box As Variant) As InterferenceType
    Dim intf As Variant
    Dim obj As Variant
    Dim bIsInterference As Boolean
    Dim objMin As Variant
    Dim objMax As Variant
    Dim sldBox As Acad3DSolid
    Dim bx1 As Double
    Dim by1 As Double
    Dim bz1 As Double
    Dim bx2 As Double
    Dim by2 As Double
    Dim bz2 As Double
           
    If box(8) <> it_unknown Then
        CheckObjectInterference = box(8)
        Exit Function
    End If
    
    CheckObjectInterference = it_none
    
    bx1 = box(0) - box(3) * (2 ^ box(4)) / 2
    by1 = box(1) - box(3) * (2 ^ box(4)) / 2
    bz1 = box(2) - box(3) * (2 ^ box(4)) / 2
    bx2 = box(0) + box(3) * (2 ^ box(4)) / 2
    by2 = box(1) + box(3) * (2 ^ box(4)) / 2
    bz2 = box(2) + box(3) * (2 ^ box(4)) / 2
       
    For Each obj In ThisDrawing.ModelSpace
        If TypeName(obj) = "IAcad3DSolid" Then
            If obj.Volume > 0 Then
                If (Not sldVoidSpace Is obj) And (Not sldVoidSpaceBackup Is obj) Then
                    obj.GetBoundingBox objMin, objMax
                    If (bx1 <= objMin(0) And bx2 <= objMin(0)) Or _
                        (bx1 >= objMax(0) And bx2 >= objMax(0)) Or _
                        (by1 <= objMin(1) And by2 <= objMin(1)) Or _
                        (by1 >= objMax(1) And by2 >= objMax(1)) Or _
                        (bz1 <= objMin(2) And bz2 <= objMin(2)) Or _
                        (bz1 >= objMax(2) And bz2 >= objMax(2)) Then
                        'the box is outside of the object's extents
                        'do nothing
                    Else
                        'the box and the object may overlap
                        Set sldBox = DrawBox(box)
                        
                        On Error GoTo errCheckObjInterference
                        sldBox.CheckInterference obj, False, bIsInterference
                        On Error GoTo 0
                                                                        
                        If bIsInterference Then
                            CheckObjectInterference = it_some
                            sldBox.Delete
                            Exit Function
                        End If
                    End If
                End If
            Else
                obj.Delete
            End If
        End If
    Next obj
    
    sldBox.Delete
    Exit Function

errCheckObjInterference:
    If Not sldBox Is Nothing Then sldBox.Delete
    CheckObjectInterference = it_error
End Function

Private Sub DetermineDrawingExtents()
    Dim obj As Variant
    Dim objMin As Variant
    Dim objMax As Variant
    Dim i As Integer
    
    For i = 0 To 2
        minDwgExtents(i) = 0: maxDwgExtents(i) = 0
    Next i
    
    For Each obj In ThisDrawing.ModelSpace
        If TypeName(obj) = "IAcad3DSolid" Then
                obj.GetBoundingBox objMin, objMax
                For i = 0 To 2
                    If objMin(i) < minDwgExtents(i) Then
                        minDwgExtents(i) = objMin(i)
                    End If
                    If objMax(i) > maxDwgExtents(i) Then
                        maxDwgExtents(i) = objMax(i)
                    End If
                Next i
        End If
    Next obj
End Sub

Private Sub bttnDrawQueue_Click()
    Dim box As Variant
    Dim cBoxes As Collection
    
    Set cBoxes = New Collection
    For Each box In cOverflow
        cBoxes.Add DrawBox(box)
    Next box
End Sub

'form events
Private Sub bttnPause_Click()
    If Not bIsPaused Then
        bttnPause.Caption = "Resume"
        lblStatus.Caption = "Pausing..."
        bttnPause.Enabled = False
        bIsPaused = True
    Else
        bttnPause.Caption = "Pause"
        SpaceBomb 0, 0, 0
    End If
End Sub

Private Sub bttnStop_Click()
    If Not bIsStopExplosion Then Unload Me
End Sub

Private Sub UserForm_Activate()
    If Not bIsRunning Then
        bIsRunning = True
        bIsPaused = False
        lblVolDiscovered.Caption = "0.000"
        lblQueue.Caption = 0
        lblOverflow.Caption = 0
        lblQueueSize.Caption = MaxQueueCount
        lblExpansionDir.Caption = CStr(middle_22)
        lblCurAreaX.Caption = ""
        lblCurAreaY.Caption = ""
        lblCurAreaZ.Caption = ""
        lblUnitDesc.Caption = UnitDescription
        bttnDrawQueue.Visible = bIsDebug
                
        ThisDrawing.SendCommand "!(command)" & vbCr
        
        SpaceBomb StartLocX, StartLocY, StartLocZ
    End If
End Sub

Private Sub UserForm_Initialize()
    bIsRunning = False
    bIsStopExplosion = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Not bIsStopExplosion Then
        lblStatus.Caption = "Shutting down..."
        Cancel = 1
        CloseMode = 1
        bIsStopExplosion = True
        If bIsPaused Then
            Cancel = 0
            SpaceBomb 0, 0, 0
        End If
    Else
        If bIsRunning Then
            Cancel = 1
        End If
    End If
End Sub
