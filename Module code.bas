Option Explicit

Const Pi     As Double = 3.141593
Dim MI       As Single
Dim k        As Single, x  As Single, y  As Single, radius As Single, LUlin As Single
Dim kIa      As Single, kIb As Single, kIc As Single
Dim kUa      As Single, kUb As Single, kUc As Single, isTHAbcAbco As Byte
Dim xLine    As Single, yLine As Single
Dim pointXUa As Single, pointYUa As Single, pointXUb As Single, pointYUb As Single, pointXUc As Single, pointYUc As Single
Dim hIa      As Single, wIa As Single, hUa As Single, wUa As Single


' Change TT scheme

Sub setTT()
    If Range("isTT").Value = "" Then
        Call setBorderUTT("TTAco")

        Range("TTAbc").Font.Color        = RGB(70, 70, 70)
        Range("TTAco").Font.Color        = RGB(0, 0, 0)
        Range("TTAbc").Interior.Color    = xlNone
        Range("TTAbc").Borders.LineStyle = FALSE
        Range("isTT").Value              = " "
        Range("nameIb").Value            = "Ic=  "
        Range("colorIb").Interior.Color  = RGB(255, 0, 0)
        Range("nameIc").Value            = "Io=  "
        Range("colorIc").Interior.Color  = RGB(0, 176, 80)

        If Range("isIDegLC").Value = "LC" Then
            If Range("isTT").Value = "" Then
                Range("IaLC").Value = "L"
                Range("IbLC").Value = "L"
                Range("IcLC").Value = "C"
            ElseIf Range("isTT").Value = " " Then
                Range("IaLC").Value = "L"
                Range("IbLC").Value = "C"
                Range("IcLC").Value = "L"
            End If
        ElseIf Range("isIDegLC").Value = "Deg" Then
            Range("IaLC").Value = "°"
            Range("IbLC").Value = "°"
            Range("IcLC").Value = "°"
        End If
    Else:
        Call setBorderUTT("TTAbc")

        Range("TTAbc").Font.Color        = RGB(0, 0, 0)
        Range("TTAco").Font.Color        = RGB(70, 70, 70)
        Range("TTAco").Interior.Color    = xlNone
        Range("TTAco").Borders.LineStyle = FALSE
        Range("isTT").Value              = ""
        Range("nameIb").Value            = "Ib=  "
        Range("colorIb").Interior.Color  = RGB(0, 176, 80)
        Range("nameIc").Value            = "Ic=  "
        Range("colorIc").Interior.Color  = RGB(255, 0, 0)

        If Range("isIDegLC").Value = "LC" Then
            Range("IaLC").Value = "L"
            Range("IbLC").Value = "L"
            Range("IcLC").Value = "C"
        ElseIf Range("isIDegLC").Value = "Deg" Then
            Range("IaLC").Value = "°"
            Range("IbLC").Value = "°"
            Range("IcLC").Value = "°"
        End If
    End If

    Call updateVectorI
    Call animFigure
End Sub


' Change U reference

Sub setUref()
    Application.ScreenUpdating = FALSE

    If Range("isRefU").Value = "" Then
        Call setBorderUTT("refUab")

        Range("refUa").Font.Color        = RGB(70, 70, 70)
        Range("refUab").Font.Color       = RGB(0, 0, 0)
        Range("refUa").Interior.Color    = xlNone
        Range("refUa").Borders.LineStyle = FALSE
        Range("isRefU").Value            = " "
    Else:
        Call setBorderUTT("refUa")

        Range("refUa").Font.Color         = RGB(0, 0, 0)
        Range("refUab").Font.Color        = RGB(70, 70, 70)
        Range("refUab").Interior.Color    = xlNone
        Range("refUab").Borders.LineStyle = FALSE
        Range("isRefU").Value             = ""
    End If

    Call updateVectorI
    Call animFigure

    Application.ScreenUpdating = TRUE
End Sub


' Border, cell color Ua, Uac, TT scheme

Sub setBorderUTT(UTT As String)
    
    Range(UTT).Interior.Color = RGB(254, 250, 160)
    
    With Range(UTT).Borders(xlEdgeLeft)
        .LineStyle    = xlContinuous
        .ColorIndex   = 0
        .TintAndShade = 0
        .Weight       = xlThin
    End With

    With Range(UTT).Borders(xlEdgeTop)
        .LineStyle    = xlContinuous
        .ColorIndex   = 0
        .TintAndShade = 0
        .Weight       = xlThin
    End With

    With Range(UTT).Borders(xlEdgeBottom)
        .LineStyle    = xlContinuous
        .ColorIndex   = 0
        .TintAndShade = 0
        .Weight       = xlThin
    End With

    With Range(UTT).Borders(xlEdgeRight)
        .LineStyle    = xlContinuous
        .ColorIndex   = 0
        .TintAndShade = 0
        .Weight       = xlThin
    End With
End Sub

' Change current Mode Deg-LC

Sub setIDegLC()
    If Range("isIDegLC").Value = "Deg" Then
        Range("setIDegLC").Value = "L / C"
        Range("isIDegLC").Value  = "LC"

        ActiveSheet.Shapes("lableL").Visible   = TRUE
        ActiveSheet.Shapes("lableC").Visible   = TRUE
        ActiveSheet.Shapes("lable90").Visible  = FALSE
        ActiveSheet.Shapes("lable270").Visible = FALSE

        If Range("isTT").Value = "" Then
            Range("IaLC").Value = "L"
            Range("IbLC").Value = "L"
            Range("IcLC").Value = "C"
        ElseIf Range("isTT").Value = " " Then
            Range("IaLC").Value = "L"
            Range("IbLC").Value = "C"
            Range("IcLC").Value = "L"
        End If

        Range("UaLC").Value = "L"
        Range("UbLC").Value = "L"
        Range("UcLC").Value = "C"
    Else:
        Range("setIDegLC").Value = "Deg. °"
        Range("isIDegLC").Value  = "Deg"
        Range("IaLC").Value = "°"
        Range("IbLC").Value = "°"
        Range("IcLC").Value = "°"
        Range("UaLC").Value = "°"
        Range("UbLC").Value = "°"
        Range("UcLC").Value = "°"
        ActiveSheet.Shapes("lable90").Visible  = TRUE
        ActiveSheet.Shapes("lable270").Visible = TRUE
        ActiveSheet.Shapes("lableL").Visible   = FALSE
        ActiveSheet.Shapes("lableC").Visible   = FALSE
    End If

    Call animFigure
    Call getkI
    Call updateVectorI
    Call updateVectorU
End Sub


' Reset the current values ​depending on TT scheme

Sub resetI()
    If Range("isTT").Value = "" Then
        Range("currentIa").Value = 0
        Range("degIa").Value     = 0
        Range("currentIb").Value = 0
        Range("degIb").Value     = 120
        Range("currentIc").Value = 0
        Range("degIc").Value     = 120
        
        If Range("isIDegLC").Value = "LC" Then
            Range("IaLC").Value = "L"
            Range("IbLC").Value = "L"
            Range("IcLC").Value = "C"
        ElseIf Range("isIDegLC").Value = "Deg" Then
            Range("IaLC").Value = "°"
            Range("IbLC").Value = "°"
            Range("IcLC").Value = "°"
        End If
    Else:
        Range("currentIa").Value = 0
        Range("degIa").Value     = 0
        Range("currentIb").Value = 0
        Range("degIb").Value     = 120
        Range("currentIc").Value = 0
        Range("degIc").Value     = 120
        
        If Range("isIDegLC").Value = "LC" Then
            Range("IaLC").Value = "L"
            Range("IbLC").Value = "C"
            Range("IcLC").Value = "L"
        ElseIf Range("isIDegLC").Value = "Deg" Then
            Range("IaLC").Value = "°"
            Range("IbLC").Value = "°"
            Range("IcLC").Value = "°"
        End If
    End If

    Call animFigure
End Sub


' Change CL Ia

Sub setCLIa()
    If Range("isIDegLC").Value = "LC" Then
        Call setCLI("IaLC")
    End If
End Sub


' Change CL Ib

Sub setCLIb()
    If Range("isIDegLC").Value = "LC" Then
        Call setCLI("IbLC")
    End If
End Sub


' Change CL Ib

Sub setCLIc()
    If Range("isIDegLC").Value = "LC" Then
        Call setCLI("IcLC")
    End If
End Sub


' Set C or L current

Sub setCLI(cell     As String)
    If Range(cell).Value    = "L" Then
        Range(cell).Value   = "C"
    Else: Range(cell).Value = "L"
    End If

    Call updateVectorI
    Call animFigure
End Sub


'---- Support Ua for current. Start -----

' Add vector Ia

Sub drawVectorIa(opora, k As Single)
    Dim degIa As Single
    
    degIa = Range("degIa").Value
    
    If Range("IaLC").Value = "C" Then
        degIa = 360 - degIa - opora
    ElseIf Range("IaLC").Value = "L" Or Range("isIDegLC").Value = "Deg" Then
        degIa = degIa - opora
    End If
    
    Call setdrawVectorI(255, 255, 0, degIa, k, "Ia")
End Sub


' Add vector Ib

Sub drawVectorIb(opora, k As Single)
    Dim degIb As Single
    
    If Range("isTT").Value = "" Then
        degIb = Range("degIb").Value

        If Range("IbLC").Value = "C" Then
            degIb = 360 - degIb - opora
        ElseIf Range("IbLC").Value = "L" Or Range("isIDegLC").Value = "Deg" Then
            degIb = degIb - opora
        End If
    Else:
        degIb = Range("degIc").Value

        If Range("IcLC").Value = "C" Then
            degIb = 360 - degIb - opora
        ElseIf Range("IcLC").Value = "L" Or Range("isIDegLC").Value = "Deg" Then
            degIb = degIb - opora
        End If
    End If
    
    Call setdrawVectorI(51, 204, 51, degIb, k, "Ib")
End Sub


' Add vector Ic

Sub drawVectorIc(opora, k As Single)
    Dim degIc As Single
    
    If Range("isTT").Value = "" Then
        degIc = Range("degIc").Value

        If Range("IcLC").Value = "C" Then
            degIc = 360 - degIc - opora
        ElseIf Range("IcLC").Value = "L" Or Range("isIDegLC").Value = "Deg" Then
            degIc = degIc - opora
        End If
    Else:
        degIc = Range("degIb").Value

        If Range("IbLC").Value = "C" Then
            degIc = 360 - degIc - opora
        ElseIf Range("IbLC").Value = "L" Or Range("isIDegLC").Value = "Deg" Then
            degIc = degIc - opora
        End If
    End If
    
    Call setdrawVectorI(255, 0, 0, degIc, k, "Ic")
End Sub


'---- Support Ua for current. End -----

' Add vector I

Sub setdrawVectorI(R, G, B, degVector, k As Single, name As String)
    
    Dim selAdress As String, degDiagramm As Single
    
    selAdress = Selection.Address
    If k <> 0 Then
        ' bring countdown point to Vector (if -deg the current ahead, if +deg current behind)
        degDiagramm = 90 + degVector

        xLine = MI * k * radius * (Cos(degDiagramm * Pi / 180))
        yLine = MI * k * radius * (Sin(degDiagramm * Pi / 180))
        
        Shapes.AddLine(x, y, x - xLine, y - yLine).Select
        
        With Selection.ShapeRange
            .Line.ForeColor.RGB     = RGB(R, G, B)
            .Line.EndArrowheadStyle = msoArrowheadStealth
            .Line.EndArrowheadWidth = msoArrowheadWidthMedium
            .Line.DashStyle         = msoLineSolid
            .Line.Weight            = 2.8
            .name                   = name
        End With
    End If

    Range(selAdress).Select

    Call addLableI(name)
End Sub


' Redraw all vectors when something changes Update

Sub updateVectorI()
    
    ' removes drawn currents
    On Error Resume Next
    ActiveSheet.Shapes("Ia").Delete
    On Error Resume Next
    ActiveSheet.Shapes("Ib").Delete
    On Error Resume Next
    ActiveSheet.Shapes("Ic").Delete
    
    ' draw currents
    If Range("isRefU").Value = "" Then
        Call drawVectorIa(0, kIa)
        Call drawVectorIb(0, kIb)
        Call drawVectorIc(0, kIc)
    Else:
        Call drawVectorIa(30, kIa)
        Call drawVectorIb(30, kIb)
        Call drawVectorIc(30, kIc)
    End If
End Sub


' Redraw when change value and angle current

Private Sub Worksheet_Change(ByVal Target As Range)
    Application.ScreenUpdating = FALSE

    If Not Intersect(Target, Range("currentIa:currentIc, degIa:degIc")) Is Nothing Then
        Call getkI
        Call updateVectorI
    End If

    If Not Intersect(Target, Range("voltageUa:voltageUc, degUa:degUc")) Is Nothing Then
        Shapes("Background").ZOrder msoSendToBack
        Call getkU
        Call updateVectorU
    End If

    Application.ScreenUpdating = TRUE
End Sub


' Find the largest current to fit into the scale automatically

Sub getkI()
    Dim Ia As Single, Ib As Single, Ic  As Single, digit As Byte

    Ia = Range("currentIa").Value
    Ib = Range("currentIb").Value
    Ic = Range("currentIc").Value
    
    ' hide zeros after the decimal point depending on the number
    If InStr(Ia, ",") Then
        digit = Len(CStr(Ia)) - InStr(CStr(Ia), ",")

        If digit = 1 Then
            Range("currentIa").NumberFormat = "0.0"
        Else:
            Range("currentIa").NumberFormat = "0.00"
        End If
    Else:
        Range("currentIa").NumberFormat = "0"
    End If
    
    If InStr(Ib, ",") Then
        digit = Len(CStr(Ib)) - InStr(CStr(Ib), ",")

        If digit = 1 Then
            Range("currentIb").NumberFormat = "0.0"
        Else:
            Range("currentIb").NumberFormat = "0.00"
        End If
    Else:
        Range("currentIb").NumberFormat = "0"
    End If
    
    If InStr(Ic, ",") Then
        digit = Len(CStr(Ic)) - InStr(CStr(Ic), ",")

        If digit = 1 Then
            Range("currentIc").NumberFormat = "0.0"
        Else:
            Range("currentIc").NumberFormat = "0.00"
        End If
    Else:
        Range("currentIc").NumberFormat = "0"
    End If
    
    If Ia >= Ib And Ia >= Ic Then
        kIa = 1
        On Error Resume Next
        kIb = Ib / Ia
        On Error Resume Next
        kIc = Ic / Ia
    End If
    
    If Ib >= Ia And Ib >= Ic Then
        On Error Resume Next
        kIa = Ia / Ib
        kIb = 1
        On Error Resume Next
        kIc = Ic / Ib
    End If
    
    If Ic >= Ia And Ic >= Ib Then
        On Error Resume Next
        kIa = Ia / Ic
        On Error Resume Next
        kIb = Ib / Ic
        kIc = 1

        If (Ia = 0 And Ib = 0 And Ic = 0) Then
            kIc = 0
        End If
    End If
End Sub


' Add lable of current

Sub addLableI(name  As String)
    Dim valIa As Single, degIa As Single, isIaLC As String, isTT As String
    Dim valIb As Single, degIb As Single, isIbLC As String
    Dim valIc As Single, degIc As Single, isIcLC As String
    
    valIa  = Range("currentIa").Value
    degIa  = Range("degIa").Value
    isIaLC = Range("IaLC").Value
    valIb  = Range("currentIb").Value
    degIb  = Range("degIb").Value
    isIbLC = Range("IbLC").Value
    valIc  = Range("currentIc").Value
    degIc  = Range("degIc").Value
    isIcLC = Range("IcLC").Value
    
    isTT = Range("isTT").Value
    
    ' label is displayed correctly when change the TT
    If isTT = " " Then
        ActiveSheet.Shapes("lableIb").TextFrame.Characters.Text = "Ic"
        ActiveSheet.Shapes("lableIc").TextFrame.Characters.Text = "Io"

        Select Case name
            Case "Ib"
                name = "Ic"
            Case "Ic"
                name = "Ib"
        End Select
    Else:
        ActiveSheet.Shapes("lableIb").TextFrame.Characters.Text = "Ib"
        ActiveSheet.Shapes("lableIc").TextFrame.Characters.Text = "Ic"
    End If
    
    ' Lable current A
    
    If name = "Ia" And valIa <> 0 Then
        
        ActiveSheet.Shapes("lableIa").Visible = TRUE
        
        If Range("isIDegLC").Value = "Deg" Then
            If degIa <= 45 Then
                ActiveSheet.Shapes("lableIa").Left = x - xLine - (wIa / 2)
                ActiveSheet.Shapes("lableIa").Top  = y - yLine - hIa
            ElseIf degIa > 45 And degIa <= 135 Then
                ActiveSheet.Shapes("lableIa").Left = x - xLine
                ActiveSheet.Shapes("lableIa").Top  = y - yLine - (hIa / 2)
            ElseIf degIa > 135 And degIa < 225 Then
                ActiveSheet.Shapes("lableIa").Left = x - xLine - (wIa / 2)
                ActiveSheet.Shapes("lableIa").Top  = y - yLine
            ElseIf degIa >= 225 And degIa < 315 Then
                ActiveSheet.Shapes("lableIa").Left = x - xLine - wIa
                ActiveSheet.Shapes("lableIa").Top  = y - yLine - (hIa / 2)
            ElseIf degIa >= 315 Then
                ActiveSheet.Shapes("lableIa").Left = x - xLine - (wIa / 2)
                ActiveSheet.Shapes("lableIa").Top  = y - yLine - hIa
            End If
        Else:
            If degIa <= 45 Then
                ActiveSheet.Shapes("lableIa").Left = x - xLine - (wIa / 2)
                ActiveSheet.Shapes("lableIa").Top  = y - yLine - hIa
            ElseIf degIa > 45 And degIa <= 135 And isIaLC = "L" Then
                ActiveSheet.Shapes("lableIa").Left = x - xLine
                ActiveSheet.Shapes("lableIa").Top  = y - yLine - (hIa / 2)        
            ElseIf degIa > 45 And degIa <= 135 And isIaLC = "C" Then
                ActiveSheet.Shapes("lableIa").Left = x - xLine - wIa
                ActiveSheet.Shapes("lableIa").Top  = y - yLine - (hIa / 2)
            ElseIf degIa > 135 And degIa <= 180 Then
                ActiveSheet.Shapes("lableIa").Left = x - xLine - (wIa / 2)
                ActiveSheet.Shapes("lableIa").Top  = y - yLine
            End If
        End If
        
    ElseIf valIa = 0 Then
        ActiveSheet.Shapes("lableIa").Visible = FALSE
    End If
    
    ' Lable current B
    
    If name = "Ib" And valIb <> 0 Then
        
        ActiveSheet.Shapes("lableIb").Visible = TRUE
        
        If Range("isIDegLC").Value = "Deg" Then
            If degIb <= 45 Then
                ActiveSheet.Shapes("lableIb").Left = x - xLine - (wIa / 2)
                ActiveSheet.Shapes("lableIb").Top  = y - yLine - hIa
            ElseIf degIb > 45 And degIb <= 135 Then
                ActiveSheet.Shapes("lableIb").Left = x - xLine
                ActiveSheet.Shapes("lableIb").Top  = y - yLine - (hIa / 2)
            ElseIf degIb > 135 And degIb < 225 Then
                ActiveSheet.Shapes("lableIb").Left = x - xLine - (wIa / 2)
                ActiveSheet.Shapes("lableIb").Top  = y - yLine
            ElseIf degIb >= 225 And degIb < 315 Then
                ActiveSheet.Shapes("lableIb").Left = x - xLine - wIa
                ActiveSheet.Shapes("lableIb").Top  = y - yLine - (hIa / 2)
            ElseIf degIb >= 315 Then
                ActiveSheet.Shapes("lableIb").Left = x - xLine - (wIa / 2)
                ActiveSheet.Shapes("lableIb").Top  = y - yLine - hIa
            End If
        Else:
            If degIb <= 45 Then
                ActiveSheet.Shapes("lableIb").Left = x - xLine - (wIa / 2)
                ActiveSheet.Shapes("lableIb").Top  = y - yLine - hIa
            ElseIf degIb > 45 And degIb <= 135 And isIbLC = "L" Then
                ActiveSheet.Shapes("lableIb").Left = x - xLine
                ActiveSheet.Shapes("lableIb").Top  = y - yLine - (hIa / 2)
            ElseIf degIb > 45 And degIb <= 135 And isIbLC = "C" Then
                ActiveSheet.Shapes("lableIb").Left = x - xLine - wIa
                ActiveSheet.Shapes("lableIb").Top  = y - yLine - (hIa / 2)
            ElseIf degIb > 135 And degIb <= 180 Then
                ActiveSheet.Shapes("lableIb").Left = x - xLine - (wIa / 2)
                ActiveSheet.Shapes("lableIb").Top  = y - yLine
            End If
        End If
        
    ElseIf valIb = 0 Then
        ActiveSheet.Shapes("lableIb").Visible = FALSE
    End If
    
    ' Lable current C
    
    If name = "Ic" And valIc <> 0 Then
        
        ActiveSheet.Shapes("lableIc").Visible = TRUE
        
        If Range("isIDegLC").Value = "Deg" Then
            If degIc <= 45 Then
                ActiveSheet.Shapes("lableIc").Left = x - xLine - (wIa / 2)
                ActiveSheet.Shapes("lableIc").Top  = y - yLine - hIa
            ElseIf degIc > 45 And degIc <= 135 Then
                ActiveSheet.Shapes("lableIc").Left = x - xLine
                ActiveSheet.Shapes("lableIc").Top  = y - yLine - (hIa / 2)
            ElseIf degIc > 135 And degIc < 225 Then
                ActiveSheet.Shapes("lableIc").Left = x - xLine - (wIa / 2)
                ActiveSheet.Shapes("lableIc").Top  = y - yLine
            ElseIf degIc >= 225 And degIc < 315 Then
                ActiveSheet.Shapes("lableIc").Left = x - xLine - wIa
                ActiveSheet.Shapes("lableIc").Top  = y - yLine - (hIa / 2)
            ElseIf degIc >= 315 Then
                ActiveSheet.Shapes("lableIc").Left = x - xLine - (wIa / 2)
                ActiveSheet.Shapes("lableIc").Top  = y - yLine - hIa
            End If
        Else:
            If degIc <= 45 Then
                ActiveSheet.Shapes("lableIc").Left = x - xLine - (wIa / 2)
                ActiveSheet.Shapes("lableIc").Top  = y - yLine - hIa
            ElseIf degIc > 45 And degIc < 130 And isIcLC = "L" Then
                ActiveSheet.Shapes("lableIc").Left = x - xLine
                ActiveSheet.Shapes("lableIc").Top  = y - yLine - (hIa / 2)
            ElseIf degIc > 45 And degIc < 130 And isIcLC = "C" Then
                ActiveSheet.Shapes("lableIc").Left = x - xLine - wIa
                ActiveSheet.Shapes("lableIc").Top  = y - yLine - (hIa / 2)
            ElseIf degIc >= 130 And degIc <= 180 Then
                ActiveSheet.Shapes("lableIc").Left = x - xLine - (wIa / 2)
                ActiveSheet.Shapes("lableIc").Top  = y - yLine
            End If
        End If
        
    ElseIf valIc = 0 Then
        ActiveSheet.Shapes("lableIc").Visible = FALSE
    End If
End Sub


' Zoom in current

Sub scaleIncrI()
    If MI < 0.9 Then
        MI = MI + 0.1
    End If

    Call animFigure
    Call updateVectorI
End Sub


' Zoom out current

Sub scaleDecrI()
    If MI > 0.1 Then
        MI = MI - 0.1
    End If

    Call animFigure
    Call updateVectorI
End Sub


'------------Voltage--------------

' Change the voltage scheme ABC, ABC0

Sub setUabc0()
    If Range("isTH").Value = "0" Then
        Call setBorderUTT("THAbco")

        Range("THAbc").Font.Color        = RGB(70, 70, 70)
        Range("THAbco").Font.Color       = RGB(0, 0, 0)
        Range("THAbc").Interior.Color    = xlNone
        Range("THAbc").Borders.LineStyle = FALSE
        Range("isTH").Value   = "1"
        Range("nameUa").Value = "Uab= "
        Range("nameUb").Value = "Ubc= "
        Range("nameUc").Value = "Uca= "

        If Range("voltageUa").Value = 0 And Range("voltageUb").Value = 0 And Range("voltageUc").Value = 0 Then
            Call resetU
        End If
    Else:
        Call setBorderUTT("THAbc")

        Range("THAbc").Font.Color         = RGB(0, 0, 0)
        Range("THAbco").Font.Color        = RGB(70, 70, 70)
        Range("THAbco").Interior.Color    = xlNone
        Range("THAbco").Borders.LineStyle = FALSE
        Range("isTH").Value   = "0"
        Range("nameUa").Value = "Ua= "
        Range("nameUb").Value = "Ub= "
        Range("nameUc").Value = "Uc= "
    End If

    Call updateVectorU
    Call animFigure
End Sub


' Change CL Ua (Uab)

Sub setCLUa()
    Call setCLU("UaLC")
End Sub


' Change CL Ub (Ubc)

Sub setCLUb()
    Call setCLU("UbLC")
End Sub


' Change CL Uc (Uca)

Sub setCLUc()
    Call setCLU("UcLC")
End Sub


' Set C or L voltage

Sub setCLU(cell     As String)
    If Range(cell).Value = "L" Then
        Range(cell).Value = "C"
    Else: Range(cell).Value = "L"
    End If

    Call updateVectorU
    Call animFigure
End Sub


'---- Phase voltage vectors U-----

' Add vector Ua

Sub drawVectorUa(k  As Single)
    Dim degUa As Single
    
    degUa = Range("degUa").Value
    
    If Range("degUa").Value <> 0 And Range("UaLC").Value = "C" Then
        degUa = 360 - degUa
    Else:
        degUa = degUa
    End If
    
    Call setdrawVectorU(255, 255, 0, degUa, k, "Ua")
End Sub


' Add vector Ub

Sub drawVectorUb(k  As Single)
    Dim degUb As Single
    
    degUb = Range("degUb").Value
    
    If Range("degUb").Value <> 0 And Range("UbLC").Value = "C" Then
        degUb = 360 - degUb
    Else:
        degUb = degUb
    End If
    
    Call setdrawVectorU(51, 204, 51, degUb, k, "Ub")
End Sub


' Add vector Uc

Sub drawVectorUc(k  As Single)
    Dim degUc As Single
    
    degUc = Range("degUc").Value
    
    If Range("degUc").Value <> 0 And Range("UcLC").Value = "C" Then
        degUc = 360 - degUc
    Else:
        degUc = degUc
    End If
    
    Call setdrawVectorU(255, 0, 0, degUc, k, "Uc")
End Sub


' Add vector U in phase measurement

Sub setdrawVectorU(R, G, B, degVector, k As Single, name As String)
    
    Dim selAdress As String, degDiagramm As Single
    
    selAdress = Selection.Address
    
    If k <> 0 Then
        ' bring countdown point to Vector (if -deg the current ahead, if +deg current behind)
        degDiagramm = 90 + degVector

        xLine = k * radius * (Cos(degDiagramm * Pi / 180))
        yLine = k * radius * (Sin(degDiagramm * Pi / 180))
        
        Shapes.AddLine(x, y, x - xLine, y - yLine).Select
        
        With Selection.ShapeRange
            .Line.ForeColor.RGB     = RGB(R, G, B)
            .Line.EndArrowheadStyle = msoArrowheadStealth
            .Line.Weight            = 1.8
            .name                   = name
        End With
    End If

    Range(selAdress).Select

    Call addLableU(name)
End Sub


'--------Draw Ulin-------

' Phase Uab

Sub drawVectorUab()
    Dim selAdress As String
    
    selAdress = Selection.Address
    
    Shapes.AddLine(pointXUb, pointYUb, pointXUa, pointYUa).Select
    
    With Selection.ShapeRange
        .Line.ForeColor.RGB     = RGB(255, 255, 0)
        .Line.EndArrowheadStyle = msoArrowheadStealth
        .Line.Weight            = 1.8
        .name                   = "Uab"
    End With
    
    Range(selAdress).Select
End Sub


' Phase Ubc

Sub drawVectorUbc()
    Dim selAdress As String
    
    selAdress = Selection.Address
    
    Shapes.AddLine(pointXUc, pointYUc, pointXUb, pointYUb).Select
    
    With Selection.ShapeRange
        .Line.ForeColor.RGB     = RGB(51, 204, 51)
        .Line.EndArrowheadStyle = msoArrowheadStealth
        .Line.Weight            = 1.8
        .name                   = "Ubc"
    End With
    
    Range(selAdress).Select
End Sub


' Phase Uca

Sub drawVectorUca()
    Dim selAdress As String
    
    selAdress = Selection.Address
    
    Shapes.AddLine(pointXUa, pointYUa, pointXUc, pointYUc).Select
    
    With Selection.ShapeRange
        .Line.ForeColor.RGB     = RGB(255, 0, 0)
        .Line.EndArrowheadStyle = msoArrowheadStealth
        .Line.Weight            = 1.8
        .name                   = "Uca"
    End With
    
    Range(selAdress).Select
End Sub


' Redraw vectors U when something changes Update

Sub updateVectorU()
    Dim valUa As Single, valUb As Single, valUc As Single
    
    valUa = Range("voltageUa").Value
    valUb = Range("voltageUc").Value
    valUc = Range("voltageUc").Value
    
    ' delete drawn current
    On Error Resume Next
    ActiveSheet.Shapes("Ua").Delete
    On Error Resume Next
    ActiveSheet.Shapes("Ub").Delete
    On Error Resume Next
    ActiveSheet.Shapes("Uc").Delete
    On Error Resume Next
    ActiveSheet.Shapes("Uab").Delete
    On Error Resume Next
    ActiveSheet.Shapes("Ubc").Delete
    On Error Resume Next
    ActiveSheet.Shapes("Uca").Delete
    
    ' draw voltage
    
    If Range("isTH").Value = "0" Then
        Call drawVectorUa(kUa)

        If valUa <> 0 Then
            pointXUa = x - xLine
            pointYUa = y - yLine
        End If

        Call drawVectorUb(kUb)

        If valUb <> 0 Then
            pointXUb = x - xLine
            pointYUb = y - yLine
        End If

        Call drawVectorUc(kUc)

        If valUc <> 0 Then
            pointXUc = x - xLine
            pointYUc = y - yLine
        End If
    Else:
        Call drawVectorUab
        Call drawVectorUbc
        Call drawVectorUca
    End If
End Sub


' Find the largest current to fit into the scale automatically

Sub getkU()
    Dim Ua As Single, Ub As Single, Uc  As Single, digit As Byte
    
    Ua = Range("voltageUa").Value
    Ub = Range("voltageUb").Value
    Uc = Range("voltageUc").Value
    
    ' hide zeros after the decimal point depending on the number
    If InStr(Ua, ",") Then
        digit = Len(CStr(Ua)) - InStr(CStr(Ua), ",")

        If digit = 1 Then
            Range("voltageUa").NumberFormat = "0.0"
        Else:
            Range("voltageUa").NumberFormat = "0.00"
        End If

    Else:
        Range("voltageUa").NumberFormat = "0"
    End If
    
    If InStr(Ub, ",") Then
        digit = Len(CStr(Ub)) - InStr(CStr(Ub), ",")

        If digit = 1 Then
            Range("voltageUb").NumberFormat = "0.0"
        Else:
            Range("voltageUb").NumberFormat = "0.00"
        End If

    Else:
        Range("voltageUb").NumberFormat = "0"
    End If
    
    If InStr(Uc, ",") Then
        digit = Len(CStr(Uc)) - InStr(CStr(Uc), ",")

        If digit = 1 Then
            Range("voltageUc").NumberFormat = "0.0"
        Else:
            Range("voltageUc").NumberFormat = "0.00"
        End If

    Else:
        Range("voltageUc").NumberFormat = "0"
    End If
    
    If Ub >= Ub And Ua >= Uc Then
        kUa = 1
        On Error Resume Next
        kUb = Ub / Ua
        On Error Resume Next
        kUc = Uc / Ua
    End If
    
    If Ub >= Ua And Ub >= Uc Then
        On Error Resume Next
        kUa = Ua / Ub
        kUb = 1
        On Error Resume Next
        kUc = Uc / Ub
    End If
    
    If Uc >= Ua And Uc >= Ub Then
        On Error Resume Next
        kUa = Ua / Uc
        On Error Resume Next
        kUb = Ub / Uc
        kUc = 1
        If (Ua = 0 And Ub = 0 And Uc = 0) Then
            kUc = 0
        End If
    End If
End Sub


' Add lable voltage

Sub addLableU(name  As String)
    Dim valUa As Single, degUa As Single, isUaLC As String
    Dim valUb As Single, degUb As Single, isUbLC As String
    Dim valUc As Single, degUc As Single, isUcLC As String
    
    valUa  = Range("voltageUa").Value
    degUa  = Range("degUa").Value
    isUaLC = Range("UaLC").Value
    valUb  = Range("voltageUb").Value
    degUb  = Range("degUb").Value
    isUbLC = Range("UbLC").Value
    valUc  = Range("voltageUc").Value
    degUc  = Range("degUc").Value
    isUcLC = Range("UcLC").Value
    
    ' Lable voltage A
    
    If name = "Ua" And valUa <> 0 Then
        
        ActiveSheet.Shapes("lableUa").Visible = TRUE
        
        If degUa <= 45 Then
            ActiveSheet.Shapes("lableUa").Left = x - xLine - (wUa / 2)
            ActiveSheet.Shapes("lableUa").Top  = y - yLine - (hUa * 1.3)
        ElseIf degUa > 45 And degUa <= 135 And isUaLC = "L" Then
            ActiveSheet.Shapes("lableUa").Left = x - xLine
            ActiveSheet.Shapes("lableUa").Top = y - yLine - (hUa / 2)
        ElseIf degUa > 45 And degUa <= 135 And isUaLC = "C" Then
            ActiveSheet.Shapes("lableUa").Left = x - xLine - wUa
            ActiveSheet.Shapes("lableUa").Top = y - yLine - (hUa / 2)
        ElseIf degUa > 135 And degUa <= 180 Then
            ActiveSheet.Shapes("lableUa").Left = x - xLine - (wUa / 2)
            ActiveSheet.Shapes("lableUa").Top = y - yLine
        End If

    ElseIf valUa = 0 Then
        ActiveSheet.Shapes("lableUa").Visible = FALSE
    End If
    
    ' Lable voltage B
    
    If name = "Ub" And valUb <> 0 Then
        
        ActiveSheet.Shapes("lableUb").Visible = TRUE
        
        If degUb <= 45 Then
            ActiveSheet.Shapes("lableUb").Left = x - xLine - (wIa / 2)
            ActiveSheet.Shapes("lableUb").Top  = y - yLine - (hUa * 1.3)
        ElseIf degUb > 45 And degUb <= 135 And isUbLC = "L" Then
            ActiveSheet.Shapes("lableUb").Left = x - xLine
            ActiveSheet.Shapes("lableUb").Top  = y - yLine - (hUa / 2)
        ElseIf degUb > 45 And degUb <= 135 And isUbLC = "C" Then
            ActiveSheet.Shapes("lableUb").Left = x - xLine - wUa
            ActiveSheet.Shapes("lableUb").Top  = y - yLine - (hUa / 2)
        ElseIf degUb > 135 And degUb <= 180 Then
            ActiveSheet.Shapes("lableUb").Left = x - xLine - (wUa / 2)
            ActiveSheet.Shapes("lableUb").Top  = y - yLine
        End If
        
    ElseIf valUb = 0 Then
        ActiveSheet.Shapes("lableUb").Visible = FALSE
    End If
    
    ' Lable voltage C
    
    If name = "Uc" And valUc <> 0 Then
        
        ActiveSheet.Shapes("lableUc").Visible = TRUE
        
        If degUc <= 45 Then
            ActiveSheet.Shapes("lableUc").Left = x - xLine - (wUa / 2)
            ActiveSheet.Shapes("lableUc").Top  = y - yLine - (hUa * 1.3)
        ElseIf degUc > 45 And degUc < 130 And isUcLC = "L" Then
            ActiveSheet.Shapes("lableUc").Left = x - xLine
            ActiveSheet.Shapes("lableUc").Top  = y - yLine - (hUa / 2)
        ElseIf degUc > 45 And degUc < 130 And isUcLC = "C" Then
            ActiveSheet.Shapes("lableUc").Left = x - xLine - wUa
            ActiveSheet.Shapes("lableUc").Top  = y - yLine - (hUa / 2)
        ElseIf degUc >= 130 And degUc <= 180 Then
            ActiveSheet.Shapes("lableUc").Left = x - xLine - (wUa / 2)
            ActiveSheet.Shapes("lableUc").Top  = y - yLine
        End If
        
    ElseIf valUc = 0 Then
        ActiveSheet.Shapes("lableUc").Visible = FALSE
    End If
End Sub


' Reset value voltage

Sub resetU()
    Range("voltageUa").Value = 0
    Range("degUa").Value     = 0

    pointXUa = 0
    pointYUa = 0

    Range("voltageUb").Value = 0
    Range("degUb").Value     = 120

    pointXUb = 0
    pointYUb = 0

    Range("voltageUc").Value = 0
    Range("degUc").Value     = 120

    pointXUc = 0
    pointYUc = 0
    
    If Range("isIDegLC").Value = "LC" Then
        Range("UaLC").Value = "L"
        Range("UbLC").Value = "L"
        Range("UcLC").Value = "C"
    ElseIf Range("isIDegLC").Value = "Deg" Then
        Range("UaLC").Value = "°"
        Range("UbLC").Value = "°"
        Range("UcLC").Value = "°"
    End If

    Call updateVectorU

    ActiveSheet.Shapes("lableUa").Visible = FALSE
    ActiveSheet.Shapes("lableUb").Visible = FALSE
    ActiveSheet.Shapes("lableUc").Visible = FALSE

    Call animFigure
End Sub


' Draw Ulin on ABC

Sub drawUlin()
    
    Dim valUa     As Single, degUa As Single, isUaLC As String
    Dim valUb     As Single, degUb As Single, isUbLC As String
    Dim valUc     As Single, degUc As Single, isUcLC As String
    Dim selAdress As String
    Dim xUlinAB   As Single, yUlinAB As Single
    Dim delX      As Single, delY As Single
    
    valUa  = Range("voltageUa").Value
    degUa  = Range("degUa").Value
    isUaLC = Range("UaLC").Value
    valUb  = Range("voltageUb").Value
    degUb  = Range("degUb").Value
    isUbLC = Range("UbLC").Value
    valUc  = Range("voltageUc").Value
    degUc  = Range("degUc").Value
    isUcLC = Range("UcLC").Value
    
    ' Voltage Uab
    
    selAdress = Selection.Address
    
    If isUaLC = "C" Then
        degUa = -degUa
    End If
    
    delX = LUlin / 2 * (1 - Sin((30 - degUa) * Pi / 180))
    delY = radius - (LUlin / 2 * Sin((30 - degUa) * Pi / 180))
    
    Shapes.AddLine(x + LUlin, y, x, y).Select
    
    xUlinAB = Selection.ShapeRange.Top
    yUlinAB = Selection.ShapeRange.Left
    
    With Selection.ShapeRange
        .Line.ForeColor.RGB     = RGB(255, 255, 0)
        .Line.EndArrowheadStyle = msoArrowheadStealth
        .Line.Weight            = 1.8
        .name                   = "Uab"
        .IncrementRotation (60 + degUa)
        .IncrementLeft -delX
        .IncrementTop -delY
    End With

    Range(selAdress).Select
    
    ' Voltage Ubc
    
    selAdress = Selection.Address
    
    If isUbLC = "L" Then
        degUb = -degUb
    End If
    
    delX = LUlin / 2 * (1 - Cos((60 - degUb) * Pi / 180))
    delY = radius / 2 * (1 - Sin((30 - degUb) * Pi / 180))
    
    Shapes.AddLine(x + LUlin, y, x, y).Select
    
    With Selection.ShapeRange
        .Line.ForeColor.RGB     = RGB(0, 255, 0)
        .Line.EndArrowheadStyle = msoArrowheadStealth
        .Line.Weight            = 1.8
        .name                   = "Ubc"
        .IncrementRotation (60 - degUb)
        .IncrementLeft -delX
        .IncrementTop -delY
    End With

    Range(selAdress).Select
    
    ' Voltage Uca
    
    selAdress = Selection.Address
    
    If isUcLC = "L" Then
        degUc = -degUc
    End If
    
    delX = LUlin / 2 * (1 - Cos((60 - degUc) * Pi / 180))
    delY = radius / 2 * (1 - Sin((30 - degUc) * Pi / 180))
    
    Shapes.AddLine(x + LUlin, y, x, y).Select
    
    With Selection.ShapeRange
        .Line.ForeColor.RGB     = RGB(255, 0, 0)
        .Line.EndArrowheadStyle = msoArrowheadStealth
        .Line.Weight            = 1.8
        .name                   = "Uca"
        .IncrementRotation (60 - degUc)
        .IncrementLeft -delX
        .IncrementTop -delY
    End With

    Range(selAdress).Select
End Sub


' Initializing the sheet on load

Private Sub Worksheet_Activate()
    hIa = ActiveSheet.Shapes("lableIa").Height
    wIa = ActiveSheet.Shapes("lableIa").Width
    hUa = ActiveSheet.Shapes("lableUa").Height
    wUa = ActiveSheet.Shapes("lableUa").Width
    
    x = XCentreDiagram()
    y = YCentreDiagram()

    radius = RadiusBackground()
    LUlin  = 1.73205 * radius
    MI     = 0.8        ' current scale
End Sub


' X center background

Public Function XCentreDiagram() As Single
    XCentreDiagram = Shapes("CircleFon").Left + Shapes("CircleFon").Width / 2
End Function


' Y center background

Public Function YCentreDiagram() As Single
    YCentreDiagram = Shapes("CircleFon").Top + Shapes("CircleFon").Height / 2
End Function


' Diameter BackgroundDiagram

Public Function RadiusBackground() As Single
    RadiusBackground = Shapes("Background").Width / 2
End Function


' Animate a shape when pressed

Sub animFigure()
    
    Dim pushName As String
    Dim Start
    
    pushName = Application.Caller
    
    With ActiveSheet.Shapes(pushName)
        .IncrementTop 1.5
        Start = Timer
        Do While Timer < Start + 0.05  ' pause sec
            DoEvents
        Loop
        .IncrementTop -1.5
    End With
End Sub

Sub copyFon()
    ActiveSheet.Range("M2:U23").CopyPicture Appearance:=xlScreen, Format:=xlBitmap
    Call animFigure
End Sub
