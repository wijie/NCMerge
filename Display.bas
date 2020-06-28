Attribute VB_Name = "Display"
Option Explicit

Dim mintF1 As Integer

'*********************************************************
' 用  途: ピクチャボックスに穴の絵をサークルを描く
' 引  数: objDisp: ピクチャボックスの名前
'         objBar: プログレスバーの名前
' 戻り値: 無し
'*********************************************************

Sub sDispNC(objDisp As Object, _
            udtNCInfo As NCInfo, _
            intFileNo As Integer, _
            objBar As Object)

    Dim lngX As Long
    Dim lngY As Long
    Dim dblX As Double
    Dim dblY As Double
    Dim dblR As Double
    Dim lngColor As Long
    Dim intTool As Integer
    Dim intDigit As Integer
    Dim dblSpace(1) As Double
    Dim dblNCWidth As Double
    Dim dblNCHeight As Double
    Dim dblFactor(1) As Double
    Dim intF0 As Integer
    Dim lngSize As Long ' ファイルサイズ
    Dim dblMin() As Double
    Dim dblMax() As Double
    Dim dblPixelPerMM As Double
    Dim lngWBOrigin(1) As Long ' WB左下の座標
    Dim dblNTOriginY As Double ' NTのY側原点の座標
    Dim lngABS(1) As Long
    Dim intTNo As Integer
    Dim intCurrentTool As Integer

    intDigit = 7
    With objBar
        .Max = 100
        .Min = 0
        .Value = objBar.Min
    End With

    mintF1 = intFileNo
    lngABS(X) = 0
    lngABS(Y) = 0
    intTNo = 1
    intCurrentTool = -32767

    With objDisp(0)
        .ScaleMode = 6 ' ScaleMode をミリに設定します
        .DrawWidth = 1 ' 線幅は1Pixel

        ' NCデータの作画エリア
        With udtNCInfo
            dblMin = .dblMin
            dblMax = .dblMax
        End With
        If dblMin(X) > 0 Then dblMin(X) = 0
        If dblMax(X) < 0 Then dblMax(X) = 0
        If dblMin(Y) > 0 Then dblMin(Y) = 0
        If dblMax(Y) < 0 Then dblMax(Y) = 0
        dblNCWidth = Abs(dblMax(X) - dblMin(X))
        dblNCHeight = Abs(dblMax(Y) - dblMin(Y))

        ' 縮尺の決定
        dblFactor(X) = Round((dblNCWidth + 20) / Abs(.ScaleWidth), intDigit)
        dblFactor(Y) = Round((dblNCHeight + 20) / Abs(.ScaleHeight), intDigit)
        If dblFactor(X) > dblFactor(Y) Then
            frmMain.ScaleFactor = dblFactor(X)
        Else
            frmMain.ScaleFactor = dblFactor(Y)
        End If

        ' 座標系の設定
        .ScaleHeight = Abs(.ScaleHeight) * frmMain.ScaleFactor * -1 '下から上をY+方向にする
        .ScaleWidth = .ScaleWidth * frmMain.ScaleFactor

        ' 余白の設定
        dblSpace(X) = Round((Abs(.ScaleWidth) - dblNCWidth) / 2, intDigit)
        dblSpace(Y) = Round((Abs(.ScaleHeight) - dblNCHeight) / 2, intDigit)

        ' 表示位置の設定
        .ScaleLeft = Round(dblMin(X) - dblSpace(X), intDigit)
        .ScaleTop = Round(dblMax(Y) + dblSpace(Y), intDigit)

        ' 正寸表示用ピクチャーボックスの設定
        With objDisp(1)
            .Width = (dblNCWidth + 20) * 56.7
            .Height = (dblNCHeight + 20) * 56.7
            .ScaleMode = 6
            .ScaleHeight = .ScaleHeight * -1
            .ScaleLeft = Round(dblMin(X) - 10, intDigit)
            .ScaleTop = Round(dblMax(Y) + 10, intDigit)
        End With

        ' 1ピクセルの大きさ(mm単位)
        dblPixelPerMM = Round(Screen.TwipsPerPixelX / 56.7, intDigit)

Display: ' 画面に出力

        intF0 = FreeFile
        Open fTempPath & "NCView._$$" For Input As #intF0
        lngSize = LOF(intF0)
        Do While Not EOF(intF0)
            Input #intF0, lngX, lngY, dblR, lngColor, intTool
            dblX = lngX / int1mm
            dblY = lngY / int1mm
            ' 全体表示
            If dblR / frmMain.ScaleFactor < dblPixelPerMM / 2 Then
                ' 画面上で1ピクセル以下は1ピクセルの点で描く
                objDisp(0).PSet (dblX, dblY), lngColor
            Else
                ' その他はサークルで描く
                objDisp(0).Circle (dblX, dblY), dblR, lngColor
            End If

            ' 正寸表示
            If (dblR * 2) - dblPixelPerMM < dblPixelPerMM Then
                ' ラインのセンタ～センタで1ピクセル以下は1ピクセルの点で描く
                objDisp(1).PSet (dblX, dblY), lngColor
            Else
                ' ラインの外側が穴径と一致する様に描く
                objDisp(1).Circle (dblX, dblY), dblR - (dblPixelPerMM / 2), lngColor
            End If

            ' プログレスバーの更新
'            objBar.Value = Int(Seek(intF0) / lngSize * 100)

            If intCurrentTool <> intTool Then
                If intCurrentTool <> -32767 Then Print #mintF1, "G80"
                Print #mintF1, "T"; CStr(intTNo)
                Print #mintF1, "G81"
                intCurrentTool = intTool
                intTNo = intTNo + 1
            End If
            Print #mintF1, "X"; CStr(lngX - lngABS(X)); "Y"; CStr(lngY - lngABS(Y))
            lngABS(X) = lngX
            lngABS(Y) = lngY
        Loop

        Print #mintF1, "G80"
        Print #mintF1, "X"; CStr(lngABS(X) * -1); "Y"; CStr(lngABS(Y) * -1)
        Print #mintF1, "M02"

        ' 全体表示用原点マーク
        objDisp(0).Line (-2.5, -2.5)-Step(5, 5), RGB(0, 0, 0)
        objDisp(0).Line (-2.5, 2.5)-Step(5, -5), RGB(0, 0, 0)

        ' 正寸表示用原点マーク
        objDisp(1).Line (-2.5, -2.5)-Step(5, 5), RGB(0, 0, 0)
        objDisp(1).Line (-2.5, 2.5)-Step(5, -5), RGB(0, 0, 0)

        Close #intF0
    End With

End Sub
