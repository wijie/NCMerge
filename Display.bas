Attribute VB_Name = "Display"
Option Explicit

Dim mintF1 As Integer

'*********************************************************
' �p  �r: �s�N�`���{�b�N�X�Ɍ��̊G���T�[�N����`��
' ��  ��: objDisp: �s�N�`���{�b�N�X�̖��O
'         objBar: �v���O���X�o�[�̖��O
' �߂�l: ����
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
    Dim lngSize As Long ' �t�@�C���T�C�Y
    Dim dblMin() As Double
    Dim dblMax() As Double
    Dim dblPixelPerMM As Double
    Dim lngWBOrigin(1) As Long ' WB�����̍��W
    Dim dblNTOriginY As Double ' NT��Y�����_�̍��W
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
        .ScaleMode = 6 ' ScaleMode ���~���ɐݒ肵�܂�
        .DrawWidth = 1 ' ������1Pixel

        ' NC�f�[�^�̍��G���A
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

        ' �k�ڂ̌���
        dblFactor(X) = Round((dblNCWidth + 20) / Abs(.ScaleWidth), intDigit)
        dblFactor(Y) = Round((dblNCHeight + 20) / Abs(.ScaleHeight), intDigit)
        If dblFactor(X) > dblFactor(Y) Then
            frmMain.ScaleFactor = dblFactor(X)
        Else
            frmMain.ScaleFactor = dblFactor(Y)
        End If

        ' ���W�n�̐ݒ�
        .ScaleHeight = Abs(.ScaleHeight) * frmMain.ScaleFactor * -1 '��������Y+�����ɂ���
        .ScaleWidth = .ScaleWidth * frmMain.ScaleFactor

        ' �]���̐ݒ�
        dblSpace(X) = Round((Abs(.ScaleWidth) - dblNCWidth) / 2, intDigit)
        dblSpace(Y) = Round((Abs(.ScaleHeight) - dblNCHeight) / 2, intDigit)

        ' �\���ʒu�̐ݒ�
        .ScaleLeft = Round(dblMin(X) - dblSpace(X), intDigit)
        .ScaleTop = Round(dblMax(Y) + dblSpace(Y), intDigit)

        ' �����\���p�s�N�`���[�{�b�N�X�̐ݒ�
        With objDisp(1)
            .Width = (dblNCWidth + 20) * 56.7
            .Height = (dblNCHeight + 20) * 56.7
            .ScaleMode = 6
            .ScaleHeight = .ScaleHeight * -1
            .ScaleLeft = Round(dblMin(X) - 10, intDigit)
            .ScaleTop = Round(dblMax(Y) + 10, intDigit)
        End With

        ' 1�s�N�Z���̑傫��(mm�P��)
        dblPixelPerMM = Round(Screen.TwipsPerPixelX / 56.7, intDigit)

Display: ' ��ʂɏo��

        intF0 = FreeFile
        Open fTempPath & "NCView._$$" For Input As #intF0
        lngSize = LOF(intF0)
        Do While Not EOF(intF0)
            Input #intF0, lngX, lngY, dblR, lngColor, intTool
            dblX = lngX / int1mm
            dblY = lngY / int1mm
            ' �S�̕\��
            If dblR / frmMain.ScaleFactor < dblPixelPerMM / 2 Then
                ' ��ʏ��1�s�N�Z���ȉ���1�s�N�Z���̓_�ŕ`��
                objDisp(0).PSet (dblX, dblY), lngColor
            Else
                ' ���̑��̓T�[�N���ŕ`��
                objDisp(0).Circle (dblX, dblY), dblR, lngColor
            End If

            ' �����\��
            If (dblR * 2) - dblPixelPerMM < dblPixelPerMM Then
                ' ���C���̃Z���^�`�Z���^��1�s�N�Z���ȉ���1�s�N�Z���̓_�ŕ`��
                objDisp(1).PSet (dblX, dblY), lngColor
            Else
                ' ���C���̊O�������a�ƈ�v����l�ɕ`��
                objDisp(1).Circle (dblX, dblY), dblR - (dblPixelPerMM / 2), lngColor
            End If

            ' �v���O���X�o�[�̍X�V
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

        ' �S�̕\���p���_�}�[�N
        objDisp(0).Line (-2.5, -2.5)-Step(5, 5), RGB(0, 0, 0)
        objDisp(0).Line (-2.5, 2.5)-Step(5, -5), RGB(0, 0, 0)

        ' �����\���p���_�}�[�N
        objDisp(1).Line (-2.5, -2.5)-Step(5, 5), RGB(0, 0, 0)
        objDisp(1).Line (-2.5, 2.5)-Step(5, -5), RGB(0, 0, 0)

        Close #intF0
    End With

End Sub
