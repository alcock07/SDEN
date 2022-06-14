Attribute VB_Name = "Module1"
Option Explicit

#If VBA7 Then
    Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#Else
    Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#End If

Public Const MAX_COMPUTERNAME_LENGTH = 15

'�R���s���[�^���擾�֐�
Public Function CP_NAME() As String

    Const COMPUTERNAMBUFFER_LENGTH = MAX_COMPUTERNAME_LENGTH + 1
    Dim strComputerNameBuffer As String * COMPUTERNAMBUFFER_LENGTH
    Dim lngComputerNameLength As Long
    Dim lngWin32apiResultCode As Long
    
    ' �R���s���[�^�[���̒�����ݒ�
    lngComputerNameLength = Len(strComputerNameBuffer)
    ' �R���s���[�^�[�����擾
    lngWin32apiResultCode = GetComputerName(strComputerNameBuffer, lngComputerNameLength)
    ' �R���s���[�^�[����\��
    CP_NAME = Strings.Left(strComputerNameBuffer, InStr(strComputerNameBuffer, vbNullChar) - 1)
    
End Function

Sub Prn_Sht()

    Dim lngR   As Long
    Dim strKNM As String
    
    lngR = 8
    If Range("AB1") = "OS" Or Range("AB1") = "HB" Then
        strKNM = "����"
    Else
        strKNM = "��������"
    End If
    
    Do
        '�K�p���Ƌ��z���u�����N�̎��͍s�̍�����0�ɂ���
        If Cells(lngR, 7) = "" And Cells(lngR + 1, 7) = "" And Cells(lngR, 13) = "" Then
                Rows(lngR & ":" & lngR + 1).RowHeight = 0
        ElseIf InStr(1, Cells(lngR, 7), "����") <> 0 Then
            Cells(lngR, 2) = "735"
            Cells(lngR + 1, 2) = "���q�^����"
            Cells(lngR + 1, 16) = strKNM
        ElseIf InStr(1, Cells(lngR, 7), "����") <> 0 Then
            Cells(lngR, 2) = "731"
            Cells(lngR + 1, 2) = "�ב��^����"
            Cells(lngR + 1, 16) = strKNM
        ElseIf InStr(1, Cells(lngR, 7), "��") <> 0 Then
            Cells(lngR, 2) = "738"
            Cells(lngR + 1, 2) = "�d�Ō���"
            Cells(lngR + 1, 16) = strKNM
        ElseIf InStr(1, Cells(lngR, 7), "���N�f�f") <> 0 Then
            Cells(lngR, 2) = "724"
            Cells(lngR + 1, 2) = "����������"
            Cells(lngR + 1, 16) = strKNM
        ElseIf InStr(1, Cells(lngR, 7), "���H��") <> 0 Then
            Cells(lngR, 2) = "745"
            Cells(lngR + 1, 2) = "�q�ɏ��Ք�"
            Cells(lngR + 1, 16) = strKNM
        ElseIf InStr(1, Cells(lngR, 7), "�X��") <> 0 Or InStr(1, Cells(lngR, 7), "�؎�") <> 0 Or InStr(1, Cells(lngR, 7), "���^�[�p�b�N") <> 0 Or InStr(1, Cells(lngR, 7), "�䂤�p�b�N") <> 0 Then
            Cells(lngR, 2) = "727"
            Cells(lngR + 1, 2) = "�ʐM��"
            Cells(lngR + 1, 16) = strKNM
        End If
        lngR = lngR + 2
        If lngR > 26 Then Exit Do
    Loop
    If lngR > 8 Then
        ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
        ActiveSheet.PageSetup.PrintGridlines = False
    End If
End Sub

Sub Add_Row()
    Dim lngR As Long
    Dim lngC As Long
    lngR = 18
    Do
        If Rows(lngR & ":" & lngR + 1).RowHeight = 0 Then
            Rows(lngR & ":" & lngR + 1).RowHeight = 16.5
            Exit Do
        End If
        lngR = lngR + 2
        If lngR > 26 Then Exit Do
    Loop
End Sub

Sub Cls_Sht()
    Range("B8:O27").Select
    Selection.ClearContents
    Range("P8:R27").Select
    Selection.ClearContents
    Range("G8:O27").Select
    Selection.ClearContents
    Range("B5") = Now()
    Range("L5:T5").Select
    Selection.ClearContents
    Range("B2").Select
    Range("B2") = "�o��"
    Rows("8:17").RowHeight = 16.5
    Rows("18:27").RowHeight = 0
End Sub

Sub Prn_Tmp()
    Call Cls_Sht
    Range("B5") = "  �N      ��      ��"
    Range("M28") = ""
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
    ActiveSheet.PageSetup.PrintGridlines = False
    Range("M28").FormulaR1C1 = "=SUM(R[-20]C:R[-1]C[2])"
End Sub

Sub Prn_Check()
    Dim lngCHK As Long
    lngCHK = Range("M28")
    If lngCHK = 0 Then
        Call Prn_Tmp
    Else
        Call Prn_Sht
    End If
End Sub

Sub Cal_Open()
'����ް�\��----------------------------
Dim DateA As Date

    UserForm1.Show
    
End Sub

Sub Cal_Close(DateA As Date)
'����ް����Ƃ����t�擾---------------
    
    UserForm1.Hide
    Sheets("�o��").Select
    Range("B5") = DateA
    
End Sub

Sub AP_END()
'==================
' �I�������@Ver2.0
'==================

    Dim myBook As Workbook
    Dim strFN As String
    Dim boolB As Boolean
    
    Application.ReferenceStyle = xlA1
    Application.MoveAfterReturnDirection = xlDown
    Application.DisplayAlerts = False
    
    strFN = ThisWorkbook.Name '���̃u�b�N�̖��O
    boolB = False
    For Each myBook In Workbooks
        If myBook.Name <> strFN Then boolB = True
    Next
    If boolB Then
        ThisWorkbook.Close False  '�t�@�C�������
    Else
        Application.Quit  'Excell���I��
        ThisWorkbook.Saved = True
        ThisWorkbook.Close False
    End If
    
End Sub

