VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "���t�I��"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2175
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private MonthView1 As clsMonthViewOnLabel

Private Sub CommandButton1_Click()

Dim DateA As Date

    DateA = MonthView1.Value
    '�N���X�C���X�^���X�p��
    MonthView1.Destroy
    Set MonthView1 = Nothing
    Cal_Close (DateA)
    
End Sub

Private Sub UserForm_Activate()
    Dim DateA As Date
    
    Set MonthView1 = New clsMonthViewOnLabel
    DateA = Now()
    With MonthView1
      .Cmd = Label1
      .UserForm = Me
      .Create    ' MonthView�𐶐�
      '�����\����{�����t�ȊO�ɂ���ꍇ�ͤCreate���
      '���̈ʒu��[Value]�v���p�e�B�ɑ������
      ' .Value = DateValue("2003/8/10")
      .Value = Strings.Format(DateA, "yyyy/mm/dd")
    End With
    
End Sub

