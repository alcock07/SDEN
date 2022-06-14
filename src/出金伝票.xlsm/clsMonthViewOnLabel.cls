VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMonthViewOnLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'========================================================
'   Label上にMonthViewを表示させるクラス
'========================================================
'   このクラスは、[VersionUp]のAKIさんが公開している
'   MonthViewをAPIで実装するVB向けクラスを、角田(AddinBox)
'   がVBA/UserForm向けに編集/移植したものです。
'
'   2004/5/4  初版公開(Ver1.0)
'   2005/3/10 Ver1.1
'         クラスやコレクション等の宣言時に付けていた[New]を外し、
'         Initializeにて[Set ～ = New ～]とする
'         [New] を付けているとTerminateでNothing代入しても再割当されてしまう為
'--------------------------------------------------------
'   AddinBox  ( http://www.h3.dion.ne.jp/~sakatsu/ )
'   VersionUp ( http://www.mtecb.co.jp/versionup/  )
'========================================================

Private Const MONTHCAL_CLASS = "SysMonthCal32"
Private Const ICC_DATE_CLASSES = &H100
Private Const MCN_FIRST = -750
Private Const MCN_LAST = -759
Private Const MCM_FIRST = &H1000

Private Const MCHT_TITLE = &H10000
Private Const MCHT_CALENDAR = &H20000
Private Const MCHT_TODAYLINK = &H30000

Private Const MCHT_NEXT = &H1000000
Private Const MCHT_PREV = &H2000000

Private Const MCHT_NOWHERE = &H0

Private Const MCHT_TITLEBK = (MCHT_TITLE)
Private Const MCHT_TITLEMONTH = (MCHT_TITLE Or &H1)
Private Const MCHT_TITLEYEAR = (MCHT_TITLE Or &H2)
Private Const MCHT_TITLEBTNNEXT = (MCHT_TITLE Or MCHT_NEXT Or &H3)
Private Const MCHT_TITLEBTNPREV = (MCHT_TITLE Or MCHT_PREV Or &H3)

Private Const MCHT_CALENDARBK = (MCHT_CALENDAR)
Private Const MCHT_CALENDARDATE = (MCHT_CALENDAR Or &H1)
Private Const MCHT_CALENDARDATENEXT = (MCHT_CALENDARDATE Or MCHT_NEXT)
Private Const MCHT_CALENDARDATEPREV = (MCHT_CALENDARDATE Or MCHT_PREV)
Private Const MCHT_CALENDARDAY = (MCHT_CALENDAR Or &H2)
Private Const MCHT_CALENDARWEEKNUM = (MCHT_CALENDAR Or &H3)

Private Const MCM_GETCURSEL = (MCM_FIRST + 1)
Private Const MCM_SETCURSEL = (MCM_FIRST + 2)
Private Const MCM_GETMAXSELCOUNT = (MCM_FIRST + 3)
Private Const MCM_SETMAXSELCOUNT = (MCM_FIRST + 4)
Private Const MCM_GETSELRANGE = (MCM_FIRST + 5)
Private Const MCM_SETSELRANGE = (MCM_FIRST + 6)
Private Const MCM_GETMINREQRECT = (MCM_FIRST + 9)
Private Const MCM_SETCOLOR = (MCM_FIRST + 10)
Private Const MCM_GETCOLOR = (MCM_FIRST + 11)
Private Const MCM_SETTODAY = (MCM_FIRST + 12)
Private Const MCM_SETFIRSTDAYOFWEEK = (MCM_FIRST + 15)
Private Const MCM_GETFIRSTDAYOFWEEK = (MCM_FIRST + 16)
Private Const MCM_GETRANGE = (MCM_FIRST + 17)
Private Const MCM_SETRANGE = (MCM_FIRST + 18)

Private Const MCSC_BACKGROUND = 0    '// the background color (between months)
Private Const MCSC_TEXT = 1          '// the dates
Private Const MCSC_TITLEBK = 2       '// background of the title
Private Const MCSC_TITLETEXT = 3
Private Const MCSC_MONTHBK = 4       '// background within the month cal
Private Const MCSC_TRAILINGTEXT = 5  '// the text color of header & trailing days

Private Const MCN_SELCHANGE = (MCN_FIRST + 1)
Private Const MCN_GETDAYSTATE = (MCN_FIRST + 3)
Private Const MCN_SELECT = (MCN_FIRST + 4)

Private Const MCS_DAYSTATE = &H1
Private Const MCS_MULTISELECT = &H2
Private Const MCS_WEEKNUMBERS = &H4
Private Const MCS_NOTODAY = &H8

Private Const GDTR_MIN = &H1
Private Const GDTR_MAX = &H2

Private Const GMR_VISIBLE = 0        '// visible portion of display
Private Const GMR_DAYSTATE = 1       '// above plus the grayed out parts of
                                    '// partially displayed months

Private Const LOGPIXELSX = 88               ' ポイント→ピクセル変換指定(横)
Private Const LOGPIXELSY = 90               ' ポイント→ピクセル変換指定(縦)
Private Const WS_VISIBLE = &H10000000
Private Const WS_CHILD = &H40000000
Private Const WS_BORDER = &H800000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_EX_TOOLWINDOW = &H80&
Private Const WS_EX_CLIENTEDGE = &H200&     'クライアント領域が凹形の3D表示
Private Const GWL_HINSTANCE As Long = (-6)
Private Const GDT_VALID = 0

' コントロール初期化の指示子
Private Type tagINITCOMMONCONTROLSEX
    dwSize          As Long
    dwICC           As Long
End Type

' コントロールから引き渡される日付時刻構造体
Private Type SYSTEMTIME
    Year            As Integer
    Month           As Integer
    DayOfWeek       As Integer
    Day             As Integer
    Hour            As Integer
    Minute          As Integer
    Second          As Integer
    Milliseconds    As Integer
End Type

' 日付選択範囲
Private Type SYSTEMTIMERANGE
    MinYear         As Integer
    MinMonth        As Integer
    MinDayOfWeek    As Integer
    MinDay          As Integer
    MinHour         As Integer
    MinMinute       As Integer
    MinSecond       As Integer
    MinMilliseconds As Integer
    MaxYear         As Integer
    MaxMonth        As Integer
    MaxDayOfWeek    As Integer
    MaxDay          As Integer
    MaxHour         As Integer
    MaxMinute       As Integer
    MaxSecond       As Integer
    MaxMilliseconds As Integer
End Type

' ウィンドウの四隅の位置
Private Type RECT
    Left        As Long
    Top         As Long
    Right       As Long
    Bottom      As Long
End Type

' コモンコントロールの初期化API
Private Declare Function InitCommonControlsEx Lib "ComCtl32" _
    (LPINITCOMMONCONTROLSEX As Any) As Long
                                
' ウィンドウハンドル取得API
Private Declare Function FindWindow Lib "user32" _
    Alias "FindWindowA" _
    (ByVal lpClassName As String, _
     ByVal lpWindowName As String) As Long

' 拡張ウィンドウハンドル取得API
Private Declare Function FindWindowEx Lib "user32" _
    Alias "FindWindowExA" _
    (ByVal hwndParent As Long, _
     ByVal hwndChildAfter As Long, _
     ByVal lpszClass As String, _
     ByVal lpszWindow As String) As Long

' アプリケーションのインスタンスを取得
Private Declare Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongA" _
    (ByVal hWnd As Long, _
     ByVal nIndex As Long) As Long

' ポイント→ピクセル変換係数取得API
Private Declare Function GetDeviceCaps Lib "gdi32" _
    (ByVal hdc As Long, ByVal nIndex As Long) As Long
    
' DeskTopWindow取得
Private Declare Function GetDesktopWindow Lib "user32" () As Long

' デバイスコンテキスト取得
Private Declare Function GetDC Lib "user32" _
    (ByVal hWnd As Long) As Long

' デバイスコンテキスト解放
Private Declare Function ReleaseDC Lib "user32" _
    (ByVal hWnd As Long, ByVal hdc As Long) As Long

' ウィンドウの作成
Private Declare Function CreateWindowEx Lib "user32" _
    Alias "CreateWindowExA" _
    (ByVal dwExStyle As Long, ByVal lpClassName As String, _
     ByVal lpWindowName As String, ByVal dwStyle As Long, _
     ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
     ByVal nHeight As Long, ByVal hwndParent As Long, _
     ByVal hMenu As Long, ByVal hInstance As Long, _
     lpParam As Any) As Long

' ウィンドウの移動､拡縮
Private Declare Function MoveWindow Lib "user32" _
    (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, _
     ByVal nWidth As Long, ByVal nHeight As Long, _
     ByVal bRepaint As Long) As Long

' ウィンドウの撤去
Private Declare Function DestroyWindow Lib "user32" _
    (ByVal hWnd As Long) As Long

' ウィンドウコントロールの操作
Private Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, _
     ByVal wParam As Long, lParam As Any) As Long
                                
' ウィンドウが作成されているかの判定
Private Declare Function IsWindow Lib "user32" _
    (ByVal hWnd As Long) As Long

' ウィンドウのサイズ(四隅の位置)を取得
Private Declare Function GetWindowRect Lib "user32" _
    (ByVal hWnd As Long, lpRect As RECT) As Long

' システムカラー定数(vbButtonFace 等)をRGB値に変換
Private Declare Function GetSysColor Lib "user32" _
    (ByVal nIndex As Long) As Long

'---------- クラスモジュール内で使う変数の宣言(実行中固定) ----------------------
Private mctlForm  As UserForm           ' 呼出元UserForm
Private mctlLabel As MSForms.Label      ' 呼出元Label
Attribute mctlLabel.VB_VarHelpID = -1

Private mlnghwndMonthView As Long       ' MonthViewのウィンドウハンドル
Private lnghWnd_Excel As Long           ' Excelのウィンドウハンドル
Private lnghWnd_Form As Long            ' UserFormのウィンドウハンドル
Private lngPixelsX As Long              ' ポイント→ピクセル係数(横)
Private lngPixelsY As Long              ' ポイント→ピクセル係数(縦)
Private blnMV3D As Boolean              ' True:3D, False:Flat
Private blnMultiSelDays As Boolean      ' True:MultiSelect False:Not
Private intMVRows As Integer            ' 縦に並べる数
Private intMVColumns As Integer         ' 横に並べる数

'*******************************************************************************
'   プロパティの記述
'*******************************************************************************
' Cmd(基盤となるLabelオブジェクトの登録)
Public Property Get Cmd() As MSForms.Label
    Set Cmd = mctlLabel
End Property

Public Property Let Cmd(ctlNewLabel As MSForms.Label)
    Set mctlLabel = ctlNewLabel
End Property

'-------------------------------------------------------------------------------
' UserForm(UserFormオブジェクトの登録)
Public Property Get UserForm() As UserForm
    Set UserForm = mctlForm
End Property

Public Property Let UserForm(ctlNewUserForm As UserForm)
    Set mctlForm = ctlNewUserForm
End Property

'-------------------------------------------------------------------------------
' NecessaryHeight/NecessaryWidth  行数/列数/3D/Flatに応じた必要Height/Width
Public Property Get NecessaryHeight() As Double
    Dim lngResult As Long
    Dim wRECT As RECT
    Dim dblMinHeight As Double
    lngResult = SendMessage(mlnghwndMonthView, _
                            MCM_GETMINREQRECT, _
                            GDT_VALID, _
                            wRECT)
    If (CBool(lngResult) <> False) Then
        dblMinHeight = (wRECT.Bottom - wRECT.Top) * 72 / lngPixelsY
        '(7):[今日日付]部分の高さ
        NecessaryHeight = intMVRows * (dblMinHeight - 7) + 7
        If (blnMV3D = True) Then
            NecessaryHeight = NecessaryHeight + 4     '(4):3D表示縁
        End If
    Else
        NecessaryHeight = 0
    End If
End Property

Public Property Get NecessaryWidth() As Double
    Dim lngResult As Long
    Dim wRECT As RECT
    Dim dblMinWidth As Double
    lngResult = SendMessage(mlnghwndMonthView, _
                            MCM_GETMINREQRECT, _
                            GDT_VALID, _
                            wRECT)
    If (CBool(lngResult) <> False) Then
        dblMinWidth = (wRECT.Right - wRECT.Left) * 72 / lngPixelsX
        '(4.5):セパレータ領域の幅
        NecessaryWidth = intMVColumns * dblMinWidth + (intMVColumns - 1) * 4.5
        If (blnMV3D = True) Then
            NecessaryWidth = NecessaryWidth + 4   '(4):3D表示縁
        End If
    Else
        NecessaryWidth = 0
    End If
End Property

'-------------------------------------------------------------------------------
' Value(シリアル値の設定と取得)‥‥ MultiSelectモード以外
Public Property Get Value() As Date
    Dim lngResult As Long
    Dim st As SYSTEMTIME
    If (blnMultiSelDays = False) Then
        lngResult = SendMessage(mlnghwndMonthView, _
                                MCM_GETCURSEL, _
                                GDT_VALID, _
                                st)
        With st
            Value = DateSerial(.Year, .Month, .Day)
        End With
    Else
        'MultiSelect モードでは不可
    End If
End Property

Public Property Let Value(dtmNewValue As Date)
    Dim lngResult As Long
    Dim st As SYSTEMTIME
    If (blnMultiSelDays = False) Then
        ' SYSTEMTIME構造体に引数の日付時間を分解して代入。
        With st
            .Year = Year(dtmNewValue)
            .Month = Month(dtmNewValue)
            .Day = Day(dtmNewValue)
            .DayOfWeek = Weekday(dtmNewValue) - 1  ' APIでは日曜日が0
            .Hour = 0
            .Minute = 0
            .Second = 0
        End With
        lngResult = SendMessage(mlnghwndMonthView, _
                                MCM_SETCURSEL, _
                                GDT_VALID, _
                                st)
    Else
        'MultiSelect モードでは不可
    End If
End Property

'-------------------------------------------------------------------------------
' MinDate(選択可能日付範囲の下限)
Public Property Get MinDate() As Date
    Dim lngResult As Long
    Dim st As SYSTEMTIMERANGE
    
    lngResult = SendMessage(mlnghwndMonthView, _
                            MCM_GETRANGE, _
                            GDT_VALID, _
                            st)
    MinDate = DateSerial(st.MinYear, st.MinMonth, st.MinDay)
End Property

Public Property Let MinDate(ByVal dtmNewValue As Date)
    Dim lngResult As Long
    Dim dtmMax As Date
    Dim st As SYSTEMTIMERANGE
    
    ' 現在設定されている内容を取得する
    lngResult = SendMessage(mlnghwndMonthView, _
                            MCM_GETRANGE, _
                            GDT_VALID, _
                            st)
    
    ' 取得したSYSTEMTIMERANGEの[Min]を引数の日付に置換
    With st
        .MinYear = Year(dtmNewValue)
        .MinMonth = Month(dtmNewValue)
        .MinDay = Day(dtmNewValue)
        .MinDayOfWeek = Weekday(dtmNewValue) - 1     ' APIでは日曜日が0
        .MinHour = 0
        .MinMinute = 0
        .MinSecond = 0
    End With
    
    If (st.MaxYear = 0) Then
        ' 上限が設定されていないので下限のみ設定する
        lngResult = SendMessage(mlnghwndMonthView, _
                                MCM_SETRANGE, _
                                GDTR_MIN, _
                                st)
    Else
        ' 既に設定されている上限とともに設定
        dtmMax = DateSerial(st.MaxYear, st.MaxMonth, st.MaxDay)
        If (dtmMax < dtmNewValue) Then
            '上限が新しい下限値より過去になるので、下限値でリセット
            With st
                .MaxYear = .MinYear
                .MaxMonth = .MinMonth
                .MaxDay = .MinDay
                .MaxDayOfWeek = .MinDayOfWeek
                .MaxHour = 0
                .MaxMinute = 0
                .MaxSecond = 0
            End With
        End If
        lngResult = SendMessage(mlnghwndMonthView, _
                                MCM_SETRANGE, _
                                GDTR_MIN Or GDTR_MAX, _
                                st)
    End If
End Property

'-------------------------------------------------------------------------------
' MaxDate(選択可能日付範囲の上限)
Public Property Get MaxDate() As Date
    Dim lngResult As Long
    Dim st As SYSTEMTIMERANGE
    
    lngResult = SendMessage(mlnghwndMonthView, _
                            MCM_GETRANGE, _
                            GDT_VALID, _
                            st)
    MaxDate = DateSerial(st.MaxYear, st.MaxMonth, st.MaxDay)
End Property

Public Property Let MaxDate(ByVal dtmNewValue As Date)
    Dim lngResult As Long
    Dim dtmMin As Date
    Dim st As SYSTEMTIMERANGE

    ' 現在設定されている内容を取得する
    lngResult = SendMessage(mlnghwndMonthView, _
                            MCM_GETRANGE, _
                            GDT_VALID, _
                            st)
    
    ' 取得したSYSTEMTIMERANGEの[Max]を引数の日付に置換
    With st
        .MaxYear = Year(dtmNewValue)
        .MaxMonth = Month(dtmNewValue)
        .MaxDay = Day(dtmNewValue)
        .MaxDayOfWeek = Weekday(dtmNewValue) - 1    ' APIでは日曜日が0
        .MaxHour = 0
        .MaxMinute = 0
        .MaxSecond = 0
    End With
    
    If (st.MinYear = 0) Then
        ' 下限が設定されていないので上限のみ設定する
        lngResult = SendMessage(mlnghwndMonthView, _
                                MCM_SETRANGE, _
                                GDTR_MAX, _
                                st)
    Else
        ' 既に設定されている下限とともに設定
        dtmMin = DateSerial(st.MinYear, st.MinMonth, st.MinDay)
        If (dtmMin > dtmNewValue) Then
            '下限が新しい上限値より未来になるので、上限値でリセット
            With st
                .MinYear = .MaxYear
                .MinMonth = .MaxMonth
                .MinDay = .MaxDay
                .MinDayOfWeek = .MaxDayOfWeek
                .MinHour = 0
                .MinMinute = 0
                .MinSecond = 0
            End With
        End If
        lngResult = SendMessage(mlnghwndMonthView, _
                                MCM_SETRANGE, _
                                GDTR_MAX Or GDTR_MIN, _
                                st)
    End If
End Property

'-------------------------------------------------------------------------------
' MaxSelCount(MultiSelectモードでの選択可能最大日数)
Public Property Get MaxSelCount() As Long
    If (blnMultiSelDays = True) Then
        MaxSelCount = SendMessage(mlnghwndMonthView, _
                                  MCM_GETMAXSELCOUNT, _
                                  GDT_VALID, _
                                  ByVal GDT_VALID)
    Else
        MaxSelCount = 1
    End If
End Property

Public Property Let MaxSelCount(lngNewSelCount As Long)
    Dim lngResult As Long
    If (blnMultiSelDays = True) Then
        lngResult = SendMessage(mlnghwndMonthView, _
                                MCM_SETMAXSELCOUNT, _
                                lngNewSelCount, _
                                ByVal GDT_VALID)
    Else
        'MultiSelectモード以外では不可
    End If
End Property

'-------------------------------------------------------------------------------
' SelStart(MultiSelectモードでの選択日付の初日)
Public Property Get SelStart() As Date
    Dim lngResult As Long
    Dim st As SYSTEMTIMERANGE
    
    If (blnMultiSelDays = True) Then
        lngResult = SendMessage(mlnghwndMonthView, _
                                MCM_GETSELRANGE, _
                                GDT_VALID, _
                                st)
        SelStart = DateSerial(st.MinYear, st.MinMonth, st.MinDay)
    Else
        'MultiSelectモード以外では不可
    End If
End Property

Public Property Let SelStart(ByVal dtmNewValue As Date)
    Dim lngResult As Long
    Dim dtmMax As Date
    Dim st As SYSTEMTIMERANGE
    
    If (blnMultiSelDays = True) Then
      ' 現在設定されている内容を取得する
      lngResult = SendMessage(mlnghwndMonthView, _
                              MCM_GETSELRANGE, _
                              GDT_VALID, _
                              st)
    
      ' 取得したSYSTEMTIMERANGEの[Min]を引数の日付に置換
      With st
        .MinYear = Year(dtmNewValue)
        .MinMonth = Month(dtmNewValue)
        .MinDay = Day(dtmNewValue)
        .MinDayOfWeek = Weekday(dtmNewValue) - 1     ' APIでは日曜日が0
        .MinHour = 0
        .MinMinute = 0
        .MinSecond = 0
      End With
    
      If (st.MaxYear = 0) Then
        ' 最終日が設定されていないので初日のみ設定する
        lngResult = SendMessage(mlnghwndMonthView, _
                                MCM_SETSELRANGE, _
                                GDTR_MIN, _
                                st)
      Else
        ' 既に設定されている最終日とともに設定
        dtmMax = DateSerial(st.MaxYear, st.MaxMonth, st.MaxDay)
        If (dtmMax < dtmNewValue) Then
            '最終日が新しい初日より過去になるので、初日でリセット
            With st
                .MaxYear = .MinYear
                .MaxMonth = .MinMonth
                .MaxDay = .MinDay
                .MaxDayOfWeek = .MinDayOfWeek
                .MaxHour = 0
                .MaxMinute = 0
                .MaxSecond = 0
            End With
        End If
        lngResult = SendMessage(mlnghwndMonthView, _
                                MCM_SETSELRANGE, _
                                GDTR_MIN Or GDTR_MAX, _
                                st)
      End If
    Else
        'MultiSelectモード以外では不可
    End If
End Property

'-------------------------------------------------------------------------------
' SelEnd(MultiSelectモードでの選択日付の最終日)
Public Property Get SelEnd() As Date
    Dim lngResult As Long
    Dim st As SYSTEMTIMERANGE
    
    If (blnMultiSelDays = True) Then
        lngResult = SendMessage(mlnghwndMonthView, _
                                MCM_GETSELRANGE, _
                                GDT_VALID, _
                                st)
        SelEnd = DateSerial(st.MaxYear, st.MaxMonth, st.MaxDay)
    Else
        'MultiSelectモード以外では不可
    End If
End Property

Public Property Let SelEnd(ByVal dtmNewValue As Date)
    Dim lngResult As Long
    Dim dtmMin As Date
    Dim st As SYSTEMTIMERANGE

    If (blnMultiSelDays = True) Then
      ' 現在設定されている内容を取得する
      lngResult = SendMessage(mlnghwndMonthView, _
                              MCM_GETSELRANGE, _
                              GDT_VALID, _
                              st)
    
      ' 取得したSYSTEMTIMERANGEの[Max]を引数の日付に置換
      With st
        .MaxYear = Year(dtmNewValue)
        .MaxMonth = Month(dtmNewValue)
        .MaxDay = Day(dtmNewValue)
        .MaxDayOfWeek = Weekday(dtmNewValue) - 1    ' APIでは日曜日が0
        .MaxHour = 0
        .MaxMinute = 0
        .MaxSecond = 0
      End With
    
      If (st.MinYear = 0) Then
        ' 初日が設定されていないので最終日のみ設定する
        lngResult = SendMessage(mlnghwndMonthView, _
                                MCM_SETSELRANGE, _
                                GDTR_MAX, _
                                st)
      Else
        ' 既に設定されている初日とともに設定
        dtmMin = DateSerial(st.MinYear, st.MinMonth, st.MinDay)
        If (dtmMin > dtmNewValue) Then
            '初日が新しい最終日より未来になるので、最終日でリセット
            With st
                .MinYear = .MaxYear
                .MinMonth = .MaxMonth
                .MinDay = .MaxDay
                .MinDayOfWeek = .MaxDayOfWeek
                .MinHour = 0
                .MinMinute = 0
                .MinSecond = 0
            End With
        End If
        lngResult = SendMessage(mlnghwndMonthView, _
                                MCM_SETSELRANGE, _
                                GDTR_MAX Or GDTR_MIN, _
                                st)
      End If
    Else
        'MultiSelectモード以外では不可
    End If
End Property

'-------------------------------------------------------------
' StartOfWeek  週の始まりの曜日(vbSunday[1]～vbSaturday[7])
' ※ここでAPIに引き継ぐ曜日値[Day of Week]は他と設定が異なるので注意
'       0:Monday    1:Tuesday   2:Wednesday 3:Thursday
'       4:Friday    5:Saturday  6:Sunday
'   外部とのインターフェースはvbSunday[1]～vbSaturday[7]で行なう
Public Property Get StartOfWeek() As Integer
    Dim lngResult As Long
    lngResult = SendMessage(mlnghwndMonthView, _
                            MCM_GETFIRSTDAYOFWEEK, _
                            GDT_VALID, _
                            ByVal GDT_VALID)
    If (lngResult = 6) Then
        StartOfWeek = vbSunday
    Else
        StartOfWeek = lngResult + 2
    End If
End Property

Public Property Let StartOfWeek(ByVal intNewValue As Integer)
    Dim lngResult As Long
    Dim lngStartWeek As Long
    If (intNewValue = vbSunday) Then
        lngStartWeek = 6
    Else
        lngStartWeek = intNewValue - 2
    End If
    lngResult = SendMessage(mlnghwndMonthView, _
                            MCM_SETFIRSTDAYOFWEEK, _
                            GDT_VALID, _
                            ByVal lngStartWeek)
End Property

'--------------------------------------------------------------
'カレンダー背景色
Public Property Get CalendarBackColor() As Long
    Dim lngResult As Long
    lngResult = SendMessage(mlnghwndMonthView, _
                            MCM_GETCOLOR, _
                            MCSC_MONTHBK, _
                            ByVal 0&)
    CalendarBackColor = lngResult
End Property

Public Property Let CalendarBackColor(ByVal lngNewColor As Long)
    Dim lngResult As Long
    Dim lngRGB As Long
    Const cstMask As Long = &H7FFFFFFF  'SystemColor定数の先頭1bitをoffに
    Select Case lngNewColor
      'vbScrollBars(&H80000000)～vbInfoBackground(&H80000018)
      Case vbScrollBars To vbInfoBackground
        lngRGB = GetSysColor(lngNewColor And cstMask)
      Case Else
        lngRGB = lngNewColor
    End Select
    lngResult = SendMessage(mlnghwndMonthView, _
                            MCM_SETCOLOR, _
                            MCSC_MONTHBK, _
                            ByVal lngRGB)
End Property

'-------------------------------------------------------------
'カレンダー文字色
Public Property Get CalendarForeColor() As Long
    Dim lngResult As Long
    lngResult = SendMessage(mlnghwndMonthView, _
                            MCM_GETCOLOR, _
                            MCSC_TEXT, _
                            ByVal 0&)
    CalendarForeColor = lngResult
End Property

Public Property Let CalendarForeColor(ByVal lngNewColor As Long)
    Dim lngResult As Long
    Dim lngRGB As Long
    Const cstMask As Long = &H7FFFFFFF
    Select Case lngNewColor
      'vbScrollBars(&H80000000)～vbInfoBackground(&H80000018)
      Case vbScrollBars To vbInfoBackground
        lngRGB = GetSysColor(lngNewColor And cstMask)
      Case Else
        lngRGB = lngNewColor
    End Select
    lngResult = SendMessage(mlnghwndMonthView, _
                            MCM_SETCOLOR, _
                            MCSC_TEXT, _
                            ByVal lngRGB)
End Property

'-------------------------------------------------------------
'[前月/翌月]文字色
Public Property Get CalendarTrailingForeColor() As Long
    Dim lngResult As Long
    lngResult = SendMessage(mlnghwndMonthView, _
                            MCM_GETCOLOR, _
                            MCSC_TRAILINGTEXT, _
                            ByVal 0&)
    CalendarTrailingForeColor = lngResult
End Property

Public Property Let CalendarTrailingForeColor(ByVal lngNewColor As Long)
    Dim lngResult As Long
    Dim lngRGB As Long
    Const cstMask As Long = &H7FFFFFFF
    Select Case lngNewColor
      'vbScrollBars(&H80000000)～vbInfoBackground(&H80000018)
      Case vbScrollBars To vbInfoBackground
        lngRGB = GetSysColor(lngNewColor And cstMask)
      Case Else
        lngRGB = lngNewColor
    End Select
    lngResult = SendMessage(mlnghwndMonthView, _
                            MCM_SETCOLOR, _
                            MCSC_TRAILINGTEXT, _
                            ByVal lngRGB)
End Property

'------------------------------------------------------------
'タイトルバー背景色/曜日文字色
Public Property Get CalendarTitleBackColor() As Long
    Dim lngResult As Long
    lngResult = SendMessage(mlnghwndMonthView, _
                            MCM_GETCOLOR, _
                            MCSC_TITLEBK, _
                            ByVal 0&)
    CalendarTitleBackColor = lngResult
End Property

Public Property Let CalendarTitleBackColor(ByVal lngNewColor As Long)
    Dim lngResult As Long
    Dim lngRGB As Long
    Const cstMask As Long = &H7FFFFFFF
    Select Case lngNewColor
      'vbScrollBars(&H80000000)～vbInfoBackground(&H80000018)
      Case vbScrollBars To vbInfoBackground
        lngRGB = GetSysColor(lngNewColor And cstMask)
      Case Else
        lngRGB = lngNewColor
    End Select
    lngResult = SendMessage(mlnghwndMonthView, _
                            MCM_SETCOLOR, _
                            MCSC_TITLEBK, _
                            ByVal lngRGB)
End Property

'-------------------------------------------------------------
'タイトルバー文字色
Public Property Get CalendarTitleForeColor() As Long
    Dim lngResult As Long
    lngResult = SendMessage(mlnghwndMonthView, _
                            MCM_GETCOLOR, _
                            MCSC_TITLETEXT, _
                            ByVal 0&)
    CalendarTitleForeColor = lngResult
End Property

Public Property Let CalendarTitleForeColor(ByVal lngNewColor As Long)
    Dim lngResult As Long
    Dim lngRGB As Long
    Const cstMask As Long = &H7FFFFFFF
    Select Case lngNewColor
      'vbScrollBars(&H80000000)～vbInfoBackground(&H80000018)
      Case vbScrollBars To vbInfoBackground
        lngRGB = GetSysColor(lngNewColor And cstMask)
      Case Else
        lngRGB = lngNewColor
    End Select
    lngResult = SendMessage(mlnghwndMonthView, _
                            MCM_SETCOLOR, _
                            MCSC_TITLETEXT, _
                            ByVal lngRGB)
End Property


'*******************************************************************************
'   メソッドの記述
'*******************************************************************************

' [Create メソッド]
'   1op  False(既定):Flat表示 or True:3D表示
'   2op  縦方向に並べる数(既定値:1)
'   3op  横方向に並べる数(既定値:1)
'   4op  選択日数(既定値:1) 2以上でMultiSelectモード
Public Sub Create(Optional ByVal MV3D As Boolean = False, _
                  Optional ByVal MVRows As Integer = 1, _
                  Optional ByVal MVColumns As Integer = 1, _
                  Optional ByVal MVMultiSelDays As Integer = 1)
    Dim icce            As tagINITCOMMONCONTROLSEX
    Dim lngResult       As Long
    Dim lnghInstance    As Long
    Dim lnghWnd_Sub     As Long
    Dim strThunder      As String
    Dim lngWindowStyle As Long
    Dim lngWindowStyleEX As Long

    blnMV3D = MV3D            '3D/Flat

    'Max12個まで展開可
    If (MVRows >= 1) And (MVRows <= 12) And _
       (MVColumns >= 1) And (MVColumns <= 12) And _
       ((MVRows * MVColumns) >= 1) And _
       ((MVRows * MVColumns) <= 12) Then
        intMVRows = MVRows        '行数
        intMVColumns = MVColumns  '列数
    Else
        intMVRows = 1
        intMVColumns = 1
    End If
    
    'MultiSelect
    If (MVMultiSelDays <= 1) Then
        blnMultiSelDays = False
        lngWindowStyle = WS_CHILD Or WS_VISIBLE Or WS_BORDER
    Else
        blnMultiSelDays = True
        lngWindowStyle = WS_CHILD Or WS_VISIBLE Or WS_BORDER Or MCS_MULTISELECT
    End If
    
    '3D 表示
    If (blnMV3D = False) Then
        lngWindowStyleEX = GDT_VALID
    Else
        lngWindowStyleEX = WS_EX_CLIENTEDGE
    End If

    If Val(Application.Version) <= 8 Then
        strThunder = "ThunderXFrame"        ' Excel97
    Else
        strThunder = "ThunderDFrame"        ' Excel2000～
    End If
    
    ' 既にウィンドウが存在する場合はｳｨﾝﾄﾞｳの破棄
    If IsWindow(mlnghwndMonthView) <> 0 Then
        Call DestroyWindow(mlnghwndMonthView)
    End If
    
    ' INITCOMMONCONTROLSEX構造体に値を代入
    With icce
        .dwICC = ICC_DATE_CLASSES
        .dwSize = Len(icce)
    End With
    
    ' コモンコントロールを初期化
    lngResult = InitCommonControlsEx(icce)
    
    ' ユーザーフォームのHWNDの取得
    lnghWnd_Form = FindWindow(strThunder, mctlForm.Caption)
    If lnghWnd_Form = 0 Then Exit Sub

    ' ポイント→ピクセル変換係数算出
    Call GetLogPixelsXY
    
    lnghWnd_Excel = FindWindow("XLMAIN", Application.Caption)
    lnghInstance = GetWindowLong(lnghWnd_Excel, GWL_HINSTANCE)
        
    ' 透明ウィンドウのHWNDの取得
    lnghWnd_Sub = FindWindowEx(lnghWnd_Form, 0&, vbNullString, vbNullString)
    
    ' Labelコントロールの左上端を基点にMonthViewの作成
    ' MonthViewの必要高/幅のMinHeight/Widthを取得できるように
    ' 一端、幅/高さをゼロサイズで作成する
    mlnghwndMonthView = _
        CreateWindowEx(lngWindowStyleEX, MONTHCAL_CLASS, _
                       vbNullString, lngWindowStyle, _
                       mctlLabel.Left * lngPixelsX / 72, _
                       mctlLabel.Top * lngPixelsY / 72, _
                       0&, 0&, _
                       lnghWnd_Sub, 0&, lnghInstance, vbNullString)
        
    ' 行数/列数/3D/Flatに応じたMonthView の必要サイズに調整する
    Call MoveWindow(mlnghwndMonthView, _
                    mctlLabel.Left * lngPixelsX / 72, _
                    mctlLabel.Top * lngPixelsY / 72, _
                    Me.NecessaryWidth * lngPixelsX / 72, _
                    Me.NecessaryHeight * lngPixelsY / 72, _
                    1&)
    
    ' MultiSelect指定時の最大選択日数を設定
    If (blnMultiSelDays = True) Then
        Me.MaxSelCount = MVMultiSelDays
    End If

    '選択可能日付範囲の初期設定
    Me.MinDate = DateValue("1900/1/1")
    Me.MaxDate = DateValue("9999/12/31")
End Sub

'-------------------------------------------------------------------------------
' Destroy
Public Sub Destroy()
    ' MonthViewのウィンドウクラスの破棄
    Call Class_Terminate
End Sub

'-------------------------------------------------------------------------------
' DispNecessarySize
Public Sub DispNecessarySize()
    MsgBox "Flat(False)/3D(True) is [ " & blnMV3D & " ]" & vbCrLf & _
           "Rows( " & intMVRows & " ) * Columns( " & intMVColumns & " )" & vbCrLf & vbCrLf & _
           "[Necessary Size]" & vbCrLf & _
           "   Height( " & Me.NecessaryHeight & " )  Width( " & Me.NecessaryWidth & " )", _
            Title:="MonthView Control by API"
End Sub


'*******************************************************************************
'   イベントの記述
'*******************************************************************************
' ClassのInitialize
Private Sub Class_Initialize()
' クラスの内の変数の初期値
    mlnghwndMonthView = 0
End Sub

' ClassのTerminate
Private Sub Class_Terminate()
    ' ウィンドウが存在する場合ウィンドウの破棄
    If (IsWindow(mlnghwndMonthView) <> 0) Then
        Call DestroyWindow(mlnghwndMonthView)
    End If
End Sub

'*******************************************************************************
'   クラス内共通処理の記述
'*******************************************************************************
' 画面精細度のポイント→ピクセル変換係数算出
Private Sub GetLogPixelsXY()
     Dim lnghwnd As Long
     Dim lngDC As Long

     'GetDCを使った後は、必ずReleaseDCで元に戻す。
     '画面表示用の専用メモリ(リソース)が消費され、
     '最後はWindowsがハングアップする危険があります。
     lnghwnd = GetDesktopWindow()
     lngDC = GetDC(lnghwnd)
     lngPixelsX = GetDeviceCaps(lngDC, LOGPIXELSX)
     lngPixelsY = GetDeviceCaps(lngDC, LOGPIXELSY)
     ReleaseDC lnghwnd, lngDC
End Sub



