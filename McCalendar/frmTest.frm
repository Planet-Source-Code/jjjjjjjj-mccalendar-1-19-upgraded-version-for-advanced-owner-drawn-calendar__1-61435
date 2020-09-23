VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "McCalendar 1.19        By, ""Jim Jose"""
   ClientHeight    =   6090
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   10125
   FillColor       =   &H000000FF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   10125
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      Height          =   4935
      Left            =   8160
      ScaleHeight     =   4875
      ScaleWidth      =   1635
      TabIndex        =   63
      Top             =   240
      Width           =   1695
      Begin VB.OptionButton optModes 
         Alignment       =   1  'Right Justify
         Caption         =   "Modes And Special Days"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   65
         Top             =   2640
         Width           =   1335
      End
      Begin VB.OptionButton optAppearance 
         Alignment       =   1  'Right Justify
         Caption         =   "Appearance"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   68
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optColors 
         Alignment       =   1  'Right Justify
         Caption         =   "Custom Colors"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   67
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optdatePicker 
         Alignment       =   1  'Right Justify
         Caption         =   "DatePicker Demonstration"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   64
         Top             =   3600
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optSettings 
         Alignment       =   1  'Right Justify
         Caption         =   "Settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   66
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Select Pages"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         TabIndex        =   69
         Top             =   0
         Width           =   1080
      End
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   5280
      Width           =   9855
      Begin AdvancedCalendar.McCalendar McCalendar2 
         Height          =   270
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Theme           =   6
         Animate         =   -1  'True
         CalendarHeight  =   363
         Mode            =   2
         CalendarBackCol =   10985207
         MonthBackCol    =   10985207
         HeaderBackCol   =   10985207
         WeekDayCol      =   11446008
         DayCol          =   16777215
         DaySelCol       =   14934998
         WeekDaySelCol   =   15857131
         DaySunCol       =   14078715
         WeekDaySunCol   =   14078715
         YearBackCol     =   10985207
         SpecialDays     =   $"frmTest.frx":000C
      End
   End
   Begin AdvancedCalendar.McCalendar McCalendar1 
      Height          =   3375
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5953
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Animate         =   -1  'True
      CalendarHeight  =   207
      CalendarBackCol =   16743805
      SpecialDays     =   $"frmTest.frx":00A0
      BorderColor     =   8388608
      ToolTipForeCol  =   -2147483625
   End
   Begin VB.Frame Frame2 
      Caption         =   "The Date informations : Skip to any Century u wish to go"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   10
      Top             =   3960
      Width           =   4815
      Begin VB.TextBox txtCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   76
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox txtyear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   2760
         TabIndex        =   72
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtDay 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   3240
         TabIndex        =   75
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtMonth 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   840
         TabIndex        =   73
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lbDay 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Day"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4335
         TabIndex        =   14
         Top             =   360
         Width           =   315
      End
      Begin VB.Label lbMonth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   525
      End
      Begin VB.Label lbYear 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4260
         TabIndex        =   12
         Top             =   600
         Width           =   390
      End
      Begin VB.Label lbDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   390
      End
   End
   Begin VB.Frame frDatePicker 
      Caption         =   "Date Picker"
      Height          =   5055
      Left            =   5280
      TabIndex        =   58
      Top             =   120
      Width           =   2775
      Begin VB.CommandButton cmdPopDown 
         Height          =   375
         Left            =   2160
         Picture         =   "frmTest.frx":0134
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   4200
         Width           =   375
      End
      Begin VB.TextBox txtDateDown 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   78
         Text            =   "DatePicker Down"
         Top             =   4200
         Width           =   1815
      End
      Begin VB.TextBox txtDateUp 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   61
         Text            =   "DatePicker Up"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CommandButton cmdPopUp 
         Height          =   375
         Left            =   240
         Picture         =   "frmTest.frx":04BE
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   1680
         Width           =   375
      End
      Begin AdvancedCalendar.McCalendar McCalendar4 
         Height          =   375
         Left            =   240
         TabIndex        =   62
         Top             =   1680
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Theme           =   3
         Animate         =   -1  'True
         CalendarHeight  =   180
         Mode            =   2
         CalendarBackCol =   11651986
         MonthGradient   =   0   'False
         MonthBackCol    =   11651986
         HeaderBackCol   =   11651986
         WeekDayCol      =   16443612
         DayCol          =   14805973
         DaySelCol       =   14334632
         WeekDaySelCol   =   15857131
         DaySunCol       =   11446008
         YearBackCol     =   11651986
         HeaderHeight    =   25
         SpecialDays     =   $"frmTest.frx":0848
      End
      Begin AdvancedCalendar.McCalendar McCalendar3 
         Height          =   375
         Left            =   -1680
         TabIndex        =   80
         Top             =   4200
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Theme           =   5
         Animate         =   -1  'True
         CalendarHeight  =   180
         Mode            =   1
         CalendarBackCol =   14004415
         MonthGradient   =   0   'False
         MonthBackCol    =   14004415
         HeaderBackCol   =   14004415
         WeekDayCol      =   13740473
         DayCol          =   16249331
         DaySelCol       =   14934998
         WeekDaySelCol   =   15857131
         DaySunCol       =   11651986
         WeekDaySunCol   =   14004415
         YearBackCol     =   14004415
         HeaderHeight    =   25
         SpecialDays     =   $"frmTest.frx":08DC
      End
      Begin VB.Label Label13 
         Caption         =   "Pop the datepicker and move this form. The calendar will capture the movement via subclassing and reposition itself again"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   81
         Top             =   3240
         Width           =   2535
      End
      Begin VB.Label Label10 
         Caption         =   "You can use McCalendar as a standard DatePicker. The PopUp operation can be stimulated externally"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   59
         Top             =   600
         Width           =   2535
      End
   End
   Begin VB.Frame frColors 
      Caption         =   "You got full range of color options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   5280
      TabIndex        =   5
      Top             =   120
      Width           =   2775
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   4095
         Left            =   120
         ScaleHeight     =   4095
         ScaleWidth      =   2535
         TabIndex        =   40
         Top             =   240
         Width           =   2535
         Begin VB.OptionButton optForeColor 
            Caption         =   "ForeColor"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   71
            Top             =   2160
            Width           =   2175
         End
         Begin VB.OptionButton optBorderCol 
            Caption         =   "BorderColor"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   70
            Top             =   2400
            Width           =   2175
         End
         Begin VB.OptionButton optHeaderBackCol 
            Caption         =   "HeaderBackCol"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   54
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton optHeaderGradientCol 
            Caption         =   "HeaderGradientCol"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   0
            TabIndex        =   53
            Top             =   360
            Width           =   2175
         End
         Begin VB.OptionButton optMonthBackCol 
            Caption         =   "MonthBackCol"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   0
            TabIndex        =   52
            Top             =   600
            Width           =   2175
         End
         Begin VB.OptionButton optMonthGradientCol 
            Caption         =   "MonthGradientCol"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   0
            TabIndex        =   51
            Top             =   840
            Width           =   2055
         End
         Begin VB.OptionButton optYearBackCol 
            Caption         =   "YearBackCol"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   50
            Top             =   1080
            Width           =   2055
         End
         Begin VB.OptionButton optYearGradientCol 
            Caption         =   "YearGradientCol"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   49
            Top             =   1320
            Width           =   2055
         End
         Begin VB.OptionButton optDaySelCol 
            Caption         =   "DaySelCol"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   48
            Top             =   2880
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton optDaySunCol 
            Caption         =   "DaySunCol"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   47
            Top             =   3120
            Width           =   2175
         End
         Begin VB.OptionButton optWeekDaySunCol 
            Caption         =   "WeekDaySunCol"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   46
            Top             =   3840
            Width           =   2175
         End
         Begin VB.OptionButton optWeekDaySelCol 
            Caption         =   "WeekDaySelCol"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   45
            Top             =   3600
            Width           =   2175
         End
         Begin VB.OptionButton optWeekDayCol 
            Caption         =   "WeekDayCol"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   44
            Top             =   3360
            Width           =   2175
         End
         Begin VB.OptionButton optDayCol 
            Caption         =   "DayCol"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   43
            Top             =   2640
            Width           =   2175
         End
         Begin VB.OptionButton optCalendarGradientCol 
            Caption         =   "CalendarGradientCol"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   42
            Top             =   1800
            Width           =   2295
         End
         Begin VB.OptionButton optCalendarBackCol 
            Caption         =   "CalendarBackCol"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   41
            Top             =   1560
            Width           =   2175
         End
      End
      Begin VB.PictureBox picSel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   360
         Picture         =   "frmTest.frx":0970
         ScaleHeight     =   405
         ScaleWidth      =   2145
         TabIndex        =   6
         Top             =   4440
         Width           =   2175
      End
   End
   Begin VB.Frame frModes 
      Caption         =   "Modes and Special Days"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   5280
      TabIndex        =   22
      Top             =   120
      Width           =   2775
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         ScaleHeight     =   735
         ScaleWidth      =   2535
         TabIndex        =   55
         Top             =   1920
         Width           =   2535
         Begin VB.CheckBox chkHeaderVisible 
            Caption         =   "Header Visible"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   480
            Value           =   1  'Checked
            Width           =   2415
         End
         Begin VB.Label Label9 
            Caption         =   "Property Using For DatePicker"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   120
            TabIndex        =   57
            Top             =   0
            Width           =   1860
         End
      End
      Begin VB.TextBox txtSpecial 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   23
         Text            =   "frmTest.frx":3742
         Top             =   3960
         Width           =   2535
      End
      Begin VB.ListBox lstMode 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         ItemData        =   "frmTest.frx":37C4
         Left            =   120
         List            =   "frmTest.frx":37D1
         TabIndex        =   25
         Top             =   720
         Width           =   2535
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   120
         X2              =   2640
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "String for Indian standard Calendar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   77
         Top             =   4560
         Width           =   2550
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Modes : Select the mode as u need"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   2520
      End
      Begin VB.Label Label7 
         Caption         =   $"frmTest.frx":380C
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Left            =   120
         TabIndex        =   24
         Top             =   2880
         Width           =   2415
      End
   End
   Begin VB.Frame frSettings 
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   5280
      TabIndex        =   19
      Top             =   120
      Width           =   2775
      Begin VB.ComboBox cmbFirstDay 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmTest.frx":38AC
         Left            =   120
         List            =   "frmTest.frx":38C5
         TabIndex        =   27
         Text            =   "[SunDay] = 1"
         Top             =   4440
         Width           =   2535
      End
      Begin VB.ComboBox cmbFormat 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmTest.frx":3933
         Left            =   120
         List            =   "frmTest.frx":3940
         TabIndex        =   20
         Text            =   "[dd-mm-yyyy]"
         Top             =   3600
         Width           =   2535
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   120
         ScaleHeight     =   3135
         ScaleWidth      =   2535
         TabIndex        =   29
         Top             =   240
         Width           =   2535
         Begin VB.CheckBox chkGradient 
            Caption         =   "HeaderGradient"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.CheckBox chkGradient 
            Caption         =   "YearGradient"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   120
            TabIndex        =   35
            Top             =   960
            Width           =   2175
         End
         Begin VB.CheckBox chkGradient 
            Caption         =   "MonthGradient"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   34
            Top             =   600
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.CheckBox chkGradient 
            Caption         =   "CalendarGradient"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   120
            TabIndex        =   33
            Top             =   1320
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.CheckBox chkSkip 
            Caption         =   "Skip Enabled"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   32
            Top             =   2280
            Width           =   2175
         End
         Begin VB.CheckBox chkSensitive 
            Caption         =   "Sensitive (PopUp mode)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   1920
            Value           =   1  'Checked
            Width           =   2535
         End
         Begin VB.CheckBox chkAnimate 
            Caption         =   "Animate"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   2640
            Value           =   1  'Checked
            Width           =   2055
         End
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "FirstDay Of Week"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   28
         Top             =   4200
         Width           =   1830
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Date Format"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   21
         Top             =   3360
         Width           =   1290
      End
   End
   Begin VB.Frame frAppearance 
      Caption         =   "Appearance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   5280
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.TextBox txtCurvature 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         TabIndex        =   1
         Text            =   "0"
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   240
         ScaleHeight     =   735
         ScaleWidth      =   2415
         TabIndex        =   37
         Top             =   480
         Width           =   2415
         Begin VB.CheckBox chkBorder 
            Caption         =   "Border"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   39
            Top             =   0
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.CheckBox chkAppearance 
            Caption         =   "3D Appearance"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   38
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.ListBox lstTheme 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         ItemData        =   "frmTest.frx":396E
         Left            =   120
         List            =   "frmTest.frx":398A
         TabIndex        =   17
         Top             =   3480
         Width           =   2535
      End
      Begin VB.TextBox txtHeight 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Text            =   "195"
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox txtHeader 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Text            =   "18"
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Themes : Eight different color themes. Select the one u need."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   18
         Top             =   3000
         Width           =   2445
      End
      Begin VB.Label Label6 
         Caption         =   "Header Height"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   9
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Calendar Height"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   4
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Curvature"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAnimate_Click()
    McCalendar1.Animate = chkAnimate
End Sub

Private Sub chkAppearance_Click()
    McCalendar1.Appearance = chkAppearance
End Sub

Private Sub chkBorder_Click()
    McCalendar1.BorderStyle = chkBorder
End Sub

Private Sub cmdBack_Click()
    Unload Me
End Sub

Private Sub chkGradient_Click(Index As Integer)
Select Case Index
    Case 0
        McCalendar1.HeaderGradient = chkGradient(0)
    Case 1
        McCalendar1.MonthGradient = chkGradient(1)
    Case 2
        McCalendar1.YearGradient = chkGradient(2)
    Case 3
        McCalendar1.CalendarGradient = chkGradient(3)
End Select
End Sub

Private Sub chkHeaderVisible_Click()
    McCalendar1.HeaderVisible = chkHeaderVisible
End Sub

Private Sub chkSensitive_Click()
    McCalendar1.Sensitive = chkSensitive
End Sub

Private Sub chkSkip_Click()
    McCalendar1.SkipEnabled = chkSkip
End Sub

Private Sub cmbFirstDay_Click()
    McCalendar1.FirstDayOfWeek = cmbFirstDay.ListIndex + 1
End Sub

Private Sub cmbFormat_Click()
    McCalendar1.DateFormat = cmbFormat.ListIndex
End Sub

Private Sub cmdPopDown_Click()
    McCalendar3.PopUpCalendar
End Sub

Private Sub cmdPopDown_LostFocus()
    McCalendar3.CollapseCalendar False
End Sub

Private Sub cmdPopUp_Click()
    McCalendar4.PopUpCalendar
End Sub

Private Sub cmdPopUp_LostFocus()
    McCalendar4.CollapseCalendar False
End Sub


Private Sub Form_Load()
    McCalendar1_DateChanged
    ' The two DatePicker calendar
    McCalendar3.HeaderVisible = False
    McCalendar4.HeaderVisible = False
End Sub

Private Sub lstMode_Click()
    McCalendar1.Mode = lstMode.ListIndex
End Sub

Private Sub lstTheme_Click()
    McCalendar1.Theme = lstTheme.ListIndex + 1
End Sub

Private Sub McCalendar1_DateChanged()
    txtDate = McCalendar1
    txtDay = McCalendar1.DayX
    txtMonth = McCalendar1.MonthX
    txtyear = McCalendar1.YearX
    txtCaption = McCalendar1.Caption(True)
End Sub

Private Sub McCalendar3_DateChanged()
    txtDateDown = McCalendar3.DateX
End Sub

Private Sub McCalendar4_DateChanged()
    txtDateUp = McCalendar4.DateX
End Sub

Private Sub optAppearance_Click()
    frAppearance.ZOrder (0)
End Sub

Private Sub optColors_Click()
    frColors.ZOrder (0)
End Sub

Private Sub optdatePicker_Click()
    frDatePicker.ZOrder (0)
End Sub

Private Sub optModes_Click()
    frModes.ZOrder (0)
End Sub

Private Sub optSettings_Click()
    frSettings.ZOrder (0)
End Sub

Private Sub picSel_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If optHeaderBackCol Then McCalendar1.HeaderBackCol = picSel.Point(x, Y)
    If optHeaderGradientCol Then McCalendar1.HeaderGradientCol = picSel.Point(x, Y)
    
    If optMonthBackCol Then McCalendar1.MonthBackCol = picSel.Point(x, Y)
    If optMonthGradientCol Then McCalendar1.MonthGradientCol = picSel.Point(x, Y)
    
    If optYearBackCol Then McCalendar1.YearBackCol = picSel.Point(x, Y)
    If optYearGradientCol Then McCalendar1.YearGradientCol = picSel.Point(x, Y)
    
    If optCalendarBackCol Then McCalendar1.CalendarBackCol = picSel.Point(x, Y)
    If optCalendarGradientCol Then McCalendar1.CalendarGradientCol = picSel.Point(x, Y)
    
    If optDayCol Then McCalendar1.DayCol = picSel.Point(x, Y)
    If optDaySelCol Then McCalendar1.DaySelCol = picSel.Point(x, Y)
    If optDaySunCol Then McCalendar1.DaySunCol = picSel.Point(x, Y)
    
    If optWeekDayCol Then McCalendar1.WeekDayCol = picSel.Point(x, Y)
    If optWeekDaySelCol Then McCalendar1.WeekDaySelCol = picSel.Point(x, Y)
    If optWeekDaySunCol Then McCalendar1.WeekDaySunCol = picSel.Point(x, Y)

    If optBorderCol Then McCalendar1.BorderColor = picSel.Point(x, Y)
    If optForeColor Then McCalendar1.ForeColor = picSel.Point(x, Y)
    
End Sub

Private Sub txtCurvature_Change()
    McCalendar1.Curvature = Val(txtCurvature)
End Sub

Private Sub txtDay_KeyDown(KeyCode As Integer, Shift As Integer)
    DoEvents
    McCalendar1.DayX = Val(txtDay)
End Sub

Private Sub txtHeader_Change()
    McCalendar1.HeaderHeight = Val(txtHeader)
End Sub

Private Sub txtHeight_Change()
    McCalendar1.CalendarHeight = Val(txtHeight)
End Sub

Private Sub txtMonth_KeyDown(KeyCode As Integer, Shift As Integer)
    DoEvents
    McCalendar1.MonthX = Val(txtMonth)
End Sub

Private Sub txtSpecial_Change()
    McCalendar1.SpecialDays = txtSpecial
End Sub

Private Sub txtyear_KeyDown(KeyCode As Integer, Shift As Integer)
    DoEvents
    McCalendar1.YearX = Val(txtyear)
End Sub

