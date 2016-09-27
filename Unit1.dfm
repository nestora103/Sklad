object Form1: TForm1
  Left = 225
  Top = 217
  Width = 665
  Height = 337
  Caption = 'Sklad'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object GroupBox1: TGroupBox
    Left = 8
    Top = 8
    Width = 641
    Height = 121
    Caption = 'ExcelToWord'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 0
    object Label1: TLabel
      Left = 8
      Top = 16
      Width = 203
      Height = 20
      Caption = #1042#1099#1073#1077#1088#1080#1090#1077' '#1092#1072#1081#1083' '#1089#1082#1083#1072#1076#1072
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Button1: TButton
      Left = 8
      Top = 48
      Width = 201
      Height = 33
      Caption = #1042#1099#1073#1088#1072#1090#1100' '#1092#1072#1081#1083' '#1089#1082#1083#1072#1076#1072
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 0
      OnClick = Button1Click
    end
    object StringGrid1: TStringGrid
      Left = 8
      Top = 88
      Width = 17
      Height = 17
      Color = clMenu
      ColCount = 10
      DefaultColWidth = 100
      DefaultRowHeight = 30
      RowCount = 10
      GridLineWidth = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goRowSizing, goColSizing]
      TabOrder = 1
    end
    object RadioButton1: TRadioButton
      Left = 256
      Top = 32
      Width = 113
      Height = 17
      Caption = #1058#1088#1077#1073#1086#1074#1072#1085#1080#1103
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 2
    end
    object RadioButton2: TRadioButton
      Left = 256
      Top = 56
      Width = 113
      Height = 17
      Caption = #1053#1072#1082#1083#1072#1076#1085#1072#1103
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 3
    end
    object Button2: TButton
      Left = 400
      Top = 40
      Width = 177
      Height = 41
      Caption = #1047#1072#1087#1086#1083#1085#1080#1090#1100
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 4
      OnClick = Button2Click
    end
    object CheckBox1: TCheckBox
      Left = 256
      Top = 96
      Width = 217
      Height = 17
      Caption = #1044#1072#1090#1072' '#1089#1086#1079#1076#1072#1085#1080#1103' '#1076#1086#1082#1091#1084#1077#1085#1090#1072
      TabOrder = 5
    end
  end
  object GroupBox2: TGroupBox
    Left = 8
    Top = 136
    Width = 641
    Height = 161
    Caption = 'ExcelOrder'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 1
    object Label5: TLabel
      Left = 16
      Top = 24
      Width = 194
      Height = 20
      Caption = #1042#1074#1077#1076#1080#1090#1077' '#1085#1086#1084#1077#1088' '#1079#1072#1082#1072#1079#1072
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Label6: TLabel
      Left = 320
      Top = 21
      Width = 170
      Height = 20
      Caption = #1057#1091#1084#1084#1072'  '#1079#1072#1082#1072#1079#1072'. ['#1088#1091#1073']'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Label4: TLabel
      Left = 384
      Top = 64
      Width = 7
      Height = 24
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -20
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Edit1: TEdit
      Left = 16
      Top = 52
      Width = 201
      Height = 24
      TabOrder = 0
      OnChange = Edit1Change
    end
    object Button4: TButton
      Left = 16
      Top = 88
      Width = 201
      Height = 33
      Caption = #1055#1086#1089#1095#1080#1090#1072#1090#1100
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 1
      OnClick = Button3Click
    end
  end
  object IdLPR1: TIdLPR
    MaxLineAction = maException
    ReadTimeout = 0
    Port = 515
    Queue = 'pr1'
    Left = 1248
    Top = 8
  end
  object WordDocument1: TWordDocument
    AutoConnect = False
    ConnectKind = ckRunningOrNew
    Left = 1144
    Top = 8
  end
  object WordApplication1: TWordApplication
    AutoConnect = False
    ConnectKind = ckRunningOrNew
    AutoQuit = False
    Left = 1184
    Top = 8
  end
  object WordParagraphFormat1: TWordParagraphFormat
    AutoConnect = False
    ConnectKind = ckRunningOrNew
    Left = 1216
    Top = 8
  end
  object SaveDialog1: TSaveDialog
    Left = 560
    Top = 72
  end
  object OpenDialog1: TOpenDialog
    Left = 196
    Top = 72
  end
  object OpenDialog2: TOpenDialog
    Left = 208
    Top = 240
  end
end
