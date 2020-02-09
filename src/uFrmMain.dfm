object frmMain: TfrmMain
  Left = 0
  Top = 0
  Caption = 'ExcelCompare'
  ClientHeight = 255
  ClientWidth = 363
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object edtFirstFile: TButtonedEdit
    AlignWithMargins = True
    Left = 3
    Top = 3
    Width = 357
    Height = 21
    Align = alTop
    TabOrder = 0
    TextHint = 'Path to the first file to comape'
  end
  object edtSecondFile: TButtonedEdit
    AlignWithMargins = True
    Left = 3
    Top = 30
    Width = 357
    Height = 21
    Align = alTop
    TabOrder = 1
    TextHint = 'Path to the second file to comape'
  end
  object btnCompare: TButton
    AlignWithMargins = True
    Left = 3
    Top = 227
    Width = 357
    Height = 25
    Align = alBottom
    Caption = 'Compare'
    TabOrder = 2
    OnClick = btnCompareClick
  end
  object mmResult: TMemo
    AlignWithMargins = True
    Left = 3
    Top = 57
    Width = 357
    Height = 164
    Align = alClient
    Color = clMenu
    ScrollBars = ssVertical
    TabOrder = 3
  end
end
