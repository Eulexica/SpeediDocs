object frmAdditionalInfo: TfrmAdditionalInfo
  Left = 0
  Top = 0
  BorderStyle = bsDialog
  Caption = 'Additional Info'
  ClientHeight = 248
  ClientWidth = 547
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Segoe UI'
  Font.Style = []
  OldCreateOrder = False
  Position = poOwnerFormCenter
  OnShow = FormShow
  DesignSize = (
    547
    248)
  PixelsPerInch = 96
  TextHeight = 13
  object BitBtn1: TBitBtn
    Left = 462
    Top = 216
    Width = 75
    Height = 25
    Anchors = [akRight, akBottom]
    Caption = 'Close'
    TabOrder = 0
    OnClick = BitBtn1Click
  end
  object Memo1: TMemo
    Left = 8
    Top = 8
    Width = 418
    Height = 42
    TabOrder = 1
  end
  object GroupBox1: TGroupBox
    Left = 8
    Top = 56
    Width = 529
    Height = 149
    Caption = 'Outlook'
    TabOrder = 2
    object txtRegResiliency: TLabel
      Left = 10
      Top = 21
      Width = 415
      Height = 32
      AutoSize = False
      WordWrap = True
    end
    object txtAlwaysLoad: TLabel
      Left = 10
      Top = 69
      Width = 415
      Height = 32
      AutoSize = False
      WordWrap = True
    end
    object Label1: TLabel
      Left = 10
      Top = 112
      Width = 423
      Height = 29
      AutoSize = False
      Caption = 
        'If you set the registry keys by clicking on either of the above ' +
        'buttons you will need to close Outlook.'
      WordWrap = True
    end
    object btnResilience: TButton
      Left = 439
      Top = 19
      Width = 75
      Height = 24
      Caption = 'Add Key'
      TabOrder = 0
      Visible = False
      OnClick = btnResilienceClick
    end
    object btnAlwaysLoad: TButton
      Left = 439
      Top = 67
      Width = 75
      Height = 24
      Caption = 'Add Key'
      TabOrder = 1
      Visible = False
      OnClick = btnAlwaysLoadClick
    end
  end
end
