object frmSaveDocDtls: TfrmSaveDocDtls
  Left = 0
  Top = 0
  BorderStyle = bsDialog
  Caption = 'Insight Save Document'
  ClientHeight = 413
  ClientWidth = 396
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = True
  Position = poMainFormCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object JvLabel1: TJvLabel
    Left = 12
    Top = 16
    Width = 34
    Height = 13
    Caption = 'Matter'
    Transparent = True
    HotTrackFont.Charset = DEFAULT_CHARSET
    HotTrackFont.Color = clWindowText
    HotTrackFont.Height = -11
    HotTrackFont.Name = 'Tahoma'
    HotTrackFont.Style = []
  end
  object JvLabel2: TJvLabel
    Left = 12
    Top = 43
    Width = 55
    Height = 13
    Caption = 'Description'
    Transparent = True
    HotTrackFont.Charset = DEFAULT_CHARSET
    HotTrackFont.Color = clWindowText
    HotTrackFont.Height = -11
    HotTrackFont.Name = 'Tahoma'
    HotTrackFont.Style = []
  end
  object JvLabel3: TJvLabel
    Left = 12
    Top = 71
    Width = 47
    Height = 13
    Caption = 'Category'
    Transparent = True
    HotTrackFont.Charset = DEFAULT_CHARSET
    HotTrackFont.Color = clWindowText
    HotTrackFont.Height = -11
    HotTrackFont.Name = 'Tahoma'
    HotTrackFont.Style = []
  end
  object JvLabel4: TJvLabel
    Left = 12
    Top = 98
    Width = 64
    Height = 13
    Caption = 'Classification'
    Transparent = True
    HotTrackFont.Charset = DEFAULT_CHARSET
    HotTrackFont.Color = clWindowText
    HotTrackFont.Height = -11
    HotTrackFont.Name = 'Tahoma'
    HotTrackFont.Style = []
  end
  object JvLabel5: TJvLabel
    Left = 12
    Top = 125
    Width = 49
    Height = 13
    Caption = 'Keywords'
    Transparent = True
    HotTrackFont.Charset = DEFAULT_CHARSET
    HotTrackFont.Color = clWindowText
    HotTrackFont.Height = -11
    HotTrackFont.Name = 'Tahoma'
    HotTrackFont.Style = []
  end
  object JvLabel6: TJvLabel
    Left = 12
    Top = 147
    Width = 86
    Height = 13
    Caption = 'Precedent Details'
    Transparent = True
    HotTrackFont.Charset = DEFAULT_CHARSET
    HotTrackFont.Color = clWindowText
    HotTrackFont.Height = -11
    HotTrackFont.Name = 'Tahoma'
    HotTrackFont.Style = []
  end
  object JvLabel7: TJvLabel
    Left = 12
    Top = 229
    Width = 35
    Height = 13
    Caption = 'Author'
    Transparent = True
    HotTrackFont.Charset = DEFAULT_CHARSET
    HotTrackFont.Color = clWindowText
    HotTrackFont.Height = -11
    HotTrackFont.Name = 'Tahoma'
    HotTrackFont.Style = []
  end
  object JvLabel8: TJvLabel
    Left = 12
    Top = 322
    Width = 61
    Height = 13
    Caption = 'File Location'
    Transparent = True
    HotTrackFont.Charset = DEFAULT_CHARSET
    HotTrackFont.Color = clWindowText
    HotTrackFont.Height = -11
    HotTrackFont.Name = 'Tahoma'
    HotTrackFont.Style = []
  end
  object TxtDocName: TLMDEdit
    Left = 98
    Top = 40
    Width = 288
    Height = 21
    Bevel.Mode = bmWindows
    Caret.BlinkRate = 530
    TabOrder = 0
    CustomButtons = <>
    PasswordChar = #0
  end
  object edKeywords: TJvEdit
    Left = 98
    Top = 121
    Width = 287
    Height = 21
    TabOrder = 1
  end
  object memoPrecDetails: TJvMemo
    Left = 98
    Top = 147
    Width = 288
    Height = 72
    ScrollBars = ssVertical
    TabOrder = 2
  end
  object cbLeaveDocOpen: TJvCheckBox
    Left = 12
    Top = 262
    Width = 182
    Height = 17
    Caption = 'Leave Document open after Save'
    Checked = True
    State = cbChecked
    TabOrder = 3
    LinkedControls = <>
    HotTrackFont.Charset = DEFAULT_CHARSET
    HotTrackFont.Color = clWindowText
    HotTrackFont.Height = -11
    HotTrackFont.Name = 'Tahoma'
    HotTrackFont.Style = []
  end
  object cbOverwriteDoc: TJvCheckBox
    Left = 12
    Top = 279
    Width = 161
    Height = 17
    Caption = 'Overwrite current document.'
    TabOrder = 4
    LinkedControls = <>
    HotTrackFont.Charset = DEFAULT_CHARSET
    HotTrackFont.Color = clWindowText
    HotTrackFont.Height = -11
    HotTrackFont.Name = 'Tahoma'
    HotTrackFont.Style = []
  end
  object cbPortalAccess: TJvCheckBox
    Left = 12
    Top = 296
    Width = 115
    Height = 17
    Caption = 'Client Portal Access'
    TabOrder = 5
    LinkedControls = <>
    HotTrackFont.Charset = DEFAULT_CHARSET
    HotTrackFont.Color = clWindowText
    HotTrackFont.Height = -11
    HotTrackFont.Name = 'Tahoma'
    HotTrackFont.Style = []
  end
  object btnTxtDocPath: TLMDBrowseEdit
    Left = 12
    Top = 338
    Width = 373
    Height = 21
    Bevel.Mode = bmWindows
    Caret.BlinkRate = 530
    TabOrder = 6
    Options = [doBrowseForComputer, doReturnFileSysDirs, doStatusText, doShowFiles, doShowPath]
    CustomButtons = <
      item
        Caption = '...'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        Index = 0
        DisplayName = 'TLMDSpecialButton'
        ImageIndex = 0
        ListIndex = 0
        UsePngGlyph = False
      end
      item
        Caption = '<<'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        Index = 1
        DisplayName = 'TLMDSpecialButton'
        ImageIndex = 0
        ListIndex = 0
        UsePngGlyph = False
      end>
    CustomButtonWidth = 18
  end
  object JvHTButton1: TJvHTButton
    Left = 80
    Top = 366
    Width = 75
    Height = 25
    Caption = 'Save'
    TabOrder = 7
    OnClick = JvHTButton1Click
  end
  object JvHTButton2: TJvHTButton
    Left = 224
    Top = 366
    Width = 75
    Height = 25
    Caption = 'Close'
    ModalResult = 2
    TabOrder = 8
    OnClick = JvHTButton2Click
  end
  object cmbCategory: TDBLookupComboBox
    Left = 98
    Top = 67
    Width = 287
    Height = 21
    KeyField = 'NPRECCATEGORY'
    ListField = 'DESCR'
    TabOrder = 9
  end
  object cmbClassification: TDBLookupComboBox
    Left = 98
    Top = 94
    Width = 287
    Height = 21
    DropDownRows = 10
    KeyField = 'NPRECCLASSIFICATION'
    ListField = 'DESCR'
    TabOrder = 10
  end
  object cmbAuthor: TDBLookupComboBox
    Left = 98
    Top = 226
    Width = 167
    Height = 21
    DropDownRows = 15
    KeyField = 'CODE'
    ListField = 'NAME'
    TabOrder = 11
  end
  object cbNewCopy: TJvCheckBox
    Left = 273
    Top = 14
    Width = 106
    Height = 17
    Caption = 'Create New Copy'
    Checked = True
    State = cbChecked
    TabOrder = 12
    LinkedControls = <>
    HotTrackFont.Charset = DEFAULT_CHARSET
    HotTrackFont.Color = clWindowText
    HotTrackFont.Height = -11
    HotTrackFont.Name = 'Tahoma'
    HotTrackFont.Style = []
  end
  object LMDDockButton1: TLMDDockButton
    Left = 233
    Top = 12
    Width = 22
    Height = 21
    TabOrder = 13
    OnClick = LMDDockButton1Click
    Glyph.Data = {
      E6000000424DE60000000000000076000000280000000D0000000E0000000100
      0400000000007000000000000000000000001000000010000000000000000000
      8000008000000080800080000000800080008080000080808000C0C0C0000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
      3000333333333333300033333333333330003333333333333000333333333333
      3000333333333333300033003003003330003300300300333000333333333333
      3000333333333333300033333333333330003333333333333000333333333333
      30003333333333333000}
    Control = btnEditMatter
  end
  object btnEditMatter: TJvEdit
    Left = 98
    Top = 12
    Width = 134
    Height = 21
    TabOrder = 14
  end
  object StatusBar: TStatusBar
    Left = 0
    Top = 394
    Width = 396
    Height = 19
    Panels = <
      item
        Width = 50
      end>
  end
end
