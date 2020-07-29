object frmFieldList: TfrmFieldList
  Left = 0
  Top = 0
  BorderStyle = bsDialog
  Caption = 'Merge Field List'
  ClientHeight = 639
  ClientWidth = 470
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'Segoe UI'
  Font.Style = []
  FormStyle = fsStayOnTop
  OldCreateOrder = False
  Position = poOwnerFormCenter
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 106
  TextHeight = 17
  object cxGrid1: TcxGrid
    Left = 0
    Top = 0
    Width = 470
    Height = 598
    Align = alClient
    TabOrder = 0
    object tvMergeFields: TcxGridDBTableView
      OnDblClick = tvMergeFieldsDblClick
      Navigator.Buttons.CustomButtons = <>
      DataController.DataSource = dsTranslate
      DataController.Summary.DefaultGroupSummaryItems = <>
      DataController.Summary.FooterSummaryItems = <>
      DataController.Summary.SummaryGroups = <>
      FilterRow.SeparatorWidth = 7
      FixedDataRows.SeparatorWidth = 7
      NewItemRow.SeparatorWidth = 7
      OptionsBehavior.IncSearch = True
      OptionsBehavior.IncSearchItem = tvMergeFieldsEXTERNALFIELD
      OptionsBehavior.PullFocusing = True
      OptionsData.Deleting = False
      OptionsData.DeletingConfirmation = False
      OptionsData.Editing = False
      OptionsData.Inserting = False
      OptionsSelection.CellSelect = False
      OptionsView.NavigatorOffset = 57
      OptionsView.CellAutoHeight = True
      OptionsView.ColumnAutoWidth = True
      OptionsView.GroupByBox = False
      OptionsView.Indicator = True
      OptionsView.IndicatorWidth = 14
      Preview.LeftIndent = 23
      Preview.RightIndent = 6
      object tvMergeFieldsEXTERNALFIELD: TcxGridDBColumn
        Caption = 'Field'
        DataBinding.FieldName = 'EXTERNALFIELD'
        MinWidth = 23
        Width = 179
      end
      object tvMergeFieldsDESCR: TcxGridDBColumn
        Caption = 'Description'
        DataBinding.FieldName = 'DESCR'
        MinWidth = 23
        Width = 87
      end
      object tvMergeFieldsSAMPLE_DATA: TcxGridDBColumn
        Caption = 'Sample Data'
        DataBinding.FieldName = 'SAMPLE_DATA'
        MinWidth = 23
        Width = 161
      end
    end
    object cxGrid1Level1: TcxGridLevel
      GridView = tvMergeFields
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 598
    Width = 470
    Height = 41
    Align = alBottom
    TabOrder = 1
    object cxButton1: TcxButton
      Left = 372
      Top = 6
      Width = 83
      Height = 28
      Caption = 'Close'
      TabOrder = 0
      OnClick = cxButton1Click
    end
  end
  object TBTranslate: TOraTable
    TableName = 'WORKFLOWFIELDTRANSLATE'
    Session = dmSaveDoc.orsInsight
    Left = 115
    Top = 299
  end
  object dsTranslate: TOraDataSource
    DataSet = TBTranslate
    Left = 110
    Top = 355
  end
end
