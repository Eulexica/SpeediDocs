object frmMain: TfrmMain
  Left = 445
  Top = 197
  BorderStyle = bsDialog
  Caption = 'SpeediDocs'
  ClientHeight = 370
  ClientWidth = 560
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  DesignSize = (
    560
    370)
  PixelsPerInch = 96
  TextHeight = 13
  object cxPageControl1: TcxPageControl
    Left = 5
    Top = 7
    Width = 548
    Height = 313
    ActivePage = cxTabSheet1
    Anchors = [akLeft, akTop, akRight, akBottom]
    ShowFrame = True
    Style = 9
    TabOrder = 0
    ClientRectBottom = 312
    ClientRectLeft = 1
    ClientRectRight = 547
    ClientRectTop = 20
    object cxTabSheet1: TcxTabSheet
      Caption = '&Data'
      ImageIndex = 0
      object cxGrid2: TcxGrid
        Left = 2
        Top = 28
        Width = 260
        Height = 253
        TabOrder = 0
        LookAndFeel.NativeStyle = True
        object cxGrid2DBTableView1: TcxGridDBTableView
          NavigatorButtons.ConfirmDelete = False
          DataController.Summary.DefaultGroupSummaryItems = <>
          DataController.Summary.FooterSummaryItems = <>
          DataController.Summary.SummaryGroups = <>
          OptionsView.GroupByBox = False
        end
        object cxGrid2Level1: TcxGridLevel
          GridView = cxGrid2DBTableView1
        end
      end
      object cxGrid3: TcxGrid
        Left = 270
        Top = 28
        Width = 260
        Height = 253
        TabOrder = 1
        LookAndFeel.NativeStyle = True
        object cxGrid3DBTableView1: TcxGridDBTableView
          NavigatorButtons.ConfirmDelete = False
          DataController.DataSource = dsMemoryData
          DataController.Summary.DefaultGroupSummaryItems = <>
          DataController.Summary.FooterSummaryItems = <>
          DataController.Summary.SummaryGroups = <>
          OptionsCustomize.ColumnFiltering = False
          OptionsView.ColumnAutoWidth = True
          OptionsView.GroupByBox = False
          OptionsView.Indicator = True
          object cxGrid3DBTableView1Column1: TcxGridDBColumn
            DataBinding.FieldName = 'FieldName'
          end
          object cxGrid3DBTableView1Column2: TcxGridDBColumn
            DataBinding.FieldName = 'FieldValue'
          end
        end
        object cxGrid3Level1: TcxGridLevel
          GridView = cxGrid3DBTableView1
        end
      end
      object btnReadDocData: TcxButton
        Left = 2
        Top = 2
        Width = 126
        Height = 23
        Caption = 'Read Document Data'
        TabOrder = 2
        OnClick = btnReadDocDataClick
        LookAndFeel.NativeStyle = True
      end
      object btnReadProjData: TcxButton
        Left = 270
        Top = 2
        Width = 126
        Height = 23
        Caption = 'Read Project Data'
        TabOrder = 3
        OnClick = btnReadProjDataClick
        LookAndFeel.NativeStyle = True
      end
    end
    object cxTabSheet2: TcxTabSheet
      Caption = '&Projects'
      ImageIndex = 1
      object Label2: TLabel
        Left = 19
        Top = 8
        Width = 61
        Height = 13
        Caption = 'Project Code'
      end
      object grdProjects: TcxGrid
        Left = 3
        Top = 32
        Width = 543
        Height = 248
        TabOrder = 0
        LookAndFeel.NativeStyle = True
        object tvProjects: TcxGridDBTableView
          NavigatorButtons.ConfirmDelete = False
          OnCellClick = tvProjectsCellClick
          DataController.DataSource = dsDataProject
          DataController.Summary.DefaultGroupSummaryItems = <>
          DataController.Summary.FooterSummaryItems = <>
          DataController.Summary.SummaryGroups = <>
          OptionsCustomize.ColumnFiltering = False
          OptionsView.ColumnAutoWidth = True
          OptionsView.GroupByBox = False
          OptionsView.Indicator = True
          object tvProjectsClientName: TcxGridDBColumn
            Caption = 'Client Name'
            DataBinding.FieldName = 'ClientName'
          end
          object tvProjectsProjectCode: TcxGridDBColumn
            Caption = 'Project Code'
            DataBinding.FieldName = 'ProjectCode'
            Options.Editing = False
            Options.Focusing = False
          end
          object tvProjectsProjectDescr: TcxGridDBColumn
            Caption = 'Project Description'
            DataBinding.FieldName = 'ProjectDescr'
          end
          object tvProjectsDateEdited: TcxGridDBColumn
            Caption = 'Date Last Edited'
            DataBinding.FieldName = 'DateEdited'
          end
        end
        object grdProjectsLevel1: TcxGridLevel
          GridView = tvProjects
        end
      end
      object edProjectCode: TcxTextEdit
        Left = 86
        Top = 5
        Style.LookAndFeel.NativeStyle = True
        StyleDisabled.LookAndFeel.NativeStyle = True
        StyleFocused.LookAndFeel.NativeStyle = True
        StyleHot.LookAndFeel.NativeStyle = True
        TabOrder = 1
        Width = 152
      end
      object btnCreateProject: TcxButton
        Left = 249
        Top = 4
        Width = 75
        Height = 23
        Caption = 'Create'
        TabOrder = 2
        OnClick = btnCreateProjectClick
        LookAndFeel.NativeStyle = True
      end
    end
    object cxTabSheet3: TcxTabSheet
      Caption = 'S&ettings'
      ImageIndex = 2
      object Label1: TLabel
        Left = 24
        Top = 15
        Width = 60
        Height = 13
        Caption = 'File Location'
        Transparent = True
      end
      object Label3: TLabel
        Left = 24
        Top = 48
        Width = 50
        Height = 13
        Caption = 'Your Email'
      end
      object Label4: TLabel
        Left = 24
        Top = 82
        Width = 84
        Height = 13
        Caption = 'Registration Code'
      end
      object Label5: TLabel
        Left = 24
        Top = 112
        Width = 55
        Height = 13
        Caption = 'Tell a friend'
      end
      object dirSettings: TJvDirectoryEdit
        Left = 119
        Top = 12
        Width = 289
        Height = 21
        DialogKind = dkWin32
        DialogText = 'GeniDocs Default Directory'
        TabOrder = 0
      end
      object Edit1: TEdit
        Left = 119
        Top = 45
        Width = 289
        Height = 21
        TabOrder = 1
      end
      object Button1: TButton
        Left = 414
        Top = 43
        Width = 75
        Height = 25
        Caption = 'Register'
        TabOrder = 2
      end
      object Edit2: TEdit
        Left = 119
        Top = 77
        Width = 121
        Height = 21
        TabOrder = 3
      end
      object Edit3: TEdit
        Left = 119
        Top = 109
        Width = 289
        Height = 21
        TabOrder = 4
      end
      object Button2: TButton
        Left = 414
        Top = 107
        Width = 75
        Height = 25
        Caption = 'Send email'
        TabOrder = 5
      end
      object memoSubject: TMemo
        Left = 119
        Top = 136
        Width = 289
        Height = 79
        Lines.Strings = (
          'Hi, I'#39've been automating some of my documents with '
          'GeniDocs.  You might like to try it.  You will get 1 month '
          'free trial and I will get an extra month free.')
        TabOrder = 6
      end
      object cbShowBuiltInProps: TCheckBox
        Left = 16
        Top = 224
        Width = 137
        Height = 17
        Caption = 'Show Built in Properties'
        Color = clBtnFace
        ParentColor = False
        TabOrder = 7
      end
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 350
    Width = 560
    Height = 20
    Align = alBottom
    TabOrder = 1
    object Label6: TLabel
      Left = 8
      Top = 3
      Width = 90
      Height = 13
      Caption = 'Current Project:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object lblPRoject: TLabel
      Left = 104
      Top = 3
      Width = 3
      Height = 13
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
    end
  end
  object btnSaveData: TcxButton
    Left = 276
    Top = 322
    Width = 98
    Height = 24
    Anchors = [akRight, akBottom]
    Caption = '&Save Data'
    TabOrder = 2
    OnClick = btnSaveDataClick
    LookAndFeel.NativeStyle = True
  end
  object MemoryData: TJvMemoryData
    FieldDefs = <
      item
        Name = 'FieldSeq'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'FieldName'
        DataType = ftString
        Size = 254
      end
      item
        Name = 'FieldType'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'FieldValue'
        DataType = ftString
        Size = 254
      end>
    Left = 40
    Top = 276
  end
  object dsMemoryData: TDataSource
    DataSet = MemoryData
    Left = 120
    Top = 276
  end
  object memDataProject: TJvMemoryData
    Active = True
    FieldDefs = <
      item
        Name = 'ClientName'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'ProjectCode'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'ProjectDescr'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'DateEdited'
        DataType = ftDateTime
      end>
    Left = 432
    Top = 284
  end
  object dsDataProject: TDataSource
    DataSet = memDataProject
    Left = 520
    Top = 284
  end
end
