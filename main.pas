unit main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxGraphics, cxControls, cxLookAndFeels, cxLookAndFeelPainters,
  cxStyles, cxCustomData, cxFilter, cxData, cxDataStorage, cxEdit, DB,
  cxDBData, Menus, cxContainer, cxTextEdit, cxButtons, cxGridLevel,
  cxClasses, cxGridCustomView, cxGridCustomTableView, cxGridTableView,
  cxGridDBTableView, cxGrid, StdCtrls, Mask, JvExMask, JvToolEdit,
  ExtCtrls, cxPC, JvMemoryDataset, Registry, ComObj;

type
  TfrmMain = class(TForm)
    cxPageControl1: TcxPageControl;
    cxTabSheet1: TcxTabSheet;
    cxTabSheet2: TcxTabSheet;
    cxTabSheet3: TcxTabSheet;
    Panel1: TPanel;
    Label6: TLabel;
    lblPRoject: TLabel;
    Label2: TLabel;
    Label1: TLabel;
    dirSettings: TJvDirectoryEdit;
    Edit1: TEdit;
    Label3: TLabel;
    Button1: TButton;
    Edit2: TEdit;
    Label4: TLabel;
    Label5: TLabel;
    Edit3: TEdit;
    Button2: TButton;
    memoSubject: TMemo;
    cbShowBuiltInProps: TCheckBox;
    tvProjects: TcxGridDBTableView;
    grdProjectsLevel1: TcxGridLevel;
    grdProjects: TcxGrid;
    cxGrid2DBTableView1: TcxGridDBTableView;
    cxGrid2Level1: TcxGridLevel;
    cxGrid2: TcxGrid;
    cxGrid3DBTableView1: TcxGridDBTableView;
    cxGrid3Level1: TcxGridLevel;
    cxGrid3: TcxGrid;
    btnReadDocData: TcxButton;
    btnReadProjData: TcxButton;
    edProjectCode: TcxTextEdit;
    btnCreateProject: TcxButton;
    MemoryData: TJvMemoryData;
    dsMemoryData: TDataSource;
    memDataProject: TJvMemoryData;
    dsDataProject: TDataSource;
    cxGrid3DBTableView1Column1: TcxGridDBColumn;
    cxGrid3DBTableView1Column2: TcxGridDBColumn;
    tvProjectsClientName: TcxGridDBColumn;
    tvProjectsProjectCode: TcxGridDBColumn;
    tvProjectsProjectDescr: TcxGridDBColumn;
    tvProjectsDateEdited: TcxGridDBColumn;
    btnSaveData: TcxButton;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure btnCreateProjectClick(Sender: TObject);
    procedure tvProjectsCellClick(Sender: TcxCustomGridTableView;
      ACellViewInfo: TcxGridTableDataCellViewInfo; AButton: TMouseButton;
      AShift: TShiftState; var AHandled: Boolean);
    procedure btnReadDocDataClick(Sender: TObject);
    procedure btnSaveDataClick(Sender: TObject);
    procedure btnReadProjDataClick(Sender: TObject);
  private
    { Private declarations }
    FValue: string;
    FProject: string;
    FProjectPath: string;

    procedure LoadProjectTable;
    procedure ReadDocProperties;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.dfm}

uses
  functions;

procedure TfrmMain.FormCreate(Sender: TObject);
begin
  memDataProject.Open;
  MemoryData.Open;
end;

procedure TfrmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  memDataProject.Close;
  MemoryData.Close;
end;

procedure TfrmMain.FormShow(Sender: TObject);
var
   reg:  TRegistry;
begin
   reg := TRegistry.Create(KEY_READ);
   reg.RootKey := HKEY_CURRENT_USER;
   if (reg.KeyExists('Software\\ProDocTools\\SpeediDocs\\')) then
   begin
      reg.OpenKey('Software\\ProDocTools\\SpeediDocs\\', False);
      FValue := reg.ReadString('FileDir');
      dirSettings.Text := FValue;
      dirSettings.Directory := FValue;
   end;
   reg.CloseKey();
   reg.Free;
   LoadProjectTable;
end;

procedure TfrmMain.btnCreateProjectClick(Sender: TObject);
var
  AFile: file;
begin
ShowMessage(FValue);
   if FValue <> '' then
   begin
      if FileExists(FValue + '\\' + edProjectCode.Text + '.gdoc') then
        ShowMessage(edProjectCode.Text + ' already exists!')
      else
      begin
        memDataProject.Append;
        memDataProject.Fields.Fields[1].Text := edProjectCode.Text;
        FProjectPath := FValue + '\' + edProjectCode.Text + '.gdoc';
        FileCreate(FProjectPath);
      end;
   end;
end;

procedure TfrmMain.tvProjectsCellClick(Sender: TcxCustomGridTableView;
  ACellViewInfo: TcxGridTableDataCellViewInfo; AButton: TMouseButton;
  AShift: TShiftState; var AHandled: Boolean);
begin
   lblProject.Caption := memDataProject.Fields.Fields[1].Text;
   FProject :=  memDataProject.Fields.Fields[1].Text;
end;

procedure TfrmMain.btnReadDocDataClick(Sender: TObject);
begin
  ReadDocProperties;
end;

procedure TfrmMain.ReadDocProperties;
var
  varWord, varDoc, item: OLEVariant;
  CustPropCount, PropCount, i: integer;
begin
   varWord := GetActiveOleObject('Word.Application');
//   ShowMessage('varword');
   varDoc := varWord.ActiveDocument;
//   ShowMessage('vardoc');
   try
     CustPropCount := varDoc.CustomDocumentProperties.Count;
//     ShowMessage(Inttostr(CustPropCount));
     PropCount := varDoc.BuiltInDocumentProperties.Count;
//     ShowMessage(Inttostr(CustPropCount) + ' custom props' + chr(13) + Inttostr(PropCount) + '  props');

     if cbShowBuiltInProps.Checked then
     begin
       for i := 1 to PropCount do
       begin
         MemoryData.Append;
         item := varDoc.BuiltInDocumentProperties[i];
         try
            MemoryData.Fields.Fields[1].AsString := Item.Name;
//            MemoryData.Fields.Fields[2].AsString := Item.Type;
            MemoryData.Fields.Fields[2].AsString := Item.Value;
         except
            //
         end;
       end;
     end;

     if (CustPropCount > 0) then
     begin
       for i := 1 to CustPropCount do
       begin
         MemoryData.Append;
         item := varDoc.CustomDocumentProperties[i];
         try
            MemoryData.Fields.Fields[1].AsString := Item.Name;
            MemoryData.Fields.Fields[2].AsString := Item.Value;
         except
            //
         end;
       end;
     end;
   except
     //
   end;
end;

procedure TfrmMain.btnSaveDataClick(Sender: TObject);
begin
  Data2XML(memDataProject, FProjectPath);
end;

procedure TfrmMain.btnReadProjDataClick(Sender: TObject);
begin
  XML2Data
end;

procedure TfrmMain.LoadProjectTable;
var
   Res: TSearchRec;
   EOFound: Boolean;
   AFileName: string;
begin
   EOFound:= False;
   if FindFirst(FValue+'\*.gdoc', faAnyFile - faDirectory, Res) < 0 then
     exit
   else
     while not EOFound do
     begin
        memDataProject.Append;
        AFileName := ExtractFileName(Res.Name);
//        MessageDlg(AFileName,mtWarning,mbYesNo,0);
        memDataProject.Fields.Fields[1].Text := Copy(AFileName,1, length(AFileName)- 5);
        EOFound:= FindNext(Res) <> 0;
     end;
   FindClose(Res) ;
end;

end.
