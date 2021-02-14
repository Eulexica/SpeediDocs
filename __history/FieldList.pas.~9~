unit FieldList;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, cxStyles, cxCustomData, cxFilter, cxData,
  cxDataStorage, cxEdit, cxNavigator, Data.DB, cxDBData, DBAccess,
  cxGridLevel, cxClasses, cxGridCustomView, cxGridCustomTableView,
  cxGridTableView, cxGridDBTableView, cxGrid, Ora, MemDS, OraSmart, Vcl.ExtCtrls,
  Word2000, cxDataControllerConditionalFormattingRulesManagerDialog, Vcl.Menus,
  Vcl.StdCtrls, cxButtons, savedoc;

type
  TfrmFieldList = class(TForm)
    tvMergeFields: TcxGridDBTableView;
    cxGrid1Level1: TcxGridLevel;
    cxGrid1: TcxGrid;
    TBTranslate: TOraTable;
    dsTranslate: TOraDataSource;
    tvMergeFieldsEXTERNALFIELD: TcxGridDBColumn;
    tvMergeFieldsDESCR: TcxGridDBColumn;
    tvMergeFieldsSAMPLE_DATA: TcxGridDBColumn;
    Panel1: TPanel;
    cxButton1: TcxButton;
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure tvMergeFieldsDblClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure cxButton1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure SetWordApp(WordApp: TWordApplication);
    function GetWordApp(): TWordApplication;
  end;

var
  frmFieldList: TfrmFieldList;

implementation


{$R *.dfm}

uses
   savedocfunc;

var
   MSWord: TWordApplication;

procedure TfrmFieldList.tvMergeFieldsDblClick(Sender: TObject);
var
   Doc: WordDocument;
   OleVar2: oleVariant;
begin
//   MSWord := GetWordApp;
   Doc := MSWord.ActiveDocument;
   OleVar2 := String(tvMergeFieldsEXTERNALFIELD.EditValue);
   Doc.MailMerge.Fields.Add(MSWord.Selection.Range, OleVar2);
end;

procedure TfrmFieldList.cxButton1Click(Sender: TObject);
begin
   Close;
end;

procedure TfrmFieldList.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
var
   OleVar: OleVariant;
   x: WordBool;
begin
   try
      TBTranslate.Close;
      OleVar := False;
//      MSWord.ActiveDocument.Close(OleVar,EmptyParam, EmptyParam);
   finally
//      dmSaveDoc.Free;
      x := False;
      MSWord.ActiveWindow.View.ShowFieldCodes := x;
//      MSWord := nil;
   end;
end;

procedure TfrmFieldList.FormCreate(Sender: TObject);
begin
//   if (not Assigned(dmSaveDoc)) then
//      dmSaveDoc := TdmSaveDoc.Create(Application);
//   if dmSaveDoc.orsInsight.Connected = False then dmSaveDoc.GetUserID;

end;

procedure TfrmFieldList.FormShow(Sender: TObject);
var
   x: WordBool;
begin
   try
//      if dmSaveDoc.orsInsight.Connected = False then dmSaveDoc.GetUserID;
      TBTranslate.Open;
      x := True;
      MSWord := GetWordApp;
      MSWord.ActiveWindow.View.ShowFieldCodes := x;
   finally
//
   end;
end;

procedure TfrmFieldList.SetWordApp(WordApp: TWordApplication);
begin
   MSWord := WordApp;
end;

function TfrmFieldList.GetWordApp(): TWordApplication;
begin
   Result := MSWord;
end;

end.
