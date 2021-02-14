unit SaveprecDetails;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComObj, StdCtrls, Menus, DB, ActnList,
  ActnMan, Vcl.ExtCtrls, Vcl.ImgList,
  Vcl.Buttons, Vcl.ComCtrls, Vcl.DBCtrls,
  SaveDoc, cxGraphics, cxControls, cxLookAndFeels, cxLookAndFeelPainters,
  cxContainer, cxEdit, cxTextEdit, cxMaskEdit, cxDropDownEdit, cxLookupEdit,
  cxDBLookupEdit, cxDBLookupComboBox, Outlook2000, cxImageComboBox, DBAccess,
  Ora, MemDS, JvBaseDlg, JvBrowseFolder, cxButtonEdit, dxLayoutControlAdapters,
  dxLayoutContainer, dxLayoutcxEditAdapters, cxClasses, dxLayoutControl,
  System.ImageList;

const
     CUSTOMPROPS: array[0..10] of string = ('MatterNo','DocID','Prec_Category','Prec_Classification','Doc_Keywords','Doc_Precedent','Doc_FileName','Doc_Author','Saved_in_DB', 'Doc_Title','Portal_Access');

type
  TfrmSavePrecDetails = class(TForm)
    btnSave: TBitBtn;
    btnClose: TBitBtn;
    ImageList1: TImageList;
    cbLeaveDocOpen: TCheckBox;
    edKeywords: TEdit;
    txtDocName: TEdit;
    StatusBar: TStatusBar;
    cmbPrecCategory: TcxLookupComboBox;
    cmbClassification: TcxLookupComboBox;
    cmbAuthor: TcxLookupComboBox;
    memoPrecDetails: TMemo;
    dblGroup: TcxImageComboBox;
    qryEmployee: TOraQuery;
    dsEmployee: TOraDataSource;
    qryPrecCategory: TOraQuery;
    dsPrecCategory: TOraDataSource;
    qryPrecClassification: TOraQuery;
    dsPrecClassification: TOraDataSource;
    BrowseDlg: TJvBrowseForFolderDialog;
    btnTxtDocPath: TcxButtonEdit;
    dxLayoutControl1Group_Root: TdxLayoutGroup;
    dxLayoutControl1: TdxLayoutControl;
    dxLayoutGroup1: TdxLayoutGroup;
    dxLayoutGroup2: TdxLayoutGroup;
    dxLayoutGroup3: TdxLayoutGroup;
    dxLayoutGroup4: TdxLayoutGroup;
    dxLayoutItem2: TdxLayoutItem;
    dxLayoutItem3: TdxLayoutItem;
    dxLayoutItem4: TdxLayoutItem;
    dxLayoutItem5: TdxLayoutItem;
    dxLayoutItem6: TdxLayoutItem;
    dxLayoutItem7: TdxLayoutItem;
    dxLayoutItem8: TdxLayoutItem;
    dxLayoutItem9: TdxLayoutItem;
    dxLayoutItem10: TdxLayoutItem;
    dxLayoutGroup5: TdxLayoutGroup;
    dxLayoutItem11: TdxLayoutItem;
    dxLayoutItem12: TdxLayoutItem;
    cmbTemplateType: TcxComboBox;
    dxLayoutItem13: TdxLayoutItem;
    cmbWorkflowType: TcxLookupComboBox;
    dxLayoutItem1: TdxLayoutItem;
    qryWorkflowType: TOraQuery;
    dsWorkflowType: TOraDataSource;
    procedure btnCloseClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnEditMatterPropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnTxtDocPathPropertiesButtonClick(Sender: TObject;
      AButtonIndex: Integer);
    procedure FormShow(Sender: TObject);
    procedure cbSaveAsPrecedentClick(Sender: TObject);
    procedure btnTxtDocPathPropertiesChange(Sender: TObject);
  private
    { Private declarations }
    nMatter: integer;

    tmpFileName: string;
    FPrec_Category: string;
    tmpdir: string;
    FSavedInDB: string;
    FDocName: string;
    FPrec_Classification: string;
    FDoc_Keywords: string;
    FDoc_Precedent: string;
    FDoc_FileName: string;
    FDoc_Author: string;
    FEditing: boolean;
    FAppType: integer;
    FFileID: string;
    FOldFileID: string;
    FMailSubject: string;
    FReceivedDate: TDateTime;
    FadxLCID: integer;
    FromWord: boolean;
    FIMail: MailItem;
//    procedure GetDetails;
  public
    { Public declarations }
    AWordProps: array[0..10] of TWordProperties;
    property DocName: string read FDocName;
    property AppType: Integer read FAppType write FAppType;
    property MailSubject: string read FMailSubject write FMailSubject;
    property ReceivedDate: TDateTime read FReceivedDate write FReceivedDate;
    property LadxLCID: integer read FadxLCID write FadxLCID;
    property IMail: _MailItem read FIMail write FIMail;
    procedure GetDetails;
  end;

var
  frmSavePrecDetails: TfrmSavePrecDetails;

function ShowDocSave: Integer; StdCall;

implementation

uses
   MatterSearch, SaveDocFunc, ActiveX, WordUnit, OutlookUnit, ExcelUnit,
   PowerPointUnit, SpeediDocs_IMPL, Office2000, SavedocDetails;

{$R *.dfm}

function ShowDocSave:integer;
var
   frmSavePrecDetails: TfrmSavePrecDetails;
begin
//   Application.Handle := AHandle;
   frmSavePrecDetails := TfrmSavePrecDetails.Create(Application);
   try
      frmSavePrecDetails.ShowModal;
      Result := frmSavePrecDetails.nMatter;
   finally
      frmSavePrecDetails.Free;
   end;
end;

procedure TfrmSavePrecDetails.btnCloseClick(Sender: TObject);
var
  Unknown: IUnknown;
  OLEResult: HResult;
  AMacro : string;
begin
   Close;
end;

procedure TfrmSavePrecDetails.FormCreate(Sender: TObject);
begin
{   try
      dmSaveDoc.qryPrecClassification.Open;
      dmSaveDoc.qryEmployee.Open;
      dmSaveDoc.qryPrecCategory.Open;
      dmSaveDoc.tbDocGroups.Open;

      cmbAuthor.EditValue := dmSaveDoc.UserCode;
      StatusBar.Panels[0].Text := 'Ver: '+ ReportVersion(SysUtils.GetModuleName(HInstance)) + ' (' +DateTimeToStr(FileDateToDateTime(FileAge(SysUtils.GetModuleName(HInstance))))+')';
   except
      Exit;
   end;   }
end;

procedure TfrmSavePrecDetails.FormShow(Sender: TObject);
var
   AItem: TcxImageComboboxItem;
begin
   try
      qryPrecClassification.Open;
      qryEmployee.Open;
      qryPrecCategory.Open;
      qryWorkFlowType.Open;
      dmConnection.tbDocGroups.Open;

      cmbAuthor.EditValue := dmConnection.UserCode;
      StatusBar.Panels[0].Text := 'Ver: '+ ReportVersion(SysUtils.GetModuleName(HInstance)) + ' (' +DateTimeToStr(FileDateToDateTime(FileAge(SysUtils.GetModuleName(HInstance))))+')';
   except
      Exit;
   end;

   dblGroup.Properties.Items.Clear;
   dmConnection.tbDocGroups.First();
   while(not dmConnection.tbDocGroups.Eof) do
   begin
      AItem := dblGroup.Properties.Items.Add;
      AItem.Description := dmConnection.tbDocGroups.FieldByName('NAME').AsString;
      AItem.Value := dmConnection.tbDocGroups.FieldByName('GROUPID').AsString;

      dmConnection.tbDocGroups.Next();
   end;

   FromWord := False;
   case AppType of
      1: ;
      2: begin
            FromWord := True;
            GetDetails;
         end;
      3: begin
            cbLeaveDocOpen.Checked := False;
            txtDocName.Text := MailSubject;
         end;
      4: ;
   end;

   btnTxtDocPath.Text := SystemString('DFLT_PRECEDENT_PATH');
end;

procedure TfrmSavePrecDetails.btnSaveClick(Sender: TObject);
var
   DocSequence: string;
//   bUsePath: boolean;
   cxTemplateTypeValue: string;
   cxWorkflowTypeValue: string;
   memoPrecDetailsValue: string;
   dblGroupKeyValue,
   cmbPrecCategoryKeyValue,
   cmbClassificationKeyValue: integer;
begin
   if btnTxtDocPath.Text <> '' then
   begin
      try
         if cmbAuthor.Text = '' then
         begin
            with Application do
            begin
               NormalizeTopMosts;
               MessageBox('Please enter an Author.','SpeediDocs',MB_OK+MB_ICONEXCLAMATION);
               RestoreTopMosts;
               exit;
            end;
         end;
         dmConnection.orsInsight.StartTransaction;
         dmConnection.qryDoctemplate.Open;

         FEditing := False;
//         bUsePath := False;
         tmpdir := IncludeTrailingPathDelimiter(GetEnvironmentVariable('TMP'));

         if btnTxtDocPath.Text = '' then
            tmpFileName := txtDocName.Text
         else
            tmpFileName := btnTxtDocPath.Text;

         try
            if cmbWorkFlowType.Text = '' then
                cxWorkflowTypeValue := ''
            else
                cxWorkflowTypeValue := cmbWorkFlowType.EditValue;

            if cmbPrecCategory.Text = '' then
               cmbPrecCategoryKeyValue := -1
            else
               cmbPrecCategoryKeyValue := cmbPrecCategory.EditValue;

            if cmbClassification.Text = '' then
               cmbClassificationKeyValue := -1
            else
               cmbClassificationKeyValue := cmbClassification.EditValue;

            if dblGroup.Text = '' then
               dblGroupKeyValue := -1
            else
               dblGroupKeyValue := dblGroup.EditValue;

            if memoPrecDetails.Text = '' then
                memoPrecDetailsValue := ''
            else
                memoPrecDetailsValue := memoPrecDetails.Text;

            if cmbTemplateType.Text = '' then
                cxTemplateTypeValue := ''
            else
                cxTemplateTypeValue := cmbTemplateType.EditValue;

            dmConnection.qryDoctemplate.insert;

            case AppType of
              1: SaveExcel(DocSequence, 1,btnTxtDocPath.Text,
                              True, True,'-1',
                              cmbAuthor.EditValue,txtDocName.Text,
                              cxWorkflowTypeValue, cxTemplateTypeValue,
                              dblGroupKeyValue, cmbPrecCategoryKeyValue, cmbClassificationKeyValue, -1,
                              memoPrecDetailsValue, edKeywords.Text, dmConnection.PrecID, True);
              2: begin
                 SaveDocument(DocSequence, 1,btnTxtDocPath.Text,
                              True, True,'-1',
                              cmbAuthor.EditValue,txtDocName.Text,
                              cxWorkflowTypeValue, cxTemplateTypeValue,
                              dblGroupKeyValue, cmbPrecCategoryKeyValue, cmbClassificationKeyValue, -1,
                              memoPrecDetailsValue, edKeywords.Text,
                              cbLeaveDocOpen.Checked, dmConnection.PrecID, True);
                  end;
              3: begin
                 SaveOutlookMessage(DocSequence, 1,btnTxtDocPath.Text,
                                    True, True, '-1',
                                    cmbAuthor.EditValue, txtDocName.Text,
                                    cxWorkflowTypeValue, cxTemplateTypeValue,
                                    dblGroupKeyValue, cmbPrecCategoryKeyValue, cmbClassificationKeyValue, -1,
                                    memoPrecDetailsValue, edKeywords.Text,
                                    ReceivedDate, IMail, True, dmConnection.PrecID );
                  end;
              4: SavePresentation(DocSequence, 1,btnTxtDocPath.Text,
                              True, True,'-1',
                              cmbAuthor.EditValue,txtDocName.Text,
                              cxWorkflowTypeValue, cxTemplateTypeValue,
                              dblGroupKeyValue, cmbPrecCategoryKeyValue, cmbClassificationKeyValue, -1,
                              memoPrecDetailsValue, edKeywords.Text);
            end;
            dmConnection.orsInsight.Commit;
         except
            raise;
         end;

      except
         dmConnection.orsInsight.Rollback;
      end;
      Self.Close;
   end
   else
   with Application do
   begin
      NormalizeTopMosts;
      MessageBox('Please enter a document name.','SpeediDocs',MB_OK+MB_ICONEXCLAMATION);
      RestoreTopMosts;
  end;
end;


procedure TfrmSavePrecDetails.btnEditMatterPropertiesValidate(
  Sender: TObject; var DisplayValue: Variant; var ErrorText: TCaption;
  var Error: Boolean);
begin
   if string(DisplayValue) <> '' then
   begin
      dmConnection.qryGetMatter.Close;
      dmConnection.qryGetMatter.ParamByName('FILEID').AsString := string(DisplayValue);
      dmConnection.qryGetMatter.Open;
      if dmConnection.qryGetMatter.Eof then
         MessageDlg('Invalid Matter Number', mtWarning, [mbOk], 0)
      else
      begin
         nMatter := dmConnection.qryGetMatter.FieldByName('NMATTER').AsInteger;
         FFileID := string(DisplayValue);
      end;
   end;
end;

procedure TfrmSavePrecDetails.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
      qryPrecClassification.Close;
      qryEmployee.Close;
      qryPrecCategory.Close;
      qryWorkFlowType.Close;
      dmConnection.tbDocGroups.Close;
end;

procedure TfrmSavePrecDetails.btnTxtDocPathPropertiesButtonClick(
  Sender: TObject; AButtonIndex: Integer);
begin
   case AButtonIndex of
      0: begin
            if BrowseDlg.Execute then
               btnTxtDocPath.Text := BrowseDlg.Directory;
         end;
      1: btnTxtDocPath.Text := SystemString('DRAG_DEFAULT_DIRECTORY');
   end;
end;

procedure TfrmSavePrecDetails.btnTxtDocPathPropertiesChange(Sender: TObject);
begin
   if (FromWord = False) then
      btnTxtDocPath.Text := SystemString('DRAG_DEFAULT_DIRECTORY')
   else
      btnTxtDocPath.Text := SystemString('DOC_DEFAULT_DIRECTORY');
end;

procedure TfrmSavePrecDetails.cbSaveAsPrecedentClick(Sender: TObject);
begin
   btnTxtDocPath.Text := SystemString('DFLT_PRECEDENT_PATH')
end;

procedure TfrmSavePrecDetails.GetDetails;
var
  varWord
  ,varDoc
  ,DocProps
  ,OLEvar
  ,Item : OleVariant;
  x
  ,i
  ,Count
  ,nRet: integer;
  Value: OleVariant;
  IProps: DocumentProperties;
  IProp: DocumentProperty;
  PropValue: OleVariant;
  PropName: widestring;

begin
   btnTxtDocPath.Text := SystemString('DOC_DEFAULT_DIRECTORY');
   try
      //*********************************
      GetWordApp.ActiveDocument.CustomDocumentProperties.QueryInterface(IID_DocumentProperties, IProps);
      if Assigned(IProps) then
      try
         IProps.Get_Count(Count);  //***values already set
         if (Count > 0) then
         begin
           for i := 1 to length(frmSavePrecDetails.AWordProps) do
           begin
              IProps.Get_Item(i, frmSavePrecDetails.LadxLCID, IProp);
              if Assigned(IProp) then
              try
                 nRet := IProp.Get_Value(frmSavePrecDetails.LadxLCID, PropValue);

                 IProp.Get_Name(frmSavePrecDetails.LadxLCID,PropName);
                 AWordProps[i].PropName := PropName;
                 AWordProps[i].PropValue := PropValue;


                 if AWordProps[I].PropName = 'DocID' then
                 begin
                     dmConnection.DocID := StrToInt(AWordProps[I].PropValue);
                     FDocName := TableString('DOC','DOCID', dmConnection.DocID, 'DOC_NAME');
                     if FDocName = '' then
                        FDocName := GetWordApp.ActiveDocument.Name;

                 end;

                 if AWordProps[I].PropName = 'Prec_Category' then
                 begin
                     try
                        FPrec_Category := AWordProps[I].PropValue;
                        cmbPrecCategory.EditValue := FPrec_Category;
                     except
                        ;// in case of errors
                     end;
                 end;

                 if AWordProps[I].PropName = 'Prec_Classification' then
                 begin
                     try
                        FPrec_Classification := AWordProps[I].PropValue;
                        cmbClassification.EditValue := FPrec_Classification;
                     except
                        ;// in case it doesnt exist
                     end;
                 end;

                 if AWordProps[I].PropName = 'Doc_Keywords' then
                 begin
                     try
                        FDoc_Keywords := AWordProps[I].PropValue;
                        edKeywords.Text := FDoc_Keywords;
                     except
                        ;// in case it doesnt exist
                     end;
                 end;

                 if AWordProps[I].PropName = 'Doc_Precedent' then
                 begin
                     try
                        FDoc_Precedent := AWordProps[I].PropValue;
                        memoPrecDetails.Text := FDoc_Precedent;
                     except
                        ;// in case it doesnt exist
                     end;
                 end;

                 if AWordProps[I].PropName = 'Doc_FileName' then
                 begin
                     try
                        FDoc_FileName := AWordProps[I].PropValue;
//                        btnTxtDocPath.Text := FDoc_FileName;
                     except
                        ;// in case it doesnt exist
                     end;
                 end;

                 if AWordProps[I].PropName = 'Doc_Title' then
                 begin
                     try
                        TxtDocName.Text := AWordProps[I].PropValue;
                     except
                        ;// in case it doesnt exist
                     end;
                 end;

                 if AWordProps[I].PropName = 'Doc_Author' then
                 begin
                     try
                        FDoc_Author := AWordProps[I].PropValue;
                        cmbAuthor.EditValue := FDoc_Author;
                     except
                        cmbAuthor.EditValue := dmConnection.UserCode;// in case it doesnt exist
                     end;
                 end;

                 if AWordProps[I].PropName = 'Saved_in_DB' then
                 begin
                    FSavedInDB := AWordProps[I].PropValue;
                    if FSavedInDB = 'Y' then
                    begin
                       btnTxtDocPath.Text := FDocName;
                    end;
                 end;

                 if (txtDocName.Text = '') and (dmConnection.DocID > 0) then
                     txtDocName.Text := TableString('DOC','DOCID', dmConnection.DocID, 'DESCR'); //  DocName;

              finally
                 IProp := nil;
              end;
           end;
         end;
//           dblGroup.EditValue := tbDocTemplatesEdit.FieldByName('GROUPID').AsInteger;
      finally
         IProps := nil;
      end;
   except
     //
   end;
end;


end.
