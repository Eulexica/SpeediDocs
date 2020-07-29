unit SavedocDetails;

{************************************************************************
 AES 24/6/2018 added ability to search for matter based on description/client name/fileid ssame as search on insight.


*************************************************************************}

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComObj, StdCtrls, Menus, DB, ActnList,
  ActnMan, Vcl.ExtCtrls, Vcl.ImgList, Vcl.Buttons, Vcl.ComCtrls, Vcl.DBCtrls,
  SaveDoc, cxGraphics, cxControls, cxLookAndFeels, cxLookAndFeelPainters,
  cxContainer, cxEdit, cxTextEdit, cxMaskEdit, cxDropDownEdit, cxLookupEdit,
  cxDBLookupEdit, cxDBLookupComboBox, Outlook2000, cxImageComboBox, cxSpinEdit,
  DBAccess, Ora, MemDS, cxMemo, cxButtonEdit, JvBaseDlg, JvBrowseFolder,
  dxLayoutcxEditAdapters, dxLayoutControlAdapters, dxLayoutContainer, cxClasses,
  dxLayoutControl, cxLabel, System.ImageList;

const
     CUSTOMPROPS: array[0..10] of string = ('MatterNo','DocID','Prec_Category','Prec_Classification','Doc_Keywords','Doc_Precedent','Doc_FileName','Doc_Author','Saved_in_DB', 'Doc_Title','Portal_Access');

type
  TfrmSaveDocDetails = class(TForm)
    btnSave: TBitBtn;
    btnClose: TBitBtn;
    ImageList1: TImageList;
    cbPortalAccess: TCheckBox;
    cbOverwriteDoc: TCheckBox;
    cbLeaveDocOpen: TCheckBox;
    cbNewCopy: TCheckBox;
    lblMatter: TLabel;
    lblDescription: TLabel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    StatusBar: TStatusBar;
    cmbPrecCategory: TcxLookupComboBox;
    cmbClassification: TcxLookupComboBox;
    cmbAuthor: TcxLookupComboBox;
    Label7: TLabel;
    neUnits: TcxSpinEdit;
    chkCreateTime: TCheckBox;
    txtDocName: TcxTextEdit;
    edKeywords: TcxTextEdit;
    memoTimeNarration: TcxMemo;
    memoPrecDetails: TcxMemo;
    cmbTasks: TcxLookupComboBox;
    btnEditMatter: TcxButtonEdit;
    btnTxtDocPath: TcxButtonEdit;
    BrowseDlg: TJvBrowseForFolderDialog;
    dxLayoutControl1Group_Root: TdxLayoutGroup;
    dxLayoutControl1: TdxLayoutControl;
    dxLayoutItem1: TdxLayoutItem;
    dxLayoutItem2: TdxLayoutItem;
    dxLayoutItem3: TdxLayoutItem;
    dxLayoutItem4: TdxLayoutItem;
    dxLayoutItem5: TdxLayoutItem;
    dxLayoutItem6: TdxLayoutItem;
    dxLayoutItem7: TdxLayoutItem;
    dxLayoutItem8: TdxLayoutItem;
    dxLayoutItem9: TdxLayoutItem;
    dxLayoutItem10: TdxLayoutItem;
    dxLayoutItem11: TdxLayoutItem;
    dxLayoutItem12: TdxLayoutItem;
    dxLayoutItem13: TdxLayoutItem;
    dxLayoutItem15: TdxLayoutItem;
    dxLayoutItem16: TdxLayoutItem;
    dxLayoutItem17: TdxLayoutItem;
    dxLayoutItem14: TdxLayoutItem;
    dxLayoutItem18: TdxLayoutItem;
    dxLayoutItem19: TdxLayoutItem;
    dxLayoutGroup1: TdxLayoutGroup;
    dxLayoutGroup2: TdxLayoutGroup;
    dxLayoutGroupTime: TdxLayoutGroup;
    dxLayoutGroup4: TdxLayoutGroup;
    dxLayoutItem20: TdxLayoutItem;
    cmbFolder: TcxLookupComboBox;
    dxLayoutItem21: TdxLayoutItem;
    dxLayoutGroupTimeFields: TdxLayoutGroup;
    Memo1: TcxMemo;
    procedure btnCloseClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure rgStorageClick(Sender: TObject);
    procedure btnEditMatterPropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnTxtDocPathPropertiesButtonClick(Sender: TObject;
      AButtonIndex: Integer);
    procedure cmbCategoryPropertiesInitPopup(Sender: TObject);
    procedure btnEditMatterExit(Sender: TObject);
    procedure dockbtnDefaultPathClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure cbSaveAsPrecedentClick(Sender: TObject);
    procedure chkCreateTimeClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure btnEditMatterPropertiesButtonClick(Sender: TObject;
      AButtonIndex: Integer);
    procedure FormCreate(Sender: TObject);
    procedure btnEditMatterKeyPress(Sender: TObject; var Key: Char);
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
    FromExcel: boolean;
    FTimeNarration: string;
    FSentEmail: boolean;
    FFolder_ID: integer;
    FCategories: string;
    FProp: Outlook2000.UserProperty;
    bMatterFound: boolean;
//    procedure GetDetails;
  public
    { Public declarations }
    AWordProps: array[1..11] of TWordProperties;
    property DocName: string read FDocName;
    property AppType: Integer read FAppType write FAppType;
    property MailSubject: string read FMailSubject write FMailSubject;
    property ReceivedDate: TDateTime read FReceivedDate write FReceivedDate;
    property LadxLCID: integer read FadxLCID write FadxLCID;
    property IMail: MailItem read FIMail write FIMail;
    property TimeNarration: string read FTimeNarration write FTimeNarration;
    property SentEmail: boolean read FSentEmail write FSentEmail default False;
    procedure GetDetails;
  end;

var
  frmSaveDocDetails: TfrmSaveDocDetails;

function ShowDocSave: Integer; StdCall;

implementation

uses
    MatterSearch, SaveDocFunc, ActiveX, WordUnit, OutlookUnit, ExcelUnit,
    PowerPointUnit, SpeediDocs_IMPL, Office2010;

{$R *.dfm}

function ShowDocSave:integer;
var
   frmSaveDocDetails: TfrmSaveDocDetails;
begin
//   Application.Handle := AHandle;
   frmSaveDocDetails := TfrmSaveDocDetails.Create(Application);
   try
      frmSaveDocDetails.ShowModal;
      Result := frmSaveDocDetails.nMatter;
   finally
      FreeAndNil(frmSaveDocDetails);
   end;
end;

procedure TfrmSaveDocDetails.btnCloseClick(Sender: TObject);
var
  Unknown: IUnknown;
  OLEResult: HResult;
  AMacro : string;
begin
   Close;
end;

procedure TfrmSaveDocDetails.FormShow(Sender: TObject);
var
   AMatter: string;
begin
   cmbAuthor.EditValue := dmConnection.UserCode;
   FromWord := False;
   bMatterFound := False;
   txtDocName.Text := MailSubject;
   Memo1.Text := '';
   FFolder_ID := 0;
   StatusBar.Panels[0].Text := 'Ver: '+ ReportVersion(GetModuleName(HInstance)) + ' (' +DateTimeToStr(FileDateToDateTime(FileAge(GetModuleName(HInstance))))+')';

   case AppType of
      1: begin
            FromExcel := True;
            GetDetails;
         end;
      2: begin
            FromWord := True;
            GetDetails;
         end;
      3: begin
            cbLeaveDocOpen.Checked := False;
            cmbAuthor.EditValue := dmConnection.UserCode;
            memoTimeNarration.Text := TimeNarration;
            FProp := nil;
            FCategories := FIMail.Categories;
            FProp := FIMail.UserProperties.Find('MATTER', True);
            if Assigned(Fprop) then
            begin
                AMatter := Fprop.Value;
                Memo1.Text := 'Email already saved to BHL Insight in matter ' + AMatter;
            end
            else
            begin
                Memo1.Text := '';
            end;

         end;
      4: ;
   end;
   if (FSavedInDB = 'N') or (FSavedInDB = '')  then
   begin
//      rgStorage.ItemIndex := SystemInteger('DFLT_DOC_SAVE_OPTION');
      if ((FromWord = True) or (FromExcel = True)) then
         btnTxtDocPath.Text := SystemString('DOC_DEFAULT_DIRECTORY')  // 'DRAG_DEFAULT_DIRECTORY');
      else
         btnTxtDocPath.Text := SystemString('DRAG_DEFAULT_DIRECTORY');
   end;
end;

procedure TfrmSaveDocDetails.btnSaveClick(Sender: TObject);
var
   DocSequence,
   lTask: string;
//   bUsePath: boolean;
   cmbPrecCategoryKeyValue,
   cmbClassificationKeyValue, cmbFolderKeyValue: integer;
begin
   try
      screen.Cursor := crHourGlass;
      if (btnEditMatter.Text = '') then
      begin
         with Application do
         begin
            NormalizeTopMosts;
            MessageBox('Please enter a Matter number.','SpeediDocs',MB_OK+MB_ICONEXCLAMATION);
            RestoreTopMosts;
            exit;
         end;
      end;
      if (Memo1.Text <> '') then
      begin
         with Application do
         begin
            NormalizeTopMosts;
            if MessageBox('Are you sure you want to file to BHL Insight again?', 'SpeediDocs', MB_ICONQUESTION+MB_YESNO) = mrNo then
            begin
                RestoreTopMosts;
                exit;
            end;
         end;
      end;
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
            dmConnection.qryMatterAttachments.Open;

            FEditing := False;
//            bUsePath := False;
            tmpdir := IncludeTrailingPathDelimiter(GetEnvironmentVariable('TMP'));

            if ((cbOverwriteDoc.Visible)  and
               (not cbOverwriteDoc.Checked)) then
            begin
               dmConnection.qryMatterAttachments.insert;
               dmConnection.qryMatterAttachments.ParamByName('docid').AsInteger := dmConnection.DocID;
            end
            else
            if (not cbOverwriteDoc.Visible) then
            begin
               dmConnection.qryMatterAttachments.Insert;
               dmConnection.qryMatterAttachments.ParamByName('docid').AsInteger := dmConnection.DocID;
            end
            else
            if (cbOverwriteDoc.Checked) then
            begin
               dmConnection.qryMatterAttachments.Edit;
               FEditing := True;
            end;

//            if bUsePath then
//            begin
//               tmpDir := btnTxtDocPath.Text + '\';
//            end;

//            if txtDocName.Text = '' then
//            begin
//               tmpFileName := tmpDir + dmSaveDoc.DocID +'.doc';
//            end
//            else
//            begin
               if btnTxtDocPath.Text = '' then
                  tmpFileName := txtDocName.Text
               else
                  tmpFileName := btnTxtDocPath.Text;

            try
               if cmbPrecCategory.Text = '' then
                  cmbPrecCategoryKeyValue := -1
               else
                  cmbPrecCategoryKeyValue := cmbPrecCategory.EditValue;

               if cmbClassification.Text = '' then
                  cmbClassificationKeyValue := -1
               else
                  cmbClassificationKeyValue := cmbClassification.EditValue;

               if cmbFolder.Text = '' then
                  cmbFolderKeyValue := -1
               else
                  cmbFolderKeyValue := cmbFolder.EditValue;

               case AppType of
                 1: begin
                       SaveExcel(DocSequence, 1,btnTxtDocPath.Text,
                                   cbNewCopy.Checked, cbOverwriteDoc.Checked,btnEditMatter.Text,
                                   cmbAuthor.EditValue, txtDocName.Text,
                                   '','',
                                   -1, cmbPrecCategoryKeyValue, cmbClassificationKeyValue, cmbFolderKeyValue,
                                   '', edKeywords.Text, dmConnection.DocID);
                    end;
                 2: begin
                       SaveDocument(DocSequence, 1,btnTxtDocPath.Text,
                                    cbNewCopy.Checked, cbOverwriteDoc.Checked, btnEditMatter.Text,
                                    cmbAuthor.EditValue, txtDocName.Text,
                                    '', '',
                                    -1,cmbPrecCategoryKeyValue, cmbClassificationKeyValue, cmbFolderKeyValue,
                                    '', edKeywords.Text,
                                    cbLeaveDocOpen.Checked, dmConnection.DocID, False, LadxLCID,
                                    lTask, chkCreateTime.Checked, neUnits.Value,
                                    memoTimeNarration.Text);
                    end;
                 3: begin
                       if (cmbTasks.Text <> '') then
                          lTask := cmbTasks.EditValue;
                       SaveOutlookMessage(DocSequence, 1,btnTxtDocPath.Text,
                                    cbNewCopy.Checked, cbOverwriteDoc.Checked, btnEditMatter.Text,
                                    cmbAuthor.EditValue, txtDocName.Text,
                                    '','',
                                    -1, cmbPrecCategoryKeyValue, cmbClassificationKeyValue, cmbFolderKeyValue,
                                    '', edKeywords.Text,
                                    ReceivedDate, IMail, True, dmConnection.DocID,
                                    chkCreateTime.Checked, memoTimeNarration.Text,
                                    neUnits.Value, SentEmail, lTask);
                    end;
                 4: begin
                       SavePresentation(DocSequence, 1,btnTxtDocPath.Text,
                                    cbNewCopy.Checked, cbOverwriteDoc.Checked,btnEditMatter.Text,
                                    cmbAuthor.EditValue, txtDocName.Text,
                                    '','',
                                    -1, cmbPrecCategoryKeyValue, cmbClassificationKeyValue, cmbFolderKeyValue,
                                    '', edKeywords.Text);
                 end;
               end;
               dmConnection.orsInsight.Commit;
//               if (rgStorage.ItemIndex = 0) and (not cbLeaveDocOpen.Checked) then
//                  DeleteFile(tmpFileName);
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
   finally
      screen.Cursor := crDefault;
   end;
end;


procedure TfrmSaveDocDetails.rgStorageClick(Sender: TObject);
begin
{   case rgStorage.ItemIndex of
      0: begin
            btnTxtDocPath.Visible := False;
            Self.Height := 275;
         end;
      1: begin}
            btnTxtDocPath.Visible := True;
            Self.Height := 307;
//         end;
//   end;
end;

procedure TfrmSaveDocDetails.btnEditMatterPropertiesButtonClick(Sender: TObject;
  AButtonIndex: Integer);
var
   frmMtrSearch: TfrmMtrSearch;
begin
   try
      FreeAndNil(frmMtrSearch);
      frmMtrSearch := TfrmMtrSearch.Create(Application);
      if (frmMtrSearch.ShowModal = mrOK) then
      begin
         btnEditMatter.Text := frmMtrSearch.tvMattersFILEID.EditValue;   // dmSaveDoc.qryMatters.FieldByName('fileid').AsString;   //  dmSaveDoc.qryMatters.FieldByName('fileid').AsString;
         nMatter := frmMtrSearch.tvMattersNMATTER.EditValue;  // dmSaveDoc.qryMatters.FieldByName('nmatter').AsInteger;
//         cmbAuthor.EditValue := TableString('MATTER','NMATTER',nMatter,'AUTHOR');
         FFileID := btnEditMatter.Text;
         cbOverwriteDoc.Enabled := (FOldFileID = FFileID);
         cbNewCopy.Visible := ((FOldFileID <> FFileID) and (FOldFileID <> ''));
         Label7.Caption := TableString('MATTER','NMATTER',nMatter,'SHORTDESCR');
         dmConnection.qryMatterFolderList.Close;
         dmConnection.qryMatterFolderList.ParamByName('nMatter').AsInteger := NMATTER;
         dmConnection.qryMatterFolderList.Open;
      end;
   finally
      FreeAndNil(frmMtrSearch);
   end;
end;

procedure TfrmSaveDocDetails.btnEditMatterPropertiesValidate(
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
//         cmbAuthor.EditValue := TableString('MATTER','NMATTER',nMatter,'AUTHOR');
         cbOverwriteDoc.Enabled := (FOldFileID = FFileID);
         cbNewCopy.Visible := ((FOldFileID <> FFileID) and (FOldFileID <> ''));
      end;
   end;
end;

procedure TfrmSaveDocDetails.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
   Action := caFree;
end;

procedure TfrmSaveDocDetails.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
   dmConnection.tbDocGroups.Close;
   dmConnection.qryPrecClassification.Close;
   dmConnection.qryEmployee.Close;
   dmConnection.qryPrecCategory.Close;
   dmConnection.qryScaleCost.Close;
   dmConnection.qryMatterFolderList.Close;
//   dmConnection.orsInsight.Disconnect;
{   if Assigned(dmSaveDoc) then
   begin
      FreeAndNil(dmSaveDoc);
   end;  }
end;

procedure TfrmSaveDocDetails.FormCreate(Sender: TObject);
begin
   if (not Assigned(dmConnection)) then
      dmConnection := TdmSaveDoc.Create(Application);

   if dmConnection.orsInsight.Connected = False then
   begin
      try
         if (dmConnection.GetUserID = True) then
         begin
            cbOverWriteDoc.Visible := False;

            StatusBar.Panels[0].Text := 'Ver: '+ ReportVersion(SysUtils.GetModuleName(HInstance)) + ' (' +DateTimeToStr(FileDateToDateTime(FileAge(SysUtils.GetModuleName(HInstance))))+')';
//            rgStorage.Enabled := (SystemString('DISABLE_SAVE_MODE') = 'N');
         end;
      except
         Exit;
      end;
   end;
   dmConnection.qryPrecClassification.Open;
   dmConnection.qryEmployee.Open;
   dmConnection.qryPrecCategory.Open;
   dmConnection.tbDocGroups.Open;
   dmConnection.qryScaleCost.Open;
   dmConnection.qryMatterFolderList.Open;
end;

procedure TfrmSaveDocDetails.btnTxtDocPathPropertiesButtonClick(
  Sender: TObject; AButtonIndex: Integer);
begin
   case AButtonIndex of
      0: begin
            if BrowseDlg.Execute then
               btnTxtDocPath.Text := BrowseDlg.Directory
         end;
      1: btnTxtDocPath.Text := SystemString('DRAG_DEFAULT_DIRECTORY');
   end;
end;

procedure TfrmSaveDocDetails.cbSaveAsPrecedentClick(Sender: TObject);
begin
    btnTxtDocPath.Text := SystemString('DOC_DEFAULT_DIRECTORY');
end;

procedure TfrmSaveDocDetails.chkCreateTimeClick(Sender: TObject);
begin
   dxLayoutGroupTimeFields.Enabled := chkCreateTime.Checked;
end;

procedure TfrmSaveDocDetails.cmbCategoryPropertiesInitPopup(
  Sender: TObject);
begin
//   dmSavedoc.qryPrecCategory.Close;
//   dmSavedoc.qryPrecCategory.Open;
end;

procedure TfrmSaveDocDetails.btnEditMatterExit(Sender: TObject);
var
   lFileID,
   lFoundFileID: string;
   nmatter: integer;
begin
   if (string(btnEditMatter.Text) <> '') and (bMatterFound = False) then
   begin
      lFileID := PadFileID(btnEditMatter.Text);
      dmConnection.LoadMatter(lFoundFileID, nmatter, lFileID);
      btnEditMatter.Text := lFoundFileID;
//      dmConnection.qryMatterFolderList.Close;
//      dmConnection.qryMatterFolderList.ParamByName('nMatter').AsInteger := nmatter;
//      dmConnection.qryMatterFolderList.Open;


{      dmConnection.qryGetMatter.Close;
      dmConnection.qryGetMatter.ParamByName('FILEID').AsString := string(btnEditMatter.Text);
      dmConnection.qryGetMatter.Open;
      if dmConnection.qryGetMatter.Eof then
         MsgErr('Invalid Matter Number')
      else     }
//      begin
//         nMatter := dmConnection.qryGetMatter.FieldByName('NMATTER').AsInteger;
         FFileID := lFileID;  //string(btnEditMatter.Text);
         Label7.Caption := TableString('MATTER','FILEID', FFileID,'SHORTDESCR');
//         cmbAuthor.EditValue := TableString('MATTER','NMATTER',nMatter,'AUTHOR');
         cbOverwriteDoc.Enabled := (FOldFileID = FFileID);
         cbNewCopy.Visible := ((FOldFileID = FFileID) and (FOldFileID = ''));
         dmConnection.qryMatterFolderList.Close;
         dmConnection.qryMatterFolderList.ParamByName('nMatter').AsInteger := NMATTER;
         dmConnection.qryMatterFolderList.Open;
         Self.ActiveControl := txtDocName;
         txtDocName.SetFocus;
//      end;
   end;
end;


procedure TfrmSaveDocDetails.btnEditMatterKeyPress(Sender: TObject;
  var Key: Char);
var
   lFileID,
   lFoundFileID: string;
   nmatter: integer;
begin
   if (key = #$D) then
   begin
      try
         bMatterFound := False;
         lFileID := PadFileID(btnEditMatter.Text);
         dmConnection.LoadMatter(lFoundFileID, nmatter, lFileID);
         btnEditMatter.Text := lFoundFileID;
      finally
         if (lFileId <> '') then
         begin
            bMatterFound := True;
            Label7.Caption := TableString('MATTER','NMATTER',nMatter,'SHORTDESCR');
            cbOverwriteDoc.Enabled := (FOldFileID = FFileID);
            cbNewCopy.Visible := ((FOldFileID = FFileID) and (FOldFileID = ''));
            dmConnection.qryMatterFolderList.Close;
            dmConnection.qryMatterFolderList.ParamByName('nMatter').AsInteger := NMATTER;
            dmConnection.qryMatterFolderList.Open;
            Self.ActiveControl := txtDocName;
            txtDocName.SetFocus;
         end;
      end;
//    btnEditMatter.ValidateEdit();
//      dmConnection.qryMatterFolderList.Closed;
//      dmConnection.qryMatterFolderList.ParamByName('nMatter').AsInteger := nmatter;
//      dmConnection.qryMatterFolderList.Open;
   end;
end;

procedure TfrmSaveDocDetails.dockbtnDefaultPathClick(Sender: TObject);
begin
   if (FromWord = False) and (FromExcel = False) then
      btnTxtDocPath.Text := SystemString('DRAG_DEFAULT_DIRECTORY')
   else
      btnTxtDocPath.Text := SystemString('DOC_DEFAULT_DIRECTORY');
end;

procedure TfrmSaveDocDetails.GetDetails;
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
      if FromWord then
         GetWordApp.ActiveDocument.CustomDocumentProperties.QueryInterface(IID_DocumentProperties, IProps);
      if FromExcel then
         GetExcelApp.ActiveWorkbook.CustomDocumentProperties.QueryInterface(IID_DocumentProperties, IProps);

      if Assigned(IProps) then
      try
         IProps.Get_Count(Count);  //***values already set
         if (Count > 0) then
         begin
           for i := 1 to length(AWordProps) do
           begin
              IProps.Get_Item(i, LadxLCID, IProp);
              if Assigned(IProp) then
              try
                 nRet := IProp.Get_Value(LadxLCID, PropValue);

                 IProp.Get_Name(LadxLCID,PropName);
                 AWordProps[i].PropName := PropName;
                 AWordProps[i].PropValue := PropValue;

                 if AWordProps[I].PropName = 'MatterNo' then
                 begin
                    try
                       FFileID := AWordProps[I].PropValue;
                       FOldFileID := FFileID;
                     except
                       ; //
                     end;
                     btnEditMatter.Text := FFileID;
                     nMatter := TableInteger('MATTER','FILEID',FFileID,'NMATTER');
                     btnEditMatterExit(nil);
                 end;

                 if AWordProps[I].PropName = 'DocID' then
                 begin
                     dmConnection.DocID := StrToInt(AWordProps[I].PropValue);
                     FDocName := TableString('DOC','DOCID', dmConnection.DocID, 'DOC_NAME');
                     if FDocName = '' then
                        FDocName := GetWordApp.ActiveDocument.Name;

                     cbOverWriteDoc.Visible := True;
                 end;

                 if AWordProps[I].PropName = 'Prec_Category_ID' then
                 begin
                     try
                        FPrec_Category := AWordProps[I].PropValue;
                        if FPrec_Category <> '' then
                           cmbPrecCategory.EditValue := FPrec_Category;
                     except
                        ;// in case of errors
                     end;
                 end;

                 if AWordProps[I].PropName = 'Prec_Classification_ID' then
                 begin
                     try
                        FPrec_Classification := AWordProps[I].PropValue;
                        if FPrec_Classification <> '' then
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
//                       rgStorage.ItemIndex := 0;
                       btnTxtDocPath.Text := FDocName;
                    end;
                 end;

                 if (txtDocName.Text = '') and (dmConnection.DocID > 0) then
                     txtDocName.Text := TableString('DOC','DOCID', dmConnection.DocID, 'DESCR'); //  DocName;

                 if AWordProps[I].PropName = 'Portal_Access' then
                     cbPortalAccess.Checked := (AWordProps[I].PropValue = 'Y');
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
