unit savedocdtls;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, JvExStdCtrls, JvCheckBox, ExtCtrls, JvExExtCtrls,
  JvRadioGroup, JvMemo, JvEdit, Mask, JvExMask, JvToolEdit, JvDBLookup,
  JvDBLookupComboEdit, LMDEdit, LMDControl, LMDCustomControl, LMDCustomPanel,
  LMDCustomBevelPanel, LMDBaseEdit, LMDCustomEdit, LMDCustomBrowseEdit,
  LMDBrowseEdit, JvExControls, JvLabel, Buttons, JvExButtons, JvButtons,
  MemDS, DBCtrls, LMDCustomButton, LMDDockButton, ComCtrls,
  ComObj, Outlook2000;

const
     CUSTOMPROPS: array[0..10] of string = ('MatterNo','DocID','Prec_Category','Prec_Classification','Doc_Keywords','Doc_Precedent','Doc_FileName','Doc_Author','Saved_in_DB', 'Doc_Title','Portal_Access');

type
  TfrmSaveDocDtls = class(TForm)
    JvLabel1: TJvLabel;
    JvLabel2: TJvLabel;
    JvLabel3: TJvLabel;
    JvLabel4: TJvLabel;
    JvLabel5: TJvLabel;
    JvLabel6: TJvLabel;
    JvLabel7: TJvLabel;
    JvLabel8: TJvLabel;
    TxtDocName: TLMDEdit;
    edKeywords: TJvEdit;
    memoPrecDetails: TJvMemo;
    cbLeaveDocOpen: TJvCheckBox;
    cbOverwriteDoc: TJvCheckBox;
    cbPortalAccess: TJvCheckBox;
    btnTxtDocPath: TLMDBrowseEdit;
    JvHTButton1: TJvHTButton;
    JvHTButton2: TJvHTButton;
    cmbCategory: TDBLookupComboBox;
    cmbClassification: TDBLookupComboBox;
    cmbAuthor: TDBLookupComboBox;
    cbNewCopy: TJvCheckBox;
    LMDDockButton1: TLMDDockButton;
    btnEditMatter: TJvEdit;
    StatusBar: TStatusBar;
    procedure JvHTButton2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure LMDDockButton1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure JvHTButton1Click(Sender: TObject);
  private
    { Private declarations }
    FUserID : string;
    FEntity : string;
    FDocID   : string;
    FOldFileID: string;
    FFileID: string;
    nMatter: integer;
    FDocName: string;
    FURLOnly: boolean;
    tmpFileName: string;

    FPrec_Category: string;
    FEditing: boolean;
    FSavedInDB: string;
    FPrec_Classification: string;
    FDoc_Keywords: string;
    FDoc_Precedent: string;
    FDoc_FileName: string;
    FDoc_Author: string;
    FAppType: integer;
    procedure GetDetails;
  public
    { Public declarations }
     function SaveDocument(DocSequence: string): boolean;
     function SaveOutlookMessage(DocSequence: string): boolean;
     property AppType: Integer read FAppType write FAppType;
  end;

var
  frmSaveDocDtls: TfrmSaveDocDtls;

implementation

uses savedoc, MatterSearch, SaveDocFunc, Office2000, ActiveX, Word2000, InsightOfficeAddIn_IMPL;

{$R *.dfm}

procedure TfrmSaveDocDtls.FormClose(Sender: TObject; var Action: TCloseAction);
begin
   dmSaveDoc.qryPrecCategory.Close;
   dmSaveDoc.qryPrecClassification.Close;
   dmSaveDoc.qryEmployee.Close;
   Action :=  caFree;
end;

procedure TfrmSaveDocDtls.FormShow(Sender: TObject);

begin
   try
      dmSaveDoc := TdmSaveDoc.Create(Application);
      if GetUserID() then
      begin
         cbOverWriteDoc.Visible := False;
         if (FSavedInDB = 'N') or (FSavedInDB = '')  then
         begin
//            rgStorage.ItemIndex := SystemInteger('DFLT_DOC_SAVE_OPTION');
//            btnTxtDocPath.Text := SystemString('DOC_DEFAULT_DIRECTORY');
         end;

         case AppType of
          1: ;
          2: GetDetails;
          3:  ;
         end;

         dmSaveDoc.qryPrecCategory.Open;
         dmSaveDoc.qryPrecClassification.Open;
         dmSaveDoc.qryEmployee.Open;
         btnTxtDocPath.Text := SystemString('DOC_DEFAULT_DIRECTORY');
//         dmSaveDoc.qryMatters.Active := True;
//         dmSavedoc.qryPrecCategory.Open;
         StatusBar.Panels.Items[0].Text := 'Ver: '+ReportVersion + ' (' +DateTimeToStr(FileDateToDateTime(FileAge(Application.ExeName)))+')';
//         rgStorage.Enabled := (SystemString('DISABLE_SAVE_MODE') = 'N');
      end
      else
         Application.MessageBox('Could not connect to Insight database5','Insight');
         Self.Close;
   except
      Application.MessageBox('Could not connect to Insight database6','Insight');
//      frmSavedoc.Close;
      ModalResult := mrCancel;
   end;
end;

procedure TfrmSaveDocDtls.JvHTButton1Click(Sender: TObject);
var
   DocSequence, tmpdir: string;
//   bUsePath: boolean;
begin
   if btnEditMatter.Text = '' then
   begin
      with Application do
      begin
         NormalizeTopMosts;
         MessageBox('Please enter a Matter number.','DocToDBSave',MB_OK+MB_ICONEXCLAMATION);
         RestoreTopMosts;
         exit;
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
               MessageBox('Please enter an Author.','DocToDBSave',MB_OK+MB_ICONEXCLAMATION);
               RestoreTopMosts;
               exit;
            end;
         end;
         dmSaveDoc.orsInsight.StartTransaction;
         dmSaveDoc.qryMatterAttachments.ParamByName('docid').AsString := dmSaveDoc.DocID;
         dmSaveDoc.qryMatterAttachments.Open;

         FEditing := False;
//         bUsePath := False;
         tmpdir := GetEnvironmentVariable('TMP')+'\';

         if ((cbOverwriteDoc.Visible)  and
            (not cbOverwriteDoc.Checked)) then
            dmSaveDoc.qryMatterAttachments.insert
         else
         if (not cbOverwriteDoc.Visible) then
            dmSaveDoc.qryMatterAttachments.Insert
         else
         if (cbOverwriteDoc.Checked) then
         begin
            dmSaveDoc.qryMatterAttachments.Edit;
            FEditing := True;
         end;

            if btnTxtDocPath.Text = '' then
               tmpFileName := txtDocName.Text
            else
               tmpFileName := btnTxtDocPath.Text;

         try
            case AppType of
              1: ;
              2: SaveDocument(DocSequence);
              3: SaveOutlookMessage(DocSequence);
            end;
            dmSaveDoc.orsInsight.Commit;
            if (not cbLeaveDocOpen.Checked) then
               DeleteFile(tmpFileName);
         except
            raise;
         end;
      except
         dmSaveDoc.orsInsight.Rollback;
      end;
      Self.Close;
   end
   else
   with Application do
   begin
      NormalizeTopMosts;
      MessageBox('Please enter a document name.','DocToDBSave',MB_OK+MB_ICONEXCLAMATION);
      RestoreTopMosts;
  end;
end;

function TfrmSaveDocDtls.SaveOutlookMessage(DocSequence: string): boolean;
var
  Mail: MailItem;
  MailName: WideString;
   lSubject: string;
   x, i, FHashPos: integer;
   FileID, AParsedDocName: string;
   up: UserProperty;
   ups: UserProperties;
   AddInModule: TAddInModule;
begin
//   Mail := MailItem(AddInModule.OutlookApp.Session.GetItemFromID(AddInModule.OutlookApp.ActiveExplorer.Selection.Item(1) ,''));

   lSubject := Mail.Subject;

   if (pos('#',lSubject) > 0) then
   begin
      // clean up subject line
      for x := i + 1 to length(lSubject) do
      begin
         if (lSubject[x] in ['/', '\', '?','"','<','>','|','*',':']) then
            lSubject[x] := ' ';
      end;
   end;

   try
      FileID := Mail.UserProperties.Find('NPR_FILEID',olText).Value;
   except
      FileID := '';
   end;

   if ((FileID = '') and (pos('#',lSubject) > 0)) then
   begin
      FHashPos := pos('#',lSubject);
      for x := FHashPos + 1 to length(lSubject) do
      begin
         if (lSubject[x] <> ' ') and ((lSubject[x] in ['A'..'Z', '0'..'9', 'a'..'z'])) then
            FileID := FileID + lSubject[x];
         if (not (lSubject[x] in ['A'..'Z', '0'..'9', 'a'..'z'])) then break;
      end;
   end;

   if (FileID <> '') then
   begin
      MailName := btnTxtDocPath.Text +'\\'+ mail.Subject + '_' + '[DOCSEQUENCE]' + '.msg';
      AParsedDocName := ParseMacros(MailName, TableInteger('MATTER','FILEID',FileID,'NMATTER'));

      if not DirectoryExists(ExtractFileDir(AParsedDocName)) then
         ForceDirectories(ExtractFileDir(AParsedDocName));

//      'D:\\InsightDocs\\DocResults\\' + mail.Subject + '.msg';
      Mail.SaveAs(AParsedDocName,olMSG);
   end;
end;


function TfrmSaveDocDtls.SaveDocument(DocSequence: string): boolean;
var
//  varWord, varDocs, PropName, varDoc: OleVariant;
//   PropName: OleVariant;
  DocName, SavedInDB: string;
  nCat, nClass: integer;
  ltmpdir, AMacro: string;
  MSWord: _Application;
  MSDoc: _Document;
  Unknown: IUnknown;
  OLEResult: HResult;
  OLEvar: OleVariant;
  CustomDocProps, Item, Value, DocProps: OleVariant;
  i, x: integer;
  ADocID, AKeyWords, APrecDetails, AExt: string;
  PropValues: TStrings;
  bMoveSuccess: boolean;
begin
   SaveDocument := False;
   bMoveSuccess := True;

   OLEResult := GetActiveObject(CLASS_WordApplication, nil, Unknown);
   if (OLEResult = MK_E_UNAVAILABLE) then
      MSWord := CoWordApplication.Create          //get MS Word running
   else
   begin
      OleCheck(OLEResult);                           //check for errors
      OleCheck(Unknown.QueryInterface(_Application, MSWord));
   end;


   if(not VarIsNull(MSWord)) then
   begin
      try
         if (FOldFileID <> FFileID) and (FOldFileID <> '') and (not cbNewCopy.Checked) then
         begin
            tmpFileName := SystemString('DRAG_DEFAULT_DIRECTORY');
            tmpFileName := tmpFileName + '\' + ExtractFileName(btnTxtDocPath.Text);

            AExt := ExtractFileExt(tmpFileName);
            tmpFileName := Copy (tmpFileName,1, Length(tmpFileName)- Length(AExt));
            tmpFileName := tmpFileName + '_' + '[DOCSEQUENCE]';
            tmpFileName := tmpFileName + AExt;

            tmpFileName := ParseMacros(tmpFileName,TableInteger('MATTER','FILEID',uppercase(FFileID),'NMATTER'));

            if FOldFileID <> '' then
               bMoveSuccess := MoveMatterDoc(tmpFileName, btnTxtDocPath.Text);
         end
         else
         if (not FEditing) then
         begin
            if btnTxtDocPath.Text = '' then
            begin
               btnTxtDocPath.Text := SystemString('DOC_DEFAULT_DIRECTORY');
//               btnTxtDocPath.Text
            end;
            tmpFileName := btnTxtDocPath.Text;

            AExt := ExtractFileExt(tmpFileName);
            tmpFileName := Copy (tmpFileName,1, Length(tmpFileName)- Length(AExt));
            tmpFileName := tmpFileName + '_[DOCSEQUENCE]';
            tmpFileName := tmpFileName + AExt;

            tmpFileName := ParseMacros(tmpFileName,TableInteger('MATTER','FILEID',FFileID,'NMATTER'));
         end
         else
         begin
            tmpFileName := tmpFileName;
         end;

         if ExtractFileName(tmpFileName) = '' then
            tmpFileName  := tmpFileName + FFileID;

         if ExtractFileExt(tmpFileName) = '' then
            tmpFileName := tmpFileName + '.' + SystemString('default_doc_ext');  //'.doc';

         if ((DocName = '') or (pos('Document', DocName) > 0) or
            (ExtractFileName(btnTxtDocPath.Text) <> DocName)) and (not cbOverwriteDoc.Checked) then
         begin
            if not DirectoryExists(ExtractFileDir(tmpFileName)) then
               ForceDirectories(ExtractFileDir(tmpFileName));
            Value := tmpFileName;
            MSWord.ActiveDocument.SaveAs(Value, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
                                         EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
         end
         else
         begin
            MSWord.ActiveDocument.Save;
         end;

         AMacro := SystemString('WORD_SAVE_MACRO');
         if AMacro <> '' then MSWord.Run(AMacro, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
                                       EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
                                       EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
                                       EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
                                       EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam );


//            Value := False;
//            MSWord.ActiveDocument.Close(Value,EmptyParam, EmptyParam);
         MSWord.Activewindow.WindowState := wdWindowStateMaximize;

         try
            if bMoveSuccess then
            begin
            // write doc properties
               PropValues := TStringList.Create;
               MSWord.Visible := True;
//             varDocs := varWord.Documents;
               MSDoc := MSWord.ActiveDocument;

                  DocName := MSDoc.Name;

                  if FFileID = '' then
                     FFileID := btnEditMatter.Text
                  else
                  begin
                     if (FOldFileID <> FFileID) and (FOldFileID <> '') then
//                     if FileID <> btnEditMatter.Text then
                        FFileID := btnEditMatter.Text;
                  end;
                  PropValues.Add(FFileID);

                  ADocID := dmSaveDoc.DocID;
                  PropValues.Add(ADocID);

                  if varIsNull(cmbCategory.KeyValue) or
                     (VarToStr(cmbCategory.KeyValue) = '') then
                     nCat := -1
                  else
                  begin
                     try
                        nCat := cmbCategory.KeyValue;
                        FPrec_Category := IntToStr(nCat);
                     except
                        nCat := -1;
                     end;
                  end;
                  PropValues.Add(IntToStr(nCat));

                  if varIsNull(cmbClassification.KeyValue) or
                     (VarToStr(cmbClassification.KeyValue) = '') then
                     nClass := -1
                  else
                  begin
                     try
                        nClass := cmbClassification.KeyValue;
                        FPrec_Classification := IntToStr(nClass);
                     except
                        nClass := -1;
                     end;
                  end;
                  PropValues.Add(IntToStr(nClass));

                  AKeyWords := edKeywords.Text;
                  PropValues.Add(AKeyWords);

                  APrecDetails := memoPrecDetails.Text;
                  PropValues.Add(APrecDetails);

                   // empty value for file name.  file name is generated and saved later
                  PropValues.Add('');

                 // add author to array
                  PropValues.Add(cmbAuthor.KeyValue);

                  PropValues.Add(SavedInDB);

                  // document description - title
                  PropValues.Add(txtDocName.Text);

                  if cbPortalAccess.Checked then
                     PropValues.Add('Y')
                  else
                     PropValues.Add('N');

                  CustomDocProps := MSDoc.CustomDocumentProperties;
                  DocProps := MSDoc.BuiltInDocumentProperties;

                  for x := 0 to (length(CUSTOMPROPS) - 1) do
                  begin
                     OLEvar := CUSTOMPROPS[x];
                     Value := PropValues.Strings[x];
                     try
                        for I := 1 to length(CUSTOMPROPS) {CustomDocProps.Count} do // Iterate
                        begin
                           try
                              if CustomDocProps.Count <= x then
                              begin
                                 CustomDocProps.Add(OLEVar, False, msoPropertyTypeString, PropValues.Strings[x] ,EmptyParam);
                                 break;
                              end
                              else
                              begin
                                 try
                                    if i > CustomDocProps.Count then
                                    begin
                                       CustomDocProps.Add(OLEVar, False, msoPropertyTypeString, PropValues.Strings[x] ,EmptyParam);
                                       break;
                                    end
                                    else
                                    begin
                                       Item := CustomDocProps.Item[i];
                                       if (Item.Name = OLEVar) then
                                       begin
                                          Item.Value := Value;
                                          break;
                                       end;
                                    end;
                                 except
                                    CustomDocProps.Add(OLEVar, False, msoPropertyTypeString, PropValues.Strings[x] ,EmptyParam);
                                 end;
                              end;
                           except
                              CustomDocProps.Add(OLEVar, False, msoPropertyTypeString, Value ,'');
                           end;
                        end; // for
                     except
                        CustomDocProps.Add(OLEVar, False, msoPropertyTypeString, PropValues.Strings[x] ,EmptyParam);
                     end;
                  end;

                  // set document title property
                  Value := txtDocName.Text;
                  Item := DocProps.Item[1];
                  Item.Value := Value;

                  // add doc name to custom properties
                  Value := tmpFileName;
                  Item := CustomDocProps.Item[7];
                  Item.Value := Value;

                  MSWord.ActiveDocument.Fields.Update;
                  MSWord.ActiveDocument.Save();
               try
                  dmSaveDoc.qryMatterAttachments.FieldByName('docid').AsString := dmSaveDoc.DocID;
                  dmSaveDoc.qryMatterAttachments.FieldByName('fileid').AsString := btnEditMatter.Text;
                  dmSaveDoc.qryMatterAttachments.FieldByName('nmatter').AsInteger := nMatter;
                  dmSaveDoc.qryMatterAttachments.FieldByName('auth1').AsString := cmbAuthor.KeyValue;  //  dmSaveDoc.UserID;
                  if not FEditing then
                     dmSaveDoc.qryMatterAttachments.FieldByName('D_CREATE').AsDateTime := Now;

                  dmSaveDoc.qryMatterAttachments.FieldByName('IMAGEINDEX').AsInteger := 2;

                  dmSaveDoc.qryMatterAttachments.FieldByName('DOC_NAME').AsString := ExtractFileName(tmpFileName);
                  dmSaveDoc.qryMatterAttachments.FieldByName('DESCR').AsString := txtDocName.Text;   // ExtractFileName(tmpFileName);
                  dmSaveDoc.qryMatterAttachments.FieldByName('FILE_EXTENSION').AsString := Copy(ExtractFileExt(tmpFileName),2, Length(ExtractFileExt(tmpFileName)));
                  dmSaveDoc.qryMatterAttachments.FieldByName('NPRECCATEGORY').AsString := FPrec_Category;
                  dmSaveDoc.qryMatterAttachments.FieldByName('precedent_details').AsString := memoPrecDetails.Text;
                  dmSaveDoc.qryMatterAttachments.FieldByName('KEYWORDS').AsString := edKeywords.Text;
                  dmSaveDoc.qryMatterAttachments.FieldByName('NPRECCLASSIFICATION').AsString := FPrec_Classification;
                  if cbPortalAccess.Checked then
                     dmSaveDoc.qryMatterAttachments.FieldByName('EXTERNAL_ACCESS').AsString := 'Y'
                  else
                     dmSaveDoc.qryMatterAttachments.FieldByName('EXTERNAL_ACCESS').AsString := 'N';

                  if FEditing then
                  begin
                     dmSaveDoc.qryMatterAttachments.FieldByName('D_MODIF').AsDateTime := Now;
                     dmSaveDoc.qryMatterAttachments.FieldByName('auth2').AsString := dmSaveDoc.UserID;
                  end;

                  dmSaveDoc.qryMatterAttachments.FieldByName('PATH').AsString := IndexPath(tmpFileName, 'DOC_SHARE_PATH');
                  dmSaveDoc.qryMatterAttachments.FieldByName('display_PATH').AsString := tmpFileName;

                  dmSaveDoc.qryMatterAttachments.Post;
                  dmSaveDoc.qryMatterAttachments.ApplyUpdates;
                  dmSaveDoc.orsInsight.Commit;

               except
                  dmSaveDoc.orsInsight.Rollback;
               end;

               SaveDocument := True;
               if (not cbLeaveDocOpen.Checked) then
               begin
                  Value := False;
                  MSWord.ActiveDocument.Close(Value,EmptyParam, EmptyParam);
//                  Value := tmpFileName;
//                  MSWord.Documents.Open(Value, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
//                                         EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
               end;
               PropValues.Free;
               MSDoc := nil;
               MSWord := nil;
               ModalResult := mrOk;
            end;
         except
//           on E: Exception do
//             begin
//                Application.MessageBox(PChar('Error during saving document: ' + E.Message), PChar('Insight'), MB_ICONERROR);
//                MessageDlg('Error during saving document: ' + E.Message, mtError, [mbOK], 0);
//                SaveDocument := False;
//             end;
         end;
      except
//         on E: Exception do
//          begin
//             Application.MessageBox(PChar('Error during saving document (trying to establish active document): ' + E.Message), PChar('Insight'), MB_ICONERROR);
//             MessageDlg('Error during saving document: ' + E.Message, mtError, [mbOK], 0);
//             SaveDocument := False;
//          end;
      end;
   end;
end;

procedure TfrmSaveDocDtls.JvHTButton2Click(Sender: TObject);
begin
  Self.Close;
end;

procedure TfrmSaveDocDtls.LMDDockButton1Click(Sender: TObject);
begin
   frmMtrSearch :=TfrmMtrSearch.Create(nil);
   try
      frmMtrSearch.MakeSql;
      if (frmMtrSearch.ShowModal = mrOK) then
      begin
         btnEditMatter.Text := dmSaveDoc.qryMatters.FieldByName('fileid').AsString;
         nMatter := dmSaveDoc.qryMatters.FieldByName('nmatter').AsInteger;
         cmbAuthor.KeyValue := TableString('MATTER','NMATTER',nMatter,'AUTHOR');
         FFileID := btnEditMatter.Text;
//         cbOverwriteDoc.Enabled := (FOldFileID = FFileID);
         cbNewCopy.Visible := ((FOldFileID <> FFileID) and (FOldFileID <> ''));
      end;
   finally
      frmMtrSearch.Free;
   end;
end;

procedure TfrmSaveDocDtls.GetDetails;
var
  varWord, varDoc, PropName : OleVariant;
begin

   try
       varWord := GetActiveOleObject('Word.Application');
   except
      on EOleSysError do
      begin
         try
            varWord := CreateOleObject('Word.Application');
         except
 //           on e: Exception do
//            begin
//               MessageDlg('Error Starting MS Word: ' + E.Message, mtError, [mbOK], 0);
//               varWord := null;
//            end;
         end;
      end;
   end;

   if(not VarIsNull(varWord)) then
   begin
      try
         PropName := 'MatterNo';
         varDoc := varWord.ActiveDocument;
         FFileID := varDoc.CustomDocumentProperties[PropName].Value;
         FOldFileID := FFileID;
         btnEditMatter.Text := FFileID;
         nMatter := TableInteger('MATTER','FILEID',FFileID,'NMATTER');

         PropName := 'DocID';
         dmSaveDoc.DocID := varDoc.CustomDocumentProperties[PropName].Value;
//         application.MessageBox(pchar(FDocID),'help',MB_OK);
         FDocName := TableString('DOC','DOCID', dmSaveDoc.DocID, 'DOC_NAME');
         if FDocName = '' then
            FDocName := varWord.ActiveDocument.Name;

         cbOverWriteDoc.Visible := True;
         PropName := 'Prec_Category';
         try
            FPrec_Category := varDoc.CustomDocumentProperties[PropName].Value;
            cmbCategory.KeyValue := FPrec_Category;
         except
            ;// in case of errors
         end;

         PropName := 'Prec_Classification';
         try
            FPrec_Classification := varDoc.CustomDocumentProperties[PropName].Value;
            cmbClassification.KeyValue := FPrec_Classification;
         except
            ;// in case it doesnt exist
         end;

         PropName := 'Doc_Keywords';
         try
            FDoc_Keywords := varDoc.CustomDocumentProperties[PropName].Value;
            edKeywords.Text := FDoc_Keywords;
         except
            ;// in case it doesnt exist
         end;

         PropName := 'Doc_Precedent';
         try
            FDoc_Precedent := varDoc.CustomDocumentProperties[PropName].Value;
            memoPrecDetails.Text := FDoc_Precedent;
         except
            ;// in case it doesnt exist
         end;

         PropName := 'Doc_FileName';
         try
            FDoc_FileName := varDoc.CustomDocumentProperties[PropName].Value;
            btnTxtDocPath.Text := FDoc_FileName;
         except
            ;// in case it doesnt exist
         end;

         PropName := 'Doc_Title';
         try
            TxtDocName.Text := varDoc.CustomDocumentProperties[PropName].Value;
         except
            ;// in case it doesnt exist
         end;

         PropName := 'Doc_Author';
         try
            FDoc_Author := varDoc.CustomDocumentProperties[PropName].Value;
            cmbAuthor.KeyValue := FDoc_Author;
         except
            cmbAuthor.KeyValue := dmSaveDoc.UserID;// in case it doesnt exist
         end;

         PropName := 'Saved_in_DB';
         FSavedInDB := varDoc.CustomDocumentProperties[PropName].Value;
         if FSavedInDB = 'Y' then
         begin
//            rgStorage.ItemIndex := 0;
            btnTxtDocPath.Text := FDocName;
         end;
//         varWord.ActiveDocument.BuiltinDocumentProperties('Category') := IntToStr(nMatter);
         if txtDocName.Text = '' then
            txtDocName.Text := TableString('DOC','DOCID', dmSaveDoc.DocID, 'DESCR'); //  DocName;

         PropName := 'Portal_Access';
         cbPortalAccess.Checked := (varDoc.CustomDocumentProperties[PropName].Value = 'Y');

      except
         // in case of errors
      end;
   end;
end;


end.
