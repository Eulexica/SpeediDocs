unit WordUnit;

interface


uses Word_TLB, Office_TLB, SpeediDocs_IMPL, ActiveX, ComObj, Variants,
     SaveDocFunc, SysUtils, SaveDoc, DB, System.Classes, savedocdetails,
     vcl.Forms, Windows, vcl.dialogs,RegularExpressions;

   function SaveDocument(DocSequence: string; rgStorageItemIndex: integer; ATxtDocPath: string;
                         ANewCopy, AOverwrite: boolean; AFileID, AAuthor, ADocName: string;
                         AWorkflowType, ATemplateType: string;
                         AGroupID, APrec_Category, APrec_Classification, ADocFolder: integer;
                         APrecedentDescr, AKeywords: string;
                         ADocOpen: boolean = TRUE; ADocID: integer = -1; ASaveAsPrecedent: boolean = False;
                         LadxLCID: integer = 0; ATask: string = ''; ASaveTime: boolean = False;
                         ATimeUnits: integer = 1; ATimeNarration: string = ''): boolean;
   procedure SetWordApp(WordApp: _Application);
   function GetWordApp(): _Application;
//   procedure GetDetails;

implementation

var
    tmpFileName: string;
    tmpdir: string;
    FFileID: string;
    FOldFileID: string;
    FEditing: boolean;
    MSWord: _Application;
    AWordProps: array[1..11] of TWordProperties;

function SaveDocument(DocSequence: string; rgStorageItemIndex: integer; ATxtDocPath: string;
                      ANewCopy, AOverwrite: boolean; AFileID, AAuthor, ADocName: string;
                      AWorkflowType, ATemplateType: string;
                      AGroupID, APrec_Category, APrec_Classification, ADocFolder: integer;
                      APrecedentDescr, AKeywords: string;
                      ADocOpen: boolean; ADocID: integer; ASaveAsPrecedent: boolean;
                      LadxLCID: integer; ATask: string; ASaveTime: boolean; ATimeUnits: integer;
                      ATimeNarration: string): boolean;
var
  DocName, SavedInDB: string;
  nCat, nClass: integer;
  ltmpdir, AMacro: string;
//  MSWord: _Application;
  MSDoc: _Document;
  Unknown: IUnknown;
  OLEResult: HResult;
  OLEvar: OleVariant;
  CustomDocProps,
  Item,
  Value,
  DocProps,
  bAddToRecentFiles,
  nFileFormat: OleVariant;
  i, x,
  nRet,
  count,
  lTimeUnits,
  LNMatter,
  LnDocID: integer;
  LDocID, AExt,
  lTask: string;
  PropValues: TStrings;
  bMoveSuccess: boolean;
  bDocSeqAppend: boolean;
  IProps: DocumentProperties;
  IProp: DocumentProperty;
  PropValue: OleVariant;
  PropName, PrecPath,
  tmpDocName: widestring;
  AFileName: string;
  lRate,
  lAmount: double;
begin
   SaveDocument := False;
   bMoveSuccess := True;

   LNMatter := TableInteger('MATTER','FILEID',uppercase(AFileID),'NMATTER');
   bDocSeqAppend := (SystemString('DOC_SEQ_APPEND') = 'Y');
   try
      case rgStorageItemIndex of
        0:  begin
               ltmpdir := ParseMacros(tmpFileName,
                                      LNMatter,
                                      StrToInt(DocSequence), ADocName);
               ltmpDir := tmpdir+ExtractFileName(ltmpdir);  // copy(ltmpDir, 1,length(ltmpdir) - 1);
//              if not DirectoryExists(ltmpdir) then
//                 ForceDirectories(ltmpdir);

               if ExtractFileExt(ltmpdir) = '' then
                  ltmpdir := ltmpdir + '.doc';

               Value := ltmpdir;
               Item := CustomDocProps.Item[7];
               Item.Value := Value;

               MSWord.ActiveDocument.SaveAs(Value, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
                                             EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
                                             EmptyParam, EmptyParam, EmptyParam, EmptyParam);
               tmpFileName := ltmpdir;
            end;
        1:  begin
               if ASaveAsPrecedent = True then
               begin
                  // DW 12 Jul 2018 added invaild character stripout
                  AFileName := ADocName;
                  for x := 1 to length(AFileName) do
                  begin
                  if (AFileName[x] in ['/', '\', '?','"','<','>','|','*',':', '.']) then
                      AFileName[x] := ' ';
                  end;

                  tmpFileName := SystemString('DFLT_PRECEDENT_PATH');
                  tmpFileName := IncludeTrailingPathDelimiter(tmpFileName) + AFileName + '.' + SystemString('DEFAULT_DOC_EXT');

                  if not DirectoryExists(ExtractFileDir(tmpFileName)) then
                       ForceDirectories(ExtractFileDir(tmpFileName));
                     Value := tmpFileName;
                   try
                     MSWord.ActiveDocument.SaveAs(Value, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
                                                 EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
                                                 EmptyParam, EmptyParam, EmptyParam, EmptyParam);
                   except
                   //
                   end;
               end
               else
               begin
//                 tmpDocName := MSWord.ActiveDocument.Name;
                  tmpDocName := ADocName;
                  // if no document name returned from document or passed documentname = document name make the docname the same
//                 if (tmpDocName = '') or (tmpDocName = ADocName) or
//                    TRegEx.IsMatch(tmpDocName, 'Document.') then tmpDocName := ADocName;

                  AExt := ExtractFileExt(tmpDocName);

                  if (FOldFileID <> AFileID) and (FOldFileID <> '') and (not ANewCopy) then
                  begin
                     tmpFileName := SystemString('DOC_DEFAULT_DIRECTORY');
                     tmpFileName := IncludeTrailingPathDelimiter(tmpFileName) + ExtractFileName(ATxtDocPath);

                     if (ExtractFileName(ATxtDocPath) = '') then
                        tmpFileName := ADocName + '.' + SystemString('DEFAULT_DOC_EXT');

                     AExt := ExtractFileExt(tmpFileName);
                     tmpFileName := Copy (tmpFileName,1, Length(tmpFileName)- Length(AExt));
                     if (SystemString('DOC_SEQ_APPEND') = 'Y') then
                         tmpFileName := tmpFileName + '_' + '[DOCSEQUENCE]';
                     tmpFileName := tmpFileName + AExt;

                     tmpFileName := ParseMacros(tmpFileName, LNMatter, ADocID, ADocName);

                     if FOldFileID <> '' then
                        bMoveSuccess := MoveMatterDoc(tmpFileName, ATxtDocPath);
                  end
                  else
                  if (not FEditing) then
                  begin
                     if ATxtDocPath = '' then
                     begin
                       ATxtDocPath := SystemString('DOC_DEFAULT_DIRECTORY');
//                       btnTxtDocPath.Text
                     end;
                     tmpFileName := ATxtDocPath; //IncludeTrailingPathDelimiter(ATxtDocPath);

                     AExt := ExtractFileExt(tmpFileName);
                     tmpFileName := Copy (tmpFileName,1, Length(tmpFileName)- Length(AExt));
                     if (SystemString('DOC_SEQ_APPEND') = 'Y') then
                        tmpFileName := tmpFileName + '_[DOCSEQUENCE]';
                     tmpFileName := tmpFileName + AExt;

                     tmpFileName := ParseMacros(tmpFileName,
                                   LNMatter,
                                   ADocID, ADocName);
                  end
                  else
                  begin
                     tmpFileName := tmpFileName;
                  end;

                  if ExtractFileName(tmpFileName) = '' then
                    tmpFileName  := tmpFileName + AFileID;

                  if (ExtractFileExt(tmpFileName) = '') then
                  begin
                     if (StrToFloat(MSWord.Version) < 12) then
                        tmpFileName := tmpFileName + '.' + SystemString('default_doc_ext')
                     else
                        tmpFileName := tmpFileName + '.' + 'docx';
                  end;

                  //*********************************
                  GetWordApp.ActiveDocument.CustomDocumentProperties.QueryInterface(IID_DocumentProperties, IProps);
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
                              IProp.Delete;
                           finally
                              IProp := nil;
                           end;
                        end;
                        GetWordApp.ActiveDocument.Fields.Update;
                     end;

                     begin
                        try
                           FFileID := AFileID;
                           PropName := 'MatterNo';
                           IProps.Add(PropName, False, msoPropertyTypeString, FFileID, EmptyParam, LadxLCID, IProp);

                           LDocID := IntToStr(ADocID);
                           PropName := 'DocID';
                           IProps.Add(PropName, False, msoPropertyTypeString, LDocID, EmptyParam, LadxLCID, IProp);

                           if varIsNull(APrec_Category) or (VarToStr(APrec_Category) = '') then
                              nCat := -1
                           else
                           begin
                              try
                                 nCat := APrec_Category;
                              except
                                 nCat := -1;
                              end;
                           end;
                           PropName := 'Prec_Category';
                           IProps.Add(PropName, False, msoPropertyTypeString, IntToStr(nCat), EmptyParam, LadxLCID, IProp);

                           if varIsNull(APrec_Classification) or (VarToStr(APrec_Classification) = '') then
                              nClass := -1
                           else
                           begin
                              try
                                 nClass := APrec_Classification;
                              except
                                 nClass := -1;
                              end;
                           end;
                           PropName := 'Prec_Classification';
                           IProps.Add(PropName, False, msoPropertyTypeString, IntToStr(nClass), EmptyParam, LadxLCID, IProp);

                           PropName := 'Doc_Keywords';
                           IProps.Add(PropName, False, msoPropertyTypeString, AKeyWords, EmptyParam, LadxLCID, IProp);

                           PropName := 'Doc_Precedent';
                           IProps.Add(PropName, False, msoPropertyTypeString, ExtractFileName(tmpFileName{ATxtDocPath}), EmptyParam, LadxLCID, IProp);


                           // add author
                           PropName := 'Doc_Author';
                           IProps.Add(PropName, False, msoPropertyTypeString, AAuthor, EmptyParam, LadxLCID, IProp);

{                          case rgStorage.ItemIndex of
                             0: SavedInDB := 'Y';
                             1: SavedInDB := 'N';
                           end; }
                           SavedInDB := 'N';
                           PropName := 'Saved_in_DB';
                           IProps.Add(PropName, False, msoPropertyTypeString, SavedInDB, EmptyParam, LadxLCID, IProp);


                           // document description - title
                           PropName := 'Doc_FileName';
                           IProps.Add(PropName, False, msoPropertyTypeString, ADocName, EmptyParam, LadxLCID, IProp);


 {                         if cbPortalAccess.Checked then
                             PropValues.Add('Y')
                          else
                             PropValues.Add('N');  }

                           PropName := 'Portal_Access';
                           PropValue := 'N';
                           IProps.Add(PropName, False, msoPropertyTypeString, PropValue, EmptyParam, LadxLCID, IProp);
                        finally
                           GetWordApp.ActiveDocument.Fields.Update;
//                          GetWordApp.ActiveDocument.Save();
                           IProp := nil;
                        end;
                     end;
                  finally
                     IProps := nil;
                  end;

            //****************************************************************************

                  if ((DocName = '') or (pos('Document', DocName) > 0) or
                    (ExtractFileName(ATxtDocPath) <> DocName)) and (AOverwrite = False) then
                  begin
                     if not DirectoryExists(ExtractFileDir(tmpFileName)) then
                       ForceDirectories(ExtractFileDir(tmpFileName));

                     Value := '"'+tmpFileName+'"';
//                   Value := tmpFileName;
                     bAddToRecentFiles := True;
                     if UpperCase(ExtractFileExt(tmpFileName)) = '.DOCX' then
                        nFileFormat := 16
                     else
                        nFileFormat := 0;

                     try
                        MSWord.ActiveDocument.SaveAs(Value, nFileFormat, EmptyParam, EmptyParam, bAddToRecentFiles, EmptyParam,
                                                 EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
                                                 EmptyParam, EmptyParam, EmptyParam, EmptyParam);
                     finally
                        if (not FileExists(tmpFileName)) then
                        begin
                           try
                              Sleep(20);
                              MSWord.ActiveDocument.SaveAs(Value, nFileFormat, EmptyParam, EmptyParam, bAddToRecentFiles, EmptyParam,
                                                 EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
                                                 EmptyParam, EmptyParam, EmptyParam, EmptyParam);
                           finally
                              MessageDlg('Document: '+ tmpFileName + ' was not saved to disk.'+chr(13)+
                                    'Please check the file name and path and try saving the document again', mtError, [mbOK], 0);
                           end;
                        end;
                     //
                     end;
                  end
                  else
                     MSWord.ActiveDocument.Save;

                  AMacro := SystemString('WORD_SAVE_MACRO');
                  try
                    if AMacro <> '' then MSWord.Run(AMacro, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
                                                EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
                                                EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
                                                EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
                                                EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
                                                EmptyParam );
                  except
                     MessageDlg('The Macro: '+ AMacro + ' has caused an error.', mtError, [mbOK], 0);
                  end;
               end;

//         Value := False;
//         MSWord.ActiveDocument.Close(Value,EmptyParam, EmptyParam);
//           MSWord.Activewindow.WindowState := wdWindowStateMaximize;

           if ASaveAsPrecedent = False then
           begin
              try
                LnDocID := dmConnection.DocID;
                dmConnection.qryMatterAttachments.FieldByName('docid').AsInteger := LnDocID;
                dmConnection.qryMatterAttachments.FieldByName('fileid').AsString := AFileid;
                dmConnection.qryMatterAttachments.FieldByName('nmatter').AsInteger := TableInteger('MATTER','FILEID',AFileID,'nMatter');
                dmConnection.qryMatterAttachments.FieldByName('auth1').AsString := AAuthor;  //  dmSaveDoc.UserID;
                if not FEditing then
                   dmConnection.qryMatterAttachments.FieldByName('D_CREATE').AsDateTime := Now;

                dmConnection.qryMatterAttachments.FieldByName('IMAGEINDEX').AsInteger := 2;
                if rgStorageItemIndex = 0 then
                   dmConnection.qryMatterAttachments.FieldByName('DOC_NAME').AsString := ExtractFileName(tmpFileName)  //txtDocName.Text + '.doc'
                else
                   dmConnection.qryMatterAttachments.FieldByName('DOC_NAME').AsString := ExtractFileName(tmpFileName);
                dmConnection.qryMatterAttachments.FieldByName('DESCR').AsString := ADocName;   // ExtractFileName(tmpFileName);
                dmConnection.qryMatterAttachments.FieldByName('SEARCH').AsString := ADocName;
                dmConnection.qryMatterAttachments.FieldByName('FILE_EXTENSION').AsString := Copy(ExtractFileExt(tmpFileName),2, Length(ExtractFileExt(tmpFileName)));
                dmConnection.qryMatterAttachments.FieldByName('precedent_details').AsString := ADocName;
                dmConnection.qryMatterAttachments.FieldByName('KEYWORDS').AsString := AKeywords;
//                dmSaveDoc.qryMatterAttachments.FieldByName('PARENTDOCID').AsInteger := LnDocID;
                if APrec_Category > -1 then
                   dmConnection.qryMatterAttachments.FieldByName('NPRECCATEGORY').AsInteger := APrec_Category;
                if APrec_Classification > -1 then
                   dmConnection.qryMatterAttachments.FieldByName('NPRECCLASSIFICATION').AsInteger := APrec_Classification;

                if ADocFolder > -1 then
                   dmConnection.qryMatterAttachments.FieldByName('FOLDER_ID').AsInteger := ADocFolder;
//                if cbPortalAccess.Checked then
//                   dmConnection.qryMatterAttachments.FieldByName('EXTERNAL_ACCESS').AsString := 'Y'
//                else
                   dmConnection.qryMatterAttachments.FieldByName('EXTERNAL_ACCESS').AsString := 'N';

                if FEditing then
                begin
                   dmConnection.qryMatterAttachments.FieldByName('D_MODIF').AsDateTime := Now;
                   dmConnection.qryMatterAttachments.FieldByName('auth2').AsString := dmConnection.UserCode;
                end;
                if rgStorageItemIndex = 0 then
                begin
                   TBlobField(dmConnection.qryMatterAttachments.fieldByname('DOCUMENT')).LoadFromFile(tmpFileName);
                end
                else
                begin
                   dmConnection.qryMatterAttachments.FieldByName('PATH').AsString := IndexPath(tmpFileName, 'DOC_SHARE_PATH');
                   dmConnection.qryMatterAttachments.FieldByName('display_PATH').AsString := tmpFileName;
                end;

//                dmSaveDoc.qryMatterAttachments.Post;
                try
                   dmConnection.qryMatterAttachments.ApplyUpdates;
                finally
                   dmConnection.orsInsight.Commit;
                   dmConnection.qryMatterAttachments.CommitUpdates;
                end;
              except
                dmConnection.orsInsight.Rollback;
              end;
           end
           else
           begin
             try
                dmConnection.qryDoctemplate.FieldByName('DOCID').AsInteger := dmConnection.PrecID;
                dmConnection.qryDoctemplate.FieldByName('DOCTYPE').AsString := 'O';  //  dmSaveDoc.UserID;
                dmConnection.qryDoctemplate.FieldByName('DOCUMENTNAME').AsString := ExtractFileName(tmpFileName);
                dmConnection.qryDoctemplate.FieldByName('DOCUMENTPATH').AsString := SystemString('DOC_DEFAULT_DIRECTORY');
                dmConnection.qryDoctemplate.FieldByName('TEMPLATEPATH').AsString := IndexPath(tmpFileName, 'DOC_SHARE_PATH');
                dmConnection.qryDoctemplate.FieldByName('DATAFILEPATH').AsString := SystemString('DFLT_MERGE_DATA_PATH');
                dmConnection.qryDoctemplate.FieldByName('WORKFLOWTYPECODE').AsString := AWorkflowType;
                if (AGroupID > -1) then
                    dmConnection.qryDoctemplate.FieldByName('GROUPID').AsInteger := AGroupID;
                dmConnection.qryDoctemplate.FieldByName('REFERREDOPTIONAL').AsString := 'N';
                dmConnection.qryDoctemplate.FieldByName('WORKFLOW_ONLY').AsString := 'N';
                dmConnection.qryDoctemplate.FieldByName('ACTIVE').AsString := 'Y';
                if (APrec_Category > -1) then
                   dmConnection.qryDoctemplate.FieldByName('NPRECCATEGORY').AsInteger := APrec_Category;
                if (APrec_Classification > -1) then
                   dmConnection.qryDoctemplate.FieldByName('NPRECCLASSIFICATION').AsInteger := APrec_Classification;
                if (ADocFolder > -1) then
                   dmConnection.qryDoctemplate.FieldByName('FOLDER_ID').AsInteger := ADocFolder;
                dmConnection.qryDoctemplate.FieldByName('IMANAGE_DOC').AsString := 'N';
                dmConnection.qryDoctemplate.FieldByName('DESCRIPTION').AsString := APrecedentDescr;   // ExtractFileName(tmpFileName);
                dmConnection.qryDoctemplate.FieldByName('TEMPLATETYPE').AsString := ATemplateType;
                try
                   dmConnection.qryDoctemplate.ApplyUpdates;
                finally
                   dmConnection.orsInsight.Commit;
                   dmConnection.qryDoctemplate.CommitUpdates;
                end;
              except
                dmConnection.orsInsight.Rollback;
              end;
           end;

           if ASaveTime then
           begin
               lTimeUnits := dmConnection.SystemInteger('TIME_UNITS');
               if ATask <> '' then
                  lTask := ATask
               else
                  lTask := dmConnection.SystemString('DFLT_EMAIL_TASK');

               lRate := CalcRate(AAuthor, lTask, Now, AFileID);
               lAmount := ATimeUnits * lRate / (60 / lTimeUnits);
               FeeTmpInsert(LNMatter, AAuthor, ATimeNarration, lAmount, lTask, ATimeUnits,
                            (ATimeUnits*lTimeUnits), lRate, 'GST');
           end;

           SaveDocument := True;
//           MSWord.ActiveDocument.Save;
           if (ADocOpen = FALSE) then
           begin
             Value := False;
             MSWord.ActiveDocument.Close(Value,EmptyParam, EmptyParam);
//         Value := tmpFileName;
//         MSWord.Documents.Open(Value, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
//                               EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
           end;
           PropValues.Free;
           MSDoc := nil;
           MSWord := nil;
        end;
      end;
   except
      on E: Exception do
      begin
//       Application.MessageBox(PChar('Error during saving document (trying to establish active document): ' + E.Message), PChar('Insight'), MB_ICONERROR);
//         MessageDlg('Error during saving document: ' + E.Message, mtError, [mbOK], 0);
         SaveDocument := False;
      end;
   end;
end;


procedure SetWordApp(WordApp: _Application);
begin
   MSWord := WordApp;
end;

function GetWordApp(): _Application;
begin
   Result := MSWord;
end;


end.
