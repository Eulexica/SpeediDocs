unit PowerPointUnit;

interface


uses MSPpt2000, Office2000, SpeediDocs_IMPL, ActiveX, ComObj, Variants,
     SaveDocFunc, SysUtils, System.Classes, Windows, SaveDoc, DB, Comserv,
     adxAddin;

function SavePresentation(DocSequence: string; rgStorageItemIndex: integer; ATxtDocPath: string;
                      ANewCopy, AOverwrite: boolean; AFileID, AAuthor, ADocName: string;
                      AWorkflowType, ATemplateType: string;
                      AGroupID, APrec_Category, APrec_Classification, ADocFolder: integer;
                      APrecedentDescr, AKeywords: string;
                      ADocID: integer = 1): boolean;

implementation

var
    tmpFileName: string;
    tmpdir: string;
    FFileID: string;
    FOldFileID: string;
    FEditing: boolean;


function SavePresentation(DocSequence: string; rgStorageItemIndex: integer; ATxtDocPath: string;
                      ANewCopy, AOverwrite: boolean; AFileID, AAuthor, ADocName: string;
                      AWorkflowType, ATemplateType: string;
                      AGroupID, APrec_Category, APrec_Classification, ADocFolder: integer;
                      APrecedentDescr, AKeywords: string;
                      ADocID: integer = 1): boolean;
var
  DocName, SavedInDB: string;
  nCat, nClass: integer;
  ltmpdir, AMacro: string;
  MSPPoint: _Application;
  MSDoc: _Presentation;
  Unknown: IUnknown;
  OLEResult: HResult;
  OLEvar,
  OLEFileType,
  OLETrueType: OleVariant;
  CustomDocProps, Item, Value, DocProps: OleVariant;
  i, x: integer;
  lDocID, AExt: string;
  PropValues: TStrings;
  bMoveSuccess: boolean;
begin
   SavePresentation := False;
   bMoveSuccess := True;

   OLEResult := GetActiveObject(CLASS_PowerPointApplication , nil, Unknown);
   if (OLEResult = MK_E_UNAVAILABLE) then
      MSPPoint := CoPowerPointApplication.Create          //get MS PowerPoint running
   else
   begin
      OleCheck(OLEResult);                           //check for errors
      OleCheck(Unknown.QueryInterface(_Application, MSPPoint));
   end;


   if(not VarIsNull(MSPPoint)) then
   begin
      try
         case rgStorageItemIndex of
           0:  begin
                  ltmpdir := ParseMacros(tmpFileName,TableInteger('MATTER','FILEID',FFileID,'NMATTER'),
                                         StrToInt(DocSequence), ADocName);
                  ltmpDir := tmpdir+ExtractFileName(ltmpdir);  // copy(ltmpDir, 1,length(ltmpdir) - 1);
//                 if not DirectoryExists(ltmpdir) then
//                    ForceDirectories(ltmpdir);

                  if ExtractFileExt(ltmpdir) = '' then
                     ltmpdir := ltmpdir + '.pptx';

                  Value := ltmpdir;
                  Item := CustomDocProps.Item[7];
                  Item.Value := Value;

                  OLETrueType := True;
                  MSPPoint.ActivePresentation.SaveAs(Value, ppSaveAsPresentation, OLETrueType);
                  tmpFileName := ltmpdir;
               end;
           1:  begin
                  if (FOldFileID <> AFileID) and (FOldFileID <> '') and (not ANewCopy) then
                  begin
                     tmpFileName := SystemString('DRAG_DEFAULT_DIRECTORY');
                     tmpFileName := IncludeTrailingPathDelimiter(tmpFileName) + ExtractFileName(ATxtDocPath);

                     AExt := ExtractFileExt(tmpFileName);
                     tmpFileName := Copy (tmpFileName,1, Length(tmpFileName)- Length(AExt));
                     tmpFileName := tmpFileName + '_' + '[DOCSEQUENCE]';
                     tmpFileName := tmpFileName + AExt;

                     tmpFileName := ParseMacros(tmpFileName,
                                    TableInteger('MATTER','FILEID',uppercase(AFileID),'NMATTER'),
                                    StrToInt(DocSequence), ADocName);

                     if FOldFileID <> '' then
                        bMoveSuccess := MoveMatterDoc(tmpFileName, ATxtDocPath);
                  end
                  else
                  if (not FEditing) then
                  begin
                     if ATxtDocPath = '' then
                     begin
                        ATxtDocPath := SystemString('DOC_DEFAULT_DIRECTORY');
//                        btnTxtDocPath.Text
                     end;
                     tmpFileName := ATxtDocPath;

                     AExt := ExtractFileExt(tmpFileName);
                     tmpFileName := Copy (tmpFileName,1, Length(tmpFileName)- Length(AExt));
                     tmpFileName := tmpFileName + '_[DOCSEQUENCE]';
                     tmpFileName := tmpFileName + AExt;

                     tmpFileName := ParseMacros(tmpFileName,
                                    TableInteger('MATTER','FILEID',AFileID,'NMATTER'),
                                    StrToInt(DocSequence), ADocName);
                  end
                  else
                  begin
                     tmpFileName := tmpFileName;
                  end;

                  if ExtractFileName(tmpFileName) = '' then
                     tmpFileName  := tmpFileName + AFileID;

                  if (MSPPoint.Version <> '12.0') and (MSPPoint.Version <> '14.0') then
                  begin
                    AExt := '.ppt';
                    OLEFileType := ppSaveAsPowerPoint4;
                  end
                  else
                  begin
 //                    If (MSPPoint.ActivePresentation.VBProject.Name <> '') Then
 //                    begin
//                        AExt := '.pptm';
//                        OLEFileType := ppSaveAsPresentation;
//                     end
//                     Else
//                     begin
                        AExt := '.pptx';
                        OLEFileType := ppSaveAsDefault;
//                     end;
                  end;
                  if ExtractFileExt(tmpFileName) = '' then
                     tmpFileName := tmpFileName + AExt;

                  if ((DocName = '') or (pos('Document', DocName) > 0) or
                     (ExtractFileName(ATxtDocPath) <> DocName)) and (AOverwrite = False) then
                  begin
                     if not DirectoryExists(ExtractFileDir(tmpFileName)) then
                        ForceDirectories(ExtractFileDir(tmpFileName));
                     Value := tmpFileName;
                     OLETrueType := True;

                     MSPPoint.ActivePresentation.SaveAs(Value, OLEFileType, OLETrueType);
                  end
                  else
                  begin
                     MSPPoint.ActivePresentation.Save;
                  end;
               end;
         end;

               try
                  dmConnection.qryMatterAttachments.FieldByName('docid').AsInteger := ADocID;
                  dmConnection.qryMatterAttachments.FieldByName('fileid').AsString := AFileid;
                  dmConnection.qryMatterAttachments.FieldByName('nmatter').AsInteger := TableInteger('MATTER','FILEID',AFileID,'nMatter');
                  dmConnection.qryMatterAttachments.FieldByName('auth1').AsString := AAuthor;  //  dmSaveDoc.UserID;
                  if not FEditing then
                     dmConnection.qryMatterAttachments.FieldByName('D_CREATE').AsDateTime := Now;

                  dmConnection.qryMatterAttachments.FieldByName('IMAGEINDEX').AsInteger := 8;
                  if rgStorageItemIndex = 0 then
                     dmConnection.qryMatterAttachments.FieldByName('DOC_NAME').AsString := ExtractFileName(tmpFileName)  //txtDocName.Text + '.doc'
                  else
                     dmConnection.qryMatterAttachments.FieldByName('DOC_NAME').AsString := ExtractFileName(tmpFileName);
                  dmConnection.qryMatterAttachments.FieldByName('DESCR').AsString := ADocName;   // ExtractFileName(tmpFileName);
                  dmConnection.qryMatterAttachments.FieldByName('SEARCH').AsString := ADocName;
                  dmConnection.qryMatterAttachments.FieldByName('FILE_EXTENSION').AsString := Copy(ExtractFileExt(tmpFileName),2, Length(ExtractFileExt(tmpFileName)));
                  dmConnection.qryMatterAttachments.FieldByName('precedent_details').AsString := ADocName;
                  dmConnection.qryMatterAttachments.FieldByName('KEYWORDS').AsString := AKeywords;
                  if APrec_Category > -1 then
                     dmConnection.qryMatterAttachments.FieldByName('NPRECCATEGORY').AsInteger := APrec_Category;
                  if APrec_Classification > -1 then
                     dmConnection.qryMatterAttachments.FieldByName('NPRECCLASSIFICATION').AsInteger := APrec_Classification;
                  if ADocFolder > -1 then
                     dmConnection.qryMatterAttachments.FieldByName('FOLDER_ID').AsInteger := ADocFolder;

//                  if cbPortalAccess.Checked then
//                     dmSaveDoc.qryMatterAttachments.FieldByName('EXTERNAL_ACCESS').AsString := 'Y'
//                  else
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

                  dmConnection.qryMatterAttachments.Post;
                  dmConnection.qryMatterAttachments.ApplyUpdates;
                  dmConnection.orsInsight.Commit;

               except
                  dmConnection.orsInsight.Rollback;
               end;
               SavePresentation := True;
      except
         on E: Exception do
          begin
             SavePresentation := False;
          end;
      end;
   end;
end;


end.
