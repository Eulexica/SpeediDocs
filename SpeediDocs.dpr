library SpeediDocs;



uses
  ComServ,
  SpeediDocs_TLB in 'SpeediDocs_TLB.pas',
  SpeediDocs_IMPL in 'SpeediDocs_IMPL.pas' {AddInModule: TAddInModule} {coSpeediDocs: CoClass},
  LoginDetails in 'LoginDetails.pas' {frmLoginSetup},
  MatterSearch in 'MatterSearch.pas' {frmMtrSearch},
  SaveDocFunc in 'SaveDocFunc.pas',
  PowerPointUnit in 'PowerPointUnit.pas',
  ExcelUnit in 'ExcelUnit.pas',
  WordUnit in 'WordUnit.pas',
  NewFee in 'NewFee.pas' {frmNewFee},
  SaveDoc in 'SaveDoc.pas' {dmSaveDoc: TDataModule},
  SavedocDetails in 'SavedocDetails.pas' {frmSaveDocDetails},
  FieldList in 'FieldList.pas' {frmFieldList},
  Matters in 'Matters.pas' {adxfrmMatters: TadxOlForm},
  DocList in 'DocList.pas' {frmDocList},
  SaveprecDetails in 'SaveprecDetails.pas' {frmSavePrecDetails};

exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer;

{$R *.TLB}

{$R *.RES}

begin
end.
