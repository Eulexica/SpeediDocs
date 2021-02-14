unit LoginDetails;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Registry, SaveDoc, SpeediDocs_IMPL, Vcl.ComCtrls, SaveDocFunc,
  Vcl.Buttons, AdditionalInfo, cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, cxContainer, cxEdit, cxCheckBox, dxSkinsCore,
  dxSkinLiquidSky, dxSkinOffice2019Colorful;

type
  TfrmLoginSetup = class(TForm)
    Button1: TButton;
    btnCancel: TButton;
    StatusBar: TStatusBar;
    BitBtn1: TBitBtn;
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    edUserName: TEdit;
    edPassword: TEdit;
    Database: TGroupBox;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    edServerName: TEdit;
    edDatabase: TEdit;
    edPort: TEdit;
    chkUseDirectConn: TcxCheckBox;
    procedure Button1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure chkUseDirectConnClick(Sender: TObject);
  private
    { Private declarations }
     LRegAxiom: TRegistry;
//     sRegistryRoot: string;
      FisOutlook: boolean;
  public
    { Public declarations }
    property isOutlook: boolean read FisOutlook write FisOutlook;
  end;

var
  frmLoginSetup: TfrmLoginSetup;

implementation


{$R *.dfm}

procedure TfrmLoginSetup.BitBtn1Click(Sender: TObject);
var
   frmAdditionalInfo : TfrmAdditionalInfo;
begin
   try
      frmAdditionalInfo := TfrmAdditionalInfo.Create(nil);
      frmAdditionalInfo.IsOutlook := Self.IsOutlook;
      frmAdditionalInfo.ShowModal;
   finally
      frmAdditionalInfo.Free;
      frmAdditionalInfo := nil;
   end;

end;

procedure TfrmLoginSetup.Button1Click(Sender: TObject);
var
  regAxiom: TRegistry;
begin
   regAxiom := TRegistry.Create;
   try
      regAxiom.RootKey := HKEY_CURRENT_USER;
      if regAxiom.OpenKey(csRegistryRoot, True) then
      begin
         if chkUseDirectConn.Checked = True then
         begin
            regAxiom.WriteString('Net','Y');
            regAxiom.WriteString('Server Name',edServerName.Text+':'+edPort.Text+':'+edDatabase.Text);
            regAxiom.WriteString('User Name',edUserName.Text);
            regAxiom.WriteString('Password',edPassword.Text);
         end
         else
         begin
            regAxiom.WriteString('Net','N');
            regAxiom.WriteString('Server Name',edDatabase.Text);
            regAxiom.WriteString('User Name',edUserName.Text);
            regAxiom.WriteString('Password',edPassword.Text);
         end;
         regAxiom.CloseKey;
      end;

      if chkUseDirectConn.Checked = True then
      begin
         dmConnection.orsInsight.Options.Direct := True;
         dmConnection.orsInsight.Server := edServerName.Text+':'+edPort.Text+':'+edDatabase.Text;
      end
      else
      begin
         dmConnection.orsInsight.Options.Direct := False;
         dmConnection.orsInsight.Server := edDatabase.Text;
      end;

   finally
      regAxiom.Free;
   end;
   Close;
end;

procedure TfrmLoginSetup.chkUseDirectConnClick(Sender: TObject);
begin
   edServerName.Enabled := chkUseDirectConn.Checked;
   edPort.Enabled := chkUseDirectConn.Checked;
end;

procedure TfrmLoginSetup.FormShow(Sender: TObject);
var
   LoginStr, s: string;
begin
   StatusBar.Panels[0].Text := 'Ver: '+ ReportVersion(SysUtils.GetModuleName(HInstance)) + ' (' +DateTimeToStr(FileDateToDateTime(FileAge(SysUtils.GetModuleName(HInstance))))+')';

   LregAxiom := TRegistry.Create;
   try
      LregAxiom.RootKey := HKEY_CURRENT_USER;
      LregAxiom.OpenKey(csRegistryRoot, False);

      s := Copy(LregAxiom.ReadString('Server Name'),1,Pos(':',LregAxiom.ReadString('Server Name'))-1);
      LoginStr := Copy(LregAxiom.ReadString('Server Name'),Pos(':',LregAxiom.ReadString('Server Name'))+1, Length(LregAxiom.ReadString('Server Name')) - Pos(':',LregAxiom.ReadString('Server Name')) );
      if (LregAxiom.ReadString('Net') = 'Y') then
      begin
         if s <> '' then
            edServerName.Text := s;

         s := Copy(LoginStr,1,Pos(':',LoginStr)-1);
         LoginStr := Copy(LoginStr,Pos(':',LoginStr)+1, Length(LoginStr));
         if s <> '' then
            edPort.Text := s;

         s := LoginStr;
         if s <> '' then
            edDatabase.Text := s;
      end
      else
      begin
         edDatabase.Text := LoginStr;
         edServerName.Enabled := False;
         edPort.Enabled := False;
      end;

      edUserName.Text := LregAxiom.ReadString('User Name');
      edPassword.Text := LregAxiom.ReadString('Password');
      chkUseDirectConn.Checked := (LregAxiom.ReadString('Net') = 'Y');
   finally
     LregAxiom.Free;
   end;
end;

end.
