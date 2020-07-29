unit GlobalPropertyPage;

{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ActiveX, AxCtrls, SpeediDocs_TLB, SpeediDocs_IMPL, StdVcl,
  Outlook2000, StdCtrls;

type
  TGlobalPropertyPage = class(TActiveForm, IGlobalPropertyPage, PropertyPage)
    Label1: TLabel;
    procedure Edit1Change(Sender: TObject);
  private
    FDirty: WordBool;
    FPropertyPageSite: PropertyPageSite;
    FEvents: IGlobalPropertyPageEvents;
    procedure ActivateEvent(Sender: TObject);
    procedure ClickEvent(Sender: TObject);
    procedure CreateEvent(Sender: TObject);
    procedure DblClickEvent(Sender: TObject);
    procedure DeactivateEvent(Sender: TObject);
    procedure DestroyEvent(Sender: TObject);
    procedure KeyPressEvent(Sender: TObject; var Key: Char);
    procedure PaintEvent(Sender: TObject);
  protected
    procedure DefinePropertyPages(DefinePropertyPage: TDefinePropertyPage); override;
    procedure EventSinkChanged(const EventSink: IUnknown); override;
    function Get_Active: WordBool; safecall;
    function Get_AutoScroll: WordBool; safecall;
    function Get_AutoSize: WordBool; safecall;
    function Get_AxBorderStyle: TxActiveFormBorderStyle; safecall;
    function Get_Caption: WideString; safecall;
    function Get_Color: Integer; safecall;
    function Get_Cursor: Smallint; safecall;
    function Get_DoubleBuffered: WordBool; safecall;
    function Get_DropTarget: WordBool; safecall;
    function Get_Enabled: WordBool; safecall;
    function Get_HelpFile: WideString; safecall;
    function Get_KeyPreview: WordBool; safecall;
    function Get_PixelsPerInch: Integer; safecall;
    function Get_PrintScale: TxPrintScale; safecall;
    function Get_Scaled: WordBool; safecall;
    function Get_Visible: WordBool; safecall;
    function Get_VisibleDockClientCount: Integer; safecall;
    procedure Set_AutoScroll(Value: WordBool); safecall;
    procedure Set_AutoSize(Value: WordBool); safecall;
    procedure Set_AxBorderStyle(Value: TxActiveFormBorderStyle); safecall;
    procedure Set_Caption(const Value: WideString); safecall;
    procedure Set_Color(Value: Integer); safecall;
    procedure Set_Cursor(Value: Smallint); safecall;
    procedure Set_DoubleBuffered(Value: WordBool); safecall;
    procedure Set_DropTarget(Value: WordBool); safecall;
    procedure Set_Enabled(Value: WordBool); safecall;
    procedure Set_HelpFile(const Value: WideString); safecall;
    procedure Set_KeyPreview(Value: WordBool); safecall;
    procedure Set_PixelsPerInch(Value: Integer); safecall;
    procedure Set_PrintScale(Value: TxPrintScale); safecall;
    procedure Set_Scaled(Value: WordBool); safecall;
    procedure Set_Visible(Value: WordBool); safecall;
    { Outlook PropertyPage }
    function GetPageInfo(var HelpFile: WideString;
      var HelpContext: Integer): HResult; stdcall;
    function Get_Dirty(out Dirty: WordBool): HResult; stdcall;
    function Apply: HResult; stdcall;
  public
    procedure Initialize; override;
    destructor Destroy; override;
    //
    procedure GetPropertyPageSite;
    procedure UpdatePropertyPageSite;
  end;

implementation

uses ComObj, ComServ;

{$R *.DFM}

{ TGlobalPropertyPage }

procedure TGlobalPropertyPage.DefinePropertyPages(DefinePropertyPage: TDefinePropertyPage);
begin
  { Define property pages here.  Property pages are defined by calling
    DefinePropertyPage with the class id of the page.  For example,
      DefinePropertyPage(Class_GlobalPropertyPage); }
end;

procedure TGlobalPropertyPage.EventSinkChanged(const EventSink: IUnknown);
begin
  FEvents := EventSink as IGlobalPropertyPageEvents;
end;

procedure TGlobalPropertyPage.Initialize;
begin
  inherited Initialize;
  OnActivate := ActivateEvent;
  OnClick := ClickEvent;
  OnCreate := CreateEvent;
  OnDblClick := DblClickEvent;
  OnDeactivate := DeactivateEvent;
  OnDestroy := DestroyEvent;
  OnKeyPress := KeyPressEvent;
  OnPaint := PaintEvent;
end;

function TGlobalPropertyPage.Get_Active: WordBool;
begin
  Result := Active;
end;

function TGlobalPropertyPage.Get_AutoScroll: WordBool;
begin
  Result := AutoScroll;
end;

function TGlobalPropertyPage.Get_AutoSize: WordBool;
begin
  Result := AutoSize;
end;

function TGlobalPropertyPage.Get_AxBorderStyle: TxActiveFormBorderStyle;
begin
  Result := Ord(AxBorderStyle);
end;

function TGlobalPropertyPage.Get_Caption: WideString;
begin
  Result := WideString(Caption);
end;

function TGlobalPropertyPage.Get_Color: Integer;
begin
  Result := Integer(Color);
end;

function TGlobalPropertyPage.Get_Cursor: Smallint;
begin
  Result := Smallint(Cursor);
end;

function TGlobalPropertyPage.Get_DoubleBuffered: WordBool;
begin
  Result := DoubleBuffered;
end;

function TGlobalPropertyPage.Get_DropTarget: WordBool;
begin
  Result := DropTarget;
end;

function TGlobalPropertyPage.Get_Enabled: WordBool;
begin
  Result := Enabled;
end;

function TGlobalPropertyPage.Get_HelpFile: WideString;
begin
  Result := WideString(HelpFile);
end;

function TGlobalPropertyPage.Get_KeyPreview: WordBool;
begin
  Result := KeyPreview;
end;

function TGlobalPropertyPage.Get_PixelsPerInch: Integer;
begin
  Result := PixelsPerInch;
end;

function TGlobalPropertyPage.Get_PrintScale: TxPrintScale;
begin
  Result := Ord(PrintScale);
end;

function TGlobalPropertyPage.Get_Scaled: WordBool;
begin
  Result := Scaled;
end;

function TGlobalPropertyPage.Get_Visible: WordBool;
begin
  Result := Visible;
end;

function TGlobalPropertyPage.Get_VisibleDockClientCount: Integer;
begin
  Result := VisibleDockClientCount;
end;

procedure TGlobalPropertyPage.ActivateEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnActivate;
end;

procedure TGlobalPropertyPage.ClickEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnClick;
end;

procedure TGlobalPropertyPage.CreateEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnCreate;
end;

procedure TGlobalPropertyPage.DblClickEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnDblClick;
end;

procedure TGlobalPropertyPage.DeactivateEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnDeactivate;
end;

procedure TGlobalPropertyPage.DestroyEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnDestroy;
end;

procedure TGlobalPropertyPage.KeyPressEvent(Sender: TObject; var Key: Char);
var
  TempKey: Smallint;
begin
  TempKey := Smallint(Key);
  if FEvents <> nil then FEvents.OnKeyPress(TempKey);
  Key := Char(TempKey);
end;

procedure TGlobalPropertyPage.PaintEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnPaint;
end;

procedure TGlobalPropertyPage.Set_AutoScroll(Value: WordBool);
begin
  AutoScroll := Value;
end;

procedure TGlobalPropertyPage.Set_AutoSize(Value: WordBool);
begin
  AutoSize := Value;
end;

procedure TGlobalPropertyPage.Set_AxBorderStyle(Value: TxActiveFormBorderStyle);
begin
  AxBorderStyle := TActiveFormBorderStyle(Value);
end;

procedure TGlobalPropertyPage.Set_Caption(const Value: WideString);
begin
  Caption := TCaption(Value);
end;

procedure TGlobalPropertyPage.Set_Color(Value: Integer);
begin
  Color := TColor(Value);
end;

procedure TGlobalPropertyPage.Set_Cursor(Value: Smallint);
begin
  Cursor := TCursor(Value);
end;

procedure TGlobalPropertyPage.Set_DoubleBuffered(Value: WordBool);
begin
  DoubleBuffered := Value;
end;

procedure TGlobalPropertyPage.Set_DropTarget(Value: WordBool);
begin
  DropTarget := Value;
end;

procedure TGlobalPropertyPage.Set_Enabled(Value: WordBool);
begin
  Enabled := Value;
end;

procedure TGlobalPropertyPage.Set_HelpFile(const Value: WideString);
begin
  HelpFile := String(Value);
end;

procedure TGlobalPropertyPage.Set_KeyPreview(Value: WordBool);
begin
  KeyPreview := Value;
end;

procedure TGlobalPropertyPage.Set_PixelsPerInch(Value: Integer);
begin
  PixelsPerInch := Value;
end;

procedure TGlobalPropertyPage.Set_PrintScale(Value: TxPrintScale);
begin
  PrintScale := TPrintScale(Value);
end;

procedure TGlobalPropertyPage.Set_Scaled(Value: WordBool);
begin
  Scaled := Value;
end;

procedure TGlobalPropertyPage.Set_Visible(Value: WordBool);
begin
  Visible := Value;
end;

destructor TGlobalPropertyPage.Destroy;
var
  ParkingHandle: HWND;
begin
  ParkingHandle := FindWindowEx(0, 0, 'DAXParkingWindow', nil);
  if ParkingHandle <> 0 then
    SendMessage(ParkingHandle, WM_CLOSE, 0, 0);
  inherited Destroy;
end;

{ Outlook PropertyPage }

function TGlobalPropertyPage.GetPageInfo(var HelpFile: WideString;
  var HelpContext: Integer): HResult;
begin
  HelpFile := '';
  HelpContext := 0;
  Result := S_OK;
end;

function TGlobalPropertyPage.Get_Dirty(out Dirty: WordBool): HResult;
begin
  Dirty := FDirty;
  Result := S_OK;
end;

function TGlobalPropertyPage.Apply: HResult;
begin
  // TODO - put your code here
  FDirty := False;
  Result := S_OK;
end;

procedure TGlobalPropertyPage.GetPropertyPageSite;
begin
  if FPropertyPageSite = nil then
    if Assigned(ActiveFormControl) then
      if Assigned(ActiveFormControl.ClientSite) then
        ActiveFormControl.ClientSite.QueryInterface(IID_PropertyPageSite, FPropertyPageSite);
end;

procedure TGlobalPropertyPage.UpdatePropertyPageSite;
begin
  if Assigned(FPropertyPageSite) and not FDirty then
  begin
    FDirty := True;
    FPropertyPageSite.OnStatusChange;
  end;
end;

procedure TGlobalPropertyPage.Edit1Change(Sender: TObject);
begin
  GetPropertyPageSite;
  // TODO - put your code here
  UpdatePropertyPageSite;
end;

initialization
  TActiveFormFactory.Create(
    ComServer,
    TActiveFormControl,
    TGlobalPropertyPage,
    Class_GlobalPropertyPage,
    1,
    '',
    OLEMISC_SIMPLEFRAME or OLEMISC_ACTSLIKELABEL,
    tmApartment);
end.
