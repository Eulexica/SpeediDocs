unit MatterList;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Outlook2000, adxolFormsManager;

type
  TadxOlForm1 = class(TadxOlForm)
  private
    { Private declarations }
  protected
    { Protected declarations }
  public
    { Public declarations }
  published
    { Published declarations }
  end;

{NOTE: The adxOlForm1 variable is intended for the exclusive use
       by the TadxOlFormsCollectionItem Designer.
       NEVER use this variable for other purposes.}
var
  adxOlForm1 : TadxOlForm1;

implementation

{$R *.DFM}

initialization
  RegisterClass(TPersistentClass(TadxOlForm1));

finalization
end.
