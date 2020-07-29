unit functions;

interface

uses
  xmldom, XMLIntf, msxmldom, XMLDoc, JvMemoryDataset, DB;

procedure Data2XML(AData: TJvMemoryData; AProjectPath: string);
procedure XML2Data;

implementation

procedure Data2XML(AData: TJvMemoryData; AProjectPath: string);
var
  XMLDoc : TXMLDocument;
  iNode : IXMLNode;
  ARec: integer;
  cNode : IXMLNode;

begin
  XMLDoc := TXMLDocument.Create(nil);
  XMLDoc.Active := True;
  iNode := XMLDoc.AddChild('GeniDocs');
  iNode.Attributes['app'] := ParamStr(0);

  ARec := 0;
  while (not AData.Eof) do
  begin
    cNode := iNode.AddChild('item');
    cNode.Attributes['Field Name'] := AData.Fields.Fields[ARec].Text;
    cNode.Attributes['Value'] := AData.Fields.Fields[ARec].AsString;
    inc(ARec);
  end;

  XMLDoc.SaveToFile(AProjectPath);

end; (* Data2XML *)

procedure XML2Data;
begin

end;

end.
