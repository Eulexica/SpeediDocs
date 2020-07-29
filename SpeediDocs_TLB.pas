unit SpeediDocs_TLB;

// ************************************************************************ //
// WARNING
// -------
// The types declared in this file were generated from data read from a
// Type Library. If this type library is explicitly or indirectly (via
// another type library referring to this type library) re-imported, or the
// 'Refresh' command of the Type Library Editor activated while editing the
// Type Library, the contents of this file will be regenerated and all
// manual modifications will be lost.
// ************************************************************************ //

// $Rev: 45604 $
// File generated on 8/11/2012 11:03:46 AM from Type Library described below.

// ************************************************************************  //
// Type Lib: D:\SpeediDocs\SpeediDocs (1)
// LIBID: {9F4133C2-47CA-4511-A74A-FA7A38955FEF}
// LCID: 0
// Helpfile:
// HelpString: SpeediDocs Library
// DepndLst:
//   (1) v2.0 stdole, (C:\Windows\SysWOW64\stdole2.tlb)
// SYS_KIND: SYS_WIN32
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers.
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
{$ALIGN 4}

interface

uses Winapi.Windows, System.Classes, System.Variants, System.Win.StdVCL, Vcl.Graphics, Vcl.OleServer, Winapi.ActiveX;


// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:
//   Type Libraries     : LIBID_xxxx
//   CoClasses          : CLASS_xxxx
//   DISPInterfaces     : DIID_xxxx
//   Non-DISP interfaces: IID_xxxx
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  SpeediDocsMajorVersion = 1;
  SpeediDocsMinorVersion = 0;

  LIBID_SpeediDocs: TGUID = '{9F4133C2-47CA-4511-A74A-FA7A38955FEF}';

  IID_IcoSpeediDocs: TGUID = '{888D05F7-E19C-40EE-AB71-78759EF6E977}';
  CLASS_coSpeediDocs: TGUID = '{404D7D9F-4CA1-4314-907C-8C8B6D5AB326}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary
// *********************************************************************//
  IcoSpeediDocs = interface;
  IcoSpeediDocsDisp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library
// (NOTE: Here we map each CoClass to its Default Interface)
// *********************************************************************//
  coSpeediDocs = IcoSpeediDocs;


// *********************************************************************//
// Interface: IcoSpeediDocs
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {888D05F7-E19C-40EE-AB71-78759EF6E977}
// *********************************************************************//
  IcoSpeediDocs = interface(IDispatch)
    ['{888D05F7-E19C-40EE-AB71-78759EF6E977}']
  end;

// *********************************************************************//
// DispIntf:  IcoSpeediDocsDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {888D05F7-E19C-40EE-AB71-78759EF6E977}
// *********************************************************************//
  IcoSpeediDocsDisp = dispinterface
    ['{888D05F7-E19C-40EE-AB71-78759EF6E977}']
  end;

// *********************************************************************//
// The Class CocoSpeediDocs provides a Create and CreateRemote method to
// create instances of the default interface IcoSpeediDocs exposed by
// the CoClass coSpeediDocs. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CocoSpeediDocs = class
    class function Create: IcoSpeediDocs;
    class function CreateRemote(const MachineName: string): IcoSpeediDocs;
  end;

implementation

uses System.Win.ComObj;

class function CocoSpeediDocs.Create: IcoSpeediDocs;
begin
  Result := CreateComObject(CLASS_coSpeediDocs) as IcoSpeediDocs;
end;

class function CocoSpeediDocs.CreateRemote(const MachineName: string): IcoSpeediDocs;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_coSpeediDocs) as IcoSpeediDocs;
end;

end.

