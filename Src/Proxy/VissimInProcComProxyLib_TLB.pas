{ /*
  MIT License

  Copyright (c) 2019 wxinixs@kld
  https://github.com/wxinix

  Permission is hereby granted, free of charge, to any person obtaining a copy
  of this software and associated documentation files (the "Software"), to deal
  in the Software without restriction, including without limitation the rights
  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
  copies of the Software, and to permit persons to whom the Software is
  furnished to do so, subject to the following conditions:

  The above copyright notice and this permission notice shall be included in all
  copies or substantial portions of the Software.

  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
  SOFTWARE.
  */ }

unit VissimInProcComProxyLib_TLB;

{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers.
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
{$ALIGN 4}

interface

uses
  System.Classes,
  System.Variants,
  System.Win.StdVCL,
  Vcl.Graphics,
  Vcl.OleServer,
  Winapi.ActiveX,
  Winapi.Windows;

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:
//   Type Libraries     : LIBID_xxxx
//   CoClasses          : CLASS_xxxx
//   DISPInterfaces     : DIID_xxxx
//   Non-DISP interfaces: IID_xxxx
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  VissimInProcComProxyLibMajorVersion = 1;
  VissimInProcComProxyLibMinorVersion = 0;

  LIBID_VissimInProcComProxyLib
    : TGUID = '{7F6948DF-58E4-42BB-8160-0D11ADE3FF10}';

  IID_IVissimInProcComProxy: TGUID  = '{A2E84260-9C33-4EBB-8B0C-DB00B164EA74}';
  CLASS_VissimInProcComProxy: TGUID = '{864FDD1D-F2AD-4A58-B089-94B64251A18E}';

type

  // *********************************************************************//
  // Forward declaration of types defined in TypeLibrary
  // *********************************************************************//
  IVissimInProcComProxy     = interface;
  IVissimInProcComProxyDisp = dispinterface;

  // *********************************************************************//
  // Declaration of CoClasses defined in Type Library
  // (NOTE: Here we map each CoClass to its Default Interface)
  // *********************************************************************//
  VissimInProcComProxy = IVissimInProcComProxy;

  // *********************************************************************//
  // Interface: IVissimInProcComProxy
  // Flags:     (4416) Dual OleAutomation Dispatchable
  // GUID:      {A2E84260-9C33-4EBB-8B0C-DB00B164EA74}
  // *********************************************************************//
  IVissimInProcComProxy = interface(IDispatch)
    ['{A2E84260-9C33-4EBB-8B0C-DB00B164EA74}']
    function GetVissimInterface: IUnknown; safecall;
    procedure SetVissimScriptEngine(const scriptEngine: IUnknown); safecall;
  end;

  // *********************************************************************//
  // DispIntf:  IVissimInProcComProxyDisp
  // Flags:     (4416) Dual OleAutomation Dispatchable
  // GUID:      {A2E84260-9C33-4EBB-8B0C-DB00B164EA74}
  // *********************************************************************//
  IVissimInProcComProxyDisp = dispinterface
    ['{A2E84260-9C33-4EBB-8B0C-DB00B164EA74}']
    function GetVissimInterface: IUnknown; dispid 201;
    procedure SetVissimScriptEngine(const scriptEngine: IUnknown); dispid 202;
  end;

  // *********************************************************************//
  // The Class CoVissimInProcComProxy provides a Create and CreateRemote method to
  // create instances of the default interface IVissimInProcComProxy exposed by
  // the CoClass VissimInProcComProxy. The functions are intended to be used by
  // clients wishing to automate the CoClass objects exposed by the
  // server of this typelibrary.
  // *********************************************************************//
  CoVissimInProcComProxy = class
    class function Create: IVissimInProcComProxy;
    class function CreateRemote(const MachineName: string)
      : IVissimInProcComProxy;
  end;

implementation

uses System.Win.ComObj;

class function CoVissimInProcComProxy.Create: IVissimInProcComProxy;
begin
  Result := CreateComObject(CLASS_VissimInProcComProxy)
    as IVissimInProcComProxy;
end;

class function CoVissimInProcComProxy.CreateRemote(const MachineName: string)
  : IVissimInProcComProxy;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_VissimInProcComProxy)
    as IVissimInProcComProxy;
end;

end.
