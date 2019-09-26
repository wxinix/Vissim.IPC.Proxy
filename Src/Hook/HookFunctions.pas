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

unit HookFunctions;

interface

uses
  WinApi.ActiveX;

var
  CoCreateInstanceNext: function(const clsid: TCLSID; unkOuter: IUnknown;
    dwClsContext: Longint; const iid: TIID; out pv): HRESULT; stdcall;

  CoRegisterClassObjectNext: function(const clsid: TCLSID; unk: IUnknown;
    dwClsContext: Longint; flags: Longint; out dwRegister: Longint)
    : HRESULT; stdcall;

  CoRevokeClassObjectNext: function(dwRegister: Longint): HRESULT; stdcall;

  CoInitializeExNext: function(pvReserved: Pointer; coInit: Longint)
    : HRESULT; stdcall;

  CoUninitializeNext: procedure; stdcall;

function CoCreateInstanceCallBack(const clsid: TCLSID; unkOuter: IUnknown;
  dwClsContext: Longint; const iid: TIID; out pv): HRESULT; stdcall;

function CoRegisterClassObjectCallBack(const clsid: TCLSID; unk: IUnknown;
  dwClsContext: Longint; flags: Longint; out dwRegister: Longint)
  : HRESULT; stdcall;

function CoRevokeClassObjectCallBack(dwRegister: Longint): HRESULT; stdcall;

function CoInitializeExCallBack(pvReserved: Pointer; coInit: Longint)
  : HRESULT; stdcall;

procedure CoUninitializeCallback; stdcall;

implementation

uses
  SysUtils,
  Windows,
  VissimInProcComProxyLib_TLB;

function CoCreateInstanceCallBack(const clsid: TCLSID; unkOuter: IUnknown;
  dwClsContext: Longint; const iid: TIID; out pv): HRESULT;
const
  IID_VBScript: TGUID      = '{B54F3741-5B07-11CF-A4B0-00AA004A55E8}';
  IID_IActiveScript: TGUID = '{BB1A2AE1-A4F9-11CF-8F20-00805F2CD064}';
var
  vissimInProcComProxy: IVissimInProcComProxy;
  scriptEngine: IInterface;
begin
  Result := CoCreateInstanceNext(clsid, unkOuter, dwClsContext, iid, pv);

  if (clsid = IID_VBScript) and (iid = IID_IActiveScript) and (Result = S_OK)
  then
  begin
    vissimInProcComProxy := CoVissimInProcComProxy.Create;
    scriptEngine := IUnknown(pv);
    vissimInProcComProxy.SetVissimScriptEngine(scriptEngine);
  end;
end;

function CoRegisterClassObjectCallBack(const clsid: TCLSID; unk: IUnknown;
  dwClsContext: Longint; flags: Longint; out dwRegister: Longint): HRESULT;
const
  fmtStr = 'CoRegisterClassObject called by Vissim[%d], with REGCLS flag = %d';
begin
  {$IFDEF DEBUG}
  OutputDebugString(PWideChar(Format(fmtStr, [GetCurrentThreadId, flags])));
  {$ENDIF}
  Result := CoRegisterClassObjectNext(clsid, unk, dwClsContext, flags,
    dwRegister);
end;

function CoRevokeClassObjectCallBack(dwRegister: Longint): HRESULT;
const
  fmtStr = 'CoRevokeClassObjectCallBack [%d], ThreadId[%d]';
begin
  {$IFDEF DEBUG}
  OutputDebugString(PWideChar(Format(fmtStr, [dwRegister,
          GetCurrentThreadId])));
  {$ENDIF}
  Result := CoRevokeClassObjectNext(dwRegister);
end;

function CoInitializeExCallBack(pvReserved: Pointer; coInit: Longint): HRESULT;
const
  fmtStr = 'CoInitializeExCallBack [%d], ThreadId[%d]';
begin
  {$IFDEF DEBUG}
  OutputDebugString(PWideChar(Format(fmtStr, [coInit, GetCurrentThreadId])));
  {$ENDIF}
  Result := CoInitializeExNext(pvReserved, coInit);
end;

procedure CoUninitializeCallback; stdcall;
const
  fmtStr = 'CoUnitializeCallBack invoked from ThreadId[%d].';
begin
  {$IFDEF DEBUG}
  OutputDebugString(PWideChar(Format(fmtStr, [GetCurrentThreadId])));
  {$ENDIF}
  CoUninitializeNext;
end;

end.
