unit uHookFunctions;

interface

uses
  WinApi.ActiveX;

var
  CoCreateInstanceNext: function(
    const clsid: TCLSID;
    unkOuter: IUnknown;
    dwClsContext: Longint;
    const iid: TIID; out pv
  ): HRESULT; stdcall;

  CoRegisterClassObjectNext: function(
    const clsid: TCLSID;
    unk: IUnknown;
    dwClsContext: Longint;
    flags: Longint; 
    out dwRegister: Longint
    ): HRESULT; stdcall;

  CoRevokeClassObjectNext: function(
    dwRegister: Longint
  ): HRESULT; stdcall;

  CoInitializeExNext: function(
    pvReserved: Pointer;
    coInit: Longint
  ): HRESULT; stdcall;

  CoUninitializeNext: procedure; stdcall;

function CoCreateInstanceCallBack(
  const clsid: TCLSID;
  unkOuter: IUnknown;
  dwClsContext: Longint;
  const iid: TIID; out pv
): HRESULT; stdcall;

function CoRegisterClassObjectCallBack(
  const clsid: TCLSID; unk: IUnknown;
  dwClsContext: Longint;
  flags: Longint;
  out dwRegister: Longint
): HRESULT; stdcall;

function CoRevokeClassObjectCallBack(
  dwRegister: Longint
): HRESULT; stdcall;

function CoInitializeExCallBack(
  pvReserved: Pointer;
  coInit: Longint
): HRESULT; stdcall;

procedure CoUninitializeCallback; stdcall;

implementation

uses
  SysUtils, Windows, VissimInProcComProxyLib_TLB;

function CoCreateInstanceCallBack(const clsid: TCLSID; unkOuter: IUnknown;
    dwClsContext: Longint; const iid: TIID; out pv): HResult;
const
  IID_VBScript: TGUID      = '{B54F3741-5B07-11CF-A4B0-00AA004A55E8}';
  IID_IActiveScript: TGUID = '{BB1A2AE1-A4F9-11CF-8F20-00805F2CD064}';
var
  vissimInProcComProxy: IVissimInProcComProxy;
  scriptEngine: IInterface;
begin
  Result := CoCreateInstanceNext(clsid, unkOuter, dwClsContext, iid, pv);

  if (clsid = IID_VBScript) and (iid = IID_IActiveScript) and (Result = S_OK) then
  begin
    vissimInProcComProxy := CoVissimInProcComProxy.Create;
    scriptEngine := IUnknown(pv);
    vissimInProcComProxy.SetVissimScriptEngine(scriptEngine);
  end;
end;

function CoRegisterClassObjectCallBack(const clsid: TCLSID; unk: IUnknown;
    dwClsContext: Longint; flags: Longint; out dwRegister: Longint): HResult;
const
  fmtStr = 'CoRegisterClassObject called by Vissim[%d], with REGCLS flag = %d';
begin
  {$IFDEF DEBUG}
  OutputDebugString(PWideChar(Format(fmtStr, [GetCurrentThreadId, flags])));
  {$ENDIF}

  Result := CoRegisterClassObjectNext(clsid, unk, dwClsContext, flags, dwRegister);
end;

function CoRevokeClassObjectCallBack(dwRegister: Longint): HRESULT;
const
  fmtStr = 'CoRevokeClassObjectCallBack [%d], ThreadId[%d]';
begin
  {$IFDEF DEBUG}
  OutputDebugString(PWideChar(Format(fmtStr, [dwRegister, GetCurrentThreadId])));
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
