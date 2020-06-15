{ /*
  MIT License

  Copyright (c) 2018-2020 Wuping Xin
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

unit VissimInProcComProxyObject;

{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  ActiveX, ComObj, StdVcl, VissimInProcComProxyLib_TLB;

type
  TVissimInProcComProxy = class(TAutoObject, IVissimInProcComProxy)
  strict private
    class var FScriptEngine: IUnknown;
    class destructor FinalizeScriptEngine;
    function GetVissimInterface: IUnknown; safecall;
    procedure SetVissimScriptEngine(const scriptEngine: IUnknown); safecall;
  public
    destructor Destroy; override;
    procedure Initialize; override;
  end;

implementation

uses
  ComServ, System.SysUtils, Windows, ActiveScriptingLib;

const
  SFailToGetVissimHostSite  = 'Fail to get Vissim host site with error code [%d]';
  SFailToGetVissimInterface = 'Fail to get Vissim interface with error code [%d]';
  SUnassignedScriptEngine   = 'Unassigned script engine.';

destructor TVissimInProcComProxy.Destroy;
begin
  inherited;
end;

class destructor TVissimInProcComProxy.FinalizeScriptEngine;
begin
  FScriptEngine := nil;
end;

function TVissimInProcComProxy.GetVissimInterface: IUnknown;
var
  comResult: HResult;
  ptrScriptSite: Pointer;
  scriptSite: IActiveScriptSite;
  scriptSiteIid: TGUID;
  typeInfo: IInterface;
  vbsEngine: IActiveScript;
  //debugMessage: string;
const
  vissimName: string = 'vissim';
begin
  Result := nil;

  //debugMessage := 'TVissimInProcComProxy.GetVissimInterface Entered.';
  //OutputDebugString(PChar(debugMessage));

  if not Assigned(FScriptEngine) then
    raise EOleSysError.Create(SUnassignedScriptEngine, - 1, 0);

  scriptSiteIid := IID_IActiveScriptSite;
  vbsEngine := FScriptEngine as IActiveScript;
  comResult := vbsEngine.GetScriptSite(scriptSiteIid, ptrScriptSite);

  if comResult <> S_OK then
    raise EOleSysError.Create(Format(SFailToGetVissimHostSite, [comResult]), comResult, 0);

  scriptSite := IActiveScriptSite(ptrScriptSite);
  comResult := scriptSite.GetItemInfo(PChar(vissimName), SCRIPTINFO_IUNKNOWN, Result, typeInfo);

  if comResult <> S_OK then
    raise EOleSysError.Create(Format(SFailToGetVissimInterface, [comResult]), comResult, 0);
end;

procedure TVissimInProcComProxy.Initialize;
begin
  inherited;
end;

procedure TVissimInProcComProxy.SetVissimScriptEngine(const scriptEngine: IUnknown);
begin
  FScriptEngine := scriptEngine;
end;

initialization

TAutoObjectFactory.Create(
  ComServer,
  TVissimInProcComProxy,
  Class_VissimInProcComProxy,
  ciSingleInstance,
  tmApartment);

end.
