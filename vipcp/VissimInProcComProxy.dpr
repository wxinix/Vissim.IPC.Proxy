library VissimInProcComProxy;

uses
  ComServ,
  uVissimInProcComProxyObject in 'uVissimInProcComProxyObject.pas',
  VissimInProcComProxyLib_TLB in 'VissimInProcComProxyLib_TLB.pas';

exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer,
  DllInstall;

{$R *.TLB}
{$R *.RES}

begin
  //
end.
