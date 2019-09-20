library VissimComHook;

uses
  madCodeHook,
  uHookFunctions in 'uHookFunctions.pas';

{$R *.res}

begin
  CollectHooks;

  HookAPI(
    'ole32.dll',
    'CoCreateInstance',
    @CoCreateInstanceCallback,
    @CoCreateInstanceNext
  );

  {$IFDEF DEBUG}
  HookAPI(
    'ole32.dll',
    'CoRegisterClassObject',
    @CoRegisterClassObjectCallback,
    @CoRegisterClassObjectNext
  );
  
  HookAPI(
    'ole32.dll',
    'CoRevokeClassObject',
    @CoRevokeClassObjectCallback,
    @CoRevokeClassObjectNext
  );
  
  HookAPI(
    'ole32.dll', 
    'CoInitializeEx',
    @CoInitializeExCallback,
    @CoInitializeExNext
  );
  
  HookAPI(
    'ole32.dll', 
    'CoUninitialize',
    @CoUninitializeCallback,
    @CoUninitializeNext
  );
  {$ENDIF}

  FlushHooks;
end.
