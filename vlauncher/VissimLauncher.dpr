program VissimLauncher;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  System.SysUtils,
  Winapi.Windows,
  madCodeHook;

var
  StartInfo   : TStartupInfo;
  ProcInfo    : TProcessInformation;
  CreateOK    : Boolean;
  VissimPath  : string;
  HookDllPath : string;
  CmdParams   : string;
  I           : Integer;
begin
  try
    {$IFDEF DEBUG}
    VissimPath  := 'C:\Program Files\PTV Vision\PTV Vissim 10\Exe\vissim100.exe';
    HookDllPath := 'X:\VissimApp\VissimAppRs\bin\Win64\Debug\VissimComHook.dll';
    {$ELSE}
    VissimPath  := ExtractFilePath(ParamStr(0)) + 'VISSIM100.exe';
    HookDllPath := ExtractFilePath(ParamStr(0)) + 'VissimComHook.dll';
    {$ENDIF}

    CmdParams := EmptyStr;

    for I := 1 to ParamCount do
      CmdParams := CmdParams + ' ' + ParamStr(I);

    FillChar(StartInfo,SizeOf(TStartupInfo),#0);
    FillChar(ProcInfo,SizeOf(TProcessInformation),#0);

    StartInfo.cb := SizeOf(TStartupInfo);

    CreateOK := CreateProcessExW(
                  PWideChar(VissimPath),
                  PWideChar(CmdParams),
                  nil,
                  nil,
                  False,
                  CREATE_NEW_PROCESS_GROUP + NORMAL_PRIORITY_CLASS,
                  nil,
                  nil,
                  StartInfo,
                  ProcInfo,
                  PChar(HookDllPath));

    if CreateOK then
      WriteLn('Successfully launched Vissim');

  except
    on E: Exception do
      Writeln(E.ClassName, ': ', E.Message);
  end;
end.
