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

program VissimLauncher;

{$APPTYPE CONSOLE}
{$R *.res}

uses
  System.SysUtils,
  Winapi.Windows,
  madCodeHook;

var
  StartInfo: TStartupInfo;
  ProcInfo: TProcessInformation;
  CreateOK: Boolean;
  VissimPath: string;
  HookDllPath: string;
  CmdParams: string;
  I: Integer;

begin
  try
    {$IFDEF DEBUG}
    VissimPath := 'C:\Program Files\PTV Vision\PTV Vissim 11\Exe\vissim110.exe';
    HookDllPath := 'X:\WX-Codes\VISSIM_APPS\bin\Win64\Debug\VissimComHook.dll';
    {$ELSE}
    VissimPath := ExtractFilePath(ParamStr(0)) + 'VISSIM110.exe';
    HookDllPath := ExtractFilePath(ParamStr(0)) + 'VissimComHook.dll';
    {$ENDIF}
    CmdParams := EmptyStr;

    for I := 1 to ParamCount do
      CmdParams := CmdParams + ' ' + ParamStr(I);

    FillChar(StartInfo, SizeOf(TStartupInfo), #0);
    FillChar(ProcInfo, SizeOf(TProcessInformation), #0);

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
      WriteLn(E.ClassName, ': ', E.Message);
  end;

end.
