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

program VissimLauncher;

{$APPTYPE CONSOLE}
{$R *.res}

uses
  System.SysUtils, Winapi.Windows, madCodeHook;

begin
  try
    var hookDllPath :=
    {$IFDEF DEBUG}
      'C:\DEVLIBS\Vissim.IPC.Proxy\Bin\Win64VissimComHook.dll';
    {$ELSE}
      ExtractFilePath(ParamStr(0)) + 'VissimComHook.dll';
    {$ENDIF}

    var vissimPath :=
    {$IFDEF DEBUG}
      'C:\Program Files\PTV Vision\PTV Vissim 2020\Exe\VISSIM200.exe';
    {$ELSE}
      ExtractFilePath(ParamStr(0)) + 'VISSIM200.exe';
    {$ENDIF}

    var cmdParams := EmptyStr;
    for var I := 1 to ParamCount do cmdParams := cmdParams + ' ' + ParamStr(I);

    var startInfo: TStartupInfo;
    startInfo.cb := SizeOf(TStartupInfo);
    FillChar(startInfo, SizeOf(TStartupInfo), #0);

    var procInfo: TProcessInformation;
    FillChar(procInfo, SizeOf(TProcessInformation), #0);

    var createOK := CreateProcessExW(
      PWideChar(vissimPath),
      PWideChar(cmdParams),
      nil,
      nil,
      False,
      CREATE_NEW_PROCESS_GROUP + NORMAL_PRIORITY_CLASS,
      nil,
      nil,
      startInfo,
      procInfo,
      PChar(hookDllPath));

    if createOK then
      WriteLn('Successfully launched Vissim.');
  except
    on E: Exception do
      WriteLn(E.ClassName, ': ', E.Message);
  end;

end.
