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

library VissimComHook;

uses
  madCodeHook,
  HookFunctions in 'HookFunctions.pas';

{$R *.res}

begin
  CollectHooks;

  HookAPI(
    'ole32.dll',
    'CoCreateInstance',
    @CoCreateInstanceCallback,
    @CoCreateInstanceNext);

  {$IFDEF DEBUG}
  HookAPI(
    'ole32.dll',
    'CoRegisterClassObject',
    @CoRegisterClassObjectCallback,
    @CoRegisterClassObjectNext);

  HookAPI(
    'ole32.dll',
    'CoRevokeClassObject',
    @CoRevokeClassObjectCallback,
    @CoRevokeClassObjectNext);

  HookAPI(
    'ole32.dll',
    'CoInitializeEx',
    @CoInitializeExCallback,
    @CoInitializeExNext);

  HookAPI(
    'ole32.dll',
    'CoUninitialize',
    @CoUninitializeCallback,
    @CoUninitializeNext);
  {$ENDIF}
  FlushHooks;

end.
