program XPLinkSync;

{$APPTYPE CONSOLE}

{$R *.res}


uses
  System.SysUtils,
  ShlObj,
  ActiveX,
  ComObj,
  System.Classes,
  Winapi.Windows;

type
  TXPFileCommandSet = (cNone, cTarger, cMask, cAddLinks, cSubLevel, cFixLinks, cPrint, cPause);

  TXPFileCommand = class(TObject)
    FTarget: string;
    FMask: String;
    FSubLevel: integer;
    FCommand: TXPFileCommandSet;
    FParams: string;
  protected
    function AddSlash(value: string): string;
    procedure ScanDir(aPath: string; aMask: string; aScanXPDirectory: boolean; var outList: TstringList);
    function ExstractCommand(value: string): TXPFileCommandSet;
    function ExstractParams(value: string): string;
    procedure check;
    procedure run_AddLinks;
    procedure run_FixLinks;
    procedure Error(sMessage: string);
    function isPathXP(aPath: string): boolean;
    function isDirNameValid(aName: string): boolean;
  public
    procedure init;
    procedure RunValue;
    procedure setValue(sValue: string);
  end;

var
  command: TXPFileCommand;

  { TCommand }

procedure CreateLink(const PathObj, PathLink: string);
var
  IObject: IUnknown;
  SLink: IShellLink;
  PFile: IPersistFile;
begin
  IObject := CreateComObject(CLSID_ShellLink);
  SLink := IObject as IShellLink;
  PFile := IObject as IPersistFile;
  with SLink do
  begin
    SetArguments(PChar(''));
    SetDescription(PChar(ExtractFileName(PathObj)));
    SetPath(PChar(PathObj));
  end;
  PFile.Save(PWidechar(WideString(PathLink)), FALSE);
end;

function CheckLink(PathLink: string): boolean;
var
  IObject: IUnknown;
  SLink: IShellLink;
  PFile: IPersistFile;
  FilePath: array [0 .. 255] of char;
  FindData: TWin32FindData;
  s: string;
begin
  IObject := CreateComObject(CLSID_ShellLink);
  SLink := IObject as IShellLink;
  PFile := IObject as IPersistFile;

  PFile.Load(PWidechar(PathLink), 0);
  SLink.GetPath(FilePath, 255, FindData, 0);
  s := FilePath;
  result := command.isDirNameValid(s) and DirectoryExists(s);
end;

function TXPFileCommand.AddSlash(value: string): string;
begin
  result := value;
  if (result <> '') and (result[result.Length] <> '\') then
    result := result + '\';
end;

procedure TXPFileCommand.check;
begin
  if not DirectoryExists(AddSlash(FTarget)) then
  begin
    Error(Format('Invalid #target "%s"', [FTarget]));
  end;
end;

procedure TXPFileCommand.Error(sMessage: string);
begin
  ExitCode := 1;
  Writeln(sMessage);
  readln;
  halt;
end;

function TXPFileCommand.ExstractCommand(value: string): TXPFileCommandSet;
var
  n: integer;
  s: string;
begin
  s := '';
  if (value <> '') and (value.Chars[0] = '#') then
  begin
    s := Copy(value, 2, Length(value));
    n := Pos(' ', s);
    if n > 0 then
      s := Copy(s, 1, n - 1);
  end;
  s := s.LowerCase(s);
  if s = 'target' then
    result := cTarger
  else if s = 'addlinks' then
    result := cAddLinks
  else if s = 'fixlinks' then
    result := cFixLinks
  else if s = 'linksublevel' then
    result := cSubLevel
  else if s = 'mask' then
    result := cMask
  else if s = 'print' then
    result := cPrint
  else if s = 'pause' then
    result := cPause
  else
    result := cNone;
end;

function TXPFileCommand.ExstractParams(value: string): string;
var
  n: integer;
begin
  result := '';
  if (value <> '') and (value.Chars[0] = '#') then
  begin
    n := Pos(' ', value);
    result := Copy(value, n + 1, Length(value));
    while (result <> '') and (result[result.Length] = ' ') do
      result := Copy(result, 1, Length(result) - 1);
  end;
end;

procedure TXPFileCommand.init;
var
  I: integer;
  sl: TstringList;
  sFilename: string;
begin
  FMask := '*.*';
  FTarget := '';
  FSubLevel := 1;

  sFilename := ChangeFileExt(ParamStr(0), '.txt');

  if ParamCount > 0 then
  begin
    sFilename := ParamStr(1);
  end;

  if not FileExists(sFilename) then
  begin
    Error(Format('txt file not found "%s"!', [sFilename]));
  end;

  sl := TstringList.Create;
  try
    sl.LoadFromFile(sFilename);
    for I := 0 to sl.Count - 1 do
      setValue(sl[I]);
  finally
    sl.Free;
  end;
end;

function TXPFileCommand.isDirNameValid(aName: string): boolean;
begin
  result := not SameText('.removed', ExtractFileExt(aName));
end;

function TXPFileCommand.isPathXP(aPath: string): boolean;
begin
  result := DirectoryExists(aPath + '\Earth nav data') or
    FileExists(aPath + '\library.txt');
end;

procedure TXPFileCommand.run_FixLinks;
var
  I: integer;
  sl: TstringList;
begin
  sl := TstringList.Create;
  try
    ScanDir(FTarget, '*.lnk', FALSE, sl);
    for I := 0 to sl.Count - 1 do
    begin
      if not CheckLink(sl[I]) then
      begin
        Writeln(Format('Deleted %s', [sl[I]]));
        System.SysUtils.DeleteFile(sl[I]);
      end;
    end;
  finally
    sl.Free;
  end;
end;

procedure TXPFileCommand.run_AddLinks;
var
  I, i2: integer;
  sl, sl2: TstringList;
  sLinkName: string;
begin
  check;
  sl := TstringList.Create;
  sl2 := TstringList.Create;
  try
    {
      case FSubLevel of
      0:
      ;
      1:
      begin
      ScanDir(FParams, FMask, true, sl);
      for I := 0 to sl.Count - 1 do
      begin
      sLinkName := AddSlash(FTarget) + ExtractFileName(sl[I]) + '.lnk';
      if Not FileExists(sLinkName) then
      begin
      CreateLink(sl[I], sLinkName);
      Writeln(Format('Add %s -> %s', [sl[I], sLinkName]));
      end;
      end;
      end;
      2:
      begin
      ScanDir(FParams, '*.*', true, sl);
      for I := 0 to sl.Count - 1 do
      begin
      ScanDir(sl[I], FMask, true, sl2);
      for i2 := 0 to sl2.Count - 1 do
      begin
      sLinkName := AddSlash(FTarget) + ExtractFileName(sl2[i2]) + '.lnk';
      if Not FileExists(sLinkName) then
      begin
      CreateLink(sl2[i2], sLinkName);
      Writeln(Format('Add %s -> %s', [sl2[i2], sLinkName]));
      end;
      end;
      end;
      end;
      end;
    }

    sl.Clear;
    Write('Scanning ' + FParams);
    ScanDir(FParams, FMask, true, sl);
    Writeln('');

    for I := 0 to sl.Count - 1 do
    begin
      sLinkName := AddSlash(FTarget) + ExtractFileName(sl[I]) + '.lnk';
      if Not FileExists(sLinkName) then
      begin
        CreateLink(sl[I], sLinkName);
        Writeln(Format('Add %s -> %s', [sl[I], sLinkName]));
      end;
    end;
    {
      for I := 0 to sl.Count - 1 do
      Writeln(sl[I])
    }

  finally
    sl.Free;
    sl2.Free;
  end;
end;

procedure TXPFileCommand.ScanDir(aPath, aMask: string; aScanXPDirectory: boolean; var outList: TstringList);
var
  sr: TsearchRec;
  FileAttrs: integer;
begin
  FileAttrs := faAnyFile;
  if System.SysUtils.findfirst(aPath + '\' + aMask, FileAttrs, sr) = 0 then
  begin
    repeat
      if (sr.Name <> '.') and (sr.Name <> '..') and isDirNameValid(sr.Name) then
        if (sr.Attr and faDirectory) <> 0 then
        begin
          if aScanXPDirectory then
          begin
            if isPathXP(aPath + '\' + sr.Name) then
            begin
              Write('.');
              outList.Add(aPath + '\' + sr.Name)
            end
            else
              ScanDir(aPath + '\' + sr.Name, aMask, aScanXPDirectory, outList);
          end;
        end
        else
        begin
          if not aScanXPDirectory then
            outList.Add(aPath + '\' + sr.Name);
        end;

    until System.SysUtils.FindNext(sr) <> 0;
    System.SysUtils.FindClose(sr);
  end;
end;

procedure TXPFileCommand.RunValue;
begin
  case FCommand of
    cNone:
      ;
    cTarger:
      FTarget := FParams;
    cAddLinks:
      run_AddLinks;
    cFixLinks:
      run_FixLinks;
    cSubLevel:
      begin
        FSubLevel := StrToIntDef(FParams, 0);
      end;
    cMask:
      FMask := FParams;
    cPrint:
      Writeln(FParams);
    cPause:
      readln;
  end;
end;

procedure TXPFileCommand.setValue(sValue: string);
begin
  FCommand := ExstractCommand(sValue);
  FParams := ExstractParams(sValue);
  RunValue;
end;

begin
  CoInitialize(nil);
  try
    try
      Writeln('XP LinkSync version 0.2 (c) Morten Isaksen');
      Writeln('');
      command := TXPFileCommand.Create;
      command.init;
    finally
      command.Free;
    end;
  except
    on E: Exception do
      Writeln(E.ClassName, ': ', E.Message);
  end;

end.
