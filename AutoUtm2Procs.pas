unit AutoUtm2Procs;
interface
uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ShellApi, Buttons, Mask, ComCtrls, IdBaseComponent,
  IdComponent, IdTCPConnection, IdTCPClient, IdTelnet, Menus,
  tlntsend, XPMan, ExtCtrls, IniFiles,FileCtrl, ExtCtrlsX;

type
 TTran = record // ���������� ���!
    id: string; // id �������� ����� � UTM
    summ: string; // ����� �������
    date: string; // ���� �������
    num: string; // ����� ���������(��������) ���������
  end;

function Parce(str: string; separator: char): TStringList;
function strtotr(str: string): TTran;
function getlogin(str: string): string;
function getdate(str: string): string;
function getbalance(str: string): Extended;
procedure SimulateKeyDown(Key: byte);
procedure SimulateKeystroke(Key: byte; extra: DWORD);
procedure SendKeys(s: string);
procedure mlclick(x: integer; y: integer);
procedure mrclick(x: integer; y: integer);
procedure FindFile(Dir: string; var list: TStringList);
procedure PressKey(keyCode: Byte);
function DateTimeToUnix(ConvDate: TDateTime): Longint;
function UnixToDateTime(USec: Longint): TDateTime;
procedure CheckAndCreatePath(Path: string);

implementation

//---------������ �� ������������ ������-------------------------------------

function Parce(str: string; separator: char): TStringList;
var
  Tmp: TStringList; {��������� ������ �� ������ �����������}
begin
  Tmp := TStringList.Create;
  Tmp.Text := StringReplace(str, separator, #13#10, [rfReplaceAll]);
  Result := Tmp;
  tmp.Destroy;
end;
//----------------������ �� ������ ����������---------------------------------

function strtotr(str: string): TTran; //�.�. ������� �������� ������
var
  Tmp: TStringList;
begin
  Tmp := TStringList.Create;
  Tmp.Text := StringReplace(str, ';', #13#10, [rfReplaceAll]);
  Result.id := Tmp[2];
  Result.summ := Tmp[3];
  Result.date := StringReplace(Tmp[8], '/', '.', [rfReplaceAll]) + ' 00:00:00';
  Result.num := Tmp[10];
  tmp.Destroy;
end;
//------------------------------------------------------------------------------

function getlogin(str: string): string; //������ �� ������ �����
var //�� ����� �����
  Tmp: TStringList;
begin
  Tmp := TStringList.Create;
  Tmp.Text := StringReplace(str, ';', #13#10, [rfReplaceAll]);
  Result := Tmp[1];
  tmp.Destroy;
end;
//------------------------------------------------------------------------------

function getdate(str: string): string; //������ �� ������ ����
var //��� ��� ��� ����� \
  Tmp: TStringList;
begin
  Tmp := TStringList.Create;
  Tmp.Text := StringReplace(str, '\', #13#10, [rfReplaceAll]);
  if Tmp.count <> 3 then
    Result := 'out fo date'
  else
    Result := Tmp[Tmp.Count - 2];
  tmp.Destroy;
end;
//-----������ �� ������ ����� ������� �����--------------------------------

function getbalance(str: string): Extended; // ������������ �� �����
var
  Tmp: TStringList;
begin
  if str = '' then
    Result := 0
  else
  begin
    Tmp := TStringList.Create;
    Tmp.Text := StringReplace(str, ';', #13#10, [rfReplaceAll]);
    Result := strtofloat(StringReplace(Tmp[5], '.', ',', [rfReplaceAll]));
    tmp.Destroy;
  end;
end;
//---------------������� �������  ---------------------------------------------

procedure SimulateKeyDown(Key: byte);
begin
  keybd_event(Key, 0, 0, 0);
end;
//----------------� �� �������-------------------------------------------------

procedure SimulateKeyUp(Key: byte);
begin
  keybd_event(Key, 0, KEYEVENTF_KEYUP, 0);
end;
//------------------------------------------------------------------------------

procedure SimulateKeystroke(Key: byte; extra: DWORD);
begin
  keybd_event(Key, extra, 0, 0);
  keybd_event(Key, extra, KEYEVENTF_KEYUP, 0);
end;
//------------------------------------------------------------------------------

procedure SendKeys(s: string);
var
  i: integer;
  flag: bool;
  w: word;
begin
  {Get the state of the caps lock key}
  flag := not GetKeyState(VK_CAPITAL) and 1 = 0;
  {If the caps lock key is on then turn it off}
  if flag then
    SimulateKeystroke(VK_CAPITAL, 0);
  for i := 1 to Length(s) do
  begin
    w := VkKeyScan(s[i]);
    {If there is not an error in the key translation}
    if ((HiByte(w) <> $FF) and (LoByte(w) <> $FF)) then
    begin
      {If the key requires the shift key down - hold it down}
      if HiByte(w) and 1 = 1 then
        SimulateKeyDown(VK_SHIFT);
      {Send the VK_KEY}
      SimulateKeystroke(LoByte(w), 0);
      {If the key required the shift key down - release it}
      if HiByte(w) and 1 = 1 then
        SimulateKeyUp(VK_SHIFT);
    end;
  end;
  {if the caps lock key was on at start, turn it back on}
  if flag then
    SimulateKeystroke(VK_CAPITAL, 0);
end;
//-----------------------------------------------------------------------------

procedure mlclick(x: integer; y: integer);
var
  Pt: TPoint;
begin
  Pt.x := Round(x * (65535 / Screen.Width));
  Pt.y := Round(y * (65535 / Screen.Height));
  Mouse_Event(MOUSEEVENTF_ABSOLUTE or MOUSEEVENTF_MOVE, Pt.x, Pt.y, 0, 0);
  Mouse_Event(MOUSEEVENTF_ABSOLUTE or MOUSEEVENTF_LEFTDOWN, Pt.x, Pt.y, 0, 0);
  Mouse_Event(MOUSEEVENTF_ABSOLUTE or MOUSEEVENTF_LEFTUP, Pt.x, Pt.y, 0, 0);
end;
//-----------------------------------------------------------------------------

procedure mrclick(x: integer; y: integer);
var
  Pt: TPoint;
begin
  Pt.x := Round(x * (65535 / Screen.Width));
  Pt.y := Round(y * (65535 / Screen.Height));
  Mouse_Event(MOUSEEVENTF_ABSOLUTE or MOUSEEVENTF_MOVE, Pt.x, Pt.y, 0, 0);
  Mouse_Event(MOUSEEVENTF_ABSOLUTE or MOUSEEVENTF_RIGHTDOWN, Pt.x, Pt.y, 0, 0);
  Mouse_Event(MOUSEEVENTF_ABSOLUTE or MOUSEEVENTF_RIGHTUP, Pt.x, Pt.y, 0, 0);
end;
//-----------------------------------------------------------------------------

procedure FindFile(Dir: string; var list: TStringList);
var
  SR: TSearchRec;
  FindRes: Integer;
begin
  FindRes := FindFirst(Dir + '*.*', faAnyFile, SR);
  while FindRes = 0 do
  begin
    if ((SR.Attr and faDirectory) = faDirectory) and
      ((SR.Name = '.') or (SR.Name = '..')) then
    begin
      FindRes := FindNext(SR);
      Continue;
    end;
    if ((SR.Attr and faDirectory) = faDirectory) then // ���� ������ �������, ��
    begin
      FindFile(Dir + SR.Name + '\', list);
      // ������ � ��������� ������ � ����������� �������� �������� + �������, ��� �� �����
      FindRes := FindNext(SR);
      // ����� ������� ���������� �������� �� ���������� ����� � ���� ��������
      Continue; // ���������� ����      �!!
    end;
    List.Add(Dir + SR.Name);
    FindRes := FindNext(SR);
  end;
  FindClose(SR);
end;
//------------------------------------------------------------------------------

procedure PressKey(keyCode: Byte);
begin
  keybd_event(keyCode, 0, 0, 0);
  keybd_event(keyCode, 0, KEYEVENTF_KEYUP, 0);
end;
//****************************************************************************

function DateTimeToUnix(ConvDate: TDateTime): Longint;
const
  UnixStartDate: TDateTime = 25569.0;
begin
  Result := Round((ConvDate - UnixStartDate) * 86400);
end;
//****************************************************************************

function UnixToDateTime(USec: Longint): TDateTime;
const
  UnixStartDate: TDateTime = 25569.0;
begin
  Result := (Usec / 86400) + UnixStartDate;
end;
//******************************************************************************
procedure CheckAndCreatePath(Path: string);
var tmp : TStringList;
i: Integer;
str:string;
begin
 Tmp := TStringList.Create;
 Tmp.Text := StringReplace(Path, '\', #13#10, [rfReplaceAll]);
 str:=tmp[0];
 for i:=1 to tmp.Count-2 do
  begin
    str:=str+'\'+tmp[i];
    if not DirectoryExists(str) then
    CreateDir(str);
    end;
  end;


end.
