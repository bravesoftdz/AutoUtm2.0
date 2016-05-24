unit AutoUtm2Main;
interface
uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ShellApi, Buttons, Mask, ComCtrls, IdBaseComponent,
  IdComponent, IdTCPConnection, IdTCPClient, IdTelnet, Menus,
  tlntsend, XPMan, ExtCtrls, IniFiles, FileCtrl, ExtCtrlsX;
type
  TfrmUtmAuto = class(TForm)
    Label5: TLabel;
    edtStartString: TEdit;
    edtSleepTime: TEdit;
    Label7: TLabel;
    StatusBar1: TStatusBar;
    grpSQL: TGroupBox;
    grpTelnet: TGroupBox;
    StaticText3: TStaticText;
    edtTLNHost: TMaskEdit;
    edtTLNUser: TEdit;
    edtTLNPwd: TEdit;
    StaticText4: TStaticText;
    edtSQLHost: TMaskEdit;
    edtSQLUser: TEdit;
    edtSQLPwd: TEdit;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    grpUTM: TGroupBox;
    StaticText2: TStaticText;
    edtUTMHost: TMaskEdit;
    Label2: TLabel;
    Label3: TLabel;
    edtUTMPass: TEdit;
    edtUTMUser: TEdit;
    grpFiles: TGroupBox;
    Label1: TLabel;
    edtPaymentPath: TEdit;
    edtReportPath: TEdit;
    Label6: TLabel;
    open: TOpenDialog;
    redtReport: TRichEdit;
    edtUtmPath: TEdit;
    lblUTMPath: TLabel;
    rgPrSel: TRadioGroup;
    mmenu: TMainMenu;
    Insert: TMenuItem;
    Check: TMenuItem;
    Balance: TMenuItem;
    Credit: TMenuItem;
    ClosePr: TMenuItem;
    grpCredit: TGroupBox;
    StaticText1: TStaticText;
    edtCreditSumm: TMaskEdit;
    xpmnfst1: TXPManifest;
    cbbUTMhost: TComboBox;
    btnAddHost: TButton;
    grpOther: TGroupBox;
    chkIsDelete: TCheckBox;
    Anoter: TMenuItem;
    InsertOld: TMenuItem;
    SQLInsert: TMenuItem;
    StartUTM: TMenuItem;
    StartOldUTM: TMenuItem;
    TrayIcon1: TTrayIcon;
    SomeOpt: TMenuItem;
    SetUtmPath: TMenuItem;
    SetPaymentPath: TMenuItem;
    SetReportPath: TMenuItem;
    btnCloneHost: TButton;
    lbl1: TLabel;
    edtOLDUTMPath: TEdit;
    lbl2: TLabel;
    pmTray: TPopupMenu;
    MoreFunc: TMenuItem;
    HahdInsert: TMenuItem;
    mniSQLQuery: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure ReadINI(Section: string);
    procedure WriteINI(Section: string);
    procedure InsertClick(Sender: TObject);
    procedure CheckClick(Sender: TObject);
    procedure BalanceClick(Sender: TObject);
    procedure CreditClick(Sender: TObject);
    procedure ClosePrClick(Sender: TObject);
    procedure cbbUTMhostChange(Sender: TObject);
    procedure btnAddHostClick(Sender: TObject);
    procedure InsertOldClick(Sender: TObject);
    procedure StartUTMClick(Sender: TObject);
    procedure edtSleepTimeChange(Sender: TObject);
    procedure SetUtmPathClick(Sender: TObject);
    procedure SetPaymentPathClick(Sender: TObject);
    procedure SetReportPathClick(Sender: TObject);
    procedure btnCloneHostClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure HahdInsertClick(Sender: TObject);
    procedure mniSQLQueryClick(Sender: TObject);
  private
    { Private declarations }
  public
    {!!!!!!!!!!!}
  end;

var
  frmUtmAuto: TfrmUtmAuto;
  Logpath, eLogpath, curentdate: string;
  log, elog: TextFile; // ���� ���
  Ressive: TStringList; // ���������� �� �������
  ini: Tinifile; // ������
  sleeptime: Integer;
implementation
uses MySQL5, AutoUtm2Procs, MainMiniUtm, Unit1;
{$R *.dfm}
//*************************service*******************************************

//******************************************************************************
  {                     ===
     \\ //   \\   //  ||  /||  �� ��� ����� ������ �������� ����� ���������!!!
      \\/     \\ //   || //||  ����� ������ ������������ �� ��������!!!
     //\\      \//    ||// ||
    //  \\     //     ||/  || }
//**********************MAIN****************************************************

procedure TfrmUtmAuto.InsertClick(Sender: TObject);
var
  fReestr, fReport: TextFile; //
  s, tmpstr, vipiska, ds, currenthostname, account: string;
  trn: TTran; // ����� ���� ���� �����
  Flist: Tstringlist; // ������ ������ � ����������
  i, j, totpay, leftpay, k, err, ret: integer; // (i,j)�������
  summ, balance: Extended; // ����� ���� ���������
  My1: TMySQL5;
  SynTel: TTelnetSend; // ������ ������ synapce
begin
  StatusBar1.Panels[1].Text := '';
  Application.ProcessMessages;
  SynTel := TTelnetSend.Create;
  SynTel.TargetHost := edtTLNHost.Text;
  SynTel.TermType := 'dumb'; // ��� ���������
  StatusBar1.Panels[1].Text := '�����������...';
  Application.ProcessMessages;
  case rgPrSel.ItemIndex of
    0: // ���� ������
      begin
        SynTel.TargetPort := '23';
        if not SynTel.Login then
        begin
          writeln(elog, DateTimeToStr(Now) + '�� ������������ � ' +
            edtTLNHost.Text);
          StatusBar1.Panels[1].Text := '�� ������������ � ' +
            edtTLNHost.Text;
          StatusBar1.Panels[0].Text := '��������';
          Exit;
        end
        else
          writeln(log, DateTimeToStr(Now) + '������������ � ' +
            edtTLNhost.Text);
        StatusBar1.Panels[0].Text := edtTLNHost.Text + ' ���������';
        Application.ProcessMessages;
        sleep(sleeptime);
        if not SynTel.WaitFor('login:') then
        begin
          writeln(elog, DateTimeToStr(Now) + '�� ������ ����� ' +
            edtTLNHost.Text);
          redtReport.Lines.Add(DateTimeToStr(Now));
          StatusBar1.Panels[1].Text := '������ ��������� (login) ';
          StatusBar1.Panels[0].Text := '��������';
          Exit;
        end;
        SynTel.Send(edtTLNUser.Text + #10#13);
        if not SynTel.WaitFor('Password:') then
        begin
          writeln(elog, DateTimeToStr(Now) + '�� ������ ������ ' +
            edtTLNHost.Text);
          StatusBar1.Panels[1].Text := '������ ��������� (Password) ';
          StatusBar1.Panels[0].Text := '��������';
          Exit;
        end;
        SynTel.Send(edtTLNPwd.Text + #10#13);
        if not SynTel.WaitFor(':~$') then
        begin
          writeln(elog, DateTimeToStr(Now) + '��� (Prompt) ' +
            edtTLNHost.Text);
          StatusBar1.Panels[1].Text := '������ ��������� (Prompt) ';
          StatusBar1.Panels[0].Text := '��������';
          Exit;
        end;
      end;
    1: // ���� ssh
      begin
        SynTel.TargetPort := '22';
        SynTel.UserName := edtTLNUser.Text;
        SynTel.Password := edtTLNPwd.Text;
        // SynTel.Login;
        if not SynTel.SSHLogin then
        begin
          writeln(elog, DateTimeToStr(Now) + '�� ������������ � ' +
            edtTLNhost.Text);
          Exit;
        end
        else
          writeln(log, DateTimeToStr(Now) + '������������ � ' +
            edtTLNhost.Text);
        StatusBar1.Panels[0].Text := edtTLNHost.Text + ' ���������';
        Application.ProcessMessages;
        sleep(sleeptime);
        if not SynTel.WaitFor(':~$') then
        begin
          writeln(elog, DateTimeToStr(Now) + '��� (Prompt) ' +
            edtTLNHost.Text);
          StatusBar1.Panels[1].Text := '������ ��������� (Prompt) ';
          StatusBar1.Panels[0].Text := '��������';
          Exit;
        end;
      end;
  else // ���� ������
    writeln(elog, DateTimeToStr(Now) + ' ���������� ������ #0001');
  end;
  // ������� ������ ------------------------------------------------
  StatusBar1.Panels[1].Text := '������� ������...';
  Application.ProcessMessages;
  sleep(sleeptime);
  My1 := TMySQL5.Create(nil);
  try
    My1.Connect(edtSQLHost.Text, edtSQLUser.Text, edtSQLPwd.Text);
    My1.Execute('select balance from UTM5.accounts where is_deleted = 0');
  except
    writeln(elog, DateTimeToStr(Now) + '������ ����������� � MySql ��:' +
      edtSQLHost.Text);
    redtReport.Lines.Add('������ ����������� � MySql ��:' +
      edtSQLHost.Text + ' ��������� �����������!');
    Exit;
  end;
  balance := 0;
  if My1.RecordCount > 0 then
    for k := 1 to My1.RecordCount do
    begin
      My1.Move(k);
      balance := balance + StrToFloat(StringReplace(My1.AsString(1), '.', ',',
        [rfReplaceAll]));
    end
  else
  begin
    writeln(elog, DateTimeToStr(Now) + '�� ������� ��������� ������!!!');
  end;
  My1.Close; // Close connection to MySQL server.
  //****************��������� ���� �������� *************************
  StatusBar1.Panels[1].Text := '������� ����� �������..';
  Application.ProcessMessages;
  sleep(sleeptime);
  AssignFile(fReport, edtReportPath.Text + '�������\������� �� ' +
    StringReplace(DateTimeToStr(now), ':', '-', [rfReplaceAll]) +
    '.txt'); //����� - ������������� �� 2 ������: 1100117
  CheckAndCreatePath(edtReportPath.Text + '�������\������� �� ' +
    StringReplace(DateTimeToStr(now), ':', '-', [rfReplaceAll]) +
    '.txt');
  Rewrite(fReport);
  //****************************************************************************
  StatusBar1.Panels[1].Text := '���� ����� ��������';
  Application.ProcessMessages;
  sleep(sleeptime);
  Flist := TStringList.Create;
  FindFile(edtPaymentPath.Text, Flist);
  // ����� ������ � ���������� ���������, �������� � Edit1
  // �������� ��������
  summ := 0; // ����� ���� ��������� ��������
  totpay := 0; // ����� �������� (����)
  leftpay := 0; // �����
  ret := 0; //���������
  err := 0; // ������ �����
  vipiska := ''; // �� ���� ��� ����� �� ��� ���������
  for i := 0 to (Flist.Count - 1) do
  begin
    if vipiska = '' then
      vipiska := getdate(Flist[i])
    else if vipiska <> getdate(Flist[i]) then
    begin
      redtReport.Lines.Add('');
      redtReport.Lines.Add('����� �� ' + DateTimeToStr(Now) + ' :');
      redtReport.Lines.Add('������ ������� : ' + inttostr(Flist.Count));
      redtReport.Lines.Add('�������� ������� : ' + inttostr(totpay));
      redtReport.Lines.Add('����� : ' + floattostr(summ));
      redtReport.Lines.Add('�������� ����� : ' + floattostr(leftpay));
      redtReport.Lines.Add('������ : ' + inttostr(err));
      redtReport.Lines.Add('�������� : ' + inttostr(ret));
      redtReport.Lines.Add('������ : ' + floattostr(balance));
      redtReport.Lines.Add('������� �� ' + vipiska);
      summ := 0; // ����� ���� ��������� ��������
      totpay := 0; // ����� �������� (����)
      leftpay := 0; // �����
      ret := 0; //���������
      err := 0; // ������ �����
      vipiska := getdate(Flist[i]);
    end;
    tmpstr := copy(Flist[i], length(Flist[i]) - 3, 4);
    if tmpstr <> '.txt' then
      continue;
    AssignFile(fReestr, Flist[i]);
    Reset(fReestr);
    StatusBar1.Panels[1].Text := Flist[i];
    for j := 1 to strtoint(edtStartString.Text) - 1 do
      readln(fReestr, s); //��������� �������� ������
    while not EOF(fReestr) do
    begin
      readln(fReestr, s);
      trn := strtotr(s);
      StatusBar1.Panels[1].Text := '������ ���� ' + inttostr(i + 1) + ' �� ' +
        inttostr(Flist.Count);
      Application.ProcessMessages;
      sleep(sleeptime);
      // ��������� ��� �� ������
      StatusBar1.Panels[1].Text := '��������� ��� �� ������';
      Application.ProcessMessages;
      sleep(sleeptime);
      if trn.id[1] <> '2' then
      begin
        StatusBar1.Panels[1].Text := '������ �� ���!';
        Application.ProcessMessages;
        sleep(sleeptime);
        writeln(fReport, s);
        leftpay := leftpay + 1;
        continue;
      end;
      // ��������� ��������� ���� ���� ������ ���� ������ ��� �������
      StatusBar1.Panels[1].Text := '��������� ������ �� ����';
      Application.ProcessMessages;
      sleep(sleeptime);
      try
        My1.Connect(edtSQLHost.Text, edtSQLUser.Text, edtSQLPwd.Text);
        My1.Execute('select id from UTM5.payment_transactions where payment_ext_number = ' + trn.num);
        if My1.RecordCount > 0 then
        begin
          My1.Close;
          ret := ret + 1;
          writeln(elog, DateTimeToStr(Now) + ' ��������� ��� ���� �������: ' +
            s);
          StatusBar1.Panels[1].Text := '��������� ��������� !!!';
          Application.ProcessMessages;
          sleep(sleeptime);
          continue;
        end;
        My1.Close;
      except
        writeln(elog, DateTimeToStr(Now) + '������ ����������� � MySql ��:' +
          edtSQLHost.Text);
        redtReport.Lines.Add('������ ����������� � MySql ��:' +
          edtSQLHost.Text + ' ��������� �����������!');
        Exit;
      end;

      // ������� ������ ������
      case StrToInt(trn.id[2]) of // ���� � ����������� ������
        3: account := copy(trn.id, 3, 5);
        //�� �������� �������� �����  �� �������� ID
        2: // ���� � �������� ��� ����������� �� �����
          begin // ������ � ���� ������ �������� ����
            StatusBar1.Panels[1].Text := '���� ID �� ����';
            Application.ProcessMessages;
            sleep(sleeptime);
            try
              My1.Connect(edtSQLHost.Text, edtSQLUser.Text, edtSQLPwd.Text);
              // 2200117  ->  177 -> tlm041117
              My1.Execute('select basic_account from UTM5.users where is_deleted = 0 and login = ''tlm041'
                + copy(trn.id, 5, 3) + '''');
              // ������ � ���� �������� ���� �� ������
              My1.First;
              account := My1.AsString(1);
              My1.Close;
            except
              writeln(elog, DateTimeToStr(Now) + '������ ����������� � MySql ��:'
                +
                edtSQLHost.Text);
              redtReport.Lines.Add('������ ����������� � MySql ��:' +
                edtSQLHost.Text + ' ��������� �����������!');
              Exit;
            end;
          end;
        1:
          begin
            StatusBar1.Panels[1].Text := '���� ID �� ����';
            Application.ProcessMessages;
            sleep(sleeptime);
            try
              My1.Connect(edtSQLHost.Text, edtSQLUser.Text, edtSQLPwd.Text);
              // 2100117  ->  177 -> chd006117
              My1.Execute('select basic_account from UTM5.users where is_deleted = 0 and login = ''chd006'
                + copy(trn.id, 5, 3) + ''''); // ������ � ����
              My1.First;
              account := My1.AsString(1);
              My1.Close;
            except
              writeln(elog, DateTimeToStr(Now) + '������ ����������� � MySql ��:'
                +
                edtSQLHost.Text);
              redtReport.Lines.Add('������ ����������� � MySql ��:' +
                edtSQLHost.Text + ' ��������� �����������!');
              Exit;
            end;
          end;
      else
        begin
          writeln(elog, DateTimeToStr(Now) + '����������� ������: ' + s);
          continue;
        end;
      end;
      // ������� ����� �������
      { /netup/utm5/bin/utm5_payment_tool -h ����
                                          -P ����
                                          -l �����
                                          -p ������
                                          -b ����� �������
                                          -a ������� ����
                                          -m ����� �������
                                          -L ���������� ������
                                          -e ����� ��������
                                          -t ����. ���� �������}
      SynTel.Send('/netup/utm5/bin/utm5_payment_tool -h ' +
        cbbUTMhost.Items[cbbUTMhost.ItemIndex]
        + ' -P 11758 -l ' + edtUTMUser.Text
        + ' -p ' + edtUTMPass.Text
        + ' -b ' + trn.summ
        + ' -a ' + inttostr(StrToInt(account))
        + ' -m 2 -L ''' + Flist[i]
        + ''' -e ' + trn.num
        + ' -t ' + IntToStr(DateTimeToUnix(StrToDate(Copy(trn.date, 1, 10))))
        +
        #10#13);
      // ��������� ������ ���������
      if not SynTel.WaitFor('successfully') then

      begin // ���� ������ �� ������ �������
        writeln(elog, DateTimeToStr(Now) + ' ��������� �� �������: ' + s);
        StatusBar1.Panels[1].Text := '������ !!!';
        Application.ProcessMessages;
        sleep(sleeptime);
        err := err + 1;
        if not SynTel.WaitFor(':~$') then //���� �� �������� �����������
        begin // �� ���� ��������� �������
          writeln(elog, DateTimeToStr(Now) + '������� �' + edtTLNHost.Text +
            ' ��������!');
          StatusBar1.Panels[0].Text := '������� �' + edtTLNHost.Text +
            ' ��������!';
          Exit;
        end;
        Continue;
      end;
      if not SynTel.WaitFor(':~$') then // ����������
      begin
        writeln(elog, DateTimeToStr(Now) + '������� �' + edtTLNHost.Text +
          ' ��������!');
        StatusBar1.Panels[0].Text := '������� �' + edtTLNHost.Text +
          ' ��������!';
        Exit;
      end;
      summ := summ + strtofloat(StringReplace(trn.summ, '.', ',',
        [rfReplaceAll])); // ������� ������� �������� �����
      totpay := totpay + 1; // ������� ����� ��������
      redtReport.Lines.Add('�� ����: ' + account + ' ������� ' + trn.summ +
        ' �� ' + trn.date);
    end;
    Flush(fReestr);
    Closefile(fReestr); //��������� ���������
    if chkIsDelete.Checked then
      DeleteFile(Flist[i]); // � ������� ���� �� �������
  end;
  StatusBar1.Panels[1].Text := '�������� ������ �������';
  Application.ProcessMessages;
  sleep(sleeptime);
  SynTel.Logout;
  StatusBar1.Panels[0].Text := '��������';
  Flush(fReport);
  Closefile(fReport); // ��������� ����� �������
  redtReport.Lines.Add('');
  redtReport.Lines.Add('����� �� ' + DateTimeToStr(Now) + ' :');
  redtReport.Lines.Add('������ ������� : ' + inttostr(Flist.Count));
  redtReport.Lines.Add('�������� ������� : ' + inttostr(totpay));
  redtReport.Lines.Add('����� : ' + floattostr(summ));
  redtReport.Lines.Add('�������� ����� : ' + floattostr(leftpay));
  redtReport.Lines.Add('������ : ' + inttostr(err));
  redtReport.Lines.Add('�������� : ' + inttostr(ret));
  redtReport.Lines.Add('������ : ' + floattostr(balance));
  redtReport.Lines.Add('������� �� ' + vipiska);
  Flist.Destroy;
  AssignFile(fReport, edtReportPath.Text + '������\�����' +
    StringReplace(DateTimeToStr(now), ':', '-', [rfReplaceAll]) +
    '.txt');
  CheckAndCreatePath(edtReportPath.Text + '������\�����' +
    StringReplace(DateTimeToStr(now), ':', '-', [rfReplaceAll]) +
    '.txt');
  Rewrite(fReport);
  for i := 1 to redtReport.Lines.Count do
    writeln(fReport, redtReport.Lines[i]);
  Flush(fReport);
  Closefile(fReport); // ���������
  StatusBar1.Panels[1].Text := '';
end;
//***************************************************************************

procedure TfrmUtmAuto.CreditClick(Sender: TObject); //������
label
  f1, f2;
var
  f, log: TextFile; // ����
  s: string;
  trn: TTran;
  err: integer;
  header: HWND; // ��� ������ ������� ����
  apchar: array[0..254] of char; // ��� ������ ������� ����
begin
  sleeptime := strtoint(edtSleepTime.Text);
  AssignFile(log, Logpath);
  Append(log);
  if open.Execute = false then
    goto f1;
  //*********************************
  ShellExecute(handle, nil, pchar(edtUtmPath.Text), nil, nil, SW_SHOWNORMAL);
  err := 0;
  while not (apchar = 'UTM5 �����������') or (apchar = 'UTM5 Login dialog') do
  begin
    err := err + 1;
    header := GetForegroundWindow; // �������� ��������� �������� ��������� ����
    GetWindowText(header, apchar, Length(apchar));
    if err > 10000000 then
    begin
      writeln(log, DateTimeToStr(Now) + ' ������ �������!!!');
      Application.Destroy;
    end;
  end;
  if (apchar = 'UTM5 �����������') then //���� ��� ������������
  begin
    sleep(sleeptime);
    mlclick(605, 588); // �� �� ������ �����
  end;
  if (apchar = 'UTM5 Login dialog') then // ���� ������ ���
  begin
    sleep(sleeptime);
    mlclick(509, 471); // ������ ����
    SendKeys(cbbUTMhost.Items[cbbUTMhost.ItemIndex]);
    sleep(sleeptime);
    mlclick(562, 504); // ������ �����
    SendKeys(edtUTMUser.Text);
    sleep(sleeptime);
    mlclick(564, 535); // ������ ������
    SendKeys(edtUTMPass.Text);
    sleep(sleeptime);
    mlclick(645, 583); //
    sleep(sleeptime); // �������� ����
    mlclick(634, 620); //
    sleep(sleeptime);
    mlclick(523, 603); //
    sleep(sleeptime); // ������ ������ ���������
    mlclick(521, 622); //
    sleep(sleeptime);
    mlclick(530, 564); // ��������� ���������
    sleep(sleeptime);
    mlclick(602, 590); // ��
  end;
  err := 0;
  while not (apchar = edtUTMUser.text + '@' +
    cbbUTMhost.Items[cbbUTMhost.ItemIndex] + ' (�������������)') do
  begin
    err := err + 1;
    header := GetForegroundWindow; // �������� ��������� �������� ��������� ����
    GetWindowText(header, apchar, Length(apchar));
    if err > 100000000 then
    begin
      writeln(log, DateTimeToStr(Now) + ' ������ �����������!!!');
      Close();
    end;
  end;
  sleep(sleeptime); //
  mlclick(935, 228); // ����� �� �������
  while not (apchar = '�����') do
  begin
    header := GetForegroundWindow; // �������� ��������� �������� ��������� ����
    GetWindowText(header, apchar, Length(apchar));
  end;
  sleep(sleeptime);
  //*********************************
  AssignFile(f, open.FileName);
  Reset(f);
  readln(f, s);
  //  if getlogin(s) <> 'ID ������������' then goto f1;
    //*********************************
  while not EOF(f) do
  begin
    readln(f, s);
    mlclick(357, 289); // �������� ���� �����
    sleep(sleeptime);
    SendKeys(getlogin(s)); //���� ����
    sleep(sleeptime);
    mlclick(959, 303); // �����
    sleep(sleeptime);
    mlclick(299, 476); // ��������
    sleep(sleeptime);
    mlclick(838, 423); // ������ ������
    err := 0;
    while not (apchar = '������ ������') do
    begin
      err := err + 1;
      header := GetForegroundWindow;
      // �������� ��������� �������� ��������� ����
      GetWindowText(header, apchar, Length(apchar));
      if err > 1000000 then
      begin
        writeln(log, DateTimeToStr(Now) + ' ������ ����� ������ ! ����:' +
          open.FileName + ' ������ ' + s);
        goto f2;
      end;
    end;
    mlclick(557, 404); // �������� ���� ����� �����
    sleep(sleeptime);
    SendKeys(edtCreditSumm.Text); //�����
    mlclick(572, 433);
    sleep(sleeptime);
    SendKeys('30.12.2010 23:59:00'); //��������
    mlclick(500, 400);
    mlclick(572, 433);
    sleep(sleeptime);
    SendKeys(DateTimeToStr(Date()) + ' 00:00:00'); //����
    mlclick(565, 514);
    sleep(sleeptime);
    mlclick(357, 289);
    sleep(sleeptime);
    mlclick(706, 461);
    sleep(sleeptime);
    mlclick(562, 484);
    sleep(sleeptime);
    SendKeys('2011'); // year
    mlclick(622, 510);
    sleep(sleeptime);
    mlclick(585, 529);
    sleep(sleeptime);
    mlclick(567, 535);
    sleep(sleeptime);
    SendKeys('12'); ////month
    mlclick(730, 631);
    sleep(sleeptime);
    mlclick(861, 569);
    sleep(sleeptime);
    mlclick(837, 638);
    sleep(sleeptime);
    mlclick(748, 715);
    while not (apchar = '�����') do
    begin
      header := GetForegroundWindow;
      // �������� ��������� �������� ��������� ����
      GetWindowText(header, apchar, Length(apchar));
    end;
    sleep(sleeptime);
    mlclick(961, 337);
    f2: sleep(sleeptime);
  end;
  f1: mlclick(1023, 228); // ��������� netup
  sleep(sleeptime);
  mlclick(1166, 162);
  sleep(sleeptime);
  Closefile(f); //���������
  Flush(log);
  Closefile(log); //��������� ���
end;

procedure TfrmUtmAuto.FormCreate(Sender: TObject);
begin
  curentdate := StringReplace(DateTimeToStr(Date), ':', '-', [rfReplaceAll]);
  sleeptime := strtoint(edtSleepTime.Text);
  Logpath := ExtractFilePath(Paramstr(0)) + curentdate + '.log';
  eLogpath := ExtractFilePath(Paramstr(0)) + curentdate + ' errors.log';
  AssignFile(log, Logpath);
  if FileExists(Logpath) then
    Append(log)
  else
    Rewrite(log);
  AssignFile(elog, eLogpath);
  if FileExists(eLogpath) then
    Append(elog)
  else
    Rewrite(elog);
  ini := Tinifile.Create(ExtractFilePath(Paramstr(0)) + 'config.ini');
  ini.ReadSections(cbbUTMHost.Items);
  if cbbUTMHost.Items.Count > 0 then
  begin
    cbbUTMHost.ItemIndex := 0;
    ReadINI(cbbUTMHost.Items[0]);
  end;
  StatusBar1.Panels[0].Text := '��������';
end;

procedure TfrmUtmAuto.CheckClick(Sender: TObject); //��������
var
  Flist: Tstringlist; // ������ ������ � ����������
  f, dump: TextFile;
  i, err, j, totpay, leftpay, sleeptime: integer;
  s, tmpstr, vipiska, ds: string;
  trn: TTran; // ����� ���� ���� �����
  summ, balance, leftSumm: Extended; // ����� ���� ���������
begin
  Flist := TStringList.Create;
  FindFile(edtPaymentPath.Text, Flist);
  // ����� ������ � ���������� ���������, �������� � Edit1
  for i := 0 to (Flist.Count - 1) do
  begin
    vipiska := getdate(Flist[i]);
    tmpstr := copy(Flist[i], length(Flist[i]) - 3, 4);
    if tmpstr <> '.txt' then
      continue;
    AssignFile(f, Flist[i]);
    Reset(f);
    for j := 1 to strtoint(edtStartString.Text) - 1 do
      readln(f, s);
    while not EOF(f) do
    begin
      readln(f, s);
      trn := strtotr(s);
      if trn.id[1] <> '2' then
      begin
        leftpay := leftpay + 1;
        leftSumm := leftSumm + strtofloat(StringReplace(trn.summ, '.', ',',
          [rfReplaceAll]));
        continue;
      end;
      err := 0;
      summ := summ + strtofloat(StringReplace(trn.summ, '.', ',',
        [rfReplaceAll]));
      totpay := totpay + 1;
    end;
    Flush(f);
    Closefile(f); //��������� ���������
  end;
  redtReport.Lines.Add('�����');
  redtReport.Lines.Add('������ ������� : ' + inttostr(Flist.Count));
  redtReport.Lines.Add('�������� ������� : ' + inttostr(totpay));
  redtReport.Lines.Add('����� : ' + floattostr(summ));
  redtReport.Lines.Add('����� �������� : ' + inttostr(leftpay));
  redtReport.Lines.Add('����� ����� �������� : ' + floattostr(leftSumm));
  Flist.Destroy;
end;

procedure TfrmUtmAuto.BalanceClick(Sender: TObject); //������
var
  ds: string;
  balance: Extended; // ����� ���� ���������
  k: integer;
  My1: TMySQL5;
begin
  My1 := TMySQL5.Create(nil);
  try
    begin
      My1.Connect(edtSQLHost.Text, edtUTMUser.Text, edtSQLPwd.Text);
      My1.Execute('select balance from UTM5.accounts where is_deleted = 0');
      balance := 0;
      for k := 1 to My1.RecordCount do
      begin
        My1.Move(k);
        balance := balance + StrToFloat(StringReplace(My1.AsString(1), '.', ',',
          [rfReplaceAll]));
      end;
      My1.Close; // Close connection to MySQL server.
    end;
  except
    begin
      balance := 0;
      writeln(elog, DateTimeToStr(Now) +
        '�� ������� ������������ � SQL �������. �� ������� ��������� ������!!!');
      redtReport.Lines.Add('�� ������� ��������� ������!!!');
    end;
  end;
  redtReport.Lines.Add('������ : ' + floattostr(balance));

end;

//**********************************************************************

procedure TfrmUtmAuto.StartUTMClick(Sender: TObject);
begin
  ShellExecute(handle, nil, pchar(edtUtmPath.Text), nil, nil, SW_SHOWNORMAL);
end;
//**********************************************************************

procedure TfrmUtmAuto.ReadINI(Section: string);
begin
  //utm
  edtUTMHost.Text := Section;
  edtUTMUser.Text := ini.ReadString(Section, 'UTMUser', '');
  edtUTMPass.Text := ini.ReadString(Section, 'UTMPass', '');
  edtUtmPath.Text := ini.ReadString(Section, 'UtmPath', '');
  //sql
  edtSQLHost.Text := ini.ReadString(Section, 'SQLHost', '');
  edtSQLUser.Text := ini.ReadString(Section, 'SQLUser', '');
  edtSQLPwd.Text := ini.ReadString(Section, 'SQLPwd', '');
  //ssh
  edtTLNHost.Text := ini.ReadString(Section, 'TLNHost', '');
  edtTLNUser.Text := ini.ReadString(Section, 'TLNUser', '');
  edtTLNPwd.Text := ini.ReadString(Section, 'TLNPwd', '');
  //file
  edtSleepTime.Text := ini.ReadString(Section, 'SleepTime', '');
  edtStartString.Text := ini.ReadString(Section, 'StartString', '');
  edtPaymentPath.Text := ini.ReadString(Section, 'PaymentPath', '');
  edtReportPath.Text := ini.ReadString(Section, 'ReportPath', '');
  //cred
  edtCreditSumm.Text := ini.ReadString(Section, 'CreditSumm', '');
  rgPrSel.ItemIndex := ini.ReadInteger(Section, 'Protocol', 0);
  chkIsDelete.Checked := ini.ReadBool(Section, 'DeleteAfter', True);
end;

procedure TfrmUtmAuto.WriteINI(Section: string);
begin
  //utm
  ini.WriteString(Section, 'UTMUser', edtUTMUser.Text);
  ini.WriteString(Section, 'UTMPass', edtUTMPass.Text);
  ini.WriteString(Section, 'UtmPath', edtUtmPath.Text);
  //sql
  ini.WriteString(Section, 'SQLHost', edtSQLHost.Text);
  ini.WriteString(Section, 'SQLUser', edtSQLUser.Text);
  ini.WriteString(Section, 'SQLPwd', edtSQLPwd.Text);
  //ssh
  ini.WriteString(Section, 'TLNHost', edtTLNHost.Text);
  ini.WriteString(Section, 'TLNUser', edtTLNUser.Text);
  ini.WriteString(Section, 'TLNPwd', edtTLNPwd.Text);
  //file
  ini.WriteString(Section, 'SleepTime', edtSleepTime.Text);
  ini.WriteString(Section, 'StartString', edtStartString.Text);
  ini.WriteString(Section, 'PaymentPath', edtPaymentPath.Text);
  ini.WriteString(Section, 'ReportPath', edtReportPath.Text);
  //other
  ini.WriteString(Section, 'CreditSumm', edtCreditSumm.Text);
  ini.WriteInteger(Section, 'Protocol', rgPrSel.ItemIndex);
  ini.WriteBool(Section, 'DeleteAfter', chkIsDelete.Checked);
end;

procedure TfrmUtmAuto.ClosePrClick(Sender: TObject);
begin
  Close();
end;

procedure TfrmUtmAuto.cbbUTMhostChange(Sender: TObject);
begin
  ReadINI(cbbUTMhost.Items[cbbUTMhost.ITemIndex]);
end;

procedure TfrmUtmAuto.btnAddHostClick(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to cbbUTMhost.Items.count - 1 do
    if cbbUTMhost.Items[i] = edtUTMHost.Text then
    begin
      WriteINI(edtUTMHost.Text);
      Exit;
    end;
  cbbUTMhost.Items.Add(edtUTMHost.Text);
  WriteINI(edtUTMHost.Text);
  cbbUTMhost.ItemIndex := cbbUTMhost.Items.count - 1;
end;

procedure TfrmUtmAuto.InsertOldClick(Sender: TObject);
// ������ ������� �� �������
label
  e2; // ��� ��� ����� ��� ��������� �����
var
  fReestr, fReport, dump: TextFile; //
  s, tmpstr, vipiska, ds: string;
  trn: TTran; // ����� ���� ���� �����
  Flist: Tstringlist; // ������ ������ � ����������
  i, err, j, totpay, leftpay, sleeptime: integer;
  // (i,j)�������  err - ��� ������ �� ������������ �����
  header: HWND; // ��� ������ ������� ����
  apchar: array[0..254] of char; // ��� ������ ������� ����
  summ, balance: Extended; // ����� ���� ���������, ������
begin
  ShellExecute(handle, nil, pchar(cbbUTMHost.Text), nil, nil, SW_SHOWNORMAL);
  err := 0;
  while not (apchar = 'UTM5 �����������') or (apchar = 'UTM5 Login dialog') do
  begin
    err := err + 1;
    header := GetForegroundWindow;
    // �������� ��������� �������� ��������� ����
    GetWindowText(header, apchar, Length(apchar));
    if err > 100000000 then
    begin
      MessageBox(0, ' ������ �������!!!', '��������', MB_OK);
      Exit;
    end;
  end;
  if (apchar = 'UTM5 �����������') then //���� ��� ������������
  begin
    sleep(sleeptime);
    mlclick(605, 588); // �� �� ������ �����
  end;
  if (apchar = 'UTM5 Login dialog') then // ���� ������ ���
  begin
    sleep(sleeptime);
    mlclick(509, 471); // ������ ����
    SendKeys(cbbUTMhost.Text);
    sleep(sleeptime);
    mlclick(562, 504); // ������ �����
    SendKeys(edtUTMUser.Text);
    sleep(sleeptime);
    mlclick(564, 535); // ������ ������
    SendKeys(edtUTMPass.Text);
    sleep(sleeptime);
    mlclick(645, 583); //
    sleep(sleeptime); // �������� ����
    mlclick(634, 620); //
    sleep(sleeptime);
    mlclick(523, 603); //
    sleep(sleeptime); // ������ ������ ���������
    mlclick(521, 622); //
    sleep(sleeptime);
    mlclick(530, 564); // ��������� ���������
    sleep(sleeptime);
    mlclick(602, 590); // ��
  end;
  err := 0;
  while not (apchar = edtUTMUser.text + '@' + cbbUTMhost.text +
    ' (�������������)') do
  begin
    err := err + 1;
    header := GetForegroundWindow;
    // �������� ��������� �������� ��������� ����
    GetWindowText(header, apchar, Length(apchar));
    if err > 10000000 then
    begin
      MessageBox(0, '������ �����������!!!', '��������', MB_OK);
      Exit;
    end;
  end;
  sleep(sleeptime); // dumping ......
  // ---------------------------------------------------------
  balance := 0;
  mrclick(650, 400);
  sleep(sleeptime);
  mlclick(700, 430);
  sleep(sleeptime);
  mlclick(505, 588);
  SendKeys(ExtractFilePath(Paramstr(0)) + 'dump.csv');
  PressKey(13);
  sleep(sleeptime + 2000);
  AssignFile(dump, ExtractFilePath(Paramstr(0)) + 'dump.csv');
  Reset(dump);
  readln(dump, ds);
  while not EOF(dump) do
  begin
    readln(dump, ds);
    balance := balance + getbalance(ds);
  end;
  Closefile(dump);
  DeleteFile(ExtractFilePath(Paramstr(0)) + 'dump.csv');
  //--------------------------------------------------------
  mlclick(935, 228); // ����� �� �������
  while not (apchar = '�����') do
  begin
    header := GetForegroundWindow;
    // �������� ��������� �������� ��������� ����
    GetWindowText(header, apchar, Length(apchar));
  end;
  sleep(sleeptime);
  //****************************************************************************
  AssignFile(fReport, edtReportPath.Text + '�������\������� �� ' +
    StringReplace(DateTimeToStr(now), ':', '-', [rfReplaceAll]) +
    '.txt'); //����� - ������������� �� 2 ������: 1100117
  CheckAndCreatePath(edtReportPath.Text + '�������\������� �� ' +
    StringReplace(DateTimeToStr(now), ':', '-', [rfReplaceAll]) +
    '.txt');
  Rewrite(fReport);
  summ := 0;
  totpay := 0;
  leftpay := 0;
  //****************************************************************************
  Flist := TStringList.Create;
  FindFile(edtPaymentPath.Text, Flist);
  // ����� ������ � ���������� ���������, �������� � Edit1
  for i := 0 to (Flist.Count - 1) do
  begin
    vipiska := getdate(Flist[i]);
    tmpstr := copy(Flist[i], length(Flist[i]) - 3, 4);
    if tmpstr <> '.txt' then
      continue;
    AssignFile(fReestr, Flist[i]);
    Reset(fReestr);
    for j := 1 to strtoint(edtStartString.Text) - 1 do
      readln(fReestr, s);
    while not EOF(fReestr) do
    begin
      readln(fReestr, s);
      trn := strtotr(s);
      case strtoint(trn.id[1]) of
        3:
          begin
            mlclick(715, 349); // �������� ���� �����
            sleep(sleeptime);
            SendKeys(inttostr(strtoint(copy(trn.id, 3, 5))));
            //�������� ������� ����
          end;
        2:
          begin
            mlclick(360, 289); // �������� ���� �����
            sleep(sleeptime);
            SendKeys('tlm041' + copy(trn.id, 5, 3)); //���� ����
          end;
        1:
          begin
            mlclick(360, 289); // �������� ���� �����
            sleep(sleeptime);
            SendKeys('chd006' + copy(trn.id, 5, 3)); //���� ����
          end;
      else
        begin
          writeln(left, s);
          leftpay := leftpay + 1;
          continue;
        end;
      end;
      sleep(sleeptime);
      mlclick(959, 303); // �����
      sleep(sleeptime);
      mlclick(299, 476); // ��������
      sleep(sleeptime);
      mlclick(838, 423); // ������ ������
      err := 0;
      while not (apchar = '������ ������') do
      begin
        err := err + 1;
        header := GetForegroundWindow;
        // �������� ��������� �������� ��������� ����
        GetWindowText(header, apchar, Length(apchar));
        if err > 1000000 then
        begin
          writeln(elog, DateTimeToStr(Now) + ' ������ ����� ������ ! ����:' +
            Flist[i] + ' ������ �' + inttostr(i + 1) + ' ������� ' + trn.id +
            ' ����� ' + trn.summ + ' �� ����: ' + trn.date + ' �� ������� �' +
            trn.num);
          goto e2;
        end;
      end;
      sleep(sleeptime);
      mlclick(557, 404); // �������� ���� ����� �����
      sleep(sleeptime);
      SendKeys(trn.summ);
      summ := summ + strtofloat(StringReplace(trn.summ, '.', ',',
        [rfReplaceAll]));
      mlclick(572, 433); // �������� ���� ����� ����
      sleep(sleeptime);
      SendKeys('11012001000000'); //��������
      mlclick(500, 400);
      mlclick(572, 433);
      sleep(sleeptime);
      SendKeys(trn.date); //����
      mlclick(565, 514); //��� ��������������
      sleep(sleeptime);
      SendKeys(Flist[i] + ' Date : ' + trn.date); //��� ��������������
      mlclick(619, 571); //
      sleep(sleeptime); // ��������
      mlclick(681, 599); // ������
      sleep(sleeptime); // ������
      mlclick(563, 596); //����� ��������
      sleep(sleeptime);
      SendKeys(trn.num);
      mlclick(748, 715);
      totpay := totpay + 1;
      while not (apchar = '�����') do
      begin
        header := GetForegroundWindow;
        // �������� ��������� �������� ��������� ����
        GetWindowText(header, apchar, Length(apchar));
      end;
      sleep(sleeptime);
      e2: mlclick(961, 337);
      sleep(sleeptime);
    end;
    Flush(fReestr);
    Closefile(fReestr); //��������� ���������
    if chkIsDelete.Checked then
      DeleteFile(Flist[i]); // � ������� ���� �� �������
  end;
  mlclick(1023, 228);
  sleep(sleeptime);
  mlclick(1166, 162);
  sleep(sleeptime);
  Flush(fReport);
  Closefile(fReport); // ��������� ����� �������
  redtReport.Lines.Add('');
  redtReport.Lines.Add('������� ������� ������ ��������:');
  redtReport.Lines.Add('����� : ' + floattostr(summ));
  redtReport.Lines.Add('������ ������� : ' + inttostr(Flist.Count));
  redtReport.Lines.Add('�������� ������� : ' + inttostr(totpay));
  redtReport.Lines.Add('������ : ' + inttostr(err));
  redtReport.Lines.Add('������ : ' + floattostr(balance));
  redtReport.Lines.Add('������� �� ' + vipiska);
  Flist.Destroy;
  AssignFile(fReport, edtReportPath.Text + '������\�����' + curentdate +
    '.txt');
  // ���������� - ������� �������� ����� ����
  Rewrite(fReport);
  for i := 1 to redtReport.Lines.Count do
    writeln(fReport, redtReport.Lines[i]);
  Flush(fReport);
  Closefile(fReport); // ���������
end;

procedure TfrmUtmAuto.edtSleepTimeChange(Sender: TObject);
begin
  sleeptime := strtoint(edtSleepTime.Text);
end;

procedure TfrmUtmAuto.SetUtmPathClick(Sender: TObject);
var
  dir: string;
begin
  open.DefaultExt := 'jar';
  if open.Execute then
    edtUtmPath.Text := open.FileName;
end;

procedure TfrmUtmAuto.SetPaymentPathClick(Sender: TObject);
var
  dir: string;
begin
  try
    if SelectDirectory(Dir, [sdAllowCreate, sdPerformCreate, sdPrompt], 0) then
      edtPaymentPath.Text := Dir + '\'
    else
      edtPaymentPath.Text := ExtractFilePath(Paramstr(0));
  except
    edtPaymentPath.Text := ExtractFilePath(Paramstr(0));
  end;

end;

procedure TfrmUtmAuto.SetReportPathClick(Sender: TObject);
var
  dir: string;
begin
  try
    if SelectDirectory(Dir, [sdAllowCreate, sdPerformCreate, sdPrompt], 0) then
      edtReportPath.Text := Dir + '\'
    else
      edtReportPath.Text := ExtractFilePath(Paramstr(0));
  except
    edtReportPath.Text := ExtractFilePath(Paramstr(0));
  end;

end;

procedure TfrmUtmAuto.btnCloneHostClick(Sender: TObject);
begin
  edtTLNHost.Text := edtUTMHost.Text; // ��� ��� ��������
  edtSQLHost.Text := edtUTMHost.Text;
end;

procedure TfrmUtmAuto.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Flush(log);
  Closefile(log); //��������� ���
  Flush(elog);
  Closefile(elog);
end;

procedure TfrmUtmAuto.HahdInsertClick(Sender: TObject);
begin
  frmMiniUtm.ShowModal;
end;

procedure TfrmUtmAuto.mniSQLQueryClick(Sender: TObject);
begin
  Form1.ShowModal;
end;

end.

