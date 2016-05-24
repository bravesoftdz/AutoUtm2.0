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
  log, elog: TextFile; // типа лог
  Ressive: TStringList; // полученное из телнета
  ini: Tinifile; // конфиг
  sleeptime: Integer;
implementation
uses MySQL5, AutoUtm2Procs, MainMiniUtm, Unit1;
{$R *.dfm}
//*************************service*******************************************

//******************************************************************************
  {                     ===
     \\ //   \\   //  ||  /||  ну или кагбэ начало основной части программы!!!
      \\/     \\ //   || //||  ололо ничего оригинальней не придумал!!!
     //\\      \//    ||// ||
    //  \\     //     ||/  || }
//**********************MAIN****************************************************

procedure TfrmUtmAuto.InsertClick(Sender: TObject);
var
  fReestr, fReport: TextFile; //
  s, tmpstr, vipiska, ds, currenthostname, account: string;
  trn: TTran; // сумма юзер дата номер
  Flist: Tstringlist; // список файлов с платежками
  i, j, totpay, leftpay, k, err, ret: integer; // (i,j)счетчик
  summ, balance: Extended; // сумма всех внесенных
  My1: TMySQL5;
  SynTel: TTelnetSend; // телнет клиент synapce
begin
  StatusBar1.Panels[1].Text := '';
  Application.ProcessMessages;
  SynTel := TTelnetSend.Create;
  SynTel.TargetHost := edtTLNHost.Text;
  SynTel.TermType := 'dumb'; // тип терминала
  StatusBar1.Panels[1].Text := 'Подключение...';
  Application.ProcessMessages;
  case rgPrSel.ItemIndex of
    0: // если телнет
      begin
        SynTel.TargetPort := '23';
        if not SynTel.Login then
        begin
          writeln(elog, DateTimeToStr(Now) + 'Не подключиться к ' +
            edtTLNHost.Text);
          StatusBar1.Panels[1].Text := 'Не подключиться к ' +
            edtTLNHost.Text;
          StatusBar1.Panels[0].Text := 'Отключен';
          Exit;
        end
        else
          writeln(log, DateTimeToStr(Now) + 'Подключились к ' +
            edtTLNhost.Text);
        StatusBar1.Panels[0].Text := edtTLNHost.Text + ' Подключен';
        Application.ProcessMessages;
        sleep(sleeptime);
        if not SynTel.WaitFor('login:') then
        begin
          writeln(elog, DateTimeToStr(Now) + 'Не ввести логин ' +
            edtTLNHost.Text);
          redtReport.Lines.Add(DateTimeToStr(Now));
          StatusBar1.Panels[1].Text := 'Ошибка протокола (login) ';
          StatusBar1.Panels[0].Text := 'Отключен';
          Exit;
        end;
        SynTel.Send(edtTLNUser.Text + #10#13);
        if not SynTel.WaitFor('Password:') then
        begin
          writeln(elog, DateTimeToStr(Now) + 'Не ввести пароль ' +
            edtTLNHost.Text);
          StatusBar1.Panels[1].Text := 'Ошибка протокола (Password) ';
          StatusBar1.Panels[0].Text := 'Отключен';
          Exit;
        end;
        SynTel.Send(edtTLNPwd.Text + #10#13);
        if not SynTel.WaitFor(':~$') then
        begin
          writeln(elog, DateTimeToStr(Now) + 'Нет (Prompt) ' +
            edtTLNHost.Text);
          StatusBar1.Panels[1].Text := 'Ошибка протокола (Prompt) ';
          StatusBar1.Panels[0].Text := 'Отключен';
          Exit;
        end;
      end;
    1: // если ssh
      begin
        SynTel.TargetPort := '22';
        SynTel.UserName := edtTLNUser.Text;
        SynTel.Password := edtTLNPwd.Text;
        // SynTel.Login;
        if not SynTel.SSHLogin then
        begin
          writeln(elog, DateTimeToStr(Now) + 'Не подключиться к ' +
            edtTLNhost.Text);
          Exit;
        end
        else
          writeln(log, DateTimeToStr(Now) + 'Подключились к ' +
            edtTLNhost.Text);
        StatusBar1.Panels[0].Text := edtTLNHost.Text + ' Подключен';
        Application.ProcessMessages;
        sleep(sleeptime);
        if not SynTel.WaitFor(':~$') then
        begin
          writeln(elog, DateTimeToStr(Now) + 'Нет (Prompt) ' +
            edtTLNHost.Text);
          StatusBar1.Panels[1].Text := 'Ошибка протокола (Prompt) ';
          StatusBar1.Panels[0].Text := 'Отключен';
          Exit;
        end;
      end;
  else // если ошибка
    writeln(elog, DateTimeToStr(Now) + ' внутренняя ошибка #0001');
  end;
  // считаем баланс ------------------------------------------------
  StatusBar1.Panels[1].Text := 'Считаем баланс...';
  Application.ProcessMessages;
  sleep(sleeptime);
  My1 := TMySQL5.Create(nil);
  try
    My1.Connect(edtSQLHost.Text, edtSQLUser.Text, edtSQLPwd.Text);
    My1.Execute('select balance from UTM5.accounts where is_deleted = 0');
  except
    writeln(elog, DateTimeToStr(Now) + 'Ошибка подключения к MySql на:' +
      edtSQLHost.Text);
    redtReport.Lines.Add('Ошибка подключения к MySql на:' +
      edtSQLHost.Text + ' программа остановлена!');
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
    writeln(elog, DateTimeToStr(Now) + 'не удалось посчитать баланс!!!');
  end;
  My1.Close; // Close connection to MySQL server.
  //****************открываем логи платежей *************************
  StatusBar1.Panels[1].Text := 'Создаем файлы отчетов..';
  Application.ProcessMessages;
  sleep(sleeptime);
  AssignFile(fReport, edtReportPath.Text + 'платежи\платежи на ' +
    StringReplace(DateTimeToStr(now), ':', '-', [rfReplaceAll]) +
    '.txt'); //левые - идентификатор не 2 пример: 1100117
  CheckAndCreatePath(edtReportPath.Text + 'платежи\платежи на ' +
    StringReplace(DateTimeToStr(now), ':', '-', [rfReplaceAll]) +
    '.txt');
  Rewrite(fReport);
  //****************************************************************************
  StatusBar1.Panels[1].Text := 'Ищем файлы платежей';
  Application.ProcessMessages;
  sleep(sleeptime);
  Flist := TStringList.Create;
  FindFile(edtPaymentPath.Text, Flist);
  // поиск файлов с начальными условиями, заданных в Edit1
  // обнуляем счетчики
  summ := 0; // сумма всех введенных платежей
  totpay := 0; // всего платежей (штук)
  leftpay := 0; // левых
  ret := 0; //повторных
  err := 0; // ошибок ввода
  vipiska := ''; // не факт что нужно но так спокойнее
  for i := 0 to (Flist.Count - 1) do
  begin
    if vipiska = '' then
      vipiska := getdate(Flist[i])
    else if vipiska <> getdate(Flist[i]) then
    begin
      redtReport.Lines.Add('');
      redtReport.Lines.Add('отчет на ' + DateTimeToStr(Now) + ' :');
      redtReport.Lines.Add('Файлов введено : ' + inttostr(Flist.Count));
      redtReport.Lines.Add('Платежей введено : ' + inttostr(totpay));
      redtReport.Lines.Add('Сумма : ' + floattostr(summ));
      redtReport.Lines.Add('Платежей левых : ' + floattostr(leftpay));
      redtReport.Lines.Add('Ошибок : ' + inttostr(err));
      redtReport.Lines.Add('Повторов : ' + inttostr(ret));
      redtReport.Lines.Add('Баланс : ' + floattostr(balance));
      redtReport.Lines.Add('выписка от ' + vipiska);
      summ := 0; // сумма всех введенных платежей
      totpay := 0; // всего платежей (штук)
      leftpay := 0; // левых
      ret := 0; //повторных
      err := 0; // ошибок ввода
      vipiska := getdate(Flist[i]);
    end;
    tmpstr := copy(Flist[i], length(Flist[i]) - 3, 4);
    if tmpstr <> '.txt' then
      continue;
    AssignFile(fReestr, Flist[i]);
    Reset(fReestr);
    StatusBar1.Panels[1].Text := Flist[i];
    for j := 1 to strtoint(edtStartString.Text) - 1 do
      readln(fReestr, s); //считываем ненужные строки
    while not EOF(fReestr) do
    begin
      readln(fReestr, s);
      trn := strtotr(s);
      StatusBar1.Panels[1].Text := 'Вводим файл ' + inttostr(i + 1) + ' из ' +
        inttostr(Flist.Count);
      Application.ProcessMessages;
      sleep(sleeptime);
      // проверяем наш ли платеж
      StatusBar1.Panels[1].Text := 'Проверяем наш ли платеж';
      Application.ProcessMessages;
      sleep(sleeptime);
      if trn.id[1] <> '2' then
      begin
        StatusBar1.Panels[1].Text := 'Платеж не наш!';
        Application.ProcessMessages;
        sleep(sleeptime);
        writeln(fReport, s);
        leftpay := leftpay + 1;
        continue;
      end;
      // проверяем повторный внос если запись есть значит уже вносили
      StatusBar1.Panels[1].Text := 'Проверяем повтор по базе';
      Application.ProcessMessages;
      sleep(sleeptime);
      try
        My1.Connect(edtSQLHost.Text, edtSQLUser.Text, edtSQLPwd.Text);
        My1.Execute('select id from UTM5.payment_transactions where payment_ext_number = ' + trn.num);
        if My1.RecordCount > 0 then
        begin
          My1.Close;
          ret := ret + 1;
          writeln(elog, DateTimeToStr(Now) + ' Квитанция уже была внесена: ' +
            s);
          StatusBar1.Panels[1].Text := 'Повторная квитанция !!!';
          Application.ProcessMessages;
          sleep(sleeptime);
          continue;
        end;
        My1.Close;
      except
        writeln(elog, DateTimeToStr(Now) + 'Ошибка подключения к MySql на:' +
          edtSQLHost.Text);
        redtReport.Lines.Add('Ошибка подключения к MySql на:' +
          edtSQLHost.Text + ' программа остановлена!');
        Exit;
      end;

      // смртрим откуда платеж
      case StrToInt(trn.id[2]) of // если с московского района
        3: account := copy(trn.id, 3, 5);
        //то получаем лицевого счета  из внешнего ID
        2: // если с тельмана или чудновского то через
          begin // запрос в базу узнаем основной счет
            StatusBar1.Panels[1].Text := 'Ищем ID по базе';
            Application.ProcessMessages;
            sleep(sleeptime);
            try
              My1.Connect(edtSQLHost.Text, edtSQLUser.Text, edtSQLPwd.Text);
              // 2200117  ->  177 -> tlm041117
              My1.Execute('select basic_account from UTM5.users where is_deleted = 0 and login = ''tlm041'
                + copy(trn.id, 5, 3) + '''');
              // запрос в базу получаем счет по логину
              My1.First;
              account := My1.AsString(1);
              My1.Close;
            except
              writeln(elog, DateTimeToStr(Now) + 'Ошибка подключения к MySql на:'
                +
                edtSQLHost.Text);
              redtReport.Lines.Add('Ошибка подключения к MySql на:' +
                edtSQLHost.Text + ' программа остановлена!');
              Exit;
            end;
          end;
        1:
          begin
            StatusBar1.Panels[1].Text := 'Ищем ID по базе';
            Application.ProcessMessages;
            sleep(sleeptime);
            try
              My1.Connect(edtSQLHost.Text, edtSQLUser.Text, edtSQLPwd.Text);
              // 2100117  ->  177 -> chd006117
              My1.Execute('select basic_account from UTM5.users where is_deleted = 0 and login = ''chd006'
                + copy(trn.id, 5, 3) + ''''); // запрос в базу
              My1.First;
              account := My1.AsString(1);
              My1.Close;
            except
              writeln(elog, DateTimeToStr(Now) + 'Ошибка подключения к MySql на:'
                +
                edtSQLHost.Text);
              redtReport.Lines.Add('Ошибка подключения к MySql на:' +
                edtSQLHost.Text + ' программа остановлена!');
              Exit;
            end;
          end;
      else
        begin
          writeln(elog, DateTimeToStr(Now) + 'неизвестный платеж: ' + s);
          continue;
        end;
      end;
      // команда вноса платежа
      { /netup/utm5/bin/utm5_payment_tool -h хост
                                          -P порт
                                          -l логин
                                          -p пароль
                                          -b сумма плалежа
                                          -a лицевой счет
                                          -m метод платежа
                                          -L коментарий админа
                                          -e номер платежки
                                          -t факт. дата платежа}
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
      // обработка ответа терминала
      if not SynTel.WaitFor('successfully') then

      begin // если платеж не принял биллинг
        writeln(elog, DateTimeToStr(Now) + ' Квитанция не внесена: ' + s);
        StatusBar1.Panels[1].Text := 'Ошибка !!!';
        Application.ProcessMessages;
        sleep(sleeptime);
        err := err + 1;
        if not SynTel.WaitFor(':~$') then //если не получено приглашение
        begin // на ввод следующей команды
          writeln(elog, DateTimeToStr(Now) + 'Коннект с' + edtTLNHost.Text +
            ' разорван!');
          StatusBar1.Panels[0].Text := 'Коннект с' + edtTLNHost.Text +
            ' разорван!';
          Exit;
        end;
        Continue;
      end;
      if not SynTel.WaitFor(':~$') then // аналогично
      begin
        writeln(elog, DateTimeToStr(Now) + 'Коннект с' + edtTLNHost.Text +
          ' разорван!');
        StatusBar1.Panels[0].Text := 'Коннект с' + edtTLNHost.Text +
          ' разорван!';
        Exit;
      end;
      summ := summ + strtofloat(StringReplace(trn.summ, '.', ',',
        [rfReplaceAll])); // считаем сколько внеслось денег
      totpay := totpay + 1; // сколько всего платежей
      redtReport.Lines.Add('на счет: ' + account + ' внесено ' + trn.summ +
        ' от ' + trn.date);
    end;
    Flush(fReestr);
    Closefile(fReestr); //закрываем введенный
    if chkIsDelete.Checked then
      DeleteFile(Flist[i]); // и удаляем чтоб не мешался
  end;
  StatusBar1.Panels[1].Text := 'Создание файлов отчетов';
  Application.ProcessMessages;
  sleep(sleeptime);
  SynTel.Logout;
  StatusBar1.Panels[0].Text := 'Отключен';
  Flush(fReport);
  Closefile(fReport); // закрываем левые платежи
  redtReport.Lines.Add('');
  redtReport.Lines.Add('отчет на ' + DateTimeToStr(Now) + ' :');
  redtReport.Lines.Add('Файлов введено : ' + inttostr(Flist.Count));
  redtReport.Lines.Add('Платежей введено : ' + inttostr(totpay));
  redtReport.Lines.Add('Сумма : ' + floattostr(summ));
  redtReport.Lines.Add('Платежей левых : ' + floattostr(leftpay));
  redtReport.Lines.Add('Ошибок : ' + inttostr(err));
  redtReport.Lines.Add('Повторов : ' + inttostr(ret));
  redtReport.Lines.Add('Баланс : ' + floattostr(balance));
  redtReport.Lines.Add('выписка от ' + vipiska);
  Flist.Destroy;
  AssignFile(fReport, edtReportPath.Text + 'отчеты\отчет' +
    StringReplace(DateTimeToStr(now), ':', '-', [rfReplaceAll]) +
    '.txt');
  CheckAndCreatePath(edtReportPath.Text + 'отчеты\отчет' +
    StringReplace(DateTimeToStr(now), ':', '-', [rfReplaceAll]) +
    '.txt');
  Rewrite(fReport);
  for i := 1 to redtReport.Lines.Count do
    writeln(fReport, redtReport.Lines[i]);
  Flush(fReport);
  Closefile(fReport); // закрываем
  StatusBar1.Panels[1].Text := '';
end;
//***************************************************************************

procedure TfrmUtmAuto.CreditClick(Sender: TObject); //кредит
label
  f1, f2;
var
  f, log: TextFile; // файл
  s: string;
  trn: TTran;
  err: integer;
  header: HWND; // для поиска нужного окна
  apchar: array[0..254] of char; // для поиска нужного окна
begin
  sleeptime := strtoint(edtSleepTime.Text);
  AssignFile(log, Logpath);
  Append(log);
  if open.Execute = false then
    goto f1;
  //*********************************
  ShellExecute(handle, nil, pchar(edtUtmPath.Text), nil, nil, SW_SHOWNORMAL);
  err := 0;
  while not (apchar = 'UTM5 Подключение') or (apchar = 'UTM5 Login dialog') do
  begin
    err := err + 1;
    header := GetForegroundWindow; // получаем заголовок текущего активного окна
    GetWindowText(header, apchar, Length(apchar));
    if err > 10000000 then
    begin
      writeln(log, DateTimeToStr(Now) + ' Ошибка запуска!!!');
      Application.Destroy;
    end;
  end;
  if (apchar = 'UTM5 Подключение') then //если уже подключались
  begin
    sleep(sleeptime);
    mlclick(605, 588); // ок на первой форме
  end;
  if (apchar = 'UTM5 Login dialog') then // если первый раз
  begin
    sleep(sleeptime);
    mlclick(509, 471); // вводим хост
    SendKeys(cbbUTMhost.Items[cbbUTMhost.ItemIndex]);
    sleep(sleeptime);
    mlclick(562, 504); // вводим юзера
    SendKeys(edtUTMUser.Text);
    sleep(sleeptime);
    mlclick(564, 535); // вводим пароль
    SendKeys(edtUTMPass.Text);
    sleep(sleeptime);
    mlclick(645, 583); //
    sleep(sleeptime); // выбираем язык
    mlclick(634, 620); //
    sleep(sleeptime);
    mlclick(523, 603); //
    sleep(sleeptime); // ставим флажок сохранить
    mlclick(521, 622); //
    sleep(sleeptime);
    mlclick(530, 564); // закрываем настройки
    sleep(sleeptime);
    mlclick(602, 590); // ок
  end;
  err := 0;
  while not (apchar = edtUTMUser.text + '@' +
    cbbUTMhost.Items[cbbUTMhost.ItemIndex] + ' (Администратор)') do
  begin
    err := err + 1;
    header := GetForegroundWindow; // получаем заголовок текущего активного окна
    GetWindowText(header, apchar, Length(apchar));
    if err > 100000000 then
    begin
      writeln(log, DateTimeToStr(Now) + ' Ошибка подключения!!!');
      Close();
    end;
  end;
  sleep(sleeptime); //
  mlclick(935, 228); // найти на главной
  while not (apchar = 'Поиск') do
  begin
    header := GetForegroundWindow; // получаем заголовок текущего активного окна
    GetWindowText(header, apchar, Length(apchar));
  end;
  sleep(sleeptime);
  //*********************************
  AssignFile(f, open.FileName);
  Reset(f);
  readln(f, s);
  //  if getlogin(s) <> 'ID пользователя' then goto f1;
    //*********************************
  while not EOF(f) do
  begin
    readln(f, s);
    mlclick(357, 289); // выделяем поле ввода
    sleep(sleeptime);
    SendKeys(getlogin(s)); //кого ищем
    sleep(sleeptime);
    mlclick(959, 303); // найти
    sleep(sleeptime);
    mlclick(299, 476); // выделяем
    sleep(sleeptime);
    mlclick(838, 423); // внести платеж
    err := 0;
    while not (apchar = 'Внести платеж') do
    begin
      err := err + 1;
      header := GetForegroundWindow;
      // получаем заголовок текущего активного окна
      GetWindowText(header, apchar, Length(apchar));
      if err > 1000000 then
      begin
        writeln(log, DateTimeToStr(Now) + ' Ошибка ввода данных ! Файл:' +
          open.FileName + ' Строка ' + s);
        goto f2;
      end;
    end;
    mlclick(557, 404); // выделяем поле ввода суммы
    sleep(sleeptime);
    SendKeys(edtCreditSumm.Text); //сумма
    mlclick(572, 433);
    sleep(sleeptime);
    SendKeys('30.12.2010 23:59:00'); //обнуляем
    mlclick(500, 400);
    mlclick(572, 433);
    sleep(sleeptime);
    SendKeys(DateTimeToStr(Date()) + ' 00:00:00'); //дата
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
    while not (apchar = 'Поиск') do
    begin
      header := GetForegroundWindow;
      // получаем заголовок текущего активного окна
      GetWindowText(header, apchar, Length(apchar));
    end;
    sleep(sleeptime);
    mlclick(961, 337);
    f2: sleep(sleeptime);
  end;
  f1: mlclick(1023, 228); // закрываем netup
  sleep(sleeptime);
  mlclick(1166, 162);
  sleep(sleeptime);
  Closefile(f); //закрываем
  Flush(log);
  Closefile(log); //закрываем лог
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
  StatusBar1.Panels[0].Text := 'Отключен';
end;

procedure TfrmUtmAuto.CheckClick(Sender: TObject); //проверка
var
  Flist: Tstringlist; // список файлов с платежками
  f, dump: TextFile;
  i, err, j, totpay, leftpay, sleeptime: integer;
  s, tmpstr, vipiska, ds: string;
  trn: TTran; // сумма юзер дата номер
  summ, balance, leftSumm: Extended; // сумма всех внесенных
begin
  Flist := TStringList.Create;
  FindFile(edtPaymentPath.Text, Flist);
  // поиск файлов с начальными условиями, заданных в Edit1
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
    Closefile(f); //закрываем введенный
  end;
  redtReport.Lines.Add('Отчет');
  redtReport.Lines.Add('Файлов введено : ' + inttostr(Flist.Count));
  redtReport.Lines.Add('Платежей введено : ' + inttostr(totpay));
  redtReport.Lines.Add('Сумма : ' + floattostr(summ));
  redtReport.Lines.Add('Левых платежей : ' + inttostr(leftpay));
  redtReport.Lines.Add('Сумма левых платежей : ' + floattostr(leftSumm));
  Flist.Destroy;
end;

procedure TfrmUtmAuto.BalanceClick(Sender: TObject); //Баланс
var
  ds: string;
  balance: Extended; // сумма всех внесенных
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
        'Не удалось подключиться к SQL серверу. Не удалось вычислить баланс!!!');
      redtReport.Lines.Add('Не удалось вычислить баланс!!!');
    end;
  end;
  redtReport.Lines.Add('Баланс : ' + floattostr(balance));

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
// вводим платежи по старому
label
  e2; // без них никак ибо вложенные циклы
var
  fReestr, fReport, dump: TextFile; //
  s, tmpstr, vipiska, ds: string;
  trn: TTran; // сумма юзер дата номер
  Flist: Tstringlist; // список файлов с платежками
  i, err, j, totpay, leftpay, sleeptime: integer;
  // (i,j)счетчик  err - для выхода из бесконечного цикла
  header: HWND; // для поиска нужного окна
  apchar: array[0..254] of char; // для поиска нужного окна
  summ, balance: Extended; // сумма всех внесенных, баланс
begin
  ShellExecute(handle, nil, pchar(cbbUTMHost.Text), nil, nil, SW_SHOWNORMAL);
  err := 0;
  while not (apchar = 'UTM5 Подключение') or (apchar = 'UTM5 Login dialog') do
  begin
    err := err + 1;
    header := GetForegroundWindow;
    // получаем заголовок текущего активного окна
    GetWindowText(header, apchar, Length(apchar));
    if err > 100000000 then
    begin
      MessageBox(0, ' Ошибка запуска!!!', 'Внимание', MB_OK);
      Exit;
    end;
  end;
  if (apchar = 'UTM5 Подключение') then //если уже подключались
  begin
    sleep(sleeptime);
    mlclick(605, 588); // ок на первой форме
  end;
  if (apchar = 'UTM5 Login dialog') then // если первый раз
  begin
    sleep(sleeptime);
    mlclick(509, 471); // вводим хост
    SendKeys(cbbUTMhost.Text);
    sleep(sleeptime);
    mlclick(562, 504); // вводим юзера
    SendKeys(edtUTMUser.Text);
    sleep(sleeptime);
    mlclick(564, 535); // вводим пароль
    SendKeys(edtUTMPass.Text);
    sleep(sleeptime);
    mlclick(645, 583); //
    sleep(sleeptime); // выбираем язык
    mlclick(634, 620); //
    sleep(sleeptime);
    mlclick(523, 603); //
    sleep(sleeptime); // ставим флажок сохранить
    mlclick(521, 622); //
    sleep(sleeptime);
    mlclick(530, 564); // закрываем настройки
    sleep(sleeptime);
    mlclick(602, 590); // ок
  end;
  err := 0;
  while not (apchar = edtUTMUser.text + '@' + cbbUTMhost.text +
    ' (Администратор)') do
  begin
    err := err + 1;
    header := GetForegroundWindow;
    // получаем заголовок текущего активного окна
    GetWindowText(header, apchar, Length(apchar));
    if err > 10000000 then
    begin
      MessageBox(0, 'Ошибка подключения!!!', 'Внимание', MB_OK);
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
  mlclick(935, 228); // найти на главной
  while not (apchar = 'Поиск') do
  begin
    header := GetForegroundWindow;
    // получаем заголовок текущего активного окна
    GetWindowText(header, apchar, Length(apchar));
  end;
  sleep(sleeptime);
  //****************************************************************************
  AssignFile(fReport, edtReportPath.Text + 'платежи\платежи на ' +
    StringReplace(DateTimeToStr(now), ':', '-', [rfReplaceAll]) +
    '.txt'); //левые - идентификатор не 2 пример: 1100117
  CheckAndCreatePath(edtReportPath.Text + 'платежи\платежи на ' +
    StringReplace(DateTimeToStr(now), ':', '-', [rfReplaceAll]) +
    '.txt');
  Rewrite(fReport);
  summ := 0;
  totpay := 0;
  leftpay := 0;
  //****************************************************************************
  Flist := TStringList.Create;
  FindFile(edtPaymentPath.Text, Flist);
  // поиск файлов с начальными условиями, заданных в Edit1
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
            mlclick(715, 349); // выделяем поле ввода
            sleep(sleeptime);
            SendKeys(inttostr(strtoint(copy(trn.id, 3, 5))));
            //выделяем лицевой счет
          end;
        2:
          begin
            mlclick(360, 289); // выделяем поле ввода
            sleep(sleeptime);
            SendKeys('tlm041' + copy(trn.id, 5, 3)); //кого ищем
          end;
        1:
          begin
            mlclick(360, 289); // выделяем поле ввода
            sleep(sleeptime);
            SendKeys('chd006' + copy(trn.id, 5, 3)); //кого ищем
          end;
      else
        begin
          writeln(left, s);
          leftpay := leftpay + 1;
          continue;
        end;
      end;
      sleep(sleeptime);
      mlclick(959, 303); // найти
      sleep(sleeptime);
      mlclick(299, 476); // выделяем
      sleep(sleeptime);
      mlclick(838, 423); // внести платеж
      err := 0;
      while not (apchar = 'Внести платеж') do
      begin
        err := err + 1;
        header := GetForegroundWindow;
        // получаем заголовок текущего активного окна
        GetWindowText(header, apchar, Length(apchar));
        if err > 1000000 then
        begin
          writeln(elog, DateTimeToStr(Now) + ' Ошибка ввода данных ! Файл:' +
            Flist[i] + ' Строка №' + inttostr(i + 1) + ' Абонент ' + trn.id +
            ' сумма ' + trn.summ + ' от даты: ' + trn.date + ' за номером №' +
            trn.num);
          goto e2;
        end;
      end;
      sleep(sleeptime);
      mlclick(557, 404); // выделяем поле ввода суммы
      sleep(sleeptime);
      SendKeys(trn.summ);
      summ := summ + strtofloat(StringReplace(trn.summ, '.', ',',
        [rfReplaceAll]));
      mlclick(572, 433); // выделяем поле ввода даты
      sleep(sleeptime);
      SendKeys('11012001000000'); //обнуляем
      mlclick(500, 400);
      mlclick(572, 433);
      sleep(sleeptime);
      SendKeys(trn.date); //дата
      mlclick(565, 514); //для администратора
      sleep(sleeptime);
      SendKeys(Flist[i] + ' Date : ' + trn.date); //для администратора
      mlclick(619, 571); //
      sleep(sleeptime); // выбираем
      mlclick(681, 599); // способ
      sleep(sleeptime); // оплаты
      mlclick(563, 596); //номер платежки
      sleep(sleeptime);
      SendKeys(trn.num);
      mlclick(748, 715);
      totpay := totpay + 1;
      while not (apchar = 'Поиск') do
      begin
        header := GetForegroundWindow;
        // получаем заголовок текущего активного окна
        GetWindowText(header, apchar, Length(apchar));
      end;
      sleep(sleeptime);
      e2: mlclick(961, 337);
      sleep(sleeptime);
    end;
    Flush(fReestr);
    Closefile(fReestr); //закрываем введенный
    if chkIsDelete.Checked then
      DeleteFile(Flist[i]); // и удаляем чтоб не мешался
  end;
  mlclick(1023, 228);
  sleep(sleeptime);
  mlclick(1166, 162);
  sleep(sleeptime);
  Flush(fReport);
  Closefile(fReport); // закрываем левые платежи
  redtReport.Lines.Add('');
  redtReport.Lines.Add('Платежи внесены старым способом:');
  redtReport.Lines.Add('Сумма : ' + floattostr(summ));
  redtReport.Lines.Add('Файлов введено : ' + inttostr(Flist.Count));
  redtReport.Lines.Add('Платежей введено : ' + inttostr(totpay));
  redtReport.Lines.Add('Ошибок : ' + inttostr(err));
  redtReport.Lines.Add('Баланс : ' + floattostr(balance));
  redtReport.Lines.Add('выписка от ' + vipiska);
  Flist.Destroy;
  AssignFile(fReport, edtReportPath.Text + 'отчеты\отчет' + curentdate +
    '.txt');
  // статистика - сколько платежей сумма дата
  Rewrite(fReport);
  for i := 1 to redtReport.Lines.Count do
    writeln(fReport, redtReport.Lines[i]);
  Flush(fReport);
  Closefile(fReport); // закрываем
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
  edtTLNHost.Text := edtUTMHost.Text; // это для удобства
  edtSQLHost.Text := edtUTMHost.Text;
end;

procedure TfrmUtmAuto.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Flush(log);
  Closefile(log); //закрываем лог
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

