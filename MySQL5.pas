unit MySQL5;

interface

uses
  SysUtils, Classes, mysql;

type
  TDynDataSet = array of array of string;
  TMySQL5 = class(TComponent)
  private
    { Private declarations }
  	MySQL: PMYSQL;
	  MyHost: string;
	  MyPort: integer;
	  MyUser: string;
	  MyPass: string;
	  MyTime: longword;
	  MyComp: integer;
    RecordNum: Integer;
    CurrRecord: Integer;
    FieldCount: Integer;
    //Dataset: array[1..1500, 1..10] of string;
      Dataset: TDynDataSet;
  protected
    { Protected declarations }
    procedure Fail(Msg: string = '');
  public
    { Public declarations }
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure Connect(Host, User, Pass: String; Port: Integer=3306);
    procedure Close;
    procedure Execute(Query: String);
    procedure First;
    procedure Next;
    procedure Prev;
    procedure Last;
    procedure Move(Pos: Integer);
    function NotEmpty :Boolean;
    function AsString (FieldId: Integer): String;
    function AsInteger(FieldId: Integer): Integer;
  published
    { Published declarations }
    property RecordCount: Integer read RecordNum;
    property FCount: Integer read FieldCount;
  end;

procedure Register;

implementation

procedure TMySQL5.Fail(Msg: string = '');
begin
	if Msg = '' then Msg := Format('%d - %s', [mysql_errno(MySQL), mysql_error(MySQL)]);
	raise Exception.Create(Msg);
end;

procedure Register;
begin
  RegisterComponents('MysqlVCL5', [TMySQL5]);
end;

constructor TMySQL5.Create(AOwner: TComponent);
begin
	MyHost := 'localhost';
	MyPort := MYSQL_PORT;
	MyUser := '';
	MyPass := '';
	MyTime := 30;
	MyComp := 0 * CLIENT_COMPRESS;

  RecordNum := 0;
  CurrRecord:= 0;

  inherited;
end;

destructor TMySQL5.Destroy;
begin
  Close;
  inherited;
end;

procedure TMySQL5.Connect(Host, User, Pass: String; Port: Integer=3306);
begin
	MySQL := mysql_init(nil);
	if MySQL = nil then Fail('Couldn''t init PMYSQL object');
	try
		if mysql_options(MySQL, MYSQL_SET_CHARSET_NAME, 'utf8') <> 0 then Fail;
    if mysql_options(MySQL, MYSQL_OPT_CONNECT_TIMEOUT, '30') <> 0 then Fail;
		if mysql_real_connect(MySQL, pChar(Host), pChar(User), pChar(Pass),
                          nil, Port, nil, MyComp) = nil then Fail;
	except
		Close;
	end;
end;

procedure TMySQL5.Close;
begin
	mysql_close(MySQL);
  RecordNum := 0;
  CurrRecord:= 0;
end;

procedure TMySQL5.Execute(Query: String);
var
	 Result: PMYSQL_RES;
	 Row: PMYSQL_ROW;
   K: Integer;
begin
 //	Query :=Format('SELECT host,user FROM %s LIMIT %d', ['mysql.user', 10]);
	if mysql_query(MySQL, pChar(Query)) <> 0 then Fail;
	Result := mysql_use_result(MySQL);
	if Result = nil then Fail;
  FieldCount := mysql_field_count(MySQL);
  if FieldCount > 10 then FieldCount := 10;
  Row := mysql_fetch_row(Result);
  while Row <> nil do begin
    Inc(RecordNum);
    SetLength(Dataset,RecordNum);
    SetLength(Dataset[RecordNum-1],FieldCount);
    for K := 1 to FieldCount do
      Dataset[RecordNum-1, K-1] := Row[K-1];
   // if RecordNum = 1500 then Break;
    Row := mysql_fetch_row(Result);
  end
end;

procedure TMySQL5.First;
begin
  if RecordNum > 0 then CurrRecord := 1;
end;

function TMySQL5.NotEmpty : Boolean;
begin
  if RecordNum > 0 then Result := True
  else  Result := False;
end;

procedure TMySQL5.Next;
begin
  if (RecordNum > 0) and (CurrRecord < RecordNum) then Inc(CurrRecord)
end;

procedure TMySQL5.Prev;
begin
  if (RecordNum > 0) and (CurrRecord > 1) then Dec(CurrRecord)
end;

procedure TMySQL5.Last;
begin
  if RecordNum > 0 then CurrRecord := RecordNum;
end;

procedure TMySQL5.Move(Pos: Integer);
begin
  if (Pos > 0) and (Pos <= RecordNum) then CurrRecord := Pos;
end;

function TMySQL5.AsString (FieldId: Integer): String;
begin
  Result := '';
  if (FieldId > 0) and (FieldId <= FieldCount) then
    Result := Dataset[CurrRecord-1, FieldId-1]
  else Fail('Error: Out of range Field Index');
end;

function TMySQL5.AsInteger(FieldId: Integer): Integer;
begin
  Result := 0;
  if (FieldId > 0) and (FieldId <= FieldCount) then
    Result := StrToInt(Dataset[CurrRecord-1, FieldId-1])
  else Fail('Error: Out of range Field Index');
end;

end.
