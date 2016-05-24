program AutoUtm2;

uses
  Forms,
  AutoUtm2Main in 'AutoUtm2Main.pas' {frmUtmAuto},
  AutoUtm2Procs in 'AutoUtm2Procs.pas',
  MainMiniUtm in '..\Miniutm\MainMiniUtm.pas' {frmMiniUtm},
  Unit1 in '..\utm-qwery\Unit1.pas' {Form1};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfrmUtmAuto, frmUtmAuto);
  Application.CreateForm(TfrmMiniUtm, frmMiniUtm);
  Application.CreateForm(TForm1, Form1);
  Application.Run;

end.
