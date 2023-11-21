program prjAjustaSpedContribuicoes;

uses
  Vcl.Forms,
  uConverteSped in 'uConverteSped.pas' {ViewPrincipal},
  Classe.DataSetToExcel in 'Classe.DataSetToExcel.pas',
  Loading in 'Loading.pas' {ViewLoaging},
  GoogleAnalyticsGlobal in 'GoogleAnalyticsGlobal.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TViewPrincipal, ViewPrincipal);
  Application.Run;
end.
