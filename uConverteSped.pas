unit uConverteSped;

interface

uses
  Winapi.Windows,
  Winapi.Messages,
  System.Variants,
  System.Classes,
  Vcl.Graphics,
  StrUtils,
  Vcl.Controls,
  Vcl.Forms,
  Vcl.Dialogs,
  Vcl.StdCtrls,
  Vcl.Buttons,
  Vcl.ExtCtrls,
  Data.DB,
  Vcl.Grids,
  Vcl.DBGrids,
  FireDAC.Stan.Intf,
  FireDAC.Stan.Option,
  FireDAC.Stan.Param,
  FireDAC.Stan.Error,
  FireDAC.DatS,
  FireDAC.Phys.Intf,
  FireDAC.DApt.Intf,
  FireDAC.Comp.DataSet,
  FireDAC.Comp.Client,
  System.Types,
  ACBrUtil.Math,
  SMDBGrid,
  Classe.DataSetToExcel,
  Loading,
//  softMeter_globalVar,
  System.Threading;

type
  TViewPrincipal = class(TForm)
    pnlPrincipal: TPanel;
    pnlGrid: TPanel;
    pnl_Botton: TPanel;
    OpenDialog: TOpenDialog;
    lblFilename: TLabel;
    dsDados: TDataSource;
    mtDados: TFDMemTable;
    mtDadosCHAVE: TStringField;
    mtDadosVLR_ICMS: TFloatField;
    mtDadosVLR_BASE: TFloatField;
    mtDadosVLR_PIS: TFloatField;
    mtDadosVLR_COFINS: TFloatField;
    gridDados: TSMDBGrid;
    mtDadosagTotal_Pis: TAggregateField;
    mtDadosagTotal_Cofins: TAggregateField;
    mtDadosVLR_TOTAL: TFloatField;
    mtDadosagTotal_Geral: TAggregateField;
    pnlGerarExcel: TPanel;
    btnGerarExcel: TSpeedButton;
    mtDadosagTotal_Icms: TAggregateField;
    Panel1: TPanel;
    btnGerar: TSpeedButton;
    chkC110: TCheckBox;
    mtDadosNUMERO: TStringField;
    mtDadosSERIE: TStringField;
    mtDadosVALOR_PIS_ANTIGO: TFloatField;
    mtDadosVALOR_COFINS_ANTIGO: TFloatField;
    mtDadosVALOR_BASE_ANTIGO: TFloatField;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnGerarExcelClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnGerarClick(Sender: TObject);
    procedure dsDadosDataChange(Sender: TObject; Field: TField);
  private
    FLista: TStringList;
    FLista1: TStringList;
    FNewList: TStringList;
    FValorICMS: Real;
    FValorPis: Real;
    FValorCofins: Real;
    FBasePisCofins: Real;
    FValorDesconto: Real;
    FAliqPis: Real;
    FAliqCofins: Real;
    FDatasetToExcel : TDatasetToExcel;
    FValorDifPis: Real;
    FValorDifCofins: Real;
    FValorPisC100: Real;
    FValorCofinsC100: Real;
    AllTasks: array of ITask;
    FLoading: TViewLoaging;
    FFLabel: PAnsiChar;
    FPValue: PAnsiChar;
    FBasePisCofinsAtual: Real;
    FValorOperacao: Real;
    FValorAtualPis: Real;
    FValorAtualCofins: Real;
    procedure GeraArquivoExcel;
    procedure AjustaArquivo;
    procedure CalculaTotais;
    procedure CalculaPisCofinsC175(AValue, AOld : TStringDynArray);
    procedure CalculaPisCofinsC170(AValue, AOld : TStringDynArray);
    function GetPathFile : String;
    function ValidaC175(AValue : TStringDynArray) : Boolean;
    function ValidaC170(AValue : TStringDynArray) : Boolean;
    function ValidaC100(AValue, AOld : TStringDynArray) : Boolean;
    function ValidaCSTPisCofC175(AValue: TStringDynArray): Boolean;
    function ValidaCSTPisCofC170(AValue: TStringDynArray): Boolean;
    function ValidaCFOP(AValue : TStringDynArray; Pos : Integer) : Boolean;
    property ValorICMS: Real read FValorICMS write FValorICMS;
    property ValorPis: Real read FValorPis write FValorPis;
    property ValorCofins: Real read FValorCofins write FValorCofins;
    property BasePisCofins: Real read FBasePisCofins write FBasePisCofins;
    property ValorDesconto: Real read FValorDesconto write FValorDesconto;
    property AliqPis: Real read FAliqPis write FAliqPis;
    property AliqCofins: Real read FAliqCofins write FAliqCofins;
    property ValorDifPis: Real read FValorDifPis write FValorDifPis;
    property ValorDifCofins: Real read FValorDifCofins write FValorDifCofins;
    property ValorPisC100: Real read FValorPisC100 write FValorPisC100;
    property ValorCofinsC100: Real read FValorCofinsC100 write FValorCofinsC100;
    property FLabel: PAnsiChar read FFLabel write FFLabel;
    property PValue: PAnsiChar read FPValue write FPValue;
    property BasePisCofinsAtual: Real read FBasePisCofinsAtual write FBasePisCofinsAtual;
    property ValorOperacao: Real read FValorOperacao write FValorOperacao;
    property ValorAtualPis: Real read FValorAtualPis write FValorAtualPis;
    property ValorAtualCofins: Real read FValorAtualCofins write FValorAtualCofins;
    procedure GravarDados(AOld: TArray<System.string>);
  public
    procedure RunTask(var aTask: ITask; aTp: Integer);
    procedure ExibirLoading;
  end;

var
  ViewPrincipal: TViewPrincipal;

implementation

uses
  System.SysUtils,
  GoogleAnalyticsGlobal;

{$R *.dfm}

procedure TViewPrincipal.AjustaArquivo;
var
  i, i2: Integer;
  Linha: String;
  Registro : TStringDynArray;
  RegistroMaisUm : TStringDynArray;
  Gerar : Boolean;
  Lista : TStringList;
begin
  mtDados.Close;
  mtDados.Open;
  Lista := TStringList.Create;
  Lista.LoadFromFile(OpenDialog.FileName);
  Gerar := True;
  for i := 0 to Lista.Count - 1 do
  begin
    Linha := Lista[i];

    if not Lista[i].IsEmpty then
      Registro := SplitString(Lista[i],'|');
    if (Registro[1] = '0000') then
      _GoogleAnalytics
        .Event('Documento','Geracao',Registro[9],1);

    if i < pred(Lista.Count) then
    begin
      if copy(Registro[1],1,3) <> '999' then
      begin
        if (chkC110.Checked) then
          RegistroMaisUm := SplitString(Lista[i+2],'|')
         else
          RegistroMaisUm := SplitString(Lista[i+1],'|');
      end;

      if (ValidaC100(Registro, RegistroMaisUm)) then
      begin
        if ValidaCSTPisCofC175(RegistroMaisUm) then
        begin
          ValorICMS := RoundABNT(StrToCurrDef(Registro[22],0),2);

          if ValorICMS > 0 then
          begin
            ValorPisC100 := RoundABNT(StrToCurrDef(Registro[26],0),2);
            ValorCofinsC100 := RoundABNT(StrToCurrDef(Registro[27],0),2);

            CalculaPisCofinsC175(RegistroMaisUm, Registro);
            Registro[26] := CurrToStr(FValorPisC100 - FValorDifPis);
            Registro[27] := CurrToStr(FValorCofinsC100 - FValorDifCofins);

          end;
        end;

        if ValidaCSTPisCofC170(RegistroMaisUm) then
        begin
          ValorICMS := RoundABNT(StrToCurrDef(Registro[22],0),2);
          if ValorICMS > 0 then
          begin
            ValorPisC100 := RoundABNT(StrToCurrDef(Registro[26],0),2);
            ValorCofinsC100 := RoundABNT(StrToCurrDef(Registro[27],0),2);

            CalculaPisCofinsC170(RegistroMaisUm, Registro);
            Registro[26] := CurrToStr(FValorPisC100 - FValorDifPis);
            Registro[27] := CurrToStr(FValorCofinsC100 - FValorDifCofins);
          end;
        end;
      end;

      if ValidaC175(Registro) and (ValorICMS > 0) then
      begin
        Linha := '';
        Registro[6] := FormatFloat('0.00', BasePisCofins);
        Registro[4] := FormatFloat('0.00', ValorDesconto);
        Registro[10] := FormatFloat('0.00', ValorPis);
        Registro[12] := FormatFloat('0.00', BasePisCofins);
        Registro[16] := FormatFloat('0.00', ValorCofins);
      end;

      if ValidaC170(Registro) and (ValorICMS > 0) then
      begin
        Linha := '';
        Registro[26] := FormatFloat('0.00', BasePisCofins);
        Registro[08] := FormatFloat('0.00', ValorDesconto);
        Registro[30] := FormatFloat('0.00', ValorPis);
        Registro[32] := FormatFloat('0.00', BasePisCofins);
        Registro[36] := FormatFloat('0.00', ValorCofins);
      end;

      Linha := '';
      for I2 := Low(Registro) to Pred(high(Registro)) do
      begin
       Linha := Linha + registro[i2] + '|';
      end;
    end;

    FLista.Add(Linha);
    FLista.SaveToFile('c:\temp\teste.txt');

  end;
  FLista.SaveToFile(GetPathFile);
end;

procedure TViewPrincipal.btnGerarExcelClick(Sender: TObject);
begin
  SetLength(AllTasks, 1);
  RunTask(AllTasks[0],2);
  ExibirLoading;
end;

procedure TViewPrincipal.CalculaPisCofinsC170(AValue, AOld: TStringDynArray);
begin
  ValorOperacao := RoundABNT(StrToFloat(AValue[5]) * StrToFloat(AValue[7]) ,2);
  ValorAtualPis := RoundABNT(StrToFloat(AValue[30]),2);
  ValorAtualCofins := RoundABNT(StrToFloat(AValue[36]),2);
  BasePisCofinsAtual := RoundABNT(StrToCurrDef(AValue[26],0),2);
  BasePisCofins := RoundABNT(StrToCurrDef(AValue[32],0),2);

  ValorDesconto := ValorICMS;

  if ValorOperacao < ValorICMS then
    ValorDesconto := ValorOperacao;

  BasePisCofins := BasePisCofins - ValorDesconto;

  AliqPis := StrToCurrDef(AValue[27],0);
  ValorPis := RoundABNT(BasePisCofins * (AliqPis / 100),2);
  ValorDifPis := ValorAtualPis - ValorPis;

  AliqCofins := StrToCurrDef(AValue[33],0);
  ValorCofins := RoundABNT(BasePisCofins * (AliqCofins / 100),2);
  ValorDifCofins := ValorAtualCofins - ValorCofins;
  GravarDados(AOld);
end;

procedure TViewPrincipal.GravarDados(AOld: TArray<System.string>);
begin
  mtDados.Insert;
  mtDadosCHAVE.AsString := AOld[9];
  mtDadosNUMERO.AsString := AOld[8];
  mtDadosSERIE.AsString := AOld[7];
  mtDadosVALOR_PIS_ANTIGO.AsFloat := ValorAtualPis;
  mtDadosVALOR_COFINS_ANTIGO.AsFloat := ValorAtualCofins;
  mtDadosVALOR_BASE_ANTIGO.AsFloat := BasePisCofinsAtual;
  mtDadosVLR_ICMS.AsFloat := FValorICMS;
  mtDadosVLR_BASE.AsFloat := FBasePisCofins;
  mtDadosVLR_PIS.AsFloat := FValorDifPis;
  mtDadosVLR_COFINS.AsFloat := FValorDifCofins;
  mtDadosVLR_TOTAL.AsFloat := FValorDifPis + FValorDifCofins;
  mtDados.Post;
end;

procedure TViewPrincipal.CalculaPisCofinsC175(AValue, AOld : TStringDynArray);
var
  ValorAtualPis, ValorAtualCofins : Real;
begin
  ValorOperacao := RoundABNT(StrToFloat(AValue[3]),2);
  ValorAtualPis := RoundABNT(StrToFloat(AValue[10]),2);
  ValorAtualCofins := RoundABNT(StrToFloat(AValue[16]),2);
  BasePisCofinsAtual := RoundABNT(StrToCurrDef(AValue[6],0),2);
  BasePisCofins := RoundABNT(StrToCurrDef(AValue[6],0),2);

  ValorDesconto := ValorICMS;

  if ValorOperacao < ValorICMS then
    ValorDesconto := ValorOperacao;

  BasePisCofins := BasePisCofins - ValorDesconto;

  AliqPis := StrToCurrDef(AValue[7],0);
  ValorPis := RoundABNT(BasePisCofins * (AliqPis / 100),2);
  ValorDifPis := ValorAtualPis - ValorPis;

  AliqCofins := StrToCurrDef(AValue[13],0);
  ValorCofins := RoundABNT(BasePisCofins * (AliqCofins / 100),2);
  ValorDifCofins := ValorAtualCofins - ValorCofins;

  GravarDados(AOld);
end;

procedure TViewPrincipal.CalculaTotais;
var i : Integer;
begin
  for i := 0 to gridDados.ColCount - 2 do
  begin
    if (gridDados.Columns[i].FieldName = 'VLR_PIS') then
        gridDados.Columns[i].FooterValue := FormatFloat('#,###,###,##0.00', StrToFloatDef(mtDadosagTotal_Pis.AsString,0));
    if (UpperCase(gridDados.Columns[i].FieldName) = 'VLR_COFINS') then
      gridDados.Columns[i].FooterValue := FormatFloat('#,###,###,##0.00', StrToCurrDef(mtDadosagTotal_Cofins.AsString,0));
    if (UpperCase(gridDados.Columns[i].FieldName) = 'VLR_TOTAL') then
      gridDados.Columns[i].FooterValue := FormatFloat('#,###,###,##0.00', StrToCurrDef(mtDadosagTotal_Geral.AsString,0));
    if (UpperCase(gridDados.Columns[i].FieldName) = 'VLR_ICMS') then
      gridDados.Columns[i].FooterValue := FormatFloat('#,###,###,##0.00', StrToCurrDef(mtDadosagTotal_Icms.AsString,0));
  end;
end;

procedure TViewPrincipal.dsDadosDataChange(Sender: TObject; Field: TField);
begin
  pnlGerarExcel.Enabled := not dsDados.DataSet.IsEmpty;
end;

procedure TViewPrincipal.ExibirLoading;
begin
  TTask.Run(
    procedure
    begin
      TThread.Synchronize(TThread.CurrentThread,
      procedure
      begin
        FLoading := TViewLoaging.Create(nil);
        FLoading.Show;
      end);
      TTask.WaitForAll(AllTasks);
      TThread.Queue(TThread.CurrentThread,
      procedure
      begin
        FLoading.Close;
        FLoading.Free;
      end);
    end);
end;

procedure TViewPrincipal.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  FLista.Free;
  Flista1.Free;
  FNewList.Free;
  _GoogleAnalytics.EndSession;
end;

procedure TViewPrincipal.FormCreate(Sender: TObject);
begin
//  dllSoftMeter.sendScreenView(PansiChar(Self.Caption));
  _GoogleAnalytics.StartSession;
end;

procedure TViewPrincipal.FormShow(Sender: TObject);
begin
  FLista := TStringList.Create;
  FNewList := TStringList.Create;
  FLista1 := TStringList.Create;
  _GoogleAnalytics
    .PageView(ExtractFileName(Application.ExeName),
              Self.Name,
              Self.Caption);
end;

procedure TViewPrincipal.GeraArquivoExcel;
var
  ANameFile, AExtFile : String;
begin
  ANameFile := ExtractFileName(OpenDialog.FileName);
  AExtFile := ExtractFileExt(OpenDialog.FileName);
  ANameFile := StringReplace(ANameFile, AExtFile, '', [rfReplaceAll]);
  mtDados.DisableControls;
  FDatasetToExcel := TDatasetToExcel.Create;
  try
    with FDatasetToExcel do
    begin
      CaminhoArquivo := ExtractFilePath(ParamStr(0)) + ANameFile + '_Novo.xls';
      NomePlanilha := 'Ajustes_Pis_Cofins';
      ExpXLS(mtDados, gridDados);
      GravarArquivo;
    end;
  finally
    mtDados.EnableControls;
    FDatasetToExcel.Free;
  end;
end;

function TViewPrincipal.GetPathFile: String;
var
  ANameFile, AExtFile : String;
begin
  ANameFile := ExtractFileName(OpenDialog.FileName);
  AExtFile := ExtractFileExt(OpenDialog.FileName);
  ANameFile := StringReplace(ANameFile, AExtFile, '', [rfReplaceAll]);
  Result := ExtractFilePath(ParamStr(0)) + ANameFile + '_Novo' + AExtFile;
end;

procedure TViewPrincipal.RunTask(var aTask: ITask; aTp: Integer);
begin
  aTask := TTask.Run(
    procedure
    begin
      mtDados.DisableControls;
      if aTp = 1 then
      begin
        AjustaArquivo;
        CalculaTotais;
      end;
      if aTp = 2 then
        GeraArquivoExcel;
      mtDados.EnableControls;
      TThread.Synchronize(nil,
      procedure
      begin
      end);
    end);
end;

function TViewPrincipal.ValidaCFOP(AValue : TStringDynArray; Pos : Integer) : Boolean;
begin
  Result := False;
  if AValue[Pos].IsEmpty then
    Exit;
  case AnsiIndexStr(AValue[Pos], ['5101', '5102','5401','6101','6102','6401']) of
   0,1,2,3,4,5 :
      begin
        Result := True;
      end;
  end;
end;

function TViewPrincipal.ValidaCSTPisCofC170(AValue: TStringDynArray): Boolean;
begin
  Result := False;
  case AnsiIndexStr(AValue[25], ['01']) of
   0 : Result := True;
  end;
end;

function TViewPrincipal.ValidaCSTPisCofC175(AValue: TStringDynArray): Boolean;
begin
  Result := False;
  case AnsiIndexStr(AValue[5], ['01']) of
   0 : Result := True;
  end;
end;

function TViewPrincipal.ValidaC100(AValue, AOld : TStringDynArray) : Boolean;
begin
  Result := False;
  if (AValue[1] = 'C100') then
  begin
    if AValue[6] = '02'  then //cancelado
      Exit;
    if (ValidaCFOP(AOld, 11) or ValidaCFOP(AOld, 2)) then
      Result := True;
  end;
end;

function TViewPrincipal.ValidaC170(AValue: TStringDynArray): Boolean;
begin
  Result := False;
  if (AValue[1] = 'C170') then
  begin
    if ValidaCFOP(AValue, 11) then
      Result := True;
  end;
end;

function TViewPrincipal.ValidaC175(AValue: TStringDynArray): Boolean;
begin
  Result := False;
  if (AValue[1] = 'C175') then
  begin
    if ValidaCFOP(AValue, 2) and (AValue[5] = '01') then
      Result := True;
  end;
end;

procedure TViewPrincipal.btnGerarClick(Sender: TObject);
begin
  if OpenDialog.Execute then
  begin
    lblFilename.Caption := OpenDialog.FileName;
    lblFilename.Update;
  end
  else
  Exit;

  SetLength(AllTasks, 1);
  RunTask(AllTasks[0],1);
  ExibirLoading;
end;

end.
