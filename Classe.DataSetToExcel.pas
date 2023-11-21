unit Classe.DataSetToExcel;

interface

uses
  DB,
  DBClient,
  Dialogs,
  SMDBGrid;

type
  TDatasetToExcel = class
  private
    FCaminhoArquivo: String;
    FExcel: Variant;
    FLinha: Integer;
    FTexto: String;
    FValorCampo: String;
    FColuna: Integer;
    FPlanilha: Variant;
    FNomePlanilha: String;
    procedure SetCaminhoArquivo(const Value: String);
    procedure SetExcel(const Value: Variant);
    procedure SetLinha(const Value: Integer);
    procedure SetTexto(const Value: String);
    procedure SetValorCampo(const Value: String);
    procedure SetColuna(const Value: Integer);
    procedure SetPlanilha(const Value: Variant);
    procedure SetNomePlanilha(const Value: String);
    function ExcelInstalado : Boolean;
  public
    constructor Create;
    destructor Destroy; override;
    property Texto : String read FTexto write SetTexto;
    property ValorCampo : String read FValorCampo write SetValorCampo;
    property Excel : Variant read FExcel write SetExcel;
    property Planilha : Variant read FPlanilha write SetPlanilha;
    property Linha : Integer read FLinha write SetLinha;
    property Coluna : Integer read FColuna write SetColuna;
    property CaminhoArquivo : String read FCaminhoArquivo write SetCaminhoArquivo;
    property NomePlanilha : String read FNomePlanilha write SetNomePlanilha;
    function ExpXLS(DataSet : TDataSet; Grid : TSMDBGrid) : Boolean;
    function FormataTexto(aValue : String) : String;
    function GravarArquivo : String;
  end;

implementation

uses
  ComObj,
  SysUtils,
  Variants,
  ActiveX;

{ TDatasetToExcel }

constructor TDatasetToExcel.Create;
begin
  if ExcelInstalado then
    Excel := CreateOleObject('Excel.Application')
  else
    raise Exception.Create('Excel Não Instalado');  
  Excel.Workbooks.Add;
  Excel.caption := 'Exportando dados do tela para o Excel';
  Excel.visible := true;
  Linha := 1;
end;

destructor TDatasetToExcel.Destroy;
begin
  inherited;
end;

function TDatasetToExcel.ExpXLS(DataSet : TDataSet; Grid : TSMDBGrid) : Boolean;
var
  i : integer;
begin
  Result := False;
  Coluna := 0;
  try
    if VarType(Planilha) = VarEmpty then
    begin
      Planilha := Excel.Workbooks[1].Sheets.Add;
      Planilha.Name := FNomePlanilha;
    end;
    if Planilha.Name <> FNomePlanilha then
    begin
      Planilha := Excel.Workbooks[1].Sheets.Add;
      Planilha.Name := FNomePlanilha;
      Linha := 1;
    end;
    for i := 0 to Grid.FieldCount - 1 do
    begin
      if Grid.Columns[i].Visible then
      begin
        Inc(FColuna);
        Planilha.cells[FLinha, FColuna] := Grid.Columns[i].Title.Caption;
        Planilha.cells[FLinha, FColuna].Font.bold := True;
      end;
    end;
    DataSet.First;
    while not DataSet.Eof do
    begin
      Inc(FLinha);
      Coluna := 0;
      for i := 0 to Grid.FieldCount - 1 do
      begin
        if Grid.Columns[i].Visible then
        begin
          Inc(FColuna);
          Texto := DataSet.FieldByName(grid.Columns[i].FieldName).AsString;
          if trim(FTexto) <> '' then
            FValorcampo := DataSet.FieldByName(grid.Columns[i].FieldName).Value
          else
            FValorcampo := '';
          Texto := Grid.Columns[i].FieldName;

          if FieldTypeNames[DataSet.FieldByName(grid.Columns[i].FieldName).DataType] = 'Integer' then
          begin
            if Trim(FValorcampo) = '' then
              FValorcampo := '0';
            Planilha.Cells[FLinha, FColuna].NumberFormat := '#.##0_);(#.##0)';
            Planilha.cells[FLinha, FColuna] := StrToFloat(FValorcampo);
          end
          else
          if FieldTypeNames[DataSet.FieldByName(grid.Columns[i].FieldName).DataType] = 'Float' then
          begin
            if trim(FValorcampo) = '' then
              FValorcampo := '0';
            Planilha.Cells[FLinha, FColuna].NumberFormat := '#.##0,00_);(#.##0,000##)';
            Planilha.cells[FLinha, FColuna] := StrToFloat(FValorcampo);
          end
          else
          if FieldTypeNames[DataSet.FieldByName(grid.Columns[i].FieldName).DataType] = 'Date' then
          begin
            if trim(FValorcampo) <> '' then
            begin
              try
                FValorcampo := FormatDateTime('mm/dd/yyyy',StrToDate(FValorcampo));
                FExcel.Cells[FLinha, FColuna].NumberFormat := AnsiString('dd/mm/aaaa');
                FExcel.cells[FLinha, FColuna] := FValorcampo;
              except
                DataSet.Next;
              end;
            end;
          end
          else
          begin
            FExcel.Cells[FLinha, FColuna].NumberFormat := AnsiChar('@');;
            FExcel.cells[FLinha, FColuna] := FValorcampo;
          end;
        end;
      end;
      DataSet.Next;
    end;
    Result := True;
  except
    on E : Exception do
    begin
      MessageDlg('Erro ao gerar o arquivo ' + e.Message ,mtInformation,[mbOK],0)
    end;
  end;
end;

function TDatasetToExcel.FormataTexto(aValue: String): String;
begin
  aValue := StringReplace(aValue, '/','', [rfReplaceAll]);
  aValue := StringReplace(aValue, '*','', [rfReplaceAll]);
  aValue := StringReplace(aValue, '-','', [rfReplaceAll]);
  aValue := StringReplace(aValue, '+','', [rfReplaceAll]);
  Result := aValue;
end;

function TDatasetToExcel.GravarArquivo: String;
begin
  Excel.WorkBooks[1].SaveAs(FCaminhoArquivo);
end;

procedure TDatasetToExcel.SetCaminhoArquivo(const Value: String);
begin
  if Value = EmptyStr then
    raise Exception.Create('Informe o caminho do Arquivo');
  FCaminhoArquivo := Value;
end;

procedure TDatasetToExcel.SetColuna(const Value: Integer);
begin
  FColuna := Value;
end;

procedure TDatasetToExcel.SetLinha(const Value: Integer);
begin
  FLinha := Value;
end;

procedure TDatasetToExcel.SetExcel(const Value: Variant);
begin
  FExcel := Value;
end;

procedure TDatasetToExcel.SetTexto(const Value: String);
begin
  FTexto := Value;
end;

procedure TDatasetToExcel.SetValorCampo(const Value: String);
begin
  FValorCampo := Value;
end;

procedure TDatasetToExcel.SetPlanilha(const Value: Variant);
begin
  FPlanilha := Value;
end;

procedure TDatasetToExcel.SetNomePlanilha(const Value: String);
begin
  FNomePlanilha := Value;
end;

function TDatasetToExcel.ExcelInstalado: Boolean;
var
  ClassID : TCLSID;
  strOleObject :  string;
begin
  strOleObject := 'Excel.Application';
  Result  := CLSIDFromProgID(PWideChar(WideString(strOleObject)),ClassID) = S_OK;
end;

end.
