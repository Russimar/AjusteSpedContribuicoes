object ViewPrincipal: TViewPrincipal
  Left = 0
  Top = 0
  Caption = 'Principal'
  ClientHeight = 489
  ClientWidth = 1068
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  WindowState = wsMaximized
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object pnlPrincipal: TPanel
    Left = 0
    Top = 57
    Width = 1068
    Height = 432
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 0
    object pnlGrid: TPanel
      Left = 0
      Top = 0
      Width = 1068
      Height = 391
      Align = alClient
      BevelOuter = bvNone
      TabOrder = 0
      object gridDados: TSMDBGrid
        Left = 0
        Top = 0
        Width = 1068
        Height = 391
        Align = alClient
        DataSource = dsDados
        Options = [dgEditing, dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgConfirmDelete, dgCancelOnExit, dgTitleClick, dgTitleHotTrack]
        TabOrder = 0
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        Flat = False
        BandsFont.Charset = DEFAULT_CHARSET
        BandsFont.Color = clWindowText
        BandsFont.Height = -11
        BandsFont.Name = 'Tahoma'
        BandsFont.Style = []
        Groupings = <>
        GridStyle.Style = gsSoftGray
        GridStyle.OddColor = 15000804
        GridStyle.EvenColor = 16119285
        TitleHeight.PixelCount = 24
        FooterColor = clBtnFace
        ExOptions = [eoDisableInsert, eoENTERlikeTAB, eoKeepSelection, eoShowFooter, eoStandardPopup, eoBLOBEditor, eoTitleWordWrap, eoShowFilterBar, eoFilterAutoApply]
        RegistryKey = 'Software\Scalabium'
        RegistrySection = 'SMDBGrid'
        WidthOfIndicator = 11
        DefaultRowHeight = 17
        ScrollBars = ssHorizontal
        Columns = <
          item
            Expanded = False
            FieldName = 'NUMERO'
            Title.Alignment = taCenter
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'SERIE'
            Title.Alignment = taCenter
            Width = 35
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'CHAVE'
            Title.Alignment = taCenter
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'VLR_ICMS'
            Title.Alignment = taCenter
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'VALOR_BASE_ANTIGO'
            Title.Alignment = taCenter
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'VALOR_PIS_ANTIGO'
            Title.Alignment = taCenter
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'VALOR_COFINS_ANTIGO'
            Title.Alignment = taCenter
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'VLR_BASE'
            Title.Alignment = taCenter
            Title.Color = clYellow
            Title.Font.Charset = DEFAULT_CHARSET
            Title.Font.Color = clWindowText
            Title.Font.Height = -11
            Title.Font.Name = 'Tahoma'
            Title.Font.Style = [fsBold]
            Width = 111
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'VLR_PIS'
            Title.Alignment = taCenter
            Title.Color = clYellow
            Title.Font.Charset = DEFAULT_CHARSET
            Title.Font.Color = clBlack
            Title.Font.Height = 15
            Title.Font.Name = 'Tahoma'
            Title.Font.Style = [fsBold, fsItalic]
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'VLR_COFINS'
            Title.Alignment = taCenter
            Title.Color = clYellow
            Title.Font.Charset = DEFAULT_CHARSET
            Title.Font.Color = clBlack
            Title.Font.Height = 15
            Title.Font.Name = 'Tahoma'
            Title.Font.Style = [fsBold, fsItalic]
            Width = 82
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'VLR_TOTAL'
            Title.Alignment = taCenter
            Title.Color = clYellow
            Title.Font.Charset = DEFAULT_CHARSET
            Title.Font.Color = clBlack
            Title.Font.Height = 15
            Title.Font.Name = 'Tahoma'
            Title.Font.Style = [fsBold, fsItalic]
            Width = 101
            Visible = True
          end>
      end
    end
    object Panel1: TPanel
      Left = 0
      Top = 391
      Width = 1068
      Height = 41
      Align = alBottom
      BevelOuter = bvNone
      Caption = 'Gerar'
      Color = clHighlight
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWhite
      Font.Height = -11
      Font.Name = 'Segoe UI'
      Font.Style = [fsBold]
      ParentBackground = False
      ParentFont = False
      TabOrder = 1
      object btnGerar: TSpeedButton
        Left = 0
        Top = 0
        Width = 1068
        Height = 41
        Align = alClient
        Flat = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Segoe UI'
        Font.Style = [fsBold]
        ParentFont = False
        OnClick = btnGerarClick
        ExplicitLeft = 304
        ExplicitTop = 8
        ExplicitWidth = 23
        ExplicitHeight = 22
      end
    end
  end
  object pnl_Botton: TPanel
    Left = 0
    Top = 0
    Width = 1068
    Height = 57
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 1
    DesignSize = (
      1068
      57)
    object lblFilename: TLabel
      Left = 16
      Top = 9
      Width = 43
      Height = 13
      Caption = 'FileName'
    end
    object pnlGerarExcel: TPanel
      Left = 926
      Top = 12
      Width = 130
      Height = 34
      Anchors = [akTop, akRight]
      Color = clHighlight
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWhite
      Font.Height = -11
      Font.Name = 'Segoe UI'
      Font.Style = []
      ParentBackground = False
      ParentFont = False
      TabOrder = 0
      object btnGerarExcel: TSpeedButton
        Left = 1
        Top = 1
        Width = 128
        Height = 32
        Hint = 'Gerar Arquivo Excel'
        ParentCustomHint = False
        Align = alClient
        BiDiMode = bdLeftToRight
        Caption = 'Gerar Excel'
        ImageIndex = 0
        HotImageIndex = 0
        Flat = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clCream
        Font.Height = -11
        Font.Name = 'Segoe UI'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ParentBiDiMode = False
        ShowHint = True
        OnClick = btnGerarExcelClick
        ExplicitLeft = 41
        ExplicitTop = 0
      end
    end
    object chkC110: TCheckBox
      Left = 16
      Top = 34
      Width = 97
      Height = 17
      Caption = 'Possui C110'
      Checked = True
      State = cbChecked
      TabOrder = 1
    end
  end
  object OpenDialog: TOpenDialog
    Filter = 'Text|*.txt'
    Left = 600
    Top = 136
  end
  object dsDados: TDataSource
    DataSet = mtDados
    OnDataChange = dsDadosDataChange
    Left = 656
    Top = 89
  end
  object mtDados: TFDMemTable
    FieldDefs = <>
    IndexDefs = <>
    AggregatesActive = True
    FetchOptions.AssignedValues = [evMode]
    FetchOptions.Mode = fmAll
    ResourceOptions.AssignedValues = [rvSilentMode]
    ResourceOptions.SilentMode = True
    UpdateOptions.AssignedValues = [uvCheckRequired, uvAutoCommitUpdates]
    UpdateOptions.CheckRequired = False
    UpdateOptions.AutoCommitUpdates = True
    StoreDefs = True
    Left = 600
    Top = 89
    object mtDadosCHAVE: TStringField
      DisplayLabel = 'Chave Acesso'
      FieldName = 'CHAVE'
      Size = 44
    end
    object mtDadosVLR_ICMS: TFloatField
      DisplayLabel = 'Valor ICMS'
      FieldName = 'VLR_ICMS'
      DisplayFormat = '#,##0.00'
    end
    object mtDadosVLR_BASE: TFloatField
      DisplayLabel = 'Valor Base C'#225'lculo'
      FieldName = 'VLR_BASE'
      DisplayFormat = '#,##0.00'
    end
    object mtDadosVLR_PIS: TFloatField
      DisplayLabel = 'Valor Pis'
      FieldName = 'VLR_PIS'
      DisplayFormat = '#,##0.00'
    end
    object mtDadosVLR_COFINS: TFloatField
      DisplayLabel = 'Valor Cofins'
      FieldName = 'VLR_COFINS'
      DisplayFormat = '#,##0.00'
    end
    object mtDadosVLR_TOTAL: TFloatField
      DisplayLabel = 'Valor Total'
      FieldName = 'VLR_TOTAL'
      DisplayFormat = '#,##0.00'
    end
    object mtDadosNUMERO: TStringField
      DisplayLabel = 'N'#250'mero NF'
      FieldName = 'NUMERO'
      Size = 9
    end
    object mtDadosSERIE: TStringField
      DisplayLabel = 'S'#233'rie'
      FieldName = 'SERIE'
      Size = 3
    end
    object mtDadosVALOR_PIS_ANTIGO: TFloatField
      DisplayLabel = 'Pis Original'
      FieldName = 'VALOR_PIS_ANTIGO'
      DisplayFormat = '#,##0.00'
    end
    object mtDadosVALOR_COFINS_ANTIGO: TFloatField
      DisplayLabel = 'Cofins Original'
      FieldName = 'VALOR_COFINS_ANTIGO'
      DisplayFormat = '#,##0.00'
    end
    object mtDadosVALOR_BASE_ANTIGO: TFloatField
      DisplayLabel = 'Base Original'
      FieldName = 'VALOR_BASE_ANTIGO'
      DisplayFormat = '#,##0.00'
    end
    object mtDadosagTotal_Pis: TAggregateField
      FieldName = 'agTotal_Pis'
      Active = True
      DisplayName = ''
      DisplayFormat = '#,##0.00'
      Expression = 'SUM(VLR_PIS)'
    end
    object mtDadosagTotal_Cofins: TAggregateField
      FieldName = 'agTotal_Cofins'
      Active = True
      DisplayName = ''
      DisplayFormat = '#,##0.00'
      Expression = 'SUM(VLR_COFINS)'
    end
    object mtDadosagTotal_Geral: TAggregateField
      FieldName = 'agTotal_Geral'
      Active = True
      DisplayName = ''
      DisplayFormat = '#,##0.00'
      Expression = 'SUM(VLR_TOTAL)'
    end
    object mtDadosagTotal_Icms: TAggregateField
      FieldName = 'agTotal_Icms'
      Active = True
      DisplayName = ''
      DisplayFormat = '#,##0.00'
      Expression = 'SUM(VLR_ICMS)'
    end
  end
end
