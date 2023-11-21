unit GoogleAnalyticsGlobal;

interface

uses
  Google.Controller.Analytics.Interfaces;

var _GoogleAnalytics : iControllerGoogleAnalytics;

implementation

uses
  Google.Controller.Analytics,
  System.SysUtils,
  Vcl.Forms;

const
  GooglePropertyID = 'UA-250239007-1';
  AppName = 'CBL Informatica';
  AppLicense = 'Comercial';
  AppEdition = 'ERP';
  VersaoSistema = '1.0.0.1';

initialization
  _GoogleAnalytics := TControllerGoogleAnalytics
                        .New(GooglePropertyID);
  _GoogleAnalytics.AppInfo
    .AppName(AppName)
    .AppVersion(VersaoSistema)
    .AppLicense(AppLicense)
    .AppEdition(ExtractFileName(Application.ExeName));

end.
