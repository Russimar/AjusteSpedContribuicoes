unit Loading;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.StdCtrls,
  Vcl.Imaging.GIFImg;

type
  TViewLoaging = class(TForm)
    Image: TImage;
    Label1: TLabel;
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  ViewLoaging: TViewLoaging;

implementation

{$R *.dfm}

procedure TViewLoaging.FormCreate(Sender: TObject);
begin
  TGIFImage(Image.Picture.Graphic).Animate := True;
end;

end.
