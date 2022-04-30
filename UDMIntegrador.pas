unit UDMIntegrador;

interface

uses
  SysUtils, Classes, DB, DBClient;

type
  TDMIntegrador = class(TDataModule)
    mAvisos: TClientDataSet;
    dsmAviso: TDataSource;
    mAvisosID: TIntegerField;
    mAvisosNome: TStringField;
    mAvisosTipo: TStringField;
    mAvisosObs: TStringField;
    mAvisosTipo_Reg: TStringField;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  DMIntegrador: TDMIntegrador;

implementation

{$R *.dfm}

end.
