unit UDMIntegrador;

interface

uses
  SysUtils, Classes, DB, DBClient, frxClass, frxDBSet, frxRich,
  frxExportMail, frxExportPDF;

type
  TDMIntegrador = class(TDataModule)
    mAvisos: TClientDataSet;
    dsmAviso: TDataSource;
    mAvisosID: TIntegerField;
    mAvisosNome: TStringField;
    mAvisosTipo: TStringField;
    mAvisosObs: TStringField;
    mAvisosTipo_Reg: TStringField;
    mAvisosReferencia: TStringField;
    frxReport1: TfrxReport;
    frxPDFExport1: TfrxPDFExport;
    frxMailExport1: TfrxMailExport;
    frxRichObject1: TfrxRichObject;
    frxmAvisos: TfrxDBDataset;
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
