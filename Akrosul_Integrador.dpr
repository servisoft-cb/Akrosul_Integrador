program Akrosul_Integrador;

uses
  Forms,
  UDMCadPessoa in '..\ssfacil\UDMCadPessoa.pas' {DMCadPessoa: TDataModule},
  LogProvider in '..\Logs\src\LogProvider.pas',
  LogTypes in '..\Logs\src\LogTypes.pas',
  DmdDatabase in '..\ssfacil\DmdDatabase.pas' {dmDatabase: TDataModule},
  uUtilPadrao in '..\ssfacil\uUtilPadrao.pas',
  UEscolhe_Filial in '..\ssfacil\UEscolhe_Filial.pas' {frmEscolhe_Filial},
  rsDBUtils in '..\rslib\nova\rsDBUtils.pas',
  DmdDatabase_NFeBD in '..\ssfacil\DmdDatabase_NFeBD.pas' {dmDatabase_NFeBD: TDataModule},
  UDMCadProduto in '..\ssfacil\UDMCadProduto.pas' {dmCadProduto: TDataModule},
  UIntegrador in 'UIntegrador.pas' {frmIntegrador},
  UDMIntegrador in 'UDMIntegrador.pas' {DMIntegrador: TDataModule},
  UDMCadNCM in '..\ssfacil\UDMCadNCM.pas' {DMCadNCM: TDataModule},
  UDMCadUnidade in '..\ssfacil\UDMCadUnidade.pas' {DMCadUnidade: TDataModule},
  UDMCadGrupo in '..\ssfacil\UDMCadGrupo.pas' {DMCadGrupo: TDataModule},
  UDMCadMarca in '..\ssfacil\UDMCadMarca.pas' {DMCadMarca: TDataModule},
  UDMEstoque in '..\ssfacil\UDMEstoque.pas' {DMEstoque: TDataModule};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Akrosul_Integrador';
  Application.CreateForm(TdmDatabase, dmDatabase);
  Application.CreateForm(TfrmIntegrador, frmIntegrador);
  Application.Run;
end.
