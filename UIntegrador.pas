unit UIntegrador;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RzTabs, Grids, DBGrids, SMDBGrid, NxCollection,
  ExtCtrls, StdCtrls, Mask, ToolEdit, ComObj, Buttons, DB, SqlExpr,
  CurrEdit, midaslib, UDMCadPessoa, UDMCadProduto, UDMIntegrador, AdvPanel, UDMCadNCM, UDMCadUnidade,
  UDMCadGrupo, UDMCadMarca, UDMEstoque, NxEdit;
  
type
  TfrmIntegrador = class(TForm)
    pnlPrincipal: TAdvPanel;
    File_Produto: TFilenameEdit;
    btnProduto: TNxButton;
    Label2: TLabel;
    ceLidos: TCurrencyEdit;
    ceTotal: TCurrencyEdit;
    Label1: TLabel;
    NxButton1: TNxButton;
    Label3: TLabel;
    RzPageControl1: TRzPageControl;
    TabSheet1: TRzTabSheet;
    SMDBGrid1: TSMDBGrid;
    Label4: TLabel;
    ceGravados: TCurrencyEdit;
    chkFornecedor: TNxCheckBox;
    btnExcel: TNxButton;
    chkGerarProdutos: TNxCheckBox;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnProdutoClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnExcelClick(Sender: TObject);
  private
    { Private declarations }
    fDMCadPessoa: TDMCadPessoa;
    fDMCadProduto: TDMCadProduto;
    fDMIntegrador: TDMIntegrador;
    fDMCadNCM: TDMCadNCM;
    fDMCadUnidade: TDMCadUnidade;
    fDMCadGrupo: TDMCadGrupo;
    fDMCadMarca: TDMCadMarca;
    fDMEstoque: TDMEstoque;

    Txt, Txt_Adi, Txt_XLS : TextFile;
    vOBSNao_Gravados: String;

    vRegistro : String;
    vArquivo : String;
    vArquivo_XLS : String;
    vArquivo_Adi : String;
    vContador  : Integer;

    gGrid: TStringGrid;
    linha, vColuna : Integer;

    vTipo_Reg: String;
    vAchou_ID: Boolean;

    procedure prc_Gravar_Produto;
    procedure prc_Gravar_Cliente;
    procedure prc_Le_XML(Tipo : String); //A= Atualizado  M=Materia Prima

    procedure prc_Carrega_Xml;

    function XlsToStringGrid2(AGrid: TStringGrid; AXLSFile: string; WorkSheet: Integer): Boolean;
    function fnc_verifica_Arquivo(NomeArquivo, Le_Grava : String) : String;
    function Replace(Str, Ant, Novo: string): string;
    function fnc_Verifica_Casas_Decimais(Campo : String) : String;

    function fnc_Grava_NCM(NCM : String) : Integer;
    procedure prc_Grava_Marca(ID, Nome : String);
    procedure prc_Grava_Grupo(ID, Nome : String);
    procedure prc_Gravar_Unidade(Unidade : String);

    procedure prc_Gravar_mAviso(ID : Integer; Nome, Tipo_Aviso, Obs, Referencia, Tipo_Reg: String);

    procedure prc_Gravar_Estoque;

  public
    { Public declarations }
  end;

var
  frmIntegrador: TfrmIntegrador;

implementation

uses rsDBUtils, StrUtils, DateUtils, uUtilPadrao, DmdDatabase,
  UInformeEndereco;

{$R *.dfm}

procedure TfrmIntegrador.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := Cafree;
end;

procedure TfrmIntegrador.btnProdutoClick(Sender: TObject);
begin
  if MessageDlg('Confirma a geração dos Produtos?',mtConfirmation,[mbYes,mbNo],0) <> mrYes then
    exit;
  vTipo_Reg := 'P';
  SMDBGrid1.ClearFilter;
  prc_Carrega_Xml;
end;

function TfrmIntegrador.XlsToStringGrid2(AGrid: TStringGrid;
  AXLSFile: string; WorkSheet: Integer): Boolean;
const
	xlCellTypeLastCell = $0000000B;
var
	XLApp, Sheet: OLEVariant;
	RangeMatrix: Variant;
	x, y, k, r: Integer;
begin
	Result := False;
	//Cria Excel- OLE Object
	XLApp  := CreateOleObject('Excel.Application');
	try
		//Esconde Excel
		XLApp.Visible:=False;

		//Abre o Workbook
		XLApp.Workbooks.Open(AXLSFile);

		//Setar na planilha desejada
		XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets[WorkSheet].Activate;

		// Para saber a dimensão do WorkSheet (o número de linhas e de colunas),
		//selecionamos a última célula não vazia do worksheet
		Sheet :=  XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets[WorkSheet];
		Sheet.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Select;

    //Pegar o número da última linha
    x:=XLApp.ActiveCell.Row;
		//x:=fDMExcel.cdsProduto.RecordCount;
    //Pegar o número da última coluna
		y:=XLApp.ActiveCell.Column;

		//Seta Stringgrid linha e coluna
		AGrid.RowCount:=x;
		AGrid.ColCount:=y;

		//Associa a variant WorkSheet com a variant do Delphi
		RangeMatrix:=XLApp.Range['A1', XLApp.Cells.Item[X, Y]].Value;

		//Cria o loop para listar os registros no TStringGrid
		k:=1;
		repeat
		  for r:=1 to y do
		  begin
			 AGrid.Cells[(r - 1),(k - 1)] := RangeMatrix[K, R];

			 //Redimensionar tamanho das colunas do grid dinamicamente
			 If (AGrid.ColWidths[r-1] < (Length(AGrid.Cells[(r - 1),(k - 1)]) * 8)) then
				AGrid.ColWidths[r-1] := Length(AGrid.Cells[(r - 1),(k - 1)]) * 8;

		  end;
		  Inc(k,1);
		until k > x;
		RangeMatrix := Unassigned;
	finally
		//Fecha o Excel
		if not VarIsEmpty(XLApp) then
		   begin
			  XLApp.Quit;
			  XLAPP:=Unassigned;
			  Sheet:=Unassigned;
			  Result:=True;
		   end;
	end;

end;


function TfrmIntegrador.fnc_verifica_Arquivo(NomeArquivo, Le_Grava : String) : String;
begin
  if copy(NomeArquivo,1,1) = '"' then
    delete(NomeArquivo,1,1);
  if copy(NomeArquivo,Length(NomeArquivo),1) = '"' then
    delete(NomeArquivo,Length(NomeArquivo),1);
  if (Le_Grava = 'G') and (copy(NomeArquivo,Length(NomeArquivo),1) = '\') then
    delete(NomeArquivo,Length(NomeArquivo),1);
  Result := NomeArquivo;
end;

function TfrmIntegrador.Replace(Str, Ant, Novo: string): string;
var
  iPos: Integer;
begin
  while Pos(Ant, Str) > 0 do
  begin
    iPos := Pos(Ant, Str);
    Str := copy(Str, 1, iPos - 1) + Novo + copy(Str, iPos + 1, Length(Str) - iPos);
  end;
  Result := Str;
end;

function TfrmIntegrador.fnc_Verifica_Casas_Decimais(Campo: String): String;
var
  i, x : Integer;
  vTexto2 : string;
begin
  Result := '';
  i := pos(',',Campo);
  if i > 0 then
    i := Length(campo) - i;
  if i < 2 then
  begin
    for x := 1 to 2 - i do
      result := result + '0';
  end;
end;

procedure TfrmIntegrador.prc_Le_XML(Tipo : String); //A= Atualizado  M=Materia Prima
begin
  ceTotal.AsInteger    := gGrid.RowCount;
  ceLidos.AsInteger    := 0;
  ceGravados.AsInteger := 0;
  Linha     := 0;
  vContador := 0;
  try
    if vTipo_Reg = 'F' then
    begin
      fDMCadPessoa.prc_Localizar(-1);
      fDMCadPessoa := TDMCadPessoa.Create(Self);
    end
    else
    begin
      fDMCadProduto := TDMCadProduto.Create(Self);
      fDMCadNCM     := TDMCadNCM.Create(Self);
      fDMCadUnidade := TDMCadUnidade.Create(Self);
      fDMCadGrupo   := TDMCadGrupo.Create(Self);
      fDMCadMarca   := TDMCadMarca.Create(Self);
      fDMEstoque    := TDMEstoque.Create(Self);
    end;

    vAchou_ID := False;
    vOBSNao_Gravados := '';
    while Linha < gGrid.RowCount -1 do
    begin
      Linha := Linha + 1;
      ceLidos.AsInteger := ceLidos.AsInteger + 1;
      if vTipo_Reg = 'P' then
        prc_Gravar_Produto
      else
        prc_Gravar_Cliente;
      ceGravados.AsInteger := vContador;
      Application.ProcessMessages;
    end;

  finally
    begin
      if vTipo_Reg = 'F' then
        FreeAndNil(fDMCadPessoa)
      else
      begin
        FreeAndNil(fDMCadProduto);
        FreeAndNil(fDMCadNCM);
        FreeAndNil(fDMCadUnidade);
        FreeAndNil(fDMCadGrupo);
        FreeAndNil(fDMCadMarca);
        FreeAndNil(fDMEstoque);
      end;
    end;
  end;
end;

procedure TfrmIntegrador.prc_Gravar_Produto;
var
  sds: TSQLDataSet;
  vCodigo : String;
  vNome1, vNome2, vNome3 : String;
  vTexto1 : String;
  vPesoLiquido, vVlrCompra, vVlrVenda : Real;
  vQtd_Estoque: Real;
  vTexto2: String;
  vPrecoVenda: Real;
begin
  vTexto1 := gGrid.Cells[0,Linha];
  if not(vAchou_ID) and (trim(UpperCase(vTexto1)) <> 'ID') then
    exit
  else
  if (trim(UpperCase(vTexto1)) = 'ID') then
  begin
    vAchou_ID := True;
    exit;
  end;
  if (trim(vTexto1) = '') then
    exit;

  vTexto1 := trim(gGrid.Cells[2,Linha]);
  if not(chkGerarProdutos.Checked) then
  begin
    if ((trim(vTexto1) = '00000000') or (Length(trim(vTexto1)) < 8)) then
    begin
      prc_Gravar_mAviso(fDMCadProduto.cdsProdutoID.AsInteger,fDMCadProduto.cdsProdutoNOME.AsString,
                        'Erro','Produto com problema no NCM (produto NÃO gravado)',gGrid.Cells[1,Linha],'Produto');
      exit;
    end;
    vTexto1     := Monta_Numero(trim(gGrid.Cells[6,Linha]),1);
    vPrecoVenda := StrToFloat(vTexto1);
    if vPrecoVenda <= 0 then
    begin
      prc_Gravar_mAviso(fDMCadProduto.cdsProdutoID.AsInteger,fDMCadProduto.cdsProdutoNOME.AsString,
                        'Erro','Preço de venda não informado (produto NÃO gravado)',gGrid.Cells[1,Linha],'Produto');
      exit;
    end;
  end;
  fDMCadProduto.prc_Localizar(StrToInt(vTexto1));

  if fDMCadProduto.cdsProdutoID.AsInteger > 0 then
    fDMCadProduto.cdsProduto.Edit
  else
  begin
    fDMCadProduto.prc_Inserir;
    fDMCadProduto.cdsProdutoID.AsInteger        := StrToInt(vTexto1);
    fDMCadProduto.cdsProdutoREFERENCIA.AsString := gGrid.Cells[1,Linha];
    fDMCadProduto.cdsProdutoCOD_ANT.AsString    := vTexto1;
  end;
  fDMCadProduto.cdsProdutoNOME.AsString          := gGrid.Cells[3,Linha];
  fDMCadProduto.cdsProdutoNOME_ORIGINAL.AsString := gGrid.Cells[3,Linha];

  vTexto1 := Monta_Numero(trim(gGrid.Cells[2,Linha]),8);
  if trim(vTexto1) = '00000000' then
    prc_Gravar_mAviso(fDMCadProduto.cdsProdutoID.AsInteger,fDMCadProduto.cdsProdutoNOME.AsString,
                      'Aviso','Produto sem NCM no arquivo EXCEL (produto gravado)',gGrid.Cells[1,Linha],'Produto')
  else
  begin
    vTexto2 := Monta_Numero(SQLLocate('TAB_NCM','NCM','ID',vTexto1),1);
    if vTexto2 <> '0' then
      fDMCadProduto.cdsProdutoID_NCM.AsInteger := StrToInt(vTexto2)
    else
      fDMCadProduto.cdsProdutoID_NCM.AsInteger := fnc_Grava_NCM(vTexto1);
  end;

  vTexto1 := UpperCase(trim(gGrid.Cells[4,Linha]));
  if trim(vTexto1) = '' then
    vTexto1  := 'UN';
  fDMCadProduto.cdsProdutoUNIDADE.AsString := vTexto1;
  vTexto2 := SQLLocate('UNIDADE','UNIDADE','UNIDADE',vTexto1);
  if trim(vTexto2) = '' then
    prc_Gravar_Unidade(vTexto1);

  vTexto1 := trim(gGrid.Cells[5,Linha]);
  fDMCadProduto.cdsProdutoPRECO_CUSTO.AsString       := vTexto1;
  fDMCadProduto.cdsProdutoPRECO_CUSTO_TOTAL.AsString := vTexto1;
  vTexto1 := trim(gGrid.Cells[6,Linha]);
  fDMCadProduto.cdsProdutoPRECO_VENDA.AsString := vTexto1;
  if fDMCadProduto.cdsProdutoPRECO_VENDA.AsFloat <= 0 then
    prc_Gravar_mAviso(fDMCadProduto.cdsProdutoID.AsInteger,fDMCadProduto.cdsProdutoNOME.AsString,
                      'Aviso','Preço de venda não informado',gGrid.Cells[1,Linha],'Produto');
  if chkFornecedor.Checked then
  begin
    vTexto1 := trim(gGrid.Cells[8,Linha]);
    if trim(vTexto1) <> '' then
    begin
      vTexto2 := SQLLocate('PESSOA','CNPJ_CPF','CODIGO',vTexto1);
      if trim(vTexto2) <> '' then
        fDMCadProduto.cdsProdutoID_FORNECEDOR.AsString := vTexto2
      else
        prc_Gravar_mAviso(fDMCadProduto.cdsProdutoID.AsInteger,fDMCadProduto.cdsProdutoNOME.AsString,
                          'Aviso','Fornecedor não encontrado no CNPJ/CPF ' + vTexto1,gGrid.Cells[1,Linha], 'Produto');
    end;
  end;

  vTexto1 := Monta_Numero(UpperCase(trim(gGrid.Cells[9,Linha])),1);
  if trim(vTexto1) <> '0' then
  begin
    vTexto2 := Monta_Numero( SQLLocate('MARCA','ID','ID',vTexto1),1);
    if vTexto2 = '0' then
      prc_Grava_Marca(vTexto1,UpperCase(gGrid.Cells[10,Linha]));
    fDMCadProduto.cdsProdutoID_MARCA.AsInteger := StrToInt(vTexto1);
  end;

  vTexto1 := Monta_Numero(UpperCase(trim(gGrid.Cells[11,Linha])),1);
  if StrToInt(vTexto1) > 0 then
  begin
    vTexto2 := Monta_Numero( SQLLocate('GRUPO','ID','ID',vTexto1),1);
    if vTexto2 = '0' then
      prc_Grava_Grupo(vTexto1,UpperCase(gGrid.Cells[12,Linha]));
    fDMCadProduto.cdsProdutoID_GRUPO.AsInteger := StrToInt(vTexto1);
  end;

  fDMCadProduto.cdsProdutoLOCALIZACAO.AsString := UpperCase(gGrid.Cells[13,Linha]);
  vTexto1 := UpperCase(gGrid.Cells[14,Linha]);
  if trim(vTexto1) <> '' then
    fDMCadProduto.cdsProdutoORIGEM_PROD.AsString := vTexto1
  else
    fDMCadProduto.cdsProdutoORIGEM_PROD.AsString := '0';

  vTexto1 := UpperCase(gGrid.Cells[15,Linha]);
  if trim(vTexto1) <> '' then
  begin
    if StrToDate(vTexto1) > 10 then
      fDMCadProduto.cdsProdutoDTCAD.AsString := vTexto1
    else
      fDMCadProduto.cdsProdutoDTCAD.AsDateTime := Date;
  end
  else
    fDMCadProduto.cdsProdutoDTCAD.AsDateTime := Date;

  vTexto1 := Monta_Numero(gGrid.Cells[16,Linha],1);
  fDMCadProduto.cdsProdutoQTD_ESTOQUE_MIN.AsString := vTexto1; 
  fDMCadProduto.cdsProdutoSPED_TIPO_ITEM.AsString  := '00';

  fDMCadProduto.cdsProdutoPESOLIQUIDO.AsFloat       := 0;
  fDMCadProduto.cdsProdutoPESOBRUTO.AsFloat         := 0;
  fDMCadProduto.cdsProdutoINATIVO.AsString          := 'N';
  fDMCadProduto.cdsProdutoTIPO_REG.AsString         := 'P';
  fDMCadProduto.cdsProdutoPOSSE_MATERIAL.AsString  := 'E';
  fDMCadProduto.cdsProdutoESTOQUE.AsString         := 'S';
  fDMCadProduto.cdsProdutoMATERIAL_OUTROS.AsString := 'M';
  fDMCadProduto.cdsProdutoUSUARIO.AsString         := 'Integrador';
  //fDMCadProduto.cdsProdutoDTCAD.AsDateTime         := Date;
  fDMCadProduto.cdsProdutoHRCAD.AsDateTime         := Now;
  fDMCadProduto.cdsProdutoPERC_MARGEMLUCRO.AsFloat := 0;
  fDMCadProduto.cdsProdutoCOD_BARRA.Clear;
  fDMCadProduto.cdsProdutoTIPO_VENDA.AsString      := 'R';
  fDMCadProduto.cdsProdutoUSA_NA_BALANCA.Clear;
  fDMCadProduto.cdsProdutoUSA_GRADE.AsString       := 'N';
  fDMCadProduto.prc_Gravar;

  vContador := vContador + 1;

  vTexto1 := Monta_Numero(gGrid.Cells[15,Linha],1);
  if (trim(vTexto1) <> '0') and (trim(copy(vTexto1,1,1)) <> '-') and (trim(copy(vTexto1,1,1)) <> '0') then
    prc_Gravar_Estoque; 
end;

procedure TfrmIntegrador.prc_Gravar_Cliente;
var
  i, i2 : Integer;
  vTexto: String;
  vTexto2: String;
begin
  fDMCadPessoa.prc_Inserir;
  fDMCadPessoa.cdsPessoaCOD_ANT.AsString  := gGrid.Cells[0,Linha];
  fDMCadPessoa.cdsPessoaNOME.AsString     := UpperCase(gGrid.Cells[1,Linha]);
  fDMCadPessoa.cdsPessoaFANTASIA.AsString := UpperCase(gGrid.Cells[2,Linha]);
  if gGrid.Cells[3,Linha] = 'F-Fornecedor' then
    fDMCadPessoa.cdsPessoaTP_FORNECEDOR.AsString := 'S'
  else
    fDMCadPessoa.cdsPessoaTP_CLIENTE.AsString := 'S';
  fDMCadPessoa.cdsPessoaCNPJ_CPF.AsString := gGrid.Cells[4,Linha];
  if length(fDMCadPessoa.cdsPessoaCNPJ_CPF.AsString) > 14 then
    fDMCadPessoa.cdsPessoaPESSOA.AsString := 'J'
  else
    fDMCadPessoa.cdsPessoaPESSOA.AsString := 'F';
  fDMCadPessoa.cdsPessoaINSCR_EST.AsString := gGrid.Cells[5,Linha];
  vTexto := gGrid.Cells[7,Linha];
  if Length(vTexto) > 0 then
  begin
    vTexto2 := '';
    i  := pos('(',vTexto);
    i2 := pos(')',vTexto);
    if i > 0 then
    begin
      vTexto2 := Monta_Numero(copy(vTexto,i+1,i2-(i+1)),0);
      if trim(vTexto2) = '' then
        fDMCadPessoa.cdsPessoaDDDFONE1.AsInteger := StrToInt(vTexto2);
      delete(vTexto,1,i2);
      fDMCadPessoa.cdsPessoaTELEFONE1.AsString := vTexto;
    end;
  end;
  vTexto  := gGrid.Cells[8,Linha];
  i := pos(',',vTexto);
  if i > 0 then
  begin
    vTexto2 := copy(vTexto,1,i-1);
    delete(vTexto,1,i);
    fDMCadPessoa.cdsPessoaENDERECO.AsString := vTexto2;
    fDMCadPessoa.cdsPessoaNUM_END.AsString  := vTexto;
  end
  else
    fDMCadPessoa.cdsPessoaENDERECO.AsString := vTexto;
  fDMCadPessoa.cdsPessoaCOMPLEMENTO_END.Clear;
  fDMCadPessoa.cdsPessoaBAIRRO.AsString := gGrid.Cells[9,Linha];
  vTexto := UpperCase(TirarAcento(gGrid.Cells[10,Linha]));
  if fDMCadPessoa.cdsCidade.locate('NOME',vTexto,[loCaseInsensitive]) then
  begin
    fDMCadPessoa.cdsPessoaID_CIDADE.AsInteger := fDMCadPessoa.cdsCidadeID.AsInteger;
    fDMCadPessoa.cdsPessoaUF.AsString         := fDMCadPessoa.cdsCidadeUF.AsString;
    fDMCadPessoa.cdsPessoaCIDADE.AsString     := fDMCadPessoa.cdsCidadeNOME.AsString;
  end
  else
  begin
    fDMCadPessoa.cdsPessoaID_CIDADE.Clear;
    fDMCadPessoa.cdsPessoaUF.AsString         := 'RS';
    fDMCadPessoa.cdsPessoaCIDADE.AsString     := UpperCase(TirarAcento(vTexto));
  end;
  fDMCadPessoa.cdsPessoaID_PAIS.AsInteger  := 1;
  fDMCadPessoa.cdsPessoaCEP.Clear;
  fDMCadPessoa.cdsPessoaUSUARIO.AsString   := 'Conversor';
  fDMCadPessoa.cdsPessoaDTCADASTRO.AsDateTime := Date;
  fDMCadPessoa.cdsPessoaHRCADASTRO.AsDateTime := Now;
  fDMCadPessoa.cdsPessoaNOME_CONTATO.Clear;
  vTexto := Monta_Numero(gGrid.Cells[12,Linha],0);
  if (trim(vTexto) <> '') and (trim(vTexto) <> '0') then
    fDMCadPessoa.cdsPessoaID_VENDEDOR.AsString := vTexto;
  fDMCadPessoa.cdsPessoaOBS.Clear;
  fDMCadPessoa.cdsPessoaCAIXAPOSTAL.Clear;
  fDMCadPessoa.cdsPessoaRG.Clear;      
  fDMCadPessoa.cdsPessoaPERC_COMISSAO.Clear;
  fDMCadPessoa.cdsPessoaINATIVO.AsString := 'N';
  fDMCadPessoa.cdsPessoaHOMEPAGE.Clear;
  fDMCadPessoa.cdsPessoaEMAIL_NFE.Clear;
  fDMCadPessoa.cdsPessoaEMAIL_PGTO.Clear; 
  fDMCadPessoa.cdsPessoaEMAIL_NFE2.Clear;
  fDMCadPessoa.cdsPessoaTP_TRANSPORTADORA.AsString := 'N';
  fDMCadPessoa.cdsPessoaID_REGIME_TRIB.Clear;
  fDMCadPessoa.cdsPessoaTIPO_COMISSAO.Clear;
  fDMCadPessoa.cdsPessoaEMAIL_COMPRAS.Clear;
  fDMCadPessoa.cdsPessoaCONTATO_COMPRAS.Clear;
  fDMCadPessoa.cdsPessoaORGAO_PUBLICO.AsString := 'N';
  fDMCadPessoa.cdsPessoaPERC_REDUCAO_INSS.Clear;
  fDMCadPessoa.cdsPessoaPAI_NOME.Clear;
  fDMCadPessoa.cdsPessoaMAE_NOME.Clear;
  fDMCadPessoa.cdsPessoaVLR_LIMITE_CREDITO.Clear;
  //if (UpperCase(fDMCadPessoa.cdsPessoaINSCR_EST.AsString) <> 'ISENTO') and (trim(fDMCadPessoa.cdsPessoaINSCR_EST.AsString) <> '') then
  //  fDMCadPessoa.cdsPessoaTIPO_CONSUMIDOR.AsString := '0'
  //else
  //if (UpperCase(fDMCadPessoa.cdsPessoaINSCR_EST.AsString) = 'ISENTO') then
    //fDMCadPessoa.cdsPessoaTIPO_CONSUMIDOR.AsString := '1';
  //if fDMProdutoSMALL.cdsCliForTP_CONTRIBUINTE.AsString = 'C' then
  //  fDMCadPessoa.cdsPessoaTIPO_CONTRIBUINTE.AsInteger := 1
  //else
    //fDMCadPessoa.cdsPessoaTIPO_CONTRIBUINTE.AsInteger := 9;
   fDMCadPessoa.cdsPessoaTIPO_CONTRIBUINTE.AsInteger := 9;
  fDMCadPessoa.cdsPessoaTIPO_CONSUMIDOR.AsInteger    := 1;

  //fDMCadPessoa.cdsPessoaCOD_ALFA.AsString := fDMProdutoSMALL.cdsCliForCD_CLIFOR.AsString;
  fDMCadPessoa.cdsPessoaPERC_COMISSAO_VEND.Clear;
  fDMCadPessoa.cdsPessoaPERC_DESC_SUFRAMA.Clear;
  fDMCadPessoa.cdsPessoaVLR_LIMITE_COMPRA.Clear;
  fDMCadPessoa.cdsPessoaUSUARIO_LOG.AsString := 'Conversor';
  fDMCadPessoa.cdsPessoaPROTESTAR.Clear;
  fDMCadPessoa.cdsPessoaDESC_MAXIMO.Clear;

  fDMCadPessoa.cdsPessoa.Post;

  fDMCadPessoa.cdsPessoa.ApplyUpdates(0);
end;

procedure TfrmIntegrador.FormShow(Sender: TObject);
begin
  fDMIntegrador := TDMIntegrador.Create(Self);
  oDBUtils.SetDataSourceProperties(Self, fDMIntegrador);
end;

procedure TfrmIntegrador.prc_Carrega_Xml;
begin
  if trim(File_Produto.Text) = '' then
  begin
    MessageDlg('*** Aquivo não informado!', mtError, [mbOk], 0);
    File_Produto.SetFocus;
    exit;
  end;
  fDMIntegrador.mAvisos.EmptyDataSet;

  gGrid := TStringGrid.Create(gGrid);
  try
    vArquivo_XLS := fnc_verifica_Arquivo(File_Produto.Text,'L');
    XlsToStringGrid2(gGrid,vArquivo_XLS,1);
    prc_Le_XML('');
  finally
    FreeAndNil(gGrid);
  end;
  if vTipo_Reg = 'P' then
    MessageDlg('*** Produtos Convertidos!', mtConfirmation, [mbOk], 0)
  else
    MessageDlg('*** Fornecedores Convertidos!', mtConfirmation, [mbOk], 0);
end;

function TfrmIntegrador.fnc_Grava_NCM(NCM: String): Integer;
var
  vID : Integer;
begin
  fDMCadNCM.prc_Inserir;
  vID    := fDMCadNCM.cdsNCMID.AsInteger;
  Result := vID; 
  fDMCadNCM.cdsNCMNCM.AsString  := Monta_Numero(NCM,8);
  fDMCadNCM.cdsNCMNOME.AsString := '';
  fDMCadNCM.cdsNCMGERAR_ST.AsString := 'N';
  fDMCadNCM.cdsNCMINATIVO.AsString  := 'N';
  fDMCadNCM.cdsNCMTIPO_AS.AsString  := 'A';
  fDMCadNCM.cdsNCMUSAR_MVA_UF_DESTINO.AsString  := 'S';
  fDMCadNCM.cdsNCMCOD_CEST.Clear;
  fDMCadNCM.prc_Gravar;
end;

procedure TfrmIntegrador.prc_Gravar_Unidade(Unidade: String);
var
  vNome: String;
begin
  fDMCadUnidade.prc_Inserir;
  vNome := Unidade + '.';
  fDMCadUnidade.cdsUnidadeUNIDADE.AsString := Unidade;
  fDMCadUnidade.cdsUnidadeNOME.AsString    := vNome;
  fDMCadUnidade.prc_Gravar;
end;

procedure TfrmIntegrador.prc_Gravar_mAviso(ID: Integer; Nome, Tipo_Aviso,
  Obs, Referencia, Tipo_Reg: String);
begin
  fDMIntegrador.mAvisos.Insert;
  fDMIntegrador.mAvisosID.AsInteger      := ID;
  fDMIntegrador.mAvisosNome.AsString     := Nome;
  fDMIntegrador.mAvisosTipo.AsString     := Tipo_Aviso;
  fDMIntegrador.mAvisosObs.AsString      := Obs;
  fDMIntegrador.mAvisosTipo_Reg.AsString := Tipo_Reg;
  fDMIntegrador.mAvisos.Post;
end;

procedure TfrmIntegrador.prc_Grava_Marca(ID, Nome: String);
begin
  fDMCadMarca.prc_Inserir;
  fDMCadMarca.cdsMarcaID.AsInteger  := StrToInt(ID);
  fDMCadMarca.cdsMarcaNOME.AsString := Nome;
  fDMCadMarca.prc_Gravar;
end;

procedure TfrmIntegrador.prc_Grava_Grupo(ID, Nome: String);
begin
  fDMCadGrupo.prc_Inserir;
  fDMCadGrupo.cdsGrupoID.AsInteger    := StrToInt(ID);
  fDMCadGrupo.cdsGrupoNOME.AsString   := Nome;
  fDMCadGrupo.cdsGrupoCODIGO.AsString := Monta_Numero(fDMCadGrupo.cdsGrupoID.AsString,3);
  fDMCadGrupo.cdsGrupoNIVEL.AsInteger := 1;
  fDMCadGrupo.cdsGrupoTIPO.AsString   := 'A';
  fDMCadGrupo.cdsGrupoCOD_PRINCIPAL.AsInteger := fDMCadGrupo.cdsGrupoID.AsInteger;
  fDMCadGrupo.cdsGrupoTIPO_PROD.AsString      := 'O';
  fDMCadGrupo.prc_Gravar;
end;

procedure TfrmIntegrador.prc_Gravar_Estoque;
var
  vQtd_Estoque : Real;
  vTexto1, vTexto2: String;
  vQtdAux: Real;
  vTipo_ES: String;
  vIDAux: Integer;
  vPrecoAux: String;
  vPrecoCustoAux: String;
  vGerarCusto: String;
begin
  vTexto1 := Monta_Numero(gGrid.Cells[0,Linha],1);
  vFilial := fnc_Buscar_Filial;
  vQtd_Estoque := fnc_Buscar_Estoque(StrToInt(vTexto1),1,0,vFilial);
  vTexto1 := Monta_Numero(gGrid.Cells[7,Linha],1);
  vQtdAux := StrToFloat(vTexto1);

  if StrToFloat(FormatFloat('0.0000',vQtd_Estoque)) = StrToFloat(FormatFloat('0.0000',vQtdAux)) then
    exit;

  vTipo_ES := 'S';
  if vQtdAux > vQtd_Estoque then
  begin
    vQtdAux  := vQtdAux - vQtd_Estoque;
    vTipo_ES := 'E';
  end
  else
    vQtdAux  := vQtd_Estoque - vQtdAux;
  vGerarCusto := 'S';
  if vTipo_ES = 'S' then
    vGerarCusto := 'N';

  vTexto1 := Monta_Numero(gGrid.Cells[0,Linha],1);
  if vTipo_ES = 'E' then
    vPrecoAux := Monta_Numero(gGrid.Cells[5,Linha],1)
  else
    vPrecoAux := Monta_Numero(gGrid.Cells[6,Linha],1);
  vPrecoCustoAux := Monta_Numero(gGrid.Cells[5,Linha],1);

  vIDAux := fDMEstoque.fnc_Gravar_Estoque(0, //ID_Estoque
                                          vFilial,
                                          1, //Local Estoque
                                          StrToInt(vTexto1), //ID Produto
                                          0, // Num Documento
                                          0, // ID Pessoa
                                          0, // ID CFOP
                                          0, // ID Nota
                                          0, // ID Centro Custo
                                          vTipo_ES,
                                          'INT', // Tipo Movimento
                                          Trim(UpperCase(gGrid.Cells[4,Linha])), //Unidade
                                          Trim(UpperCase(gGrid.Cells[4,Linha])), //Unidade
                                          '', //Serie
                                          '', //Tamanho
                                          Date,
                                          StrToFloat(vPrecoAux),
                                          vQtdAux,
                                          0, // % ICMS
                                          0, // % IPI
                                          0, // Valor Desconto
                                          0, // % Tributação
                                          0, // Valor Frete
                                          vQtdAux,
                                          StrToFloat(vPrecoAux),
                                          0, // Valor Desconto Original
                                          0, // Qtd Pacote
                                          '', // Unidade Interna
                                          0, // ID Cor
                                          '', // Número Lote Controle
                                          vGerarCusto,
                                          StrToFloat(vPrecoCustoAux),
                                          0, // Comprimento
                                          0, // Largura
                                          0, // Espessura
                                          0, // ID Operação
                                          0, // ID Pedido
                                          0); //Item Pedido );


end;

procedure TfrmIntegrador.btnExcelClick(Sender: TObject);
var
  vAno,vMes, vDia: Word;
begin
  frmInformeEndereco := TfrmInformeEndereco.Create(self);
  try
    frmInformeEndereco.ShowModal;
  finally
    FreeAndNil(frmInformeEndereco);
  end;
  if copy(vEndereco_Arq,Length(vEndereco_Arq),1) <> '\' then
    vEndereco_Arq := vEndereco_Arq + '\';
  DecodeDate(Date,vAno,vMes,vDia);
  prc_Preencher_CSV(SMDBGrid1.DataSource, SMDBGrid1,'Avisos_' + FormatFloat('0000',vAno) + '_' + FormatFloat('00',vMes) + '.CSV')
end;

end.
