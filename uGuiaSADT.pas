unit uGuiaSADT;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Data.DB,
  Datasnap.DBClient, Vcl.ExtCtrls, Vcl.Buttons, Vcl.Mask, Vcl.DBCtrls,
  Vcl.Grids, Vcl.DBGrids, Vcl.ComCtrls, IdHashMessageDigest, Xml.xmldom,
  Xml.XMLIntf, Xml.XMLDoc, System.Win.ComObj;

type
  TForm6 = class(TForm)
    DataSource1: TDataSource;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    Panel1: TPanel;
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Edit1: TDBEdit;
    Edit2: TDBEdit;
    GroupBox2: TGroupBox;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Edit3: TDBEdit;
    Edit4: TDBEdit;
    Edit5: TDBEdit;
    GroupBox3: TGroupBox;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Edit6: TDBEdit;
    Edit7: TDBEdit;
    Edit8: TDBEdit;
    GroupBox4: TGroupBox;
    GroupBox5: TGroupBox;
    Label9: TLabel;
    Label31: TLabel;
    Edit9: TDBEdit;
    Edit31: TDBEdit;
    GroupBox6: TGroupBox;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Edit11: TDBEdit;
    Edit12: TDBEdit;
    Edit13: TDBEdit;
    Edit14: TDBEdit;
    Edit15: TDBEdit;
    GroupBox7: TGroupBox;
    Label16: TLabel;
    Edit16: TDBEdit;
    GroupBox8: TGroupBox;
    Label17: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    Edit17: TDBEdit;
    Edit18: TDBEdit;
    Edit19: TDBEdit;
    GroupBox9: TGroupBox;
    Label20: TLabel;
    Label21: TLabel;
    Edit20: TDBEdit;
    Edit21: TDBEdit;
    GroupBox10: TGroupBox;
    Label22: TLabel;
    Label23: TLabel;
    Label24: TLabel;
    Label25: TLabel;
    Label26: TLabel;
    Label27: TLabel;
    Label28: TLabel;
    Label29: TLabel;
    Label30: TLabel;
    Label32: TLabel;
    Edit22: TDBEdit;
    Edit23: TDBEdit;
    Edit24: TDBEdit;
    Edit25: TDBEdit;
    Edit26: TDBEdit;
    Edit27: TDBEdit;
    Edit28: TDBEdit;
    Edit29: TDBEdit;
    Edit30: TDBEdit;
    Edit32: TDBEdit;
    GroupBox11: TGroupBox;
    Label33: TLabel;
    Label34: TLabel;
    Label35: TLabel;
    Label36: TLabel;
    Label37: TLabel;
    Label38: TLabel;
    Label39: TLabel;
    Edit33: TDBEdit;
    Edit34: TDBEdit;
    Edit35: TDBEdit;
    Edit36: TDBEdit;
    Edit37: TDBEdit;
    Edit38: TDBEdit;
    Edit39: TDBEdit;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    DBGrid1: TDBGrid;
    cdsTISS: TClientDataSet;
    cdsTISSRegistroANS: TStringField;
    cdsTISSnumeroLote: TStringField;
    cdsTISSnumeroGuiaPrestador: TStringField;
    cdsTISSdataAutorizacao: TStringField;
    cdsTISSsenha: TStringField;
    cdsTISSdataValidadeSenha: TStringField;
    cdsTISSnumeroCarteira: TStringField;
    cdsTISSatendimentoRN: TStringField;
    cdsTISSnomeBeneficiario: TStringField;
    cdsTISScodigoPrestadorNaOperadora: TStringField;
    cdsTISSnomeContratado: TStringField;
    cdsTISSnomeProfissional: TStringField;
    cdsTISSconselhoProfissional: TStringField;
    cdsTISSnumeroConselhoProfissional: TStringField;
    cdsTISSUF: TStringField;
    cdsTISSCBOS: TStringField;
    cdsTISScaraterAtendimento: TStringField;
    cdsTISSCNES: TStringField;
    cdsTISStipoAtendimento: TStringField;
    cdsTISSindicacaoAcidente: TStringField;
    cdsTISSdataExecucao: TStringField;
    cdsTISScodigoTabela: TStringField;
    cdsTISScodigoProcedimento: TStringField;
    cdsTISSdescricaoProcedimento: TStringField;
    cdsTISSquantidadeExecutada: TStringField;
    cdsTISSviaAcesso: TStringField;
    cdsTISStecnicaUtilizada: TStringField;
    cdsTISSreducaoAcrescimo: TStringField;
    cdsTISSvalorUnitario: TStringField;
    cdsTISSvalorTotal: TStringField;
    cdsTISSgrauPart: TStringField;
    cdsTISScodigoPrestadorNaOperadora2: TStringField;
    cdsTISSnomeProf: TStringField;
    cdsTISSconselho: TStringField;
    Panel2: TPanel;
    DBNavigator1: TDBNavigator;
    BitBtn3: TBitBtn;
    DBEdit1: TDBEdit;
    Label10: TLabel;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    Button1: TButton;
    Button2: TButton;
    XMLDocument1: TXMLDocument;
    cdsTISSnomeExecutante: TStringField;
    BitBtn4: TBitBtn;
    FileOpenDialog1: TFileOpenDialog;
    cdsTISSdataSolicitacao: TStringField;
    Label40: TLabel;
    DBEdit2: TDBEdit;
    cdsTISStipoConsulta: TStringField;
    Label41: TLabel;
    DBEdit3: TDBEdit;
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
  private
    procedure AdicionarTag(texto: string);
    procedure ValidarXML(caminho: string);
    { Private declarations }
  public
    { Public declarations }
    XML: TStringList;
  end;

var
  Form6: TForm6;

implementation

{$R *.dfm}

procedure TForm6.BitBtn1Click(Sender: TObject);
begin
  cdsTISS.Post;
  cdsTISS.SaveToFile( Extractfilepath(application.ExeName)+'SADT.xml',dfXML);
end;

procedure TForm6.BitBtn2Click(Sender: TObject);

  procedure CopyRecord;
  var varCopyData: Variant;
      i: Integer;
  begin
    with cdsTiss do
    begin
      varCopyData := VarArrayCreate([0, FieldCount-1], varVariant);
      for i := 0 to FieldCount-1 do
        varCopyData[i] := Fields[i].Value;

      Insert;

      for i := 0 to FieldCount-1 do
        Fields[i].Value := varCopyData[i];
    end;
  end;

begin

  CopyRecord;
  cdsTISSsenha.value := '';
  cdsTISSnumeroCarteira.value := '';
  cdsTISSdataAutorizacao.Value:= '';
  cdsTISSnumeroLote.Value:= '';

end;

procedure TForm6.AdicionarTag(texto: string);
begin
   XML.Add( texto );
end;

function MD5String(const Value: string): string;
var
  xMD5: TIdHashMessageDigest5;
begin
  xMD5 := TIdHashMessageDigest5.Create;
  try
    Result := xMD5.HashStringAsHex(Value);
  finally
    xMD5.Free;
  end;
end;

function MD5Arquivo(const Value: string): string;
var
  xMD5: TIdHashMessageDigest5;
  xArquivo: TFileStream;
begin
  xMD5 := TIdHashMessageDigest5.Create;
  xArquivo := TFileStream.Create(Value, fmOpenRead OR fmShareDenyWrite);
  try
    Result := xMD5.HashStreamAsHex(xArquivo);
  finally
    xArquivo.Free;
    xMD5.Free;
  end;
end;

procedure TForm6.BitBtn3Click(Sender: TObject);
var
  Conteudo: string;

  RegistroANS: String;
  numeroLote: String;
  numeroGuiaPrestador: String;
  dataAutorizacao: String;
  senha: String;
  dataValidadeSenha: String;
  numeroCarteira: String;
  atendimentoRN: String;
  nomeBeneficiario: String;
  codigoPrestadorNaOperadora: String;
  nomeContratado: String;
  nomeProfissional: String;
  conselhoProfissional: String;
  numeroConselhoProfissional: String;
  UF: String;
  CBOS: String;
  caraterAtendimento: String;
  CNES: String;
  tipoAtendimento: String;
  indicacaoAcidente: String;
  dataExecucao: String;
  codigoTabela: String;
  codigoProcedimento: String;
  descricaoProcedimento: String;
  quantidadeExecutada: String;
  viaAcesso: String;
  tecnicaUtilizada: String;
  reducaoAcrescimo: String;
  valorUnitario: String;
  valorTotal: String;
  grauPart: String;
  codigoPrestadorNaOperadora2: String;
  nomeProf: String;
  conselho: String;
  dataSolicitacao: string;
  tipoConsulta: string;
begin
   XML:= TStringList.Create;

   AdicionarTag('<?xml version="1.0" encoding="iso-8859-1"?>');

   AdicionarTag('<ansTISS:mensagemTISS xmlns:ansTISS="http://www.ans.gov.br/padroes/tiss/schemas">');

   AdicionarTag('<ansTISS:cabecalho>');
   AdicionarTag('<ansTISS:identificacaoTransacao>');
   AdicionarTag('  <ansTISS:tipoTransacao>ENVIO_LOTE_GUIAS</ansTISS:tipoTransacao>');
   AdicionarTag('  <ansTISS:sequencialTransacao>1</ansTISS:sequencialTransacao>');
   AdicionarTag(Format('  <ansTISS:dataRegistroTransacao>%s</ansTISS:dataRegistroTransacao>',[FormatDatetime('YYYY-MM-DD',date)] ));
   AdicionarTag(Format('  <ansTISS:horaRegistroTransacao>%s</ansTISS:horaRegistroTransacao>',[timetostr(time)] ));
   AdicionarTag(' </ansTISS:identificacaoTransacao>');

   AdicionarTag(' <ansTISS:origem> ');
   AdicionarTag('  <ansTISS:identificacaoPrestador> ');
   AdicionarTag(Format('     <ansTISS:codigoPrestadorNaOperadora>%s</ansTISS:codigoPrestadorNaOperadora>',['44631669']));
   AdicionarTag('   </ansTISS:identificacaoPrestador> ');
   AdicionarTag(' </ansTISS:origem>');

   AdicionarTag('<ansTISS:destino>');
   AdicionarTag(Format('  <ansTISS:registroANS>%s</ansTISS:registroANS>',['326305']));
   AdicionarTag('</ansTISS:destino>');

   AdicionarTag(' <ansTISS:Padrao>3.03.01</ansTISS:Padrao> ');
   AdicionarTag('</ansTISS:cabecalho>');

   AdicionarTag('		<ansTISS:prestadorParaOperadora> ');
	 AdicionarTag('		  <ansTISS:loteGuias>  ');
	 AdicionarTag(Format('		  	<ansTISS:numeroLote>%s</ansTISS:numeroLote>  ',[cdsTISSnumeroLote.Asstring]));
	 AdicionarTag('		      <ansTISS:guiasTISS>  ');

   cdsTISS.First;
   while not cdsTISS.Eof do
   begin
     RegistroANS                 := cdsTISSRegistroANS.Asstring;
     numeroGuiaPrestador         := cdsTISSnumeroGuiaPrestador.Asstring;
     dataAutorizacao             := cdsTISSdataAutorizacao.Asstring;
     senha                       := cdsTISSsenha.Asstring;
     dataValidadeSenha           := cdsTISSdataValidadeSenha.Asstring;
     numeroCarteira              := cdsTISSnumeroCarteira.Asstring;
     atendimentoRN               := cdsTISSatendimentoRN.Asstring;
     nomeBeneficiario            := cdsTISSnomeBeneficiario.Asstring;
     codigoPrestadorNaOperadora  := cdsTISScodigoPrestadorNaOperadora.Asstring;
     nomeContratado              := cdsTISSnomeContratado.Asstring;
     nomeProfissional            := cdsTISSnomeProfissional.Asstring;
     conselhoProfissional        := cdsTISSconselhoProfissional.Asstring;
     numeroConselhoProfissional  := cdsTISSnumeroConselhoProfissional.Asstring;
     UF                          := cdsTISSUF.Asstring;
     CBOS                        := cdsTISSCBOS.Asstring;
     caraterAtendimento          := cdsTISScaraterAtendimento.Asstring;
     CNES                        := cdsTISSCNES.Asstring;
     tipoAtendimento             := cdsTISStipoAtendimento.Asstring;
     indicacaoAcidente           := cdsTISSindicacaoAcidente.Asstring;
     dataExecucao                := cdsTISSdataExecucao.Asstring;
     codigoTabela                := cdsTISScodigoTabela.Asstring;
     codigoProcedimento          := cdsTISScodigoProcedimento.Asstring;
     descricaoProcedimento       := cdsTISSdescricaoProcedimento.Asstring;
     quantidadeExecutada         := cdsTISSquantidadeExecutada.Asstring;
     viaAcesso                   := cdsTISSviaAcesso.Asstring;
     tecnicaUtilizada            := cdsTISStecnicaUtilizada.Asstring;
     reducaoAcrescimo            := cdsTISSreducaoAcrescimo.Asstring;
     valorUnitario               := cdsTISSvalorUnitario.Asstring;
     valorTotal                  := cdsTISSvalorTotal.Asstring;
     grauPart                    := cdsTISSgrauPart.Asstring;
     codigoPrestadorNaOperadora2 := cdsTISScodigoPrestadorNaOperadora2.Asstring;
     nomeProf                    := cdsTISSnomeProf.Asstring;
     conselho                    := cdsTISSconselho.Asstring;
     dataSolicitacao             := cdsTISSdataSolicitacao.Asstring;
     tipoConsulta                := cdsTISStipoConsulta.AsString;

      AdicionarTag('<ansTISS:guiaSP-SADT>');
      AdicionarTag('<ansTISS:cabecalhoGuia>');
      AdicionarTag(Format('	<ansTISS:registroANS>%s</ansTISS:registroANS>',[registroANS]));
      AdicionarTag(Format('	<ansTISS:numeroGuiaPrestador>%s</ansTISS:numeroGuiaPrestador> ',[numeroGuiaPrestador]));
      AdicionarTag('</ansTISS:cabecalhoGuia>');

      AdicionarTag('<ansTISS:dadosAutorizacao> ');
      AdicionarTag(Format('	<ansTISS:dataAutorizacao>%s</ansTISS:dataAutorizacao>',[dataAutorizacao]));
      AdicionarTag(Format('	<ansTISS:senha>%s</ansTISS:senha>',[senha]));
      AdicionarTag(Format('	<ansTISS:dataValidadeSenha>%s</ansTISS:dataValidadeSenha>',[dataValidadeSenha]));
      AdicionarTag('</ansTISS:dadosAutorizacao>');

      AdicionarTag('<ansTISS:dadosBeneficiario> ');
      AdicionarTag(Format('	<ansTISS:numeroCarteira>%s</ansTISS:numeroCarteira> ',[numeroCarteira]));
      AdicionarTag(Format('	<ansTISS:atendimentoRN>%s</ansTISS:atendimentoRN> ',[atendimentoRN]));
      AdicionarTag(Format('	<ansTISS:nomeBeneficiario>%s</ansTISS:nomeBeneficiario>',[nomeBeneficiario]));
      AdicionarTag('</ansTISS:dadosBeneficiario>');

      AdicionarTag('<ansTISS:dadosSolicitante>');
      AdicionarTag('	<ansTISS:contratadoSolicitante> ');
      AdicionarTag(Format('		<ansTISS:codigoPrestadorNaOperadora>%s</ansTISS:codigoPrestadorNaOperadora> ',[codigoPrestadorNaOperadora]));
      AdicionarTag(Format('		<ansTISS:nomeContratado>%s</ansTISS:nomeContratado>',[nomeContratado]));
      AdicionarTag('	</ansTISS:contratadoSolicitante>');

      AdicionarTag('	<ansTISS:profissionalSolicitante>');
      AdicionarTag(Format('		<ansTISS:nomeProfissional>%s</ansTISS:nomeProfissional> ',[nomeProfissional]));
      AdicionarTag(Format('		<ansTISS:conselhoProfissional>%s</ansTISS:conselhoProfissional>',[conselhoProfissional]));
      AdicionarTag(Format('		<ansTISS:numeroConselhoProfissional>%s</ansTISS:numeroConselhoProfissional>',[numeroConselhoProfissional]));
      AdicionarTag(Format('		<ansTISS:UF>%s</ansTISS:UF>',[UF]));
      AdicionarTag(Format('		<ansTISS:CBOS>%s</ansTISS:CBOS>',[CBOS]));
      AdicionarTag('	</ansTISS:profissionalSolicitante> ');
      AdicionarTag('</ansTISS:dadosSolicitante> ');

      AdicionarTag('<ansTISS:dadosSolicitacao>');
      AdicionarTag(Format(' <ansTISS:dataSolicitacao>%s</ansTISS:dataSolicitacao>',[dataSolicitacao]));
      AdicionarTag(Format('	<ansTISS:caraterAtendimento>%s</ansTISS:caraterAtendimento>',[caraterAtendimento]));
      AdicionarTag('</ansTISS:dadosSolicitacao> ');

      AdicionarTag('<ansTISS:dadosExecutante>');
      AdicionarTag('	<ansTISS:contratadoExecutante>');
      AdicionarTag(Format('		<ansTISS:codigoPrestadorNaOperadora>%s</ansTISS:codigoPrestadorNaOperadora>',[codigoPrestadorNaOperadora]));
      AdicionarTag(Format('		<ansTISS:nomeContratado>%s</ansTISS:nomeContratado>',[nomeContratado]));
      AdicionarTag('	</ansTISS:contratadoExecutante>');
      AdicionarTag(Format('	<ansTISS:CNES>%s</ansTISS:CNES>',[CNES]));
      AdicionarTag('</ansTISS:dadosExecutante> ');

      AdicionarTag('<ansTISS:dadosAtendimento> ');
      AdicionarTag(Format('	<ansTISS:tipoAtendimento>%s</ansTISS:tipoAtendimento> ',[tipoAtendimento]));
      AdicionarTag(Format('	<ansTISS:indicacaoAcidente>%s</ansTISS:indicacaoAcidente> ',[indicacaoAcidente]));
      AdicionarTag(Format(' <ansTISS:tipoConsulta>%s</ansTISS:tipoConsulta>',[tipoConsulta]));
      AdicionarTag('</ansTISS:dadosAtendimento> ');

      AdicionarTag('<ansTISS:procedimentosExecutados>');
      AdicionarTag('	<ansTISS:procedimentoExecutado>');
      AdicionarTag(Format('		<ansTISS:dataExecucao>%s</ansTISS:dataExecucao>',[dataExecucao]));
      AdicionarTag('		<ansTISS:procedimento>');
      AdicionarTag(Format('			<ansTISS:codigoTabela>%s</ansTISS:codigoTabela>',[codigoTabela]));
      AdicionarTag(Format('			<ansTISS:codigoProcedimento>%s</ansTISS:codigoProcedimento>',[codigoProcedimento]));
      AdicionarTag(Format('			<ansTISS:descricaoProcedimento>%s</ansTISS:descricaoProcedimento>',[descricaoProcedimento]));
      AdicionarTag('		</ansTISS:procedimento>');

      AdicionarTag(Format('		<ansTISS:quantidadeExecutada>%s</ansTISS:quantidadeExecutada>',[quantidadeExecutada]));

      if viaAcesso <> '' then
      AdicionarTag(Format('		<ansTISS:viaAcesso>%s</ansTISS:viaAcesso>',[viaAcesso]));
       if tecnicaUtilizada <> '' then
      AdicionarTag(Format('		<ansTISS:tecnicaUtilizada>%s</ansTISS:tecnicaUtilizada>',[tecnicaUtilizada]));
       if reducaoAcrescimo <> '' then
      AdicionarTag(Format('		<ansTISS:reducaoAcrescimo>%s</ansTISS:reducaoAcrescimo>',[reducaoAcrescimo]));

      AdicionarTag(Format('		<ansTISS:valorUnitario>%s</ansTISS:valorUnitario>',[valorUnitario]));
      AdicionarTag(Format('		<ansTISS:valorTotal>%s</ansTISS:valorTotal>',[valorTotal]));
      {AdicionarTag('		<ansTISS:equipeSadt>');
      AdicionarTag(Format('			<ansTISS:grauPart>%s</ansTISS:grauPart>',[grauPart]));

      AdicionarTag('			<ansTISS:codProfissional>');
      AdicionarTag(Format('				<ansTISS:codigoPrestadorNaOperadora>%s</ansTISS:codigoPrestadorNaOperadora>',[codigoPrestadorNaOperadora]));
      AdicionarTag('			</ansTISS:codProfissional>');

      AdicionarTag(Format('			<ansTISS:nomeProf>%s</ansTISS:nomeProf>',[nomeProf]));
      AdicionarTag(Format('			<ansTISS:conselho>%s</ansTISS:conselho>',[conselho]));
      AdicionarTag(Format('			<ansTISS:numeroConselhoProfissional>%s</ansTISS:numeroConselhoProfissional>',[numeroConselhoProfissional]));
      AdicionarTag(Format('			<ansTISS:UF>%s</ansTISS:UF>',[UF]));
      AdicionarTag(Format('			<ansTISS:CBOS>%s</ansTISS:CBOS>',[CBOS]));
      AdicionarTag('		</ansTISS:equipeSadt>'); }
      AdicionarTag('	</ansTISS:procedimentoExecutado>');
      AdicionarTag('</ansTISS:procedimentosExecutados>');

      AdicionarTag('<ansTISS:valorTotal>');
      AdicionarTag(Format('	<ansTISS:valorProcedimentos>%s</ansTISS:valorProcedimentos>',[ valorTotal]));
      AdicionarTag(Format('	<ansTISS:valorTotalGeral>%s</ansTISS:valorTotalGeral>',[ valorTotal]));
      AdicionarTag('</ansTISS:valorTotal>');

      AdicionarTag('</ansTISS:guiaSP-SADT>');

      cdsTISS.Next;
   end;

	 AdicionarTag('			</ansTISS:guiasTISS> ');
	 AdicionarTag('		</ansTISS:loteGuias> ');
	 AdicionarTag('	</ansTISS:prestadorParaOperadora>');

	 AdicionarTag('	    <ansTISS:epilogo>');
	 AdicionarTag(Format('		<ansTISS:hash>%s</ansTISS:hash> ',['733ecc3f3dfe45f4b7d95ccecbcb8879']));
	 AdicionarTag('	</ansTISS:epilogo>');

   AdicionarTag('</ansTISS:mensagemTISS> ');

   XML.SaveToFile( Extractfilepath(application.ExeName)+'Lote_Envio_GuiasSADT.xml');

   ValidarXML( Extractfilepath(application.ExeName)+'Lote_Envio_GuiasSADT.xml' );

end;

procedure TForm6.ValidarXML(caminho: string );
var
  path: string;
  XSDL:OleVariant;
  XML: OleVariant;
begin
    path:=ExtractFilePath(caminho);
    XSDL := CreateOLEObject('MSXML2.XMLSchemaCache.4.0');

    TRY
       XSDL.add('http://www.ans.gov.br/padroes/tiss/schemas', (path + '\schemas\tissV3_03_03.xsd'))
    EXCEPT
       showmessage('Schema tissV3.03.03 não encontrado!');
       exit;
    END;

   XML := CreateOLEObject('MSXML2.DOMDocument.4.0');
   XML.validateOnParse := True;
   XML.resolveExternals := True;
   XML.schemas := XSDL;
   XML.load(caminho);

   if (XML.parseError.errorCode <> 0) then
   begin
      showmessage('Descrição: ' +  StringReplace(trim(XML.parseError.Reason),
      '{http://www.ans.gov.br/padroes/tiss/schemas}','',[rfReplaceAll])+
      'Tag inválida: ' + trim(XML.parseError.srcText) + 'Linha ' +
      intToStr(XML.parseError.Line));
   end
   else
      showmessage('XML Válido!');
end;


procedure TForm6.BitBtn4Click(Sender: TObject);
begin
  if FileOpenDialog1.Execute then
  begin
    ValidarXML( FileOpenDialog1.FileName )
  end
  else
    exit;

end;

procedure TForm6.Button1Click(Sender: TObject);
begin
  cdsTISS.filter:= 'dataAutorizacao >= '''+ datetostr(DateTimePicker1.Date)+
            ''' and dataAutorizacao <= '''+ datetostr(DateTimePicker2.Date)+'''';
  cdsTISS.filtered:= true;
end;

procedure TForm6.Button2Click(Sender: TObject);
var
  Node, NodeChild: IXMLNode;
  i:integer;
  texto: string;

begin
  if not cdsTISS.Active then
    cdsTISS.CreateDataSet;

  if FileOpenDialog1.Execute then
  begin
    XMLDocument1.LoadFromFile(FileOpenDialog1.FileName);
    XMLDocument1.Active := true;
  end
  else
    exit;

  Node := XMLDocument1.DocumentElement.
           childNodes.FindNode('prestadorParaOperadora').
           childNodes.FindNode('loteGuias').
           childNodes.FindNode('guiasTISS').
           childNodes.FindNode('guiaSP-SADT');

  while (Node <> nil) do
  begin
    cdsTISS.append;
  //cdsTISSnumeroLote:= Node5.ChildNodes.FindNode('numeroGuiaPrestador').text;
    NodeChild := Node.childNodes.FindNode('cabecalhoGuia');

    texto:= NodeChild.ChildNodes.FindNode('registroANS').text;
    cdsTISSRegistroANS.AsString         := texto;

    texto:= NodeChild.ChildNodes.FindNode('numeroGuiaPrestador').text;
    cdsTISSnumeroGuiaPrestador.AsString := texto;

    NodeChild := Node.childNodes.FindNode('dadosAutorizacao');
    cdsTISSdataAutorizacao.AsString:= NodeChild.ChildNodes.FindNode('dataAutorizacao').text;
    cdsTISSsenha.AsString:= NodeChild.ChildNodes.FindNode('senha').text;
    cdsTISSdataValidadeSenha.AsString:= NodeChild.ChildNodes.FindNode('dataValidadeSenha').text;

    NodeChild := Node.childNodes.FindNode('dadosBeneficiario');
    cdsTISSnumeroCarteira.AsString:= NodeChild.ChildNodes.FindNode('numeroCarteira').text;
    cdsTISSatendimentoRN.AsString:= NodeChild.ChildNodes.FindNode('atendimentoRN').text;
    cdsTISSnomeBeneficiario.AsString:= NodeChild.ChildNodes.FindNode('nomeBeneficiario').text;

    NodeChild := Node.childNodes.FindNode('dadosSolicitante').childNodes.FindNode('contratadoSolicitante');
    cdsTISScodigoPrestadorNaOperadora.AsString:= NodeChild.ChildNodes.FindNode('codigoPrestadorNaOperadora').text;
    cdsTISSnomeContratado.AsString:= NodeChild.ChildNodes.FindNode('nomeContratado').text;

    NodeChild := Node.childNodes.FindNode('dadosSolicitante').childNodes.FindNode('profissionalSolicitante');

    cdsTISSnomeProfissional.AsString:= NodeChild.ChildNodes.FindNode('nomeProfissional').text;
    cdsTISSconselhoProfissional.AsString:= NodeChild.ChildNodes.FindNode('conselhoProfissional').text;
    cdsTISSnumeroConselhoProfissional.AsString:= NodeChild.ChildNodes.FindNode('numeroConselhoProfissional').text;
    cdsTISSUF.AsString:= NodeChild.ChildNodes.FindNode('UF').text;
    cdsTISSCBOS.AsString:= NodeChild.ChildNodes.FindNode('CBOS').text;

    NodeChild := Node.childNodes.FindNode('dadosSolicitacao');
    cdsTISSdataSolicitacao.AsString:= NodeChild.ChildNodes.FindNode('dataSolicitacao').text;
    cdsTISScaraterAtendimento.AsString:= NodeChild.ChildNodes.FindNode('caraterAtendimento').text;

    NodeChild := Node.childNodes.FindNode('dadosExecutante').childNodes.FindNode('contratadoExecutante');
    cdsTISScodigoPrestadorNaOperadora2.AsString:=  NodeChild.ChildNodes.FindNode('codigoPrestadorNaOperadora').text;
    cdsTISSnomeExecutante.AsString := NodeChild.ChildNodes.FindNode('nomeContratado').text;

    NodeChild := Node.childNodes.FindNode('dadosExecutante');
    cdsTISSCNES.AsString:= NodeChild.ChildNodes.FindNode('CNES').text;

     NodeChild := Node.childNodes.FindNode('dadosAtendimento');
    cdsTISStipoAtendimento.AsString:= NodeChild.ChildNodes.FindNode('tipoAtendimento').text;
    cdsTISSindicacaoAcidente.AsString:= NodeChild.ChildNodes.FindNode('indicacaoAcidente').text;
    cdsTISStipoConsulta.AsString:= NodeChild.ChildNodes.FindNode('tipoConsulta').text;

    NodeChild := Node.childNodes.FindNode('procedimentosExecutados').childNodes.FindNode('procedimentoExecutado');
    cdsTISSdataExecucao.AsString:= NodeChild.ChildNodes.FindNode('dataExecucao').text;
    cdsTISSquantidadeExecutada.AsString:= NodeChild.ChildNodes.FindNode('quantidadeExecutada').text;

    if NodeChild.ChildNodes.FindNode('viaAcesso') <> nil then
       cdsTISSviaAcesso.AsString:= NodeChild.ChildNodes.FindNode('viaAcesso').text;
    if NodeChild.ChildNodes.FindNode('tecnicaUtilizada') <> nil then
       cdsTISStecnicaUtilizada.AsString:= NodeChild.ChildNodes.FindNode('tecnicaUtilizada').text;
    if NodeChild.ChildNodes.FindNode('reducaoAcrescimo') <> nil then
       cdsTISSreducaoAcrescimo.AsString:= NodeChild.ChildNodes.FindNode('reducaoAcrescimo').text;

    cdsTISSvalorUnitario.AsString:= NodeChild.ChildNodes.FindNode('valorUnitario').text;
    cdsTISSvalorTotal.AsString:= NodeChild.ChildNodes.FindNode('valorTotal').text;

    NodeChild := NodeChild.childNodes.FindNode('procedimento');

    cdsTISScodigoTabela.AsString:= NodeChild.ChildNodes.FindNode('codigoTabela').text;
    cdsTISScodigoProcedimento.AsString:= NodeChild.ChildNodes.FindNode('codigoProcedimento').text;
    cdsTISSdescricaoProcedimento.AsString:= NodeChild.ChildNodes.FindNode('descricaoProcedimento').text;

    cdsTISS.Post;

    Node := Node.NextSibling;
  end;

end;

procedure TForm6.FormCreate(Sender: TObject);
begin
  // if FILEEXISTS( Extractfilepath(application.ExeName)+'SADT.xml' ) then
  //    cdsTISS.loadfromfile( Extractfilepath(application.ExeName)+'SADT.xml')
   //else
   //   cdsTISS.CreateDataSet;
end;

end.
