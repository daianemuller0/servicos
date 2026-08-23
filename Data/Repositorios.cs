using HowdenServicos.Poc.Models;

namespace HowdenServicos.Poc.Data;

/// <summary>Propostas de serviço gravadas (Banco de Dados › Propostas).</summary>
public class PropostaRepository
{
    private const string Entidade = "propostas";
    private readonly ParquetStore _store;
    public PropostaRepository(ParquetStore store) => _store = store;

    public List<Proposta> All() => _store.ReadLatest(Entidade,
        "id, cliente, cidade, contatoNome, contatoEmail, contatoTelefone, projeto, referencia, " +
        "numero, revisao, ano, data, bu, idioma, moeda, validadeDias, prazoEntregaDias, " +
        "preparadaPor, revisadaPor, representante, representante2, estado, segmento, marketSegment, vendaPara, destino, " +
        "assinaNome, assinaCargo, assinaEmail, assinaFones, " +
        "itensMoJson, itensDespesaJson, pricingJson, custoTotal, total, criadaEm, status",
        r => new Proposta
        {
            Id = S(r, 0), Cliente = S(r, 1), Cidade = S(r, 2), ContatoNome = S(r, 3),
            ContatoEmail = S(r, 4), ContatoTelefone = S(r, 5), Projeto = S(r, 6), Referencia = S(r, 7),
            Numero = S(r, 8), Revisao = S(r, 9), Ano = S(r, 10), Data = S(r, 11), Bu = S(r, 12),
            Idioma = S(r, 13), Moeda = S(r, 14), ValidadeDias = S(r, 15), PrazoEntregaDias = S(r, 16),
            PreparadaPor = S(r, 17), RevisadaPor = S(r, 18), Representante = S(r, 19),
            Representante2 = S(r, 20), Estado = S(r, 21),
            Segmento = S(r, 22), MarketSegment = S(r, 23), VendaPara = S(r, 24), Destino = S(r, 25),
            AssinaNome = S(r, 26), AssinaCargo = S(r, 27), AssinaEmail = S(r, 28), AssinaFones = S(r, 29),
            ItensMoJson = S(r, 30), ItensDespesaJson = S(r, 31), PricingJson = S(r, 32),
            CustoTotal = S(r, 33), Total = S(r, 34), CriadaEm = S(r, 35), Status = S(r, 36),
        }, "criadaEm DESC");

    public void Save(Proposta p) => _store.WriteRow(Entidade, new KeyValuePair<string, object?>[]
    {
        new("id", p.Id), new("cliente", p.Cliente), new("cidade", p.Cidade),
        new("contatoNome", p.ContatoNome), new("contatoEmail", p.ContatoEmail),
        new("contatoTelefone", p.ContatoTelefone), new("projeto", p.Projeto),
        new("referencia", p.Referencia), new("numero", p.Numero), new("revisao", p.Revisao),
        new("ano", p.Ano), new("data", p.Data), new("bu", p.Bu), new("idioma", p.Idioma),
        new("moeda", p.Moeda), new("validadeDias", p.ValidadeDias),
        new("prazoEntregaDias", p.PrazoEntregaDias), new("preparadaPor", p.PreparadaPor),
        new("revisadaPor", p.RevisadaPor), new("representante", p.Representante),
        new("representante2", p.Representante2), new("estado", p.Estado), new("segmento", p.Segmento), new("marketSegment", p.MarketSegment),
        new("vendaPara", p.VendaPara), new("destino", p.Destino), new("assinaNome", p.AssinaNome),
        new("assinaCargo", p.AssinaCargo), new("assinaEmail", p.AssinaEmail),
        new("assinaFones", p.AssinaFones), new("itensMoJson", p.ItensMoJson),
        new("itensDespesaJson", p.ItensDespesaJson), new("pricingJson", p.PricingJson),
        new("custoTotal", p.CustoTotal), new("total", p.Total), new("criadaEm", p.CriadaEm),
        new("status", p.Status),
    });

    public void Delete(string id) => _store.WriteRow(Entidade,
        new KeyValuePair<string, object?>[] { new("id", id) }, deleted: true);

    private static string S(System.Data.IDataReader r, int i) => r.IsDBNull(i) ? "" : r.GetString(i);
}

/// <summary>Tabela de custos padrão (custo/hora e despesas) — guia CUSTO.</summary>
public class ParametroRepository
{
    private const string Entidade = "parametros";
    private readonly ParquetStore _store;
    public ParametroRepository(ParquetStore store) => _store = store;

    public List<Parametro> All() => _store.ReadLatest(Entidade,
        "id, tipo, descricao, obs, horas, valor, porTecnico, ordem, mult",
        r => new Parametro
        {
            Id = S(r, 0), Tipo = S(r, 1), Descricao = S(r, 2), Obs = S(r, 3),
            Horas = S(r, 4), Valor = S(r, 5), PorTecnico = S(r, 6), Ordem = S(r, 7),
            Mult = S(r, 8) == "" ? "0" : S(r, 8),
        }, "ordem");

    public void Save(Parametro p) => _store.WriteRow(Entidade, new KeyValuePair<string, object?>[]
    {
        new("id", p.Id), new("tipo", p.Tipo), new("descricao", p.Descricao), new("obs", p.Obs),
        new("horas", p.Horas), new("valor", p.Valor), new("porTecnico", p.PorTecnico),
        new("ordem", p.Ordem), new("mult", p.Mult),
    });

    public void Delete(string id) => _store.WriteRow(Entidade,
        new KeyValuePair<string, object?>[] { new("id", id) }, deleted: true);

    private static string S(System.Data.IDataReader r, int i) => r.IsDBNull(i) ? "" : r.GetString(i);
}

/// <summary>Dados de faturamento por BU.</summary>
public class FaturamentoRepository
{
    private const string Entidade = "faturamento";
    private readonly ParquetStore _store;
    public FaturamentoRepository(ParquetStore store) => _store = store;

    public List<BillingInfo> All() => _store.ReadLatest(Entidade,
        "id, razao, endereco, registro, bancoNome, agencia, conta",
        r => new BillingInfo
        {
            Id = S(r, 0), Razao = S(r, 1), Endereco = S(r, 2), Registro = S(r, 3),
            BancoNome = S(r, 4), Agencia = S(r, 5), Conta = S(r, 6),
        }, "id");

    public void Save(BillingInfo b) => _store.WriteRow(Entidade, new KeyValuePair<string, object?>[]
    {
        new("id", b.Id), new("razao", b.Razao), new("endereco", b.Endereco),
        new("registro", b.Registro), new("bancoNome", b.BancoNome),
        new("agencia", b.Agencia), new("conta", b.Conta),
    });

    private static string S(System.Data.IDataReader r, int i) => r.IsDBNull(i) ? "" : r.GetString(i);
}

/// <summary>Representantes comerciais e suas comissões (BD_pricing da planilha).</summary>
public class RepresentanteRepository
{
    private const string Entidade = "representantes";
    private readonly ParquetStore _store;
    public RepresentanteRepository(ParquetStore store) => _store = store;

    public List<Representante> All() => _store.ReadLatest(Entidade,
        "id, nome, local, comissaoPct, contato",
        r => new Representante
        {
            Id = S(r, 0), Nome = S(r, 1), Local = S(r, 2),
            ComissaoPct = S(r, 3), Contato = S(r, 4),
        }, "nome");

    public void Save(Representante x) => _store.WriteRow(Entidade, new KeyValuePair<string, object?>[]
    {
        new("id", x.Id), new("nome", x.Nome), new("local", x.Local),
        new("comissaoPct", x.ComissaoPct), new("contato", x.Contato),
    });

    public void Delete(string id) => _store.WriteRow(Entidade,
        new KeyValuePair<string, object?>[] { new("id", id) }, deleted: true);

    private static string S(System.Data.IDataReader r, int i) => r.IsDBNull(i) ? "" : r.GetString(i);
}

/// <summary>Identidade visual (logo do cabeçalho da proposta), como data URI.</summary>
public class BrandingRepository
{
    private const string LogoId = "logo";
    private readonly ParquetStore _store;
    public BrandingRepository(ParquetStore store) => _store = store;

    public string? GetLogo() =>
        _store.ReadLatest("branding", "id, valor",
                r => new { Id = r.IsDBNull(0) ? "" : r.GetString(0), Valor = r.IsDBNull(1) ? "" : r.GetString(1) })
            .Where(x => x.Id == LogoId)
            .Select(x => string.IsNullOrWhiteSpace(x.Valor) ? null : x.Valor)
            .FirstOrDefault();

    public void SaveLogo(string dataUri) => _store.WriteRow("branding",
        new KeyValuePair<string, object?>[] { new("id", LogoId), new("valor", dataUri) });

    public void ClearLogo() => _store.WriteRow("branding",
        new KeyValuePair<string, object?>[] { new("id", LogoId), new("valor", "") }, deleted: true);
}
