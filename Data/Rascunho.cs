using System.Text.Json;
using HowdenServicos.Poc.Models;

namespace HowdenServicos.Poc.Data;

/// <summary>
/// A proposta que está sendo montada agora. É um serviço "scoped": no Blazor
/// Server isso equivale ao circuito do usuário, então o rascunho sobrevive à
/// navegação entre Custo → Pricing → Proposta sem gravar nada no banco.
/// </summary>
public class Rascunho
{
    public Proposta Proposta { get; private set; } = new();
    public List<ItemMO> ItensMO { get; set; } = new();
    public List<ItemDespesa> ItensDespesa { get; set; } = new();
    public PricingParams Params { get; private set; } = new();

    /// <summary>Verdadeiro depois que a proposta foi gravada no banco.</summary>
    public bool Gravada => !string.IsNullOrWhiteSpace(Proposta.Id);

    /// <summary>Rascunho ainda não iniciado (nenhum item carregado).</summary>
    public bool Vazio => ItensMO.Count == 0 && ItensDespesa.Count == 0;

    public Pricing.Documento Documento() => Pricing.Montar(ItensMO, ItensDespesa, Params);
    public Pricing.Resultado Calculo() => Pricing.Calcular(ItensMO, ItensDespesa, Params);

    /// <summary>Começa uma proposta nova, já com os itens padrão da tabela de custos.</summary>
    public void Novo(List<Parametro> parametros, string usuario)
    {
        Proposta = new Proposta
        {
            Data = DateTime.Today.ToString("yyyy-MM-dd"),
            Ano = DateTime.Today.Year.ToString(),
            PreparadaPor = usuario,
        };
        Params = new PricingParams();
        ItensMO = parametros.Where(p => p.Tipo == "MO").Select(p => new ItemMO
        {
            Servico = p.Descricao, Obs = p.Obs, Horas = p.Horas,
            CustoHora = p.Valor, QtdDiaria = "0", PorTecnico = p.PorTecnico == "Sim",
        }).ToList();
        ItensDespesa = parametros.Where(p => p.Tipo == "DESPESA").Select(p => new ItemDespesa
        {
            Despesa = p.Descricao, Obs = p.Obs, Qtd = "0",
            CustoUnitario = p.Valor, PorTecnico = p.PorTecnico == "Sim",
        }).ToList();
    }

    /// <summary>Carrega uma proposta gravada de volta para edição.</summary>
    public void Carregar(Proposta p)
    {
        Proposta = p;
        ItensMO = Des<List<ItemMO>>(p.ItensMoJson) ?? new();
        ItensDespesa = Des<List<ItemDespesa>>(p.ItensDespesaJson) ?? new();
        Params = Des<PricingParams>(p.PricingJson) ?? new PricingParams();
    }

    /// <summary>Copia o estado atual para o objeto que vai ser gravado.</summary>
    public Proposta ParaGravar()
    {
        var doc = Documento();
        Proposta.ItensMoJson = JsonSerializer.Serialize(ItensMO);
        Proposta.ItensDespesaJson = JsonSerializer.Serialize(ItensDespesa);
        Proposta.PricingJson = JsonSerializer.Serialize(Params);
        Proposta.CustoTotal = doc.Calculo.CustoTotal.ToString(System.Globalization.CultureInfo.InvariantCulture);
        Proposta.Total = doc.Total.ToString(System.Globalization.CultureInfo.InvariantCulture);
        return Proposta;
    }

    // ---- persistência no navegador (sobrevive a F5 e a abrir a página direto) ----
    private sealed record Estado(Proposta Proposta, List<ItemMO> MO, List<ItemDespesa> Despesas, PricingParams Params);

    public string ToJson() =>
        JsonSerializer.Serialize(new Estado(Proposta, ItensMO, ItensDespesa, Params));

    public bool FromJson(string? json)
    {
        var e = Des<Estado>(json ?? "");
        if (e is null || e.Proposta is null) return false;
        Proposta = e.Proposta;
        ItensMO = e.MO ?? new();
        ItensDespesa = e.Despesas ?? new();
        Params = e.Params ?? new PricingParams();
        return true;
    }

    private static T? Des<T>(string json)
    {
        if (string.IsNullOrWhiteSpace(json)) return default;
        try { return JsonSerializer.Deserialize<T>(json); } catch { return default; }
    }
}
