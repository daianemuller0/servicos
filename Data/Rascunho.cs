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

    /// <summary>
    /// Aplica a hora-base aos serviços com multiplicador (planilha: E9=E8×1,5;
    /// E10=E8×2…). Itens com multiplicador zero não são tocados, e qualquer
    /// valor pode ser editado depois, item a item, só nesta proposta.
    /// </summary>
    public void AplicarHoraBase()
    {
        var baseHora = Pricing.Num(Params.HoraBase);
        if (baseHora <= 0) return;
        foreach (var i in ItensMO)
        {
            var m = Pricing.Num(i.Mult);
            if (m > 0) i.CustoHora = (baseHora * m).ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
        }
    }

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
            Mult = string.IsNullOrWhiteSpace(p.Mult) ? "0" : p.Mult,
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
        MigrarMultiplicadores();
    }

    /// <summary>
    /// Rascunhos e propostas gravados antes do multiplicador da hora-base não
    /// carregam essa marcação — sem ela, as regras de diária (sáb/dom = 2×,
    /// HE = 1,5×/2×) e a ferramenta de proposta anterior não reconhecem as
    /// linhas. Re-identifica pelo nome do serviço.
    /// </summary>
    private void MigrarMultiplicadores()
    {
        foreach (var i in ItensMO)
        {
            if (Pricing.Num(i.Mult) > 0) continue;
            var servico = (i.Servico ?? "").ToUpperInvariant();
            var obs = (i.Obs ?? "").ToUpperInvariant();
            var fimDeSemana = servico.Contains("SAB") || servico.Contains("DOM") || servico.Contains("FER");

            if (servico.StartsWith("DIARIAS NORMAIS"))
                i.Mult = obs.Contains("2O") || obs.Contains("2°") ? "1.5" : "1";
            else if (servico.Contains("DIARIAS EXTRAS") && fimDeSemana)
                i.Mult = "2";
            else if (servico.Contains("HORAS EXTRAS"))
                i.Mult = fimDeSemana ? "2" : "1.5";
        }
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
        MigrarMultiplicadores();
        return true;
    }

    private static T? Des<T>(string json)
    {
        if (string.IsNullOrWhiteSpace(json)) return default;
        try { return JsonSerializer.Deserialize<T>(json); } catch { return default; }
    }
}
