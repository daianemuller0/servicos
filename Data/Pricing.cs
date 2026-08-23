using System.Globalization;
using HowdenServicos.Poc.Models;

namespace HowdenServicos.Poc.Data;

/// <summary>
/// Motor de cálculo da proposta de serviço — a tradução em C# das guias
/// CUSTO e PRICING da planilha.
///
/// Cadeia de preço (idêntica à planilha):
///   custo (MO + despesas)
///     → + risco                                  (PRICING E36/E38)
///     → ÷ (1 − provisões − margem alvo)          (PRICING E52, "montar preço a partir da margem")
///     = VENDA LÍQUIDA (valor sem impostos)       (PRICING Q18)
///     → ÷ (1 − PIS − COFINS − ISS)               (PRICING Q20)
///     = VALOR COM IMPOSTOS
///
/// O preço de cada item da proposta sai por participação no custo:
/// participação = custo do item ÷ custo total; valor do item = participação ×
/// valor com impostos (é o que a coluna L da guia CUSTO fazia).
/// </summary>
public static class Pricing
{
    // ---------- conversões tolerantes (aceita "1.234,56" e "1,234.56") ----------
    public static double Num(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return 0;
        s = s.Trim().Replace("R$", "").Replace("$", "").Replace("%", "").Replace(" ", "");
        if (s.Contains(',') && s.LastIndexOf(',') > s.LastIndexOf('.'))
            s = s.Replace(".", "").Replace(',', '.');
        else
            s = s.Replace(",", "");
        return double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v) ? v : 0;
    }

    /// <summary>Lê um campo digitado em porcentagem ("3" = 3%) como fração (0,03).</summary>
    public static double Pct(string? s) => Num(s) / 100.0;

    public static int Inteiro(string? s)
    {
        var v = (int)Math.Round(Num(s));
        return v < 0 ? 0 : v;
    }

    /// <summary>Arredondamento para cima usado pelas classes VBA (RoundUp).</summary>
    public static double ParaCima(double v) => Math.Ceiling(Math.Round(v, 6));

    public static string Moeda(double v) => v.ToString("#,##0.00", CultureInfo.GetCultureInfo("pt-BR"));
    public static string Moeda0(double v) => v.ToString("#,##0", CultureInfo.GetCultureInfo("pt-BR"));
    public static string Porcento(double fracao) =>
        (fracao * 100).ToString("#,##0.00", CultureInfo.GetCultureInfo("pt-BR")) + "%";

    // ---------- custo ----------
    public static double CustoMO(ItemMO i, int tecnicos) =>
        Num(i.QtdDiaria) * Num(i.CustoHora) * Num(i.Horas) * (i.PorTecnico ? Math.Max(tecnicos, 1) : 1);

    public static double CustoDespesa(ItemDespesa d, int tecnicos) =>
        Num(d.Qtd) * Num(d.CustoUnitario) * (d.PorTecnico ? Math.Max(tecnicos, 1) : 1);

    // ---------- resultado ----------
    public sealed record Resultado(
        double TotalMO, double TotalDespesas, double CustoTotal,
        double Risco, double CustoComRisco, double Fianca,
        double Comissoes, double MargemNegociacao, double PmSach, double GarantiaProjeto, double Portal,
        double VendaLiquida, double ComPisCofins, double ComImpostos,
        double Pis, double Cofins, double Iss,
        double FatorLiquido, double FatorPisCofins, double FatorComImpostos,
        double Markup, double ProjectMargin, double ContributionMargin);

    public static Resultado Calcular(IEnumerable<ItemMO> mo, IEnumerable<ItemDespesa> desp, PricingParams p)
    {
        var tec = Math.Max(Inteiro(p.QtdTecnicos), 1);
        var totalMO = mo.Sum(i => CustoMO(i, tec));
        var totalDesp = desp.Sum(d => CustoDespesa(d, tec));
        var custo = totalMO + totalDesp;

        var risco = custo * Pct(p.RiscoPct);
        var custoRisco = custo + risco;
        var fianca = Num(p.Fianca);

        var provisoes = Pct(p.ComissoesPct) + Pct(p.MargemNegociacaoPct) + Pct(p.PmSachPct)
                      + Pct(p.GarantiaProjetoPct) + Pct(p.PortalPct);
        var margem = Pct(p.MargemAlvoPct);

        // Venda líquida a partir da margem alvo (equivale ao Goal Seek do markup na planilha).
        var den = 1 - provisoes - margem;
        var vendaLiquida = den > 0.0001 ? (custoRisco + fianca) / den : 0;

        var pis = Pct(p.PisPct);
        var cofins = Pct(p.CofinsPct);
        var iss = Pct(p.IssPct);
        var denImp = 1 - pis - cofins - iss;
        var comImpostos = denImp > 0.0001 ? vendaLiquida / denImp : 0;
        var comPisCofins = comImpostos * (1 - iss);

        var comissoes = vendaLiquida * Pct(p.ComissoesPct);
        var mn = vendaLiquida * Pct(p.MargemNegociacaoPct);
        var pmSach = vendaLiquida * Pct(p.PmSachPct);
        var garantia = vendaLiquida * Pct(p.GarantiaProjetoPct);
        var portal = vendaLiquida * Pct(p.PortalPct);

        var pm = vendaLiquida > 0
            ? (vendaLiquida - (custoRisco + fianca + comissoes + mn + pmSach + garantia + portal)) / vendaLiquida
            : 0;
        var cm = vendaLiquida > 0
            ? (vendaLiquida - (custoRisco + fianca + mn + pmSach + garantia + portal)) / vendaLiquida
            : 0;
        var markup = custoRisco > 0 ? (vendaLiquida * (1 - provisoes) - fianca) / custoRisco : 0;

        return new Resultado(
            totalMO, totalDesp, custo,
            risco, custoRisco, fianca,
            comissoes, mn, pmSach, garantia, portal,
            vendaLiquida, comPisCofins, comImpostos,
            comImpostos * pis, comImpostos * cofins, comImpostos * iss,
            custo > 0 ? vendaLiquida / custo : 0,
            custo > 0 ? comPisCofins / custo : 0,
            custo > 0 ? comImpostos / custo : 0,
            markup, pm, cm);
    }

    // ---------- itens já precificados (o que sai impresso na proposta) ----------
    public sealed record LinhaMO(string Servico, string Obs, double Horas, double QtdDiaria,
        double Custo, double Participacao, double ValorTotal, double ValorDiaria, double ValorHora);

    public sealed record LinhaDespesa(string Despesa, string Obs, double Qtd,
        double Custo, double Participacao, double ValorUnitario, double ValorTotal);

    public sealed record Complementar(string Descricao, string Obs, double Qtd, double Valor);

    public sealed record Documento(
        List<LinhaMO> MO, List<LinhaDespesa> Despesas, List<Complementar> Complementares,
        double TotalMO, double TotalDespesas, double Total, Resultado Calculo);

    /// <summary>
    /// Distribui o preço de venda entre os itens (proporcional ao custo) e monta
    /// as linhas da proposta, com os mesmos arredondamentos das classes VBA:
    /// valor total do item para cima em reais inteiros; diária = total ÷ qtd.
    /// </summary>
    public static Documento Montar(List<ItemMO> mo, List<ItemDespesa> desp, PricingParams p)
    {
        var calc = Calcular(mo, desp, p);
        var tec = Math.Max(Inteiro(p.QtdTecnicos), 1);
        var custoTotal = calc.CustoTotal;

        var linhasMO = new List<LinhaMO>();
        foreach (var i in mo)
        {
            var custo = CustoMO(i, tec);
            if (custo <= 0) continue;
            var part = custoTotal > 0 ? custo / custoTotal : 0;
            var total = ParaCima(part * calc.ComImpostos);
            var qtdDiaria = Num(i.QtdDiaria);
            var horas = Num(i.Horas);
            var diaria = qtdDiaria > 0 ? Math.Round(total / qtdDiaria, 2) : 0;
            var hora = horas > 0 ? Math.Round(diaria / horas, 2) : 0;
            linhasMO.Add(new LinhaMO(i.Servico, i.Obs, horas, qtdDiaria, custo, part, total, diaria, hora));
        }

        var linhasDesp = new List<LinhaDespesa>();
        foreach (var d in desp)
        {
            var custo = CustoDespesa(d, tec);
            if (custo <= 0) continue;
            var part = custoTotal > 0 ? custo / custoTotal : 0;
            var total = ParaCima(part * calc.ComImpostos);
            var qtd = Num(d.Qtd) * (d.PorTecnico ? tec : 1);
            var unit = qtd > 0 ? ParaCima(total / qtd) : 0;
            linhasDesp.Add(new LinhaDespesa(d.Despesa, d.Obs, qtd, custo, part, unit, total));
        }

        var totalMO = linhasMO.Sum(l => l.ValorTotal);
        var totalDesp = linhasDesp.Sum(l => l.ValorTotal);

        return new Documento(linhasMO, linhasDesp, Complementares(linhasMO, linhasDesp),
            totalMO, totalDesp, totalMO + totalDesp, calc);
    }

    /// <summary>
    /// "INFORMAÇÕES COMPLEMENTARES — NÃO INCLUSO": diária adicional e horas
    /// extras, na mesma regra da ClasseInfoComplementares do VBA.
    /// </summary>
    public static List<Complementar> Complementares(List<LinhaMO> mo, List<LinhaDespesa> desp)
    {
        // Diária adicional = diárias normais + despesas que se repetem por dia.
        var diaria = mo.Where(l => l.Servico.StartsWith("DIARIAS NORMAIS", StringComparison.OrdinalIgnoreCase))
                       .Sum(l => l.ValorDiaria);
        string[] porDia = { "HOSPEDAGEM", "LOCAÇÃO DE CARRO", "COMBUSTIVEL", "COMBUSTÍVEL", "REFEIÇÕES" };
        diaria += desp.Where(d => porDia.Contains(d.Despesa.Trim().ToUpperInvariant()))
                      .Sum(d => d.ValorUnitario);

        // Hora normal = diária do 1º turno ÷ 8.
        var primeiro = mo.FirstOrDefault(l =>
            l.Servico.StartsWith("DIARIAS NORMAIS", StringComparison.OrdinalIgnoreCase) &&
            l.Obs.Contains("1o", StringComparison.OrdinalIgnoreCase));
        var valorHora = primeiro is null ? 0 : primeiro.ValorDiaria / 8;

        return new List<Complementar>
        {
            new("DIARIA ADICIONAL", "DAS 8:00 AS 17:00 (DIAS UTEIS DE SEG. A SEX.)", 1, ParaCima(diaria)),
            new("HORA EXTRA", "SUPERIOR A 8/H DIA (DIAS UTEIS DE SEG. A SEX.)", 1, ParaCima(valorHora * 1.5)),
            new("HORA EXTRA", "SAB., DOM. E FERIADOS", 1, ParaCima(valorHora * 2)),
        };
    }
}
