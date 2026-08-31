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

    // ---------- seguro garantia / carta de fiança (listas!Custo_Garantia) ----------
    /// <summary>Instrumentos de garantia e taxa anual (% a.a.) da planilha.</summary>
    public static readonly (string Nome, double TaxaAnualPct)[] GarantiaTipos =
    {
        ("Não", 0),
        ("HSA C. Fiança (Banco)", 5.0),
        ("HSA S. Garantia (Corretora)", 1.0),
        ("HCHL/HPU C. Fiança (Banco)", 3.75),
        ("HCHL/HPU S. Garantia (Corretora)", 1.75),
    };

    /// <summary>"OUTROS" é a linha de acerto de valores: nunca sai na proposta ao cliente.</summary>
    public static bool EhOutros(string? nome) =>
        (nome ?? "").Trim().Equals("OUTROS", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Despesas "por dia de trabalho": a quantidade acompanha a soma das
    /// diárias lançadas (hospedagem, carro, combustível e refeições).
    /// </summary>
    public static bool EhDespesaDiaria(string? nome)
    {
        var n = (nome ?? "").ToUpperInvariant();
        return n.Contains("HOSPEDAGEM") || n.Contains("LOCA") || n.Contains("COMBUST") || n.Contains("REFEI");
    }

    /// <summary>
    /// Despesas de deslocamento (táxi + passagem aérea): no modo sem despesas,
    /// saem numa linha própria abaixo do total, a custo + taxa administrativa.
    /// </summary>
    public static bool EhDeslocamento(string? nome)
    {
        var n = (nome ?? "").ToUpperInvariant();
        return n.Contains("TAXI") || n.Contains("TÁXI") || n.Contains("PASSAGEM");
    }

    /// <summary>
    /// Margem (Project Margin) que resulta se o valor final com impostos for
    /// exatamente a meta informada — a função "chegar no valor".
    /// </summary>
    public static double MargemParaMeta(Resultado c, PricingParams p, double meta, double prazoDias = 0)
    {
        if (meta <= 0 || c.CustoComRisco <= 0) return 0;
        var denImp = 1 - Pct(p.PisPct) - Pct(p.CofinsPct) - Pct(p.IssPct);
        var vendaLiquida = meta * denImp;
        if (vendaLiquida <= 0) return 0;
        var provisoes = ComissoesFrac(p) + Pct(p.MargemNegociacaoPct) + Pct(p.PmSachPct)
                      + Pct(p.GarantiaProjetoPct) + Pct(p.PortalPct);
        return 1 - provisoes - FiancaFrac(p, prazoDias) - c.CustoComRisco / vendaLiquida;
    }

    public static double TaxaGarantia(string tipo) =>
        GarantiaTipos.FirstOrDefault(g => g.Nome == tipo).TaxaAnualPct / 100.0;

    /// <summary>
    /// Comissão total (fração), como a tabela de comissões da planilha:
    /// PPR + Sales Director + Sales Industrial + DSR×(comissões de Sales) + Rep1 + Rep2.
    /// </summary>
    public static double ComissoesFrac(PricingParams p)
    {
        var sales = Pct(p.SalesDirPct) + Pct(p.SalesIndPct);
        return Pct(p.PprPct) + sales + Pct(p.DsrFatorPct) * sales + Pct(p.Rep1Pct) + Pct(p.Rep2Pct);
    }

    /// <summary>
    /// Fração da venda consumida pela fiança: % coberto × dias × taxa anual ÷ 365.
    /// Dias em branco (ou zero) = acompanha o PRAZO DE ENTREGA, como na planilha
    /// (cobertura R49 = prazo − evento de pagamento).
    /// </summary>
    public static double FiancaFrac(PricingParams p, double prazoDias = 0)
    {
        var taxa = TaxaGarantia(p.FiancaTipo);
        if (taxa <= 0) return 0;
        var dias = Num(p.FiancaDias);
        if (dias <= 0) dias = prazoDias;
        return Pct(p.FiancaPctVenda) * dias * taxa / 365.0;
    }

    // ---------- resultado ----------
    public sealed record Resultado(
        double TotalMO, double TotalDespesas, double CustoTotal,
        double Risco, double CustoComRisco, double Fianca,
        double Comissoes, double MargemNegociacao, double PmSach, double GarantiaProjeto, double Portal,
        double VendaLiquida, double ComPisCofins, double ComImpostos,
        double Pis, double Cofins, double Iss,
        double FatorLiquido, double FatorPisCofins, double FatorComImpostos,
        double Markup, double ProjectMargin, double ContributionMargin,
        double ComissoesFracTotal, double Dsr);

    public static Resultado Calcular(IEnumerable<ItemMO> mo, IEnumerable<ItemDespesa> desp, PricingParams p, double prazoDias = 0)
    {
        var tec = Math.Max(Inteiro(p.QtdTecnicos), 1);
        var totalMO = mo.Sum(i => CustoMO(i, tec));
        var totalDesp = desp.Sum(d => CustoDespesa(d, tec));
        var custo = totalMO + totalDesp;

        var risco = custo * Pct(p.RiscoPct);
        var custoRisco = custo + risco;

        var comissoesFrac = ComissoesFrac(p);
        var provisoes = comissoesFrac + Pct(p.MargemNegociacaoPct) + Pct(p.PmSachPct)
                      + Pct(p.GarantiaProjetoPct) + Pct(p.PortalPct);
        var margem = Pct(p.MargemAlvoPct);

        // Venda líquida a partir da margem alvo. A fiança é proporcional à própria
        // venda (fração k), então entra no denominador — mesma conta que a planilha
        // fecha por iteração: venda×(1−prov−margem) = custoRisco + venda×k.
        var k = FiancaFrac(p, prazoDias);
        var den = 1 - provisoes - margem - k;
        var vendaLiquida = den > 0.0001 ? custoRisco / den : 0;
        var fianca = vendaLiquida * k;

        var pis = Pct(p.PisPct);
        var cofins = Pct(p.CofinsPct);
        var iss = Pct(p.IssPct);
        var denImp = 1 - pis - cofins - iss;
        var comImpostos = denImp > 0.0001 ? vendaLiquida / denImp : 0;
        var comPisCofins = comImpostos * (1 - iss);

        var salesFrac = Pct(p.SalesDirPct) + Pct(p.SalesIndPct);
        var comissoes = vendaLiquida * comissoesFrac;
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
            markup, pm, cm,
            comissoesFrac, vendaLiquida * Pct(p.DsrFatorPct) * salesFrac);
    }

    // ---------- itens já precificados (o que sai impresso na proposta) ----------
    public sealed record LinhaMO(string Servico, string Obs, double Horas, double QtdDiaria,
        double Custo, double Participacao, double ValorTotal, double ValorDiaria, double ValorHora,
        double Mult = 0);

    public sealed record LinhaDespesa(string Despesa, string Obs, double Qtd,
        double Custo, double Participacao, double ValorUnitario, double ValorTotal);

    public sealed record Complementar(string Descricao, string Obs, double Qtd, double Valor);

    public sealed record Documento(
        List<LinhaMO> MO, List<LinhaDespesa> Despesas, List<Complementar> Complementares,
        double TotalMO, double TotalDespesas, double Total, Resultado Calculo,
        double Deslocamento = 0);

    /// <summary>
    /// Distribui o preço de venda entre os itens (proporcional ao custo) e monta
    /// as linhas da proposta, com os mesmos arredondamentos das classes VBA:
    /// valor total do item para cima em reais inteiros; diária = total ÷ qtd.
    /// </summary>
    public static Documento Montar(List<ItemMO> mo, List<ItemDespesa> desp, PricingParams p, double prazoDias = 0)
    {
        var calc = Calcular(mo, desp, p, prazoDias);
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
            linhasMO.Add(new LinhaMO(i.Servico, i.Obs, horas, qtdDiaria, custo, part, total, diaria, hora, Num(i.Mult)));
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
    /// Versão do documento APRESENTADA ao cliente. O pricing interno não muda —
    /// só a distribuição entre as tabelas:
    ///
    ///  - "ComDespesas": as despesas saem a custo + taxa administrativa (ex.: 30%);
    ///    a diferença até o preço real delas é embutida nas diárias da assessoria.
    ///  - "SemDespesas": a tabela de despesas some e todo o valor vai para as diárias.
    ///
    /// O TOTAL C/ IMPOSTOS é exatamente o mesmo nas duas formas.
    /// </summary>
    public static Documento Apresentar(Documento doc, string modo, double taxaAdmPct, double diariaTravada = 0)
    {
        // Sem linhas de assessoria não há onde embutir — mantém como está.
        if (doc.MO.Count == 0) return doc;

        List<LinhaDespesa> desp;
        double deslocamento = 0;
        var admFrac = 1 + taxaAdmPct / 100.0;
        if (modo == "SemDespesas")
        {
            desp = new List<LinhaDespesa>();
            // Táxi + passagem aérea saem numa linha própria, a custo + taxa adm;
            // o restante do valor (e as demais despesas) sobe para as diárias.
            deslocamento = doc.Despesas
                .Where(d => EhDeslocamento(d.Despesa))
                .Sum(d => ParaCima(d.Custo * admFrac));
        }
        else
        {
            var adm = 1 + taxaAdmPct / 100.0;
            // "OUTROS" nunca sai na proposta — o valor dele é diluído nas diárias.
            desp = doc.Despesas.Where(d => !EhOutros(d.Despesa)).Select(d =>
            {
                var custoUnit = d.Qtd > 0 ? d.Custo / d.Qtd : d.Custo;
                var unit = ParaCima(custoUnit * adm);
                var total = unit * Math.Max(d.Qtd, 1);
                return new LinhaDespesa(d.Despesa, d.Obs, d.Qtd, d.Custo, d.Participacao, unit, total);
            }).ToList();
        }

        // O que as despesas deixam de mostrar vai para a assessoria.
        var alvoMO = doc.Total - desp.Sum(d => d.ValorTotal) - deslocamento;
        if (alvoMO <= 0 || doc.TotalMO <= 0) return doc;

        // Excedente do DESLOCAMENTO (valor real do táxi + passagem além do que é
        // mostrado a custo + taxa adm): sobe SÓ para as diárias normais — as
        // linhas derivadas (2º turno, sáb/dom, HE) não carregam deslocamento,
        // porque o técnico já está no local.
        var deslocInterno = doc.Despesas.Where(d => EhDeslocamento(d.Despesa)).Sum(d => d.ValorTotal);
        var deslocMostrado = modo == "SemDespesas"
            ? deslocamento
            : desp.Where(d => EhDeslocamento(d.Despesa)).Sum(d => d.ValorTotal);
        var excedenteDesloc = deslocInterno - deslocMostrado;

        // As linhas com multiplicador obedecem às regras fixas sobre a hora normal
        // (2º turno/HE semana = 1,5×; sáb/dom/feriado = 2×), todas com impostos.
        // As linhas sem multiplicador (equipamentos, terceiros…) mantêm o valor
        // distribuído por participação no custo.
        var comMult = doc.MO.Where(l => l.Mult > 0 && l.QtdDiaria > 0).ToList();
        var pesoW = comMult.Sum(l => l.Mult * Math.Max(l.Horas, 1) * l.QtdDiaria);
        var livres = doc.MO.Where(l => !(l.Mult > 0 && l.QtdDiaria > 0)).Sum(l => l.ValorTotal);
        var pool = alvoMO - livres;

        var mo = new List<LinhaMO>();
        if (pesoW > 0 && pool > 0)
        {
            // As diárias normais (mult = 1) absorvem o excedente do deslocamento;
            // a base limpa é distribuída pelos pesos entre todas as linhas.
            var normais = doc.MO.Where(l => l.Mult is > 0.99 and < 1.01 && l.QtdDiaria > 0).ToList();
            var qtdNormais = normais.Sum(l => l.QtdDiaria);
            var extra = normais.Count > 0 ? excedenteDesloc : 0;   // sem diária normal, fica no pool
            var poolBase = pool - extra;
            if (poolBase <= 0) { poolBase = pool; extra = 0; }

            // hora normal (base, sem deslocamento): pool ÷ pesos, arredondado p/ cima
            var horaNormal = Math.Ceiling(poolBase / pesoW * 100) / 100;
            foreach (var l in doc.MO)
            {
                if (l.Mult > 0 && l.QtdDiaria > 0)
                {
                    var hora = Math.Round(l.Mult * horaNormal, 2);
                    var diaria = Math.Round(hora * Math.Max(l.Horas, 1), 2);
                    if (extra != 0 && qtdNormais > 0 && l.Mult is > 0.99 and < 1.01)
                        diaria = Math.Round(diaria + extra / qtdNormais, 2);
                    // Diária cravada pela ferramenta de proposta anterior: aparar o
                    // resíduo de arredondamento (no máx. R$ 1) para o valor exato.
                    if (diariaTravada > 0 && l.Mult is > 0.99 and < 1.01 &&
                        Math.Abs(diaria - diariaTravada) <= 1.0)
                        diaria = diariaTravada;
                    var total = Math.Round(diaria * l.QtdDiaria, 2);
                    var horaLinha = Math.Round(diaria / Math.Max(l.Horas, 1), 2);
                    mo.Add(new LinhaMO(l.Servico, l.Obs, l.Horas, l.QtdDiaria, l.Custo, l.Participacao, total, diaria, horaLinha, l.Mult));
                }
                else
                {
                    mo.Add(l);
                }
            }
        }
        else
        {
            // fallback: distribuição proporcional ao custo (nenhuma linha com multiplicador)
            double acumulado = 0;
            for (var i = 0; i < doc.MO.Count; i++)
            {
                var l = doc.MO[i];
                double total = i == doc.MO.Count - 1
                    ? alvoMO - acumulado
                    : ParaCima(alvoMO * (l.ValorTotal / doc.TotalMO));
                acumulado += total;
                var diaria = l.QtdDiaria > 0 ? Math.Round(total / l.QtdDiaria, 2) : 0;
                var hora = l.Horas > 0 ? Math.Round(diaria / l.Horas, 2) : 0;
                mo.Add(new LinhaMO(l.Servico, l.Obs, l.Horas, l.QtdDiaria, l.Custo, l.Participacao, total, diaria, hora, l.Mult));
            }
        }

        var totalMO = mo.Sum(l => l.ValorTotal);
        var totalDesp = desp.Sum(d => d.ValorTotal);
        return new Documento(mo, desp, Complementares(mo, desp), totalMO, totalDesp,
            totalMO + totalDesp + deslocamento, doc.Calculo, deslocamento);
    }

    /// <summary>Diária normal (1×, com impostos) como sai na proposta apresentada.</summary>
    public static double DiariaNormalApresentada(Documento apresentado)
    {
        var normal = apresentado.MO.FirstOrDefault(l => l.Mult is > 0.99 and < 1.01 && l.QtdDiaria > 0);
        if (normal is not null) return normal.ValorDiaria;
        var outra = apresentado.MO.FirstOrDefault(l => l.Mult > 0 && l.QtdDiaria > 0);
        return outra is null ? 0 : Math.Round(outra.ValorDiaria / outra.Mult, 2);
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
