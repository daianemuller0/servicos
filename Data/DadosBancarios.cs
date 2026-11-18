namespace HowdenServicos.Poc.Data;

/// <summary>Uma linha do bloco bancário: coluna da esquerda e da direita (vazia = a da esquerda ocupa a linha).</summary>
public sealed record LinhaBancaria(string Esquerda, string Direita);

/// <summary>
/// Os dados bancários que saem em "Dados para faturamento", por BU e pela
/// MOEDA da proposta. A razão social e o endereço da BU (tela Dados de
/// Faturamento) saem em todas as moedas; embaixo vem o bloco da moeda.
///
/// <para>
/// Cada bloco é um texto, uma linha por linha do quadro: "esquerda | direita"
/// (sem "|", a linha ocupa a largura toda) e *assim* sai em negrito. O padrão
/// de fábrica é o das contas da equipe; editado na tela Dados de Faturamento,
/// fica salvo na configuração (chave "banco.&lt;BU&gt;.&lt;moeda&gt;").
/// Sem bloco para a moeda, a proposta usa o banco do cadastro da BU.
/// </para>
/// </summary>
public static class DadosBancarios
{
    public static string Chave(string bu, string moeda) => $"banco.{bu}.{moeda}";

    private const string ItauEndereco = "Bank's Address: RUA URURAI 111, PREDIO B, FLOOR 2, SAO PAULO, Brazil";
    private const string HsaCliente = "Beneficiary Customer: HOWDEN SOUTH AMERICA VENTILADORES E COMPRESSORES IND E";
    private const string HsaConta = "Branch number / Account number: 4892 / 00284805";
    private const string HsaIban = "IBAN: BR7460701190048920000284805C1";

    private static readonly Dictionary<(string Bu, string Moeda), string> Fabrica = new()
    {
        [("HSA-SP", "BRL")] = "*Banco: Banco Itaú* – Agência: 4892 Conta: 28480-5",
        [("HSA-SP", "USD")] = string.Join("\n",
            "Intermediary's Bank: | Beneficiary's Bank:",
            "*JPMorgan Chase Bank, N.A* | *In favor of Itaú Unibanco S.A.*",
            "Swift Code: CHASUS33 | Swift code: ITAUBRSP",
            $"Account nr: 544705690 | {ItauEndereco}",
            $"ABA: 021000021 | {HsaCliente}",
            $" | {HsaConta}",
            $" | {HsaIban}"),
        [("HSA-SP", "EUR")] = string.Join("\n",
            "Intermediary's Bank: | Beneficiary's Bank:",
            "Swift code: SCBLDEFX | Account with: *Commerzbank Ag*",
            "Account nr: 18250407 | Swift Code: COBADEFF",
            "In favor of Itaú Unibanco S.A. | Beneficiary Bank Itaú Unibanco S.A.",
            "Swift code: ITAUBRSP | Swift Code: ITAUBRSP",
            $" | {ItauEndereco}",
            $" | {HsaCliente}",
            $" | {HsaConta}",
            $" | {HsaIban}"),

        [("HCHL", "CLP")] = string.Join("\n",
            "Account Holder Name | HOWDEN CHILE SPA",
            "Account Holder Address | CORDILLERA 575, PUDAHUEL, SANTIAGO - CHILE",
            "Current Account Number | 000-059261-001",
            "TAX ID | 76.274.455-4",
            "Beneficiary Bank | HSBC Bank (CHILE)",
            "Swift Code | BLICCLRM",
            "Beneficiary Bank Address | Isidora Goyenechea 2800 – 23rd floor, Las Condes",
            "Beneficiary Bank City/Country | Santiago, Chile"),
        [("HCHL", "USD")] = string.Join("\n",
            "Account Holder Name | HOWDEN CHILE SPA",
            "Account Holder Address | CORDILLERA 575, PUDAHUEL, SANTIAGO - CHILE",
            "Current Account Number | 000-059261-060",
            "Beneficiary Bank Name | HSBC Bank (Chile)",
            "Beneficiary Swift Code | BLICCLRM",
            "Beneficiary Bank Address | Isidora Goyenechea 2800 – 23rd floor, Las Condes",
            "Beneficiary Bank City/Country | Santiago, Chile",
            "Correspondent Bank Name | HSBC Bank USA",
            "Correspondent Bank Swift Code | MRMDUS33",
            "Correspondent Bank Account Number | 000056634",
            "Correspondent Bank Address | 452 FIFTH AVENUE",
            "Correspondent Bank City/Country | New York/USA"),
        [("HCHL", "EUR")] = string.Join("\n",
            "Account Holder Name | HOWDEN CHILE SPA",
            "Account Holder Address | CORDILLERA 575, PUDAHUEL, SANTIAGO - CHILE",
            "Current Account Number | 000-059261-061",
            "Beneficiary Bank Swift Code | BLICCLRM",
            "Beneficiary Bank Address | Isidora Goyenechea 2800 – 23rd floor, Las Condes",
            "Beneficiary Bank City/Country | Santiago, Chile",
            "Correspondent Bank Name | HSBC France",
            "Correspondent Bank Swift Code | CCFRFRPP",
            "Correspondent Bank Account Number - IBAN | FR7630056000100010000407283",
            "Correspondent Bank Address | 103 Avenue des Champs-Élysées",
            "Correspondent Bank City/Country | Paris/France"),
    };

    /// <summary>O padrão de fábrica do bloco (vazio quando não há).</summary>
    public static string DeFabrica(string bu, string moeda) =>
        Fabrica.TryGetValue((bu, moeda), out var t) ? t : "";

    /// <summary>O bloco da BU e da moeda: o salvo na configuração ou, sem nada salvo, o de fábrica.</summary>
    public static string Texto(IReadOnlyDictionary<string, string>? cfg, string bu, string moeda)
    {
        if (cfg is not null && cfg.TryGetValue(Chave(bu, moeda), out var salvo))
            return salvo;   // salvo vazio = sem dados bancários nessa moeda (vale o banco do cadastro)
        return DeFabrica(bu, moeda);
    }

    /// <summary>As linhas do bloco ("esquerda | direita").</summary>
    public static List<LinhaBancaria> Linhas(string? texto) => (texto ?? "")
        .Replace("\r\n", "\n").Split('\n')
        .Where(l => l.Trim().Length > 0)
        .Select(l =>
        {
            var i = l.IndexOf('|');
            return i < 0 ? new LinhaBancaria(l.Trim(), "") : new LinhaBancaria(l[..i].Trim(), l[(i + 1)..].Trim());
        })
        .ToList();

    /// <summary>
    /// A largura da coluna da esquerda (em %), pelo texto mais longo de cada
    /// coluna: entre 30% e 50%, para a da direita (endereços, IBAN) não quebrar à toa.
    /// </summary>
    public static int LarguraDaEsquerda(List<LinhaBancaria> linhas)
    {
        int Maior(Func<LinhaBancaria, string> col) => linhas.Select(l => col(l).Replace("*", "").Length).DefaultIfEmpty(0).Max();
        var e = Maior(l => l.Esquerda);
        var d = Maior(l => l.Direita);
        if (e + d == 0) return 50;
        return Math.Clamp((int)Math.Round(100.0 * e / (e + d)) + 4, 30, 50);
    }

    /// <summary>Os trechos de uma célula: (texto, negrito) — "*assim*" é negrito.</summary>
    public static List<(string Texto, bool Negrito)> Trechos(string celula)
    {
        var saida = new List<(string, bool)>();
        var partes = (celula ?? "").Split('*');
        for (var i = 0; i < partes.Length; i++)
            if (partes[i].Length > 0) saida.Add((partes[i], i % 2 == 1));
        return saida;
    }

    /// <summary>
    /// A célula no idioma da proposta, pelas regras da proposta completa
    /// ("Banco: X – Agência: Y Conta: Z" → "Bank: X – Branch: Y Account: Z"),
    /// mantendo o negrito. O que não tem tradução sai como está (os dados em
    /// inglês das contas no exterior já servem nos três idiomas).
    /// </summary>
    public static string NoIdioma(string celula, string idioma)
    {
        var lingua = IdiomasDaProposta.Normalizar(idioma);
        if (lingua == IdiomasDaProposta.Portugues || celula.Length == 0) return celula;
        var trechos = Trechos(celula);
        var plano = string.Concat(trechos.Select(t => t.Texto));
        if (TraducaoDaProposta.Traduzir(plano, lingua) is not { } traducao || traducao == plano.Trim()) return celula;
        var partes = traducao.Split(TraducaoDaProposta.Parte);
        if (partes.Length == trechos.Count)
            return string.Concat(partes.Select((p, i) => trechos[i].Negrito ? $"*{p}*" : p));
        return trechos.Count == 1 && trechos[0].Negrito ? $"*{traducao.Replace(TraducaoDaProposta.Parte.ToString(), "")}*"
            : traducao.Replace(TraducaoDaProposta.Parte.ToString(), "");
    }

    /// <summary>O título do bloco: "Real – BRL", "Dólar Americano – USD"... no idioma da proposta.</summary>
    public static string Titulo(string moeda, string idioma)
    {
        var en = IdiomasDaProposta.Normalizar(idioma) == IdiomasDaProposta.Ingles;
        var es = IdiomasDaProposta.Normalizar(idioma) == IdiomasDaProposta.Espanhol;
        var nome = moeda switch
        {
            "BRL" => en ? "Brazilian Real" : es ? "Real Brasileño" : "Real",
            "USD" => en ? "US Dollar" : es ? "Dólar Estadounidense" : "Dólar Americano",
            "EUR" => "Euro",
            "CLP" => en ? "Chilean Peso" : "Peso Chileno",
            "PEN" => en ? "Peruvian Sol" : "Sol Peruano",
            _ => moeda,
        };
        return $"{nome} – {moeda}";
    }
}
