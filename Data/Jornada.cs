using System.Text.Json;

namespace HowdenServicos.Poc.Data;

/// <summary>
/// O parágrafo do horário de trabalho ("Proposta considera a mobilização e
/// desmobilização durante semana ... das 08h00 às 17h00 incluindo 01 hora de
/// descanso"). A proposta guarda o TEXTO escolhido (em branco = o padrão de
/// fábrica); os padrões ficam na configuração (chave <see cref="Chave"/>) —
/// o que alguém salva como padrão aparece para todos.
/// </summary>
public static class Jornada
{
    public const string Chave = "padroes.jornada";

    public const string DeFabrica =
        "Proposta considera a mobilização e desmobilização durante semana (Segunda a Sexta, exceto feriados) no período diurno das 08h00 às 17h00 incluindo 01 hora de descanso";

    private const string DeFabricaEn =
        "This proposal considers mobilization and demobilization during weekdays (Monday to Friday, except holidays) in the daytime period from 8:00 a.m. to 5:00 p.m., including a 1-hour break";

    private const string DeFabricaEs =
        "La propuesta considera la movilización y desmovilización durante la semana (lunes a viernes, excepto feriados) en el período diurno de 08h00 a 17h00, incluida 01 hora de descanso";

    /// <summary>Os padrões: o de fábrica sempre primeiro, depois os salvos.</summary>
    public static List<string> Padroes(IReadOnlyDictionary<string, string>? cfg)
    {
        var lista = new List<string> { DeFabrica };
        if (cfg is not null && cfg.TryGetValue(Chave, out var json) && !string.IsNullOrWhiteSpace(json))
        {
            try
            {
                foreach (var t in JsonSerializer.Deserialize<List<string>>(json) ?? new())
                    if (Limpar(t) is { Length: > 0 } l && !lista.Any(x => Igual(x, l))) lista.Add(l);
            }
            catch (JsonException) { }
        }
        return lista;
    }

    /// <summary>Os padrões salvos (sem o de fábrica), para gravar na configuração.</summary>
    public static string Gravar(IEnumerable<string> padroes) =>
        JsonSerializer.Serialize(padroes.Select(Limpar).Where(t => t.Length > 0 && !Igual(t, DeFabrica)).ToList());

    /// <summary>Sem espaços nas pontas e sem a pontuação do fim (o ";" é posto na proposta).</summary>
    public static string Limpar(string? texto) =>
        string.Join(" ", (texto ?? "").Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)).TrimEnd(';', '.', ',').Trim();

    public static bool Igual(string? a, string? b) => Textos.Igual(Limpar(a), Limpar(b));

    /// <summary>O texto da proposta (em branco = o padrão de fábrica).</summary>
    public static string DaProposta(string? jornada) => Limpar(jornada) is { Length: > 0 } t ? t : DeFabrica;

    /// <summary>A proposta usa o padrão de fábrica?</summary>
    public static bool EhDeFabrica(string? jornada) => Igual(DaProposta(jornada), DeFabrica);

    /// <summary>
    /// O parágrafo no idioma da proposta, sem pontuação no fim: o padrão de
    /// fábrica sai traduzido; um texto digitado sai como foi escrito.
    /// </summary>
    public static string NoIdioma(string? jornada, string idioma)
    {
        if (!EhDeFabrica(jornada)) return DaProposta(jornada);
        return IdiomasDaProposta.Normalizar(idioma) switch
        {
            IdiomasDaProposta.Ingles => DeFabricaEn,
            IdiomasDaProposta.Espanhol => DeFabricaEs,
            _ => DeFabrica,
        };
    }
}
