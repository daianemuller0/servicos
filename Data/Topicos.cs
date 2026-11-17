using System.Text.Json;

namespace HowdenServicos.Poc.Data;

/// <summary>Um grupo de tópicos: título ("Assessoria de Montagem") e os itens dele.</summary>
public sealed class GrupoDeTopicos
{
    public string Titulo { get; set; } = "";
    public List<string> Itens { get; set; } = new();
}

/// <summary>
/// Os tópicos com tic da proposta: a "Descrição das Atividades de
/// Assessoria" e os padrões do "Escopo do Serviço". Os padrões ficam na
/// configuração (tela E-mails e Padrões / banco): o que alguém adiciona na
/// proposta vira padrão para todos. A proposta guarda o TEXTO do que foi
/// marcado — mudar um padrão depois não altera proposta já feita.
/// </summary>
public static class Topicos
{
    public const string ChaveAtividades = "padroes.atividades";
    public const string ChaveEscopo = "padroes.escopo";

    public const string TituloAtividades = "Descrição das Atividades de Assessoria";

    /// <summary>As atividades de fábrica (um grupo só).</summary>
    public static List<GrupoDeTopicos> AtividadesDeFabrica() => new()
    {
        new GrupoDeTopicos
        {
            Titulo = TituloAtividades,
            Itens = new()
            {
                "Inspeção inicial do equipamento",
                "Verificação do estado geral de todos os componentes",
                "Inspeção visual dos mancais",
                "Inspeção visual do acoplamento",
                "Inspeção dos colarinhos flexíveis",
                "Inspeção e teste de abertura e fechamento do damper",
                "Abertura das portas de inspeção para acesso interno ao ventilador",
                "Abertura do bipartido do mancal monobloco para inspeção interna",
            },
        },
    };

    /// <summary>Os dois padrões de escopo de fábrica.</summary>
    public static List<GrupoDeTopicos> EscopoDeFabrica() => new()
    {
        new GrupoDeTopicos
        {
            Titulo = "Assessoria de Montagem",
            Itens = new()
            {
                "Apoio ao planejamento e preparação das atividades de campo",
                "Acompanhamento da manutenção dos componentes",
                "Acompanhamento da desmontagem dos componentes existentes",
                "Acompanhamento da montagem mecânica dos equipamentos e acessórios",
                "Acompanhamento do alinhamento e posicionamento dos conjuntos girantes",
                "Acompanhamento da instalação dos elementos de fixação e respectivos torques de aperto",
                "Acompanhamento da instalação de sistemas auxiliares, quando aplicável",
                "Interface técnica com a equipe de montagem para esclarecimento de dúvidas e orientação de boas práticas",
            },
        },
        new GrupoDeTopicos
        {
            Titulo = "Comissionamento e Start-up",
            Itens = new()
            {
                "Inspeção final de montagem mecânica",
                "Verificação dimensional e visual dos componentes instalados",
                "Conferência do aperto e travamento dos elementos de fixação",
                "Verificação dos níveis e condições de lubrificação",
                "Verificação das interligações elétricas e de instrumentação relacionadas ao equipamento",
                "Inspeção dos sistemas de proteção e segurança do equipamento",
                "Acompanhamento do teste de giro livre e confirmação da ausência de interferências mecânicas",
                "Verificação do sentido de rotação do conjunto girante",
                "Acompanhamento da partida a frio do equipamento",
                "Monitoramento das condições operacionais durante a elevação gradual da carga",
                "Medição e análise dos níveis de vibração",
                "Monitoramento de temperatura de mancais, motor e demais componentes críticos",
                "Verificação da estabilidade operacional do equipamento",
                "Orientação operacional à equipe do cliente quanto às melhores práticas de operação",
                "Registro de ocorrências e recomendações para otimização da operação",
            },
        },
    };

    /// <summary>Os padrões gravados na configuração; sem nada gravado, os de fábrica.</summary>
    public static List<GrupoDeTopicos> Padroes(IReadOnlyDictionary<string, string>? cfg, string chave)
    {
        if (cfg is not null && cfg.TryGetValue(chave, out var json) && Ler(json) is { Count: > 0 } gravados)
            return gravados;
        return chave == ChaveAtividades ? AtividadesDeFabrica() : EscopoDeFabrica();
    }

    /// <summary>Lê uma lista de grupos (padrão gravado ou seleção da proposta). Inválido/vazio = lista vazia.</summary>
    public static List<GrupoDeTopicos> Ler(string? json)
    {
        if (string.IsNullOrWhiteSpace(json)) return new();
        try
        {
            return (JsonSerializer.Deserialize<List<GrupoDeTopicos>>(json) ?? new())
                .Where(g => g is not null)
                .Select(g => new GrupoDeTopicos { Titulo = (g.Titulo ?? "").Trim(), Itens = Limpar(g.Itens) })
                .ToList();
        }
        catch (JsonException) { return new(); }
    }

    public static string Gravar(List<GrupoDeTopicos> grupos) => JsonSerializer.Serialize(
        grupos.Select(g => new GrupoDeTopicos { Titulo = g.Titulo, Itens = Limpar(g.Itens) }).ToList());

    /// <summary>A seleção da proposta, só com os grupos que têm item marcado.</summary>
    public static List<GrupoDeTopicos> Marcados(string? json) =>
        Ler(json).Where(g => g.Itens.Count > 0).ToList();

    /// <summary>Tira espaço, ";" e "." do fim (a pontuação é posta na hora de imprimir) e repetidos.</summary>
    public static string LimparItem(string? item) => (item ?? "").Trim().TrimEnd(';', '.', ',').Trim();

    private static List<string> Limpar(IEnumerable<string>? itens) => (itens ?? Enumerable.Empty<string>())
        .Select(LimparItem).Where(i => i.Length > 0)
        .Distinct(StringComparer.OrdinalIgnoreCase).ToList();

    public static bool Igual(string? a, string? b) => Textos.Igual(LimparItem(a), LimparItem(b));

    /// <summary>
    /// Liga/desliga um item na seleção, mantendo a ordem dos padrões (itens
    /// marcados que não estão mais no padrão ficam no fim do grupo).
    /// </summary>
    public static List<GrupoDeTopicos> Marcar(List<GrupoDeTopicos> padroes, List<GrupoDeTopicos> selecao,
        string titulo, IEnumerable<string> itens, bool ligar)
    {
        var mudar = itens.Select(LimparItem).Where(i => i.Length > 0).ToList();
        var grupo = selecao.FirstOrDefault(g => Igual(g.Titulo, titulo));
        if (grupo is null) { grupo = new GrupoDeTopicos { Titulo = titulo }; selecao.Add(grupo); }
        if (ligar) grupo.Itens.AddRange(mudar.Where(i => !grupo.Itens.Any(x => Igual(x, i))));
        else grupo.Itens.RemoveAll(x => mudar.Any(i => Igual(x, i)));
        return Ordenar(padroes, selecao);
    }

    private static List<GrupoDeTopicos> Ordenar(List<GrupoDeTopicos> padroes, List<GrupoDeTopicos> selecao)
    {
        var saida = new List<GrupoDeTopicos>();
        foreach (var p in padroes)
        {
            var s = selecao.FirstOrDefault(g => Igual(g.Titulo, p.Titulo));
            if (s is null || s.Itens.Count == 0) continue;
            var itens = p.Itens.Where(i => s.Itens.Any(x => Igual(x, i))).ToList();
            itens.AddRange(s.Itens.Where(x => !p.Itens.Any(i => Igual(x, i))));
            saida.Add(new GrupoDeTopicos { Titulo = p.Titulo, Itens = itens });
        }
        saida.AddRange(selecao.Where(s => s.Itens.Count > 0 && !padroes.Any(p => Igual(p.Titulo, s.Titulo))));
        return saida;
    }

    /// <summary>"item;" e o último "item." — como na proposta.</summary>
    public static string ComPontuacao(string item, bool ultimo) => item + (ultimo ? "." : ";");
}
