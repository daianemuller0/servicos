namespace HowdenServicos.Poc.Data;

/// <summary>
/// Tradução dos textos que vêm da TABELA DE CUSTOS (nomes de serviços,
/// despesas e observações) e das linhas geradas pelo sistema. A tabela é
/// cadastrada em português; o documento sai no idioma da proposta.
///
/// O que não estiver no dicionário sai como foi digitado — itens criados à
/// mão pela equipe continuam aparecendo exatamente como escritos.
/// </summary>
public static partial class Traducoes
{
    private static readonly Dictionary<string, (string Es, string En)> Base = new()
    {
        // ---- serviços (mão de obra) ----
        ["DIARIAS NORMAIS:"] = ("DÍAS NORMALES:", "REGULAR DAILY RATES:"),
        ["DIARIAS NORMAIS"] = ("DÍAS NORMALES", "REGULAR DAILY RATES"),
        ["DIARIAS EXTRAS SAB, DOM E FER"] = ("DÍAS EXTRAS SÁB, DOM Y FER", "EXTRA DAYS SAT, SUN & HOLIDAYS"),
        ["HORAS EXTRAS SUP. 8H/DIA"] = ("HORAS EXTRAS SUP. 8h/DÍA", "OVERTIME OVER 8h/DAY"),
        ["HORAS EXTRAS SAB, DOM E FER"] = ("HORAS EXTRAS SÁB, DOM Y FER", "OVERTIME SAT, SUN & HOLIDAYS"),
        ["EQUIPAMENTOS P/ MONTAGEM"] = ("EQUIPOS PARA MONTAJE", "ASSEMBLY EQUIPMENT"),
        ["FERRAMENTAS"] = ("HERRAMIENTAS", "TOOLS"),
        ["TERCEIROS"] = ("TERCEROS", "THIRD PARTIES"),
        ["TREINAMENTOS"] = ("CAPACITACIONES", "TRAINING"),

        // ---- observações dos serviços ----
        ["1O. TURNO"] = ("1er. TURNO", "1st SHIFT"),
        ["2O. TURNO NOITE"] = ("2do. TURNO NOCHE", "2nd SHIFT (NIGHT)"),

        // ---- despesas ----
        ["TAXI"] = ("TAXI", "TAXI"),
        ["TÁXI"] = ("TAXI", "TAXI"),
        ["PASSAGEM AEREA"] = ("PASAJE AÉREO", "AIRFARE"),
        ["PASSAGEM AÉREA"] = ("PASAJE AÉREO", "AIRFARE"),
        ["HOSPEDAGEM"] = ("HOSPEDAJE", "ACCOMMODATION"),
        ["LOCACAO DE CARRO"] = ("ALQUILER DE AUTO", "CAR RENTAL"),
        ["COMBUSTIVEL"] = ("COMBUSTIBLE", "FUEL"),
        ["REFEICOES"] = ("COMIDAS", "MEALS"),
        ["PEDAGIOS"] = ("PEAJES", "TOLLS"),
        ["ESTACIONAMENTO"] = ("ESTACIONAMIENTO", "PARKING"),
        ["EXAMES / ASO"] = ("EXÁMENES MÉDICOS", "MEDICAL EXAMS"),
        ["OUTROS"] = ("OTROS", "OTHERS"),

        // ---- observações das despesas ----
        ["AEROPORTO SP IDA E VOLTA"] = ("AEROPUERTO SP IDA Y VUELTA", "SP AIRPORT ROUND TRIP"),
        ["IDA E VOLTA"] = ("IDA Y VUELTA", "ROUND TRIP"),
        ["DIARIA HOTEL"] = ("DÍA DE HOTEL", "HOTEL NIGHT"),
        ["ALMOCO /JANTAR"] = ("ALMUERZO / CENA", "LUNCH / DINNER"),
        ["ALMOCO / JANTAR"] = ("ALMUERZO / CENA", "LUNCH / DINNER"),
        ["ADMINISTRATIVA"] = ("ADMINISTRATIVA", "ADMINISTRATIVE"),

        // ---- linhas geradas pelo sistema (diárias adicionais) ----
        ["DIARIA ADICIONAL"] = ("DÍA ADICIONAL", "ADDITIONAL DAY"),
        ["DIARIA ADICIONAL: 2O. TURNO NOITE"] = ("DÍA ADICIONAL: 2do. TURNO NOCHE", "ADDITIONAL DAY: 2nd SHIFT (NIGHT)"),
        ["DIARIA EXTRA SAB, DOM E FER"] = ("DÍA EXTRA SÁB, DOM Y FER", "EXTRA DAY SAT, SUN & HOLIDAYS"),
        ["HORA EXTRA SUP. 8H/DIA"] = ("HORA EXTRA SUP. 8h/DÍA", "OVERTIME HOUR OVER 8h/DAY"),
        ["HORA EXTRA SAB, DOM E FER"] = ("HORA EXTRA SÁB, DOM Y FER", "OVERTIME HOUR SAT, SUN & HOLIDAYS"),
        ["1O. TURNO — 8 HORAS (DIAS UTEIS DE SEG. A SEX.)"] =
            ("1er. TURNO — 8 HORAS (DÍAS HÁBILES DE LUN. A VIE.)", "1st SHIFT — 8 HOURS (WORKING DAYS, MON–FRI)"),
        ["DAS 8:00 AS 17:00 (DIAS UTEIS DE SEG. A SEX.)"] =
            ("DE 8:00 A 17:00 (DÍAS HÁBILES DE LUN. A VIE.)", "FROM 8:00 TO 17:00 (WORKING DAYS, MON–FRI)"),
        ["SUPERIOR A 8/H DIA (DIAS UTEIS DE SEG. A SEX.)"] =
            ("SUPERIOR A 8 h/DÍA (DÍAS HÁBILES DE LUN. A VIE.)", "OVER 8h/DAY (WORKING DAYS, MON–FRI)"),
        ["SAB., DOM. E FERIADOS"] = ("SÁB., DOM. Y FERIADOS", "SAT., SUN. AND HOLIDAYS"),
        ["HORA EXTRA"] = ("HORA EXTRA", "OVERTIME HOUR"),

        // ---- escopos de serviço sugeridos (título da assessoria) ----
        ["COMISSIONAMENTO E STARTUP"] = ("Puesta en marcha y comisionamiento", "Commissioning and start-up"),
        ["TREINAMENTO DE PRINCIPIOS BASICOS DE VENTILADORES"] =
            ("Capacitación en principios básicos de ventiladores", "Training on fan fundamentals"),
        ["TREINAMENTO PARA OPERACAO E MANUTENCAO DE VENTILADORES CENTRIFUGOS"] =
            ("Capacitación para operación y mantenimiento de ventiladores centrífugos",
             "Training on operation and maintenance of centrifugal fans"),
        ["MONTAGEM, COMISSIONAMENTO E STARTUP"] =
            ("Montaje, comisionamiento y puesta en marcha", "Assembly, commissioning and start-up"),
        ["SERVICO DE INSPECAO"] = ("Servicio de inspección", "Inspection service"),
        ["LEVANTAMENTO DE CAMPO"] = ("Levantamiento en campo", "Field survey"),
        ["ESTUDO DE VENTILACAO"] = ("Estudio de ventilación", "Ventilation study"),
        ["SUPERVISAO TECNICA DE ESPECIALISTA HOWDEN"] =
            ("Supervisión técnica de especialista Howden", "Technical supervision by a Howden specialist"),
        ["LEVANTAMENTO DOS PONTOS OPERACIONAIS DO EQUIPAMENTO"] =
            ("Levantamiento de los puntos operativos del equipo", "Survey of the equipment operating points"),
    };

    // chave normalizada (maiúsculas, sem acento) → (espanhol, inglês).
    // Entradas que normalizam para a mesma chave (ex.: "TAXI" e "TÁXI") são
    // a mesma tradução: fica a primeira.
    private static readonly Dictionary<string, (string Es, string En)> Itens = Base
        .GroupBy(x => Chave(x.Key))
        .ToDictionary(g => g.Key, g => g.First().Value);

    /// <summary>
    /// Traduz um texto da tabela de custos para o idioma da proposta.
    /// Português (ou texto desconhecido) volta igual.
    /// </summary>
    public static string Item(string? texto, string idioma)
    {
        if (string.IsNullOrWhiteSpace(texto) || idioma == "Português") return texto ?? "";
        if (!Itens.TryGetValue(Chave(texto), out var t)) return texto;
        return idioma == "English" ? t.En : t.Es;
    }

    private static string Chave(string texto)
    {
        var normal = texto.Trim().ToUpperInvariant().Normalize(System.Text.NormalizationForm.FormD);
        var sb = new System.Text.StringBuilder();
        foreach (var c in normal)
            if (System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c)
                != System.Globalization.UnicodeCategory.NonSpacingMark)
                sb.Append(c);
        return sb.ToString();
    }
}
