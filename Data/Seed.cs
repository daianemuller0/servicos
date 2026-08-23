using HowdenServicos.Poc.Models;

namespace HowdenServicos.Poc.Data;

/// <summary>
/// Valores padrão vindos da planilha "Ferramenta para propostas de serviço"
/// (guia CUSTO): serviços de mão de obra e despesas com seus preços de custo.
/// </summary>
public static class Seed
{
    public static List<Parametro> Parametros() => new()
    {
        // ---- mão de obra (CUSTO B8:H17) ----
        P("mo-diaria-1turno",  "MO", "DIARIAS NORMAIS:", "1o. TURNO", "8", "320", "Sim", 1),
        P("mo-diaria-2turno",  "MO", "DIARIAS NORMAIS:", "2o. TURNO NOITE", "8", "480", "Sim", 2),
        P("mo-diaria-extra",   "MO", "DIARIAS EXTRAS SAB, DOM E FER", "", "8", "640", "Sim", 3),
        P("mo-he-8h",          "MO", "HORAS EXTRAS SUP. 8h/DIA", "", "1", "480", "Sim", 4),
        P("mo-he-sab",         "MO", "HORAS EXTRAS SAB, DOM E FER", "", "1", "640", "Não", 5),
        P("mo-equipamentos",   "MO", "EQUIPAMENTOS P/ MONTAGEM", "", "1", "0", "Não", 6),
        P("mo-ferramentas",    "MO", "FERRAMENTAS", "", "1", "0", "Não", 7),
        P("mo-terceiros",      "MO", "TERCEIROS", "", "1", "0", "Não", 8),
        P("mo-treinamentos",   "MO", "TREINAMENTOS", "", "1", "0", "Sim", 9),

        // ---- despesas (CUSTO B22:H31) ----
        P("de-taxi",       "DESPESA", "TAXI", "AEROPORTO SP IDA E VOLTA", "", "280", "Sim", 20),
        P("de-passagem",   "DESPESA", "PASSAGEM AEREA", "IDA E VOLTA", "", "3500", "Sim", 21),
        P("de-hospedagem", "DESPESA", "HOSPEDAGEM", "DIARIA HOTEL", "", "400", "Sim", 22),
        P("de-carro",      "DESPESA", "LOCAÇÃO DE CARRO", "", "", "160", "Sim", 23),
        P("de-combustivel","DESPESA", "COMBUSTIVEL", "", "", "30", "Sim", 24),
        P("de-refeicoes",  "DESPESA", "REFEIÇÕES", "ALMOÇO /JANTAR", "", "100", "Sim", 25),
        P("de-pedagio",    "DESPESA", "PEDÁGIOS", "", "", "0", "Sim", 26),
        P("de-estacion",   "DESPESA", "ESTACIONAMENTO", "", "", "0", "Sim", 27),
        P("de-exames",     "DESPESA", "EXAMES / ASO", "ADMINISTRATIVA", "", "350", "Sim", 28),
        P("de-outros",     "DESPESA", "OUTROS", "ADMINISTRATIVA", "", "0", "Sim", 29),
    };

    public static List<BillingInfo> Faturamento() => new()
    {
        new BillingInfo
        {
            Id = "HSA-SP",
            Razao = "Howden South America Ventiladores e Compressores Indústria e Comércio Ltda.",
            Endereco = "Av. Osvaldo Berto, 475, Distrito Industrial Alfredo Rela, 13255-405 – Itatiba – SP",
            Registro = "CNPJ: 01.094.363/0001-04 – Inscrição Estadual: 382.062.017.112",
            BancoNome = "Banco Itaú", Agencia = "4892", Conta = "28480-5",
        },
        new BillingInfo { Id = "HSA-ES" },
        new BillingInfo { Id = "HCHL" },
        new BillingInfo { Id = "HPU" },
    };

    private static Parametro P(string id, string tipo, string desc, string obs,
        string horas, string valor, string porTecnico, int ordem) => new()
    {
        Id = id, Tipo = tipo, Descricao = desc, Obs = obs,
        Horas = horas, Valor = valor, PorTecnico = porTecnico, Ordem = ordem.ToString("D3"),
    };
}

/// <summary>Semeia as tabelas de apoio na primeira execução (pasta vazia).</summary>
public static class DbInitializer
{
    public static void Initialize(ParquetStore store)
    {
        if (store.IsEmpty("parametros"))
        {
            var repo = new ParametroRepository(store);
            foreach (var p in Seed.Parametros()) repo.Save(p);
        }

        if (store.IsEmpty("faturamento"))
        {
            var repo = new FaturamentoRepository(store);
            foreach (var b in Seed.Faturamento()) repo.Save(b);
        }
    }
}
