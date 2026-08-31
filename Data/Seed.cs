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
        // ---- mão de obra (CUSTO B8:H17) — mult = multiplicador sobre a hora-base ----
        P("mo-diaria-1turno",  "MO", "DIARIAS NORMAIS:", "1o. TURNO", "8", "320", "Sim", 1, "1"),
        P("mo-diaria-2turno",  "MO", "DIARIAS NORMAIS:", "2o. TURNO NOITE", "8", "480", "Sim", 2, "1.5"),
        P("mo-diaria-extra",   "MO", "DIARIAS EXTRAS SAB, DOM E FER", "", "8", "640", "Sim", 3, "2"),
        P("mo-he-8h",          "MO", "HORAS EXTRAS SUP. 8h/DIA", "", "1", "480", "Sim", 4, "1.5"),
        P("mo-he-sab",         "MO", "HORAS EXTRAS SAB, DOM E FER", "", "1", "480", "Não", 5, "1.5"),
        P("mo-equipamentos",   "MO", "EQUIPAMENTOS P/ MONTAGEM", "", "1", "0", "Não", 6, "0"),
        P("mo-ferramentas",    "MO", "FERRAMENTAS", "", "1", "0", "Não", 7, "0"),
        P("mo-terceiros",      "MO", "TERCEIROS", "", "1", "0", "Não", 8, "0"),
        P("mo-treinamentos",   "MO", "TREINAMENTOS", "", "1", "0", "Sim", 9, "0"),

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

    /// <summary>Representantes e comissões, direto de BD_pricing A67:E99 da planilha.</summary>
    public static List<Representante> Representantes() => new()
    {
        Rep("douglas-mezza", "Douglas (Mezza & Baga)", "Brasil", "3", "Douglas M. Matavelli por (11) 97144-3085 ou e-mail: douglas.matavelli@chartindustries.com"),
        Rep("alexandre-artman", "Alexandre (Artman)", "Brasil", "3", "Alexandre B. Pereira por (91) 98883-8142 / (16) 99429-1786 ou e-mail: alexandre.pereira@artman.net.br"),
        Rep("gerson-lizan", "Gerson (Lizan)", "Brasil", "3", ""),
        Rep("ivars-dzelme", "Ivars (Dzelme & Leite Ltda)", "Brasil", "3", "Ivars Janis Dzelme por (81) 3221-0250 / (81) 99946-0506 ou e-mail: ivars@hotlink.com.br"),
        Rep("julio-doulus", "Júlio (Doulus)", "Brasil", "3", "Júlio Augusto Afro por (27) 3314-1000 / (27) 98122-1177 ou e-mail: howden@doulus.com.br"),
        Rep("mauricio-livimat", "Maurício (Livimat)", "Brasil", "3", "Mauricio A. de Araujo por (21) 99908-1687 ou e-mail: Livimat.comercial@outlook.com"),
        Rep("ricardo-sesbras", "Ricardo (Sesbras)", "Brasil", "5", "Ricardo V. F. Martins por (21) 2532-7404 / (21) 99764-5297 ou e-mail: aviabras@aviabras.com.br"),
        Rep("sander-provent", "Sander (Provent)", "Brasil", "5", ""),
        Rep("thais-intec", "Thais (InTec)", "Brasil", "3", "InTec – Engª Thais Werner de Lima por (71) 3289-3611 / (71) 9 9961-9278 ou e-mail: intec@inovacaotecnologia.com.br"),
        Rep("adolpho-atric", "Adolpho (Atric)", "Brasil / Exterior", "5", "Adolpho Procópio Rossi Neto por (11) 99976-1952 ou e-mail: rossi@atric.com.br"),
        Rep("polonio-giovanni", "Polonio (Giovanni e Polonio)", "Brasil", "3", "Walter Luiz Polonio por (14) 9 9616-0560 ou e-mail: wlpolonio@terra.com.br"),
        Rep("aseq-500k", "ASESORIA Y EQUIPO < =USD 500K", "Exterior", "10", "Pablo Santamarina por +502 24285468, 24285478, 23658515, 23658669 ou e-mail: aseqsa@gmail.com"),
        Rep("aseq-1mm", "ASESORIA Y EQUIPO < USD 1MM", "Exterior", "7.5", "Pablo Santamarina por +502 24285468, 24285478, 23658515, 23658669 ou e-mail: aseqsa@gmail.com"),
        Rep("aseq-mais1mm", "ASESORIA Y EQUIPO > USD 1MM", "Exterior", "5", "Pablo Santamarina por +502 24285468, 24285478, 23658515, 23658669 ou e-mail: aseqsa@gmail.com"),
        Rep("ferrunion", "FERRUNION", "Exterior", "5", "Gilmer Vasquez por: +51 1 4754560 ou e-mail: gsvasquez@ferrunion.net"),
        Rep("het", "H&T", "Exterior", "5", "Jorge Gonzalo Hernández Cabeza por +56 2 29970179 / +56 9 98871135 ou e-mail: jhernandez@ghis.cl"),
        Rep("hca", "HCA", "Exterior", "6", "Angelo Ramirez por +56 9 4478 3695 / +56 2 5725-7371 o e-mail: angelo@hcamineria.cl"),
        Rep("hri", "HRI S.A.", "Exterior", "5", "Rury Harms Orrego por +56 2 2592 3500 ou e-mail: rharms@hri.cl"),
        Rep("ipt-125mm", "IPT Colômbia < =EUR 1,25MM", "Exterior", "3.5", "Ricardo Morales Castro por: +57 3125866426 / 3206737171 ou e-mail: rmorales@iptcolombia.com"),
        Rep("ipt-1mm", "IPT Colômbia < =EUR 1MM", "Exterior", "5", "Ricardo Morales Castro por: +57 3125866426 / 3206737171 ou e-mail: rmorales@iptcolombia.com"),
        Rep("ipt-500k", "IPT Colômbia < =EUR 500K", "Exterior", "8", "Ricardo Morales Castro por: +57 3125866426 / 3206737171 ou e-mail: rmorales@iptcolombia.com"),
        Rep("ipt-750k", "IPT Colômbia < =EUR 750K", "Exterior", "6.5", "Ricardo Morales Castro por: +57 3125866426 / 3206737171 ou e-mail: rmorales@iptcolombia.com"),
        Rep("ipt-mais125", "IPT Colômbia >EUR 1,25MM", "Exterior", "3", "Ricardo Morales Castro por: +57 3125866426 / 3206737171 ou e-mail: rmorales@iptcolombia.com"),
        Rep("siminco", "SIMINCO", "Exterior", "5", "Alejandro Cadavid L. por +57 323 460 0551 o e-mail comercial@siminco.com.co o Carlos Contreras U. por +57 311 588 4883 o e-mail coordinadortecnico@siminco.com.co"),
        Rep("sistagua", "SISTAGUA", "Exterior", "5", "Melissa Melville por +502 5990 7944 o e-mail: mmelville@sistagua.com"),
        Rep("tejada", "TEJADA", "Exterior", "5", "Luis Felipe Tejada por :+57 315-505-5397 ou e-mail: Tejadaingenieros@tejadaingenieros.com"),
        Rep("turbo-100k", "Turbomaquinarias <= EUR 100 k", "Exterior", "4", "Carlos Daniel Weihmuller por e-mail: cweihmuller@turbomaquinarias.com"),
        Rep("turbo-25m", "Turbomaquinarias <= EUR 2,5 M", "Exterior", "2.5", "Carlos Daniel Weihmuller por e-mail: cweihmuller@turbomaquinarias.com"),
        Rep("turbo-500k", "Turbomaquinarias <= EUR 500 k", "Exterior", "3", "Carlos Daniel Weihmuller por e-mail: cweihmuller@turbomaquinarias.com"),
        Rep("turbo-70m", "Turbomaquinarias <= EUR 7,0 M", "Exterior", "2", "Carlos Daniel Weihmuller por e-mail: cweihmuller@turbomaquinarias.com"),
        Rep("turbo-mais70m", "Turbomaquinarias > EUR 7,0 M", "Exterior", "1", "Carlos Daniel Weihmuller por e-mail: cweihmuller@turbomaquinarias.com"),
    };

    private static Representante Rep(string id, string nome, string local, string com, string contato) =>
        new() { Id = id, Nome = nome, Local = local, ComissaoPct = com, Contato = contato };

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
        string horas, string valor, string porTecnico, int ordem, string mult = "0") => new()
    {
        Id = id, Tipo = tipo, Descricao = desc, Obs = obs,
        Horas = horas, Valor = valor, PorTecnico = porTecnico, Ordem = ordem.ToString("D3"),
        Mult = mult,
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

        if (store.IsEmpty("representantes"))
        {
            var repo = new RepresentanteRepository(store);
            foreach (var x in Seed.Representantes()) repo.Save(x);
        }
    }
}
