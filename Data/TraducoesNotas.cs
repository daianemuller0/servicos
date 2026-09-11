namespace HowdenServicos.Poc.Data;

/// <summary>
/// Notas da proposta (INCLUSO / EXCLUSOS / Notas Gerais) em espanhol e
/// inglês. A chave é a mesma usada no tic de seleção da tela, então ligar
/// ou desligar uma nota vale para os três idiomas.
/// </summary>
public static partial class Traducoes
{
    private static readonly Dictionary<string, (string Es, string En)> NotasTraduzidas = new()
    {
        // ---- itens variáveis (entram no incluso ou no excluso conforme o custo) ----
        ["var.passagem"] = ("Pasaje aéreo;", "Airfare;"),
        ["var.traslado"] = ("Traslado;", "Transfer;"),
        ["var.hospedagem"] = ("Hospedaje;", "Accommodation;"),
        ["var.locomocoes"] = ("Desplazamientos;", "Local transportation;"),
        ["var.estadias"] = ("Estadías;", "Lodging;"),
        ["var.lanches"] = ("Refrigerios;", "Snacks;"),
        ["var.alimentacao"] = ("Alimentación;", "Meals;"),
        ["var.transporte"] = ("Transporte (auto, peajes, combustible, etc.);",
                              "Transportation (car, tolls, fuel, etc.);"),

        // ---- inclusos fixos ----
        ["inc.tempo-viagem"] = (
            "El tiempo de viaje y/o de desplazamiento en el origen y en el destino y las horas de inducción se computarán como período trabajado;",
            "Travel and commuting time at origin and destination, as well as site induction hours, will be counted as worked time;"),
        ["inc.20dias"] = (
            "El Cliente podrá, con un plazo mínimo de 20 días, solicitar a Howden la ejecución del Servicio. La falta de manifestación del Cliente dentro del plazo y la ejecución del Servicio por cuenta propia o por terceros cancela automáticamente la garantía contractual concedida inicialmente;",
            "The Customer may, with a minimum notice of 20 days, request Howden to perform the Service. Failure to notify within that period, or performing the Service in-house or through third parties, automatically terminates the contractual warranty initially granted;"),
        ["inc.prazo-tecnico"] = (
            "Las partes deben acordar un plazo técnico razonable para la ejecución del Servicio; a partir del aviso del Cliente, Howden definirá el o los Técnicos responsables y preparará la documentación necesaria para su inducción, de acuerdo con las exigencias del Cliente;",
            "The parties shall agree on a reasonable technical schedule for performing the Service; upon the Customer's notice, Howden will assign the responsible Technician(s) and arrange the documentation required for their site induction, according to the Customer's requirements;"),
        ["inc.epi"] = ("Equipo de Protección Individual (EPI).", "Personal Protective Equipment (PPE)."),

        // ---- exclusos: treinamento ----
        ["exc.trein.didatico-digital"] = (
            "Suministro de material didáctico de la capacitación en medio digital (correo electrónico) – EN CASO DE CAPACITACIONES;",
            "Supply of training materials in digital format (e-mail) – FOR TRAINING SESSIONS;"),
        ["exc.trein.projetor"] = (
            "Suministro de proyector para las presentaciones de la capacitación – EN CASO DE CAPACITACIONES;",
            "Supply of a projector for training presentations – FOR TRAINING SESSIONS;"),
        ["exc.trein.didatico-impresso"] = (
            "Suministro de material didáctico impreso de la capacitación – EN CASO DE CAPACITACIONES;",
            "Supply of printed training materials – FOR TRAINING SESSIONS;"),
        ["exc.trein.desenhos"] = (
            "Suministro de planos e información constructiva – EN CASO DE CAPACITACIONES;",
            "Supply of drawings and construction information – FOR TRAINING SESSIONS;"),
        ["exc.trein.procedimentos"] = (
            "Suministro de procedimientos de fabricación – EN CASO DE CAPACITACIONES;",
            "Supply of manufacturing procedures – FOR TRAINING SESSIONS;"),

        // ---- exclusos fixos ----
        ["exc.area-classificada"] = (
            "No se considera trabajo en área clasificada/explosiva;",
            "Work in classified/explosive areas is not included;"),
        ["exc.art"] = ("No se considera la emisión de ART (registro técnico);",
                       "Technical responsibility registration (ART) is not included;"),
        ["exc.pgr"] = (
            "No se contemplan PGR, PCMSO ni otros documentos específicos para este servicio;",
            "PGR, PCMSO and other service-specific health and safety documents are not included;"),
        ["exc.noturno"] = ("Trabajo en período nocturno;", "Night-shift work;"),
        ["exc.covid"] = ("No se prevé tiempo para protocolos de COVID-19;",
                         "Time for COVID-19 protocols is not included;"),
        ["exc.pintura"] = ("No se prevé pintura en nuestro alcance;", "Painting is not included in our scope;"),
        ["exc.materiais-instalacao"] = ("Suministro de materiales de instalación/aplicación;",
                                        "Supply of installation/application materials;"),
        ["exc.eletricista"] = ("Suministro de mano de obra de electricista para los servicios;",
                               "Supply of electrician labor for the services;"),
        ["exc.andaimes"] = ("Suministro de andamios, plataformas elevadoras y grúas;",
                            "Supply of scaffolding, aerial platforms and cranes;"),
        ["exc.civil"] = ("Mano de obra civil;", "Civil works labor;"),
        ["exc.eletrica-instrumentacao"] = ("Materiales y accesorios de eléctrica e instrumentación;",
                                           "Electrical and instrumentation materials and accessories;"),
        ["exc.fundacoes"] = ("Proyecto, equipos y mano de obra para fundaciones y obras civiles;",
                             "Design, equipment and labor for foundations and civil works;"),
        ["exc.engenharia-adicional"] = (
            "Servicios adicionales de nuestro departamento de Ingeniería generados por cambios posteriores a la orden de compra solicitados por el cliente;",
            "Additional services from our Engineering department arising from changes requested by the customer after the purchase order;"),
        ["exc.protocolos-comunicacao"] = (
            "Cualquier protocolo de comunicación, como Hart, Modbus, Profibus, etc., para los instrumentos incluidos en esta oferta;",
            "Any communication protocol, such as Hart, Modbus, Profibus, etc., for the instruments included in this quotation;"),
        ["exc.nao-consta"] = (
            "Queda excluido de nuestro suministro cualquier ítem o accesorio que no conste claramente en nuestra oferta;",
            "Any item or accessory not expressly stated in our quotation is excluded from our supply;"),
        ["exc.quantidade-total"] = (
            "Los valores presentados son válidos solamente para la cantidad total ofertada; si esa cantidad cambia, los valores deberán recalcularse y presentarse en una revisión de la oferta;",
            "The prices presented are valid only for the total quantity offered; if that quantity changes, prices must be recalculated and presented in a revised quotation;"),
        ["exc.lucros-cessantes"] = (
            "<b>Exclusión de lucro cesante.</b> Howden no será, en ninguna hipótesis, responsable por lucro cesante y/o daños indirectos de cualquier tipo, incluyendo, entre otros, pérdida de negocio, de beneficio o de productividad, <b>conforme al ítem 12 de las condiciones de suministro</b>;",
            "<b>Exclusion of consequential loss.</b> Howden shall in no event be liable for loss of profit and/or indirect damages of any kind, including but not limited to loss of business, profit or productivity, <b>as per item 12 of the conditions of supply</b>;"),

        // ---- notas gerais ----
        ["ger.adicional-noturno"] = (
            "(*) Recargo – adicional nocturno (22h00 a 05h00) del 50% sobre los precios informados arriba;",
            "(*) Surcharge – night premium (10:00 p.m. to 5:00 a.m.) of 50% over the prices stated above;"),
        ["ger.iss"] = (
            "Impuesto: ISS incluido — \"ISS recaudado en el municipio del prestador, conforme a la ley 3667/2003.\";",
            "Tax: ISS included — \"ISS collected in the service provider's municipality, pursuant to Law 3667/2003.\";"),
        ["ger.medicao"] = (
            "El valor total se determinará al término de los servicios y conforme a la medición realizada;",
            "The total amount will be determined upon completion of the services and according to the measured work;"),
        ["ger.valor-minimo"] = (
            "El valor mínimo corresponde a 01 (un) día normal de trabajo, es decir, 08 horas;",
            "The minimum charge corresponds to 01 (one) regular working day, i.e. 08 hours;"),
        ["ger.carga-horaria"] = (
            "Jornada máxima de 08 horas por día; en casos especiales se cobrarán horas adicionales (50% en semana, 100% sábados, domingos y feriados);",
            "Maximum working time of 08 hours per day; in special cases additional hours will be charged (50% on weekdays, 100% on Saturdays, Sundays and holidays);"),
        ["ger.reajuste-180"] = (
            "Los valores podrán reajustarse después de 180 días de la aceptación de la orden de compra;",
            "Prices may be adjusted 180 days after acceptance of the purchase order;"),
        ["ger.despesas-howden"] = (
            "Pasajes, traslados, desplazamientos, hospedajes, estadías, refrigerios y alimentación serán provistos y costeados por Howden exclusivamente para esta oferta;",
            "Airfare, transfers, local transportation, accommodation, lodging, snacks and meals will be arranged and paid by Howden exclusively for this quotation;"),
        ["ger.materia-prima"] = (
            "Howden no considera el suministro de materia prima y/o accesorios;",
            "Howden does not include the supply of raw material and/or accessories;"),
        ["ger.politica-viagens"] = (
            "En caso de que la empresa contratante asuma los gastos, deberá seguir la política de viajes y estadías adoptada por Howden;",
            "Should the contracting company cover the expenses, it must follow the travel and accommodation policy adopted by Howden;"),
        ["ger.periculosidade"] = (
            "En caso de trabajos en áreas clasificadas con peligrosidad, se agregará una tasa del 30% al valor ofertado por técnico;",
            "For work in hazardous classified areas, a 30% surcharge will be added to the price quoted per technician;"),
        ["ger.subcontratacao"] = ("La ejecución de los servicios está sujeta a subcontratación;",
                                  "Performance of the services is subject to subcontracting;"),
        ["ger.legislacao-tributaria"] = (
            "<b>Legislación tributaria y reforma tributaria:</b> informamos que esta oferta comercial fue elaborada con base en la legislación tributaria vigente a la fecha de su emisión. Cualquier cambio derivado de la reforma tributaria o de la legislación vigente podrá impactar los valores y condiciones aquí presentados. Por ello, nos reservamos el derecho de revisar y ajustar esta oferta según sea necesario. Agradecemos su comprensión y quedamos a disposición para aclaraciones adicionales;",
            "<b>Tax legislation and tax reform:</b> this commercial quotation was prepared based on the tax legislation in force on its issue date. Any change arising from tax reform or applicable legislation may affect the prices and conditions presented herein. We therefore reserve the right to review and adjust this quotation as required. We appreciate your understanding and remain available for further clarification;"),
        ["ger.kyc"] = (
            "Este presupuesto u oferta está sujeto a la conclusión satisfactoria de nuestros procedimientos habituales de conocimiento del cliente y de cumplimiento (KYC). Las condiciones contractuales finales se acordarán posteriormente por escrito;",
            "This estimate or quotation is subject to the satisfactory completion of our usual know-your-customer and compliance (KYC) procedures. Final contractual terms will be subsequently agreed in writing;"),
        ["ger.prazo-90dias"] = (
            "<b>Los servicios deberán ejecutarse en un plazo máximo de 90 días después de la recepción de la Orden de Compra.</b>",
            "<b>The services must be performed within a maximum of 90 days after receipt of the Purchase Order.</b>"),
    };

    /// <summary>
    /// Texto de uma nota no idioma da proposta. Devolve null quando não há
    /// tradução — aí vale o texto em português.
    /// </summary>
    public static string? Nota(string chave, string idioma)
    {
        if (idioma == "Português" || !NotasTraduzidas.TryGetValue(chave, out var t)) return null;
        return idioma == "English" ? t.En : t.Es;
    }

    /// <summary>Títulos das seções de notas no idioma da proposta.</summary>
    public static (string Incluso, string Excluso, string Gerais) TitulosNotas(string idioma) => idioma switch
    {
        "English" => ("INCLUDED IN HOWDEN'S SUPPLY:", "EXCLUDED FROM HOWDEN'S SUPPLY:", "General notes:"),
        "Español" => ("INCLUIDO EN EL SUMINISTRO DE HOWDEN:", "EXCLUIDO DEL SUMINISTRO DE HOWDEN:", "Notas generales:"),
        _ => ("INCLUSO NO FORNECIMENTO DA HOWDEN:", "EXCLUSOS DO FORNECIMENTO DA HOWDEN:", "Notas Gerais:"),
    };
}
