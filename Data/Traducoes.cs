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

        // ---- descrição das atividades de assessoria e padrões de escopo (Data/Topicos.cs) ----
        ["DESCRICAO DAS ATIVIDADES DE ASSESSORIA"] = ("Descripción de las Actividades de Asesoría", "Description of the Technical Assistance Activities"),
        ["INSPECAO INICIAL DO EQUIPAMENTO"] = ("Inspección inicial del equipo", "Initial inspection of the equipment"),
        ["VERIFICACAO DO ESTADO GERAL DE TODOS OS COMPONENTES"] = ("Verificación del estado general de todos los componentes", "Check of the general condition of all components"),
        ["INSPECAO VISUAL DOS MANCAIS"] = ("Inspección visual de los descansos", "Visual inspection of the bearings"),
        ["INSPECAO VISUAL DO ACOPLAMENTO"] = ("Inspección visual del acoplamiento", "Visual inspection of the coupling"),
        ["INSPECAO DOS COLARINHOS FLEXIVEIS"] = ("Inspección de las juntas flexibles", "Inspection of the flexible connectors"),
        ["INSPECAO E TESTE DE ABERTURA E FECHAMENTO DO DAMPER"] = ("Inspección y prueba de apertura y cierre del damper", "Inspection and opening/closing test of the damper"),
        ["ABERTURA DAS PORTAS DE INSPECAO PARA ACESSO INTERNO AO VENTILADOR"] = ("Apertura de las puertas de inspección para acceso interno al ventilador", "Opening of the inspection doors for internal access to the fan"),
        ["ABERTURA DO BIPARTIDO DO MANCAL MONOBLOCO PARA INSPECAO INTERNA"] = ("Apertura del bipartido del descanso monobloque para inspección interna", "Opening of the split housing of the monobloc bearing for internal inspection"),

        ["ASSESSORIA DE MONTAGEM"] = ("Asesoría de Montaje", "Assembly Assistance"),
        ["APOIO AO PLANEJAMENTO E PREPARACAO DAS ATIVIDADES DE CAMPO"] = ("Apoyo a la planificación y preparación de las actividades de campo", "Support for the planning and preparation of field activities"),
        ["ACOMPANHAMENTO DA MANUTENCAO DOS COMPONENTES"] = ("Acompañamiento del mantenimiento de los componentes", "Follow-up of the maintenance of the components"),
        ["ACOMPANHAMENTO DA DESMONTAGEM DOS COMPONENTES EXISTENTES"] = ("Acompañamiento del desmontaje de los componentes existentes", "Follow-up of the disassembly of the existing components"),
        ["ACOMPANHAMENTO DA MONTAGEM MECANICA DOS EQUIPAMENTOS E ACESSORIOS"] = ("Acompañamiento del montaje mecánico de los equipos y accesorios", "Follow-up of the mechanical assembly of the equipment and accessories"),
        ["ACOMPANHAMENTO DO ALINHAMENTO E POSICIONAMENTO DOS CONJUNTOS GIRANTES"] = ("Acompañamiento de la alineación y posicionamiento de los conjuntos rotativos", "Follow-up of the alignment and positioning of the rotating assemblies"),
        ["ACOMPANHAMENTO DA INSTALACAO DOS ELEMENTOS DE FIXACAO E RESPECTIVOS TORQUES DE APERTO"] = ("Acompañamiento de la instalación de los elementos de fijación y sus torques de apriete", "Follow-up of the installation of the fasteners and their tightening torques"),
        ["ACOMPANHAMENTO DA INSTALACAO DE SISTEMAS AUXILIARES, QUANDO APLICAVEL"] = ("Acompañamiento de la instalación de sistemas auxiliares, cuando corresponda", "Follow-up of the installation of auxiliary systems, when applicable"),
        ["INTERFACE TECNICA COM A EQUIPE DE MONTAGEM PARA ESCLARECIMENTO DE DUVIDAS E ORIENTACAO DE BOAS PRATICAS"] = ("Interfaz técnica con el equipo de montaje para aclarar dudas y orientar sobre buenas prácticas", "Technical interface with the assembly team to clarify doubts and provide guidance on good practices"),

        ["COMISSIONAMENTO E START-UP"] = ("Comisionamiento y Puesta en Marcha", "Commissioning and Start-up"),
        ["INSPECAO FINAL DE MONTAGEM MECANICA"] = ("Inspección final del montaje mecánico", "Final inspection of the mechanical assembly"),
        ["VERIFICACAO DIMENSIONAL E VISUAL DOS COMPONENTES INSTALADOS"] = ("Verificación dimensional y visual de los componentes instalados", "Dimensional and visual check of the installed components"),
        ["CONFERENCIA DO APERTO E TRAVAMENTO DOS ELEMENTOS DE FIXACAO"] = ("Verificación del apriete y bloqueo de los elementos de fijación", "Check of the tightening and locking of the fasteners"),
        ["VERIFICACAO DOS NIVEIS E CONDICOES DE LUBRIFICACAO"] = ("Verificación de los niveles y condiciones de lubricación", "Check of the lubrication levels and conditions"),
        ["VERIFICACAO DAS INTERLIGACOES ELETRICAS E DE INSTRUMENTACAO RELACIONADAS AO EQUIPAMENTO"] = ("Verificación de las interconexiones eléctricas y de instrumentación relacionadas con el equipo", "Check of the electrical and instrumentation interconnections related to the equipment"),
        ["INSPECAO DOS SISTEMAS DE PROTECAO E SEGURANCA DO EQUIPAMENTO"] = ("Inspección de los sistemas de protección y seguridad del equipo", "Inspection of the protection and safety systems of the equipment"),
        ["ACOMPANHAMENTO DO TESTE DE GIRO LIVRE E CONFIRMACAO DA AUSENCIA DE INTERFERENCIAS MECANICAS"] = ("Acompañamiento de la prueba de giro libre y confirmación de la ausencia de interferencias mecánicas", "Follow-up of the free rotation test and confirmation that there is no mechanical interference"),
        ["VERIFICACAO DO SENTIDO DE ROTACAO DO CONJUNTO GIRANTE"] = ("Verificación del sentido de rotación del conjunto rotativo", "Check of the direction of rotation of the rotating assembly"),
        ["ACOMPANHAMENTO DA PARTIDA A FRIO DO EQUIPAMENTO"] = ("Acompañamiento de la partida en frío del equipo", "Follow-up of the cold start of the equipment"),
        ["MONITORAMENTO DAS CONDICOES OPERACIONAIS DURANTE A ELEVACAO GRADUAL DA CARGA"] = ("Monitoreo de las condiciones operacionales durante el aumento gradual de la carga", "Monitoring of the operating conditions during the gradual load increase"),
        ["MEDICAO E ANALISE DOS NIVEIS DE VIBRACAO"] = ("Medición y análisis de los niveles de vibración", "Measurement and analysis of the vibration levels"),
        ["MONITORAMENTO DE TEMPERATURA DE MANCAIS, MOTOR E DEMAIS COMPONENTES CRITICOS"] = ("Monitoreo de temperatura de descansos, motor y demás componentes críticos", "Temperature monitoring of bearings, motor and other critical components"),
        ["VERIFICACAO DA ESTABILIDADE OPERACIONAL DO EQUIPAMENTO"] = ("Verificación de la estabilidad operacional del equipo", "Check of the operational stability of the equipment"),
        ["ORIENTACAO OPERACIONAL A EQUIPE DO CLIENTE QUANTO AS MELHORES PRATICAS DE OPERACAO"] = ("Orientación operacional al equipo del cliente sobre las mejores prácticas de operación", "Operational guidance to the customer's team on best operating practices"),
        ["REGISTRO DE OCORRENCIAS E RECOMENDACOES PARA OTIMIZACAO DA OPERACAO"] = ("Registro de ocurrencias y recomendaciones para la optimización de la operación", "Record of occurrences and recommendations for optimizing operation"),

        // ---- entregáveis, recursos da Howden e infraestrutura do contratante (Data/Topicos.cs) ----
        ["RELATORIO DOS SERVICOS EXECUTADOS (COM DESCRICAO DOS SERVICOS, FOTOS ETC.)"] = ("Informe de los servicios ejecutados (con descripción de los servicios, fotos, etc.)", "Report of the services performed (with description of the services, photos, etc.)"),
        ["RELATORIO DE HORAS DE ATENDIMENTO"] = ("Informe de horas de atención", "Timesheet report"),
        ["MATERIAL DIDATICO DE TREINAMENTO (FORMATO ELETRONICO), QUANDO CONTRATADO ESTE SERVICO"] = ("Material didáctico de capacitación (formato electrónico), cuando se contrate este servicio", "Training material (electronic format), when this service is contracted"),
        ["CASO SEJAM SOLICITADOS DOCUMENTOS ADICIONAIS AOS LISTADOS ACIMA, APOS AS NEGOCIACOES E/OU FECHAMENTO DO CONTRATO, SERAO COBRADOS COMO ADICIONAIS DE ESCOPO"] = ("En caso de que se soliciten documentos adicionales a los listados arriba, después de las negociaciones y/o del cierre del contrato, serán cobrados como adicionales de alcance", "Should documents in addition to those listed above be requested after the negotiations and/or the signing of the contract, they will be charged as scope additions"),
        ["ESPECIALISTA HOWDEN PARA ASSESSORIA TECNICA"] = ("Especialista HOWDEN para Asesoría Técnica", "HOWDEN specialist for Technical Assistance"),
        ["EPIS OBRIGATORIOS (CAPACETE, OCULOS, LUVAS, PROTETOR AUDITIVO, CALCADO DE SEGURANCA ETC.)"] = ("EPP obligatorios (casco, lentes, guantes, protector auditivo, calzado de seguridad, etc.)", "Mandatory PPE (helmet, goggles, gloves, hearing protection, safety shoes, etc.)"),
        ["INSTRUMENTACAO DE MEDICAO DE VIBRACAO E ALINHAMENTO (QUANDO CONTRATADO)"] = ("Instrumentación de medición de vibración y alineación (cuando se contrate)", "Vibration and alignment measuring instruments (when contracted)"),

        ["INFRAESTRUTURA E RECURSOS – RESPONSABILIDADE CONTRATANTE"] = ("Infraestructura y Recursos – Responsabilidad del CONTRATANTE", "Infrastructure and Resources – CUSTOMER's Responsibility"),
        ["FERRAMENTAL GERAL PARA ATIVIDADE"] = ("Herramental general para la actividad", "General tools for the activity"),
        ["JOGO DE FERRAMENTAS COMPLETO"] = ("Juego de herramientas completo", "Complete tool set"),
        ["CHAVE DE IMPACTO E ACESSORIOS"] = ("Llave de impacto y accesorios", "Impact wrench and accessories"),
        ["TORQUIMETRO DE ESTALO OU SIMILAR COM ACESSORIOS"] = ("Torquímetro de chasquido o similar con accesorios", "Click-type torque wrench or similar with accessories"),
        ["RELOGIO COMPARADOR COM BASE MAGNETICA"] = ("Reloj comparador con base magnética", "Dial indicator with magnetic base"),
        ["ALINHADOR A LASER"] = ("Alineador láser", "Laser alignment tool"),
        ["OUTRAS FERRAMENTAS QUE O ESPECIALISTA HOWDEN JULGAR NECESSARIO"] = ("Otras herramientas que el Especialista Howden considere necesarias", "Other tools deemed necessary by the Howden Specialist"),
        ["FERRAMENTAL GERAL PARA CORTE E SOLDA DA CARCACA (SE APLICAVEL)"] = ("Herramental general para corte y soldadura de la carcasa (SI CORRESPONDE)", "General tools for cutting and welding the casing (IF APPLICABLE)"),
        ["CONJUNTO DE OXICORTE COMPOSTO POR MACARICO DE CORTE OXICOMBUSTIVEL, CILINDROS/REGULADORES DE OXIGENIO E GAS COMBUSTIVEL, MANGUEIRAS E ACESSORIOS DE SEGURANCA PARA EXECUCAO DOS CORTES NA CARCACA DO VENTILADOR"] = ("Equipo de oxicorte compuesto por soplete de corte oxicombustible, cilindros/reguladores de oxígeno y gas combustible, mangueras y accesorios de seguridad para ejecutar los cortes en la carcasa del ventilador", "Oxy-fuel cutting set consisting of an oxy-fuel cutting torch, oxygen and fuel gas cylinders/regulators, hoses and safety accessories for cutting the fan casing"),
        ["MAQUINA DE SOLDA PARA REMOCAO DE SOLDAS EXISTENTES E PREPARACAO DE JUNTAS"] = ("Máquina de soldar para remoción de soldaduras existentes y preparación de juntas", "Welding machine for removal of existing welds and joint preparation"),
        ["ESMERILHADEIRAS ANGULARES DE DIVERSOS PORTES, EQUIPADAS COM DISCOS DE CORTE, DESBASTE E FLAP PARA ACABAMENTO E PREPARACAO DAS SUPERFICIES"] = ("Esmeriladoras angulares de diversos tamaños, equipadas con discos de corte, desbaste y flap para acabado y preparación de superficies", "Angle grinders of various sizes, fitted with cutting, grinding and flap discs for finishing and surface preparation"),
        ["FURADEIRAS ELETRICAS E FERRAMENTAS DE CORTE AUXILIARES PARA EXECUCAO DE FUROS, REMOCAO DE INTERFERENCIAS E ADEQUACAO DIMENSIONAL"] = ("Taladros eléctricos y herramientas de corte auxiliares para ejecución de perforaciones, remoción de interferencias y ajuste dimensional", "Electric drills and auxiliary cutting tools for drilling, removal of interferences and dimensional adjustment"),
        ["EQUIPAMENTOS DE MEDICAO E CONTROLE DIMENSIONAL, INCLUINDO TRENAS, REGUAS, NIVEIS, ESQUADROS, PAQUIMETROS E INSTRUMENTOS DE ALINHAMENTO"] = ("Equipos de medición y control dimensional, incluyendo huinchas, reglas, niveles, escuadras, pies de metro e instrumentos de alineación", "Dimensional measuring and control equipment, including tape measures, rulers, levels, squares, calipers and alignment instruments"),
        ["DISPOSITIVOS DE MARCACAO, GABARITOS E INSTRUMENTOS DE TRACAGEM PARA DELIMITACAO DAS AREAS DE CORTE E CONTROLE GEOMETRICO"] = ("Dispositivos de marcación, plantillas e instrumentos de trazado para delimitación de las áreas de corte y control geométrico", "Marking devices, templates and layout instruments for delimiting the cutting areas and geometric control"),
        ["ANDAIMES"] = ("Andamios", "Scaffolding"),
        ["PLATAFORMA ELEVATORIA"] = ("Plataforma elevadora", "Aerial work platform"),
        ["CAMINHAO MUNCK, GUINDASTES COM ACESSORIOS"] = ("Camión pluma, grúas con accesorios", "Truck-mounted crane, cranes with accessories"),
        ["CAMINHAO PARA MOVIMENTACAO INTERNA"] = ("Camión para movimiento interno", "Truck for internal handling"),
        ["CONTAINER ALMOXARIFADO, ADAPTADO PARA BANHEIROS E VESTIARIOS"] = ("Contenedor bodega, adaptado para baños y camarines", "Storage container, adapted for restrooms and changing rooms"),
        ["POSTO DE TRABALHO COM INTERNET, IMPRESSORA E OUTROS"] = ("Puesto de trabajo con internet, impresora y otros", "Workstation with internet, printer and others"),
        ["INSTALACOES ADEQUADAS, COMO BANHEIROS, REFEITORIOS, PARA USO DO ESPECIALISTA, PROMOVENDO UM AMBIENTE DE TRABALHO DIGNO E FUNCIONAL"] = ("Instalaciones adecuadas, como baños y comedores, para uso del especialista, promoviendo un ambiente de trabajo digno y funcional", "Suitable facilities, such as restrooms and canteens, for the specialist's use, providing a decent and functional work environment"),
        ["EPCS, PCMSO, PGR, PERMISSAO DE TRABALHO E OUTROS DOCUMENTOS DE SEGURANCA"] = ("EPC, PCMSO, PGR, Permiso de Trabajo y otros documentos de seguridad", "Collective protective equipment, PCMSO, PGR, Work Permit and other safety documents"),
        ["O TRANSPORTE AEREO POR HELICOPTERO SERA DE RESPONSABILIDADE DA CONTRATANTE, CONTEMPLANDO O DESLOCAMENTO ENTRE A BASE OPERACIONAL E A PLATAFORMA. A HOWDEN NAO FORNECERA ESTE MEIO DE TRANSPORTE ESPECIFICO"] = ("El transporte aéreo en helicóptero será responsabilidad del CONTRATANTE, incluyendo el traslado entre la base operacional y la plataforma. HOWDEN no proporcionará este medio de transporte específico", "Helicopter air transport will be the CUSTOMER's responsibility, including travel between the operational base and the platform. HOWDEN will not provide this specific means of transport"),
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
