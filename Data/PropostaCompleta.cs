using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HowdenServicos.Poc.Models;

namespace HowdenServicos.Poc.Data;

/// <summary>
/// O SEGUNDO PADRÃO da proposta: a proposta completa, PREENCHENDO o modelo em
/// Word da equipe (<c>Recursos/proposta-completa.docx</c>, o mesmo modelo
/// "Proposta Técnica Comercial" do AFM_HSA). O arquivo que sai é o de vocês —
/// capa, logos, estilos, cabeçalho e rodapé — e daqui só entram os dados.
///
/// <para>
/// O modelo é escrito para ser preenchido À MÃO e traz alternativas com a
/// instrução de escolher ou apagar. Gerar a proposta é escolher e APAGAR. Na
/// proposta de serviço ficam: capa, contatos, introdução, preço, notas
/// comerciais, impostos, pagamento, prazo, validade, faturamento, notas,
/// sistema de gestão, Assessoria Técnica no Campo (as diárias), Descrição da
/// Assessoria, Anexos e Comentários. As partes de ventilador (Proposta
/// Técnica, limites, itens e exclusos de fornecimento, desenho, tolerâncias,
/// Coanda, padrões e mancais) e o local de entrega saem.
/// </para>
///
/// <para>
/// As ferramentas de preenchimento (Escrever, EscreverNaCelula, Achar,
/// ApagarEntre, os campos ligados, a tradução) são as do AFM_HSA
/// (PropostaWordPreenchimento), que vieram do axiaisvent.
/// </para>
/// </summary>
public static class PropostaCompleta
{
    public static string CaminhoDoModelo =>
        Path.Combine(AppContext.BaseDirectory, "Recursos", "proposta-completa.docx");

    /// <summary>O que o Word precisa, já calculado pela tela.</summary>
    /// <param name="Base">o documento do pricing (antes da apresentação): é dele que sai a diária sem impostos e sem despesas</param>
    /// <param name="Apresentado">o documento como o cliente vê (o mesmo da proposta simples)</param>
    /// <param name="ContatoDoRepresentante">o "Contato" do cadastro do representante, pelo nome</param>
    public sealed record Dados(
        Proposta P, PricingParams Params, Pricing.Documento Base, Pricing.Documento Apresentado,
        BillingInfo? Faturamento, Func<string, string?> ContatoDoRepresentante);

    /// <summary>"Proposta-HSAXXX.AFM.000123-rev0.docx" (Quote / Propuesta no inglês e no espanhol).</summary>
    public static string NomeDoArquivo(Proposta p, string extensao)
    {
        var prefixo = p.Idioma switch { "English" => "Quote", "Español" => "Propuesta", _ => "Proposta" };
        var numero = string.IsNullOrWhiteSpace(p.Numero) ? "rascunho" : p.Numero.Trim();
        var nome = $"{prefixo}-{numero}-rev{PropostaWordOu(p.Revisao, "0")}.{extensao}";
        foreach (var c in Path.GetInvalidFileNameChars()) nome = nome.Replace(c, '_');
        return nome;
    }

    public static byte[] Gerar(Dados dados, string? caminhoDoModelo = null)
    {
        var caminho = caminhoDoModelo ?? CaminhoDoModelo;
        var fluxo = new MemoryStream();
        using (var modelo = File.OpenRead(caminho)) modelo.CopyTo(fluxo);
        fluxo.Position = 0;
        using (var doc = WordprocessingDocument.Open(fluxo, true))
        {
            var termos = TermosDasBusDeFora.Ler(Path.GetDirectoryName(Path.GetFullPath(caminho)) ?? ".");
            new PreenchimentoCompleto(doc, dados, termos).Preencher();
        }
        return fluxo.ToArray();
    }

    /// <summary>O mesmo Word, em PDF (Word do computador ou LibreOffice — ver <see cref="PdfDaProposta"/>).</summary>
    public static byte[] GerarPdf(Dados dados) => PdfDaProposta.Converter(Gerar(dados));

    internal static string PropostaWordOu(string? valor, string padrao) =>
        string.IsNullOrWhiteSpace(valor) ? padrao : valor.Trim();
}

/// <summary>O preenchimento do modelo, seção por seção (ver <see cref="PropostaCompleta"/>).</summary>
internal sealed class PreenchimentoCompleto
{
    private readonly WordprocessingDocument _doc;
    private readonly PropostaCompleta.Dados _d;
    private readonly Proposta _p;
    private readonly Pricing.Documento _a;
    private readonly Body _corpo;
    private readonly string _idioma;
    private readonly IReadOnlyList<TermosDasBusDeFora.Bloco> _termos;
    private static readonly CultureInfo Br = new("pt-BR");

    /// <summary>Chile/Peru (ou sem imposto calculado): a proposta mostra só o valor líquido.</summary>
    private readonly bool _semImpostos;

    /// <summary>BU de fora do Brasil (HCHL / HPU): notas a mais e os Términos y Condiciones no fim.</summary>
    private bool DeFora => Servicos.PaisDaBu(_p.Bu) != "Brasil";

    public PreenchimentoCompleto(WordprocessingDocument doc, PropostaCompleta.Dados dados, IReadOnlyList<TermosDasBusDeFora.Bloco> termos)
    {
        _doc = doc;
        _d = dados;
        _p = dados.P;
        _a = dados.Apresentado;
        _termos = termos;
        _corpo = doc.MainDocumentPart!.Document.Body!;
        _idioma = IdiomasDaProposta.Normalizar(_p.Idioma);
        var c = _a.Calculo;
        _semImpostos = c.Pis + c.Cofins + c.Iss <= 0.005;
    }

    public void Preencher()
    {
        Propriedades();
        CamposLigados();
        Capa();
        EnderecoDaBu();
        Contatos();
        Representante();
        TirarPaginaDeConsulta();
        Revisoes();
        Introducao();
        Precos();
        NotasComerciais();
        Impostos();
        Pagamento();
        Prazo();
        TirarLocalDeEntrega();
        Validade();
        Faturamento();
        AssessoriaNoCampo();
        TirarParteTecnica();
        DescricaoDaAssessoria();
        InformacoesAdicionais();
        TermosNoFim();
        Indice();
        Traduzir();
        TirarRealce();

        _doc.MainDocumentPart!.Document.Save();
    }

    // ================= textos e valores =================

    /// <summary>O texto no idioma da proposta.</summary>
    private string T3(string pt, string en, string es) => _idioma switch
    {
        IdiomasDaProposta.Ingles => en,
        IdiomasDaProposta.Espanhol => es,
        _ => pt,
    };

    /// <summary>Nome vindo da tabela de custos (português), no idioma da proposta.</summary>
    private string Item(string? texto) => Traducoes.Item(texto, _p.Idioma);

    private string M(double v) => $"{Servicos.Simbolo(_p.Moeda)} {Pricing.Moeda(v)}";

    /// <summary>Valor dentro de uma frase: no inglês, com ponto decimal (a tabela é convertida na tradução).</summary>
    private string Mx(double v) => _idioma == IdiomasDaProposta.Ingles
        ? $"{Servicos.Simbolo(_p.Moeda)} {v.ToString("#,##0.00", CultureInfo.InvariantCulture)}" : M(v);

    private static string Qtd(double v) => v.ToString("0.##", Br);

    private string Pct(double v) => v.ToString("0.##", _idioma == IdiomasDaProposta.Ingles ? CultureInfo.InvariantCulture : Br) + "%";

    private string Data() => Servicos.FmtData(_p.Data);

    private string Referencia() => Servicos.NumeroCompleto(_p);

    // ================= propriedades e campos ligados =================

    private void Propriedades()
    {
        var props = _doc.PackageProperties;
        props.Title = Referencia();
        props.Category = _p.Projeto;
        props.Description = _p.ContatoNome;
        props.Keywords = _p.Referencia;
        props.Creator = _p.PreparadaPor;

        if (_doc.ExtendedFilePropertiesPart?.Properties is { } extendidas)
        {
            extendidas.Company ??= new DocumentFormat.OpenXml.ExtendedProperties.Company();
            extendidas.Company.Text = _p.Cliente;
            extendidas.Save();
        }

        foreach (var parte in _doc.MainDocumentPart!.CustomXmlParts)
        {
            string xml;
            using (var leitura = parte.GetStream(FileMode.Open, FileAccess.Read))
            using (var leitor = new StreamReader(leitura))
                xml = leitor.ReadToEnd();
            if (!xml.Contains("CoverPageProperties")) continue;

            // o campo de data lê a propriedade como data (storeMappedDataAs=dateTime): vai em ISO
            xml = TrocarTag(xml, "PublishDate", DateTime.TryParse(_p.Data, out var dia)
                ? dia.ToString("yyyy-MM-dd'T'00:00:00", CultureInfo.InvariantCulture) : "");
            xml = TrocarTag(xml, "CompanyAddress", Cidade());

            using var escrita = parte.GetStream(FileMode.Create, FileAccess.Write);
            using var escritor = new StreamWriter(escrita);
            escritor.Write(xml);
        }
    }

    /// <summary>Cidade da proposta; no Brasil, com o estado ("Itatiba - SP").</summary>
    private string Cidade()
    {
        var cidade = (_p.Cidade ?? "").Trim();
        if (_p.Moeda != "BRL" || Servicos.PaisDaBu(_p.Bu) != "Brasil") return cidade;
        var uf = (_p.Estado ?? "").Trim();
        return cidade.Length == 0 ? uf : uf.Length == 0 ? cidade : $"{cidade} - {uf}";
    }

    private static string TrocarTag(string xml, string tag, string valor)
    {
        var seguro = System.Security.SecurityElement.Escape(valor ?? "") ?? "";
        var vazia = $"<{tag}/>";
        if (xml.Contains(vazia)) return xml.Replace(vazia, $"<{tag}>{seguro}</{tag}>");
        var abre = xml.IndexOf($"<{tag}>", StringComparison.Ordinal);
        var fecha = xml.IndexOf($"</{tag}>", StringComparison.Ordinal);
        if (abre < 0 || fecha < abre) return xml;
        return xml[..(abre + tag.Length + 2)] + seguro + xml[fecha..];
    }

    /// <summary>
    /// Os campos da capa, do cabeçalho e da primeira nota, ligados às
    /// propriedades. Sem valor o campo SAI — senão o Word volta a mostrar o
    /// exemplo, e a proposta iria para o cliente escrito "[Categoria]".
    /// </summary>
    private void CamposLigados()
    {
        var valores = new Dictionary<string, string>
        {
            ["Empresa"] = _p.Cliente,
            ["Categoria"] = _p.Projeto,
            ["Comentários"] = _p.ContatoNome,
            ["Título"] = Referencia(),
            ["Endereço da Empresa"] = Cidade(),
            ["Palavras-chave"] = _p.Referencia,
            ["Data de Publicação"] = Data(),
        };

        foreach (var parte in Partes())
        {
            foreach (var sdt in parte.Descendants<SdtElement>().ToList())
            {
                var alias = sdt.SdtProperties?.GetFirstChild<SdtAlias>()?.Val?.Value;
                if (alias is null || !valores.TryGetValue(alias, out var valor)) continue;

                if ((valor ?? "").Trim().Length == 0)
                {
                    Desembrulhar(sdt);
                    continue;
                }

                sdt.SdtProperties?.GetFirstChild<ShowingPlaceholder>()?.Remove();
                // a data fica como texto: sem o vínculo, o Word e o LibreOffice mostram o que foi escrito aqui
                if (alias == "Data de Publicação") sdt.SdtProperties?.GetFirstChild<DataBinding>()?.Remove();
                foreach (var cinza in sdt.Descendants<Color>().ToList()) cinza.Remove();

                var textos = sdt.Descendants<Text>().ToList();
                if (textos.Count == 0) continue;
                textos[0].Text = valor!;
                textos[0].Space = SpaceProcessingModeValues.Preserve;
                foreach (var sobra in textos.Skip(1)) sobra.Remove();
            }
            parte.Save();
        }
    }

    private static void Desembrulhar(SdtElement sdt)
    {
        var conteudo = sdt.ChildElements.FirstOrDefault(e =>
            e is SdtContentBlock or SdtContentRun or SdtContentCell or SdtContentRow);
        if (conteudo is null || sdt.Parent is null) { sdt.Remove(); return; }

        foreach (var texto in conteudo.Descendants<Text>().ToList()) texto.Text = "";
        foreach (var filho in conteudo.ChildElements.ToList())
        {
            filho.Remove();
            sdt.Parent.InsertBefore(filho, sdt);
        }
        sdt.Remove();
    }

    private static void Abrir(SdtElement sdt)
    {
        var conteudo = sdt.ChildElements.FirstOrDefault(e =>
            e is SdtContentBlock or SdtContentRun or SdtContentCell or SdtContentRow);
        if (conteudo is null || sdt.Parent is null) { sdt.Remove(); return; }
        foreach (var filho in conteudo.ChildElements.ToList())
        {
            filho.Remove();
            sdt.Parent.InsertBefore(filho, sdt);
        }
        sdt.Remove();
    }

    // ================= capa =================

    private Table Capa0() => Tabelas().First();

    private void Capa()
    {
        var capa = Capa0();
        Rotulo(capa, "E-mail:", _p.ContatoEmail);
        Rotulo(capa, "Fone:", _p.ContatoTelefone);
        MesmaFonteDoCliente(capa);
        Rotulo(capa, "Preparada por:", _p.PreparadaPor);

        // a faixa azul: Howden Chile / Howden Perú (a HSA fica com Howden South America)
        if (DeFora && capa.Descendants<Paragraph>().FirstOrDefault(p => Texto(p).Trim() == "Howden South America") is { } faixa)
            Escrever(faixa, Servicos.NomeCurtoBu(_p.Bu));
    }

    private void MesmaFonteDoCliente(Table capa)
    {
        var empresa = capa.Descendants<SdtElement>()
            .FirstOrDefault(s => s.SdtProperties?.GetFirstChild<SdtAlias>()?.Val?.Value == "Empresa");
        var molde = empresa?.Descendants<Run>().FirstOrDefault(r => r.Elements<Text>().Any(t => t.Text.Trim().Length > 0))?.RunProperties;
        if (molde is null) return;

        var alvos = new List<OpenXmlElement>();
        foreach (var rotulo in new[] { "E-mail:", "Fone:", "Preparada por:" })
            if (Celula(capa, rotulo)?.NextSibling<TableCell>() is { } valor) alvos.Add(valor);
        alvos.AddRange(capa.Descendants<SdtElement>()
            .Where(s => s.SdtProperties?.GetFirstChild<SdtAlias>()?.Val?.Value is
                "Comentários" or "Endereço da Empresa" or "Categoria" or "Título" or "Data de Publicação" or "Palavras-chave"));
        foreach (var alvo in alvos)
            foreach (var run in alvo.Descendants<Run>().Where(r => r.Elements<Text>().Any()))
            {
                run.RunProperties?.Remove();
                run.PrependChild((RunProperties)molde.CloneNode(true));
            }
    }

    private static void Rotulo(Table tabela, string rotulo, string valor)
    {
        if (Celula(tabela, rotulo) is not { } celula) return;
        if (celula.NextSibling<TableCell>() is { } seguinte)
            EscreverNaCelula(seguinte, new[] { valor ?? "" });
    }

    /// <summary>O endereço da BU: o modelo traz os quatro; fica o da BU da proposta e o "Web:".</summary>
    private void EnderecoDaBu()
    {
        var marcas = new Dictionary<string, string>
        {
            ["HSA-SP"] = "Osvaldo Berto", ["HSA-ES"] = "Rua 4E", ["HCHL"] = "Cordillera", ["HPU"] = "Guillermo Marconi",
        };
        var escolhida = marcas.GetValueOrDefault((_p.Bu ?? "").Trim(), "Osvaldo Berto");

        var celula = Capa0().Descendants<TableCell>()
            .FirstOrDefault(c => Texto(c).Contains("Osvaldo Berto", StringComparison.Ordinal));
        if (celula is null) return;

        var blocos = new List<List<Paragraph>> { new() };
        foreach (var p in celula.Elements<Paragraph>())
        {
            if (Texto(p).Trim().Length == 0) blocos.Add(new());
            else blocos[^1].Add(p);
        }
        foreach (var bloco in blocos)
        {
            var texto = string.Join(" ", bloco.Select(Texto));
            if (texto.Contains(escolhida, StringComparison.Ordinal) || texto.Contains("Web:", StringComparison.Ordinal)) continue;
            foreach (var p in bloco) p.Remove();
        }

        var brancos = celula.Elements<Paragraph>().Where(p => Texto(p).Trim().Length == 0).ToList();
        var web = celula.Elements<Paragraph>().FirstOrDefault(p => Texto(p).Contains("Web:", StringComparison.Ordinal));
        foreach (var b in brancos) b.Remove();
        if (web is not null && brancos.Count > 0) celula.InsertBefore((Paragraph)brancos[0].CloneNode(true), web);
    }

    /// <summary>
    /// Os contatos Howden: na capa fica a linha do diretor de vendas e, ao lado
    /// dele, quem assina a proposta (o vendedor) — com o cartão que o modelo já
    /// tem dele, ou com os dados digitados na proposta.
    /// </summary>
    private void Contatos()
    {
        var capa = Capa0();
        var doDiretor = capa.Elements<TableRow>().FirstOrDefault(l => Texto(l).Contains("Geraldini", StringComparison.Ordinal));
        if (doDiretor is null) return;
        var celulas = doDiretor.Elements<TableCell>().ToList();
        if (celulas.Count < 2) return;
        var molde = celulas[^1];

        var linhas = LinhasDoVendedor();
        for (var i = 0; i < celulas.Count - 1; i++)
        {
            if (i == 0 && linhas.Count > 0) ComoOMolde(celulas[i], molde, linhas);
            else EscreverNaCelula(celulas[i], new[] { "" });
        }
    }

    private List<string> LinhasDoVendedor()
    {
        var nomeVendedor = (_p.AssinaNome ?? "").Trim();
        if (nomeVendedor.Length == 0) return new();
        // o mesmo diretor que já está na capa não se repete
        if (Textos.Simples(nomeVendedor).Contains("geraldini")) return new();

        var nome = Textos.Simples(nomeVendedor);
        var partes = nome.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var contatos = Tabelas().Skip(1).FirstOrDefault(t => Texto(t).Contains("@chartindustries.com", StringComparison.Ordinal));
        if (contatos is not null && partes.Length > 0)
        {
            foreach (var cel in contatos.Descendants<TableCell>())
            {
                var ps = cel.Elements<Paragraph>().Select(Texto).Where(t => t.Trim().Length > 0).ToList();
                if (ps.Count == 0) continue;
                var doCartao = Textos.Simples(ps[0]);
                if (partes.All(doCartao.Contains) || doCartao == nome) return ps;
            }
        }
        var fones = (_p.AssinaFones ?? "").Split(new[] { '/', '|', ';' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        return new[] { nomeVendedor, _p.AssinaCargo ?? "" }.Concat(fones).Append(_p.AssinaEmail ?? "")
            .Where(s => s.Trim().Length > 0).Select(s => s.Trim()).ToList();
    }

    private static void ComoOMolde(TableCell celula, TableCell molde, IReadOnlyList<string> linhas)
    {
        foreach (var antigo in celula.Elements<Paragraph>().ToList()) antigo.Remove();
        foreach (var p in molde.Elements<Paragraph>()) celula.Append((Paragraph)p.CloneNode(true));
        if (molde.TableCellProperties?.GetFirstChild<TableCellVerticalAlignment>() is { } alinhamento)
        {
            celula.TableCellProperties ??= new TableCellProperties();
            celula.TableCellProperties.GetFirstChild<TableCellVerticalAlignment>()?.Remove();
            celula.TableCellProperties.Append((TableCellVerticalAlignment)alinhamento.CloneNode(true));
        }
        EscreverNaCelula(celula, linhas);
    }

    /// <summary>
    /// O representante: fica só o desta proposta (Rep 1 e Rep 2), com o
    /// contato do cadastro — ou o que o modelo traz dele. Sem representante, a linha sai.
    /// </summary>
    private void Representante()
    {
        var celula = Capa0().Descendants<TableCell>()
            .FirstOrDefault(c => Texto(c).TrimStart().StartsWith("Representante:", StringComparison.Ordinal));
        if (celula is null) return;

        var reps = new[] { _p.Representante, _p.Representante2 }
            .Select(r => (r ?? "").Trim()).Where(r => r.Length > 0 && r != "-").Distinct().ToList();
        if (reps.Count == 0)
        {
            celula.Ancestors<TableRow>().FirstOrDefault()?.Remove();
            return;
        }

        var linhas = new List<string> { "Representante:" };
        linhas.AddRange(reps.Select(r => _d.ContatoDoRepresentante(r) is { Length: > 0 } c ? c : ContatoNoModelo(r)));
        EscreverNaCelula(celula, linhas);
    }

    private string ContatoNoModelo(string nome)
    {
        var chave = nome;
        foreach (var corte in new[] { " (", " <", " >" })
            if (chave.IndexOf(corte, StringComparison.Ordinal) is var i and > 0) chave = chave[..i];
        chave = Textos.Simples(chave);
        if (chave.Length < 3) return nome;

        var inicio = Achar("Representante:");
        var fim = Achar("Consulta REPRESENTANTES");
        if (inicio is null || fim is null) return nome;
        var linha = Entre(inicio, fim).OfType<Paragraph>()
            .FirstOrDefault(p => Textos.Simples(Texto(p)).Contains(chave));
        return linha is null ? nome : Texto(linha).Trim();
    }

    /// <summary>A página de CONSULTA (contatos e representantes) sai inteira, com a quebra de página dela.</summary>
    private void TirarPaginaDeConsulta()
    {
        var contatos = Tabelas().Skip(1).FirstOrDefault(t => Texto(t).Contains("@chartindustries.com", StringComparison.Ordinal));
        var consulta = Achar("Consulta REPRESENTANTES");
        if (contatos is null || consulta is null) return;

        var fim = _corpo.ChildElements.SkipWhile(e => e != consulta).Skip(1)
            .OfType<Paragraph>().FirstOrDefault(p => p.Descendants<Break>().Any(b => b.Type?.Value == BreakValues.Page));

        var apagar = new List<OpenXmlElement>();
        var dentro = false;
        foreach (var e in _corpo.ChildElements)
        {
            if (e == contatos) dentro = true;
            if (dentro) apagar.Add(e);
            if (e == (fim ?? (OpenXmlElement)consulta)) break;
        }
        foreach (var e in apagar) e.Remove();

        if (Capa0().NextSibling<Paragraph>() is { } quebra &&
            quebra.Descendants<Break>().Any(b => b.Type?.Value == BreakValues.Page) && Texto(quebra).Trim().Length == 0 &&
            quebra.NextSibling<Paragraph>() is { } seguinte)
        {
            quebra.Remove();
            seguinte.ParagraphProperties ??= new ParagraphProperties();
            seguinte.ParagraphProperties.PageBreakBefore = new PageBreakBefore();
        }
    }

    private void Revisoes()
    {
        var tabela = Tabelas().FirstOrDefault(t => Texto(t).StartsWith("Rev.", StringComparison.Ordinal));
        if (tabela is null) return;
        var linhas = tabela.Elements<TableRow>().ToList();
        if (linhas.Count < 2) return;

        var celulas = linhas[1].Elements<TableCell>().ToList();
        if (celulas.Count < 4) return;
        EscreverNaCelula(celulas[0], new[] { PropostaCompleta.PropostaWordOu(_p.Revisao, "0") });
        EscreverNaCelula(celulas[1], new[] { _p.PreparadaPor ?? "" });
        EscreverNaCelula(celulas[2], new[] { _p.RevisadaPor ?? "" });
        foreach (var vazia in linhas.Skip(2)) vazia.Remove();

        // a coluna "Descrição:" sai (como no AFM, pedido da equipe): a largura dela vai para "Exec"
        var colunas = tabela.GetFirstChild<TableGrid>()?.Elements<GridColumn>().ToList();
        if (colunas is { Count: 4 })
        {
            int N(StringValue? v) => int.TryParse(v?.Value, out var n) ? n : 0;
            colunas[1].Width = (N(colunas[1].Width) + N(colunas[3].Width)).ToString();
            colunas[3].Remove();
            foreach (var linha in tabela.Elements<TableRow>())
                if (linha.Elements<TableCell>().ToList() is { Count: 4 } cels) cels[3].Remove();
            var total = colunas.Take(3).Sum(c => N(c.Width));
            if (total > 0)
                foreach (var linha in tabela.Elements<TableRow>())
                {
                    var cels = linha.Elements<TableCell>().ToList();
                    if (cels.Count != 3) continue;
                    for (var i = 0; i < 3; i++)
                    {
                        var props = cels[i].TableCellProperties ??= new TableCellProperties();
                        props.TableCellWidth = new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = (N(colunas[i].Width) * 5000 / total).ToString() };
                    }
                }
        }
    }

    // ================= introdução e preço =================

    private void Introducao()
    {
        var contato = (_p.ContatoNome ?? "").Trim();
        if (Achar("Prezado Sr.") is { } saudacao)
            Escrever(saudacao, contato.Length > 0
                ? T3($"Prezado(a) Sr(a). {contato},", $"Dear {contato},", $"Estimado/a {contato},")
                : T3("Prezados Senhores,", "Dear Sirs,", "Estimados señores,"));

        var cliente = PropostaCompleta.PropostaWordOu(_p.Cliente, "—");
        var escopo = string.IsNullOrWhiteSpace(_p.EscopoServico) ? "" : " – " + Item(_p.EscopoServico);
        if (Achar("Howden tem a satisfação de apresentar") is { } apresentacao)
            Escrever(apresentacao, T3(
                $"Howden tem a satisfação de apresentar para {cliente} sua oferta para a prestação de serviços de assessoria técnica{escopo}.",
                $"Howden is pleased to present {cliente} with its proposal for technical assistance services{escopo}.",
                $"Howden tiene la satisfacción de presentar a {cliente} su oferta para la prestación de servicios de asesoría técnica{escopo}."));
    }

    private void Precos()
    {
        // não há proposta técnica de ventilador: a frase que apontava para ela sai
        Achar("Preços conforme descrição da Proposta Técnica")?.Remove();
        TabelaDePrecos();

        // a quebra de página depois da tabela sai: as Notas Comerciais seguem logo abaixo
        if (Achar("A descrição do produto deverá estar presente na Ordem de Compra") is { } nota
            && nota.NextSibling() is Paragraph seguinte && Texto(seguinte).Trim().Length == 0)
        {
            foreach (var quebra in seguinte.Descendants<Break>().Where(b => b.Type?.Value == BreakValues.Page).ToList()) quebra.Remove();
            foreach (var p in seguinte.Descendants<ParagraphProperties>()) p.RemoveAllChildren<PageBreakBefore>();
        }
    }

    /// <summary>Uma linha da tabela de preços: título, observação, quantidade, unitário e total.</summary>
    private sealed record LinhaDePreco(string Titulo, string Obs, double Qtd, double Total);

    /// <summary>
    /// As linhas do preço: as da assessoria (diárias e horas) e, mostrando as
    /// despesas, as despesas a custo + taxa; sem despesas, o deslocamento numa
    /// linha própria — o mesmo que a proposta simples mostra.
    /// </summary>
    private List<LinhaDePreco> LinhasDePreco()
    {
        var linhas = new List<LinhaDePreco>();
        foreach (var l in _a.MO.Where(l => l.ValorTotal > 0.004))
            linhas.Add(new(Item(l.Servico), Item(l.Obs), l.QtdDiaria > 0 ? l.QtdDiaria : 1, l.ValorTotal));
        foreach (var d in _a.Despesas.Where(d => d.ValorTotal > 0.004))
            linhas.Add(new(Item(d.Despesa), Item(d.Obs), d.Qtd > 0 ? d.Qtd : 1, d.ValorTotal));
        if (_a.Deslocamento > 0.004)
        {
            var l = Servicos.Labels(_p.Idioma);
            linhas.Add(new(l.Deslocamento, l.DeslocamentoObs, 1, _a.Deslocamento));
        }
        return linhas;
    }

    private void TabelaDePrecos()
    {
        var tabela = Tabelas().FirstOrDefault(t => Texto(t).StartsWith("Item", StringComparison.Ordinal) && Texto(t).Contains("Descrição"));
        if (tabela is null) return;
        var linhas = tabela.Elements<TableRow>().ToList();
        if (linhas.Count < 3) return;

        var cab = linhas[0].Elements<TableCell>().ToList();
        if (cab.Count >= 5)
        {
            EscreverNaCelula(cab[3], new[] { "Preço Unitário" });
            EscreverNaCelula(cab[4], new[] { "Preço Total" });
        }

        var total = linhas[^1];
        foreach (var sdt in total.Descendants<SdtElement>().ToList()) Abrir(sdt);
        var celTotal = total.Elements<TableCell>().ToList();
        if (celTotal.Count >= 2)
        {
            EscreverNaCelula(celTotal[0], new[]
            {
                _semImpostos ? T3("VALOR TOTAL", "TOTAL AMOUNT", "VALOR TOTAL")
                    : T3("VALOR TOTAL C/ PIS, COFINS E ISS", "TOTAL AMOUNT INCL. PIS, COFINS AND ISS", "VALOR TOTAL CON PIS, COFINS E ISS"),
            });
            EscreverNaCelula(celTotal[^1], new[] { M(_a.Total) });
            RotuloEmNegrito(celTotal[0], celTotal[^1]);
        }

        var modelo = (TableRow)linhas[1].CloneNode(true);
        modelo.TableRowProperties ??= new TableRowProperties();
        if (!modelo.TableRowProperties.Elements<CantSplit>().Any()) modelo.TableRowProperties.Append(new CantSplit());
        var itens = LinhasDePreco();
        for (var i = 0; i < itens.Count; i++)
        {
            var linha = (TableRow)modelo.CloneNode(true);
            PreencherItem(linha, i, itens[i]);
            tabela.InsertBefore(linha, total);
        }
        foreach (var antiga in linhas.Skip(1).Take(linhas.Count - 2)) antiga.Remove();
    }

    private void PreencherItem(TableRow linha, int indice, LinhaDePreco item)
    {
        var celulas = linha.Elements<TableCell>().ToList();
        if (celulas.Count < 5) return;

        EscreverNaCelula(celulas[0], new[] { (indice + 1).ToString("D2") });
        Descricao(celulas[1], item);
        EscreverNaCelula(celulas[2], new[] { Qtd(item.Qtd) });
        EscreverNaCelula(celulas[3], new[] { M(Math.Round(item.Total / item.Qtd, 2)) });
        EscreverNaCelula(celulas[4], new[] { M(item.Total) });
    }

    /// <summary>
    /// A descrição do item nos parágrafos do modelo (é deles que vem o negrito
    /// e o centro do título): o serviço no título e a observação embaixo. O
    /// resto (referências e escopo de peça, prazo, NCM) é de ventilador e sai.
    /// </summary>
    private static void Descricao(TableCell celula, LinhaDePreco item)
    {
        foreach (var sdt in celula.Elements<SdtElement>().ToList()) Abrir(sdt);
        var paragrafos = celula.Elements<Paragraph>().ToList();
        var titulo = paragrafos.FirstOrDefault(p => Texto(p).Contains("PREENCHER!!", StringComparison.Ordinal));
        var obs = paragrafos.FirstOrDefault(p => Texto(p).Trim() == "XXXXXX");
        foreach (var e in celula.ChildElements.Where(e => e is Paragraph or Table or SdtElement).ToList())
            if (e != titulo && e != obs) e.Remove();
        if (titulo is not null) Escrever(titulo, item.Titulo);
        if (obs is not null)
        {
            if (item.Obs.Trim().Length > 0) Escrever(obs, item.Obs);
            else obs.Remove();
        }
        if (!celula.Elements<Paragraph>().Any()) celula.Append(new Paragraph(new Run(new Text(item.Titulo))));
    }

    private static void RotuloEmNegrito(TableCell rotulo, TableCell valor)
    {
        var cor = valor.Descendants<Run>().FirstOrDefault(r => r.Elements<Text>().Any(t => t.Text.Trim().Length > 0))
            ?.RunProperties?.GetFirstChild<Color>();
        foreach (var run in rotulo.Descendants<Run>().Where(r => r.Elements<Text>().Any()))
        {
            run.RunProperties ??= new RunProperties();
            var rp = run.RunProperties;
            if (rp.GetFirstChild<Bold>() is null) rp.Append(new Bold());
            if (rp.GetFirstChild<BoldComplexScript>() is null) rp.Append(new BoldComplexScript());
            rp.GetFirstChild<Color>()?.Remove();
            rp.Append(cor is not null ? (Color)cor.CloneNode(true) : new Color { Val = "000000" });
            NaOrdem(rp);
        }
    }

    private static readonly Type[] OrdemDoRPr =
    {
        typeof(RunStyle), typeof(RunFonts), typeof(Bold), typeof(BoldComplexScript), typeof(Italic), typeof(ItalicComplexScript),
        typeof(Caps), typeof(SmallCaps), typeof(Strike), typeof(DoubleStrike), typeof(Outline), typeof(Shadow), typeof(Emboss),
        typeof(Imprint), typeof(NoProof), typeof(SnapToGrid), typeof(Vanish), typeof(WebHidden), typeof(Color), typeof(Spacing),
        typeof(CharacterScale), typeof(Kern), typeof(Position), typeof(FontSize), typeof(FontSizeComplexScript), typeof(Highlight),
        typeof(Underline), typeof(TextEffect), typeof(Border), typeof(Shading), typeof(FitText), typeof(VerticalTextAlignment),
        typeof(RightToLeftText), typeof(ComplexScript), typeof(Emphasis), typeof(Languages),
    };

    private static void NaOrdem(RunProperties rp)
    {
        var filhos = rp.ChildElements.ToList();
        int Posicao(OpenXmlElement e) => Array.IndexOf(OrdemDoRPr, e.GetType()) is var i and >= 0 ? i : OrdemDoRPr.Length;
        var ordenados = filhos.Select((e, i) => (e, i)).OrderBy(x => Posicao(x.e)).ThenBy(x => x.i).Select(x => x.e).ToList();
        if (ordenados.SequenceEqual(filhos)) return;
        foreach (var f in filhos) f.Remove();
        foreach (var f in ordenados) rp.Append(f);
    }

    /// <summary>HCHL / HPU: as Notas Comerciais começam com as duas notas a mais da equipe de lá.</summary>
    private void NotasComerciais()
    {
        if (!DeFora || _corpo.Descendants<Paragraph>().FirstOrDefault(p => Texto(p).Contains("Excluso de nosso fornecimento qualquer item", StringComparison.Ordinal)) is not { } primeira)
            return;
        foreach (var texto in new[]
        {
            "Os preços e condições considerados nesta oferta aplicam-se somente se forem comprados todos os itens aqui ofertados. Caso contrário, os preços e condições podem estar sujeitos a alterações.",
            "Não está incluído no escopo qualquer documento associado para realizar a importação (certificado EUR1, catálogos, ficha técnica etc.). Caso necessário, solicitar junto com o pedido ou com a ordem de compra para ser incorporado ao pedido.",
        })
        {
            var nota = (Paragraph)primeira.CloneNode(true);
            Escrever(nota, texto);
            primeira.InsertBeforeSelf(nota);
        }
    }

    // ================= impostos =================

    /// <summary>
    /// Impostos Vigentes de serviço: PIS, COFINS e ISS, inclusos no preço (as
    /// alíquotas da proposta). ICMS e IPI são de produto e saem. Chile/Peru:
    /// só o valor líquido, sem a lei brasileira e sem o detalhamento.
    /// </summary>
    private void Impostos()
    {
        var pis = Achar("PIS: 1,65%");
        var cofins = Achar("COFINS: 7,6%");
        var iss = Achar("ISS: 2%");
        var apagar = new List<Paragraph?>
        {
            Achar("ITATIBA:"), Achar("ICMS: (18% - SP)"), Achar("ICMS: (12% - MG"), Achar("ICMS: (7% - INDICAR"), Achar("VITORIA:"),
            Achar("ICMS: (12% - INDICAR"), Achar("ICMS: (17% - ES)"), Achar("IPI: 0%"), Achar("IPI: 3,25%"), Achar("IPI: 9,75%"),
            Achar("ISS: 5%"), Achar("Fundamento Legal"),
        };

        if (_semImpostos)
        {
            var empresa = Servicos.NomeCurtoBu(_p.Bu);
            if (pis is not null) Escrever(pis, T3(
                $"Os preços cotados não incluem nenhum tipo de imposto ou retenção; são os valores líquidos a serem pagos à {empresa}.",
                $"The quoted prices do not include any kind of tax or withholding; they are the net amounts to be paid to {empresa}.",
                $"Los precios cotizados no incluyen ningún tipo de impuesto o retención; son los valores netos a pagar a {empresa}."));
            apagar.AddRange(new[] { cofins, iss });
            apagar.AddRange(new[]
            {
                Achar("Nota sobre Legislação Tributária"), Achar("A presente proposta comercial foi elaborada com base na legislação tributária"),
                Achar("Dessa forma, a Howden South America se reserva o direito de revisar"),
            });
        }
        else
        {
            var (aPis, aCofins, aIss) = Servicos.AliquotasDoc(_a);
            var incluso = T3("incluso", "included", "incluido");
            if (pis is not null) Escrever(pis, $"PIS: {Pct(aPis)} {incluso};");
            if (cofins is not null) Escrever(cofins, $"COFINS: {Pct(aCofins)} {incluso};");
            if (iss is not null)
            {
                if (aIss > 0.004) Escrever(iss, $"ISS: {Pct(aIss)} {incluso};");
                else apagar.Add(iss);
            }
        }
        foreach (var p in apagar) p?.Remove();

        // os parágrafos vazios que separavam ICMS/IPI ficam sobrando entre os impostos
        if (Titulo2("Impostos Vigentes") is { } titulo)
            foreach (var vazio in _corpo.ChildElements.SkipWhile(e => e != titulo).Skip(1)
                         .TakeWhile(e => e is Paragraph p && Texto(p).Trim() is var t && !t.StartsWith("Nota sobre") && !t.StartsWith("Detalhamento"))
                         .OfType<Paragraph>().Where(p => Texto(p).Trim().Length == 0 && !p.Descendants<Break>().Any()).ToList())
                vazio.Remove();

        DetalhamentoDosImpostos();
    }

    /// <summary>"Detalhamento para adição dos Impostos": líquido, com PIS/COFINS e com PIS/COFINS/ISS (= o total).</summary>
    private void DetalhamentoDosImpostos()
    {
        var tabela = Tabelas().FirstOrDefault(t => Texto(t).StartsWith("Valor Total Líquido", StringComparison.Ordinal));
        if (tabela is null) return;
        if (_semImpostos)
        {
            Achar("Detalhamento para adição dos Impostos")?.Remove();
            tabela.Remove();
            return;
        }
        var linhas = tabela.Elements<TableRow>().ToList();
        if (linhas.Count < 4) return;

        var (semImp, comPisCofins) = Pricing.ResumoDoTotal(_a);
        void Valor(TableRow linha, double v) => EscreverNaCelula(linha.Elements<TableCell>().Last(), new[] { M(v) });
        Valor(linhas[0], semImp);
        Valor(linhas[1], comPisCofins);
        EscreverNaCelula(linhas[2].Elements<TableCell>().First(), new[]
        {
            T3("Valor Total c/ PIS, COFINS e ISS:", "Total Value incl. PIS, COFINS and ISS:", "Valor Total c/ PIS, COFINS e ISS:"),
        });
        Valor(linhas[2], _a.Total);
        foreach (var l in linhas.Skip(3)) l.Remove();
    }

    // ================= pagamento, prazo, entrega e validade =================

    /// <summary>
    /// O modelo traz as faixas de pagamento por valor; fica só a condição
    /// desta proposta (a escolhida na tela, no idioma da proposta). Cada "•"
    /// vira uma linha com o marcador do modelo.
    /// </summary>
    private void Pagamento()
    {
        var faixa = Achar("Propostas com valor < USD 75 mil");
        var escolhida = Achar("100% – 30 ddl da retirada");
        var fim = Ultimo("aprovação do data book");
        if (escolhida is null || fim is null) return;
        ApagarEntre(escolhida, fim);
        faixa?.Remove();

        var texto = Servicos.PagamentoNoIdioma(_p.CondicaoPagamento, _p.Idioma);
        var partes = texto.Replace("\r\n", "\n").Split('\n', '•')
            .Select(l => l.Trim()).Where(l => l.Length > 0).ToList();
        if (partes.Count == 0) partes.Add(T3("A combinar.", "To be agreed.", "A convenir."));

        Escrever(escolhida, partes[0]);
        var depois = escolhida;
        foreach (var parte in partes.Skip(1))
        {
            var p = (Paragraph)escolhida.CloneNode(true);
            Escrever(p, parte);
            depois.InsertAfterSelf(p);
            depois = p;
        }
    }

    private void Prazo()
    {
        if (Achar("Prazos conforme mencionado a cada item") is not { } p) return;
        var dias = (_p.PrazoEntregaDias ?? "").Trim();
        Escrever(p, dias.Length == 0
            ? T3("A combinar.", "To be agreed.", "A convenir.")
            : T3($"{dias} dias, com início após a confirmação do pedido e o esclarecimento de todos os dados técnicos e comerciais.",
                 $"{dias} days, starting after confirmation of the order and clarification of all technical and commercial data.",
                 $"{dias} días, contados a partir de la confirmación del pedido y de la aclaración de todos los datos técnicos y comerciales."));
    }

    /// <summary>Local de entrega (incoterms) é de equipamento: a seção sai inteira.</summary>
    private void TirarLocalDeEntrega()
    {
        var titulo = Paragrafos().FirstOrDefault(p => Texto(p).Trim().StartsWith("Local de Entrega", StringComparison.Ordinal));
        var fim = Achar("Equipamento fornecido desmontado");
        if (titulo is null || fim is null) return;
        ApagarEntre(titulo, fim);
        titulo.Remove();
    }

    private void Validade()
    {
        if (Achar("Esta proposta é válida por") is { } texto)
            Escrever(texto, $"Esta proposta é válida por {PropostaCompleta.PropostaWordOu(_p.ValidadeDias, "30")} dias.");
    }

    // ================= dados para faturamento =================

    /// <summary>
    /// "Dados para faturamento": a empresa da BU (tela Faturamento) no lugar
    /// da que o modelo traz, e a conta bancária. Sem banco cadastrado, a caixa sai.
    /// </summary>
    private void Faturamento()
    {
        var fat = _d.Faturamento ?? Servicos.FaturamentoPadrao(_p.Bu);
        var razao = Achar("Howden South America Ventiladores e Compressores");
        var endereco = Achar("Av. Osvaldo Berto, 475");
        var registro = Achar("CNPJ: 01.094.363");
        var banco = Tabelas().FirstOrDefault(t => Texto(t).TrimStart().StartsWith("Banco:", StringComparison.Ordinal));

        if (razao is not null) Escrever(razao, PropostaCompleta.PropostaWordOu(fat.Razao, Servicos.RazaoSocialBu(_p.Bu)));
        if (endereco is not null)
        {
            if ((fat.Endereco ?? "").Trim().Length > 0) Escrever(endereco, fat.Endereco!.Trim());
            else endereco.Remove();
        }
        if (registro is not null)
        {
            if ((fat.Registro ?? "").Trim().Length > 0) Escrever(registro, fat.Registro!.Trim());
            else registro.Remove();
        }
        if (banco is null) return;
        if ((fat.BancoNome ?? "").Trim().Length == 0) { banco.Remove(); return; }

        // "Banco: X" em negrito e o resto normal, como no modelo
        var p = banco.Descendants<Paragraph>().FirstOrDefault(x => Texto(x).Contains("Banco:", StringComparison.Ordinal));
        if (p is null) return;
        var runs = p.Elements<Run>().Where(r => r.Elements<Text>().Any()).ToList();
        var negrito = runs.FirstOrDefault()?.RunProperties?.CloneNode(true) as RunProperties;
        var normal = runs.FirstOrDefault(r => r.RunProperties?.Bold is null)?.RunProperties?.CloneNode(true) as RunProperties
                     ?? (negrito?.CloneNode(true) as RunProperties);
        if (normal is not null) { normal.Bold = null; normal.BoldComplexScript = null; }
        p.RemoveAllChildren<Run>();
        Run R(RunProperties? rp, string t)
        {
            var r = new Run();
            if (rp is not null) r.Append(rp.CloneNode(true));
            r.Append(new Text(t) { Space = SpaceProcessingModeValues.Preserve });
            return r;
        }
        p.Append(R(negrito, $"{T3("Banco", "Bank", "Banco")}: {fat.BancoNome!.Trim()}"));
        var resto = "";
        if ((fat.Agencia ?? "").Trim().Length > 0) resto += $" – {T3("Agência", "Branch", "Agencia")}: {fat.Agencia!.Trim()}";
        if ((fat.Conta ?? "").Trim().Length > 0) resto += $" {T3("Conta", "Account", "Cuenta")}: {fat.Conta!.Trim()}";
        if (resto.Length > 0) p.Append(R(normal, resto));
    }

    // ================= assessoria técnica no campo (as diárias) =================

    /// <summary>
    /// A diária normal SEM impostos e SEM despesas, na moeda da proposta: a
    /// parte da mão de obra no preço (antes de embutir despesas), tirados PIS,
    /// COFINS e ISS — como o modelo pede ("sem impostos e sem despesas").
    /// </summary>
    private double DiariaSemImpostos()
    {
        var c = _d.Base.Calculo;
        var bruta = Pricing.DiariaNormalApresentada(_d.Base);
        if (bruta <= 0 || c.ComImpostos <= 0) return 0;
        var liquida = bruta * c.VendaLiquida / c.ComImpostos;
        var fx = Servicos.PrecisaConverter(_p, _d.Params)
            ? Pricing.CambioEfetivo(Pricing.TaxaCambioNormalizada(_d.Params), Pricing.Num(_d.Params.SegurancaCambioPct)) : 1;
        return Math.Round(liquida / fx, 2);
    }

    /// <summary>De onde o técnico sai, pela BU.</summary>
    private string LocalDeSaida() => _p.Bu switch
    {
        "HSA-ES" => "Serra/ES",
        "HCHL" => "Santiago (Chile)",
        "HPU" => T3("Lima (Peru)", "Lima (Peru)", "Lima (Perú)"),
        _ => "Itatiba/SP",
    };

    private void AssessoriaNoCampo()
    {
        var diaUtil = Achar("Preço por dia úteis das 08:00");
        var fimDeSemana = Achar("Preço por dia (Sábado, Domingos e Feriados)");
        var diaria = DiariaSemImpostos();
        if (diaria > 0)
        {
            var sufixo = T3("sem impostos e sem despesas", "excluding taxes and expenses", "sin impuestos y sin gastos");
            if (diaUtil is not null) Escrever(diaUtil, T3(
                $"Preço por dia útil das 08:00 às 17:00: {Mx(diaria)} {sufixo}",
                $"Price per working day from 8:00 a.m. to 5:00 p.m.: {Mx(diaria)} {sufixo}",
                $"Precio por día hábil de 08:00 a 17:00: {Mx(diaria)} {sufixo}"));
            if (fimDeSemana is not null) Escrever(fimDeSemana, T3(
                $"Preço por dia (Sábados, Domingos e Feriados): {Mx(2 * diaria)} {sufixo}",
                $"Price per day (Saturdays, Sundays and Holidays): {Mx(2 * diaria)} {sufixo}",
                $"Precio por día (Sábados, Domingos y Feriados): {Mx(2 * diaria)} {sufixo}"));
        }
        else
        {
            diaUtil?.Remove();
            fimDeSemana?.Remove();
        }

        var local = LocalDeSaida();
        foreach (var p in Paragrafos().Where(p => Texto(p).Contains("Local de saída para execução dos serviços", StringComparison.Ordinal)))
            Escrever(p, T3($"Local de saída para execução dos serviços: é sempre considerado nossa fábrica {local}.",
                           $"Departure point for the services: always considered to be our factory in {local}.",
                           $"Lugar de salida para la ejecución de los servicios: se considera siempre nuestra fábrica de {local}."));

        var adm = Pricing.Moeda0(Pricing.Num(_d.Params.TaxaAdmPct));
        if (Achar("Todas as despesas de viagem (passagem aérea") is { } despesas)
            Escrever(despesas, T3(
                $"Todas as despesas de viagem (passagem aérea, hotel, transporte, táxi, alimentação, certificações, exames de saúde especiais e outros) estão EXCLUÍDAS nos preços acima indicados. Se o comprador solicitar que as despesas sejam por conta da Howden, as despesas serão acrescidas de taxa administrativa ({adm}%) somadas no valor computado. Em caso de contratação das despesas pela empresa contratante, essa deverá seguir a política de viagens e estadias seguindo o padrão adotado pela Howden;",
                $"All travel expenses (airfare, hotel, transport, taxi, meals, certifications, special medical exams and others) are EXCLUDED from the prices above. If the buyer requests that the expenses be borne by Howden, an administrative fee ({adm}%) will be added to the amount computed. If the expenses are contracted by the contracting company, it shall follow the travel and accommodation policy according to the standard adopted by Howden;",
                $"Todos los gastos de viaje (pasaje aéreo, hotel, transporte, taxi, alimentación, certificaciones, exámenes de salud especiales y otros) están EXCLUIDOS de los precios arriba indicados. Si el comprador solicita que los gastos sean por cuenta de Howden, se agregará una tasa administrativa ({adm}%) al valor computado. En caso de contratación de los gastos por la empresa contratante, esta deberá seguir la política de viajes y estadías según el estándar adoptado por Howden;"));

        // ISS é do Brasil
        if (DeFora) Achar("Imposto: ISS a ser incluído")?.Remove();
    }

    // ================= escopo: só a assessoria =================

    private Paragraph? Titulo1(string texto) => Paragrafos().FirstOrDefault(p =>
        Texto(p).Trim() == texto &&
        (p.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "").Contains("Ttulo1", StringComparison.OrdinalIgnoreCase));

    private Paragraph? Titulo2(string inicio) => Paragrafos().FirstOrDefault(p =>
        Texto(p).Trim().StartsWith(inicio, StringComparison.Ordinal) &&
        (p.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "").Contains("TTULO2", StringComparison.OrdinalIgnoreCase));

    /// <summary>O número da seção, num parágrafo próprio logo acima do título ("2", "3"...).</summary>
    private static Paragraph? NumeroDe(Paragraph titulo) =>
        titulo.PreviousSibling<Paragraph>() is { } p && Texto(p).Trim() is { Length: > 0 } t && t.All(char.IsDigit) ? p : null;

    /// <summary>
    /// A Proposta Técnica (dados de ventilador, curvas, construção, ruído) sai
    /// inteira; do Escopo de Fornecimento saem limite de bateria, condições de
    /// processo, itens inclusos, documentação, itens opcionais, exclusos e o
    /// desenho — fica a Descrição da Assessoria Técnica.
    /// </summary>
    private void TirarParteTecnica()
    {
        var tecnica = Titulo1("Proposta Técnica");
        var escopo = Titulo1("Escopo de Fornecimento");
        if (tecnica is not null && escopo is not null)
        {
            OpenXmlElement de = NumeroDe(tecnica) ?? tecnica;
            OpenXmlElement ate = NumeroDe(escopo) ?? escopo;
            foreach (var e in Entre(de, ate)) e.Remove();
            de.Remove();
        }

        var limite = Titulo2("Limite de bateria");
        var descricao = Titulo2("Descrição - Assessoria Técnica");
        if (limite is not null && descricao is not null)
        {
            foreach (var e in Entre(limite, descricao)) e.Remove();
            limite.Remove();
        }

        var opcionais = Titulo2("Itens Opcionais");
        var adicionais = Titulo1("Informações Adicionais");
        if (opcionais is not null && adicionais is not null)
        {
            OpenXmlElement ate = NumeroDe(adicionais) ?? adicionais;
            foreach (var e in Entre(opcionais, ate)) e.Remove();
            opcionais.Remove();
            // Informações Adicionais continua começando numa página nova
            if (ate is Paragraph numero)
            {
                numero.ParagraphProperties ??= new ParagraphProperties();
                numero.ParagraphProperties.PageBreakBefore = new PageBreakBefore();
            }
        }
    }

    /// <summary>
    /// "Descrição - Assessoria Técnica": o escopo detalhado digitado na
    /// proposta e o que está incluso — técnicos e as linhas de diárias/horas
    /// (as mesmas quantidades do preço). As notas repetidas da seção (que já
    /// estão em Assessoria Técnica no Campo) saem.
    /// </summary>
    private void DescricaoDaAssessoria()
    {
        var titulo = Titulo2("Descrição - Assessoria Técnica");
        var intro = Achar("Estamos considerando incluso em nosso escopo");
        var tecnicos = Achar("XXXXX técnico/técnicos");
        var diarias = Achar("XXXXX total de diárias de assessoria técnica");
        var sub = Achar("XXXXX diárias de deslocamento");
        var ultima = Achar("Despesas de transporte, estadia e alimentação a cargo de V.Sas");
        if (titulo is null || intro is null || tecnicos is null || diarias is null || sub is null) return;

        // as notas do fim da seção (de "Preço para trabalhos por dia útil" até as despesas) saem
        var viagens = Achar("XXXXX viagens para o total de diárias inclusas");
        if (ultima is not null) ApagarEntre(viagens ?? sub, ultima);
        viagens?.Remove();
        Achar("XXXXX dias de integração")?.Remove();

        var n = Math.Max(Pricing.Inteiro(_d.Params.QtdTecnicos), 1);
        Escrever(tecnicos, T3($"{n} {(n == 1 ? "técnico" : "técnicos")}", $"{n} {(n == 1 ? "technician" : "technicians")}",
                              $"{n} {(n == 1 ? "técnico" : "técnicos")}"));

        var linhas = _a.MO.Where(l => l.QtdDiaria > 0).ToList();
        if (linhas.Count == 0)
        {
            diarias.Remove();
            sub.Remove();
        }
        else
        {
            Escrever(diarias, T3("Diárias e horas de assessoria técnica, considerando:", "Technical assistance days and hours, considering:",
                                 "Días y horas de asesoría técnica, considerando:"));
            OpenXmlElement depois = sub;
            foreach (var l in linhas)
            {
                var obs = Item(l.Obs).Trim();
                var p = (Paragraph)sub.CloneNode(true);
                Escrever(p, $"{Qtd(l.QtdDiaria)} × {Item(l.Servico).TrimEnd(':')}{(obs.Length > 0 ? $" – {obs}" : "")}");
                depois.InsertAfterSelf(p);
                depois = p;
            }
            sub.Remove();
        }

        // o escopo detalhado digitado na proposta, logo abaixo do título
        var escopo = (_p.EscopoDetalhado ?? "").Replace("\r\n", "\n").Split('\n').Select(l => l.TrimEnd()).ToList();
        while (escopo.Count > 0 && escopo[^1].Length == 0) escopo.RemoveAt(escopo.Count - 1);
        if (escopo.Count > 0)
        {
            var moldeTexto = (Paragraph)intro.CloneNode(true);
            OpenXmlElement depois = titulo;
            foreach (var linha in escopo)
            {
                var p = (Paragraph)moldeTexto.CloneNode(true);
                Escrever(p, linha);
                foreach (var r in p.Elements<Run>()) { if (r.RunProperties is { } rp) { rp.Bold = null; rp.BoldComplexScript = null; } }
                depois.InsertAfterSelf(p);
                depois = p;
            }
            // um espaço entre o escopo e o "Estamos considerando incluso"
            intro.ParagraphProperties ??= new ParagraphProperties();
            intro.ParagraphProperties.SpacingBetweenLines ??= new SpacingBetweenLines();
            intro.ParagraphProperties.SpacingBetweenLines.Before = "160";
        }
    }

    /// <summary>
    /// Informações Adicionais: em Anexos fica a T&amp;C (que a nota comercial
    /// cita como anexa) e ficam os Comentários e Desvios; o resto (LTSA,
    /// desenhos, FD, documentos do cliente, tolerâncias ISO 13348, Coanda,
    /// padrões, mancais) é de ventilador e sai.
    /// </summary>
    private void InformacoesAdicionais()
    {
        var ltsa = Achar("LTSA - Long Term");
        var comentarios = Titulo2("Comentários e Desvios");
        if (ltsa is not null && comentarios is not null)
        {
            foreach (var e in Entre(ltsa, comentarios)) e.Remove();
            ltsa.Remove();
        }
        var tolerancias = Titulo2("Tolerâncias De Performance");
        if (tolerancias is not null)
        {
            foreach (var e in _corpo.ChildElements.SkipWhile(e => e != tolerancias).ToList())
            {
                if (e is SectionProperties) continue;
                if (e is Paragraph p && p.ParagraphProperties?.SectionProperties is not null)
                {
                    foreach (var filho in p.ChildElements.Where(c => c is not ParagraphProperties).ToList()) filho.Remove();
                    continue;
                }
                e.Remove();
            }
        }
    }

    // ================= HCHL / HPU: os Términos y Condiciones no fim =================

    private void TermosNoFim()
    {
        if (!DeFora || _termos.Count == 0) return;
        var ancora = _corpo.ChildElements.LastOrDefault(e => e is SectionProperties);
        var quebra = new Paragraph(new Run(new Break { Type = BreakValues.Page }));
        if (ancora is not null) ancora.InsertBeforeSelf(quebra); else _corpo.Append(quebra);
        char anterior = ' ';
        foreach (var b in _termos)
        {
            var p = ParagrafoDosTermos(b, anterior);
            if (ancora is not null) ancora.InsertBeforeSelf(p); else _corpo.Append(p);
            anterior = b.Tipo;
        }
    }

    private static Paragraph ParagrafoDosTermos(TermosDasBusDeFora.Bloco b, char anterior)
    {
        const int Texto = 17, Secao = 18, Titulo = 22;
        Run R(string t, bool negrito = false, int tamanho = Texto, string? cor = null)
        {
            var rp = new RunProperties();
            if (negrito) { rp.Append(new Bold()); rp.Append(new BoldComplexScript()); }
            if (cor is not null) rp.Append(new Color { Val = cor });
            rp.Append(new FontSize { Val = tamanho.ToString(CultureInfo.InvariantCulture) });
            rp.Append(new FontSizeComplexScript { Val = tamanho.ToString(CultureInfo.InvariantCulture) });
            return new Run(rp, new Text(t) { Space = SpaceProcessingModeValues.Preserve });
        }
        Run Tab()
        {
            var r = R("");
            r.RemoveAllChildren<Text>();
            r.Append(new TabChar());
            return r;
        }
        Paragraph P(int antes, int depois, int esquerda, int pendurado, bool juntoAoProximo, JustificationValues alinhar, params Run[] runs)
        {
            var pp = new ParagraphProperties();
            if (juntoAoProximo) pp.Append(new KeepNext());
            pp.Append(new SpacingBetweenLines { Before = antes.ToString(CultureInfo.InvariantCulture), After = depois.ToString(CultureInfo.InvariantCulture) });
            if (esquerda > 0 || pendurado > 0)
                pp.Append(new Indentation { Left = esquerda.ToString(CultureInfo.InvariantCulture), Hanging = pendurado > 0 ? pendurado.ToString(CultureInfo.InvariantCulture) : null });
            pp.Append(new Justification { Val = alinhar });
            var p = new Paragraph(pp);
            foreach (var r in runs) p.Append(r);
            return p;
        }

        switch (b.Tipo)
        {
            case 'T':
                return P(anterior == ' ' ? 360 : 480, 160, 0, 0, true, JustificationValues.Left, R(b.Texto, true, Titulo, "1F4E79"));
            case 'S':
                return P(200, 80, 0, 0, true, JustificationValues.Left, R(b.Texto, true, Secao));
            case 'D':
                return P(0, 100, 567, 0, false, JustificationValues.Both, R(b.Termo + ": ", true), R(b.Texto));
            case 'I':
            {
                var m = System.Text.RegularExpressions.Regex.Match(b.Texto, @"^(\([ivx]+\))\s+(.*)$");
                return m.Success
                    ? P(0, 100, 1134, 567, false, JustificationValues.Both, R(m.Groups[1].Value), Tab(), R(m.Groups[2].Value))
                    : P(0, 100, 1134, 0, false, JustificationValues.Both, R(b.Texto));
            }
            case 'B':
                return P(0, 60, 1134, 283, false, JustificationValues.Both, R("–"), Tab(), R(b.Texto));
            default:
            {
                var m = System.Text.RegularExpressions.Regex.Match(b.Texto, @"^(\d{1,2}\.\d{1,2}\.?)\s+(.*)$");
                if (m.Success)
                    return P(0, 100, 567, 567, false, JustificationValues.Both, R(m.Groups[1].Value, true), Tab(), R(m.Groups[2].Value));
                return P(0, 100, anterior == 'T' ? 0 : 567, 0, false, JustificationValues.Both, R(b.Texto));
            }
        }
    }

    // ================= índice =================

    /// <summary>Os títulos que saíram (também do índice).</summary>
    private static readonly string[] TitulosQueSaem =
    {
        "Proposta Técnica", "Dados Técnicos", "Curva de Performance", "Curva de Torque", "Características Construtivas",
        "Nível de Ruído", "Limite de bateria", "Condições de processo", "Itens inclusos", "Documentação Incluída",
        "Itens Opcionais", "Exclusos do Fornecimento", "Desenho de linha de eixo", "Tolerâncias De Performance",
        "Efeito Coanda", "Padrões e Códigos", "Mancais Monobloco", "Local de Entrega",
    };

    /// <summary>
    /// O índice: as entradas do que saiu são tiradas, e o documento pede ao
    /// Word para atualizar os campos ao abrir — as páginas mudaram.
    /// </summary>
    private void Indice()
    {
        var toc = _corpo.Elements<SdtBlock>().FirstOrDefault(s => Texto(s).Contains("Proposta Comercial", StringComparison.Ordinal)
                                                                   && Texto(s).Contains("Introdução", StringComparison.Ordinal));
        if (toc?.SdtContentBlock is { } conteudo)
        {
            var entradas = conteudo.Elements<Paragraph>().ToList();
            foreach (var p in entradas.Skip(1).Take(Math.Max(0, entradas.Count - 2)))
            {
                var t = Texto(p).Trim();
                if (TitulosQueSaem.Any(tt => t.StartsWith(tt, StringComparison.Ordinal))) p.Remove();
            }
        }

        var settings = _doc.MainDocumentPart!.DocumentSettingsPart?.Settings;
        if (settings is not null)
        {
            settings.RemoveAllChildren<UpdateFieldsOnOpen>();
            var depois = settings.ChildElements.FirstOrDefault(e => e is HeaderShapeDefaults or FootnoteDocumentWideProperties
                or EndnoteDocumentWideProperties or Compatibility or DocumentVariables or Rsids
                or DocumentFormat.OpenXml.Math.MathProperties or AttachedSchema or ThemeFontLanguages or ColorSchemeMapping
                or DoNotIncludeSubdocsInStats or DoNotAutoCompressPictures or ForceUpgrade or Captions or ReadModeInkLockDown
                or ShapeDefaults or DecimalSymbol or ListSeparator);
            var atualizar = new UpdateFieldsOnOpen { Val = true };
            if (depois is not null) settings.InsertBefore(atualizar, depois);
            else settings.Append(atualizar);
            settings.Save();
        }
    }

    // ================= inglês e espanhol =================

    /// <summary>
    /// O texto fixo do modelo em inglês ou espanhol (<see cref="TraducaoDaProposta"/>).
    /// O que este preenchimento escreve já sai no idioma da proposta; o que
    /// não tem tradução (nome do cliente, texto digitado) fica como está.
    /// </summary>
    private void Traduzir()
    {
        if (_idioma == IdiomasDaProposta.Portugues) return;
        foreach (var parte in Partes())
        {
            // o modelo escreve alguns hífens como símbolo ("reserva‑se", "mão‑de‑obra"): viram texto,
            // senão sobram soltos depois da tradução ("financiero.----")
            foreach (var sym in parte.Descendants<SymbolChar>().Where(s => s.Char?.Value is "F02D" or "002D").ToList())
            {
                sym.InsertAfterSelf(new Text("-"));
                sym.Remove();
            }
            foreach (var p in parte.Descendants<Paragraph>().Where(p => !p.Descendants<Paragraph>().Any()).ToList())
            {
                var comCampo = p.Descendants<FieldChar>().Any() || p.Descendants<SimpleField>().Any();
                foreach (var trecho in Trechos(p))
                {
                    var texto = TextoDos(trecho);
                    if (TraducaoDaProposta.Traduzir(texto, _idioma, comCampo) is not { } traducao) continue;
                    if (traducao == texto.Trim()) continue;
                    Reescrever(trecho, texto, traducao);
                }
            }
            parte.Save();
        }
        Lingua();
    }

    private static List<List<Run>> Trechos(Paragraph p)
    {
        var trechos = new List<List<Run>>();
        var atual = new List<Run>();
        var campos = new Stack<bool>();
        void Fechar()
        {
            if (atual.Count > 0 && atual.Any(r => r.Elements<Text>().Any(t => t.Text.Length > 0))) trechos.Add(atual);
            atual = new List<Run>();
        }
        void Andar(OpenXmlElement pai)
        {
            foreach (var e in pai.ChildElements)
            {
                switch (e)
                {
                    case Run r:
                        if (r.GetFirstChild<FieldChar>() is { } fc)
                        {
                            Fechar();
                            var tipo = fc.FieldCharType?.Value;
                            if (tipo == FieldCharValues.Begin) campos.Push(true);
                            else if (tipo == FieldCharValues.Separate && campos.Count > 0) { campos.Pop(); campos.Push(false); }
                            else if (tipo == FieldCharValues.End && campos.Count > 0) campos.Pop();
                        }
                        else if (campos.Contains(true) || r.GetFirstChild<FieldCode>() is not null) Fechar();
                        else if (r.Elements<Text>().Any() || r.Elements<TabChar>().Any() || r.Elements<NoBreakHyphen>().Any()) atual.Add(r);
                        break;
                    case Hyperlink or InsertedRun or CustomXmlRun:
                        Andar(e);
                        break;
                    case SdtElement or SimpleField:
                        Fechar();
                        break;
                }
            }
        }
        Andar(p);
        Fechar();
        return trechos;
    }

    private static string TextoDos(List<Run> runs) => string.Concat(runs.SelectMany(r => r.ChildElements)
        .Select(e => e switch { Text t => t.Text, TabChar => "\t", NoBreakHyphen => "-", _ => "" }));

    private static void Reescrever(List<Run> trecho, string original, string traducao)
    {
        var partes = traducao.Split(TraducaoDaProposta.Parte);
        if (original.Length > 0 && char.IsWhiteSpace(original[0]) && !char.IsWhiteSpace(partes[0].FirstOrDefault(' ')))
            partes[0] = " " + partes[0];
        if (original.Length > 0 && original[^1] == ' ' && partes[^1].Length > 0 && partes[^1][^1] != ' ')
            partes[^1] += " ";

        var comTexto = trecho.Where(r => r.Elements<Text>().Any() || r.Elements<NoBreakHyphen>().Any()).ToList();
        var grupos = new List<List<Run>>();
        string? anterior = null;
        foreach (var r in comTexto)
        {
            var formato = Formato(r);
            if (grupos.Count == 0 || formato != anterior) grupos.Add(new List<Run>());
            grupos[^1].Add(r);
            anterior = formato;
        }
        if (grupos.Count == 0) return;
        if (partes.Length != grupos.Count)
        {
            partes = new[] { string.Concat(partes) };
            grupos = new List<List<Run>> { comTexto };
        }
        for (var i = 0; i < grupos.Count; i++)
        {
            var primeiro = true;
            foreach (var r in grupos[i])
            {
                foreach (var h in r.Elements<NoBreakHyphen>().ToList()) h.Remove();
                foreach (var h in r.Elements<SoftHyphen>().ToList()) h.Remove();
                var textos = r.Elements<Text>().ToList();
                if (primeiro && textos.Count == 0)
                {
                    r.Append(new Text());
                    textos = r.Elements<Text>().ToList();
                }
                foreach (var t in textos)
                {
                    if (primeiro)
                    {
                        t.Text = partes[i];
                        t.Space = SpaceProcessingModeValues.Preserve;
                        primeiro = false;
                    }
                    else t.Remove();
                }
            }
        }
    }

    private static string Formato(Run r)
    {
        var rp = r.RunProperties;
        if (rp is null) return "";
        return string.Join("|", new OpenXmlElement?[] { rp.Bold, rp.Italic, rp.Underline, rp.VerticalTextAlignment, rp.Color, rp.FontSize, rp.Highlight }
            .Select(e => e?.OuterXml ?? "-"));
    }

    private void Lingua()
    {
        var lingua = _idioma == IdiomasDaProposta.Ingles ? "en-US" : "es-ES";
        var raizes = new List<OpenXmlPartRootElement>(Partes());
        if (_doc.MainDocumentPart!.StyleDefinitionsPart?.Styles is { } estilos) raizes.Add(estilos);
        foreach (var raiz in raizes)
        {
            foreach (var l in raiz.Descendants<Languages>())
                if (l.Val?.Value is { } v && v.StartsWith("pt", StringComparison.OrdinalIgnoreCase)) l.Val = lingua;
            raiz.Save();
        }
    }

    /// <summary>Tira o realce de todo o documento: no modelo ele é a marca de "preencher/escolher".</summary>
    private void TirarRealce()
    {
        foreach (var parte in Partes())
        {
            foreach (var realce in parte.Descendants<Highlight>().ToList()) realce.Remove();
            parte.Save();
        }
    }

    // ================= ferramentas (as do AFM_HSA / axiaisvent) =================

    private List<OpenXmlPartRootElement> Partes()
    {
        var partes = new List<OpenXmlPartRootElement> { _doc.MainDocumentPart!.Document };
        partes.AddRange(_doc.MainDocumentPart.HeaderParts.Where(h => h.Header is not null).Select(h => (OpenXmlPartRootElement)h.Header));
        partes.AddRange(_doc.MainDocumentPart.FooterParts.Where(f => f.Footer is not null).Select(f => (OpenXmlPartRootElement)f.Footer));
        return partes;
    }

    private List<Paragraph> Paragrafos() => _corpo.Elements<Paragraph>().ToList();
    private List<Table> Tabelas() => _corpo.Elements<Table>().ToList();

    private Paragraph? Achar(string trecho) => Paragrafos()
        .FirstOrDefault(p => Texto(p).Contains(trecho, StringComparison.Ordinal));

    private Paragraph? Ultimo(string trecho) => Paragrafos()
        .LastOrDefault(p => Texto(p).Contains(trecho, StringComparison.Ordinal));

    private static TableCell? Celula(Table tabela, string rotulo) => tabela.Descendants<TableCell>()
        .FirstOrDefault(c => Texto(c).Trim() == rotulo.Trim());

    private static string Texto(OpenXmlElement e)
    {
        var partes = new List<string>();
        foreach (var n in e.Descendants())
        {
            if (n is Text t) partes.Add(t.Text);
            else if (n is TabChar) partes.Add(" ");
        }
        return string.Join("", partes);
    }

    /// <summary>Escreve o texto no parágrafo com a formatação do primeiro run que tinha texto; "\n" vira quebra de linha e "\t" tabulação.</summary>
    private static void Escrever(Paragraph p, string texto)
    {
        var modelo = p.Elements<Run>().FirstOrDefault(r => r.Elements<Text>().Any())
                     ?? p.Descendants<Run>().FirstOrDefault(r => r.Elements<Text>().Any())
                     ?? p.Elements<Run>().FirstOrDefault();
        var formato = modelo?.RunProperties?.CloneNode(true) as RunProperties;

        p.RemoveAllChildren<Run>();
        p.RemoveAllChildren<Hyperlink>();
        foreach (var campo in p.ChildElements.OfType<SdtElement>().ToList()) campo.Remove();

        var run = new Run();
        if (formato is not null) run.Append(formato);
        var primeiro = true;
        foreach (var linha in texto.Split('\n'))
        {
            if (!primeiro) run.Append(new Break());
            var pedacos = linha.Split('\t');
            for (var i = 0; i < pedacos.Length; i++)
            {
                if (i > 0) run.Append(new TabChar());
                run.Append(new Text(pedacos[i]) { Space = SpaceProcessingModeValues.Preserve });
            }
            primeiro = false;
        }
        p.Append(run);
    }

    private static void EscreverNaCelula(TableCell celula, IReadOnlyList<string> linhas)
    {
        var paragrafos = celula.Elements<Paragraph>().ToList();
        if (paragrafos.Count == 0)
        {
            celula.Append(new Paragraph());
            paragrafos = celula.Elements<Paragraph>().ToList();
        }
        for (var i = 0; i < linhas.Count; i++)
        {
            if (i < paragrafos.Count) { Escrever(paragrafos[i], linhas[i]); continue; }
            var novo = (Paragraph)paragrafos[^1].CloneNode(true);
            Escrever(novo, linhas[i]);
            celula.Append(novo);
        }
        foreach (var sobra in paragrafos.Skip(Math.Max(linhas.Count, 1))) sobra.Remove();
    }

    private List<OpenXmlElement> Entre(OpenXmlElement de, OpenXmlElement ate)
    {
        var achou = false;
        var lista = new List<OpenXmlElement>();
        foreach (var filho in _corpo.ChildElements)
        {
            if (filho == de) { achou = true; continue; }
            if (filho == ate) break;
            if (achou) lista.Add(filho);
        }
        return lista;
    }

    private void ApagarEntre(OpenXmlElement de, OpenXmlElement ate)
    {
        foreach (var elemento in Entre(de, ate)) elemento.Remove();
        ate.Remove();
    }
}
