using System.Globalization;
using HowdenServicos.Poc.Models;

namespace HowdenServicos.Poc.Data;

/// <summary>
/// Regras da proposta de serviço: listas dos campos, rótulos traduzidos e a
/// geração do documento (impressão/PDF e Word), no mesmo modelo visual da
/// proposta oficial.
/// </summary>
public static class Servicos
{
    public static readonly string[] Bus = { "HSA-SP", "HSA-ES", "HCHL", "HPU" };
    public static readonly string[] Idiomas = { "Português", "English", "Español" };
    public static readonly string[] Moedas = { "BRL", "USD", "EUR", "CLP", "PEN" };
    public static readonly string[] Segmentos = { "Service", "NB", "AFM", "Intercompany" };
    public static readonly string[] VendasPara = { "Cliente Final", "Industrialização", "Revenda" };
    public static readonly string[] Destinos = { "Nacional", "Exportação" };
    public static readonly string[] MarketSegments =
    {
        "Mining", "Industrial - Cement", "Industrial - Chemicals", "Industrial - Food & Beverages",
        "Industrial - Metals Processing", "Industrial - Petrochemical", "Industrial - Pulp and Paper",
        "Industrial - Other Industries", "Energy & Renewable Power - Refinery",
        "Infrastructure & Mobility - Tunnel", "Infrastructure & Mobility - Water (Municipal)", "Other",
    };
    public static readonly string[] Estados =
    {
        "AC","AL","AM","AP","BA","CE","DF","ES","GO","MA","MG","MS","MT","PA","PB","PE","PI",
        "PR","RJ","RN","RO","RR","RS","SC","SE","SP","TO","Chile","Peru","Exportação",
    };

    /// <summary>Alíquota de ISS conforme a BU emissora (2% Itatiba / 5% Serra).</summary>
    public static string IssPadrao(string bu) => bu == "HSA-ES" ? "5" : "2";

    /// <summary>Símbolo da moeda usado no documento.</summary>
    public static string Simbolo(string moeda) => moeda switch
    {
        "USD" => "US$", "EUR" => "€", "CLP" => "CLP$", "PEN" => "S/", _ => "R$",
    };

    public static string FmtData(string iso) =>
        DateTime.TryParse(iso, out var d) ? d.ToString("dd/MM/yyyy") : (iso ?? "");

    public static string ValidadeData(Proposta p)
    {
        if (!DateTime.TryParse(p.Data, out var d)) return $"{p.ValidadeDias} dias";
        var dias = (int)Pricing.Num(p.ValidadeDias);
        return d.AddDays(dias <= 0 ? 30 : dias).ToString("dd/MM/yyyy");
    }

    public static string NumeroCompleto(Proposta p) =>
        string.IsNullOrWhiteSpace(p.Numero) ? "—" : $"{p.Numero} · Rev. {p.Revisao}";

    public static BillingInfo FaturamentoPadrao(string bu) =>
        Seed.Faturamento().FirstOrDefault(b => b.Id == bu) ?? new BillingInfo { Id = bu };

    // ---- rótulos traduzidos (PT/EN/ES) ----
    public sealed record DocLabels(
        string DadosCliente, string Cliente, string AosCuidados, string Email, string Telefone,
        string Proposta, string Data, string Validade, string Projeto, string Cidade, string Estado,
        string PrazoEntrega, string Dias, string Assessoria, string Servico, string Obs, string Horas,
        string ValorHora, string ValorDiaria, string QtdDiaria, string ValorTotal,
        string Despesas, string Qtd, string ValorUnit, string TotalAssessoria, string TotalDespesas,
        string TotalComImpostos, string SemImpostos, string ComPisCofins, string ComPisCofinsIss,
        string Complementares, string Descricao, string Faturamento, string Banco, string Agencia,
        string Conta, string PreparadaPor, string RevisadaPor);

    public static DocLabels Labels(string idioma) => idioma switch
    {
        "English" => new DocLabels(
            "CUSTOMER DATA", "Customer:", "To the care of:", "E-mail:", "Phone:",
            "QUOTE", "DATE", "EXPIRATION DATE", "PROJECT", "CITY", "STATE / PROVINCE",
            "Delivery time:", "days", "TECHNICAL ASSISTANCE", "SERVICE", "NOTES", "HOURS",
            "HOUR RATE", "DAILY RATE", "DAYS", "TOTAL",
            "EXPENSES", "QTY", "UNIT PRICE", "TOTAL TECHNICAL ASSISTANCE (WITH TAXES)",
            "TOTAL EXPENSES (WITH TAXES)", "TOTAL WITH TAXES", "AMOUNT WITHOUT TAXES",
            "AMOUNT WITH PIS AND COFINS", "AMOUNT WITH PIS, COFINS AND ISS",
            "ADDITIONAL INFORMATION — NOT INCLUDED", "DESCRIPTION", "INVOICING INFORMATION",
            "Bank", "Branch", "Account", "Prepared by:", "Reviewed by:"),
        "Español" => new DocLabels(
            "DATOS DEL CLIENTE", "Cliente:", "Al cuidado de:", "E-mail:", "Fono:",
            "OFERTA COMERCIAL", "FECHA", "VALIDEZ", "PROYECTO", "CIUDAD", "ESTADO / PROVINCIA",
            "Plazo de entrega:", "días", "ASESORÍA TÉCNICA", "SERVICIO", "OBS", "HORAS",
            "VALOR HORA", "VALOR DÍA", "CTD. DÍAS", "VALOR TOTAL",
            "GASTOS", "CTD", "VALOR UNITARIO", "TOTAL ASESORÍA (CON IMPUESTOS)",
            "TOTAL GASTOS (CON IMPUESTOS)", "TOTAL CON IMPUESTOS", "VALOR SIN IMPUESTOS",
            "VALOR CON PIS Y COFINS", "VALOR CON PIS, COFINS E ISS",
            "INFORMACIONES COMPLEMENTARIAS — NO INCLUIDO", "DESCRIPCIÓN", "DATOS PARA FACTURACIÓN",
            "Banco", "Sucursal", "Cuenta", "Preparado por:", "Revisado por:"),
        _ => new DocLabels(
            "DADOS DO CLIENTE", "Cliente:", "Aos cuidados de:", "E-mail:", "Telefone:",
            "Proposta", "DATA", "VALIDADE", "PROJETO", "CIDADE", "ESTADO",
            "Prazo de entrega:", "dias", "ASSESSORIA TÉCNICA", "SERVIÇOS", "OBS", "HORAS",
            "VALOR HORA", "VALOR 1 DIARIA", "QTD. DIARIA", "VALOR TOTAL",
            "DESPESAS", "QTD", "VALOR UNITARIO", "TOTAL ASSESSORIA C/ IMPOSTOS",
            "TOTAL DESPESAS C/ IMPOSTOS", "TOTAL C/ IMPOSTOS", "VALOR SEM IMPOSTOS",
            "VALOR C/ PIS E COFINS", "VALOR C/ PIS, COFINS E ISS",
            "INFORMAÇÕES COMPLEMENTARES — NÃO INCLUSO", "DESCRIÇÃO", "DADOS PARA FATURAMENTO",
            "Banco", "Agência", "Conta", "Preparada por:", "Revisada por:"),
    };

    /// <summary>Logo do cabeçalho: imagem enviada em Identidade Visual ou o wordmark padrão.</summary>
    public static string LogoHtml(string? logoDataUri = null)
    {
        if (!string.IsNullOrEmpty(logoDataUri))
            return $"<img src='{logoDataUri}' alt='Howden' style='height:72px' />";

        return
            "<div style='font-family:Arial,sans-serif;line-height:1'>" +
            "<span style='font-weight:bold;font-size:30pt;color:#00539B;letter-spacing:.5px'>Howden</span>" +
            "<div style='font-style:italic;font-size:9pt;color:#00539B;margin-top:2px'>A Chart Industries Company</div>" +
            "</div>";
    }

    /// <summary>Documento completo em HTML (usado no download em Word).</summary>
    public static string DocHtml(Proposta p, Pricing.Documento doc, string? logo, BillingInfo? fat) =>
        $@"<html><head><meta charset='utf-8'><title>Proposta {System.Net.WebUtility.HtmlEncode(p.Numero)}</title>
<style>body {{ font-family: Arial, sans-serif; color: #3C465A; font-size: 9pt; margin: 32px; }}</style>
</head><body>{DocBody(p, doc, logo, fat)}</body></html>";

    /// <summary>
    /// Miolo do documento — compartilhado entre a impressão/PDF e o Word.
    /// Mesmas cores do modelo oficial: Arial, títulos navy #141E32, texto
    /// #3C465A e caixas/tabelas em azul #004785. Estilos inline por causa do Word.
    /// </summary>
    public static string DocBody(Proposta p, Pricing.Documento doc, string? logo, BillingInfo? billing)
    {
        const string azul = "#004785";
        const string navy = "#141E32";
        const string corpo = "#3C465A";
        const string sec = "#46506E";
        const string rod = "#6E7890";
        const string ft = "font-family:Arial,sans-serif";
        const string bd = "border:1px solid #d5dae4;padding:6px 8px";
        static string E(string? s) => System.Net.WebUtility.HtmlEncode(s ?? "");
        static string D(string? s) => string.IsNullOrWhiteSpace(s) ? "—" : System.Net.WebUtility.HtmlEncode(s);

        var L = Labels(p.Idioma);
        var cif = Simbolo(p.Moeda);
        string M(double v) => $"{cif} {Pricing.Moeda(v)}";

        string Th(string t, string align = "left") =>
            $"<th style='background:{azul};color:#fff;padding:6px 8px;text-align:{align};font-size:8pt'>{t}</th>";

        var linhasMO = string.Join("", doc.MO.Select(i =>
            "<tr>" +
            $"<td style='{bd};color:{corpo}'><b style='color:{navy}'>{E(i.Servico)}</b></td>" +
            $"<td style='{bd};color:{corpo}'>{E(i.Obs)}</td>" +
            $"<td style='{bd};text-align:center;color:{corpo}'>{Pricing.Moeda0(i.Horas)}</td>" +
            $"<td style='{bd};text-align:right;color:{corpo}'>{Pricing.Moeda(i.ValorHora)}</td>" +
            $"<td style='{bd};text-align:right;color:{corpo}'>{Pricing.Moeda(i.ValorDiaria)}</td>" +
            $"<td style='{bd};text-align:center;color:{corpo}'>{Pricing.Moeda0(i.QtdDiaria)}</td>" +
            $"<td style='{bd};text-align:right;color:{corpo}'>{M(i.ValorTotal)}</td></tr>"));

        var linhasDesp = string.Join("", doc.Despesas.Select(i =>
            "<tr>" +
            $"<td style='{bd};color:{corpo}'><b style='color:{navy}'>{E(i.Despesa)}</b></td>" +
            $"<td style='{bd};color:{corpo}'>{E(i.Obs)}</td>" +
            $"<td style='{bd};text-align:center;color:{corpo}'>{Pricing.Moeda0(i.Qtd)}</td>" +
            $"<td style='{bd};text-align:right;color:{corpo}'>{Pricing.Moeda(i.ValorUnitario)}</td>" +
            $"<td style='{bd};text-align:right;color:{corpo}'>{M(i.ValorTotal)}</td></tr>"));

        var linhasComp = string.Join("", doc.Complementares.Select(c =>
            "<tr>" +
            $"<td style='{bd};color:{corpo}'><b style='color:{navy}'>{E(c.Descricao)}</b></td>" +
            $"<td style='{bd};color:{corpo}'>{E(c.Obs)}</td>" +
            $"<td style='{bd};text-align:center;color:{corpo}'>{Pricing.Moeda0(c.Qtd)}</td>" +
            $"<td style='{bd};text-align:right;color:{corpo}'>{M(c.Valor)}</td></tr>"));

        var fat = billing ?? FaturamentoPadrao(p.Bu);
        var bancoLinha = string.IsNullOrWhiteSpace(fat.BancoNome) ? "" :
            $"{L.Banco}: {E(fat.BancoNome)} – {L.Agencia}: {E(fat.Agencia)} {L.Conta}: {E(fat.Conta)}";

        var totalTabela = $@"
<table style='width:100%;border-collapse:collapse;margin-top:10px'><tr>
<td style='padding:0'></td>
<td style='width:210px;background:{azul};color:#fff;padding:9px 14px;font-weight:bold'>{L.TotalComImpostos}</td>
<td style='width:150px;background:{azul};color:#fff;padding:9px 14px;font-weight:bold;text-align:right'>{M(doc.Total)}</td>
</tr></table>";

        return $@"
<div style='{ft};color:{corpo};font-size:9pt'>
<table style='width:100%;border-collapse:collapse'><tr>
<td style='vertical-align:top;padding:0'>
  <p style='margin:0;{ft};font-weight:bold;font-size:12pt;color:{navy}'>Howden South America</p>
  <p style='margin:10px 0 0;{ft};font-weight:bold;font-size:8.5pt;color:{sec}'>{D(p.AssinaNome)}</p>
  <p style='margin:0;{ft};font-size:8.5pt;color:{sec}'>{E(p.AssinaCargo)}<br/>
  <a href='mailto:{E(p.AssinaEmail)}' style='color:{sec}'>{E(p.AssinaEmail)}</a><br/>{E(p.AssinaFones)}</p>
</td>
<td style='vertical-align:top;text-align:right;padding:0'>{LogoHtml(logo)}</td></tr></table>

<table style='width:100%;border-collapse:collapse;margin-top:22px'><tr>
<td style='vertical-align:top;padding:0'>
  <p style='margin:0;font-weight:bold;font-size:10pt;color:{navy}'>{L.DadosCliente}</p>
  <p style='margin:8px 0 0;color:{corpo}'>{L.Cliente} {D(p.Cliente)}</p>
  {(string.IsNullOrWhiteSpace(p.ContatoNome) ? "" : $"<p style='margin:2px 0 0;color:{corpo}'>{L.AosCuidados} {E(p.ContatoNome)}</p>")}
  {(string.IsNullOrWhiteSpace(p.ContatoEmail) ? "" : $"<p style='margin:2px 0 0;color:{corpo}'>{L.Email} {E(p.ContatoEmail)}</p>")}
  {(string.IsNullOrWhiteSpace(p.ContatoTelefone) ? "" : $"<p style='margin:2px 0 0;color:{corpo}'>{L.Telefone} {E(p.ContatoTelefone)}</p>")}
</td>
<td style='vertical-align:top;width:310px;padding:0'>
  <table style='width:100%;border-collapse:collapse'>
    <tr><td colspan='2' style='background:{azul};color:#fff;padding:8px 14px;font-weight:bold;font-size:8.5pt;border-bottom:2px solid #fff'>{L.Proposta} {E(p.Numero)} · Rev. {E(p.Revisao)}</td></tr>
    <tr><td style='background:{azul};color:#fff;padding:8px 14px;font-size:8.5pt;font-weight:bold;border-bottom:2px solid #fff'>{L.Data}</td>
        <td style='background:{azul};color:#fff;padding:8px 14px;font-size:8.5pt;text-align:right;border-bottom:2px solid #fff'>{FmtData(p.Data)}</td></tr>
    <tr><td style='background:{azul};color:#fff;padding:8px 14px;font-size:8.5pt;font-weight:bold'>{L.Validade}</td>
        <td style='background:{azul};color:#fff;padding:8px 14px;font-size:8.5pt;text-align:right'>{ValidadeData(p)}</td></tr>
  </table>
</td></tr></table>

<hr style='border:none;border-top:3px solid {azul};margin:16px 0 12px'/>

<table style='width:100%;border-collapse:collapse'><tr>
<td style='padding:0;width:40%'><b style='color:{navy}'>{L.Projeto}</b><br/><span style='color:{corpo}'>{D(p.Projeto)}</span></td>
<td style='padding:0;width:30%'><b style='color:{navy}'>{L.Cidade}</b><br/><span style='color:{corpo}'>{D(p.Cidade)}</span></td>
<td style='padding:0;width:30%'><b style='color:{navy}'>{L.Estado}</b><br/><span style='color:{corpo}'>{D(p.Estado)}</span></td>
</tr></table>
{(string.IsNullOrWhiteSpace(p.Referencia) ? "" : $"<p style='color:{corpo};margin:10px 0 0'><b style='color:{navy}'>Ref.:</b> {E(p.Referencia)}</p>")}
<p style='color:{corpo};margin:10px 0 16px'>{L.PrazoEntrega} {D(p.PrazoEntregaDias)} {L.Dias}</p>

<p style='margin:0 0 4px;font-weight:bold;font-size:10pt;color:{navy}'>{L.Assessoria}</p>
<table style='width:100%;border-collapse:collapse;font-size:8.5pt'>
<tr>{Th(L.Servico)}{Th(L.Obs)}{Th(L.Horas, "center")}{Th(L.ValorHora, "right")}{Th(L.ValorDiaria, "right")}{Th(L.QtdDiaria, "center")}{Th(L.ValorTotal, "right")}</tr>
{linhasMO}
<tr><td colspan='6' style='{bd};font-weight:bold;color:{navy}'>{L.TotalAssessoria}</td>
    <td style='{bd};text-align:right;font-weight:bold;color:{navy}'>{M(doc.TotalMO)}</td></tr>
</table>

{(doc.Despesas.Count == 0 ? "" : $@"
<p style='margin:18px 0 4px;font-weight:bold;font-size:10pt;color:{navy}'>{L.Despesas}</p>
<table style='width:100%;border-collapse:collapse;font-size:8.5pt'>
<tr>{Th(L.Despesas)}{Th(L.Obs)}{Th(L.Qtd, "center")}{Th(L.ValorUnit, "right")}{Th(L.ValorTotal, "right")}</tr>
{linhasDesp}
<tr><td colspan='4' style='{bd};font-weight:bold;color:{navy}'>{L.TotalDespesas}</td>
    <td style='{bd};text-align:right;font-weight:bold;color:{navy}'>{M(doc.TotalDespesas)}</td></tr>
</table>")}

{totalTabela}

<table style='width:100%;border-collapse:collapse;margin-top:16px;font-size:8.5pt'>
<tr><td style='{bd};color:{corpo}'>{L.SemImpostos}</td><td style='{bd};text-align:right;color:{corpo}'>{M(doc.Calculo.VendaLiquida)}</td></tr>
<tr><td style='{bd};color:{corpo}'>{L.ComPisCofins}</td><td style='{bd};text-align:right;color:{corpo}'>{M(doc.Calculo.ComPisCofins)}</td></tr>
<tr><td style='{bd};color:{navy};font-weight:bold'>{L.ComPisCofinsIss}</td><td style='{bd};text-align:right;color:{navy};font-weight:bold'>{M(doc.Calculo.ComImpostos)}</td></tr>
</table>

<p style='margin:24px 0 4px;font-weight:bold;font-size:10pt;color:{navy}'>{L.Complementares}</p>
<table style='width:100%;border-collapse:collapse;font-size:8.5pt'>
<tr>{Th(L.Descricao)}{Th(L.Obs)}{Th(L.Qtd, "center")}{Th(L.ComPisCofinsIss, "right")}</tr>
{linhasComp}
</table>

<p style='margin:26px 0 4px;font-weight:bold;font-size:10pt;color:{navy}'>{L.Faturamento}</p>
<p style='margin:0;color:{corpo}'><b style='color:{navy}'>{E(fat.Razao)}</b><br/>{E(fat.Endereco)}{(string.IsNullOrWhiteSpace(fat.Registro) ? "" : $"<br/><span style='color:{sec}'>{E(fat.Registro)}</span>")}</p>
{(bancoLinha == "" ? "" : $"<p style='margin:12px 0 0'><span style='border:1px solid #d5dae4;border-radius:3px;padding:5px 12px;display:inline-block;color:#23253F'><b>{bancoLinha}</b></span></p>")}

<p style='color:{rod};font-style:italic;margin-top:26px'>{L.PreparadaPor} {D(p.PreparadaPor)}{(string.IsNullOrWhiteSpace(p.RevisadaPor) ? "" : $" · {L.RevisadaPor} {E(p.RevisadaPor)}")} · {E(p.Ano)} · BU {E(p.Bu)}</p>
</div>";
    }
}
