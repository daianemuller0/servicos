namespace HowdenServicos.Poc.Models;

/// <summary>
/// Proposta de serviço — equivale a uma pasta inteira da planilha antiga
/// (guias CUSTO + PRICING + PROPOSTA). Os campos são texto porque a camada
/// de dados (Parquet) grava tudo como VARCHAR; as conversões ficam no
/// motor de cálculo (<see cref="Data.Pricing"/>).
/// </summary>
public class Proposta
{
    public string Id { get; set; } = "";

    // ---- cadastro do cliente (guia PROPOSTA, coluna C) ----
    public string Cliente { get; set; } = "";
    public string Cidade { get; set; } = "";
    public string ContatoNome { get; set; } = "";
    public string ContatoEmail { get; set; } = "";
    public string ContatoTelefone { get; set; } = "";
    public string Projeto { get; set; } = "";
    public string Referencia { get; set; } = "";

    // ---- cabeçalho / identificação ----
    public string Numero { get; set; } = "";           // ex.: HCHLROU.AFM.000048
    public string Revisao { get; set; } = "0";
    public string Ano { get; set; } = DateTime.Today.Year.ToString();
    public string Data { get; set; } = DateTime.Today.ToString("yyyy-MM-dd");
    public string Bu { get; set; } = "HSA-SP";
    public string Idioma { get; set; } = "Português";
    public string Moeda { get; set; } = "BRL";
    public string ValidadeDias { get; set; } = "30";
    public string PrazoEntregaDias { get; set; } = "30";
    public string PreparadaPor { get; set; } = "";
    public string RevisadaPor { get; set; } = "";
    public string Representante { get; set; } = "-";
    public string Estado { get; set; } = "SP";
    public string Segmento { get; set; } = "Service";
    public string MarketSegment { get; set; } = "Mining";
    public string VendaPara { get; set; } = "Cliente Final";
    public string Destino { get; set; } = "Nacional";

    // ---- contato que assina o documento (cabeçalho impresso) ----
    public string AssinaNome { get; set; } = "";
    public string AssinaCargo { get; set; } = "";
    public string AssinaEmail { get; set; } = "";
    public string AssinaFones { get; set; } = "";

    // ---- conteúdo serializado ----
    public string ItensMoJson { get; set; } = "";
    public string ItensDespesaJson { get; set; } = "";
    public string PricingJson { get; set; } = "";

    // ---- resumo gravado (para as listagens) ----
    public string CustoTotal { get; set; } = "0";
    public string Total { get; set; } = "0";
    public string CriadaEm { get; set; } = "";
    public string Status { get; set; } = "Rascunho";
}

/// <summary>Item de mão de obra (guia CUSTO, linhas 8–17).</summary>
public class ItemMO
{
    public string Servico { get; set; } = "";
    public string Obs { get; set; } = "";
    public string Horas { get; set; } = "8";
    public string CustoHora { get; set; } = "0";
    public string QtdDiaria { get; set; } = "0";
    /// <summary>Quando verdadeiro, o custo é multiplicado pela quantidade de técnicos.</summary>
    public bool PorTecnico { get; set; } = true;
}

/// <summary>Item de despesa (guia CUSTO, linhas 22–31).</summary>
public class ItemDespesa
{
    public string Despesa { get; set; } = "";
    public string Obs { get; set; } = "";
    public string Qtd { get; set; } = "0";
    public string CustoUnitario { get; set; } = "0";
    public bool PorTecnico { get; set; } = true;
}

/// <summary>Parâmetros da guia PRICING (riscos, provisões, margem e impostos).</summary>
public class PricingParams
{
    public string QtdTecnicos { get; set; } = "1";

    // riscos
    public string RiscoPct { get; set; } = "5";          // PRICING D37 (risco adicional)

    // provisões sobre a venda líquida
    public string ComissoesPct { get; set; } = "1.65644"; // PRICING D46
    public string MargemNegociacaoPct { get; set; } = "3";// PRICING D47
    public string PmSachPct { get; set; } = "1";          // PRICING D48
    public string GarantiaProjetoPct { get; set; } = "2"; // PRICING D49
    public string PortalPct { get; set; } = "0";          // PRICING D50 (0,7% quando há portal)
    public string Fianca { get; set; } = "0";             // PRICING E43 (seguro garantia / fiança)

    // margem alvo (Project Margin) — "montar preço a partir da margem"
    public string MargemAlvoPct { get; set; } = "36";

    // impostos (serviço: PIS, COFINS e ISS)
    public string PisPct { get; set; } = "1.65";
    public string CofinsPct { get; set; } = "7.6";
    public string IssPct { get; set; } = "7";             // 2% Itatiba + 5% outro município

    // taxa administrativa das despesas informativas (CUSTO C51/C61)
    public string TaxaAdmPct { get; set; } = "40";
}
