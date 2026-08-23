using ClosedXML.Excel;
using HowdenServicos.Poc.Models;

namespace HowdenServicos.Poc.Data;

/// <summary>
/// Importa a planilha antiga ("Ferramenta para propostas de serviço", .xlsm/.xlsx):
/// lê o cadastro da guia PROPOSTA e os itens de custo da guia CUSTO, no mesmo
/// mapeamento de células que as macros usavam.
///
///   PROPOSTA: C5 cliente · C6 cidade · C7 contato · C8 e-mail · C9 telefone ·
///             C10 projeto · C11 referência · C12 nº da proposta · C15 preparada por ·
///             C16 revisada por · C17 representante
///   CUSTO:    H6 qtd. de técnicos · linhas 8–17 mão de obra (B,C,D,E,G) ·
///             linhas 22–31 despesas (B,C,F,G)
/// </summary>
public static class CustoImport
{
    public sealed record Resultado(Proposta Proposta, List<ItemMO> MO, List<ItemDespesa> Despesas,
        string QtdTecnicos, List<string> Avisos);

    public static Resultado Parse(Stream stream)
    {
        var avisos = new List<string>();
        using var wb = new XLWorkbook(stream);

        var p = new Proposta
        {
            Data = DateTime.Today.ToString("yyyy-MM-dd"),
            Ano = DateTime.Today.Year.ToString(),
        };

        if (Aba(wb, "PROPOSTA") is { } prop)
        {
            p.Cliente = Txt(prop, "C5");
            p.Cidade = Txt(prop, "C6");
            p.ContatoNome = Txt(prop, "C7");
            p.ContatoEmail = Txt(prop, "C8");
            p.ContatoTelefone = Txt(prop, "C9");
            p.Projeto = Txt(prop, "C10");
            p.Referencia = Txt(prop, "C11");
            p.Numero = Txt(prop, "C12");
            p.PreparadaPor = Txt(prop, "C15");
            p.RevisadaPor = Txt(prop, "C16");
            p.Representante = Txt(prop, "C17");
        }
        else
        {
            avisos.Add("A guia PROPOSTA não foi encontrada — o cadastro do cliente ficou em branco.");
        }

        var mo = new List<ItemMO>();
        var desp = new List<ItemDespesa>();
        var tecnicos = "1";

        if (Aba(wb, "CUSTO") is { } custo)
        {
            var t = Val(custo, "H6");
            if (t >= 1) tecnicos = ((int)t).ToString();

            for (var linha = 8; linha <= 17; linha++)
            {
                var servico = Txt(custo, $"B{linha}");
                if (string.IsNullOrWhiteSpace(servico)) continue;
                mo.Add(new ItemMO
                {
                    Servico = servico,
                    Obs = Txt(custo, $"C{linha}"),
                    Horas = N(Val(custo, $"D{linha}")),
                    CustoHora = N(Val(custo, $"E{linha}")),
                    QtdDiaria = N(Val(custo, $"G{linha}")),
                    PorTecnico = linha is 8 or 9 or 10 or 11 or 16,
                });
            }

            for (var linha = 22; linha <= 31; linha++)
            {
                var nome = Txt(custo, $"B{linha}");
                if (string.IsNullOrWhiteSpace(nome)) continue;
                desp.Add(new ItemDespesa
                {
                    Despesa = nome,
                    Obs = Txt(custo, $"C{linha}"),
                    Qtd = N(Val(custo, $"F{linha}")),
                    CustoUnitario = N(Val(custo, $"G{linha}")),
                    PorTecnico = true,
                });
            }
        }
        else
        {
            avisos.Add("A guia CUSTO não foi encontrada — nenhum item foi importado.");
        }

        if (mo.Count == 0 && desp.Count == 0)
            avisos.Add("Nenhuma linha de custo foi lida da planilha.");

        return new Resultado(p, mo, desp, tecnicos, avisos);
    }

    private static IXLWorksheet? Aba(XLWorkbook wb, string nome) =>
        wb.Worksheets.FirstOrDefault(w => string.Equals(w.Name.Trim(), nome, StringComparison.OrdinalIgnoreCase));

    private static string Txt(IXLWorksheet ws, string end)
    {
        try
        {
            var c = ws.Cell(end);
            if (c.IsEmpty()) return "";
            var v = c.HasFormula ? c.CachedValue.ToString() : c.Value.ToString();
            return (v ?? "").Trim();
        }
        catch { return ""; }
    }

    private static double Val(IXLWorksheet ws, string end)
    {
        try
        {
            var c = ws.Cell(end);
            if (c.IsEmpty()) return 0;
            var obj = c.HasFormula ? c.CachedValue : c.Value;
            if (obj.IsNumber) return obj.GetNumber();
            return Pricing.Num(obj.ToString());
        }
        catch { return 0; }
    }

    private static string N(double v) =>
        v == 0 ? "0" : v.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture);
}
