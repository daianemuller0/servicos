using ClosedXML.Excel;
using HowdenServicos.Poc.Models;

namespace HowdenServicos.Poc.Data;

/// <summary>
/// Devolve a PLANILHA ORIGINAL (subir.xlsm — "Ferramenta para propostas de
/// serviço") preenchida com os dados da proposta feita no sistema, com as
/// macros e fórmulas preservadas, pronta para salvar na pasta de registro.
///
/// O modelo vai embutido no deploy (Recursos/planilha-modelo.xlsm). São
/// preenchidas as células de ENTRADA das guias CUSTO, PRICING e PROPOSTA; as
/// fórmulas da própria planilha recalculam ao abrir no Excel.
/// </summary>
public static class PlanilhaExport
{
    private static string CaminhoModelo =>
        Path.Combine(AppContext.BaseDirectory, "Recursos", "planilha-modelo.xlsm");

    public static byte[] Gerar(Proposta p, List<ItemMO> itensMO, List<ItemDespesa> itensDespesa,
        PricingParams par, Pricing.Documento apresentado)
    {
        using var wb = new XLWorkbook(CaminhoModelo);
        var tec = Math.Max(Pricing.Inteiro(par.QtdTecnicos), 1);

        // ================= CUSTO =================
        var custo = wb.Worksheet("CUSTO");
        custo.Cell("H6").Value = tec;

        // Mão de obra: linhas 8..17, na mesma ordem da Tabela de Custos.
        for (var i = 0; i < Math.Min(itensMO.Count, 10); i++)
        {
            var r = 8 + i;
            var item = itensMO[i];
            custo.Cell(r, 2).Value = item.Servico;
            custo.Cell(r, 3).Value = item.Obs;
            custo.Cell(r, 4).Value = Pricing.Num(item.Horas);
            custo.Cell(r, 5).Value = Pricing.Num(item.CustoHora);     // sobrescreve =E8*1,5 etc. pelo valor real
            custo.Cell(r, 7).Value = Pricing.Num(item.QtdDiaria);
            // H tem fórmula =G*F*$H$6; itens sem "por técnico" no sistema:
            if (!item.PorTecnico) custo.Cell(r, 8).FormulaA1 = $"=G{r}*F{r}";
        }

        // Despesas: linhas 22..31, mesma ordem.
        for (var i = 0; i < Math.Min(itensDespesa.Count, 10); i++)
        {
            var r = 22 + i;
            var item = itensDespesa[i];
            custo.Cell(r, 2).Value = item.Despesa;
            custo.Cell(r, 3).Value = item.Obs;
            custo.Cell(r, 6).Value = Pricing.Num(item.Qtd);           // F27 tinha =$F$24 — vira valor
            custo.Cell(r, 7).Value = Pricing.Num(item.CustoUnitario);
            if (!item.PorTecnico) custo.Cell(r, 8).FormulaA1 = $"=G{r}*F{r}";
        }

        // ================= PRICING =================
        var pricing = wb.Worksheet("PRICING");
        pricing.Cell("E5").Value = Pricing.Num(p.Ano) > 0 ? Pricing.Num(p.Ano) : DateTime.Today.Year;
        pricing.Cell("E7").Value = Pricing.Num(p.Revisao);
        pricing.Cell("E8").Value = p.Bu;
        pricing.Cell("J3").Value = p.Segmento;
        pricing.Cell("J4").Value = p.VendaPara;
        pricing.Cell("J5").Value = p.Destino;
        pricing.Cell("J6").Value = p.Estado;
        pricing.Cell("J8").Value = Pricing.Num(p.PrazoEntregaDias);
        pricing.Cell("P5").Value = string.IsNullOrWhiteSpace(p.Representante) ? "-" : p.Representante;
        pricing.Cell("P6").Value = string.IsNullOrWhiteSpace(p.Representante2) ? "-" : p.Representante2;

        // Margem alvo ("montar preço a partir da margem") e fiança.
        pricing.Cell("P26").Value = Pricing.Pct(par.MargemAlvoPct);
        var temFianca = par.FiancaTipo != "Não" && Pricing.TaxaGarantia(par.FiancaTipo) > 0;
        pricing.Cell("P3").Value = temFianca ? "Sim" : "Não";
        if (temFianca)
        {
            pricing.Cell("R46").Value = par.FiancaTipo;
            pricing.Cell("N49").Value = Pricing.Pct(par.FiancaPctVenda);
        }

        // ================= PROPOSTA =================
        var prop = wb.Worksheet("PROPOSTA");
        prop.Cell("C5").Value = p.Cliente;
        prop.Cell("C6").Value = string.IsNullOrWhiteSpace(p.Cidade) ? p.Estado : $"{p.Cidade} - {p.Estado}";
        prop.Cell("C7").Value = p.ContatoNome;
        prop.Cell("C8").Value = p.ContatoEmail;
        prop.Cell("C9").Value = p.ContatoTelefone;
        prop.Cell("C10").Value = p.Projeto;
        prop.Cell("C11").Value = p.Referencia;
        prop.Cell("C12").Value = Servicos.NumeroCompleto(p);
        prop.Cell("C13").Value = DateTime.TryParse(p.Data, out var data) ? data : DateTime.Today;
        prop.Cell("C15").Value = p.PreparadaPor;
        if (!string.IsNullOrWhiteSpace(p.AssinaNome)) prop.Cell("C16").Value = p.AssinaNome;
        prop.Cell("C17").Value = string.IsNullOrWhiteSpace(p.Representante) ? "-" : p.Representante;
        prop.Cell("L14").Value = p.Moeda;

        // ---- tabela ASSESSORIA (o modelo traz 2 linhas: 25 e 26) ----
        var mo = apresentado.MO.Where(l => l.ValorTotal > 0).ToList();
        var slotsMO = Ajustar(prop, 25, 2, mo.Count);
        for (var i = 0; i < mo.Count; i++)
        {
            var r = 25 + i;
            var l = mo[i];
            prop.Cell(r, 2).Value = "- " + l.Servico;
            prop.Cell(r, 3).Value = l.Obs;
            prop.Cell(r, 4).Value = l.Horas;
            prop.Cell(r, 5).Value = l.ValorHora;
            prop.Cell(r, 6).Value = l.ValorDiaria;
            prop.Cell(r, 7).Value = l.QtdDiaria;
            prop.Cell(r, 8).Value = l.ValorTotal;
        }
        var off1 = slotsMO - 2;
        prop.Cell(27 + off1, 8).Value = apresentado.TotalMO;   // TOTAL ASSESSORIA C/ IMPOSTOS

        // ---- tabela DESPESAS (o modelo traz 1 linha: 31) ----
        var desp = apresentado.Despesas.Where(d => d.ValorTotal > 0)
            .Select(d => (Nome: d.Despesa, d.Obs, Qtd: (double?)d.Qtd, Unit: (double?)d.ValorUnitario, Total: d.ValorTotal))
            .ToList();
        if (apresentado.Deslocamento > 0)
            desp.Add(("DESPESAS DE DESLOCAMENTO", "TÁXI + PASSAGEM AÉREA, TAXA ADM. INCLUSA",
                null, null, apresentado.Deslocamento));

        var inicioDesp = 31 + off1;
        var slotsDesp = Ajustar(prop, inicioDesp, 1, desp.Count);
        for (var i = 0; i < desp.Count; i++)
        {
            var r = inicioDesp + i;
            var d = desp[i];
            prop.Cell(r, 2).Value = "- " + d.Nome;
            prop.Cell(r, 3).Value = d.Obs;
            if (d.Qtd is { } q) prop.Cell(r, 6).Value = q; else prop.Cell(r, 6).Clear(XLClearOptions.Contents);
            if (d.Unit is { } u) prop.Cell(r, 7).Value = u; else prop.Cell(r, 7).Clear(XLClearOptions.Contents);
            prop.Cell(r, 8).Value = d.Total;
        }
        var off2 = slotsDesp - 1;
        prop.Cell(32 + off1 + off2, 8).Value = apresentado.TotalDespesas + apresentado.Deslocamento;

        // ---- totais e resumo: os valores APRESENTADOS, batendo zerado ----
        var (semImp, pisCofins) = Pricing.ResumoDoTotal(apresentado);
        prop.Cell(34 + off1 + off2, 8).Value = apresentado.Total;      // TOTAL C/ IMPOSTOS
        prop.Cell(36 + off1 + off2, 3).Value = semImp;
        prop.Cell(37 + off1 + off2, 3).Value = pisCofins;
        prop.Cell(38 + off1 + off2, 3).Value = apresentado.Total;

        // ---- "INFORMAÇÕES COMPLEMENTARES" vira DIÁRIAS ADICIONAIS (5 linhas) ----
        var adicionais = Pricing.DiariasAdicionais(apresentado);
        if (adicionais.Count > 0)
        {
            var titulo = 40 + off1 + off2;
            prop.Cell(titulo, 2).Value = "DIÁRIAS ADICIONAIS — VALORES PARA DIAS E HORAS ALÉM DO CONTRATADO";
            var inicioAdic = titulo + 2;                               // modelo traz 3 linhas
            Ajustar(prop, inicioAdic, 3, adicionais.Count);
            for (var i = 0; i < adicionais.Count; i++)
            {
                var r = inicioAdic + i;
                var a = adicionais[i];
                prop.Cell(r, 2).Value = "- " + a.Servico;
                prop.Cell(r, 3).Value = a.Obs;
                prop.Cell(r, 4).Value = 1;
                prop.Cell(r, 5).Value = a.Valor;
            }
        }

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms.ToArray();
    }

    /// <summary>
    /// Garante que um bloco de linhas do modelo tenha espaço para a quantidade
    /// necessária: insere linhas (herdando o estilo) ou apaga as sobras.
    /// Devolve o número final de linhas do bloco (mínimo 1).
    /// </summary>
    private static int Ajustar(IXLWorksheet ws, int inicio, int linhasModelo, int necessarias)
    {
        var alvo = Math.Max(necessarias, 1);
        if (alvo > linhasModelo)
            ws.Row(inicio + linhasModelo - 1).InsertRowsBelow(alvo - linhasModelo);
        else if (alvo < linhasModelo)
            ws.Rows(inicio + alvo, inicio + linhasModelo - 1).Delete();
        return alvo;
    }
}
