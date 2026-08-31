using HowdenServicos.Poc.Data;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.JSInterop;

namespace HowdenServicos.Poc.Components;

/// <summary>
/// Base das três telas que trabalham na mesma proposta (Custo → Pricing → Proposta).
///
/// O rascunho vive no <see cref="Rascunho"/>, que é um serviço "scoped" — ou seja,
/// dura enquanto o circuito do Blazor estiver de pé. Como um F5 (ou abrir a URL
/// direto) começa um circuito novo, cada render também grava o rascunho no
/// localStorage do navegador e o restaura na primeira renderização.
/// </summary>
public abstract class PaginaProposta : ComponentBase, IDisposable
{
    private DotNetObjectReference<PaginaProposta>? _jsRef;

    [Inject] protected Rascunho R { get; set; } = default!;
    [Inject] protected IJSRuntime JS { get; set; } = default!;
    [Inject] protected ParametroRepository Parametros { get; set; } = default!;
    [CascadingParameter] protected Task<AuthenticationState>? AuthState { get; set; }

    protected string Usuario { get; private set; } = "";
    private string _ultimoSalvo = "";

    protected override async Task OnInitializedAsync()
    {
        if (AuthState is not null)
        {
            var auth = await AuthState;
            Usuario = auth.User.Identity?.Name ?? "";
        }
    }

    protected override async Task OnAfterRenderAsync(bool primeiraVez)
    {
        if (primeiraVez)
        {
            var json = await JS.InvokeAsync<string>("appLoadDraft");
            if (R.Vazio) R.FromJson(json);
            if (R.Vazio) R.Novo(Parametros.All(), Usuario);
            _ultimoSalvo = R.ToJson();

            // Escuta alterações feitas em OUTRAS guias do navegador.
            _jsRef = DotNetObjectReference.Create(this);
            await JS.InvokeVoidAsync("appWatchDraft", _jsRef);

            StateHasChanged();
            return;
        }

        var atual = R.ToJson();
        if (atual == _ultimoSalvo) return;
        _ultimoSalvo = atual;
        await JS.InvokeVoidAsync("appSaveDraft", atual);
    }

    /// <summary>Começa uma proposta do zero, limpando também o rascunho do navegador.</summary>
    protected async Task NovaProposta()
    {
        R.Novo(Parametros.All(), Usuario);
        _ultimoSalvo = R.ToJson();
        await JS.InvokeVoidAsync("appSaveDraft", _ultimoSalvo);
    }

    /// <summary>Grava o rascunho agora.</summary>
    protected async Task SalvarRascunho()
    {
        _ultimoSalvo = R.ToJson();
        await JS.InvokeVoidAsync("appSaveDraft", _ultimoSalvo);
    }

    /// <summary>
    /// Soma das diárias lançadas (DIARIAS NORMAIS 1º/2º turno + DIARIAS EXTRAS)
    /// preenche sozinha a quantidade das despesas por dia: hospedagem, locação
    /// de carro, combustível e refeições. Continua editável depois.
    /// </summary>
    protected void SincronizarDiarias()
    {
        var dias = R.ItensMO
            .Where(i => (i.Servico ?? "").TrimStart().ToUpperInvariant().StartsWith("DIARIAS"))
            .Sum(i => Data.Pricing.Num(i.QtdDiaria));
        var texto = dias.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
        foreach (var d in R.ItensDespesa)
            if (Data.Pricing.EhDespesaDiaria(d.Despesa)) d.Qtd = texto;
    }

    // ================= ferramentas de preço =================

    /// <summary>O documento como apresentado ao cliente (para ler a diária normal).</summary>
    private Data.Pricing.Documento Apresentado() =>
        Data.Pricing.Apresentar(R.Documento(), R.Proposta.ModoApresentacao,
            Data.Pricing.Num(R.Params.TaxaAdmPct), Data.Pricing.Num(R.Params.DiariaTravada),
            Data.Pricing.Num(R.Params.TotalTravado));

    /// <summary>Diária normal (com impostos) da proposta atual, como sai para o cliente.</summary>
    protected double DiariaAtual => Data.Pricing.DiariaNormalApresentada(Apresentado());

    /// <summary>Meta de DIÁRIA: diária normal da proposta anterior × (1 + % a mais).</summary>
    protected double MetaAnterior =>
        Data.Pricing.Num(R.Params.PropAnteriorValor) * (1 + Data.Pricing.Pct(R.Params.PropAnteriorPct));

    /// <summary>Quanto falta na diária normal para superar a da proposta anterior.</summary>
    protected double FaltaParaAnterior
    {
        get
        {
            var alvo = MetaAnterior;
            if (alvo <= 0) return 0;
            var falta = alvo - DiariaAtual;
            return falta > 0 ? falta : 0;
        }
    }

    /// <summary>
    /// Injeta em OUTROS o custo necessário para a DIÁRIA NORMAL (com impostos)
    /// superar a diária da proposta anterior + % a mais. Como as demais linhas
    /// são múltiplos fixos da diária normal (sáb/dom = 2×; HE = 1,5×/2×), todas
    /// acompanham. Retorna a mensagem para exibir ao usuário.
    /// </summary>
    private Models.ItemDespesa GarantirOutros()
    {
        var outros = R.ItensDespesa.FirstOrDefault(d => Data.Pricing.EhOutros(d.Despesa));
        if (outros is null)
        {
            outros = new Models.ItemDespesa { Despesa = "OUTROS", Obs = "ADMINISTRATIVA", PorTecnico = false };
            R.ItensDespesa.Add(outros);
        }
        if (Data.Pricing.Num(outros.Qtd) < 1) outros.Qtd = "1";
        return outros;
    }

    /// <summary>
    /// Ajusta o custo de OUTROS (centavos, para cima ou para baixo) até a diária
    /// normal apresentada cravar no alvo; aplica o pino cosmético. Retorna a
    /// diária final.
    /// </summary>
    private double BuscarDiariaExata(double alvo, Models.ItemDespesa outros)
    {
        var inv = System.Globalization.CultureInfo.InvariantCulture;
        double Unit() => Data.Pricing.Num(outros.CustoUnitario);
        void SetUnit(double v) => outros.CustoUnitario = Math.Max(v, 0).ToString("0.##", inv);

        R.Params.DiariaTravada = "";              // solta o pino durante a busca
        for (var i = 0; i < 40; i++)
        {
            var dia = DiariaAtual;
            var erro = alvo - dia;
            if (Math.Abs(erro) <= 0.05) break;

            var u = Unit();
            if (u <= 0 && erro < 0) break;        // não dá para reduzir além de zero

            SetUnit(u + 10);
            var sensibilidade = (DiariaAtual - dia) / 10.0;
            SetUnit(u);
            if (sensibilidade <= 1e-9) break;

            SetUnit(u + erro / sensibilidade);
        }

        if (Math.Abs(DiariaAtual - alvo) <= 1.0)
            R.Params.DiariaTravada = alvo.ToString("0.00", inv);
        return DiariaAtual;
    }

    /// <summary>Grava a margem que fecha o total na meta e a trava do total.</summary>
    private void AplicarMargemParaTotal(double meta)
    {
        var inv = System.Globalization.CultureInfo.InvariantCulture;
        var m = Data.Pricing.MargemParaMeta(R.Calculo(), R.Params, meta,
            Data.Pricing.Num(R.Proposta.PrazoEntregaDias));
        R.Params.MargemAlvoPct = (m * 100).ToString("0.######", inv);
        R.Params.TotalTravado = meta.ToString("0.00", inv);
    }

    /// <summary>
    /// Quando as duas metas não fecham juntas (ex.: só diárias normais
    /// lançadas), explica a relação e sugere os números possíveis.
    /// </summary>
    private string AvisoConflito(double diaAlvo, double totalAlvo)
    {
        var doc = Apresentado();
        var normal = doc.MO.FirstOrDefault(l => l.Mult is > 0.99 and < 1.01 && l.QtdDiaria > 0);
        if (normal is null) return "";
        var outras = doc.MO.Where(l => l != normal).Sum(l => l.ValorTotal) + doc.TotalDespesas + doc.Deslocamento;
        var totalNecessario = diaAlvo * normal.QtdDiaria + outras;
        var diariaPossivel = normal.QtdDiaria > 0 ? (totalAlvo - outras) / normal.QtdDiaria : 0;
        return $" ⚠ As duas metas não fecham juntas com as linhas lançadas: para diária R$ {Data.Pricing.Moeda(diaAlvo)} o total seria R$ {Data.Pricing.Moeda(totalNecessario)}; para total R$ {Data.Pricing.Moeda(totalAlvo)} a diária seria R$ {Data.Pricing.Moeda(diariaPossivel)}. Ajuste uma das metas ou lance 2º turno / sáb-dom / horas extras.";
    }

    /// <summary>
    /// Fecha as DUAS metas ao mesmo tempo: a margem (GM) é resolvida para o
    /// total e OUTROS para a diária — alternando até os dois cravarem.
    /// </summary>
    private (double Dia, double Total) ReconciliarTravas(double diaAlvo, double totalAlvo, Models.ItemDespesa outros)
    {
        double dia = 0, total = 0;
        for (var i = 0; i < 15; i++)
        {
            AplicarMargemParaTotal(totalAlvo);            // GM fecha o total (custo atual)
            dia = BuscarDiariaExata(diaAlvo, outros);     // OUTROS fecha a diária (muda o custo)
            AplicarMargemParaTotal(totalAlvo);            // GM refeita para o custo novo
            total = Apresentado().Total;
            if (Math.Abs(dia - diaAlvo) <= 0.05 && Math.Abs(total - totalAlvo) <= 0.05) break;
        }
        return (DiariaAtual, Apresentado().Total);
    }

    protected string InjetarEmOutros()
    {
        var alvo = Math.Round(MetaAnterior, 2);
        if (alvo <= 0) return "Informe a diária normal da proposta anterior.";

        R.Params.DiariaTravada = "";
        if (DiariaAtual <= 0) return "Lance as diárias normais (1º turno) antes de ajustar.";

        var outros = GarantirOutros();
        var unitInicial = Data.Pricing.Num(outros.CustoUnitario);

        double diaFinal, totalFinal;
        var totalAlvo = Data.Pricing.Num(R.Params.TotalTravado);
        if (totalAlvo > 0)
        {
            (diaFinal, totalFinal) = ReconciliarTravas(alvo, totalAlvo, outros);
        }
        else
        {
            diaFinal = BuscarDiariaExata(alvo, outros);
            totalFinal = Apresentado().Total;
        }

        var variacao = (Data.Pricing.Num(outros.CustoUnitario) - unitInicial)
                     * Math.Max(Data.Pricing.Num(outros.Qtd), 1)
                     * (outros.PorTecnico ? Math.Max(Data.Pricing.Inteiro(R.Params.QtdTecnicos), 1) : 1);

        if (Math.Abs(diaFinal - alvo) > 0.05)
        {
            var explicacao = totalAlvo > 0
                ? AvisoConflito(alvo, totalAlvo)
                : " Os custos reais já superam a meta (OUTROS chegou ao mínimo).";
            return $"Diária ficou em R$ {Data.Pricing.Moeda(diaFinal)} (meta R$ {Data.Pricing.Moeda(alvo)}).{explicacao}";
        }

        var verbo = variacao >= 0 ? $"+R$ {Data.Pricing.Moeda(variacao)}" : $"−R$ {Data.Pricing.Moeda(-variacao)}";
        var aviso = totalAlvo > 0 && Math.Abs(totalFinal - totalAlvo) > 0.05
            ? AvisoConflito(alvo, totalAlvo)
            : "";
        return $"OUTROS {verbo} → diária normal CRAVADA em R$ {Data.Pricing.Moeda(diaFinal)} · total R$ {Data.Pricing.Moeda(totalFinal)} ✓{aviso}";
    }

    /// <summary>Margem que resulta da meta de valor informada (função "chegar no valor").</summary>
    protected double MargemDaMeta =>
        Data.Pricing.MargemParaMeta(R.Calculo(), R.Params, Data.Pricing.Num(R.Params.MetaValor),
            Data.Pricing.Num(R.Proposta.PrazoEntregaDias));

    /// <summary>Aplica a margem calculada pela meta. Retorna a mensagem para o usuário.</summary>
    protected string AplicarMargemDaMeta()
    {
        var meta = Math.Round(Data.Pricing.Num(R.Params.MetaValor), 2);
        if (meta <= 0) return "Informe a meta de valor final.";
        if (R.Calculo().CustoComRisco <= 0) return "Lance algum custo antes de usar a meta.";

        var diaAlvo = Math.Round(MetaAnterior, 2);

        double totalFinal;
        string aviso = "";
        if (diaAlvo > 0)
        {
            // As duas metas ativas: GM fecha o total e OUTROS refecha a diária.
            var (dia, tot) = ReconciliarTravas(diaAlvo, meta, GarantirOutros());
            totalFinal = tot;
            if (Math.Abs(dia - diaAlvo) > 0.05 || Math.Abs(tot - meta) > 0.05)
                aviso = AvisoConflito(diaAlvo, meta);
        }
        else
        {
            AplicarMargemParaTotal(meta);
            totalFinal = Apresentado().Total;
        }

        var margem = Data.Pricing.Porcento(R.Calculo().ProjectMargin);
        return $"Margem ajustada para {margem} — total CRAVADO em R$ {Data.Pricing.Moeda(totalFinal)} ✓{aviso}";
    }

    /// <summary>
    /// Chamado pelo JavaScript quando OUTRA guia do navegador altera o rascunho:
    /// aplica o novo estado e re-renderiza — as duas guias ficam 100% espelhadas.
    /// </summary>
    [JSInvokable]
    public Task DraftAtualizado(string json)
    {
        if (json == _ultimoSalvo) return Task.CompletedTask;
        if (R.FromJson(json))
        {
            _ultimoSalvo = json;
            return InvokeAsync(StateHasChanged);
        }
        return Task.CompletedTask;
    }

    public void Dispose() => _jsRef?.Dispose();
}
