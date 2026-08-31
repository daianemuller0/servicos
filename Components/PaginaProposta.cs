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
        Data.Pricing.Apresentar(R.Documento(), R.Proposta.ModoApresentacao, Data.Pricing.Num(R.Params.TaxaAdmPct));

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
    protected string InjetarEmOutros()
    {
        var alvo = MetaAnterior;
        if (alvo <= 0) return "Informe a diária normal da proposta anterior.";
        var dia = DiariaAtual;
        if (dia <= 0) return "Lance as diárias normais (1º turno) antes de ajustar.";
        if (dia >= alvo)
            return $"A diária atual (R$ {Data.Pricing.Moeda(dia)}) já supera a meta (R$ {Data.Pricing.Moeda(alvo)}) — nada a fazer.";

        var outros = R.ItensDespesa.FirstOrDefault(d => Data.Pricing.EhOutros(d.Despesa));
        if (outros is null)
        {
            outros = new Models.ItemDespesa { Despesa = "OUTROS", Obs = "ADMINISTRATIVA", PorTecnico = false };
            R.ItensDespesa.Add(outros);
        }
        var tec = Math.Max(Data.Pricing.Inteiro(R.Params.QtdTecnicos), 1);
        var qtd = Math.Max(Data.Pricing.Num(outros.Qtd), 1);
        outros.Qtd = qtd.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
        var mult = qtd * (outros.PorTecnico ? tec : 1);

        // A diária cresce quase linearmente com o custo total: aproxima e refina.
        double injetado = 0;
        for (var i = 0; i < 12 && dia < alvo; i++)
        {
            var c = R.Calculo();
            if (c.CustoTotal <= 0 || c.FatorComImpostos <= 0) break;
            var delta = Data.Pricing.ParaCima(Math.Max((alvo - dia) / dia * c.CustoTotal, 1));
            var novoUnit = Data.Pricing.Num(outros.CustoUnitario) + Math.Ceiling(delta / mult * 100) / 100;
            outros.CustoUnitario = novoUnit.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
            injetado += delta;
            dia = DiariaAtual;
        }

        var total = Apresentado().Total;
        return $"OUTROS +R$ {Data.Pricing.Moeda(injetado)} → diária normal R$ {Data.Pricing.Moeda(dia)} (meta R$ {Data.Pricing.Moeda(alvo)}) · total R$ {Data.Pricing.Moeda(total)} ✓";
    }

    /// <summary>Margem que resulta da meta de valor informada (função "chegar no valor").</summary>
    protected double MargemDaMeta =>
        Data.Pricing.MargemParaMeta(R.Calculo(), R.Params, Data.Pricing.Num(R.Params.MetaValor),
            Data.Pricing.Num(R.Proposta.PrazoEntregaDias));

    /// <summary>Aplica a margem calculada pela meta. Retorna a mensagem para o usuário.</summary>
    protected string AplicarMargemDaMeta()
    {
        var meta = Data.Pricing.Num(R.Params.MetaValor);
        if (meta <= 0) return "Informe a meta de valor final.";
        if (R.Calculo().CustoComRisco <= 0) return "Lance algum custo antes de usar a meta.";
        var m = MargemDaMeta;
        R.Params.MargemAlvoPct = (m * 100).ToString("0.####", System.Globalization.CultureInfo.InvariantCulture);
        return $"Margem ajustada para {Data.Pricing.Porcento(m)} — total fecha na meta de R$ {Data.Pricing.Moeda(meta)} ✓";
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
