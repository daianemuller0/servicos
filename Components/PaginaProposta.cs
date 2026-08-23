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
