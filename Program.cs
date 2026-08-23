using System.Security.Claims;
using HowdenServicos.Poc.Components;
using HowdenServicos.Poc.Data;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;

var builder = WebApplication.CreateBuilder(args);

// Porta padrão (modo por-usuário). Para "servidor central", rode com:
//   HowdenServicos.Poc.exe --urls http://0.0.0.0:5081
if (string.IsNullOrEmpty(builder.Configuration["urls"]) &&
    string.IsNullOrEmpty(Environment.GetEnvironmentVariable("ASPNETCORE_URLS")))
{
    builder.WebHost.UseUrls("http://localhost:5081");
}

// --- Blazor Server (componentes interativos no servidor) ---
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

// --- Autenticação por cookie (sessão fica no navegador do usuário) ---
builder.Services.AddCascadingAuthenticationState();
builder.Services
    .AddAuthentication(CookieAuthenticationDefaults.AuthenticationScheme)
    .AddCookie(options =>
    {
        options.LoginPath = "/login";
        options.ExpireTimeSpan = TimeSpan.FromDays(7);
        options.SlidingExpiration = true;
    });
builder.Services.AddAuthorization();

// --- Dados: DuckDB (motor) sobre Parquet numa pasta de rede ---
var dataFolder = builder.Configuration["Data:Folder"] ?? "data";
builder.Services.AddSingleton(new ParquetStore(dataFolder));
builder.Services.AddScoped<PropostaRepository>();
builder.Services.AddScoped<ParametroRepository>();
builder.Services.AddScoped<FaturamentoRepository>();
builder.Services.AddScoped<BrandingRepository>();
builder.Services.AddScoped<RepresentanteRepository>();

// Rascunho da proposta: vive no circuito do usuário (Custo → Pricing → Proposta).
builder.Services.AddScoped<Rascunho>();

var app = builder.Build();

// Semeia a tabela de custos padrão na primeira execução.
using (var scope = app.Services.CreateScope())
{
    DbInitializer.Initialize(scope.ServiceProvider.GetRequiredService<ParquetStore>());
}

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
}

app.UseStaticFiles();
app.UseAntiforgery();
app.UseAuthentication();
app.UseAuthorization();

// --- Login/logout (precisam do HttpContext para gravar o cookie) ---
app.MapPost("/auth/login", async (HttpContext http, IConfiguration cfg) =>
{
    var form = await http.Request.ReadFormAsync();
    var usuario = form["usuario"].ToString().Trim();
    var senha = form["senha"].ToString();

    var cfgUsuario = cfg["Auth:Usuario"] ?? "howden";
    var cfgSenha = cfg["Auth:Senha"] ?? "howden2026";

    if (!usuario.Equals(cfgUsuario, StringComparison.OrdinalIgnoreCase) || senha != cfgSenha)
        return Results.Redirect("/login?error=1");

    var claims = new List<Claim>
    {
        new(ClaimTypes.NameIdentifier, "equipe"),
        new(ClaimTypes.Name, "Equipe Howden"),
        new(ClaimTypes.Role, "admin"),
    };
    var identity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);
    await http.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, new ClaimsPrincipal(identity));
    return Results.Redirect("/");
}).DisableAntiforgery();

app.MapPost("/auth/logout", async (HttpContext http) =>
{
    await http.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
    return Results.Redirect("/login");
}).DisableAntiforgery();

// Exporta as propostas gravadas em CSV (BOM UTF-8, separador ';' p/ Excel pt-BR).
app.MapGet("/servicos/propostas/export", (PropostaRepository repo) =>
{
    static string C(string s) => s.Contains(';') || s.Contains('"') || s.Contains('\n')
        ? $"\"{s.Replace("\"", "\"\"")}\"" : s;

    var sb = new System.Text.StringBuilder();
    sb.AppendLine("Número;Rev.;Data;Cliente;Cidade;Projeto;BU;Moeda;Custo total;Total c/ impostos;Status;Preparada por");
    foreach (var p in repo.All())
    {
        sb.AppendLine(string.Join(';', new[]
        {
            C(p.Numero), C(p.Revisao), Servicos.FmtData(p.Data), C(p.Cliente), C(p.Cidade),
            C(p.Projeto), C(p.Bu), C(p.Moeda),
            Pricing.Num(p.CustoTotal).ToString("0.00", System.Globalization.CultureInfo.InvariantCulture),
            Pricing.Num(p.Total).ToString("0.00", System.Globalization.CultureInfo.InvariantCulture),
            C(p.Status), C(p.PreparadaPor),
        }));
    }
    var bytes = System.Text.Encoding.UTF8.GetPreamble()
        .Concat(System.Text.Encoding.UTF8.GetBytes(sb.ToString())).ToArray();
    return Results.File(bytes, "text/csv; charset=utf-8", "propostas-servicos.csv");
}).RequireAuthorization();

app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

// Modo por-usuário: abre o navegador sozinho ao iniciar.
if (builder.Configuration.GetValue("OpenBrowser", true))
{
    app.Lifetime.ApplicationStarted.Register(() =>
    {
        try
        {
            var url = (app.Urls.FirstOrDefault() ?? "http://localhost:5081")
                .Replace("0.0.0.0", "localhost").Replace("[::]", "localhost");
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(url) { UseShellExecute = true });
        }
        catch { /* sem navegador disponível: apenas ignora */ }
    });
}

app.Run();
