using System.Diagnostics;

namespace HowdenServicos.Launcher;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
        Application.Run(new FormAtualizacao());
    }
}

/// <summary>
/// Fluxo de publicação dinâmica:
///  - na REDE (pasta deste .exe): app\versao.txt aponta a versão atual, e
///    app\&lt;versão&gt;\ tem os arquivos publicados (imutáveis);
///  - na MÁQUINA: %LOCALAPPDATA%\HowdenSV\&lt;versão&gt;\ guarda a cópia local.
/// Ao abrir: versão local já existe → abre na hora; versão nova na rede →
/// copia com barra de % e abre. Ou seja: publicou na rede, todo mundo
/// recebe na próxima vez que abrir o programa.
/// </summary>
public sealed class FormAtualizacao : Form
{
    private const string ExeDoApp = "HowdenServicos.exe";

    private readonly ProgressBar _barra = new() { Style = ProgressBarStyle.Continuous, Minimum = 0, Maximum = 100 };
    private readonly Label _status = new() { Text = "Verificando a versão…", ForeColor = Color.FromArgb(70, 80, 110) };
    private readonly Label _pct = new() { Text = "0%", TextAlign = ContentAlignment.MiddleRight, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };

    public FormAtualizacao()
    {
        FormBorderStyle = FormBorderStyle.None;
        StartPosition = FormStartPosition.CenterScreen;
        Size = new Size(480, 150);
        BackColor = Color.White;
        TopMost = true;
        Font = new Font("Segoe UI", 9f);

        var faixa = new Panel { Dock = DockStyle.Top, Height = 54, BackColor = Color.FromArgb(0, 71, 133) };
        faixa.Controls.Add(new Label
        {
            Text = "SV · Propostas de Serviço",
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 13f, FontStyle.Bold),
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(20, 0, 0, 0),
        });
        Controls.Add(faixa);

        _status.SetBounds(20, 70, 370, 20);
        _pct.SetBounds(392, 70, 66, 20);
        _barra.SetBounds(20, 96, 438, 14);
        Controls.Add(_status);
        Controls.Add(_pct);
        Controls.Add(_barra);
        Paint += (_, e) => e.Graphics.DrawRectangle(new Pen(Color.FromArgb(213, 218, 228)), 0, 0, Width - 1, Height - 1);

        Shown += async (_, _) => await ExecutarAsync();
    }

    private void Reportar(int pct, string msg)
    {
        if (IsDisposed) return;
        if (InvokeRequired) { BeginInvoke(() => Reportar(pct, msg)); return; }
        var v = Math.Clamp(pct, 0, 100);
        if (v < 100) _barra.Value = v + 1;
        _barra.Value = v;
        _pct.Text = $"{v}%";
        _status.Text = msg;
    }

    private async Task ExecutarAsync()
    {
        try
        {
            Reportar(5, "Verificando a versão…");

            var raizRede = AppContext.BaseDirectory;                    // \\servidor\...\SV\
            var appRede = Path.Combine(raizRede, "app");
            var raizLocal = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "HowdenSV");
            Directory.CreateDirectory(raizLocal);

            string? versao = null;
            try { versao = File.ReadAllText(Path.Combine(appRede, "versao.txt")).Trim(); }
            catch { /* rede fora do ar: cai no plano B abaixo */ }

            string exe;
            if (string.IsNullOrWhiteSpace(versao))
            {
                // Sem acesso à rede: abre a última versão já instalada na máquina.
                var ultima = Directory.GetDirectories(raizLocal, "v*")
                    .OrderByDescending(d => d, StringComparer.OrdinalIgnoreCase)
                    .FirstOrDefault(d => File.Exists(Path.Combine(d, ExeDoApp)));
                if (ultima is null)
                    throw new Exception(
                        "Sem acesso à pasta de rede do SV e nenhuma versão instalada nesta máquina ainda.\n" +
                        "Conecte-se à rede e tente de novo.");
                Reportar(70, "Sem rede — abrindo a última versão instalada…");
                exe = Path.Combine(ultima, ExeDoApp);
            }
            else
            {
                var destino = Path.Combine(raizLocal, versao);
                exe = Path.Combine(destino, ExeDoApp);
                if (!File.Exists(exe))
                {
                    Reportar(8, $"Atualizando para a versão {versao}…");
                    await Task.Run(() => CopiarComProgresso(Path.Combine(appRede, versao), destino));
                    LimparVersoesAntigas(raizLocal, versao);
                }
                else
                {
                    Reportar(85, "Versão em dia — abrindo…");
                }
            }

            Reportar(95, "Abrindo o SV…");
            Process.Start(new ProcessStartInfo(exe)
            {
                WorkingDirectory = Path.GetDirectoryName(exe)!,
                UseShellExecute = true,
            });
            Reportar(100, "Pronto!");
            await Task.Delay(300);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Não foi possível abrir o SV:\n\n" + ex.Message,
                "SV", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        Close();
    }

    /// <summary>
    /// Copia a pasta da versão para uma pasta temporária local (barra de 10% a
    /// 90%, proporcional aos bytes) e só no fim renomeia para o nome definitivo —
    /// se a cópia falhar no meio, nada fica pela metade.
    /// </summary>
    private void CopiarComProgresso(string origem, string destino)
    {
        var arquivos = Directory.GetFiles(origem, "*", SearchOption.AllDirectories);
        if (arquivos.Length == 0)
            throw new Exception($"A pasta da versão está vazia na rede: {origem}");
        var totalBytes = arquivos.Sum(f => new FileInfo(f).Length);

        var tmp = destino + ".baixando";
        if (Directory.Exists(tmp)) Directory.Delete(tmp, recursive: true);

        long copiados = 0;
        foreach (var arq in arquivos)
        {
            var relativo = Path.GetRelativePath(origem, arq);
            var alvo = Path.Combine(tmp, relativo);
            Directory.CreateDirectory(Path.GetDirectoryName(alvo)!);
            File.Copy(arq, alvo, overwrite: true);
            copiados += new FileInfo(arq).Length;
            var pct = 10 + (int)(copiados * 80 / Math.Max(totalBytes, 1));
            Reportar(pct, $"Atualizando… {copiados / 1048576} de {totalBytes / 1048576} MB");
        }

        if (Directory.Exists(destino)) Directory.Delete(destino, recursive: true);
        Directory.Move(tmp, destino);
    }

    /// <summary>Apaga versões locais antigas (mantém a atual e a anterior).</summary>
    private static void LimparVersoesAntigas(string raizLocal, string versaoAtual)
    {
        try
        {
            var antigas = Directory.GetDirectories(raizLocal, "v*")
                .Where(d => !d.EndsWith(versaoAtual, StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(d => d, StringComparer.OrdinalIgnoreCase)
                .Skip(1);   // deixa a anterior como reserva
            foreach (var d in antigas)
                Directory.Delete(d, recursive: true);
        }
        catch { /* limpeza é melhor esforço — versões antigas não atrapalham */ }
    }
}
