// Utilitários chamados pelo Blazor via JS interop.
window.appPrint = () => window.print();

window.appDownload = (fileName, mime, base64) => {
    const a = document.createElement('a');
    a.href = `data:${mime};base64,${base64}`;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
};

// Rascunho da proposta guardado no navegador: o trabalho em andamento sobrevive
// a um F5 ou a abrir /servicos/pricing direto na barra de endereços.
const RASCUNHO = 'howden-servicos-rascunho';

window.appSaveDraft = (json) => {
    try { localStorage.setItem(RASCUNHO, json); } catch { /* modo privativo: ignora */ }
};

window.appLoadDraft = () => {
    try { return localStorage.getItem(RASCUNHO) || ''; } catch { return ''; }
};

window.appClearDraft = () => {
    try { localStorage.removeItem(RASCUNHO); } catch { /* ignora */ }
};
