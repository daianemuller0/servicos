namespace HowdenServicos.Poc.Data;

/// <summary>As línguas da proposta completa (Word): pt, en, es — a partir do idioma da proposta.</summary>
public static class IdiomasDaProposta
{
    public const string Portugues = "pt";
    public const string Ingles = "en";
    public const string Espanhol = "es";

    /// <summary>"English" → en, "Español" → es, o resto → pt (aceita também os códigos).</summary>
    public static string Normalizar(string? idioma) => idioma switch
    {
        "English" or Ingles => Ingles,
        "Español" or Espanhol => Espanhol,
        _ => Portugues,
    };
}
