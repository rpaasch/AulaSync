namespace AulaSync.Core;

public class AulaException(string message, Exception? inner = null) : Exception(message, inner);

public sealed class SessionExpiredException() : AulaException("Aula-sessionen er udløbet");

// Aula svarer 403 / status.code 401 eller 403: ingen adgang (fx ikke medlem af gruppen, ingen ret til et lokale). Ikke et udløbet login.
public sealed class ForbiddenException(string message = "Ingen adgang i Aula") : AulaException(message);

// Aula status.code 451 (adgang ikke givet endnu) eller 452 (bruger deaktiveret): hele kontoen, ingen genforsøg.
public sealed class AccessDeniedException(string message) : AulaException(message);
