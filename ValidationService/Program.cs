using System.Text.RegularExpressions;

var builder = WebApplication.CreateBuilder(args);

var app = builder.Build();

// Enforce HTTPS. Entra CIAM authenticates the call via a Bearer token.
app.UseHttpsRedirection();

// Hardcoded list of allowed email name (local-part) patterns.
var allowedPatterns = new[]
{
    @"^admin$",
    @"^support$",
    @"^info$",
    @"^.*\.team$",
    @"^dev\d*$",
    @"^grr$",
    @"^test$",
    @"^ayoubseddiki132$",
    @"^ingenieur\.ayoub\.seddiki$",
};

var compiledPatterns = allowedPatterns
    .Select(p => new Regex(p, RegexOptions.IgnoreCase | RegexOptions.Compiled))
    .ToArray();

app.MapPost("/api/validate", (AttributeCollectionStartRequest request) =>
{
    var identities = request.Data?.UserSignUpInfo?.Identities;
    var email = identities is { Length: > 0 } ? identities[0]?.IssuerAssignedId : null;

    var allowed = false;
    if (!string.IsNullOrWhiteSpace(email))
    {
        var atIndex = email.IndexOf('@');
        var localPart = atIndex > 0 ? email[..atIndex] : email;
        allowed = compiledPatterns.Any(p => p.IsMatch(localPart));
    }

    var action = allowed
        ? new ResponseAction
        {
            ODataType = "microsoft.graph.attributeCollectionStart.continueWithDefaultBehavior"
        }
        : new ResponseAction
        {
            ODataType = "microsoft.graph.attributeCollectionStart.showBlockPage",
            Message = "Your email address is not authorized to register. Please contact support."
        };

    return Results.Ok(new AttributeCollectionStartResponse
    {
        Data = new ResponseData
        {
            ODataType = "microsoft.graph.onAttributeCollectionStartResponseData",
            Actions = new[] { action }
        }
    });
});

app.Run();

// ----- Request model (Entra CIAM OnAttributeCollectionStart) -----

record AttributeCollectionStartRequest(RequestData? Data);

record RequestData(UserSignUpInfo? UserSignUpInfo);

record UserSignUpInfo(Identity[]? Identities);

record Identity(string? IssuerAssignedId);

// ----- Response model (Entra CIAM OnAttributeCollectionStart) -----

class AttributeCollectionStartResponse
{
    [System.Text.Json.Serialization.JsonPropertyName("data")]
    public ResponseData Data { get; set; } = default!;
}

class ResponseData
{
    [System.Text.Json.Serialization.JsonPropertyName("@odata.type")]
    public string ODataType { get; set; } = default!;

    [System.Text.Json.Serialization.JsonPropertyName("actions")]
    public ResponseAction[] Actions { get; set; } = default!;
}

class ResponseAction
{
    [System.Text.Json.Serialization.JsonPropertyName("@odata.type")]
    public string ODataType { get; set; } = default!;

    [System.Text.Json.Serialization.JsonPropertyName("title")]
    [System.Text.Json.Serialization.JsonIgnore(Condition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull)]
    public string? Title { get; set; }

    [System.Text.Json.Serialization.JsonPropertyName("message")]
    [System.Text.Json.Serialization.JsonIgnore(Condition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; set; }
}
