using System.Diagnostics.CodeAnalysis;

using AIStudio.Tools.Web;

namespace AIStudio.Tools.ToolCallingSystem.ToolCallingImplementations;

/// <summary>
/// The address of a site of the organization which a tool was configured with, such as its wiki
/// or its SharePoint.
/// </summary>
/// <remarks>
/// A search tool sends the query of the model and the sign-in of the user to this address, so two
/// questions decide where both may go: whether the configured value is an address a tool may use
/// at all, and whether another address still lies within it. The tools which search a configured
/// site share the answers, so that none of them judges a redirect more loosely than the others.
/// </remarks>
internal static class ConfiguredSiteUrl
{
    /// <summary>
    /// Reads the configured address of a site.
    /// </summary>
    /// <remarks>
    /// Only a plain HTTPS address counts: plain HTTP would expose the sign-in, credentials in the
    /// address would travel with every request, and a query or fragment has no place in the root
    /// of a site. The result ends with exactly one slash, so that a path resolved against it lands
    /// below it instead of replacing its last segment.
    /// </remarks>
    /// <param name="value">The value as configured.</param>
    /// <param name="baseUrl">The address, when the value is one a tool may use.</param>
    /// <returns>True when the value is a plain HTTPS address.</returns>
    public static bool TryParse(string? value, [NotNullWhen(true)] out Uri? baseUrl)
    {
        baseUrl = null;
        if (!Uri.TryCreate(value?.Trim(), UriKind.Absolute, out var uri) ||
            uri.Scheme is not "https" ||
            !string.IsNullOrWhiteSpace(uri.UserInfo) ||
            !string.IsNullOrWhiteSpace(uri.Query) ||
            !string.IsNullOrWhiteSpace(uri.Fragment))
            return false;

        baseUrl = new Uri(uri.AbsoluteUri.TrimEnd('/') + '/');
        return true;
    }

    public static bool IsHost(Uri baseUrl, string host) => WebHostHelper.Normalize(host) == WebHostHelper.Normalize(baseUrl.Host);

    /// <summary>
    /// Whether an address lies within the configured site: same scheme, host, and port, and a path
    /// below the one of the site.
    /// </summary>
    public static bool IsWithin(Uri baseUrl, Uri url) =>
        url.Scheme == baseUrl.Scheme &&
        IsHost(baseUrl, url.Host) &&
        url.Port == baseUrl.Port &&
        url.AbsolutePath.StartsWith(baseUrl.AbsolutePath, StringComparison.Ordinal);
}
