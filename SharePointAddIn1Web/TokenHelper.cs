using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Globalization;
using System.IdentityModel.Tokens;
using System.IdentityModel.Tokens.Jwt;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Claims;
using System.Security.Cryptography.X509Certificates;
using System.Security.Principal;
using System.ServiceModel;
using System.Text;
using System.Web;
using System.Web.Configuration;
using System.Web.Script.Serialization;
using SigningCredentials = Microsoft.IdentityModel.Tokens.SigningCredentials;
using SymmetricSecurityKey = Microsoft.IdentityModel.Tokens.SymmetricSecurityKey;
using TokenValidationParameters = Microsoft.IdentityModel.Tokens.TokenValidationParameters;
using X509SigningCredentials = Microsoft.IdentityModel.Tokens.X509SigningCredentials;

namespace SharePointAddIn1Web
{
    public static class TokenHelper
    {
        #region 公共字段

        /// <summary>
        /// SharePoint 主体。
        /// </summary>
        public const string SharePointPrincipal = "00000003-0000-0ff1-ce00-000000000000";

        /// <summary>
        /// HighTrust 访问令牌的生存期(12 小时)。
        /// </summary>
        public static readonly TimeSpan HighTrustAccessTokenLifetime = TimeSpan.FromHours(12.0);

        #endregion public fields

        #region 公共方法

        /// <summary>
        ///通过查找已知参数名称，从指定请求中检索上下文标记字符串
        ///从而从指定请求中检索上下文标记字符串。 如果找不到上下文标记，则返回 null。
        /// </summary>
        /// <param name="request">要从中查找上下文标记的 HttpRequest</param>
        /// <returns>上下文标记字符串</returns>
        public static string GetContextTokenFromRequest(HttpRequest request)
        {
            return GetContextTokenFromRequest(new HttpRequestWrapper(request));
        }

        /// <summary>
        ///通过查找已知参数名称，从指定请求中检索上下文标记字符串
        ///从而从指定请求中检索上下文标记字符串。 如果找不到上下文标记，则返回 null。
        /// </summary>
        /// <param name="request">要从中查找上下文标记的 HttpRequest</param>
        /// <returns>上下文标记字符串</returns>
        public static string GetContextTokenFromRequest(HttpRequestBase request)
        {
            string[] paramNames = { "AppContext", "AppContextToken", "AccessToken", "SPAppToken" };
            foreach (string paramName in paramNames)
            {
                if (!string.IsNullOrEmpty(request.Form[paramName]))
                {
                    return request.Form[paramName];
                }
                if (!string.IsNullOrEmpty(request.QueryString[paramName]))
                {
                    return request.QueryString[paramName];
                }
            }
            return null;
        }

        /// <summary>
        /// 基于参数验证指定的上下文标识字符串用于此应用程序
        ///在 web.config 中指定。web.config 中使用的参数用于验证，包括 ClientId，
        ///HostedAppHostNameOverride、HostedAppHostName、ClientSecret 和 Realm (如果指定它)。如果存在 HostedAppHostNameOverride，
        ///则使用其进行验证。否则，如果 <paramref name="appHostName"/> 不是
        ///如果为 null，则使用它而非 web.config 的 HostedAppHostName 进行验证。如果该标记无效，则
        ///引发异常。 如果标记有效，则根据标记内容更新 TokenHelper 的静态 STS 元数据 URL，
        ///并且返回基于上下文标记的 JwtSecurityToken。
        /// </summary>
        /// <param name="contextTokenString">要验证的上下文标记</param>
        /// <param name="appHostName">URL 颁发机构，包含域名系统 (DNS) 主机名或 IP 地址及端口号，用于标记的访问群体验证。
        ///如果为 null，则改用 HostedAppHostName web.config 设置。如果存在 HostedAppHostNameOverride web.config 设置，则将使用该设置
        ///代替 <paramref name="appHostName"/> 进行验证。</param>
        /// <returns>基于上下文标记的 JwtSecurityToken。</returns>
        public static SharePointContextToken ReadAndValidateContextToken(string contextTokenString, string appHostName = null)
        {
            List<SymmetricSecurityKey> securityKeys = new List<SymmetricSecurityKey>
            {
                new SymmetricSecurityKey(Convert.FromBase64String(ClientSecret))
            };

            if (!string.IsNullOrEmpty(SecondaryClientSecret))
            {
                securityKeys.Add(new SymmetricSecurityKey(Convert.FromBase64String(SecondaryClientSecret)));
            }

            JwtSecurityTokenHandler tokenHandler = CreateJwtSecurityTokenHandler();
            TokenValidationParameters parameters = new TokenValidationParameters
            {
                ValidateIssuer = false,
                ValidateAudience = false, // 以下已验证
                IssuerSigningKeys = securityKeys // 验证签名
            };

            tokenHandler.ValidateToken(contextTokenString, parameters, out Microsoft.IdentityModel.Tokens.SecurityToken securityToken);
            SharePointContextToken token = SharePointContextToken.Create(securityToken as JwtSecurityToken);

            string stsAuthority = (new Uri(token.SecurityTokenServiceUri)).Authority;
            int firstDot = stsAuthority.IndexOf('.');

            GlobalEndPointPrefix = stsAuthority.Substring(0, firstDot);
            AcsHostUrl = stsAuthority.Substring(firstDot + 1);

            string[] acceptableAudiences;
            if (!String.IsNullOrEmpty(HostedAppHostNameOverride))
            {
                acceptableAudiences = HostedAppHostNameOverride.Split(';');
            }
            else if (appHostName == null)
            {
                acceptableAudiences = new[] { HostedAppHostName };
            }
            else
            {
                acceptableAudiences = new[] { appHostName };
            }

            bool validationSuccessful = false;
            string realm = Realm ?? token.Realm;
            foreach (var audience in acceptableAudiences)
            {
                string principal = GetFormattedPrincipal(ClientId, audience, realm);
                if (token.Audiences.First<string>(item => StringComparer.OrdinalIgnoreCase.Equals(item, principal)) != null)
                {
                    validationSuccessful = true;
                    break;
                }
            }

            if (!validationSuccessful)
            {
                throw new AudienceUriValidationFailedException(
                    String.Format(CultureInfo.CurrentCulture,
                    "\"{0}\" is not the intended audience \"{1}\"", String.Join(";", acceptableAudiences),
                    String.Join(";", token.Audiences.ToArray<string>())));
            }

            return token;
        }

        /// <summary>
        /// 从 ACS 检索访问令牌，以在指定 targetHost 中调用指定上下文标记 
        /// 的源。 必须为发送上下文标记的主体注册 targetHost。
        /// </summary>
        /// <param name="contextToken">由预期的访问令牌群体颁发的上下文标记</param>
        /// <param name="targetHost">的目标主体名称</param>
        /// <returns>带有与上下文标记源匹配的访问群体的访问令牌</returns>
        public static OAuthTokenResponse GetAccessToken(SharePointContextToken contextToken, string targetHost)
        {
            string targetPrincipalName = contextToken.TargetPrincipalName;

            // 从上下文标记提取 refreshToken
            string refreshToken = contextToken.RefreshToken;

            if (String.IsNullOrEmpty(refreshToken))
            {
                return null;
            }

            string targetRealm = Realm ?? contextToken.Realm;

            return GetAccessToken(refreshToken,
                                  targetPrincipalName,
                                  targetHost,
                                  targetRealm);
        }

        /// <summary>
        /// 使用指定的授权代码从 ACS 检索访问令牌，以调用指定主体
        ///在指定的 targetHost。必须为目标主体注册 targetHost。如果指定的领域为 
        /// 空，则改用 web.config 中的 "Realm" 设置。
        /// </summary>
        /// <param name="authorizationCode">用于交换访问令牌的授权代码</param>
        /// <param name="targetPrincipalName">要检索目标主体 URL 颁发机构的访问令牌</param>
        /// <param name="targetHost">的目标主体名称</param>
        /// <param name="targetRealm">用于访问令牌的 nameid 和访问群体的 Realm</param>
        /// <param name="redirectUri">已为此外接程序注册重定向 URI</param>
        /// <returns>带有目标主体访问群体的访问令牌</returns>
        public static OAuthTokenResponse GetAccessToken(
            string authorizationCode,
            string targetPrincipalName,
            string targetHost,
            string targetRealm,
            Uri redirectUri)
        {
            if (targetRealm == null)
            {
                targetRealm = Realm;
            }

            string resource = GetFormattedPrincipal(targetPrincipalName, targetHost, targetRealm);
            string clientId = GetFormattedPrincipal(ClientId, null, targetRealm);
            string acsUri = AcsMetadataParser.GetStsUrl(targetRealm);
            OAuthTokenResponse oauthResponse = null;

            try
            {
                oauthResponse = OAuthClient.GetAccessTokenWithAuthorizationCode(acsUri, clientId, ClientSecret,
                    authorizationCode, redirectUri.AbsoluteUri, resource);
            }
            catch (WebException wex)
            {
                if (!string.IsNullOrEmpty(SecondaryClientSecret))
                {
                    oauthResponse = OAuthClient.GetAccessTokenWithAuthorizationCode(acsUri, clientId, SecondaryClientSecret,
                        authorizationCode, redirectUri.AbsoluteUri, resource);
                }
                else
                {
                    using (StreamReader sr = new StreamReader(wex.Response.GetResponseStream()))
                    {
                        string responseText = sr.ReadToEnd();
                        throw new WebException(wex.Message + " - " + responseText, wex);
                    }
                }
            }

            return oauthResponse;
        }

        /// <summary>
        /// 使用指定的刷新标记从 ACS 检索访问令牌，以调用指定主体
        ///在指定的 targetHost。必须为目标主体注册 targetHost。如果指定的领域为 
        /// 空，则改用 web.config 中的 "Realm" 设置。
        /// </summary>
        /// <param name="refreshToken">用于交换访问令牌的刷新标记</param>
        /// <param name="targetPrincipalName">要检索目标主体 URL 颁发机构的访问令牌</param>
        /// <param name="targetHost">的目标主体名称</param>
        /// <param name="targetRealm">用于访问令牌的 nameid 和访问群体的 Realm</param>
        /// <returns>带有目标主体访问群体的访问令牌</returns>
        public static OAuthTokenResponse GetAccessToken(
            string refreshToken,
            string targetPrincipalName,
            string targetHost,
            string targetRealm)
        {
            if (targetRealm == null)
            {
                targetRealm = Realm;
            }

            string resource = GetFormattedPrincipal(targetPrincipalName, targetHost, targetRealm);
            string clientId = GetFormattedPrincipal(ClientId, null, targetRealm);
            string acsUri = AcsMetadataParser.GetStsUrl(targetRealm);
            OAuthTokenResponse oauthResponse = null;

            try
            {
                oauthResponse = OAuthClient.GetAccessTokenWithRefreshToken(acsUri, clientId, ClientSecret, refreshToken, resource);
            }
            catch (WebException wex)
            {
                if (!string.IsNullOrEmpty(SecondaryClientSecret))
                {
                    oauthResponse = OAuthClient.GetAccessTokenWithRefreshToken(acsUri, clientId, SecondaryClientSecret, refreshToken, resource);
                }
                else
                {
                    using (StreamReader sr = new StreamReader(wex.Response.GetResponseStream()))
                    {
                        string responseText = sr.ReadToEnd();
                        throw new WebException(wex.Message + " - " + responseText, wex);
                    }
                }
            }

            return oauthResponse;
        }

        /// <summary>
        /// 从 ACS 检索只允许应用程序使用的访问令牌，以调用指定主体
        ///在指定的 targetHost。必须为目标主体注册 targetHost。如果指定的领域为 
        /// 空，则改用 web.config 中的 "Realm" 设置。
        /// </summary>
        /// <param name="targetPrincipalName">要检索目标主体 URL 颁发机构的访问令牌</param>
        /// <param name="targetHost">的目标主体名称</param>
        /// <param name="targetRealm">用于访问令牌的 nameid 和访问群体的 Realm</param>
        /// <returns>带有目标主体访问群体的访问令牌</returns>
        public static OAuthTokenResponse GetAppOnlyAccessToken(
            string targetPrincipalName,
            string targetHost,
            string targetRealm)
        {

            if (targetRealm == null)
            {
                targetRealm = Realm;
            }

            string resource = GetFormattedPrincipal(targetPrincipalName, targetHost, targetRealm);
            string clientId = GetFormattedPrincipal(ClientId, HostedAppHostName, targetRealm);
            string acsUri = AcsMetadataParser.GetStsUrl(targetRealm);
            OAuthTokenResponse oauthResponse = null;

            try
            {
                oauthResponse = OAuthClient.GetAccessTokenWithClientCredentials(acsUri, clientId, ClientSecret, resource);
            }
            catch (WebException wex)
            {
                if (!string.IsNullOrEmpty(SecondaryClientSecret))
                {
                    oauthResponse = OAuthClient.GetAccessTokenWithClientCredentials(acsUri, clientId, SecondaryClientSecret, resource);
                }
                else
                {
                    using (StreamReader sr = new StreamReader(wex.Response.GetResponseStream()))
                    {
                        string responseText = sr.ReadToEnd();
                        throw new WebException(wex.Message + " - " + responseText, wex);
                    }
                }
            }

            return oauthResponse;
        }

        /// <summary>
        /// 根据远程事件接收器的属性创建客户端上下文
        /// </summary>
        /// <param name="properties">远程事件接收器的属性</param>
        /// <returns>ClientContext 准备调用发起事件的 Web</returns>
        public static ClientContext CreateRemoteEventReceiverClientContext(SPRemoteEventProperties properties)
        {
            Uri sharepointUrl;
            if (properties.ListEventProperties != null)
            {
                sharepointUrl = new Uri(properties.ListEventProperties.WebUrl);
            }
            else if (properties.ItemEventProperties != null)
            {
                sharepointUrl = new Uri(properties.ItemEventProperties.WebUrl);
            }
            else if (properties.WebEventProperties != null)
            {
                sharepointUrl = new Uri(properties.WebEventProperties.FullUrl);
            }
            else
            {
                return null;
            }

            if (IsHighTrustApp())
            {
                return GetS2SClientContextWithWindowsIdentity(sharepointUrl, null);
            }

            return CreateAcsClientContextForUrl(properties, sharepointUrl);
        }

        /// <summary>
        /// 基于外接程序事件的属性创建客户端上下文
        /// </summary>
        /// <param name="properties">外接程序事件的属性</param>
        /// <param name="useAppWeb">如果定位到应用程序 Web，则为 true；如果定位到主机 Web，则为 false</param>
        /// <returns>ClientContext 准备调用应用程序 Web 或父网站</returns>
        public static ClientContext CreateAppEventClientContext(SPRemoteEventProperties properties, bool useAppWeb)
        {
            if (properties.AppEventProperties == null)
            {
                return null;
            }

            Uri sharepointUrl = useAppWeb ? properties.AppEventProperties.AppWebFullUrl : properties.AppEventProperties.HostWebFullUrl;
            if (IsHighTrustApp())
            {
                return GetS2SClientContextWithWindowsIdentity(sharepointUrl, null);
            }

            return CreateAcsClientContextForUrl(properties, sharepointUrl);
        }

        /// <summary>
        /// 使用指定的授权代码从 ACS 检索访问令牌，并使用该访问令牌 
        /// 创建客户端上下文
        /// </summary>
        /// <param name="targetUrl">目标 SharePoint 网站的 URL</param>
        /// <param name="authorizationCode">从 ACS 检索访问令牌时使用的授权代码</param>
        /// <param name="redirectUri">已为此外接程序注册重定向 URI</param>
        /// <returns>ClientContext 准备使用有效访问令牌调用 targetUrl</returns>
        public static ClientContext GetClientContextWithAuthorizationCode(
            string targetUrl,
            string authorizationCode,
            Uri redirectUri)
        {
            return GetClientContextWithAuthorizationCode(targetUrl, SharePointPrincipal, authorizationCode, GetRealmFromTargetUrl(new Uri(targetUrl)), redirectUri);
        }

        /// <summary>
        /// 使用指定的授权代码从 ACS 检索访问令牌，并使用该访问令牌 
        /// 创建客户端上下文
        /// </summary>
        /// <param name="targetUrl">目标 SharePoint 网站的 URL</param>
        /// <param name="targetPrincipalName">目标 SharePoint 主体的名称</param>
        /// <param name="authorizationCode">从 ACS 检索访问令牌时使用的授权代码</param>
        /// <param name="targetRealm">用于访问令牌的 nameid 和访问群体的 Realm</param>
        /// <param name="redirectUri">已为此外接程序注册重定向 URI</param>
        /// <returns>ClientContext 准备使用有效访问令牌调用 targetUrl</returns>
        public static ClientContext GetClientContextWithAuthorizationCode(
            string targetUrl,
            string targetPrincipalName,
            string authorizationCode,
            string targetRealm,
            Uri redirectUri)
        {
            Uri targetUri = new Uri(targetUrl);

            string accessToken =
                GetAccessToken(authorizationCode, targetPrincipalName, targetUri.Authority, targetRealm, redirectUri).AccessToken;

            return GetClientContextWithAccessToken(targetUrl, accessToken);
        }

        /// <summary>
        /// 使用指定的访问令牌创建客户端上下文
        /// </summary>
        /// <param name="targetUrl">目标 SharePoint 网站的 URL</param>
        /// <param name="accessToken">调用指定 targetUrl 时使用的访问令牌</param>
        /// <returns>ClientContext 准备使用指定访问令牌调用 targetUrl</returns>
        public static ClientContext GetClientContextWithAccessToken(string targetUrl, string accessToken)
        {
            ClientContext clientContext = new ClientContext(targetUrl);

            clientContext.AuthenticationMode = ClientAuthenticationMode.Anonymous;
            clientContext.FormDigestHandlingEnabled = false;
            clientContext.ExecutingWebRequest +=
                delegate (object oSender, WebRequestEventArgs webRequestEventArgs)
                {
                    webRequestEventArgs.WebRequestExecutor.RequestHeaders["Authorization"] =
                        "Bearer " + accessToken;
                };

            return clientContext;
        }

        /// <summary>
        /// 使用指定的上下文标记从 ACS 检索访问令牌，并使用该令牌创建
        /// 客户端上下文
        /// </summary>
        /// <param name="targetUrl">目标 SharePoint 网站的 URL</param>
        /// <param name="contextTokenString">从目标 SharePoint 网站接收的上下文标记</param>
        /// <param name="appHostUrl">托管外接程序的 URL 授权。如果它为 NULL，则改用 HostedAppHostName 中的值
        /// 中 HostedAppHostName 的值</param>
        /// <returns>ClientContext 准备使用有效访问令牌调用 targetUrl</returns>
        public static ClientContext GetClientContextWithContextToken(
            string targetUrl,
            string contextTokenString,
            string appHostUrl)
        {
            SharePointContextToken contextToken = ReadAndValidateContextToken(contextTokenString, appHostUrl);

            Uri targetUri = new Uri(targetUrl);

            string accessToken = GetAccessToken(contextToken, targetUri.Authority).AccessToken;

            return GetClientContextWithAccessToken(targetUrl, accessToken);
        }

        /// <summary>
        /// 返回一个 SharePoint URL，外接程序应将浏览器重定向到该 URL，以请求许可并返回
        ///授权代码。
        /// </summary>
        /// <param name="contextUrl">SharePoint 网站的绝对 URL</param>
        /// <param name="scope">以“速记”格式从 SharePoint 网站进行请求的空格分隔权限
        /// (例如 "Web.Read Site.Write")</param>
        /// <returns>SharePoint 网站 OAuth 授权页面的 URL</returns>
        public static string GetAuthorizationUrl(string contextUrl, string scope)
        {
            return string.Format(
                "{0}{1}?IsDlg=1&client_id={2}&scope={3}&response_type=code",
                EnsureTrailingSlash(contextUrl),
                AuthorizationPage,
                ClientId,
                scope);
        }

        /// <summary>
        /// 返回一个 SharePoint URL，外接程序应将浏览器重定向到该 URL，以请求许可并返回
        ///授权代码。
        /// </summary>
        /// <param name="contextUrl">SharePoint 网站的绝对 URL</param>
        /// <param name="scope">以“速记”格式从 SharePoint 网站进行请求的空格分隔权限
        /// (例如 "Web.Read Site.Write")</param>
        /// <param name="redirectUri">在获得同意后，SharePoint 应将浏览器重定向到的 URI
        ///已授予</param>
        /// <returns>SharePoint 网站 OAuth 授权页面的 URL</returns>
        public static string GetAuthorizationUrl(string contextUrl, string scope, string redirectUri)
        {
            return string.Format(
                "{0}{1}?IsDlg=1&client_id={2}&scope={3}&response_type=code&redirect_uri={4}",
                EnsureTrailingSlash(contextUrl),
                AuthorizationPage,
                ClientId,
                scope,
                redirectUri);
        }

        /// <summary>
        /// 返回一个 SharePoint URL，外接程序应将浏览器重定向到该 URL，以请求新的上下文标记。
        /// </summary>
        /// <param name="contextUrl">SharePoint 网站的绝对 URL</param>
        /// <param name="redirectUri">SharePoint 应使用上下文标记将浏览器重定向到的 URL</param>
        /// <returns>SharePoint 网站的上下文标记重定向页面的 URL</returns>
        public static string GetAppContextTokenRequestUrl(string contextUrl, string redirectUri)
        {
            return string.Format(
                "{0}{1}?client_id={2}&redirect_uri={3}",
                EnsureTrailingSlash(contextUrl),
                RedirectPage,
                ClientId,
                redirectUri);
        }

        /// <summary>
        ///检索由应用程序的专有证书签名的 S2S 访问令牌
        /// WindowsIdentity 并用于 targetApplicationUri 处的 SharePoint。如果未指定领域
        /// Realm，将向 targetApplicationUri 发出身份验证质询以发现它。
        /// </summary>
        /// <param name="targetApplicationUri">目标 SharePoint 网站的 URL</param>
        /// <param name="identity">代表用户创建访问令牌的 Windows 标识</param>
        /// <returns>带有目标主体访问群体的访问令牌</returns>
        public static string GetS2SAccessTokenWithWindowsIdentity(
            Uri targetApplicationUri,
            WindowsIdentity identity)
        {
            string realm = string.IsNullOrEmpty(Realm) ? GetRealmFromTargetUrl(targetApplicationUri) : Realm;

            Claim[] claims = identity != null ? GetClaimsWithWindowsIdentity(identity) : null;

            return GetS2SAccessTokenWithClaims(targetApplicationUri.Authority, realm, claims);
        }

        /// <summary>
        ///使用由应用程序的专有证书签名的访问令牌检索 S2S 客户端上下文
        ///代表指定的 WindowsIdentity 并拟用于 targetApplicationUri 处的应用程序
        /// targetRealm。如果 web.config 中未指定领域，则会将身份验证质询发给
        /// 以发现它。
        /// </summary>
        /// <param name="targetApplicationUri">目标 SharePoint 网站的 URL</param>
        /// <param name="identity">代表用户创建访问令牌的 Windows 标识</param>
        /// <returns>使用带有目标应用程序访问群体的访问令牌的 ClientContext</returns>
        public static ClientContext GetS2SClientContextWithWindowsIdentity(
            Uri targetApplicationUri,
            WindowsIdentity identity)
        {
            string realm = string.IsNullOrEmpty(Realm) ? GetRealmFromTargetUrl(targetApplicationUri) : Realm;

            Claim[] claims = identity != null ? GetClaimsWithWindowsIdentity(identity) : null;

            string accessToken = GetS2SAccessTokenWithClaims(targetApplicationUri.Authority, realm, claims);

            return GetClientContextWithAccessToken(targetApplicationUri.ToString(), accessToken);
        }

        /// <summary>
        /// 从 SharePoint 获取身份验证领域
        /// </summary>
        /// <param name="targetApplicationUri">目标 SharePoint 网站的 URL</param>
        /// <returns>领域 GUID 的字符串表示形式</returns>
        public static string GetRealmFromTargetUrl(Uri targetApplicationUri)
        {
            WebRequest request = WebRequest.Create(targetApplicationUri + "/_vti_bin/client.svc");
            request.Headers.Add("Authorization: Bearer ");

            try
            {
                using (request.GetResponse())
                {
                }
            }
            catch (WebException e)
            {
                if (e.Response == null)
                {
                    return null;
                }

                string bearerResponseHeader = e.Response.Headers["WWW-Authenticate"];
                if (string.IsNullOrEmpty(bearerResponseHeader))
                {
                    return null;
                }

                const string bearer = "Bearer realm=\"";
                int bearerIndex = bearerResponseHeader.IndexOf(bearer, StringComparison.Ordinal);
                if (bearerIndex < 0)
                {
                    return null;
                }

                int realmIndex = bearerIndex + bearer.Length;

                if (bearerResponseHeader.Length >= realmIndex + 36)
                {
                    string targetRealm = bearerResponseHeader.Substring(realmIndex, 36);

                    Guid realmGuid;

                    if (Guid.TryParse(targetRealm, out realmGuid))
                    {
                        return targetRealm;
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// 确定这是否是一个高信任外接程序。
        /// </summary>
        /// <returns>若是一个高信任外接程序，则为 True。</returns>
        public static bool IsHighTrustApp()
        {
            return SigningCredentials != null;
        }

        /// <summary>
        /// 确保指定的 URL 在不为 Null 或空时以“/”结束。
        /// </summary>
        /// <param name="url">URL。</param>
        /// <returns>如果 URL 不为 Null 或空，则该 URL 将以“/”结束。</returns>
        public static string EnsureTrailingSlash(string url)
        {
            if (!string.IsNullOrEmpty(url) && url[url.Length - 1] != '/')
            {
                return url + "/";
            }

            return url;
        }

        /// <summary>
        ///返回当前的时期时间(秒)
        /// </summary>
        /// <returns>以秒表示的时期时间</returns>
        public static long EpochTimeNow()
        {
            return (long)(DateTime.UtcNow - new DateTime(1970, 1, 1).ToUniversalTime()).TotalSeconds;
        }

        #endregion

        #region 私有字段

        //
        // 配置常数
        //

        private const string AuthorizationPage = "_layouts/15/OAuthAuthorize.aspx";
        private const string RedirectPage = "_layouts/15/AppRedirect.aspx";
        private const string AcsPrincipalName = "00000001-0000-0000-c000-000000000000";
        private const string AcsMetadataEndPointRelativeUrl = "metadata/json/1";
        private const string S2SProtocol = "OAuth2";
        private const string DelegationIssuance = "DelegationIssuance1.0";
        private const string NameIdentifierClaimType = "nameid";
        private const string TrustedForImpersonationClaimType = "trustedfordelegation";
        private const string ActorTokenClaimType = "actortoken";

        //
        // 环境常数
        //

        private static string GlobalEndPointPrefix = "accounts";
        private static string AcsHostUrl = "accesscontrol.windows.net";

        //
        // 托管外接程序配置
        //
        private static readonly string ClientId = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("ClientId")) ? WebConfigurationManager.AppSettings.Get("HostedAppName") : WebConfigurationManager.AppSettings.Get("ClientId");
        private static readonly string IssuerId = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("IssuerId")) ? ClientId : WebConfigurationManager.AppSettings.Get("IssuerId");
        private static readonly string HostedAppHostNameOverride = WebConfigurationManager.AppSettings.Get("HostedAppHostNameOverride");
        private static readonly string HostedAppHostName = WebConfigurationManager.AppSettings.Get("HostedAppHostName");
        private static readonly string ClientSecret = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("ClientSecret")) ? WebConfigurationManager.AppSettings.Get("HostedAppSigningKey") : WebConfigurationManager.AppSettings.Get("ClientSecret");
        private static readonly string SecondaryClientSecret = WebConfigurationManager.AppSettings.Get("SecondaryClientSecret");
        private static readonly string Realm = WebConfigurationManager.AppSettings.Get("Realm");
        private static readonly string ServiceNamespace = WebConfigurationManager.AppSettings.Get("Realm");

        private static readonly string ClientSigningCertificatePath = WebConfigurationManager.AppSettings.Get("ClientSigningCertificatePath");
        private static readonly string ClientSigningCertificatePassword = WebConfigurationManager.AppSettings.Get("ClientSigningCertificatePassword");
        private static readonly X509Certificate2 ClientCertificate = (string.IsNullOrEmpty(ClientSigningCertificatePath) || string.IsNullOrEmpty(ClientSigningCertificatePassword)) ? null : new X509Certificate2(ClientSigningCertificatePath, ClientSigningCertificatePassword);
        private static readonly X509SigningCredentials SigningCredentials = (ClientCertificate == null) ? null : new X509SigningCredentials(ClientCertificate, Microsoft.IdentityModel.Tokens.SecurityAlgorithms.RsaSha256);

        #endregion

        #region 私有方法

        private static ClientContext CreateAcsClientContextForUrl(SPRemoteEventProperties properties, Uri sharepointUrl)
        {
            string contextTokenString = properties.ContextToken;

            if (String.IsNullOrEmpty(contextTokenString))
            {
                return null;
            }

            SharePointContextToken contextToken = ReadAndValidateContextToken(contextTokenString, OperationContext.Current.IncomingMessageHeaders.To.Host);
            string accessToken = GetAccessToken(contextToken, sharepointUrl.Authority).AccessToken;

            return GetClientContextWithAccessToken(sharepointUrl.ToString(), accessToken);
        }

        private static string GetAcsMetadataEndpointUrl()
        {
            return Path.Combine(GetAcsGlobalEndpointUrl(), AcsMetadataEndPointRelativeUrl);
        }

        private static string GetFormattedPrincipal(string principalName, string hostName, string realm)
        {
            if (!String.IsNullOrEmpty(hostName))
            {
                return String.Format(CultureInfo.InvariantCulture, "{0}/{1}@{2}", principalName, hostName, realm);
            }

            return String.Format(CultureInfo.InvariantCulture, "{0}@{1}", principalName, realm);
        }

        private static string GetAcsPrincipalName(string realm)
        {
            return GetFormattedPrincipal(AcsPrincipalName, new Uri(GetAcsGlobalEndpointUrl()).Host, realm);
        }

        private static string GetAcsGlobalEndpointUrl()
        {
            return String.Format(CultureInfo.InvariantCulture, "https://{0}.{1}/", GlobalEndPointPrefix, AcsHostUrl);
        }

        private static JwtSecurityTokenHandler CreateJwtSecurityTokenHandler()
        {
            return new JwtSecurityTokenHandler();
        }

        private static string GetS2SAccessTokenWithClaims(
            string targetApplicationHostName,
            string targetRealm,
            IEnumerable<Claim> claims)
        {
            return IssueToken(
                ClientId,
                IssuerId,
                targetRealm,
                SharePointPrincipal,
                targetRealm,
                targetApplicationHostName,
                true,
                claims,
                claims == null);
        }

        private static Claim[] GetClaimsWithWindowsIdentity(WindowsIdentity identity)
        {
            Claim[] claims = new Claim[]
            {
                new Claim(NameIdentifierClaimType, identity.User.Value.ToLower()),
                new Claim("nii", "urn:office:idp:activedirectory")
            };
            return claims;
        }

        private static string IssueToken(
            string sourceApplication,
            string issuerApplication,
            string sourceRealm,
            string targetApplication,
            string targetRealm,
            string targetApplicationHostName,
            bool trustedForDelegation,
            IEnumerable<Claim> claims,
            bool appOnly = false)
        {
            if (null == SigningCredentials)
            {
                throw new InvalidOperationException("SigningCredentials was not initialized");
            }

            #region 参与者标记

            string issuer = string.IsNullOrEmpty(sourceRealm) ? issuerApplication : string.Format("{0}@{1}", issuerApplication, sourceRealm);
            string nameid = string.IsNullOrEmpty(sourceRealm) ? sourceApplication : string.Format("{0}@{1}", sourceApplication, sourceRealm);
            string audience = string.Format("{0}/{1}@{2}", targetApplication, targetApplicationHostName, targetRealm);

            List<Claim> actorClaims = new List<Claim>();
            actorClaims.Add(new Claim(NameIdentifierClaimType, nameid));
            if (trustedForDelegation && !appOnly)
            {
                actorClaims.Add(new Claim(TrustedForImpersonationClaimType, "true"));
            }

            // 创建标记
            JwtSecurityToken actorToken = new JwtSecurityToken(
                issuer: issuer,
                audience: audience,
                claims: actorClaims,
                notBefore: DateTime.UtcNow,
                expires: DateTime.UtcNow.Add(HighTrustAccessTokenLifetime),
                signingCredentials: SigningCredentials
                );

            string actorTokenString = new JwtSecurityTokenHandler().WriteToken(actorToken);

            if (appOnly)
            {
                // 在委托情况下，只允许应用程序使用的标记与参与者标记相同
                return actorTokenString;
            }

            #endregion Actor token

            #region 外部标记

            List<Claim> outerClaims = null == claims ? new List<Claim>() : new List<Claim>(claims);
            outerClaims.Add(new Claim(ActorTokenClaimType, actorTokenString));

            JwtSecurityToken jsonToken = new JwtSecurityToken(
                nameid, // 外部标记颁发者应与参与者标记的 nameid 匹配
                audience,
                outerClaims,
                DateTime.UtcNow,
                DateTime.UtcNow.Add(HighTrustAccessTokenLifetime)
                );

            string accessToken = new JwtSecurityTokenHandler().WriteToken(jsonToken);

            #endregion Outer token

            return accessToken;
        }

        #endregion

        #region AcsMetadataParser

        // 该类用于从全局 STS 终结点获取元数据文档。 它包含
        // 分析元数据文档以及获取终结点和 STS 证书的方法。
        public static class AcsMetadataParser
        {
            public static X509Certificate2 GetAcsSigningCert(string realm)
            {
                JsonMetadataDocument document = GetMetadataDocument(realm);

                if (null != document.keys && document.keys.Count > 0)
                {
                    JsonKey signingKey = document.keys[0];

                    if (null != signingKey && null != signingKey.keyValue)
                    {
                        return new X509Certificate2(Encoding.UTF8.GetBytes(signingKey.keyValue.value));
                    }
                }

                throw new Exception("Metadata document does not contain ACS signing certificate.");
            }

            public static string GetDelegationServiceUrl(string realm)
            {
                JsonMetadataDocument document = GetMetadataDocument(realm);

                JsonEndpoint delegationEndpoint = document.endpoints.SingleOrDefault(e => e.protocol == DelegationIssuance);

                if (null != delegationEndpoint)
                {
                    return delegationEndpoint.location;
                }
                throw new Exception("Metadata document does not contain Delegation Service endpoint Url");
            }

            private static JsonMetadataDocument GetMetadataDocument(string realm)
            {
                string acsMetadataEndpointUrlWithRealm = String.Format(CultureInfo.InvariantCulture, "{0}?realm={1}",
                                                                       GetAcsMetadataEndpointUrl(),
                                                                       realm);
                byte[] acsMetadata;
                using (WebClient webClient = new WebClient())
                {

                    acsMetadata = webClient.DownloadData(acsMetadataEndpointUrlWithRealm);
                }
                string jsonResponseString = Encoding.UTF8.GetString(acsMetadata);

                JavaScriptSerializer serializer = new JavaScriptSerializer();
                JsonMetadataDocument document = serializer.Deserialize<JsonMetadataDocument>(jsonResponseString);

                if (null == document)
                {
                    throw new Exception("No metadata document found at the global endpoint " + acsMetadataEndpointUrlWithRealm);
                }

                return document;
            }

            public static string GetStsUrl(string realm)
            {
                JsonMetadataDocument document = GetMetadataDocument(realm);

                JsonEndpoint s2sEndpoint = document.endpoints.SingleOrDefault(e => e.protocol == S2SProtocol);

                if (null != s2sEndpoint)
                {
                    return s2sEndpoint.location;
                }

                throw new Exception("Metadata document does not contain STS endpoint url");
            }

            private class JsonMetadataDocument
            {
                public string serviceName { get; set; }
                public List<JsonEndpoint> endpoints { get; set; }
                public List<JsonKey> keys { get; set; }
            }

            private class JsonEndpoint
            {
                public string location { get; set; }
                public string protocol { get; set; }
                public string usage { get; set; }
            }

            private class JsonKeyValue
            {
                public string type { get; set; }
                public string value { get; set; }
            }

            private class JsonKey
            {
                public string usage { get; set; }
                public JsonKeyValue keyValue { get; set; }
            }
        }

        #endregion
    }

    /// <summary>
    /// 由 SharePoint 生成的 JwtSecurityToken，它可对第三方应用程序进行身份验证，并允许使用刷新标记进行回拨
    /// </summary>
    public class SharePointContextToken : JwtSecurityToken
    {
        public static SharePointContextToken Create(JwtSecurityToken contextToken)
        {
            return new SharePointContextToken(contextToken.Issuer, contextToken.Audiences.FirstOrDefault<string>(), contextToken.ValidFrom, contextToken.ValidTo, contextToken.Claims);
        }

        public SharePointContextToken(string issuer, string audience, DateTime validFrom, DateTime validTo, IEnumerable<Claim> claims)
        : base(issuer, audience, claims, validFrom, validTo)
        {
        }

        public SharePointContextToken(string issuer, string audience, DateTime validFrom, DateTime validTo, IEnumerable<Claim> claims, SecurityToken issuerToken, JwtSecurityToken actorToken)
            : base(issuer, audience, claims, validFrom, validTo, actorToken.SigningCredentials)
        {
            //此方法用于与 TokenHelper 之前版本的后向兼容。
            //当前版本的 JwtSecurityToken 没有采用以上所有参数的构造函数。
        }

        public SharePointContextToken(string issuer, string audience, DateTime validFrom, DateTime validTo, IEnumerable<Claim> claims, SigningCredentials signingCredentials)
            : base(issuer, audience, claims, validFrom, validTo, signingCredentials)
        {
        }

        public string NameId
        {
            get
            {
                return GetClaimValue(this, "nameid");
            }
        }

        /// <summary>
        /// 上下文标记 "appctxsender" 声明的主体名称部分
        /// </summary>
        public string TargetPrincipalName
        {
            get
            {
                string appctxsender = GetClaimValue(this, "appctxsender");

                if (appctxsender == null)
                {
                    return null;
                }

                return appctxsender.Split('@')[0];
            }
        }

        /// <summary>
        /// 上下文标记的 "refreshtoke" 声明
        /// </summary>
        public string RefreshToken
        {
            get
            {
                return GetClaimValue(this, "refreshtoken");
            }
        }

        /// <summary>
        /// 上下文标记的 "CacheKey" 声明
        /// </summary>
        public string CacheKey
        {
            get
            {
                string appctx = GetClaimValue(this, "appctx");
                if (appctx == null)
                {
                    return null;
                }

                ClientContext ctx = new ClientContext("http://tempuri.org");
                Dictionary<string, object> dict = (Dictionary<string, object>)ctx.ParseObjectFromJsonString(appctx);
                string cacheKey = (string)dict["CacheKey"];

                return cacheKey;
            }
        }

        /// <summary>
        /// 上下文标记的 "SecurityTokenServiceUri" 声明
        /// </summary>
        public string SecurityTokenServiceUri
        {
            get
            {
                string appctx = GetClaimValue(this, "appctx");
                if (appctx == null)
                {
                    return null;
                }

                ClientContext ctx = new ClientContext("http://tempuri.org");
                Dictionary<string, object> dict = (Dictionary<string, object>)ctx.ParseObjectFromJsonString(appctx);
                string securityTokenServiceUri = (string)dict["SecurityTokenServiceUri"];

                return securityTokenServiceUri;
            }
        }

        /// <summary>
        /// 上下文标记 "audience" 声明的领域部分
        /// </summary>
        public string Realm
        {
            get
            {
                //获取第一个非 null 领域
                foreach (string aud in Audiences)
                {
                    string tokenRealm = aud.Substring(aud.IndexOf('@') + 1);

                    if (string.IsNullOrEmpty(tokenRealm))
                    {
                        continue;
                    }
                    else
                    {
                        return tokenRealm;
                    }
                }

                return null;
            }
        }

        private static string GetClaimValue(JwtSecurityToken token, string claimType)
        {
            if (token == null)
            {
                throw new ArgumentNullException("token");
            }

            foreach (Claim claim in token.Claims)
            {
                if (StringComparer.Ordinal.Equals(claim.Type, claimType))
                {
                    return claim.Value;
                }
            }

            return null;
        }

    }

    /// <summary>
    /// 表示含有多个使用对称算法生成的安全密钥的安全标记。
    /// </summary>
    public class MultipleSymmetricKeySecurityToken : SecurityToken
    {
        /// <summary>
        /// 对 MultipleSymmetricKeySecurityToken 类的新实例进行初始化。
        /// </summary>
        /// <param name="keys">包含对称密钥的字节数组枚举。</param>
        public MultipleSymmetricKeySecurityToken(IEnumerable<byte[]> keys)
            : this(Microsoft.IdentityModel.Tokens.UniqueId.CreateUniqueId(), keys)
        {
        }

        /// <summary>
        /// 对 MultipleSymmetricKeySecurityToken 类的新实例进行初始化。
        /// </summary>
        /// <param name="tokenId">安全标记的唯一标识符。</param>
        /// <param name="keys">包含对称密钥的字节数组枚举。</param>
        public MultipleSymmetricKeySecurityToken(string tokenId, IEnumerable<byte[]> keys)
        {
            if (keys == null)
            {
                throw new ArgumentNullException("keys");
            }

            if (String.IsNullOrEmpty(tokenId))
            {
                throw new ArgumentException("Value cannot be a null or empty string.", "tokenId");
            }

            foreach (byte[] key in keys)
            {
                if (key.Length <= 0)
                {
                    throw new ArgumentException("The key length must be greater then zero.", "keys");
                }
            }

            id = tokenId;
            effectiveTime = DateTime.UtcNow;
            securityKeys = CreateSymmetricSecurityKeys(keys);
        }

        /// <summary>
        /// 获取安全标记的唯一标识符。
        /// </summary>
        public override string Id
        {
            get
            {
                return id;
            }
        }

        /// <summary>
        /// 获取与安全标记关联的加密密钥。
        /// </summary>
        public override ReadOnlyCollection<SecurityKey> SecurityKeys
        {
            get
            {
                return securityKeys.AsReadOnly();
            }
        }

        /// <summary>
        /// 在安全标记有效后及时获取第一个瞬间。
        /// </summary>
        public override DateTime ValidFrom
        {
            get
            {
                return effectiveTime;
            }
        }

        /// <summary>
        /// 在安全标记有效后及时获取最后一个瞬间。
        /// </summary>
        public override DateTime ValidTo
        {
            get
            {
                // 永不过期
                return DateTime.MaxValue;
            }
        }

        /// <summary>
        /// 返回一个值，用于指示此实例的密钥标识符是否可以分析到指定密钥标识符。
        /// </summary>
        /// <param name="keyIdentifierClause">要与此实例进行比较的 SecurityKeyIdentifierClause</param>
        /// <returns>如果 keyIdentifierClause 为 SecurityKeyIdentifierClause，并且具有与 ID 属性相同的唯一标识符，则为 true；否则为 false。</returns>
        public override bool MatchesKeyIdentifierClause(SecurityKeyIdentifierClause keyIdentifierClause)
        {
            if (keyIdentifierClause == null)
            {
                throw new ArgumentNullException("keyIdentifierClause");
            }

            return base.MatchesKeyIdentifierClause(keyIdentifierClause);
        }

        #region 私有成员

        private List<SecurityKey> CreateSymmetricSecurityKeys(IEnumerable<byte[]> keys)
        {
            List<SecurityKey> symmetricKeys = new List<SecurityKey>();
            foreach (byte[] key in keys)
            {
                symmetricKeys.Add(new InMemorySymmetricSecurityKey(key));
            }
            return symmetricKeys;
        }

        private string id;
        private DateTime effectiveTime;
        private List<SecurityKey> securityKeys;

        #endregion
    }

    /// <summary>
    ///表示 ACS 服务器调用中的 OAuth 响应。
    /// </summary>
    public class OAuthTokenResponse
    {
        /// <summary>
        ///默认构造函数。
        /// </summary>
        public OAuthTokenResponse() { }

        /// <summary>
        ///在从 ACS 服务器返回的字节数组中构造 OAuthTokenResponse 对象。
        /// </summary>
        /// <param name="response">从 ACS 获得的原始字节数组。</param>
        public OAuthTokenResponse(byte[] response)
        {
            var serializer = new JavaScriptSerializer();
            this.Data = serializer.DeserializeObject(Encoding.UTF8.GetString(response)) as Dictionary<string, object>;

            this.AccessToken = this.GetValue("access_token");
            this.TokenType = this.GetValue("token_type");
            this.Resource = this.GetValue("resource");
            this.UserType = this.GetValue("user_type");

            long epochTime = 0;
            if (long.TryParse(this.GetValue("expires_in"), out epochTime))
            {
                this.ExpiresIn = epochTime;
            }
            if (long.TryParse(this.GetValue("expires_on"), out epochTime))
            {
                this.ExpiresOn = epochTime;
            }
            if (long.TryParse(this.GetValue("not_before"), out epochTime))
            {
                this.NotBefore = epochTime;
            }
            if (long.TryParse(this.GetValue("extended_expires_in"), out epochTime))
            {
                this.ExtendedExpiresIn = epochTime;
            }
        }

        /// <summary>
        ///获取访问令牌。
        /// </summary>
        public string AccessToken { get; private set; }

        /// <summary>
        ///从原始响应获取分析的数据。
        /// </summary>
        public IDictionary<string, object> Data { get; }

        /// <summary>
        ///获取以 Epoch 时间表示的到期时间。
        /// </summary>
        public long ExpiresIn { get; }

        /// <summary>
        ///获取时期时间中的过期时间。
        /// </summary>
        public long ExpiresOn { get; }

        /// <summary>
        ///获取 extended expires in 时期时间。
        /// </summary>
        public long ExtendedExpiresIn { get; }

        /// <summary>
        ///获取时期时间之前的过期时间。
        /// </summary>
        public long NotBefore { get; }

        /// <summary>
        ///获取资源。
        /// </summary>
        public string Resource { get; }

        /// <summary>
        ///获取标记类型。
        /// </summary>
        public string TokenType { get; }

        /// <summary>
        ///获取用户类型。
        /// </summary>
        public string UserType { get; }

        /// <summary>
        ///通过键从数据中获取值。
        /// </summary>
        /// <param name="key">键。</param>
        /// <returns>如果键值存在，则为键值，否则为空字符串。</returns>
        private string GetValue(string key)
        {
            if (this.Data.TryGetValue(key, out object value))
            {
                return value as string;
            }
            else
            {
                return string.Empty;
            }
        }
    }

    /// <summary>
    ///表示 Web 客户端，用于向 ACS 服务器发出 OAuth 调用。
    /// </summary>
    public class OAuthClient
    {
        /// <summary>
        ///使用刷新标记获得 OAuthTokenResponse。
        /// </summary>
        /// <param name="uri">ACS 服务器的 URI。</param>
        /// <param name="clientId">客户端 ID。</param>
        /// <param name="ClientSecret">客户端密钥。</param>
        /// <param name="refreshToken">刷新标记。</param>
        /// <param name="resource">资源。</param>
        /// <returns>从 ACS 服务器响应。</returns>
        public static OAuthTokenResponse GetAccessTokenWithRefreshToken(string uri, string clientId,
            string ClientSecret, string refreshToken, string resource)
        {
            WebClient client = new WebClient();
            NameValueCollection values = new NameValueCollection
            {
                { "grant_type", "refresh_token" },
                { "client_id", clientId },
                { "client_secret", ClientSecret },
                { "refresh_token", refreshToken },
                { "resource", resource }
            };

            byte[] response = client.UploadValues(uri, "POST", values);

            return new OAuthTokenResponse(response);
        }

        /// <summary>
        ///使用客户端凭据获得 OAuthTokenResponse。
        /// </summary>
        /// <param name="uri">ACS 服务器的 URI。</param>
        /// <param name="clientId">客户端 ID。</param>
        /// <param name="ClientSecret">客户端密钥。</param>
        /// <param name="resource">资源。</param>
        /// <returns>从 ACS 服务器响应。</returns>
        public static OAuthTokenResponse GetAccessTokenWithClientCredentials(string uri, string clientId,
            string ClientSecret, string resource)
        {
            WebClient client = new WebClient();
            NameValueCollection values = new NameValueCollection
            {
                { "grant_type", "client_credentials" },
                { "client_id", clientId },
                { "client_secret", ClientSecret },
                { "resource", resource }
            };

            byte[] response = client.UploadValues(uri, "POST", values);

            return new OAuthTokenResponse(response);
        }

        /// <summary>
        ///使用授权代码获得 OAuthTokenResponse。
        /// </summary>
        /// <param name="uri">ACS 服务器的 URI。</param>
        /// <param name="clientId">客户端 ID。</param>
        /// <param name="ClientSecret">客户端密钥。</param>
        /// <param name="authorizationCode">授权代码。</param>
        /// <param name="redirectUri">重定向 Uri。</param>
        /// <param name="resource">资源。</param>
        /// <returns>从 ACS 服务器响应。</returns>
        public static OAuthTokenResponse GetAccessTokenWithAuthorizationCode(string uri, string clientId,
            string ClientSecret, string authorizationCode, string redirectUri, string resource)
        {
            WebClient client = new WebClient();
            NameValueCollection values = new NameValueCollection
            {
                { "grant_type", "authorization_code" },
                { "client_id", clientId },
                { "client_secret", ClientSecret },
                { "code", authorizationCode },
                { "redirect_uri", redirectUri },
                { "resource", resource }
            };

            byte[] response = client.UploadValues(uri, "POST", values);

            return new OAuthTokenResponse(response);
        }
    }
}
