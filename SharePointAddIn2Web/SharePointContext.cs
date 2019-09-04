using Microsoft.SharePoint.Client;
using System;
using System.IdentityModel.Tokens;
using System.Net;
using System.Security.Principal;
using System.Web;

namespace SharePointAddIn2Web
{
    /// <summary>
    /// 封装 SharePoint 中的所有信息。
    /// </summary>
    public abstract class SharePointContext
    {
        public const string SPHostUrlKey = "SPHostUrl";
        public const string SPAppWebUrlKey = "SPAppWebUrl";
        public const string SPLanguageKey = "SPLanguage";
        public const string SPClientTagKey = "SPClientTag";
        public const string SPProductNumberKey = "SPProductNumber";

        protected static readonly long AccessTokenLifetimeTolerance = 5 * 60; //5 分钟

        private readonly Uri spHostUrl;
        private readonly Uri spAppWebUrl;
        private readonly string spLanguage;
        private readonly string spClientTag;
        private readonly string spProductNumber;

        // <AccessTokenString，以 Epoch 时间表示的 UtcExpiresOn>
        protected Tuple<string, long> userAccessTokenForSPHost;
        protected Tuple<string, long> userAccessTokenForSPAppWeb;
        protected Tuple<string, long> appOnlyAccessTokenForSPHost;
        protected Tuple<string, long> appOnlyAccessTokenForSPAppWeb;

        /// <summary>
        /// 从指定的 HTTP 请求的 QueryString 中获取 SharePoint 宿主 URL。
        /// </summary>
        /// <param name="httpRequest">指定的 HTTP 请求。</param>
        /// <returns>SharePoint 宿主 URL。如果 HTTP 请求不包含 SharePoint 宿主 URL，则返回 <c>null</c>。</returns>
        public static Uri GetSPHostUrl(HttpRequestBase httpRequest)
        {
            if (httpRequest == null)
            {
                throw new ArgumentNullException("httpRequest");
            }

            string spHostUrlString = TokenHelper.EnsureTrailingSlash(httpRequest.QueryString[SPHostUrlKey]);

            if (Uri.TryCreate(spHostUrlString, UriKind.Absolute, out Uri spHostUrl) &&
                (spHostUrl.Scheme == Uri.UriSchemeHttp || spHostUrl.Scheme == Uri.UriSchemeHttps))
            {
                return spHostUrl;
            }

            return null;
        }

        /// <summary>
        /// 从指定的 HTTP 请求的 QueryString 中获取 SharePoint 宿主 URL。
        /// </summary>
        /// <param name="httpRequest">指定的 HTTP 请求。</param>
        /// <returns>SharePoint 宿主 URL。如果 HTTP 请求不包含 SharePoint 宿主 URL，则返回 <c>null</c>。</returns>
        public static Uri GetSPHostUrl(HttpRequest httpRequest)
        {
            return GetSPHostUrl(new HttpRequestWrapper(httpRequest));
        }

        /// <summary>
        /// SharePoint 宿主 URL。
        /// </summary>
        public Uri SPHostUrl
        {
            get { return this.spHostUrl; }
        }

        /// <summary>
        /// SharePoint 应用程序网站 URL。
        /// </summary>
        public Uri SPAppWebUrl
        {
            get { return this.spAppWebUrl; }
        }

        /// <summary>
        /// SharePoint 语言。
        /// </summary>
        public string SPLanguage
        {
            get { return this.spLanguage; }
        }

        /// <summary>
        /// SharePoint 客户端标记。
        /// </summary>
        public string SPClientTag
        {
            get { return this.spClientTag; }
        }

        /// <summary>
        /// SharePoint 产品编号。
        /// </summary>
        public string SPProductNumber
        {
            get { return this.spProductNumber; }
        }

        /// <summary>
        /// 适用于 SharePoint 宿主的用户访问令牌。
        /// </summary>
        public abstract string UserAccessTokenForSPHost
        {
            get;
        }

        /// <summary>
        /// 适用于 SharePoint 应用程序网站的用户访问令牌。
        /// </summary>
        public abstract string UserAccessTokenForSPAppWeb
        {
            get;
        }

        /// <summary>
        /// 适用于 SharePoint Web 宿主的应用程序专用访问令牌。
        /// </summary>
        public abstract string AppOnlyAccessTokenForSPHost
        {
            get;
        }

        /// <summary>
        /// 适用于 SharePoint 应用程序网站的应用程序专用访问令牌。
        /// </summary>
        public abstract string AppOnlyAccessTokenForSPAppWeb
        {
            get;
        }

        /// <summary>
        /// 构造函数。
        /// </summary>
        /// <param name="spHostUrl">SharePoint 宿主 URL。</param>
        /// <param name="spAppWebUrl">SharePoint 应用程序网站 URL。</param>
        /// <param name="spLanguage">SharePoint 语言。</param>
        /// <param name="spClientTag">SharePoint 客户端标记。</param>
        /// <param name="spProductNumber">SharePoint 产品编号。</param>
        protected SharePointContext(Uri spHostUrl, Uri spAppWebUrl, string spLanguage, string spClientTag, string spProductNumber)
        {
            if (spHostUrl == null)
            {
                throw new ArgumentNullException("spHostUrl");
            }

            if (string.IsNullOrEmpty(spLanguage))
            {
                throw new ArgumentNullException("spLanguage");
            }

            if (string.IsNullOrEmpty(spClientTag))
            {
                throw new ArgumentNullException("spClientTag");
            }

            if (string.IsNullOrEmpty(spProductNumber))
            {
                throw new ArgumentNullException("spProductNumber");
            }

            this.spHostUrl = spHostUrl;
            this.spAppWebUrl = spAppWebUrl;
            this.spLanguage = spLanguage;
            this.spClientTag = spClientTag;
            this.spProductNumber = spProductNumber;
        }

        /// <summary>
        /// 为 SharePoint 宿主创建一个用户 ClientContext。
        /// </summary>
        /// <returns>ClientContext 实例。</returns>
        public ClientContext CreateUserClientContextForSPHost()
        {
            return CreateClientContext(this.SPHostUrl, this.UserAccessTokenForSPHost);
        }

        /// <summary>
        /// 创建适用于 SharePoint 应用程序网站的用户 ClientContext。
        /// </summary>
        /// <returns>ClientContext 实例。</returns>
        public ClientContext CreateUserClientContextForSPAppWeb()
        {
            return CreateClientContext(this.SPAppWebUrl, this.UserAccessTokenForSPAppWeb);
        }

        /// <summary>
        /// 创建适用于 SharePoint 宿主的应用程序专用 ClientContext。
        /// </summary>
        /// <returns>ClientContext 实例。</returns>
        public ClientContext CreateAppOnlyClientContextForSPHost()
        {
            return CreateClientContext(this.SPHostUrl, this.AppOnlyAccessTokenForSPHost);
        }

        /// <summary>
        /// 创建适用于 SharePoint 应用程序网站的应用程序专用 ClientContext。
        /// </summary>
        /// <returns>ClientContext 实例。</returns>
        public ClientContext CreateAppOnlyClientContextForSPAppWeb()
        {
            return CreateClientContext(this.SPAppWebUrl, this.AppOnlyAccessTokenForSPAppWeb);
        }

        /// <summary>
        /// 从自动托管外接程序的 SharePoint 获取数据库连接字符串。
        ///由于不再提供自动托管选项，已弃用此方法。
        /// </summary>
        [ObsoleteAttribute("This method is deprecated because the autohosted option is no longer available.", true)]
        public string GetDatabaseConnectionString()
        {
            throw new NotSupportedException("This method is deprecated because the autohosted option is no longer available.");
        }

        /// <summary>
        /// 确定指定的访问令牌是否有效。
        /// 如果访问令牌为 null 或者已过期，则将其视为无效。
        /// </summary>
        /// <param name="accessToken">要验证的访问令牌。</param>
        /// <returns>如果访问令牌有效，则为 True。</returns>
        protected static bool IsAccessTokenValid(Tuple<string, long> accessToken)
        {
            return accessToken != null &&
                   !string.IsNullOrEmpty(accessToken.Item1) &&
                   accessToken.Item2 > TokenHelper.EpochTimeNow();
        }

        /// <summary>
        /// 使用指定的 SharePoint 网站 URL 和访问令牌创建 ClientContext。
        /// </summary>
        /// <param name="spSiteUrl">站点 URL。</param>
        /// <param name="accessToken">访问令牌。</param>
        /// <returns>ClientContext 实例。</returns>
        private static ClientContext CreateClientContext(Uri spSiteUrl, string accessToken)
        {
            if (spSiteUrl != null && !string.IsNullOrEmpty(accessToken))
            {
                return TokenHelper.GetClientContextWithAccessToken(spSiteUrl.AbsoluteUri, accessToken);
            }

            return null;
        }
    }

    /// <summary>
    /// 重定向状态。
    /// </summary>
    public enum RedirectionStatus
    {
        Ok,
        ShouldRedirect,
        CanNotRedirect
    }

    /// <summary>
    /// 提供 SharePointContext 实例。
    /// </summary>
    public abstract class SharePointContextProvider
    {
        private static SharePointContextProvider current;

        /// <summary>
        /// 当前的 SharePointContextProvider 实例。
        /// </summary>
        public static SharePointContextProvider Current
        {
            get { return SharePointContextProvider.current; }
        }

        /// <summary>
        /// 初始化默认的 SharePointContextProvider 实例。
        /// </summary>
        static SharePointContextProvider()
        {
            if (!TokenHelper.IsHighTrustApp())
            {
                SharePointContextProvider.current = new SharePointAcsContextProvider();
            }
            else
            {
                SharePointContextProvider.current = new SharePointHighTrustContextProvider();
            }
        }

        /// <summary>
        /// 将指定的 SharePointContextProvider 实例注册为当前项。
        /// 它应由 Global.asax 中的 Application_Start() 调用。
        /// </summary>
        /// <param name="provider">将 SharePointContextProvider 设置为当前项。</param>
        public static void Register(SharePointContextProvider provider)
        {
            if (provider == null)
            {
                throw new ArgumentNullException("provider");
            }

            SharePointContextProvider.current = provider;
        }

        /// <summary>
        /// 检查是否必须重定向到 SharePoint 供用户进行身份验证。
        /// </summary>
        /// <param name="httpContext">HTTP 上下文。</param>
        /// <param name="redirectUrl">如果状态为“ShouldRedirect”，则返回指向 SharePoint 的重定向 URL。如果状态为“Ok”或“CanNotRedirect”，则返回 <c>Null</c>。</param>
        /// <returns>重定向状态。</returns>
        public static RedirectionStatus CheckRedirectionStatus(HttpContextBase httpContext, out Uri redirectUrl)
        {
            if (httpContext == null)
            {
                throw new ArgumentNullException("httpContext");
            }

            redirectUrl = null;
            bool contextTokenExpired = false;

            try
            {
                if (SharePointContextProvider.Current.GetSharePointContext(httpContext) != null)
                {
                    return RedirectionStatus.Ok;
                }
            }
            catch (SecurityTokenExpiredException)
            {
                contextTokenExpired = true;
            }

            const string SPHasRedirectedToSharePointKey = "SPHasRedirectedToSharePoint";

            if (!string.IsNullOrEmpty(httpContext.Request.QueryString[SPHasRedirectedToSharePointKey]) && !contextTokenExpired)
            {
                return RedirectionStatus.CanNotRedirect;
            }

            Uri spHostUrl = SharePointContext.GetSPHostUrl(httpContext.Request);

            if (spHostUrl == null)
            {
                return RedirectionStatus.CanNotRedirect;
            }

            if (StringComparer.OrdinalIgnoreCase.Equals(httpContext.Request.HttpMethod, "POST"))
            {
                return RedirectionStatus.CanNotRedirect;
            }

            Uri requestUrl = httpContext.Request.Url;

            var queryNameValueCollection = HttpUtility.ParseQueryString(requestUrl.Query);

            // 移除 {StandardTokens} 中包括的值，因为 {StandardTokens} 将插入在查询字符串的开头。
            queryNameValueCollection.Remove(SharePointContext.SPHostUrlKey);
            queryNameValueCollection.Remove(SharePointContext.SPAppWebUrlKey);
            queryNameValueCollection.Remove(SharePointContext.SPLanguageKey);
            queryNameValueCollection.Remove(SharePointContext.SPClientTagKey);
            queryNameValueCollection.Remove(SharePointContext.SPProductNumberKey);

            // 添加 SPHasRedirectedToSharePoint=1。
            queryNameValueCollection.Add(SPHasRedirectedToSharePointKey, "1");

            UriBuilder returnUrlBuilder = new UriBuilder(requestUrl);
            returnUrlBuilder.Query = queryNameValueCollection.ToString();

            // 插入 StandardTokens。
            const string StandardTokens = "{StandardTokens}";
            string returnUrlString = returnUrlBuilder.Uri.AbsoluteUri;
            returnUrlString = returnUrlString.Insert(returnUrlString.IndexOf("?") + 1, StandardTokens + "&");

            // 构造重定向 URL。
            string redirectUrlString = TokenHelper.GetAppContextTokenRequestUrl(spHostUrl.AbsoluteUri, Uri.EscapeDataString(returnUrlString));

            redirectUrl = new Uri(redirectUrlString, UriKind.Absolute);

            return RedirectionStatus.ShouldRedirect;
        }

        /// <summary>
        /// 检查是否必须重定向到 SharePoint 供用户进行身份验证。
        /// </summary>
        /// <param name="httpContext">HTTP 上下文。</param>
        /// <param name="redirectUrl">如果状态为“ShouldRedirect”，则返回指向 SharePoint 的重定向 URL。如果状态为“Ok”或“CanNotRedirect”，则返回 <c>Null</c>。</param>
        /// <returns>重定向状态。</returns>
        public static RedirectionStatus CheckRedirectionStatus(HttpContext httpContext, out Uri redirectUrl)
        {
            return CheckRedirectionStatus(new HttpContextWrapper(httpContext), out redirectUrl);
        }

        /// <summary>
        /// 使用指定的 HTTP 请求创建 SharePointContext 实例。
        /// </summary>
        /// <param name="httpRequest">HTTP 请求。</param>
        /// <returns>SharePointContext 实例。如果出现错误，则返回 <c>null</c>。</returns>
        public SharePointContext CreateSharePointContext(HttpRequestBase httpRequest)
        {
            if (httpRequest == null)
            {
                throw new ArgumentNullException("httpRequest");
            }

            // SPHostUrl
            Uri spHostUrl = SharePointContext.GetSPHostUrl(httpRequest);
            if (spHostUrl == null)
            {
                return null;
            }

            // SPAppWebUrl
            string spAppWebUrlString = TokenHelper.EnsureTrailingSlash(httpRequest.QueryString[SharePointContext.SPAppWebUrlKey]);
            Uri spAppWebUrl;
            if (!Uri.TryCreate(spAppWebUrlString, UriKind.Absolute, out spAppWebUrl) ||
                !(spAppWebUrl.Scheme == Uri.UriSchemeHttp || spAppWebUrl.Scheme == Uri.UriSchemeHttps))
            {
                spAppWebUrl = null;
            }

            // SPLanguage
            string spLanguage = httpRequest.QueryString[SharePointContext.SPLanguageKey];
            if (string.IsNullOrEmpty(spLanguage))
            {
                return null;
            }

            // SPClientTag
            string spClientTag = httpRequest.QueryString[SharePointContext.SPClientTagKey];
            if (string.IsNullOrEmpty(spClientTag))
            {
                return null;
            }

            // SPProductNumber
            string spProductNumber = httpRequest.QueryString[SharePointContext.SPProductNumberKey];
            if (string.IsNullOrEmpty(spProductNumber))
            {
                return null;
            }

            return CreateSharePointContext(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber, httpRequest);
        }

        /// <summary>
        /// 使用指定的 HTTP 请求创建 SharePointContext 实例。
        /// </summary>
        /// <param name="httpRequest">HTTP 请求。</param>
        /// <returns>SharePointContext 实例。如果出现错误，则返回 <c>null</c>。</returns>
        public SharePointContext CreateSharePointContext(HttpRequest httpRequest)
        {
            return CreateSharePointContext(new HttpRequestWrapper(httpRequest));
        }

        /// <summary>
        /// 获取与指定的 HTTP 上下文关联的 SharePointContext 实例。
        /// </summary>
        /// <param name="httpContext">HTTP 上下文。</param>
        /// <returns>SharePointContext 实例。如果未找到该实例并且无法创建新实例，则返回 <c>null</c>。</returns>
        public SharePointContext GetSharePointContext(HttpContextBase httpContext)
        {
            if (httpContext == null)
            {
                throw new ArgumentNullException("httpContext");
            }

            Uri spHostUrl = SharePointContext.GetSPHostUrl(httpContext.Request);
            if (spHostUrl == null)
            {
                return null;
            }

            SharePointContext spContext = LoadSharePointContext(httpContext);

            if (spContext == null || !ValidateSharePointContext(spContext, httpContext))
            {
                spContext = CreateSharePointContext(httpContext.Request);

                if (spContext != null)
                {
                    SaveSharePointContext(spContext, httpContext);
                }
            }

            return spContext;
        }

        /// <summary>
        /// 获取与指定的 HTTP 上下文关联的 SharePointContext 实例。
        /// </summary>
        /// <param name="httpContext">HTTP 上下文。</param>
        /// <returns>SharePointContext 实例。如果未找到该实例并且无法创建新实例，则返回 <c>null</c>。</returns>
        public SharePointContext GetSharePointContext(HttpContext httpContext)
        {
            return GetSharePointContext(new HttpContextWrapper(httpContext));
        }

        /// <summary>
        /// 创建 SharePointContext 实例。
        /// </summary>
        /// <param name="spHostUrl">SharePoint 宿主 URL。</param>
        /// <param name="spAppWebUrl">SharePoint 应用程序网站 URL。</param>
        /// <param name="spLanguage">SharePoint 语言。</param>
        /// <param name="spClientTag">SharePoint 客户端标记。</param>
        /// <param name="spProductNumber">SharePoint 产品编号。</param>
        /// <param name="httpRequest">HTTP 请求。</param>
        /// <returns>SharePointContext 实例。如果出现错误，则返回 <c>null</c>。</returns>
        protected abstract SharePointContext CreateSharePointContext(Uri spHostUrl, Uri spAppWebUrl, string spLanguage, string spClientTag, string spProductNumber, HttpRequestBase httpRequest);

        /// <summary>
        /// 验证给定的 SharePointContext 是否能与指定的 HTTP 上下文一起使用。
        /// </summary>
        /// <param name="spContext">SharePointContext。</param>
        /// <param name="httpContext">HTTP 上下文。</param>
        /// <returns>如果给定的 SharePointContext 能与指定的 HTTP 上下文一起使用，则为 True。</returns>
        protected abstract bool ValidateSharePointContext(SharePointContext spContext, HttpContextBase httpContext);

        /// <summary>
        /// 加载与指定的 HTTP 上下文关联的 SharePointContext 实例。
        /// </summary>
        /// <param name="httpContext">HTTP 上下文。</param>
        /// <returns>SharePointContext 实例。如果未找到，则返回 <c>null</c>。</returns>
        protected abstract SharePointContext LoadSharePointContext(HttpContextBase httpContext);

        /// <summary>
        /// 保存与指定的 HTTP 上下文关联的指定的 SharePointContext 实例。
        /// 接受 <c>null</c> 可用于清除与 HTTP 上下文关联的 SharePointContext 实例。
        /// </summary>
        /// <param name="spContext">要保存的 SharePointContext 实例，或 <c>null</c>。</param>
        /// <param name="httpContext">HTTP 上下文。</param>
        protected abstract void SaveSharePointContext(SharePointContext spContext, HttpContextBase httpContext);
    }

    #region ACS

    /// <summary>
    /// 采用 ACS 模式封装 SharePoint 中的所有信息。
    /// </summary>
    public class SharePointAcsContext : SharePointContext
    {
        private readonly string contextToken;
        private readonly SharePointContextToken contextTokenObj;

        /// <summary>
        /// 上下文标记。
        /// </summary>
        public string ContextToken
        {
            get { return this.contextTokenObj.ValidTo > DateTime.UtcNow ? this.contextToken : null; }
        }

        /// <summary>
        /// 上下文标记的“CacheKey”声明。
        /// </summary>
        public string CacheKey
        {
            get { return this.contextTokenObj.ValidTo > DateTime.UtcNow ? this.contextTokenObj.CacheKey : null; }
        }

        /// <summary>
        /// 上下文标记的“refreshtoken”声明。
        /// </summary>
        public string RefreshToken
        {
            get { return this.contextTokenObj.ValidTo > DateTime.UtcNow ? this.contextTokenObj.RefreshToken : null; }
        }

        public override string UserAccessTokenForSPHost
        {
            get
            {
                return GetAccessTokenString(ref this.userAccessTokenForSPHost,
                                            () => TokenHelper.GetAccessToken(this.contextTokenObj, this.SPHostUrl.Authority));
            }
        }

        public override string UserAccessTokenForSPAppWeb
        {
            get
            {
                if (this.SPAppWebUrl == null)
                {
                    return null;
                }

                return GetAccessTokenString(ref this.userAccessTokenForSPAppWeb,
                                            () => TokenHelper.GetAccessToken(this.contextTokenObj, this.SPAppWebUrl.Authority));
            }
        }

        public override string AppOnlyAccessTokenForSPHost
        {
            get
            {
                return GetAccessTokenString(ref this.appOnlyAccessTokenForSPHost,
                                            () => TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, this.SPHostUrl.Authority, TokenHelper.GetRealmFromTargetUrl(this.SPHostUrl)));
            }
        }

        public override string AppOnlyAccessTokenForSPAppWeb
        {
            get
            {
                if (this.SPAppWebUrl == null)
                {
                    return null;
                }

                return GetAccessTokenString(ref this.appOnlyAccessTokenForSPAppWeb,
                                            () => TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, this.SPAppWebUrl.Authority, TokenHelper.GetRealmFromTargetUrl(this.SPAppWebUrl)));
            }
        }

        public SharePointAcsContext(Uri spHostUrl, Uri spAppWebUrl, string spLanguage, string spClientTag, string spProductNumber, string contextToken, SharePointContextToken contextTokenObj)
            : base(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber)
        {
            if (string.IsNullOrEmpty(contextToken))
            {
                throw new ArgumentNullException("contextToken");
            }

            if (contextTokenObj == null)
            {
                throw new ArgumentNullException("contextTokenObj");
            }

            this.contextToken = contextToken;
            this.contextTokenObj = contextTokenObj;
        }

        /// <summary>
        /// 确保访问令牌有效并返回该令牌。
        /// </summary>
        /// <param name="accessToken">要验证的访问令牌。</param>
        /// <param name="tokenRenewalHandler">标记续订处理程序。</param>
        /// <returns>访问令牌字符串。</returns>
        private static string GetAccessTokenString(ref Tuple<string, long> accessToken, Func<OAuthTokenResponse> tokenRenewalHandler)
        {
            RenewAccessTokenIfNeeded(ref accessToken, tokenRenewalHandler);

            return IsAccessTokenValid(accessToken) ? accessToken.Item1 : null;
        }

        /// <summary>
        /// 如果访问令牌无效，则应续订访问令牌。
        /// </summary>
        /// <param name="accessToken">要续订的访问令牌。</param>
        /// <param name="tokenRenewalHandler">标记续订处理程序。</param>
        private static void RenewAccessTokenIfNeeded(ref Tuple<string, long> accessToken, Func<OAuthTokenResponse> tokenRenewalHandler)
        {
            if (IsAccessTokenValid(accessToken))
            {
                return;
            }

            try
            {
                OAuthTokenResponse oauthTokenResponse = tokenRenewalHandler();

                long expiresOn = oauthTokenResponse.ExpiresOn;

                if ((expiresOn - oauthTokenResponse.NotBefore) > AccessTokenLifetimeTolerance)
                {
                    // 在访问令牌到期之前稍微提前一些进行续订
                    // 以便使用它的对 SharePoint 的调用将有足够的时间来成功完成操作。
                    expiresOn -= AccessTokenLifetimeTolerance;
                }

                accessToken = Tuple.Create(oauthTokenResponse.AccessToken, expiresOn);
            }
            catch (WebException)
            {
            }
        }
    }

    /// <summary>
    /// SharePointAcsContext 的默认提供程序。
    /// </summary>
    public class SharePointAcsContextProvider : SharePointContextProvider
    {
        private const string SPContextKey = "SPContext";
        private const string SPCacheKeyKey = "SPCacheKey";

        protected override SharePointContext CreateSharePointContext(Uri spHostUrl, Uri spAppWebUrl, string spLanguage, string spClientTag, string spProductNumber, HttpRequestBase httpRequest)
        {
            string contextTokenString = TokenHelper.GetContextTokenFromRequest(httpRequest);
            if (string.IsNullOrEmpty(contextTokenString))
            {
                return null;
            }

            SharePointContextToken contextToken = null;
            try
            {
                contextToken = TokenHelper.ReadAndValidateContextToken(contextTokenString, httpRequest.Url.Authority);
            }
            catch (WebException)
            {
                return null;
            }
            catch (AudienceUriValidationFailedException)
            {
                return null;
            }

            return new SharePointAcsContext(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber, contextTokenString, contextToken);
        }

        protected override bool ValidateSharePointContext(SharePointContext spContext, HttpContextBase httpContext)
        {
            SharePointAcsContext spAcsContext = spContext as SharePointAcsContext;

            if (spAcsContext != null)
            {
                Uri spHostUrl = SharePointContext.GetSPHostUrl(httpContext.Request);
                string contextToken = TokenHelper.GetContextTokenFromRequest(httpContext.Request);
                HttpCookie spCacheKeyCookie = httpContext.Request.Cookies[SPCacheKeyKey];
                string spCacheKey = spCacheKeyCookie != null ? spCacheKeyCookie.Value : null;

                return spHostUrl == spAcsContext.SPHostUrl &&
                       !string.IsNullOrEmpty(spAcsContext.CacheKey) &&
                       spCacheKey == spAcsContext.CacheKey &&
                       !string.IsNullOrEmpty(spAcsContext.ContextToken) &&
                       (string.IsNullOrEmpty(contextToken) || contextToken == spAcsContext.ContextToken);
            }

            return false;
        }

        protected override SharePointContext LoadSharePointContext(HttpContextBase httpContext)
        {
            return httpContext.Session[SPContextKey] as SharePointAcsContext;
        }

        protected override void SaveSharePointContext(SharePointContext spContext, HttpContextBase httpContext)
        {
            SharePointAcsContext spAcsContext = spContext as SharePointAcsContext;

            if (spAcsContext != null)
            {
                HttpCookie spCacheKeyCookie = new HttpCookie(SPCacheKeyKey)
                {
                    Value = spAcsContext.CacheKey,
                    Secure = true,
                    HttpOnly = true
                };

                httpContext.Response.AppendCookie(spCacheKeyCookie);
            }

            httpContext.Session[SPContextKey] = spAcsContext;
        }
    }

    #endregion ACS

    #region HighTrust

    /// <summary>
    /// 采用 HighTrust 模式封装 SharePoint 中的所有信息。
    /// </summary>
    public class SharePointHighTrustContext : SharePointContext
    {
        private readonly WindowsIdentity logonUserIdentity;

        /// <summary>
        /// 当前用户的 Windows 标识。
        /// </summary>
        public WindowsIdentity LogonUserIdentity
        {
            get { return this.logonUserIdentity; }
        }

        public override string UserAccessTokenForSPHost
        {
            get
            {
                return GetAccessTokenString(ref this.userAccessTokenForSPHost,
                                            () => TokenHelper.GetS2SAccessTokenWithWindowsIdentity(this.SPHostUrl, this.LogonUserIdentity));
            }
        }

        public override string UserAccessTokenForSPAppWeb
        {
            get
            {
                if (this.SPAppWebUrl == null)
                {
                    return null;
                }

                return GetAccessTokenString(ref this.userAccessTokenForSPAppWeb,
                                            () => TokenHelper.GetS2SAccessTokenWithWindowsIdentity(this.SPAppWebUrl, this.LogonUserIdentity));
            }
        }

        public override string AppOnlyAccessTokenForSPHost
        {
            get
            {
                return GetAccessTokenString(ref this.appOnlyAccessTokenForSPHost,
                                            () => TokenHelper.GetS2SAccessTokenWithWindowsIdentity(this.SPHostUrl, null));
            }
        }

        public override string AppOnlyAccessTokenForSPAppWeb
        {
            get
            {
                if (this.SPAppWebUrl == null)
                {
                    return null;
                }

                return GetAccessTokenString(ref this.appOnlyAccessTokenForSPAppWeb,
                                            () => TokenHelper.GetS2SAccessTokenWithWindowsIdentity(this.SPAppWebUrl, null));
            }
        }

        public SharePointHighTrustContext(Uri spHostUrl, Uri spAppWebUrl, string spLanguage, string spClientTag, string spProductNumber, WindowsIdentity logonUserIdentity)
            : base(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber)
        {
            if (logonUserIdentity == null)
            {
                throw new ArgumentNullException("logonUserIdentity");
            }

            this.logonUserIdentity = logonUserIdentity;
        }

        /// <summary>
        /// 确保访问令牌有效并返回该令牌。
        /// </summary>
        /// <param name="accessToken">要验证的访问令牌。</param>
        /// <param name="tokenRenewalHandler">标记续订处理程序。</param>
        /// <returns>访问令牌字符串。</returns>
        private static string GetAccessTokenString(ref Tuple<string, long> accessToken, Func<string> tokenRenewalHandler)
        {
            RenewAccessTokenIfNeeded(ref accessToken, tokenRenewalHandler);

            return IsAccessTokenValid(accessToken) ? accessToken.Item1 : null;
        }

        /// <summary>
        /// 如果访问令牌无效，则应续订访问令牌。
        /// </summary>
        /// <param name="accessToken">要续订的访问令牌。</param>
        /// <param name="tokenRenewalHandler">标记续订处理程序。</param>
        private static void RenewAccessTokenIfNeeded(ref Tuple<string, long> accessToken, Func<string> tokenRenewalHandler)
        {
            if (IsAccessTokenValid(accessToken))
            {
                return;
            }

            long expiresOn = TokenHelper.EpochTimeNow() + (long)TokenHelper.HighTrustAccessTokenLifetime.TotalSeconds;

            if (TokenHelper.HighTrustAccessTokenLifetime.TotalSeconds > AccessTokenLifetimeTolerance)
            {
                // 在访问令牌到期之前稍微提前一些进行续订
                // 以便使用它的对 SharePoint 的调用将有足够的时间来成功完成操作。
                expiresOn -= AccessTokenLifetimeTolerance;
            }

            accessToken = Tuple.Create(tokenRenewalHandler(), expiresOn);
        }
    }

    /// <summary>
    /// SharePointHighTrustContext 的默认提供程序。
    /// </summary>
    public class SharePointHighTrustContextProvider : SharePointContextProvider
    {
        private const string SPContextKey = "SPContext";

        protected override SharePointContext CreateSharePointContext(Uri spHostUrl, Uri spAppWebUrl, string spLanguage, string spClientTag, string spProductNumber, HttpRequestBase httpRequest)
        {
            WindowsIdentity logonUserIdentity = httpRequest.LogonUserIdentity;
            if (logonUserIdentity == null || !logonUserIdentity.IsAuthenticated || logonUserIdentity.IsGuest || logonUserIdentity.User == null)
            {
                return null;
            }

            return new SharePointHighTrustContext(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber, logonUserIdentity);
        }

        protected override bool ValidateSharePointContext(SharePointContext spContext, HttpContextBase httpContext)
        {
            SharePointHighTrustContext spHighTrustContext = spContext as SharePointHighTrustContext;

            if (spHighTrustContext != null)
            {
                Uri spHostUrl = SharePointContext.GetSPHostUrl(httpContext.Request);
                WindowsIdentity logonUserIdentity = httpContext.Request.LogonUserIdentity;

                return spHostUrl == spHighTrustContext.SPHostUrl &&
                       logonUserIdentity != null &&
                       logonUserIdentity.IsAuthenticated &&
                       !logonUserIdentity.IsGuest &&
                       logonUserIdentity.User == spHighTrustContext.LogonUserIdentity.User;
            }

            return false;
        }

        protected override SharePointContext LoadSharePointContext(HttpContextBase httpContext)
        {
            return httpContext.Session[SPContextKey] as SharePointHighTrustContext;
        }

        protected override void SaveSharePointContext(SharePointContext spContext, HttpContextBase httpContext)
        {
            httpContext.Session[SPContextKey] = spContext as SharePointHighTrustContext;
        }
    }

    #endregion HighTrust
}
