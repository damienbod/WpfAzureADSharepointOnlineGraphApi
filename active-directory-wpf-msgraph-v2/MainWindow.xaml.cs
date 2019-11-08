using Microsoft.Identity.Client;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using System;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;

namespace active_directory_wpf_msgraph_v2
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //Set the API Endpoint to Graph 'me' endpoint
        string graphAPIEndpoint = "https://graph.microsoft.com/v1.0/me";

        //Set the scope for API call to user.read
        string[] scopes = new string[] {   "user.read" };

        public ClientContext clientContext { get; set; }
        public string access_token = string.Empty;

        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Call AcquireToken - to acquire a token requiring user to sign-in
        /// </summary>
        private async void CallGraphButton_Click(object sender, RoutedEventArgs e)
        {
            //AuthenticationManager mgr = new AuthenticationManager();
            //using (var context = mgr.GetAzureADNativeApplicationAuthenticatedContext(
            //    "https://damienbodsharepoint.sharepoint.com",
            //    "411f99c2-19f6-403d-b6db-376bf9a597ad",
            //    "urn:ietf:wg:oauth:2.0:oob"))
            //{
            //    context.Load(context.Web, web => web.Title);
            //    context.ExecuteQuery();
            //    Console.WriteLine(context.Web.Title);
            //    Console.ReadKey();
            //}

            AuthenticationResult authResult = null;
            var app = App.PublicClientApp;
            ResultText.Text = string.Empty;
            //TokenInfoText.Text = string.Empty;

            var accounts = await app.GetAccountsAsync();
            var firstAccount = accounts.FirstOrDefault();

            try
            {
                var tt = app.Authority;
                authResult = await app.AcquireTokenSilent(scopes, firstAccount)
                    .ExecuteAsync();
            }
            catch (MsalUiRequiredException ex)
            {
                // A MsalUiRequiredException happened on AcquireTokenSilent. 
                // This indicates you need to call AcquireTokenInteractive to acquire a token
                System.Diagnostics.Debug.WriteLine($"MsalUiRequiredException: {ex.Message}");

                try
                {
                    authResult = await app.AcquireTokenInteractive(scopes)
                        .WithAccount(accounts.FirstOrDefault())
                        .WithParentActivityOrWindow(new WindowInteropHelper(this).Handle) // optional, used to center the browser on the window
                        .WithPrompt(Prompt.SelectAccount)
                        .ExecuteAsync();
                }
                catch (MsalException msalex)
                {
                    ResultText.Text = $"Error Acquiring Token:{System.Environment.NewLine}{msalex}";
                }
            }
            catch (Exception ex)
            {
                ResultText.Text = $"Error Acquiring Token Silently:{System.Environment.NewLine}{ex}";
                return;
            }

            if (authResult != null)
            {
                access_token = authResult.AccessToken;


                //AuthenticationManager mgr = new AuthenticationManager();
                //using (var context = mgr.GetAzureADAccessTokenAuthenticatedContext(
                //    "https://damienbodsharepoint.sharepoint.com/sites/listview",
                //    access_token))
                //{
                //    context.Load(context.Web, web => web.Title);
                //    context.ExecuteQuery();
                //    Console.WriteLine(context.Web.Title);
                //    Console.ReadKey();
                //}

                await GetDamienTest();
                // http://server/site/_api/site
                var url = "https://damienbodsharepoint.sharepoint.com/search/docs/Forms/AllItems.aspx";
                ResultText.Text = await GetHttpContentWithToken(url, authResult.AccessToken);

                // https://damienbodsharepoint.sharepoint.com/search/docs/Forms/AllItems.aspx
                //ResultText.Text = await GetHttpContentWithToken(graphAPIEndpoint, authResult.AccessToken);
                ResultText.Text = await GetHttpContentWithToken("https://damienbodsharepoint.sharepoint.com/sites/listview", authResult.AccessToken);
                DisplayBasicTokenInfo(authResult);
                this.SignOutButton.Visibility = Visibility.Visible;
            }
        }

        /// <summary>
        /// Perform an HTTP GET request to a URL using an HTTP Authorization header
        /// </summary>
        /// <param name="url">The URL</param>
        /// <param name="token">The token</param>
        /// <returns>String containing the results of the GET operation</returns>
        public async Task<string> GetHttpContentWithToken(string url, string token)
        {
            var httpClient = new System.Net.Http.HttpClient();
            System.Net.Http.HttpResponseMessage response;
            try
            {
                var request = new System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Get, url);
                //Add the token in Authorization header
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                response = await httpClient.SendAsync(request);
                var content = await response.Content.ReadAsStringAsync();
                return content;
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
        }

        /// <summary>
        /// Sign out the current user
        /// </summary>
        private async void SignOutButton_Click(object sender, RoutedEventArgs e)
        {
            var accounts = await App.PublicClientApp.GetAccountsAsync();
            if (accounts.Any())
            {
                try
                {
                    await App.PublicClientApp.RemoveAsync(accounts.FirstOrDefault());
                    this.ResultText.Text = "User has signed-out";
                    this.CallGraphButton.Visibility = Visibility.Visible;
                    this.SignOutButton.Visibility = Visibility.Collapsed;
                }
                catch (MsalException ex)
                {
                    ResultText.Text = $"Error signing-out user: {ex.Message}";
                }
            }
        }

        /// <summary>
        /// Display basic information contained in the token
        /// </summary>
        private void DisplayBasicTokenInfo(AuthenticationResult authResult)
        {
            TokenInfoText.Text = "";
            if (authResult != null)
            {
                TokenInfoText.Text += $"Username: {authResult.Account.Username}" + Environment.NewLine;
                TokenInfoText.Text += $"Token Expires: {authResult.ExpiresOn.ToLocalTime()}" + Environment.NewLine;
            }
        }

        private async Task GetDamienTest()
        {
            try
            {
                string spSiteUrl = "https://damienbodsharepoint.sharepoint.com/sites/listview";
                ClientContext context = new ClientContext(spSiteUrl);
                context.ExecutingWebRequest += context_ExecutingWebRequest;
                List documents = context.Web.Lists.GetByTitle("newnew");
                context.Load(documents, list => list.DefaultViewUrl);
                // Execute the query to server.
                await context.ExecuteQueryAsync();
            }
            catch (Exception ex)
            {
                string dd = ex.Message;
            }
        }

        void context_ExecutingWebRequest(object sender, WebRequestEventArgs e)
        {
            e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + access_token;
        }
    }
}
