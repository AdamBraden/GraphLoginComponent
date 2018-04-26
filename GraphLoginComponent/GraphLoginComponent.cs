using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Toolkit.Services.MicrosoftGraph;

namespace GraphLoginComponent
{
    public partial class GraphLoginComponent : Component
    {
        private string clientId;
        private string[] scopes;
        private GraphServiceClient graphServiceClient;
        private string displayName;
        private string jobTitle;
        private string email;
        private System.Drawing.Image photo;

        public GraphLoginComponent()
        {
            InitializeComponent();
        }

        public GraphLoginComponent(IContainer container)
        {
            container.Add(this);

            InitializeComponent();
        }
        public string ClientId { get => clientId; set => clientId = value; }
        public string[] Scopes { get => scopes; set => scopes = value; }
        public System.Drawing.Image Photo { get => photo; set => photo = value; }
        public string DisplayName { get => displayName; set => displayName = value; }
        public string JobTitle { get => jobTitle; set => jobTitle = value; }
        public string Email { get => email; set => email = value; }
        public GraphServiceClient GraphServiceClient { get => graphServiceClient; set => graphServiceClient = value; }

        /// <summary>
        /// LoginAsync provides entrypoint into the MicrosoftGraphService LoginAsync
        /// </summary>
        /// <returns>A MicrosoftGraphService reference</returns>
        public async Task<bool> LoginAsync()
        {
            // check inputs
            if (string.IsNullOrEmpty(clientId))
            {
                //error
                return false;
            }

            if (!MicrosoftGraphService.Instance.Initialize(clientId,delegatedPermissionScopes: Scopes))
            {
                return false;
            }

            // login and return
            try
            {
                await MicrosoftGraphService.Instance.LoginAsync();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                return false;
            }

            // Initialize the User
            var user = await MicrosoftGraphService.Instance.GraphProvider.Me.Request().GetAsync();
            displayName = user.DisplayName;
            jobTitle = user.JobTitle;
            email = user.Mail;
            
            // get the profile picture 
            using (Stream photoStream = await MicrosoftGraphService.Instance.GraphProvider.Me.Photo.Content.Request().GetAsync())
            {
                if (photoStream != null)
                {
                    photo = System.Drawing.Image.FromStream(photoStream);
                }
            }

            // Return MicrosoftGraphService or GraphProvider (SDK's GraphServiceClient)?
            // return MicrosoftGraphService.Instance.GraphProvider;
            graphServiceClient = MicrosoftGraphService.Instance.GraphProvider;
            return true;
            
        }
    }
}
