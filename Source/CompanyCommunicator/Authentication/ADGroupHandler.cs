// <copyright file="ADGroupHandler.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Authentication
{
    using System;
    using System.Net.Http.Headers;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.Graph;
    using Microsoft.Identity.Web;

    /// <summary>
    /// This class is an authorization handler, which handles the authorization requirement based on AD Groups for the CC web methods.
    /// </summary>
    public class ADGroupHandler : AuthorizationHandler<ADGroupRequirement>
    {
        private readonly ITokenAcquisition tokenAcquisition; // token acquisition service

        /// <summary>
        /// Initializes a new instance of the <see cref="ADGroupHandler"/> class.
        /// </summary>
        /// <param name="tokenAcquisition">MSAL.NET token acquisition service.</param>
        public ADGroupHandler(ITokenAcquisition tokenAcquisition)
        {
            this.tokenAcquisition = tokenAcquisition;
        }

        /// <summary>
        /// This method handles the authorization requirement.
        /// </summary>
        /// <param name="context">AuthorizationHandlerContext instance.</param>
        /// <param name="requirement">IAuthorizationRequirement instance.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task HandleRequirementAsync(AuthorizationHandlerContext context, ADGroupRequirement requirement)
        {
            // get the token to connect to the graph API
            var accessToken = await this.tokenAcquisition.GetAccessTokenForUserAsync(requirement.Scopes);

            // graph client
            var graphClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    (requestMessage) =>
                    {
                        if (!string.IsNullOrEmpty(accessToken))
                        {
                            // Configure the HTTP bearer Authorization Header
                            requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken);
                        }
                        else
                        {
                            throw new Exception("Invalid authorization context");
                        }

                        return Task.FromResult(0);
                    }
                ));
            // get all the groups for the logged user
            var groups = graphClient.Me.MemberOf.Request().GetAsync().Result;

            // check if the group specified on the configuration is part of the list for the user
            var result = false;
            foreach (var group in groups)
            {
                // if we find the group, return true
                if (requirement.Group.Equals(group.Id))
                {
                    result = true;
                }
            }

            if (result)
            {
                context.Succeed(requirement);
            }

        }
    }
}
