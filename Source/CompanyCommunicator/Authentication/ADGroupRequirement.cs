// <copyright file="ADGroupRequirement.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Authentication
{
    using Microsoft.AspNetCore.Authorization;

    /// <summary>
    /// This class is an authorization policy requirement.
    /// It specifies that an access token must contain group.read.all scope.
    /// </summary>
    public class ADGroupRequirement : IAuthorizationRequirement
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ADGroupRequirement"/> class.
        /// </summary>
        /// <param name="scopes">Microsoft Graph Scopes.</param>
        /// <param name="adGroup">AD Group.</param>
        public ADGroupRequirement(string[] scopes, string adGroup)
        {
            this.Scopes = scopes;
            this.Group = adGroup;
        }

        /// <summary>
        /// Gets microsoft Graph Scopes.
        /// </summary>
        public string[] Scopes { get; private set; }

       /// <summary>
       /// Gets microsoft Graph group.
       /// </summary>
        public string Group { get; private set; }
    }
}
