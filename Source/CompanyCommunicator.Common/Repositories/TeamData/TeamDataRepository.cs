// <copyright file="TeamDataRepository.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TeamData
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;

    /// <summary>
    /// Repository of the team data stored in the table storage.
    /// </summary>
    public class TeamDataRepository : BaseRepository<TeamDataEntity>, ITeamDataRepository
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="TeamDataRepository"/> class.
        /// </summary>
        /// <param name="logger">The logging service.</param>
        /// <param name="repositoryOptions">Options used to create the repository.</param>
        public TeamDataRepository(
            ILogger<TeamDataRepository> logger,
            IOptions<RepositoryOptions> repositoryOptions)
            : base(
                  logger,
                  storageAccountConnectionString: repositoryOptions.Value.StorageAccountConnectionString,
                  tableName: TeamDataTableNames.TableName,
                  defaultPartitionKey: TeamDataTableNames.TeamDataPartition,
                  ensureTableExists: repositoryOptions.Value.EnsureTableExists)
        {
        }

        /// <inheritdoc/>
        public async Task<IEnumerable<TeamDataEntity>> GetTeamDataEntitiesByIdsAsync(IEnumerable<string> teamIds)
        {
            if (teamIds == null || !teamIds.Any())
            {
                return new List<TeamDataEntity>();
            }

            List<TeamDataEntity> teamDataEntities = new List<TeamDataEntity>();
            string rowKeyFilter = string.Empty;

            // batch the calls to azure storage, Url was too long (max url (2048 char) / channelId (60 char))
            var maxNoPerFilter = 20;
            var batchAmount = (int)Math.Ceiling((double)teamIds.Count() / maxNoPerFilter);

            for (var i = 0; i < batchAmount; i++)
            {
                var currentIds = teamIds.Skip(i * maxNoPerFilter).Take(maxNoPerFilter);

                rowKeyFilter = this.GetRowKeysFilter(currentIds);

                var data = await this.GetWithFilterAsync(rowKeyFilter);

                teamDataEntities.AddRange(data);

                rowKeyFilter = string.Empty;
            }

            return teamDataEntities;
        }

        /// <inheritdoc/>
        public async Task<IEnumerable<string>> GetTeamNamesByIdsAsync(IEnumerable<string> ids)
        {
            IEnumerable<TeamDataEntity> teamDataEntities = await this.GetTeamDataEntitiesByIdsAsync(ids);

            return teamDataEntities.Select(p => p.Name).OrderBy(p => p);
        }

        /// <inheritdoc/>
        public async Task<IEnumerable<TeamDataEntity>> GetAllSortedAlphabeticallyByNameAsync()
        {
            var teamDataEntities = await this.GetAllAsync();
            var sortedSet = new SortedSet<TeamDataEntity>(teamDataEntities, new TeamDataEntityComparer());
            return sortedSet;
        }

        private class TeamDataEntityComparer : IComparer<TeamDataEntity>
        {
            public int Compare(TeamDataEntity x, TeamDataEntity y)
            {
                return x.Name.CompareTo(y.Name);
            }
        }
    }
}
