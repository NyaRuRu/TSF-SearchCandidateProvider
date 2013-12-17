using System;
using System.Collections.Generic;
using System.Linq;
using TSF.TypeLib;

namespace TSF.SearchCandidateProviderInternal
{
    internal static partial class TSFEx
    {
        /// <summary>
        /// An iterator-based implementation that encapsulates COM reference counting operation around ITfFnSearchCandidateProvider.GetSearchCandidates.
        /// </summary>
        public static IEnumerable<string> GetSearchCandidates(this ITfFnSearchCandidateProvider provider, string query, string applicationId = "")
        {
            using (var releaser = new ComReleaser())
            {
                var candidateList = releaser.ReceiveObject((out ITfCandidateList _) => provider.GetSearchCandidates(query, applicationId, out _));
                if (candidateList == null)
                {
                    yield break;
                }
                var numResults = default(uint);
                if (!candidateList.GetCandidateNum(out numResults))
                {
                    yield break;
                }
                var iter = Enumerable.Range(0, (int)numResults)
                                     .Select(index => releaser.ReceiveObject((out ITfCandidateString _) => candidateList.GetCandidate((uint)index, out _)))
                                     .Where(candidateString => candidateString != null)
                                     .Select(candidateString =>
                                     {
                                         var resultString = string.Empty;
                                         candidateString.GetString(out resultString);
                                         return resultString;
                                     })
                                     .Where(str => !String.IsNullOrEmpty(str));
                foreach (var candiate in iter)
                {
                    yield return candiate;
                }
            }
        }
    }
}
