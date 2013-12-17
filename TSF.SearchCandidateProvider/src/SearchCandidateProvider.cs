using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using TSF.InteropTypes;
using TSF.SearchCandidateProviderInternal;
using TSF.TypeLib;

namespace TSF
{
    /// <summary>
    /// Represents the identity of a Text Input Processor that is supposed to support ITfFnSearchCandidateProvider.
    /// </summary>
    public class SearchCandidateProviderProfile
    {
        public SearchCandidateProviderProfile(CultureInfo cultureInfo, Guid clsid, Guid profileGuid, string name = "")
        {
            Name = name;
            LangId = cultureInfo.ToLangID();
            Clsid = clsid;
            ProfileGuid = profileGuid;
        }
        public readonly string Name;
        public readonly LangID LangId;
        public readonly Guid Clsid;
        public readonly Guid ProfileGuid;
    }

    /// <summary>
    /// Represents the execution result of ITfFnSearchCandidateProvider.GetSearchCandidates.
    /// </summary>
    public class SearchCandidateResponse
    {
        public SearchCandidateResponse(SearchCandidateProviderProfile provider, string queryString, string[] candidates, TimeSpan elapsed, Exception exception = null)
        {
            Provider = provider;
            QueryString = queryString;
            Candidates = candidates;
            Elaplsed = elapsed;
            Exception = exception;
        }
        public readonly SearchCandidateProviderProfile Provider;
        public readonly string QueryString;
        public readonly string[] Candidates;
        public readonly TimeSpan Elaplsed;
        public readonly Exception Exception;
    }

    /// <summary>
    /// A managed wrapper of ITfFnSearchCandidateProvider.
    /// </summary>
    public class SearchCandidateProvider : IDisposable
    {
        /// <summary>
        /// Returns if the IME settings is per-thread mode or not.
        /// </summary>
        /// <param name="defaultValue">The value to be returned when fails to retrieve the settings.</param>
        /// <returns>True if the IME settings is per-thread mode.</returns>
        /// <remarks>SearchCandidateProvider.FromProfile is available as long as the IME settings is per-thread mode.</remarks>
        public static bool GetThreadLocalInputSettings(bool defaultValue)
        {
            return SystemParametersInfo.GetThreadLocalInputSettings(defaultValue);
        }
        /// <summary>
        /// Updates the IME settings to be per-thread or not.
        /// </summary>
        /// <param name="value">True if the IME settings should be set to per-thread mode.</param>
        public static void SetThreadLocalInputSettings(bool value)
        {
            SystemParametersInfo.SetThreadLocalInputSettings(value);
        }

        /// <summary>
        /// Starts initialization of SearchCandidateProvider with <paramref name="profile"/>.
        /// </summary>
        /// <remarks>The caller must ensure that the IME settings is kept to be per-thread mode from the method is called until the returned task is completed.</remarks>
        /// <param name="profile">The profile information with which SearchCandidateProvider is instantiated.</param>
        /// <returns>A task object which represents an asynchronous operation to initialize a SearchCandidateProvider object.
        /// The resolved value can be null if any error occurs.</returns>
        static public Task<SearchCandidateProvider> FromProfile(SearchCandidateProviderProfile profile)
        {
            if (!(GetThreadLocalInputSettings(defaultValue: true)))
            {
                throw new InvalidOperationException("Aborted because this operation is dangerous while SPI_GETTHREADLOCALINPUTSETTINGS is returning true.");
            }
            if (profile == null)
            {
                throw new ArgumentException("null is not allowed", "profile");
            }
            return CreateInternal(profile);
        }

        /// <summary>
        /// Starts initialization of SearchCandidateProvider with <paramref name="profile"/>.
        /// </summary>
        /// <returns>A task object which represents an asynchronous operation to initialize a SearchCandidateProvider object.
        /// The resolved value can be null if any error occurs.</returns>
        static public Task<SearchCandidateProvider> FromCurrentProfile()
        {
            return CreateInternal(null);
        }

        public Task<SearchCandidateResponse> GetSearchCandidates(string query)
        {
            var response = new TaskCompletionSource<SearchCandidateResponse>();
            RequestQueue.Add(new RequestQueueItem() { Query = query, Response = response });
            return response.Task;
        }

        static private Task<SearchCandidateProvider> CreateInternal(SearchCandidateProviderProfile profile)
        {
            var createComplete = new TaskCompletionSource<SearchCandidateProvider>();
            var thread = new Thread(() => SearchCandidateProvider.ThreadMain(profile, createComplete));
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            var threadName = "Query Thread" + (profile == null ? "" : ": " + profile.Name);
            thread.Name = threadName;
            return createComplete.Task;
        }
        
        public void Dispose()
        {
            // Note: |RequestQueue| is onwed by the background thread.  Should not call Dispose here.
            RequestQueue.CompleteAdding();
        }

        internal static void ThreadMain(SearchCandidateProviderProfile profile, TaskCompletionSource<SearchCandidateProvider> createComplete)
        {
            using (var mainReleaser = new ComReleaser())
            {
                try
                {
                    var profileMgr = mainReleaser.CreateComObject<ITfInputProcessorProfileMgr>();
                    if (profile == null)
                    {
                        var keyboardGuid = Guids.GUID_TFCAT_TIP_KEYBOARD;
                        var activeProfile = default(TF_INPUTPROCESSORPROFILE);
                        if (!profileMgr.GetActiveProfile(ref keyboardGuid, out activeProfile))
                        {
                            return;
                        }
                        if (activeProfile.dwProfileType == TF_PROFILETYPE.TF_PROFILETYPE_KEYBOARDLAYOUT)
                        {
                            // This is not a TIP.
                            return;
                        }
                        profile = new SearchCandidateProviderProfile(activeProfile.langid.ToCultureInfo(), activeProfile.clsid, activeProfile.guidProfile);
                    }
                    var name = profile.Name;
                    var langid = profile.LangId;
                    var clsid = profile.Clsid;
                    var profileGuid = profile.ProfileGuid;

                    if (!profileMgr.ActivateProfile(TF_PROFILETYPE.TF_PROFILETYPE_INPUTPROCESSOR, langid, ref clsid, ref profileGuid, IntPtr.Zero, TF_IPPMF.TF_IPPMF_DONTCARECURRENTINPUTLANGUAGE))
                    {
                        return;
                    }

                    var threadMgr = mainReleaser.CreateComObject<ITfThreadMgr2>();
                    var clientId = TextFrameworkDeclarations.TF_CLIENTID_NULL;
                    if (!threadMgr.ActivateEx(out clientId, default(TF_TMAE)))
                    {
                        return;
                    }
                    mainReleaser.RegisterCleanup(() => threadMgr.Deactivate());

                    var functionProvider = mainReleaser.ReceiveObject((out ITfFunctionProvider _) => threadMgr.GetFunctionProvider(ref clsid, out _));
                    if (functionProvider == null)
                    {
                        return;
                    }
                    if (string.IsNullOrEmpty(name))
                    {
                        functionProvider.GetDescription(out name);
                    }
                    var guidNull = new Guid();
                    var iid = typeof(ITfFnSearchCandidateProvider).GUID;
                    var searchCandidateProvider = mainReleaser.ReceiveObject((out object _) => functionProvider.GetFunction(ref guidNull, ref iid, out _)) as ITfFnSearchCandidateProvider;
                    if (searchCandidateProvider == null)
                    {
                        return;
                    }
                    if (string.IsNullOrEmpty(name))
                    {
                        searchCandidateProvider.GetDisplayName(out name);
                    }

                    // Refresh the |profile| with the latest |name|.
                    profile = new SearchCandidateProviderProfile(profile.LangId.ToCultureInfo(), profile.Clsid, profile.ProfileGuid, name);

                    // It's time to enter the event loop.
                    using (var queue = new BlockingCollection<RequestQueueItem>())
                    {
                        createComplete.SetResult(new SearchCandidateProvider { RequestQueue = queue, Profile = profile });
                        foreach (var task in queue.GetConsumingEnumerable())
                        {
                            using (var releaser = new ComReleaser())
                            {
                                var stopWatch = Stopwatch.StartNew();
                                var response = task.Response;
                                try
                                {
                                    var result = searchCandidateProvider.GetSearchCandidates(task.Query).ToArray();
                                    response.SetResult(new SearchCandidateResponse(profile, task.Query, result, stopWatch.Elapsed));
                                }
                                catch (Exception e)
                                {
                                    if (!response.Task.IsCompleted)
                                    {
                                        response.SetResult(new SearchCandidateResponse(profile, task.Query, Enumerable.Empty<string>().ToArray(), stopWatch.Elapsed, e));
                                    }
                                }
                                finally
                                {
                                    if (!response.Task.IsCompleted)
                                    {
                                        response.SetResult(new SearchCandidateResponse(profile, task.Query, Enumerable.Empty<string>().ToArray(), stopWatch.Elapsed));
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    if (!createComplete.Task.IsCompleted)
                    {
                        createComplete.SetException(e);
                    }
                }
                finally
                {
                    if (!createComplete.Task.IsCompleted)
                    {
                        createComplete.SetResult(null);
                    }
                }
            }
        }

        public SearchCandidateProviderProfile Profile { get; private set; }

        private class RequestQueueItem
        {
            public string Query { get; set; }
            public TaskCompletionSource<SearchCandidateResponse> Response { get; set; }
        }
        private BlockingCollection<RequestQueueItem> RequestQueue { get; set; }
    }
}
