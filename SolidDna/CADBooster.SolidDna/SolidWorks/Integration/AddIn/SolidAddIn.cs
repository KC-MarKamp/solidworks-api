using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using System;

namespace CADBooster.SolidDna
{
    /// <summary>
    /// Integrates into SolidWorks as an add-in and registers for callbacks provided by SolidWorks
    /// IMPORTANT: The class that overrides <see cref="ISwAddin" /> MUST be the same class that
    /// contains the ComRegister and ComUnregister functions due to how SolidWorks loads add-ins
    /// </summary>
    public abstract class SolidAddIn : ISwAddin
    {
        #region Protected Members
        /// <summary>
        /// Flag if we have loaded into memory (as ConnectedToSolidWorks can happen multiple times if unloaded/reloaded)
        /// </summary>
        protected bool mLoaded;
        #endregion

        #region Public Properties
        /// <summary>
        /// The title displayed for this SolidWorks Add-in
        /// </summary>
        public abstract string SolidWorksAddInTitle { get; }

        /// <summary>
        /// The description displayed for this SolidWorks Add-in
        /// </summary>
        public abstract string SolidWorksAddInDescription { get; }
        #endregion

        /// <summary>
        /// Default constructor
        /// </summary>
        public SolidAddIn() { }

        #region Public Abstract / Virtual Methods
        /// <summary>
        /// Specific application startup code when SolidWorks is connected
        /// and before any plug-ins or listeners are informed.
        /// Runs after <see cref="PreConnectToSolidWorks" /> and after <see cref="PreLoadPlugIns" />.
        /// NOTE: This call will not be in the same AppDomain as the SolidDna plug-ins
        /// </summary>
        /// <returns></returns>
        public abstract void ApplicationStartup();

        /// <summary>
        /// Run immediately when <see cref="ConnectToSW(object, int)" /> is called to do any pre-setup.
        /// For example, call <see cref="Logger.AddFileLogger{TAddIn}" /> to add a file logger for SolidDna messages.
        /// Runs before <see cref="PreLoadPlugIns" /> and before <see cref="ApplicationStartup" />.
        /// </summary>
        public abstract void PreConnectToSolidWorks();

        public abstract void ConnectedToSolidWorks();

        public abstract void DisconnectedFromSolidWorks();
        #endregion

        #region SolidWorks Add-in Callbacks
        /// <summary>
        /// Receives all callbacks from command manager items and flyouts.
        /// We tell SolidWorks to call a method in the <see cref="SolidAddIn" /> class in <see cref="SetUpCallbacks" />
        /// We tell it to call this method in <see cref="CommandManagerGroup.AddCommandItem" /> and <see cref="CommandManagerFlyout.AddCommandItem" />.
        /// We forward this to <see cref="PlugInIntegration.OnCallback" />, which then finds the correct command manager item or flyout and calls its OnClick method.
        /// </summary>
        /// <param name="arg"></param>
        public void Callback(string arg)
        {
            // Log it
            Logger.LogDebugSource($"SolidWorks Callback fired {arg}");

            AddInIntegration.OnCallback(arg);
        }

        /// <summary>
        /// Called when SolidWorks has loaded our add-in and wants us to do our connection logic
        /// </summary>
        /// <param name="thisSw">The current SolidWorks instance</param>
        /// <param name="cookie">The current SolidWorks cookie ID</param>
        /// <returns></returns>
        public bool ConnectToSW(object thisSw, int cookie)
        {
            try
            {
                // Add this add-in to the list of currently active add-ins.
                AddInIntegration.AddAddIn(this);

                // Fire event
                PreConnectToSolidWorks();

                // Log it
                Logger.LogTraceSource($"Fired PreConnectToSolidWorks...");

                // Get the directory path to this actual add-in dll
                var assemblyPath = this.AssemblyPath();

                // Log it
                Logger.LogDebugSource($"{SolidWorksAddInTitle} Connected to SolidWorks...");

                //
                //   NOTE: Do not need to create it here, as we now create it inside PlugInIntegration.Setup in its own AppDomain
                //         If we change back to loading directly (not in an app domain) then uncomment this 
                //
                // Store a reference to the current SolidWorks instance
                // Initialize SolidWorks (SolidDNA class)
                //SolidWorks = new SolidWorksApplication((SldWorks)ThisSW, Cookie);

                // Tell solidworks which method to call when it receives a button click on a command manager item or flyout.
                SetUpCallbacks(thisSw, cookie);

                // Log it
                Logger.LogDebugSource($"Storing the SOLIDWORKS instance...");

                // Set up the current SolidWorks instance as a SolidDNA class.
                AddInIntegration.ConnectToActiveSolidWorks(((SldWorks)thisSw).RevisionNumber(), cookie);

                // Log it
                Logger.LogDebugSource($"Firing PreLoadPlugIns...");

                // If this is the first load
                if (!mLoaded)
                {
                    // Now loaded so don't do it again
                    mLoaded = true;
                }

                // Log it
                Logger.LogDebugSource($"Firing ApplicationStartup...");

                // Call the application startup function for an entry point to the application
                ApplicationStartup();

                // Log it
                Logger.LogDebugSource($"Firing ConnectedToSolidWorks...");

                // Inform listeners
                ConnectedToSolidWorks();

                // Return ok
                return true;
            }
            catch (Exception ex)
            {
                // Log it
                Logger.LogCriticalSource($"Unexpected error: {ex}");

                return false;
            }
        }

        /// <summary>
        /// Called when SolidWorks is about to unload our add-in and wants us to do our disconnection logic
        /// </summary>
        /// <returns></returns>
        public bool DisconnectFromSW()
        {
            // Log it
            Logger.LogDebugSource($"{SolidWorksAddInTitle} Disconnected from SolidWorks...");

            // Log it
            Logger.LogDebugSource($"Firing DisconnectedFromSolidWorks...");

            // Inform listeners
            DisconnectedFromSolidWorks();

            // Log it
            Logger.LogDebugSource($"Tearing down...");

            // Remove it from the list and tear down SOLIDWORKS when it was the last add-in.
            AddInIntegration.RemoveAddInAndTearDownSolidWorksWhenLast(this);

            // Remove the loggers for this add-in
            Logger.RemoveLoggers(this);

            // Reset mLoaded so we can restart this add-in
            mLoaded = false;

            // Return ok
            return true;
        }

        /// <summary>
        /// Tell SolidWorks that it should call the <see cref="Callback" /> method in this class whenever it receives a Command Manager item or flyout button click.
        /// We forward this to <see cref="PlugInIntegration.OnCallback" />, which then finds the correct command manager item or flyout and calls its OnClick method.
        /// </summary>
        /// <param name="thisSw"></param>
        /// <param name="cookie"></param>
        private void SetUpCallbacks(object thisSw, int cookie)
        {
            // Log it
            Logger.LogDebugSource($"Setting AddInCallbackInfo...");

            ((SldWorks)thisSw).SetAddinCallbackInfo2(0, this, cookie);
        }
        #endregion

        #region Connected to SolidWorks Event Calls
        /// <summary>
        /// When the add-in has connected to SolidWorks
        /// </summary>
        public void OnConnectedToSolidWorks()
        {
            // Log it
            Logger.LogDebugSource($"Firing ConnectedToSolidWorks event...");

            ConnectedToSolidWorks();
        }

        /// <summary>
        /// When the add-in has disconnected to SolidWorks
        /// </summary>
        public void OnDisconnectedFromSolidWorks()
        {
            // Log it
            Logger.LogDebugSource($"Firing DisconnectedFromSolidWorks event...");

            DisconnectedFromSolidWorks();
        }
        #endregion

        #region Com Registration
        ///// <summary>
        ///// The COM registration call to add our registry entries to the SolidWorks add-in registry
        ///// </summary>
        ///// <param name="t"></param>
        //[ComRegisterFunction] protected static void ComRegister(Type t)
        //{
        //    try
        //    {
        //        // Create new instance of a blank add-in
        //        var addIn = new BlankSolidAddIn();

        //        // Get assembly name
        //        var assemblyName = t.Assembly.Location;

        //        // Log it
        //        Logger.LogInformationSource($"Registering {assemblyName}");

        //        // Get registry key path
        //        var keyPath = string.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);

        //        // Create our registry folder for the add-in
        //        using (var rk = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(keyPath))
        //        {
        //            // Load add-in when SolidWorks opens
        //            rk.SetValue(null, 1);

        //            // Set SolidWorks add-in title and description
        //            rk.SetValue("Title", addIn.SolidWorksAddInTitle);
        //            rk.SetValue("Description", addIn.SolidWorksAddInDescription);

        //            Logger.LogInformationSource($"COM Registration successful. '{addIn.SolidWorksAddInTitle}' : '{addIn.SolidWorksAddInDescription}'");
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Debugger.Break();

        //        // Get the path to this DLL
        //        var assemblyLocation = typeof(SolidAddIn).AssemblyFilePath();

        //        // Create a path for a text file. The assembly location is always lowercase.
        //        var changeExtension = assemblyLocation.Replace(".dll", ".fatal.log.txt");

        //        // Log an error to a new or existing text file 
        //        File.AppendAllText(changeExtension, $"\r\nUnexpected error: {ex}");

        //        Logger.LogCriticalSource($"COM Registration error. {ex}");
        //        throw;
        //    }
        //}

        ///// <summary>
        ///// The COM unregister call to remove our custom entries we added in the COM register function
        ///// </summary>
        ///// <param name="t"></param>
        //[ComUnregisterFunction] protected static void ComUnregister(Type t)
        //{
        //    // Get registry key path
        //    var keyPath = string.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);

        //    // Remove our registry entry
        //    Microsoft.Win32.Registry.LocalMachine.DeleteSubKeyTree(keyPath);
        //}
        #endregion
    }
}
