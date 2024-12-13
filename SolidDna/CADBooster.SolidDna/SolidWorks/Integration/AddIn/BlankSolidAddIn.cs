namespace CADBooster.SolidDna
{
    /// <summary>
    /// Creates a blank AddIn integration class
    /// </summary>
    public class BlankSolidAddIn : SolidAddIn
    {
        #region Constructor
        /// <summary>
        /// Default constructor
        /// </summary>
        public BlankSolidAddIn() : base() { }
        #endregion

        #region AddIn Methods
        public override string SolidWorksAddInTitle => "SolidDNA Add-In Title";
        public override string SolidWorksAddInDescription => "SolidDNA Add-In Description";
        public override void ApplicationStartup() { }
        public override void PreConnectToSolidWorks() { }
        public override void ConnectedToSolidWorks() { }
        public override void DisconnectedFromSolidWorks() { }
        #endregion
    }
}
