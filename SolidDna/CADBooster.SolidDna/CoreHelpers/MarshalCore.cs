using System;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Security;

namespace CADBooster.SolidDna.CoreHelpers
{
    /// <summary>
    ///     Provides methods for marshalling objects and interacting with COM objects.
    /// </summary>
    /// <remarks>
    ///     This class is security-critical, which means it cannot be used by partially trusted code.
    /// </remarks>
    public static class MarshalCore
    {
        private const string Ole32 = "ole32.dll";
        private const string Oleaut32 = "oleaut32.dll";

        /// <summary>
        ///     Retrieves the active COM object with the specified ProgID.
        /// </summary>
        /// <param name="progId">The ProgID of the COM object.</param>
        /// <returns>The active COM object associated with the specified ProgID.</returns>
        /// <remarks>
        ///     This method first attempts to call CLSIDFromProgIDEx to get the CLSID for the specified ProgID.
        ///     If CLSIDFromProgIDEx does not exist, it falls back on CLSIDFromProgID.
        /// </remarks>
        /// <exception cref="Exception">Thrown when neither CLSIDFromProgIDEx nor CLSIDFromProgID can provide a valid CLSID.</exception>
        [SecurityCritical] public static object GetActiveObject(string progId)
        {
            Guid clsid;

            // Call CLSIDFromProgIDEx first then fall back on CLSIDFromProgID if CLSIDFromProgIDEx doesn't exist.
            try
            {
                CLSIDFromProgIDEx(progId, out clsid);
            }
            catch (Exception)
            {
                CLSIDFromProgID(progId, out clsid);
            }

            GetActiveObject(ref clsid, IntPtr.Zero, out var obj);
            return obj;
        }

        [DllImport(Ole32, PreserveSig = false)] [ResourceExposure(ResourceScope.None)] [SuppressUnmanagedCodeSecurity] [SecurityCritical]
        private static extern void CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string progId, out Guid clsid);

        [DllImport(Ole32, PreserveSig = false)] [ResourceExposure(ResourceScope.None)] [SuppressUnmanagedCodeSecurity] [SecurityCritical]
        private static extern void CLSIDFromProgIDEx([MarshalAs(UnmanagedType.LPWStr)] string progId, out Guid clsid);

        [DllImport(Oleaut32, PreserveSig = false)] [ResourceExposure(ResourceScope.None)] [SuppressUnmanagedCodeSecurity] [SecurityCritical]
        private static extern void GetActiveObject(ref Guid rclsid, IntPtr reserved, [MarshalAs(UnmanagedType.Interface)] out object ppunk);
    }
}