namespace OPCLibrary
{
    using System;
    using System.Net;
    using System.Runtime.InteropServices;

    namespace DZ.Opc.Integration.Internal
    {
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct COSERVERINFO
        {
            public uint dwReserved1;

            [MarshalAs(UnmanagedType.LPWStr)]
            public string pwszName;

            public IntPtr pAuthInfo;
            public uint dwReserved2;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct COAUTHIDENTITY
        {
            public IntPtr User;
            public uint UserLength;
            public IntPtr Domain;
            public uint DomainLength;
            public IntPtr Password;
            public uint PasswordLength;
            public uint Flags;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct COAUTHINFO
        {
            public uint dwAuthnSvc;
            public uint dwAuthzSvc;
            public IntPtr pwszServerPrincName;
            public uint dwAuthnLevel;
            public uint dwImpersonationLevel;
            public IntPtr pAuthIdentityData;
            public uint dwCapabilities;
        }


        public class ServerInfo
        {
            // Fields
            private GCHandle m_hAuthInfo;
            private GCHandle m_hDomain;
            private GCHandle m_hIdentity;
            private GCHandle m_hPassword;
            private GCHandle m_hUserName;

            // Methods
            public COSERVERINFO Allocate(string hostName, NetworkCredential credential)
            {
                string userName = null;
                string password = null;
                string domain = null;
                if (credential != null)
                {
                    userName = credential.UserName;
                    password = credential.Password;
                    domain = credential.Domain;
                }
                this.m_hUserName = GCHandle.Alloc(userName, GCHandleType.Pinned);
                this.m_hPassword = GCHandle.Alloc(password, GCHandleType.Pinned);
                this.m_hDomain = GCHandle.Alloc(domain, GCHandleType.Pinned);
                this.m_hIdentity = new GCHandle();
                if ((userName != null) && (userName != string.Empty))
                {
                    COAUTHIDENTITY coauthidentity = new COAUTHIDENTITY
                    {
                        User = this.m_hUserName.AddrOfPinnedObject(),
                        UserLength = (userName != null) ? ((uint)userName.Length) : 0,
                        Password = this.m_hPassword.AddrOfPinnedObject(),
                        PasswordLength = (password != null) ? ((uint)password.Length) : 0,
                        Domain = this.m_hDomain.AddrOfPinnedObject(),
                        DomainLength = (domain != null) ? ((uint)domain.Length) : 0,
                        Flags = 2
                    };
                    this.m_hIdentity = GCHandle.Alloc(coauthidentity, GCHandleType.Pinned);
                }
                COAUTHINFO coauthinfo = new COAUTHINFO
                {
                    dwAuthnSvc = 10,
                    dwAuthzSvc = 0,
                    pwszServerPrincName = IntPtr.Zero,
                    dwAuthnLevel = 2,
                    dwImpersonationLevel = 3,
                    pAuthIdentityData = this.m_hIdentity.IsAllocated ? this.m_hIdentity.AddrOfPinnedObject() : IntPtr.Zero,
                    dwCapabilities = 0
                };
                this.m_hAuthInfo = GCHandle.Alloc(coauthinfo, GCHandleType.Pinned);
                return new COSERVERINFO
                {
                    pwszName = hostName,
                    pAuthInfo = (credential != null) ? this.m_hAuthInfo.AddrOfPinnedObject() : IntPtr.Zero,
                    dwReserved1 = 0,
                    dwReserved2 = 0
                };
            }

            public void Deallocate()
            {
                if (this.m_hUserName.IsAllocated)
                {
                    this.m_hUserName.Free();
                }
                if (this.m_hPassword.IsAllocated)
                {
                    this.m_hPassword.Free();
                }
                if (this.m_hDomain.IsAllocated)
                {
                    this.m_hDomain.Free();
                }
                if (this.m_hIdentity.IsAllocated)
                {
                    this.m_hIdentity.Free();
                }
                if (this.m_hAuthInfo.IsAllocated)
                {
                    this.m_hAuthInfo.Free();
                }
            }
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct MULTI_QI
        {
            public IntPtr iid;

            [MarshalAs(UnmanagedType.IUnknown)]
            public object pItf;

            public uint hr;
        }


        public class Interop
        {
            private static readonly Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");

            [DllImport("ole32.dll")]
            private static extern void CoCreateInstanceEx(ref Guid clsid,
                                                          [MarshalAs(UnmanagedType.IUnknown)] object punkOuter,
                                                          uint dwClsCtx, [In] ref COSERVERINFO pServerInfo, uint dwCount,
                                                          [In, Out] MULTI_QI[] pResults);

            [DllImport("Kernel32.dll")]
            private static extern int FormatMessageW(int dwFlags, IntPtr lpSource, int dwMessageId, int dwLanguageId,
                                                     IntPtr lpBuffer, int nSize, IntPtr Arguments);


            public static string GetSystemMessage(int error)
            {
                IntPtr lpBuffer = Marshal.AllocCoTaskMem(0x400);
                FormatMessageW(0x1000, IntPtr.Zero, error, 0, lpBuffer, 0x3ff, IntPtr.Zero);
                string str = Marshal.PtrToStringUni(lpBuffer);
                Marshal.FreeCoTaskMem(lpBuffer);
                if ((str != null) && (str.Length > 0))
                {
                    return str;
                }
                return string.Format("0x{0,0:X}", error);
            }

            public static object CreateInstance(Guid clsid, string hostName)
            {
                return CreateInstance(clsid, hostName, null);
            }

            public static object CreateInstance(Guid clsid, string hostName, NetworkCredential credential)
            {
                ServerInfo info = new ServerInfo();
                COSERVERINFO pServerInfo = info.Allocate(hostName, credential);
                //pServerInfo.pwszName = "192.168.0.102";
                GCHandle handle = GCHandle.Alloc(IID_IUnknown, GCHandleType.Pinned);
                MULTI_QI[] pResults = new MULTI_QI[1];
                pResults[0].iid = handle.AddrOfPinnedObject();
                pResults[0].pItf = null;
                pResults[0].hr = 0;
                try
                {
                    uint dwClsCtx = 5;
                    if (((hostName != null) && (hostName.Length > 0)) && (hostName != "localhost"))
                    {
                        dwClsCtx = 0x10;
                    }
                    //  CoCreateInstanceEx()
                    CoCreateInstanceEx(ref clsid, null, dwClsCtx, ref pServerInfo, 1, pResults);
                }
                catch (Exception exception)
                {
                    throw exception;
                }
                finally
                {
                    if (handle.IsAllocated)
                    {
                        handle.Free();
                    }
                    info.Deallocate();
                }

                if (pResults[0].hr != 0)
                {
                    throw new ExternalException("CoCreateInstanceEx: " + GetSystemMessage((int)pResults[0].hr));
                }
                return pResults[0].pItf;
            }

            public static void ReleaseServer(object server)
            {
                if ((server != null) && server.GetType().IsCOMObject)
                {
                    Marshal.ReleaseComObject(server);
                }
            }

            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
            private struct SOLE_AUTHENTICATION_SERVICE
            {
                public uint dwAuthnSvc;
                public uint dwAuthzSvc;
                [MarshalAs(UnmanagedType.LPWStr)]
                public string pPrincipalName;
                public int hr;
            }

            [DllImport("ole32.dll")]
            private static extern int CoInitializeSecurity(IntPtr pSecDesc, int cAuthSvc, SOLE_AUTHENTICATION_SERVICE[] asAuthSvc, IntPtr pReserved1, uint dwAuthnLevel, uint dwImpLevel, IntPtr pAuthList, uint dwCapabilities, IntPtr pReserved3);

            /// <summary>
            /// This is a real magic method !
            /// </summary>
            public static void InitializeSecurity()
            {
                int errorCode = CoInitializeSecurity(IntPtr.Zero, -1, null, IntPtr.Zero, 1, 2, IntPtr.Zero, 0, IntPtr.Zero);
                if (errorCode != 0)
                {
                    throw new ExternalException("CoInitializeSecurity: " + GetSystemMessage(errorCode), errorCode);
                }
            }
        }
    }
}
