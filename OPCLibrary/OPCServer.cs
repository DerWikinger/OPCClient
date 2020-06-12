using System;
using System.Collections.Generic;
using OpcEnumLib;
using opcprox;

namespace OPCLibrary
{
    public class OPCServer
    {
        private IOPCServer m_pOPCServer = null;

        public OPCServer()
        {

        }

        ~OPCServer()
        {
            Disconnect();
        }

        private string serverName;
        public string ServerName
        {
            get { return serverName; }
            set { serverName = value; }
        }

        private string progId;
        public string ProgId
        {
            get { return progId; }
            set { progId = value; }
        }

        private string serverDescription;
        public string ServerDescription
        {
            get { return serverDescription; }
            set { serverDescription = value; }
        }

        private Guid guid;
        public Guid Guid
        {
            get { return guid; }
            set { guid = value; }
        }

        private string hostname;
        public string HostName
        {
            get { return hostname; }
            set { hostname = value; }
        }

        private bool enabled;
        public bool Enabled
        {
            get { return enabled; }
            set { enabled = value; }
        }

        public void Connect()
        {
            m_pOPCServer = (IOPCServer)DZ.Opc.Integration.Internal.Interop.CreateInstance(guid, hostname);
        }

        public void Disconnect()
        {
            if (m_pOPCServer != null)
            {
                
                DZ.Opc.Integration.Internal.Interop.ReleaseServer(m_pOPCServer);
            }
            m_pOPCServer = null;
        }

        public static List<OPCServer> BrowseServers(string hostname)
        {
            List<OPCServer> servers = new List<OPCServer>();
            Guid clsidcat = new Guid("{63D5F432-CFE4-11D1-B2C8-0060083BA1FB}");
            Guid clsidenum = new Guid("{13486D51-4821-11D2-A494-3CB306C10000}");
            try
            {
                IOPCServerList2 serverList = (IOPCServerList2)DZ.Opc.Integration.Internal.Interop.CreateInstance(clsidenum, hostname);

                IOPCEnumGUID pIOPCEnumGuid;

                serverList.EnumClassesOfCategories(1, ref clsidcat, 0, ref clsidcat, out pIOPCEnumGuid);

                string pszProgID;
                string pszUserType;
                string pszVerIndProgID;
                Guid guid = new Guid();
                int nServerCnt = 0;
                uint iRetSvr;

                pIOPCEnumGuid.Next(1, out guid, out iRetSvr);
                while (iRetSvr != 0)
                {
                    nServerCnt++;
                    serverList.GetClassDetails(ref guid, out pszProgID,
                         out pszUserType, out pszVerIndProgID);
                    OPCServer server = new OPCServer()
                    {
                        Guid = guid,
                        ProgId = pszProgID,
                        HostName = hostname,
                    };
                    servers.Add(server);
                    pIOPCEnumGuid.Next(1, out guid, out iRetSvr);
                }
                DZ.Opc.Integration.Internal.Interop.ReleaseServer(serverList);
            }
            catch (System.Exception ex)
            {
                Console.Out.Write(ex.Message);
            }
            return servers;
        }
    }
}
