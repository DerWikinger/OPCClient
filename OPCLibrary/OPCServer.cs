using System;
using System.Collections.Generic;
using OpcEnumLib;
using opcprox;
using System.Runtime.InteropServices;
using System.Linq;


namespace OPCLibrary
{
    public class OPCServer
    {
        public OPCServer()
        {
            IsConnected = false;
            Enabled = true;
        }

        ~OPCServer()
        {
            Disconnect();
        }

        private object m_pOPCServer = null;
        public IOPCServer Server
        {
            get { return (IOPCServer)m_pOPCServer; }
            set { m_pOPCServer = value; }
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

        public bool IsConnected
        { get; set; }

        public List<OPCItem> GetItems(bool reinitizlize = false)
        {
            List<OPCItem> opcItems = new List<OPCItem>();
            IOPCBrowseServerAddressSpace pBrowse = (IOPCBrowseServerAddressSpace)Internal.Interop.CreateInstance(Guid, HostName);
            if(m_pOPCServer == null) m_pOPCServer = pBrowse;
            GetItemChildren(opcItems, "", pBrowse);
            return opcItems;
        }
        
        private void GetItemProperties(string szItemID)
        {
            IOPCItemProperties properties = (IOPCItemProperties)m_pOPCServer;
            //uint pdwPropertyIDs = 0;
            uint pdwCount = 0;
            IntPtr ppvData = IntPtr.Zero;
            IntPtr ppPropertyIDs = IntPtr.Zero;
            IntPtr ppszNewItemIDs = IntPtr.Zero;
            uint pdwPropertyIDs = 0;
            IntPtr ppDescriptions = IntPtr.Zero;
            IntPtr ppvtDataTypes = IntPtr.Zero;
            IntPtr ppErrors = IntPtr.Zero;
            //properties.LookupItemIDs(szItemID, 1, ref pdwPropertyIDs, ppszNewItemIDs, ppErrors);
            //properties.QueryAvailableProperties(szItemID, out pdwCount, out ppPropertyIDs, out ppDescriptions, out ppvtDataTypes);
            //properties.GetItemProperties(szItemID, 1, ref pdwPropertyIDs, out ppvData, out ppErrors);
        }

        private void GetItemChildren(List<OPCItem> items, string szNameFilter, IOPCBrowseServerAddressSpace pParent, OPCItem parentItem = null)
        {
            opcprox.IEnumString pEnum;
            uint cnt;
            string strName;
            string szItemID;
            //IOPCItemProperties properties = (IOPCItemProperties)DZ.Opc.Integration.Internal.Interop.CreateInstance(Guid, HostName);

            try
            {
                pParent.BrowseOPCItemIDs(tagOPCBROWSETYPE.OPC_LEAF, szNameFilter, (ushort)VarEnum.VT_EMPTY, 0, out pEnum);
                pEnum.RemoteNext(1, out strName, out cnt);
                int nLeavesCount = 0;
                while (cnt != 0)
                {
                    pParent.GetItemID(strName, out szItemID); // получает полный идентификатор тега
                    items.Add(
                        new OPCItem()
                        {
                            Parent = parentItem,
                            ItemName = strName,
                            Enabled = true,
                            ItemType = OPCItemType.LEAF,
                            ItemID = szItemID
                        }
                    );
                    pEnum.RemoteNext(1, out strName, out cnt);
                    nLeavesCount++;
                }
                pParent.BrowseOPCItemIDs(tagOPCBROWSETYPE.OPC_BRANCH, szNameFilter, (ushort)VarEnum.VT_EMPTY, 0, out pEnum);
                pEnum.RemoteNext(1, out strName, out cnt);
                int nBranchesCount = 0;
                while (cnt != 0)
                {
                    pParent.GetItemID(strName, out szItemID);
                    OPCItem item = new OPCItem()
                    {
                        Parent = parentItem,
                        ItemName = strName,
                        Enabled = true,
                        ItemType = OPCItemType.BRANCH,
                        ItemID = szItemID
                    };
                    items.Add(item);
                    pParent.ChangeBrowsePosition(tagOPCBROWSEDIRECTION.OPC_BROWSE_TO, szItemID);
                    GetItemChildren(items, "", pParent, item);
                    pParent.ChangeBrowsePosition(tagOPCBROWSEDIRECTION.OPC_BROWSE_UP, szItemID);
                    pEnum.RemoteNext(1, out strName, out cnt);
                    nBranchesCount++;
                }
            }
            catch (System.Exception ex)
            {
                Console.Out.WriteLine(ex.Message);
            }

        }

        public void Connect()
        {
            try
            {
                m_pOPCServer = (IOPCServer)Internal.Interop.CreateInstance(guid, hostname);
                IsConnected = true;
            }
            catch (System.Exception ex)
            {
                Console.Out.WriteLine(ex.Message);
            }
        }

        public void Disconnect()
        {
            if (m_pOPCServer != null)
            {
                Internal.Interop.ReleaseServer(m_pOPCServer);
            }
            m_pOPCServer = null;
            IsConnected = false;
        }

        public static List<OPCServer> BrowseServers(string hostname)
        {
            List<OPCServer> servers = new List<OPCServer>();

            Guid clsidcat = new Guid("{63D5F432-CFE4-11D1-B2C8-0060083BA1FB}"); //OPC Data Access Servers Version 2.0
            Guid clsidenum = new Guid("{13486D51-4821-11D2-A494-3CB306C10000}"); //OPCEnum.exe
            try
            {
                IOPCServerList2 serverList = (IOPCServerList2)Internal.Interop.CreateInstance(clsidenum, hostname);

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
                Internal.Interop.ReleaseServer(serverList);
            }
            catch (System.Exception ex)
            {
                Console.Out.WriteLine(ex.Message);
            }
            return servers;
        }

        public static OPCServer FindServerByProgID(string progID, string hostname)
        {

            foreach (OPCServer srv in OPCServer.BrowseServers(hostname))
            {
                if (srv.ProgId.ToUpper().Contains(progID.ToUpper())) return srv;
            }

            return null;
        }
    }

    public class ServerException : Exception
    {
        public ServerException(int hRes, string msg = "Ошибка OPC сервера") : base(msg)
        {
            base.HResult = hRes;
        }
    }
}
