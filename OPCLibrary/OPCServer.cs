using System;
using System.Collections.Generic;
using OpcEnumLib;
using opcprox;
using System.Runtime.InteropServices;


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

        public List<OPCItem> GetItems()
        {
            List<OPCItem> items = new List<OPCItem>();
            Guid guid = new Guid("{39227004-A18F-4b57-8B0A-5235670F4468}");
            IOPCBrowseServerAddressSpace pBrowse = (IOPCBrowseServerAddressSpace)DZ.Opc.Integration.Internal.Interop.CreateInstance(Guid, HostName);
            GetItemChildren(items, "", pBrowse);
            return items;
        }

        private void GetItemChildren(List<OPCItem> items, string szNameFilter, IOPCBrowseServerAddressSpace pParent, OPCItem parentItem = null)
        {
            opcprox.IEnumString pEnum;
            uint cnt;
            string strName;
            string szItemID;
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
                Console.Out.Write(ex.Message);
            }

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
 
            Guid clsidcat = new Guid("{63D5F432-CFE4-11D1-B2C8-0060083BA1FB}"); //OPC Data Access Servers Version 2.0
            Guid clsidenum = new Guid("{13486D51-4821-11D2-A494-3CB306C10000}"); //OPCEnum.exe
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


        public static OPCServer FindServerByProgID(string progID, string hostname)
        {

            foreach(OPCServer srv in OPCServer.BrowseServers(hostname))
            {
                if (srv.ProgId.ToUpper().Contains(progID.ToUpper())) return srv;
            }

            return null;
        }
    }
}
