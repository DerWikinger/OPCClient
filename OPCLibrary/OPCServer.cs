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

        }

        ~OPCServer()
        {
            Disconnect();
        }

        private IOPCServer m_pOPCServer = null;
        private uint m_hGroup = 0;
        private uint[] m_hItems = null;
        int[] m_hRes;
        private int m_dwReadCount = 0;
        List<uint> indecies = new List<uint>();
/*        List<OPCItem> opcItems = new List<OPCItem>();*/
        IOPCItemMgt m_pItemMgt = null;
        IOPCSyncIO m_pSyncIO = null;

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

        private void RegInServer(string groupName)
        {
            int TimeBias = 0;
            Guid riid = typeof(IOPCItemMgt).GUID;
            uint fUpdateRate = 1000;
            float fDeadband = 0;
            uint hClientGroup = 1;
            uint dwLCID = 2;
            object pUnknown;
            try
            {
                m_pOPCServer.AddGroup(groupName, System.Convert.ToInt32(Enabled), fUpdateRate, hClientGroup, 
                    ref TimeBias, ref fDeadband, dwLCID, out m_hGroup, out fUpdateRate, riid, out pUnknown);
                m_pItemMgt = (IOPCItemMgt)pUnknown;
            }
            catch (System.Exception ex)
            {
                Console.Out.Write(ex.Message);
            }
        }

        public void RegItemsInServer(List<OPCItem> opcItems)
        {
            RegInServer("My_Group");

            int dwCount = opcItems.Count;

            Type type = typeof(tagOPCITEMDEF);
            IntPtr pItems = Marshal.AllocCoTaskMem((int)dwCount * Marshal.SizeOf(type));
            m_dwReadCount = 0;
            indecies.Clear();
            tagOPCITEMDEF[] itemDefs = new tagOPCITEMDEF[dwCount];
            for (int i = 0; i < dwCount; i++)
            {
/*                if (opcItems[i].ItemType == OPCItemType.LEAF && opcItems[i].Enabled)*/
                if (opcItems[i].Enabled)
                {
                    itemDefs[m_dwReadCount] = opcItems[i].GetItemDef();
                    Marshal.StructureToPtr(itemDefs[m_dwReadCount], pItems + m_dwReadCount * Marshal.SizeOf(type), false);
                    indecies.Add((uint)m_dwReadCount);
                    m_dwReadCount++;
                }
            }
            IntPtr iptrErrors = IntPtr.Zero;
            IntPtr iptrResults = IntPtr.Zero;
            IntPtr ppValidationResults = IntPtr.Zero;

            try
            {
                // Добавляем элемент данных в группу
                m_pItemMgt.AddItems((uint)m_dwReadCount, pItems, out iptrResults, out iptrErrors);
            }
            catch (ApplicationException ex)
            {
                Console.Out.Write(ex.Message);
            }
            m_hItems = new uint[m_dwReadCount];
            tagOPCITEMRESULT[] pResults = new tagOPCITEMRESULT[m_dwReadCount];
            m_hRes = new int[dwCount];
            Marshal.Copy(iptrErrors, m_hRes, 0, dwCount);
            try
            {
                for (int i = 0; i < m_dwReadCount; i++)
                {
                    pResults[i] = (tagOPCITEMRESULT)Marshal.PtrToStructure(iptrResults +
                        i * Marshal.SizeOf(typeof(tagOPCITEMRESULT)), typeof(tagOPCITEMRESULT));
                    m_hItems[i] = pResults[i].hServer;
                    //Генерируем исключение в случае ошибки в HRESULT
                    Marshal.ThrowExceptionForHR(m_hRes[i]);
                }
            } 
            catch (System.Exception ex)
            {
                Console.Out.Write(ex.Message);
                string msg;
                //Получаем HRESULT соответствующий сгененрированному исключению
                int hRes = Marshal.GetHRForException(ex);
                //Запрашиваем у сервера текст ошибки, соответствующий текущему HRESULT 
                m_pOPCServer.GetErrorString(hRes, 2, out msg);
                //Показываем сообщение ошибки
                Console.Out.Write(msg, "Ошибка");
            }

            Marshal.FreeCoTaskMem(pItems);
            Marshal.FreeCoTaskMem(iptrErrors);
            Marshal.FreeCoTaskMem(iptrResults);
        }

/*        public void RemoveItemsFromServer(List<OPCItem> opcItems)
        {
            foreach (OPCItem opcItem in opcItems)
            {
                opcItem.RemoveFromServer(m_pItemMgt);
            }
        }*/

        public int SyncRead(List<OPCItem> opcItems)
        {
            /*unsafe IOPCEnumGUID* pIOPCEnumGuid;*/

            RegItemsInServer(opcItems);

            int iRet = 0;
            try
            {
                if (null == m_pSyncIO)
                {
                    m_pSyncIO = (IOPCSyncIO)m_pItemMgt;
                }
            }
            catch (ApplicationException ex)
            {
                Console.Out.Write(ex.Message);
            }

            tagOPCDATASOURCE ds = new tagOPCDATASOURCE();
            ds = tagOPCDATASOURCE.OPC_DS_DEVICE;
/*            if (OPCDataSource.DSCache == DataSource)
                ds = tagOPCDATASOURCE.OPC_DS_CACHE;
            else
                ds = tagOPCDATASOURCE.OPC_DS_DEVICE;*/
            IntPtr iptrItemValue = IntPtr.Zero;
            IntPtr iptrErrors = IntPtr.Zero;
            try
            {
                m_pSyncIO.Read(ds, (uint)m_dwReadCount, ref m_hItems[0], out iptrItemValue, out iptrErrors);
            }
            catch (System.Exception ex)
            {
                iRet = -1;
                Console.Out.Write(ex.Message);
            }
            try
            {
                tagOPCITEMSTATE[] itemSates = new tagOPCITEMSTATE[m_dwReadCount];
                Type type = typeof(tagOPCITEMSTATE);
                Marshal.Copy(iptrErrors, m_hRes, 0, 1);

                for (int i = 0; i < m_dwReadCount; i++)
                {
                    itemSates[i] = (tagOPCITEMSTATE)Marshal.PtrToStructure(iptrItemValue + i * Marshal.SizeOf(type), type);
                    uint index = indecies[i];
                    opcItems[(int)index].Value = itemSates[i].vDataValue.ToString();
                    string tst = OPCLibrary.Converter.GetFTSting(itemSates[i].ftTimeStamp);
                    opcItems[(int)index].TimeStamp = tst;                                     
                    opcItems[(int)index].Quality = itemSates[i].wQuality.ToString();

                    //Генерируем исключение в случае ошибки в HRESULT
                    Marshal.ThrowExceptionForHR(m_hRes[i]);
                }

            }
            catch (System.Exception ex)
            {
                Console.Out.Write(ex.Message);
                string msg;
                //Получаем HRESULT соответствующий сгененрированному исключению
                int hRes = Marshal.GetHRForException(ex);
                //Запрашиваем у сервера текст ошибки, соответствующий текущему HRESULT 
                m_pOPCServer.GetErrorString(hRes, 2, out msg);
                //Показываем сообщение ошибки
                Console.Out.Write(msg, "Ошибка");
            }
            finally
            {
                Marshal.FreeCoTaskMem(iptrItemValue);
                Marshal.FreeCoTaskMem(iptrErrors);
            }
            return iRet;
        }

        public List<OPCItem> GetItems(bool reinitizlize = false)
        {
            List<OPCItem> opcItems = new List<OPCItem>();
/*            if (!reinitizlize) return opcItems;*/
/*            opcItems.Clear();*/
/*            Guid guid = new Guid("{39227004-A18F-4b57-8B0A-5235670F4468}");*/
            IOPCBrowseServerAddressSpace pBrowse = (IOPCBrowseServerAddressSpace)DZ.Opc.Integration.Internal.Interop.CreateInstance(Guid, HostName);
            GetItemChildren(opcItems, "", pBrowse);
/*            var result = from item in items
                         where item.ItemType == OPCItemType.LEAF
                         select item;*/
            return opcItems;
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
