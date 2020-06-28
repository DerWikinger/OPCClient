using System;
using System.Collections.Generic;
using OpcEnumLib;
using System.Linq;
using opcprox;
using System.Runtime.InteropServices;
using System.Linq;
using System.Reflection;
using System.Net;

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

        public List<PropertyData> GetItemProperties(string szItemID)
        {
            IOPCItemProperties properties = (IOPCItemProperties)m_pOPCServer;
            uint pdwCount = 0;
            IntPtr ppvData = IntPtr.Zero;
            IntPtr ppPropertyIDs = IntPtr.Zero;
            IntPtr ppDescriptions = IntPtr.Zero;
            IntPtr ppvtDataTypes = IntPtr.Zero;
            IntPtr ppErrors = IntPtr.Zero;

            properties.QueryAvailableProperties(szItemID, out pdwCount, out ppPropertyIDs, out ppDescriptions, out ppvtDataTypes);

            int[] propertiesIDs = new int[pdwCount];
            Marshal.Copy(ppPropertyIDs, propertiesIDs, 0, (int)pdwCount);
            IntPtr pdwPropertyIDs = Marshal.AllocCoTaskMem((int)pdwCount * sizeof(int));
            Marshal.Copy(propertiesIDs, 0, pdwPropertyIDs, (int)pdwCount);

            IntPtr[] pDescriptions = new IntPtr[pdwCount];
            Marshal.Copy(ppDescriptions, pDescriptions, 0, (int)pdwCount);
            string[] descriptions = new string[pdwCount];
            for (int i = 0; i < pdwCount; i++)
            {
                string description;
                int length = 0;
                description = Marshal.PtrToStringAuto(pDescriptions[i]);
                length += description.Length;
                descriptions[i] = description;
                Marshal.FreeCoTaskMem(pDescriptions[i]);
            }

            short[] propertyTypes = new short[pdwCount];
            Marshal.Copy(ppvtDataTypes, propertyTypes, 0, (int)pdwCount);

            var strPropertyTypes = from type in propertyTypes
                                   let sType = Converter.GetVTString((ushort)type)
                                   let typeSize = Converter.GetVTSize((ushort)type)
                                   select sType + " : " + typeSize;
            string[] list = strPropertyTypes.ToArray();

            properties.GetItemProperties(szItemID, pdwCount, pdwPropertyIDs, out ppvData, out ppErrors);

            List<PropertyData> pValues = new List<PropertyData>();
            int start = 0;
            short vt;
            for (int i = 0; i < pdwCount; i++)
            {
                int propertyID = propertiesIDs[i];
                ulong value = (ulong)Marshal.ReadInt64(ppvData + start * sizeof(ulong));
                int sz = Converter.GetVTSize((ushort)value);
                vt = propertyTypes[i];
                if (vt == 8)
                {
                    start++;
                    IntPtr[] strPtr = new IntPtr[1];
                    Marshal.Copy(ppvData + start * sizeof(long), strPtr, 0, 1);
                    string val = Marshal.PtrToStringUni(strPtr[0]);
                    pValues.Add(new PropertyData() { ID = propertyID, Name = descriptions[i], Value = val });
                    start++;
                }
                else if (vt == 8197)
                {
                    start++;
                    IntPtr[] strPtr = new IntPtr[1];
                    Marshal.Copy(ppvData + start * sizeof(long), strPtr, 0, 1);
                    string val = Marshal.PtrToStringUni(strPtr[0]);
                    pValues.Add(new PropertyData() { ID = propertyID, Name = descriptions[i], Value = val });
                    start++;
                } else if(vt == 4) {
                    start++;
                    value = (ulong)Marshal.ReadInt64(ppvData + start * sizeof(ulong));
                    object obj = Converter.GetPropertyValue(vt, value);
                    pValues.Add(new PropertyData() { ID = propertyID, Name = descriptions[i], Value = obj });
                    start++;
                } else if(vt == 7) {
                    // Пока не разобрался с конвертированием данных в формат даты
                    // поэтому присваивается текущее значение даты и времени
                    start++;
                    IntPtr pDate = ppvData + start * sizeof(ulong);
                    //IntPtr pp = Marshal.AllocCoTaskMem(8);
                    //byte[] bb = new byte[8];
                    //Marshal.GetNativeVariantForObject(pDate, pp);
                    //Marshal.Copy(pp, bb, 0, 8);

                    Type type = typeof(DateTime);

                    long lValue = Marshal.ReadInt64(pDate);

                    uint lt = (uint)Marshal.ReadInt32(pDate, 4);
                    uint ht = (uint)Marshal.ReadInt32(pDate, 0);

                    lValue = ((long)ht << 32) + lt;
                    //DateTime ft = (DateTime)Marshal.PtrToStructure(pDate, type);
                    
                    _FILETIME ft = new _FILETIME() { dwHighDateTime = ht, dwLowDateTime = lt };
                    //DateTime dt = DateTime.FromBinary(lValue);
                    //string s = Converter.GetFTSting(ft);
                    //DateTime dt = new DateTime(lValue);
                    string strDate = DateTime.Now.ToString();
                    pValues.Add(new PropertyData() { ID = propertyID, Name = descriptions[i], Value = strDate });
                    //Marshal.FreeCoTaskMem(pDate);
                    start++;
                } else {
                    start++;
                    value = (ulong)Marshal.ReadInt64(ppvData + start * sizeof(ulong));
                    pValues.Add(new PropertyData() { ID = propertyID, Name = descriptions[i], Value = value });
                    start++;
                }
            }

            Marshal.FreeCoTaskMem(ppDescriptions);
            Marshal.FreeCoTaskMem(ppPropertyIDs);
            Marshal.FreeCoTaskMem(ppvtDataTypes);
            Marshal.FreeCoTaskMem(ppvData);
            Marshal.FreeCoTaskMem(ppErrors);

            return pValues;
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
                    GetItemProperties(szItemID);
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

        public static List<OPCServer> BrowseServers(string hostname, NetworkCredential networkCredential = null, bool DAVersion3 = false)
        {
            List<OPCServer> servers = new List<OPCServer>();

            Guid clsidcat;

            if(DAVersion3) {
                clsidcat = new Guid("{CC603642-66D7-48F1-B69A-B625E73652D7}"); //OPC Data Access Servers Version 3.
            } else {
                clsidcat = new Guid("{63D5F432-CFE4-11D1-B2C8-0060083BA1FB}"); //OPC Data Access Servers Version 2.
            } 

            Guid clsidenum = new Guid("{13486D51-4821-11D2-A494-3CB306C10000}"); //OPCEnum.exe
            try
            {
                IOPCServerList2 serverList = (IOPCServerList2)Internal.Interop.CreateInstance(clsidenum, hostname, networkCredential);

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

    public struct PropertyData
    {
        public int ID;
        public string Name;
        public object Value;
        public string ToString()
        {
            return ID + " : " + Name + " = " + Value.ToString();
        }
    }
}
