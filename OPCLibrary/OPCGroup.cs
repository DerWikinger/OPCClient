using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpcEnumLib;
using opcprox;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace OPCLibrary
{
    public class OPCGroup
    {
        private IEnumerable<OPCItem> items;
        public IEnumerable<OPCItem> Items
        { 
            get { return items; }
            set { items = value; }
        }

        public string Name
        { get; set; }

        public bool Enabled
        { get; set; }

        uint fUpdateRate;
        public uint UpdateRate
        {
            get { return fUpdateRate; }
            set { fUpdateRate = value; }
        }

        float fDeadband = 0;
        public float Deadband
        {
            get { return fDeadband; }
            set { fDeadband = value; }
        }

        public uint GroupID
        { get; private set; }

        //private OPCServer server = null;
        private static uint m_hGroup = 0;
        private uint[] m_hItems = null;
        int[] m_hRes;
        private int m_dwReadCount = 0;
        List<uint> indecies = new List<uint>();
        IOPCItemMgt m_pItemMgt = null;
        //IOPCSyncIO m_pSyncIO = null;
        
        public IConnectionPoint m_pDataCallback; // Точка подключения к событиям сервера
        int m_dwCookie; // Описатель подписки к событиям сервера

        public OPCGroup(string name, IEnumerable<OPCItem> items)
        {
            this.items = items;
            Name = name;
            Enabled = true;
            GroupID = ++m_hGroup;
        }

        ~OPCGroup()
        {
            m_hGroup--;
        }

        public void RegInServer(OPCServer server)
        {
            //this.server = server;

            int TimeBias = 0;
            Guid riid = typeof(IOPCItemMgt).GUID;

            uint hClientGroup = 1;
            uint dwLCID = 2;
            object pUnknown;
            try
            {
                server.Server.AddGroup(this.Name, System.Convert.ToInt32(this.Enabled), this.UpdateRate, hClientGroup,
                    ref TimeBias, ref this.fDeadband, dwLCID, out m_hGroup, out fUpdateRate, riid, out pUnknown);
                m_pItemMgt = (IOPCItemMgt)pUnknown;
            }
            catch (System.Exception ex)
            {
                Console.Out.WriteLine(ex.Message);
            }

            try
            {
                RegItemsInServer();
            }
            catch(ServerException ex)
            {
                string msg;
                //Запрашиваем у сервера текст ошибки, соответствующий текущему HRESULT 
                server.Server.GetErrorString(ex.HResult, 2, out msg);
                //Показываем сообщение ошибки
                Console.Out.WriteLine(msg, "Ошибка");
            }            
        }

        public void AdviseOnServer(OPCServer server, OPCItem item)
        {
            if (m_pItemMgt == null) return;
            try
            {
                IConnectionPointContainer pCPC;
                pCPC = (IConnectionPointContainer)m_pItemMgt;
                Guid riid = typeof(IOPCDataCallback).GUID;
                //IEnumConnectionPoints ppEnum;
                //pCPC.EnumConnectionPoints(out ppEnum);
                pCPC.FindConnectionPoint(ref riid, out m_pDataCallback);

                //Если ранее была активирована подписка, то отменить ее
                if (m_dwCookie != 0) m_pDataCallback.Unadvise(m_dwCookie);

                DataCallback m_pSink = new DataCallback(item);
                m_pDataCallback.Advise(m_pSink, out m_dwCookie);
            }
            catch(COMException ex)
            {
                string msg;
                //Запрашиваем у сервера текст ошибки, соответствующий текущему HRESULT 
                server.Server.GetErrorString(ex.HResult, 2, out msg);
                //Показываем сообщение ошибки
                Console.Out.WriteLine(msg, "Ошибка");
            }
            catch (ServerException ex)
            {
                string msg;
                //Запрашиваем у сервера текст ошибки, соответствующий текущему HRESULT 
                server.Server.GetErrorString(ex.HResult, 2, out msg);
                //Показываем сообщение ошибки
                Console.Out.WriteLine(msg, "Ошибка");
            }
        }

        private void RegItemsInServer()
        {
            int dwCount = Items.Count();

            Type type = typeof(tagOPCITEMDEF);
            IntPtr pItems = Marshal.AllocCoTaskMem((int)dwCount * Marshal.SizeOf(type));
            m_dwReadCount = 0;
            indecies.Clear();
            tagOPCITEMDEF[] itemDefs = new tagOPCITEMDEF[dwCount];
            for (int i = 0; i < dwCount; i++)
            {
                OPCItem item = ((List<OPCItem>)Items)[i];
                if (item.Enabled)
                {
                    itemDefs[m_dwReadCount] = item.GetItemDef();
                    Marshal.StructureToPtr(itemDefs[m_dwReadCount], pItems + m_dwReadCount * Marshal.SizeOf(type), false);
                    indecies.Add((uint)m_dwReadCount);
                    m_dwReadCount++;
                }
            }
            IntPtr iptrErrors = IntPtr.Zero;
            IntPtr iptrResults = IntPtr.Zero;

            try
            {
                // Добавляем элемент данных в группу
                m_pItemMgt.AddItems((uint)m_dwReadCount, pItems, out iptrResults, out iptrErrors);
            }
            catch (ApplicationException ex)
            {
                Console.Out.WriteLine(ex.Message);
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
                Console.Out.WriteLine(ex.Message);
                string msg;
                //Получаем HRESULT соответствующий сгененрированному исключению
                int hRes = Marshal.GetHRForException(ex);
                throw new ServerException(hRes);
            }

            Marshal.FreeCoTaskMem(pItems);
            Marshal.FreeCoTaskMem(iptrErrors);
            Marshal.FreeCoTaskMem(iptrResults);
        }

        public void RemoveFromServer()
        {
            IntPtr iptrErrors = IntPtr.Zero;
            if (m_hItems == null) return;
            for (int index = 0; index < m_hItems.Length; index++)
            {
                if (m_hItems[index] != 0)
                {
                    this.m_pItemMgt.RemoveItems(1, m_hItems[index], out iptrErrors);
                    Marshal.FreeCoTaskMem(iptrErrors);
                }
                m_hItems[index] = 0;
            }
        }

        public int SyncRead()
        {
            //if (!this.server.IsConnected) this.server.Connect();
            //if(m_pItemMgt != null) RemoveFromServer();

            IOPCSyncIO m_pSyncIO = null;

            int iRet = 0;
            try
            {
                m_pSyncIO = (IOPCSyncIO)m_pItemMgt;
                //При наличии подписки событий группы на сервере удаляем её
                if (m_dwCookie != 0)
                {
                    m_pDataCallback.Unadvise(m_dwCookie);
                    m_dwCookie = 0;
                }
            }
            catch (ApplicationException ex)
            {
                Console.Out.WriteLine(ex.Message);
            }

            tagOPCDATASOURCE ds = new tagOPCDATASOURCE();
            ds = tagOPCDATASOURCE.OPC_DS_DEVICE;
            IntPtr iptrItemValue = IntPtr.Zero;
            IntPtr iptrErrors = IntPtr.Zero;
            try
            {
                m_pSyncIO.Read(ds, (uint)m_dwReadCount, ref m_hItems[0], out iptrItemValue, out iptrErrors);
            }
            catch (System.Exception ex)
            {
                iRet = -1;
                Console.Out.WriteLine(ex.Message);
            }
            try
            {
                OPCItem[] items = Items.ToArray();
                tagOPCITEMSTATE[] itemSates = new tagOPCITEMSTATE[m_dwReadCount];
                Type type = typeof(tagOPCITEMSTATE);
                Marshal.Copy(iptrErrors, m_hRes, 0, m_dwReadCount);

                for (int i = 0; i < items.Length; i++)
                {
                    OPCItem item = items[i];
                    itemSates[i] = (tagOPCITEMSTATE)Marshal.PtrToStructure(iptrItemValue + i * Marshal.SizeOf(type), type);
                    uint index = indecies[i];
                    item.Value = itemSates[i].vDataValue != null ? itemSates[i].vDataValue.ToString() : null;
                    string tst = OPCLibrary.Converter.GetFTSting(itemSates[i].ftTimeStamp);
                    item.TimeStamp = tst;
                    item.Quality = itemSates[i].wQuality.ToString();

                    //Генерируем исключение в случае ошибки в HRESULT
                    Marshal.ThrowExceptionForHR(m_hRes[i]);
                }

            }
            catch (System.Exception ex)
            {
                Console.Out.WriteLine(ex.Message);
                string msg;
                //Получаем HRESULT соответствующий сгененрированному исключению
                int hRes = Marshal.GetHRForException(ex);
                throw new ServerException(hRes);
            }
            finally
            {
                Marshal.FreeCoTaskMem(iptrItemValue);
                Marshal.FreeCoTaskMem(iptrErrors);
            }
            return iRet;
        }

        public int ASyncRead()
        {
            IOPCAsyncIO2 m_pASyncIO = null;

            int iRet = 0;
            try
            {
                //IntPtr iptrItemValue = IntPtr.Zero;
                IntPtr iptrErrors = IntPtr.Zero;
                uint dwTransactionID = GroupID;
                uint pdwCancelID;
                m_pASyncIO = (IOPCAsyncIO2)m_pItemMgt;
                m_pASyncIO.Read((uint)m_dwReadCount, ref m_hItems[0], dwTransactionID, out pdwCancelID, out iptrErrors);
                Marshal.Copy(iptrErrors, m_hRes, 0, m_dwReadCount);
            }
            catch (ApplicationException ex)
            {
                Console.Out.WriteLine(ex.Message);
            }

            return iRet;
        }
    }
}
