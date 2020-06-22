using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using opcprox;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace OPCLibrary
{
    public class DataCallback : IOPCDataCallback
    {
        private string m_szItemID;
        private OPCItem m_item;
        
        public DataCallback(OPCItem item)
        {
            m_item = item;
            SetItemID(item.ItemID);
        }

        public void SetItemID(string szItemID)
        {
            m_szItemID = szItemID;
        }

        public void OnDataChange(uint dwTransid, uint hGroup, int hrMasterquality, int hrMastererror, uint dwCount, 
            ref uint phClientItems, ref object pvValues, ref ushort pwQualities, ref _FILETIME pftTimeStamps, ref int pErrors)
        {
            if(dwTransid == 0) return;
            int count = (int)dwCount;
            IntPtr iptrValues = Marshal.AllocCoTaskMem(count * 2);
            Marshal.GetNativeVariantForObject(pvValues, iptrValues);
            byte[] vt = new byte[count * 2];
            Marshal.Copy(iptrValues, vt, 0, count * 2);
            Marshal.FreeCoTaskMem(iptrValues);
            ushort[] usVt = new ushort[count];
            for(int i = 0; i < count; i++)
            {
                usVt[i] = (ushort)(vt[i * 2] + vt[i*2 + 1] * 255);
            }

            m_item.DataCallback.Invoke(usVt);
        }

        public void OnReadComplete(uint dwTransid, uint hGroup, int hrMasterquality, int hrMastererror, uint dwCount, 
            ref uint phClientItems, ref object pvValues, ref ushort pwQualities, ref _FILETIME pftTimeStamps, ref int pErrors)
        {
            int count = (int)dwCount;
            tagOPCITEMVQT[] tags = new tagOPCITEMVQT[count];
            
            
            IntPtr iptrValues = Marshal.AllocCoTaskMem(count * 2);
            Marshal.GetNativeVariantForObject(pvValues, iptrValues);
            byte[] vt = new byte[count * 2];
            Marshal.Copy(iptrValues, vt, 0, count * 2);
            Marshal.FreeCoTaskMem(iptrValues);

            IntPtr iptrClientItems = Marshal.AllocCoTaskMem(count * sizeof(uint));
            Marshal.GetNativeVariantForObject(phClientItems, iptrClientItems);
            byte[] ct = new byte[count * sizeof(uint)];
            Marshal.Copy(iptrClientItems,ct, 0, count * sizeof(uint));
            Marshal.FreeCoTaskMem(iptrClientItems);

            ushort[] usVt = new ushort[count];
            for (int i = 0; i < count; i++)
            {
                usVt[i] = (ushort)(vt[i * 2] + vt[i * 2 + 1] * 255);
            }

            m_item.DataCallback.Invoke(usVt);
        }

        public void OnWriteComplete(uint dwTransid, uint hGroup, int hrMastererr, uint dwCount, ref uint pClienthandles, ref int pErrors)
        {
            throw new NotImplementedException();
        }

        void IOPCDataCallback.OnCancelComplete(uint dwTransid, uint hGroup)
        {
            throw new NotImplementedException();
        }
    }
}
