using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using opcprox;
using System.Runtime.InteropServices;

namespace OPCLibrary
{
    public enum OPCItemType { LEAF, BRANCH };
    public class OPCItem
    {
        private OPCItem parent;
        
        public OPCItem Parent
        {
            get { return parent; }
            set { parent = value; }
        }

        private string itemName;
        public string ItemName
        {
            get { return itemName; }
            set { itemName = value; }
        }

        private string description;
        public string Description
        {
            set { description = value; }
            get { return description; }
        }

        private OPCItemType itemType;
        public OPCItemType ItemType
        {
            get { return itemType; }
            set { itemType = value; }
        }

        private string itemID;
        public string ItemID
        {
            get { return itemID; }
            set { itemID = value; }
        }

        private bool enabled;
        public bool Enabled
        {
            set { enabled = value; }
            get { return enabled; }
        }

        private string dataValue;
        public string Value
        {
            get { return dataValue; }
            set { dataValue = value; }
        }

        public string TimeStamp
        { get; set; }

        private ushort wQuality;
        public string Quality
        {
            get { return OPCLibrary.Converter.GetQualityString(wQuality); }
            set { wQuality = (ushort)Int16.Parse(value); }
        }


        private uint m_hItem = 0;
        public uint ItemHandle
        {
            get { return m_hItem; }
            set { m_hItem = value; }
        }

        public OPCItem(OPCItem parent = null)
        { 
            Parent = parent; 
        }

        public tagOPCITEMDEF GetItemDef()
        {
            tagOPCITEMDEF itemDef = new tagOPCITEMDEF();
            itemDef.szItemID = ItemID;
            itemDef.szAccessPath = null;
            itemDef.bActive = Convert.ToInt32(Enabled);
            itemDef.hClient = 1;
            itemDef.vtRequestedDataType = (ushort)VarEnum.VT_EMPTY;
            itemDef.dwBlobSize = 0;
            itemDef.pBlob = IntPtr.Zero;
            return itemDef;
        }

        public override string ToString()
        {
            return ItemID;
        }

    }
}
