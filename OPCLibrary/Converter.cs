using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using opcprox;

namespace OPCLibrary
{
    public static class Converter
    {

        public static string GetQualityString(ushort usQuality)
        {
            switch (usQuality)
            {
                case 0x00: return "Bad";
                case 0x04: return "Config Error";
                case 0x08: return "Not Connected";
                case 0x0C: return "Device Failure";
                case 0x10: return "Sensor Failure";
                case 0x14: return "Last Known";
                case 0x18: return "Comm Failure";
                case 0x1C: return "Out of Service";
                case 0x20: return "Initializing";
                case 0x40: return "Uncertain";
                case 0x44: return "Last Usable";
                case 0x50: return "Sensor Calibration";
                case 0x54: return "EGU Exceeded";
                case 0x58: return "Sub Normal";
                case 0xC0: return "Good";
                case 0xD8: return "Local Override";
                default: return "Unknown";
            }
        }

        public static string GetVTString(ushort vt)
        {
            return ((VarEnum)vt).ToString();
        }

        public static Type GetVTType(ushort vt)
        {
            //VarEnum.
            string t = GetVTString(vt);
            switch (t.ToString())
            {
                case "VT_ARRAY":
                    return new int[1].GetType();
                case "VT_BOOL":
                    return true.GetType();
                case "VT_BSTR":
                    return "".GetType();
                case "VT_DATE":
                    return DateTime.Now.GetType();
                case "VT_I1":
                    return ((char)1).GetType();
                case "VT_I2":
                    return ((short)1).GetType();
                case "VT_I4":
                    return ((int)1).GetType();
                case "VT_I8":
                    return ((long)1).GetType();
                case "VT_R4":
                    return ((float)1).GetType();
                case "VT_R8":
                    return ((double)1).GetType();
                case "VT_DECIMAL":
                    return ((decimal)1).GetType();
                case "VT_INT":
                    return ((int)1).GetType();
                case "VT_PTR":
                    return (IntPtr.Zero).GetType();
                default:
                    return new Object().GetType();
            }
        }

        public static int GetVTSize(ushort vt)
        {
            //VarEnum.
            //DateTime now = DateTime.Now;
            string t = GetVTString(vt);
            switch (t.ToString())
            {
                case "VT_BOOL":
                    return sizeof(bool);
                case "VT_I1":
                    return sizeof(char);
                case "VT_I2":
                    return sizeof(short);
                case "VT_I4":
                    return sizeof(int);
                case "VT_I8":
                    return sizeof(long);
                case "VT_R4":
                    return sizeof(float);
                case "VT_R8":
                    return sizeof(double);
                case "VT_DECIMAL":
                    return sizeof(decimal);
                case "VT_INT":
                    return sizeof(int);
                case "VT_PTR":
                    return sizeof(uint);
                case "VT_DATE":
                    return 4;
                default:
                    return -1;
            }
        }

        public static object GetPropertyValue(short type, object value)
        {
            string strVal = value.ToString();
            switch (type)
            {
                case 16: // char
                    return char.Parse(strVal);
                case 2: // short
                    return short.Parse(strVal);
                case 4: // float
                    {
                        IntPtr pValue = Marshal.AllocCoTaskMem(8);
                        Marshal.WriteInt64(pValue, long.Parse(strVal));
                        byte[] bytes = new byte[8];
                        Marshal.Copy(pValue, bytes, 0, 8);
                        float[] result = new float[1];
                        Marshal.Copy(pValue, result, 0, 1);
                        Marshal.FreeCoTaskMem(pValue);
                        return result[0];
                    }
                case 3: // int
                    return int.Parse(strVal);
                case 20: // long
                    return long.Parse(strVal);
                case 7: // DateTime
                    {
                        
/*                        long now = DateTime.Now.ToBinary();
                        IntPtr pNow = Marshal.AllocCoTaskMem(8);
                        Marshal.WriteInt64(pNow, now);
                        byte[] nByte = new byte[8];
                        Marshal.Copy(pNow, nByte, 0, 8);*/

                        IntPtr pValue = Marshal.AllocCoTaskMem(8);
                        Marshal.WriteInt64(pValue, long.Parse(value.ToString()));
                        byte[] bytes = new byte[8];
                        Marshal.Copy(pValue, bytes, 0, 8);
                        byte[] res = new byte[8];
                        res[0] = bytes[3];
                        res[1] = bytes[2];
                        res[2] = bytes[1];
                        res[3] = bytes[0];
                        res[4] = bytes[7];
                        res[5] = bytes[6];
                        res[6] = bytes[5];
                        res[7] = bytes[4];
                        IntPtr pRes = Marshal.AllocCoTaskMem(8);
                        Marshal.Copy(res, 0, pRes, 8);
                        long l = Marshal.ReadInt64(pRes);
                        Marshal.FreeCoTaskMem(pValue);
                        Marshal.FreeCoTaskMem(pRes);
                        return DateTime.FromBinary(l);
                    }
                    
                default:
                    return 0;
            }
        }

        public static string GetFTSting(_FILETIME ft)
        {
            long lFT = (((long)ft.dwHighDateTime) << 32) + ft.dwLowDateTime;
            try
            {
                DateTime dt = System.DateTime.FromFileTime(lFT);
                return dt.ToString();
            }
            catch (ApplicationException ex)
            {
                return ex.Message;
            }

        }
    }
}
