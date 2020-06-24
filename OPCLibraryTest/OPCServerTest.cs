using System;
using System.Threading;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OPCLibrary;
using OPCLibrary.Internal;
using System.Runtime.InteropServices;
using System.Linq;
using opcprox;



namespace OPCLibraryTest
{
    [TestClass]
    public class OPCServerTest
    {
        [TestMethod]
        public void TestOPCServer()
        {
            OPCServer server = new OPCServer();
            Assert.AreNotEqual(null, server);
            server.ServerName = "Name";
            Assert.AreEqual("Name", server.ServerName);
            server.ServerDescription = "Description";
            Assert.AreEqual("Description", server.ServerDescription);
            Guid g = new Guid("00000000-0000-0000-0000-000000000000");
            server.Guid = g;
            Assert.AreEqual(g, server.Guid);
            server.HostName = "localhost";
            Assert.AreEqual("localhost", server.HostName);
            server.ProgId = "ProgID";
            Assert.AreEqual("ProgID", server.ProgId);

            string hostname = "127.0.0.1";
            List<OPCServer> servers = OPCServer.BrowseServers(hostname);
            Assert.AreNotEqual(null, servers);
            Assert.AreNotEqual(0, servers.Count);

            string progID = "Insat";
            server = OPCServer.FindServerByProgID(progID, hostname);
            Assert.AreNotEqual(null, server);

            server.Connect();
            List<OPCItem> arr = server.GetItems(true);
            string itemID = "PN_SIMULATOR.PD_SIMULATOR.Sin";
            var item = from elem in arr
                       where elem.ItemID == itemID
                       select elem;
            List < PropertyData > list = server.GetItemProperties(item.ToArray()[0].ItemID);
            Assert.AreEqual(1, list[0].ID);
            Assert.AreEqual("Item Canonical DataType", list[0].Name);
            Assert.AreEqual("4", list[0].Value.ToString());
            Assert.AreNotEqual(null, arr);
            Assert.AreEqual(27, arr.Count); // 27 items are defiened in OPC Server config
            server.Disconnect();
        }

        [TestMethod]
        public void TestInterop()
        {
            string message = OPCLibrary.Internal.Interop.GetSystemMessage(0);
            Assert.AreNotEqual(null, message);
            Assert.AreNotEqual("", message);
        }

        [TestMethod]
        public void TestServerInfo()
        {
            ServerInfo sInfo = new ServerInfo();
            Assert.AreNotEqual(null, sInfo);
            COSERVERINFO info = sInfo.Allocate("localhost", new System.Net.NetworkCredential());
            Assert.AreNotEqual(null, info); 
        }

        [TestMethod]
        public void TestConverter()
        {
            Assert.AreEqual("Bad", OPCLibrary.Converter.GetQualityString(0x00));
            Assert.AreEqual("Config Error", OPCLibrary.Converter.GetQualityString(0x04));
            Assert.AreEqual("Not Connected", OPCLibrary.Converter.GetQualityString(0x08));
            Assert.AreEqual("Device Failure", OPCLibrary.Converter.GetQualityString(0x0C));
            Assert.AreEqual("Sensor Failure", OPCLibrary.Converter.GetQualityString(0x10));
            Assert.AreEqual("Last Known", OPCLibrary.Converter.GetQualityString(0x14));
            Assert.AreEqual("Comm Failure", OPCLibrary.Converter.GetQualityString(0x18));
            Assert.AreEqual("Out of Service", OPCLibrary.Converter.GetQualityString(0x1C));
            Assert.AreEqual("Initializing", OPCLibrary.Converter.GetQualityString(0x20));
            Assert.AreEqual("Uncertain", OPCLibrary.Converter.GetQualityString(0x40));
            Assert.AreEqual("Last Usable", OPCLibrary.Converter.GetQualityString(0x44));
            Assert.AreEqual("Sensor Calibration", OPCLibrary.Converter.GetQualityString(0x50));
            Assert.AreEqual("EGU Exceeded", OPCLibrary.Converter.GetQualityString(0x54));
            Assert.AreEqual("Sub Normal", OPCLibrary.Converter.GetQualityString(0x58));
            Assert.AreEqual("Good", OPCLibrary.Converter.GetQualityString(0xC0));
            Assert.AreEqual("Local Override", OPCLibrary.Converter.GetQualityString(0xD8));
            Assert.AreEqual("Unknown", OPCLibrary.Converter.GetQualityString(0xFF));

            Assert.AreEqual("VT_EMPTY", OPCLibrary.Converter.GetVTString(0));

            float f = 1;
            IntPtr pF = Marshal.AllocCoTaskMem(4);
            Marshal.Copy(new float[1] { f }, 0, pF, 1);
            int iValue = Marshal.ReadInt32(pF);
            Marshal.FreeCoTaskMem(pF);

            Assert.AreEqual(char.Parse("1"), Converter.GetPropertyValue(16, 1));
            Assert.AreEqual(short.Parse("1"), Converter.GetPropertyValue(2, 1));
            Assert.AreEqual(float.Parse("1"), Converter.GetPropertyValue(4, iValue));
            Assert.AreEqual(1, Converter.GetPropertyValue(3, 1));
            Assert.AreEqual(long.Parse("1"), Converter.GetPropertyValue(20, 1));
            Assert.AreEqual(0, Converter.GetPropertyValue(999, 1));

            DateTime dt = DateTime.Now;
            long vt = dt.ToFileTime();
            uint ht = (uint)(vt >> 32);
            uint lt = (uint)(vt - ((long)ht << 32));
            _FILETIME ft = new _FILETIME() { dwHighDateTime = ht, dwLowDateTime = lt };

            Assert.AreEqual(dt.ToString(), Converter.GetFTSting(ft));

            Assert.AreEqual(typeof(char), Converter.GetVTType(16));
            Assert.AreEqual(typeof(string), Converter.GetVTType(8));
            Assert.AreEqual(typeof(short), Converter.GetVTType(2));
            Assert.AreEqual(typeof(float), Converter.GetVTType(4));
            Assert.AreEqual(typeof(int), Converter.GetVTType(3));
            Assert.AreEqual(typeof(long), Converter.GetVTType(20));

            Assert.AreEqual(sizeof(char), Converter.GetVTSize(16));
            Assert.AreEqual(sizeof(short), Converter.GetVTSize(2));
            Assert.AreEqual(sizeof(float), Converter.GetVTSize(4));
            Assert.AreEqual(sizeof(int), Converter.GetVTSize(3));
            Assert.AreEqual(sizeof(long), Converter.GetVTSize(20));
        }

        [TestMethod]
        public void TestOPCGroupSyncRead()
        {
            string progID = "Insat";
            int errorCount = 0;
            OPCServer server = OPCServer.FindServerByProgID(progID, "localhost");
            server.Connect();
            List<OPCItem> items = server.GetItems();
            var opcBranches = from item in items
                              where item.ItemType == OPCItemType.BRANCH
                              let leaves = from elem in items
                                           where elem.ItemType == OPCItemType.LEAF
                                           where elem.Parent != null && elem.Parent.ItemName == item.ItemName
                                           select elem
                              where leaves.Count() > 0
                              where !item.ItemName.Contains("МВ110")
                              select item;
            foreach(OPCItem branch in opcBranches)
            {
                string groupName = branch.ItemID;                
                var opcItems = from item in items
                               where item.ItemType == OPCItemType.LEAF
                               where item.Parent != null
                               where item.Parent.ItemName == branch.ItemName
                               select item;
                List<OPCItem> list = new List<OPCItem>(opcItems);
                uint udateRate = 100;
                OPCGroup group = new OPCGroup(name: groupName, items: list)
                {
                    UpdateRate = udateRate
                };
                Assert.AreEqual(groupName, group.Name);
                Assert.AreSame(list, group.Items);
                Assert.AreEqual(udateRate, group.UpdateRate);
                group.RegInServer(server);
                try
                {
                    group.SyncRead();
                }
                catch(ServerException ex)
                {
                    string msg;
                    //Запрашиваем у сервера текст ошибки, соответствующий текущему HRESULT 
                    server.Server.GetErrorString(ex.HResult, 2, out msg);
                    //Показываем сообщение ошибки
                    Console.Out.WriteLine(msg, "Ошибка");
                }
                List<string> results1 = new List<string>(group.Items.Select(i => i.TimeStamp ?? "1"));
                System.Threading.Thread.Sleep(1000);
                try
                {
                    group.SyncRead();
                }
                catch (ServerException ex)
                {
                    string msg;
                    //Запрашиваем у сервера текст ошибки, соответствующий текущему HRESULT 
                    server.Server.GetErrorString(ex.HResult, 2, out msg);
                    //Показываем сообщение ошибки
                    Console.Out.WriteLine(msg, "Ошибка");
                }
                List<string> results2 = new List<string>(group.Items.Select(i => i.TimeStamp ?? "2"));
                Assert.AreEqual(results1.Count, results2.Count);
                Assert.AreNotEqual(results1[1], results2[1]);
                group.RemoveFromServer();
                try
                {
                    group.SyncRead();
                }
                catch (ServerException ex)
                {
                    errorCount++;
                    string msg;
                    //Запрашиваем у сервера текст ошибки, соответствующий текущему HRESULT 
                    server.Server.GetErrorString(ex.HResult, 2, out msg);
                    //Показываем сообщение ошибки
                    Console.Out.WriteLine(msg, "Ошибка");
                }
            }
            Assert.AreEqual(errorCount, opcBranches.Count());
            server.Disconnect();      
        }

        [TestMethod]
        public void TestOPCGroupASyncRead()
        {
            string progID = "Insat";
            int errorCount = 0;
            OPCServer server = OPCServer.FindServerByProgID(progID, "localhost");
            server.Connect();
            List<OPCItem> items = server.GetItems();
            var opcBranches = from item in items
                              where item.ItemType == OPCItemType.BRANCH
                              let leaves = from elem in items
                                           where elem.ItemType == OPCItemType.LEAF
                                           where elem.Parent != null && elem.Parent.ItemName == item.ItemName
                                           select elem
                              where leaves.Count() > 0
                              where !item.ItemName.Contains("МВ110")
                              select item;
            foreach (OPCItem branch in opcBranches)
            {
                string groupName = branch.ItemID;
                var opcItems = from item in items
                               where item.ItemType == OPCItemType.LEAF
                               where item.Parent != null
                               where item.Parent.ItemName == branch.ItemName
                               select item;
                List<OPCItem> list = new List<OPCItem>(opcItems);
                uint udateRate = 100;
                OPCGroup group = new OPCGroup(name: groupName, items: list)
                {
                    UpdateRate = udateRate
                };
                Assert.AreEqual(groupName, group.Name);
                Assert.AreSame(list, group.Items);
                Assert.AreEqual(udateRate, group.UpdateRate);
                group.RegInServer(server);
                group.AdviseOnServer(server, group.Items.ToArray()[1]);
                
                try
                {
                    group.ASyncRead();
                }
                catch (ServerException ex)
                {
                    string msg;
                    //Запрашиваем у сервера текст ошибки, соответствующий текущему HRESULT 
                    server.Server.GetErrorString(ex.HResult, 2, out msg);
                    //Показываем сообщение ошибки
                    Console.Out.WriteLine(msg, "Ошибка");
                }
                List<string> results1 = new List<string>(group.Items.Select(i => i.TimeStamp ?? "1"));
                System.Threading.Thread.Sleep(1000);
/*                try
                {
                    group.SyncRead();
                }
                catch (ServerException ex)
                {
                    string msg;
                    //Запрашиваем у сервера текст ошибки, соответствующий текущему HRESULT 
                    server.Server.GetErrorString(ex.HResult, 2, out msg);
                    //Показываем сообщение ошибки
                    Console.Out.WriteLine(msg, "Ошибка");
                }
*/                List<string> results2 = new List<string>(group.Items.Select(i => i.TimeStamp ?? "2"));
                Assert.AreEqual(results1.Count, results2.Count);
                Assert.AreNotEqual(results1[1], results2[1]);
                group.RemoveFromServer();
/*                try
                {
                    group.SyncRead();
                }
                catch (ServerException ex)
                {
                    errorCount++;
                    string msg;
                    //Запрашиваем у сервера текст ошибки, соответствующий текущему HRESULT 
                    server.Server.GetErrorString(ex.HResult, 2, out msg);
                    //Показываем сообщение ошибки
                    Console.Out.WriteLine(msg, "Ошибка");
                }*/
            }
            //Assert.AreEqual(errorCount, opcBranches.Count());
            server.Disconnect();
        }
    }
}
