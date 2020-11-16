using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Schema;
using Microsoft.Win32;
using Neo4j.Driver;

namespace COM_Mapper
{
    class Program
    {
        static private IDriver _driver;

        /*
         * Given a registry key of a COM class, initialize the class and call QueryInterface on each
         * interface passed in the list of interface registry keys. Returns a list of interface RegistryKeys
         * that the class supports.
        */
        static List<RegistryKey> GetAllSupportedInterfaces(RegistryKey cls, List<RegistryKey> interfaces)
        {
            List<RegistryKey> supportedInterfaces = new List<RegistryKey>();

            var clsId = cls.Name.Split('\\').Last();
            Type clsType;
            object clsObj = null;

            Console.WriteLine("[*] Getting all supported interfaces for class : " + clsId);

            try
            {
                clsType = Type.GetTypeFromCLSID(Guid.Parse(clsId), false);
            }
            catch (System.FormatException e)
            {
                Console.WriteLine("[!] Unable to convert to GUID: " + clsId);
                return supportedInterfaces;
            }

            try
            {
                clsObj = Activator.CreateInstance(clsType);
            }
            catch(Exception e)
            {
                Console.WriteLine("[!] " + e.ToString());
                return supportedInterfaces;
            }

            var iUnknown = Marshal.GetIUnknownForObject(clsObj);

            foreach (RegistryKey ifaceKey in interfaces)
            {
                String iid;
                IntPtr ppv;

                iid = ifaceKey.Name.Split('\\').Last().ToUpper();

                Guid ifaceGuid = new Guid(iid);
                var hr = Marshal.QueryInterface(iUnknown, ref ifaceGuid, out ppv);

                if (hr == 0)
                {
                    supportedInterfaces.Add(ifaceKey);
                }
            }

            return supportedInterfaces;
        }

        /* 
         * Get a dictionary of all the registered COM classes where the key
         * is the CLSID string and the value is the corresponding RegistryKey
         */
        static Dictionary<String, RegistryKey> GetAllCOMClasses()
        {
            var keys = new Dictionary<String, RegistryKey>();

            RegistryKey key = Registry.ClassesRoot.OpenSubKey("CLSID");

            foreach (var cls in key.GetSubKeyNames())
            {
                List<String> blocked_guids = new List<String>();
                
                RegistryKey class_key = key.OpenSubKey(cls);
                Console.WriteLine("\t[*] CLSID: " + cls);
                keys[cls.ToUpper()] = class_key;
            }

            key = Registry.LocalMachine.OpenSubKey("SOFTWARE\\WOW6432Node\\Classes\\CLSID");
            foreach (var cls in key.GetSubKeyNames())
            {
                RegistryKey class_key = key.OpenSubKey(cls);

                if (!keys.ContainsKey(cls))
                {
                    Console.WriteLine("\t[*] CLSID: " + cls);
                    keys[cls.ToUpper()] = class_key;
                }
            }

            return keys;
        }

        /*
         * Gets a list of all the COM interfaces registered on the OS. Returns a list of registry keys.
         */
        static List<RegistryKey> GetAllCOMInterfaces()
        {
            List<RegistryKey> interfaces = new List<RegistryKey>();
            
            RegistryKey key = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Classes\\Interface");

            foreach (var iface in key.GetSubKeyNames())
            {
                interfaces.Add(key.OpenSubKey(iface));
            }

            key = Registry.ClassesRoot.OpenSubKey("Interface");

            foreach (var iface in key.GetSubKeyNames())
            {
                interfaces.Add(key.OpenSubKey(iface));
            }

            return interfaces;
        }

        /*
         * Creates a neo4j node that represents a COM class. Returns true if successful.
         */ 
        static bool CreateClassNode(RegistryKey cls)
        {
            String clsId = cls.Name.Split('\\').Last().ToUpper();
            String className = "";

            var defaultValue = cls.GetValue("");
            if (defaultValue != null)
            {
                className = defaultValue.ToString();
            }


            RegistryKey inProcServer =  cls.OpenSubKey("InprocServer32");
            RegistryKey progId = cls.OpenSubKey("ProgID");
            RegistryKey inProcHandler = cls.OpenSubKey("InProcHandler32");
            RegistryKey localServer32 = cls.OpenSubKey("LocalServer32");
            RegistryKey version = cls.OpenSubKey("Version");


            using (var session = _driver.Session())
            {
                try
                {
                    var newNode = session.WriteTransaction(
                        tx =>
                        {
                            var query = "CREATE (newClass:ComClass {clsid:\"" + clsId + "\"}) ";
                            if (className != null)
                            {
                                query += "SET newClass.name = \"" + className + "\" ";
                            }
                            if (inProcServer != null)
                            {
                                String tmp = inProcServer.GetValue("").ToString().Replace("\\", "\\\\");
                                query += "SET newClass.InProcServer32 = \"" + tmp + "\" ";
                            }
                            if (progId != null)
                            {
                                String tmp = progId.GetValue("").ToString().Replace("\\", "\\\\");
                                query += "SET newClass.ProgId = \"" + tmp + "\" ";
                            }
                            if (inProcHandler != null)
                            {
                                String tmp = inProcHandler.GetValue("").ToString().Replace("\\", "\\\\");
                                query += "SET newClass.InProcHandler32 = \"" + tmp + "\" ";
                            }
                            if (localServer32 != null)
                            {
                                String tmp = localServer32.GetValue("").ToString().Replace("\\", "\\\\");
                                query += "SET newClass.LocalServer32 = \"" + tmp + "\" ";
                            }
                            if (version != null)
                            {
                                String tmp = version.GetValue("").ToString().Replace("\\", "\\\\");
                                query += "SET newClass.version = \"" + tmp + "\" ";
                            }
                            query += "RETURN newClass";
                            var result = tx.Run(query);
                            return result.Single()[0].As<string>();
                        }
                    );
                }
                catch (Neo4j.Driver.ClientException e)
                {
                    // If error message isn't "Object already exists"
                    return true;
                }
                catch (Exception e)
                {
                    Console.WriteLine("[!] " + e.ToString());
                    return false;
                }
            }
            return true;
        }

        /*
         * Creates an interface node in Neo4j. Returns true if successful.
         */
        static bool CreateInterfaceNode(RegistryKey ifaceKey)
        {
            using (var session = _driver.Session())
            {
                String ifaceId = ifaceKey.Name.Split('\\').Last().ToUpper();
                try
                {
                    var newNode = session.WriteTransaction(
                        tx =>
                        {
                            var query = "CREATE (newIface:ComInterface {iid:\"" + ifaceId + "\"}) ";
                            var ifaceName = ifaceKey.GetValue("");
                            if (ifaceName != null)
                            {
                                query += "SET newIface.name = \"" + ifaceName + "\" ";
                            }

                            query += "RETURN newIface";
                            var result = tx.Run(query);
                            return result.Single()[0].As<string>();
                        }
                    );
                    Console.WriteLine("[*] Created interface in neo4j : " + ifaceId);
                }
                catch (Neo4j.Driver.ClientException e)
                {
                    // If error message isn't "Object already exists"
                    return true;
                }
                catch(Exception e)
                {
                    Console.WriteLine("[!] " + e.ToString());
                    return false;
                }
            }
            return true;
        }

        /*
         * Creates a "implements" relationship between a COM class and a COM interface. Returns true if successful.
         */
        static bool CreateRelationship(String clsId, String ifaceId)
        {
            using (var session = _driver.Session())
            {
                try
                {
                    var newNode = session.WriteTransaction(
                        tx =>
                        {
                            var query = "MATCH (cclass:ComClass {clsid:\"" + clsId + "\"})" + 
                                        "MATCH(iface: ComInterface { iid: \"" + ifaceId + "\"})" +
                                        "MERGE(cclass) - [i: implements]-> (iface)" +
                                        "RETURN cclass, iface";
                            var result = tx.Run(query);
                            return result.Single()[0].As<string>();
                        }
                    );
                    Console.WriteLine("[*] Created relationship in neo4j : " + clsId + " -> " + ifaceId);
                }
                catch (Neo4j.Driver.ClientException e)
                {
                    // If error message isn't "Object already exists"
                    return true;
                }
                catch(Exception e)
                {
                    Console.WriteLine("[!] " + e.ToString());
                    return false;
                }
            }
            return true;
        }

        /*
         * Given a COM class and a list of supported interfaces, this function will create the
         * requisite nodes and then create the appropriate relationships. Returns true if successful.
         */
        static bool SendToNeo4J(RegistryKey cls, List<RegistryKey> ifaces)
        {
            String clsId = cls.Name.Split('\\').Last().ToUpper();
            
            if (!CreateClassNode(cls))
            {
                return false;
            }

            foreach (RegistryKey iface in ifaces)
            {
                String iface_iid = iface.Name.Split('\\').Last().ToUpper();
                CreateInterfaceNode(iface);
                CreateRelationship(clsId, iface_iid);
            }

            return true;
        }

        /*
         * You will need to babysit this, as many COM classes will produce pop-ups or crash the process.
         * If the process crashes, you can provide the last COM class ID as an argument, and the program will pick
         * up where it left off.
         */
        static void Main(string[] args)
        {
            Console.WriteLine("[*] Starting COM mapper");
            String start_class = "";

            try
            {
                start_class = args[0].ToUpper();
            }
            catch (System.IndexOutOfRangeException)
            {
                Console.WriteLine("[*] No starting class provided. Starting from beginning");
            }

            Dictionary<String, RegistryKey> classes = GetAllCOMClasses();

            List<RegistryKey> interfaces = GetAllCOMInterfaces();

            _driver = GraphDatabase.Driver("bolt://127.0.0.1:7687", AuthTokens.Basic("neo4j", "neo4j"));

            foreach (KeyValuePair<String, RegistryKey> entry in classes)
            {

                var classId = entry.Key;
                var classRegKey = entry.Value;

                Console.WriteLine("[*] Checking : " + classId);
                if (String.Compare(classId, start_class) <= 0)
                {
                    continue;
                }

                Console.WriteLine("[*] Getting interface for " + classId);
                List<RegistryKey> supportedInterfaces = GetAllSupportedInterfaces(classRegKey, interfaces);
                if (supportedInterfaces.Count > 0)
                {
                    try
                    {
                        SendToNeo4J(classRegKey, supportedInterfaces);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("[!] " + e.ToString());
                        continue;
                    }
                    
                }
            }

            Console.WriteLine("[*] Finished.");

        }
    }
}
