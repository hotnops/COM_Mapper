# COM_Mapper
A tool to create COM class/interface relationships in neo4j

# Intro
COM_Mapper is a C# tool that maps out COM class/interface relationships in neo4j, enabling the user to perform custom on COM class implementation. COM_Mapper works by iterating through each COM class registered under HKCR\CLSID and HKLM\Software\WOW64Node\Classes\CLSID. COM_Mapper will initialize each class and attempt to call QueryInterface for every interface registered under HKLM\Software\Classes\Interface and HKCR\Interface. If the COM class supports the interface, a relationship will be created in neo4j. 

# Usage
You need to stand up a neo4j server and modify Program.cs to point to the instance IP and port as well as the credentails. These will be command line arguments on the next update. Once neo4j is up and running, run COM_Mapper.exe. You will need to babysit this process as many COM classes will produce popups and initializing some COM classes will cause the process to crash. If the process crashes, you can provide the last CLSID processed as a command line argument and COM_Mapper will resume at that CLSID. Once COM_Mapper is finished, you will have all COM classes and interfaces saved in a neo4j database  that you can then query for research.

# Examples
## Querying for all classes that implement a set of interfaces
```
MATCH (c:ComClass) - [:implements] -> (:ComInterface {name:"IPersistFile"})
MATCH (c:ComClass) - [:implements] -> (:ComInterface {name:"IOleObject"})
RETURN c
```
## Querying for all interfaces implemented by a class
```
MATCH (c:ComClass {clsid:"{1C82EAD9-508E-11D1-8DCF-00C04FB951F9}"}) - [:implements] -> (i:ComInterface)
RETURN c, i
```
## Querying for all COM classes implemented by a dll
```
MATCH (c:ComClass {InProcServer32:"C:\\Windows\\SysWOW64\\mshtml.dll"})
RETURN c
```
## Querying for all COM classes that have a LocalServer32
```
MATCH (c:ComClass) WHERE EXISTS (c.LocalServer32)
RETURN c
```
