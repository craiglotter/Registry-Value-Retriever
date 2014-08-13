Registry Value Retriever
========================

Registry Value Retriever is a simple console application that prints out the data found in the specified registry key value as a string.

Usage:
   ApplicationName [RegistryRoot] [RegistryKey] [ValueName]
   where
     [RegistryRoot] is the registry root key, e.g. HKEY_LOCAL_MACHINE
     [RegistryKey] is the actual fully qualified key minus the root key, e.g. SOFTWARE\Microsoft
     [ValueName] is the name of the Value for which the data is to be returned
Example:
  "Registry Value Retriever.exe" "HKEY_LOCAL_MACHINE" "SOFTWARE\Microsoft\.NETFramework" "InstallRoot"
Output:
  C:\WINDOWS\Microsoft.NET\Framework\

Created by Craig Lotter, January 2006

*********************************

Project Details:

Coded in Visual Basic .NET using Visual Studio .NET 2003
Implements concepts such as Registry Programming.
Level of Complexity: simple
