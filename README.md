Registry Value Retriever
========================

Registry Value Retriever is a simple console application that can create a specified registry key within the registry.

Usage:
   ApplicationName [RegistryRoot] [RegistryKey] {RegistryValue} {ValueData}
   where
     [RegistryRoot] is the registry root key, e.g. HKEY_LOCAL_MACHINE
     [RegistryKey] is the actual fully qualified key to create minus the root key, e.g. SOFTWARE\Microsoft
     {RegistryValue}: Optional Registry Value name to be created")
     {ValueData}: Optional Registry Value Data to be set")
Example:
  "Registry Value Retriever.exe" "HKEY_LOCAL_MACHINE" "SOFTWARE\Microsoft\.NETFramework" 
Output:
  Success.

Created by Craig Lotter, January 2006

*********************************

Project Details:

Coded in Visual Basic .NET using Visual Studio .NET 2003
Implements concepts such as Registry Programming.
Level of Complexity: simple
