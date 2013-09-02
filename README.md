CL-WMI
=======

CL-WMI is a package for making WMI querys and manipulating the results in Common Lisp. 
It accesses the .NET platform using the RDNZL library, which must first be installed on your system.
The source files can be found at http://weitz.de/rdnzl/

If you use quicklisp, it is advisable to put the RDNZL Lisp files in your quicklisp/local-projects directory
with the appropriate line the the systems.txt file, so that quicklisp can load RDNZL.

Overview
-----------

The package provides two functions for making WMI queries. WMI-QUERY makes a query possibly on a 
remote machine with login credentials if required. LOCAL-WMI-QUERY makes a query on the local machine
using current user privileges. These two functions are typically used to do get information, e.g.

(wmi-query "Select * From MSiSCSIInitiator_TargetClass" 
	   :host "ahost1"
	   :namespace "root\\wmi"
	   :username "administrator"
	   :password "password")

This returns a ManagementObjectCollection RDNZL container. Getting useful information out of this container
requires further .NET calls, and is implemented by various utility functions provided with CL-WMI.

* collection-list collection
This converts a ManagementObjectCollection into a list of the objects it contains.

* object-properties
This returns an assoc list of the properties the object contains, with property names as keywords.

* collection-properties	
This maps over the objects in the collection, getting the object properties using the above.

* foreach
General purpose macro for iterating over .NET objects. 

* mapcoll
This is basically a MAPCAR but for collections

* get-type-name
Often the objects found from the ManagementObjectCollections are returned as RDNZL containers with the type
"System.Object". This function extracts the true type name of the object.

* cast-from-type-name
This casts a RDNZL container using the true type name

* unbox-from-type-name
This unboxes the RDNZL container using the true type name, and also handles unsigned integers which RDNZL doesn't
support.

* unbox-object
This unboxes each of the properties of the ManagementObject, recursively calling itself on properties that 
are either arrays or other ManagementObjects

* unbox-collection
This just applies unbox-object to each object in the collection

