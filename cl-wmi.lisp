

(defpackage #:cl-wmi
  (:use #:cl #:rdnzl)
  (:export #:wmi-query
	   #:local-wmi-query

	   #:collection-list
	   #:object-properties
	   #:collection-properties	   
	   #:foreach
	   #:mapcoll
	   #:get-type-name
	   #:cast-from-type-name
	   #:unbox-from-type-name

	   #:unbox-object
	   #:unbox-collection))

(in-package #:cl-wmi)

(enable-rdnzl-syntax)

(import-types "System.Management"
	      "ConnectionOptions" "ObjectQuery" "ManagementObjectSearcher" "ManagementScope" 
	      "ManagementObject" "ManagementBaseObject" "ManagementBaseObject[]")

(use-namespace "System")
(use-namespace "System.Management")

;;; ----------- WMI queries ------------------------------------------

(defun wmi-query (querystr &key host (namespace "root\\wmi") username password domain)
  "Execute the WMI query possibly on a remote machine"
  (let ((scopepath (if host
		       (format nil "\\\\~A\\~A" host namespace)
		       namespace))
	(options (new "ConnectionOptions"))
	(query (new "ObjectQuery" querystr)))
    (if username
	(setf [%Username options] username))
    (if password
	(setf [%Password options] password))
    (if domain
	(setf [%Username options] (format nil "~A\\~A" domain (if username username "."))))

    (let ((scope (new "ManagementScope" scopepath options)))
      [Connect scope]
      (let ((searcher (new "ManagementObjectSearcher" scope query)))
	[Get searcher]))))

(defun local-wmi-query (querystr &optional namespace)
  "Execute the WMI query on the local machine with current user privalidegs"
  (let ((searcher 
	 (if namespace
	     (new "ManagementObjectSearcher" (new "ManagementScope" namespace) querystr)
	     (new "ManagementObjectSearcher" querystr))))
    [Get searcher]))

;;; ------------ various utilities follow ------------------------------

(defun get-type-name (container)  
  "Get the type name of the container"
  [%FullName [GetType container]])

(defun cast-from-type-name (container)
  "Auto cast the container based off the type name returned from .NET"
  (cast container (get-type-name container)))

(defun unbox-from-type-name (container)
  "Unbox the container based on the full type name. RDNZL doesn't handle UInts so we deal with those specially"
  (when container
    (let ((type-name (get-type-name container)))
      (if (member type-name '("System.UInt8" "System.UInt16" "System.UInt32" "System.UInt64") :test #'string=)
	  (unbox (cast container "System.Int64"))
	  (unbox (cast-from-type-name container))))))
    
(defmacro foreach ((var collection) &body body)
  "Iterate over the collection using .NET iterators"
  (let ((genum (gensym "ENUMERATOR"))
	(gx (gensym)))
    `(let* ((,genum [GetEnumerator ,collection]))
       (do ((,gx [MoveNext ,genum] [MoveNext ,genum]))
	   ((not ,gx))
	 (let ((,var [%Current ,genum]))
	   ,@body)))))

(defun collection-list (collection)
  "Convert a collection into a list"
  (let ((lst nil))
    (foreach (ob collection)
      (push ob lst))
    (nreverse lst)))

(defun mapcoll (function collection &rest more-collections)
  "Like MAPCAR but for collections"
  (apply #'mapcar 
	 function 
	 (mapcar #'collection-list (cons collection more-collections))))

(defun object-properties (management-object)			  
  "Extract the properties of the collection as an assoc list"
  (mapcar (lambda (property)
	    (let ((val [%Value property]))
	      (cons (intern [%Name property] "KEYWORD")
		    (unbox-from-type-name val))))
	  (collection-list [%Properties management-object])))

(defun collection-properties (collection)
  "Extract the properties of each of the objects in the collection"
  (mapcar #'object-properties (collection-list collection)))

;; ------------------------------------------------------------------------

;; possibly unneeded now .....
(defun cimtype-name (cim-enum)
  "Convert a CIM type enum into a string name"
  (case cim-enum
    (0 "None")
    (2 "SInt16")
    (3 "SInt32")
    (4 "Real32")
    (5 "Real64")
    (8 "String")
    (11 "Boolean")
    (13 "Object")
    (16 "SInt8")
    (17 "UInt8")
    (18 "UInt16")
    (19 "UInt32")
    (20 "SInt64")
    (21 "UInt64")
    (101 "DateTime")
    (102 "Reference")
    (103 "Char16")))

(defun unbox-from-cimtype-name (container type-name)
  "Like UNBOX-FROM-TYPE-NAME but uses the CIM type name"
  (let (cast-type)
    (cond
      ((member type-name '("String" "DateTime") :test #'string=)
       (setf cast-type "System.String"))
      ((string= type-name "Boolean")
       (setf cast-type "System.Boolean"))
      ((string= type-name "Char16")
       (setf cast-type "System.Char"))
      ((member type-name '("SInt8" "SInt16" "SInt32" "SInt64" 
			   "UInt8" "UInt16" "UInt32" "UInt64") 
	       :test #'string=)
       (setf cast-type "System.Int64")))
    (if (and container cast-type)
	(progn
	  (cast container cast-type)
	  (unbox container))
	container)))


;; ------------------------------------------------------------------------


(defun unbox-object (management-object)
  "An attempt to unbox all properties of the object, calling itself recursively if required"
  (mapcar (lambda (property)
	    (destructuring-bind (name . value) property
	      (cons name
		    (if (container-p value)
			(if [%IsArray [GetType value]]
			    (mapcar #'unbox-object (rdnzl-array-to-list value))
			    (unbox-object value))
			value))))
	  (object-properties management-object)))

(defun unbox-collection (collection)
  "Unbox each object in the collection"
  (mapcar #'unbox-object (collection-list collection)))

