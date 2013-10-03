

(defpackage #:cl-wmi
  (:use #:cl #:rdnzl)
  (:export #:wmi-query
	   #:local-wmi-query

	   #:collection-list
	   #:object-properties
	   #:collection-properties
	   #:object-methods

	   #:foreach
	   #:mapcoll
	   #:get-type-name
	   #:cast-from-type-name
	   #:unbox-from-type-name

	   #:unbox-object
	   #:unbox-collection
	   #:invoke-method
	   #:invoke-class-method
	   #:make-management-class

	   #:registry-values
	   #:registry-subkeys
	   #:registry-query
	   #:set-registry-value
	   #:delete-registry-value
	   #:create-registry-key
	   #:delete-registry-key))

(in-package #:cl-wmi)

(enable-rdnzl-syntax)

(import-types "System.Management"
	      "ConnectionOptions" "ObjectQuery" "ManagementObjectSearcher" "ManagementScope" 
	      "ManagementObject" "ManagementBaseObject" "ManagementBaseObject[]"
	      "ManagementClass" "ObjectGetOptions" "ManagementPath"
	      "AuthenticationLevel" "ImpersonationLevel" "InvokeMethodOptions")

(use-namespace "System")
(use-namespace "System.Management")

;;; ----------- WMI queries ------------------------------------------

(defun make-management-scope (&key host (namespace "root\\wmi") username password domain)
  "Create a management scope"
  (let ((scopepath (if host
		       (format nil "\\\\~A\\~A" host namespace)
		       namespace))
	(options (new "ConnectionOptions")))
    (if username
	(setf [%Username options] username))
    (if password
	(setf [%Password options] password))
    (if domain
	(setf [%Username options] (format nil "~A\\~A" domain (if username username "."))))
    (new "ManagementScope" scopepath options)))

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
	     (new "ManagementObjectSearcher" namespace querystr)
	     (new "ManagementObjectSearcher" querystr))))
    [Get searcher]))

;;; ------------ various utilities follow ------------------------------

(defun get-type-name (container)  
  "Get the type name of the container"
  [%FullName [GetType container]])

(defun rdnzl-array-p (container) 
  "Is this a rdnzl array?"
  [%IsArray [GetType container]])

(defun cast-from-type-name (container)
  "Auto cast the container based off the type name returned from .NET"
  (cast container (get-type-name container)))

(defun unbox-from-type-name (container)
  "Unbox the container based on the full type name. RDNZL doesn't handle UInts so we deal with those specially"
  (if (container-p container)      
      (let ((type [GetType container]))
	(cast-from-type-name container)
	(cond
	  ([%IsArray type]
	   (rdnzl-array-to-list container))
	  ((member [%FullName type] '("System.UInt8" "System.UInt16" "System.UInt32" "System.UInt64") :test #'string=)
	   (unbox (cast container "System.Int64")))
	  (t 
	   (unbox (cast-from-type-name container)))))
      container))
    
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
  "Extract the properties of the ManagementObject as an assoc list"
  (mapcar (lambda (property)
	    (let ((val [%Value property]))
	      (cons (intern [%Name property] "KEYWORD")
		    (unbox-from-type-name val))))
	  (collection-list [%Properties management-object])))

(defun collection-properties (collection)
  "Extract the properties of each of the objects in the collection"
  (mapcar #'object-properties (collection-list collection)))

(defun object-methods (management-object)
  "Extract the methods of the ManagementObject"
  (let ((object-class (new "ManagementClass" [%ClassPath management-object])))
    [%Methods object-class]))


;; ------------------------------------------------------------------------

(defun unbox-object (management-object)
  "An attempt to unbox all properties of the object, calling itself recursively if required"
  (if (container-p management-object)
      (mapcar (lambda (property)
		(destructuring-bind (name . value) property
		  (cons name
			;; if the value is a container then it's a managementobject (or array of management objects)
			;; object-properties already unboxes basic types
			(cond
			  ((container-p value)
			   (if [%IsArray [GetType value]]
			       (mapcar #'unbox-object (rdnzl-array-to-list value))
			       (unbox-object value)))
			  ((listp value)
			   (mapcar #'unbox-object value))
			  (t value)))))
	      (object-properties management-object))
      management-object))

(defun unbox-collection (collection)
  "Unbox each object in the collection"
  (mapcar #'unbox-object (collection-list collection)))


;;; ---------------------------


(defun invoke-method (management-object method-name &rest args)
  "Invoke the named method on the managementobject with the specified arguments"
  (apply #'invoke 	 
	 management-object
	  "InvokeMethod"
	  method-name
	  args))

(defun invoke-class-method (class-name method-name &rest args)
  "Invoke the method of the named ManagementClass"
  (apply #'invoke-method 
	 (new "ManagementClass" class-name)
	 method-name
	 args))

(defun make-management-class (class-name &key host (namespace "root\\wmi") username password domain wow64)
  "Create a management class, possibly on a remote machine"
  (let ((scopepath (if host
		       (format nil "\\\\~A\\~A" host namespace)
		       namespace))
	(options (new "ConnectionOptions")))
    (if username
	(setf [%Username options] username))
    (if password
	(setf [%Password options] password))
    (if domain
	(setf [%Username options] (format nil "~A\\~A" domain (if username username "."))))

    (setf [%EnablePrivileges options] t
	  [%Authentication options] [$AuthenticationLevel.Default]
	  [%Impersonation options] [$ImpersonationLevel.Impersonate])

    (let ((scope (new "ManagementScope" scopepath options))
	  (getoptions (new "ObjectGetOptions"))
	  (path (new "ManagementPath" class-name)))

      ;; hack to access 64 bit hive on remote machine
      ;; whoever would've guessed to do this?
      (when wow64
	[Add [%Context [%Options scope]] "__ProviderArchitecture" 64]
	[Add [%Context [%Options scope]] "__RequiredArchitecture" t])

      [Connect scope]
      (new "ManagementClass" scope path getoptions))))


;;;; --------------- registry queries -----------------------------------------


;; definitions for HKEY_LOCAL_MACHINE, these are normally given in hex,
;; 80000001, 80000002 etc.
(defparameter *hkey-trees* '((:classes-root . 2147483648) (:current-user . 2147483649) 
			    (:local-machine . 2147483650) (:users . 2147483651) (:current-config . 2147483653)))

;; the names of the "getter" methods
(defparameter *get-methods* '((:string . "GetStringValue") (:binary . "GetBinaryValue") (:dword . "GetDWORDValue")
			     (:expanded-string . "GetExpandedStringValue") (:multi-string . "GetMultiStringValue")))

;; the names of the "setter" methods
(defparameter *set-methods* '((:string . "SetStringValue") (:binary . "SetBinaryValue") (:dword . "SetDWORDValue")
			     (:expanded-string . "SetExpandedStringValue") (:multi-string . "SetMultiStringValue")))


;; A note on the wow64 keyword argument:
;; Windows has two different registry databases, a 32 bit and a 64 bit "hive"
;; If you call the WMI functions from a windows process running either on a 32 bit machine or under 32 bit 
;; windows-on-windows-64 (i.e. a 32 program on a 64-bit windows machine)
;; then the WMI query executes as a 32 bit query. This means you'll only have access to the 32 bit 
;; registry, even if you are querying a remote machine that is a 64-bit system!
;; Most Lisps on windows are 32 bit but typically you'll want to access the 64 bit hive.
;; Setting this key to t should allow access to the 64 hives in such situations

(defun registry-values (key-name &key tree host username password domain (wow64 t))
  "Returns a list of values with their types on this registry key"
  (let* ((mc (make-management-class "StdRegProv" 
				    :host host
				    :namespace "root\\cimv2"
				    :username username
				    :password password
				    :domain domain
				    :wow64 wow64))
	 (inparams [GetMethodParameters mc "EnumValues"]))

    ;; the tree is HKEY_LOCAL_MACHINE by default
    (if tree
	[SetPropertyValue inparams "hDefKey" (cast (box (cdr (assoc tree *hkey-trees*))) "System.UInt32")])
    [SetPropertyValue inparams "sSubKeyName" key-name]

    (let* ((outparams (invoke-method mc "EnumValues" 
				     inparams (make-null-object "System.Management.InvokeMethodOptions"))))
      (values 
       (mapcar (lambda (name type)
		 (list (intern name "KEYWORD")
		       (case type
			 (1 :string)
			 (2 :expanded-string)
			 (3 :binary)
			 (4 :dword)
			 (5 :multi-string)
			 (otherwise :unknown))))		      		 
	       (unbox-from-type-name [GetPropertyValue outparams "sNames"])
	       (unbox-from-type-name [GetPropertyValue outparams "Types"]))
       (unbox-from-type-name [GetPropertyValue outparams "ReturnValue"])))))

(defun registry-subkeys (key-name &key tree host username password domain (wow64 t))
  "Returns a list of sub-keys on this registry key"
  (let* ((mc (make-management-class "StdRegProv" 
				    :host host
				    :namespace "root\\cimv2"
				    :username username
				    :password password
				    :domain domain
				    :wow64 wow64))
	 (inparams [GetMethodParameters mc "EnumKey"]))

    (if tree
	[SetPropertyValue inparams "hDefKey" (cast (box (cdr (assoc tree +hkey-trees+))) "System.UInt32")])
    [SetPropertyValue inparams "sSubKeyName" key-name]

    (let* ((outparams (invoke-method mc "EnumKey" 
				     inparams (make-null-object "System.Management.InvokeMethodOptions"))))
      (values 
       (unbox-from-type-name [GetPropertyValue outparams "sNames"])
       (unbox-from-type-name [GetPropertyValue outparams "ReturnValue"])))))

(defun registry-query (key-name &key tree host username password domain (wow64 t))
  "Return all the values on the given registry key as an assoc list"
  (multiple-value-bind (values err-code) (registry-values key-name :tree tree :host host :username username 
							  :password password :domain domain :wow64 wow64)
    (when (zerop err-code)
      (mapcar (lambda (valpair)
		(destructuring-bind (value-name type) valpair
		  (let ((meth-name (cdr (assoc type *get-methods*))))
		    (let* ((mc (make-management-class "StdRegProv" 
						      :host host
						      :namespace "root\\cimv2"
						      :username username
						      :password password
						      :domain domain
						      :wow64 wow64))
			   (inparams [GetMethodParameters mc meth-name]))

		      (if tree
			  [SetPropertyValue inparams "hDefKey" (cast (box (cdr (assoc tree *hkey-trees*))) "System.UInt32")])
		      [SetPropertyValue inparams "sSubKeyName" key-name]
		      [SetPropertyValue inparams "sValueName" (symbol-name value-name)]
		      
		      (let* ((outparams (invoke-method mc meth-name
						       inparams 
						       (make-null-object "System.Management.InvokeMethodOptions"))))
			(values (cons value-name
				      (unbox-from-type-name 				 
				       [GetPropertyValue outparams
				       ;; Microsoft uses hungarian notation so we need different names depending on the type
				       (if (member type '(:string :multi-string :expanded-string))
					   "sValue"
					   "uValue")]))
				(unbox-from-type-name [GetPropertyValue outparams "ReturnValue"])))))))
	      values))))

(defun set-registry-value (key-name value-name value type &key tree host username password domain (wow64 t))
  "Sets the value on the remote machine. Type should be :string, :dword etc"
  (let* ((meth-name (cdr (assoc type *set-methods*)))
	 (mc (make-management-class "StdRegProv" 
				    :host host
				    :namespace "root\\cimv2"
				    :username username
				    :password password
				    :domain domain
				    :wow64 wow64))
	 (inparams [GetMethodParameters mc meth-name]))

    (if tree
	[SetPropertyValue inparams "hDefKey" (cast (box (cdr (assoc tree *hkey-trees*))) "System.UInt32")])
    [SetPropertyValue inparams "sSubKeyName" key-name]
    [SetPropertyValue inparams "sValueName" value-name]
    [SetPropertyValue inparams 
                      (if (member type '(:string :extended-string :multi-string))
			  "sValue"
			  "uValue")
		      value]

    (let ((outparams (invoke-method mc meth-name 
				     inparams (make-null-object "System.Management.InvokeMethodOptions"))))
      (unbox-from-type-name [GetPropertyValue outparams "ReturnValue"]))))


(defun delete-registry-value (key-name value-name &key tree host username password domain (wow64 t))
  "Deletes the value on the remote machine. Type should be :string, :dword etc"
  (let* ((mc (make-management-class "StdRegProv" 
				    :host host
				    :namespace "root\\cimv2"
				    :username username
				    :password password
				    :domain domain
				    :wow64 wow64))
	 (inparams [GetMethodParameters mc "DeleteValue"]))

    (if tree
	[SetPropertyValue inparams "hDefKey" (cast (box (cdr (assoc tree +hkey-trees+))) "System.UInt32")])
    [SetPropertyValue inparams "sSubKeyName" key-name]
    [SetPropertyValue inparams "sValueName" value-name]

    (let ((outparams (invoke-method mc "DeleteValue"
				     inparams (make-null-object "System.Management.InvokeMethodOptions"))))
      (unbox-from-type-name [GetPropertyValue outparams "ReturnValue"]))))

(defun create-registry-key (key-name &key tree host username password domain (wow64 t))
  "Creates the key on the remote machine. Type should be :string, :dword etc"
  (let* ((mc (make-management-class "StdRegProv" 
				    :host host
				    :namespace "root\\cimv2"
				    :username username
				    :password password
				    :domain domain
				    :wow64 wow64))
	 (inparams [GetMethodParameters mc "CreateKey"]))

    (if tree
	[SetPropertyValue inparams "hDefKey" (cast (box (cdr (assoc tree *hkey-trees*))) "System.UInt32")])
    [SetPropertyValue inparams "sSubKeyName" key-name]

    (let ((outparams (invoke-method mc "CreateKey"
				     inparams (make-null-object "System.Management.InvokeMethodOptions"))))
      (unbox-from-type-name [GetPropertyValue outparams "ReturnValue"]))))

(defun delete-registry-key (key-name &key tree host username password domain (wow64 t))
  "Delets the key on the remote machine. Type should be :string, :dword etc"
  (let* ((mc (make-management-class "StdRegProv" 
				    :host host
				    :namespace "root\\cimv2"
				    :username username
				    :password password
				    :domain domain
				    :wow64 wow64))
	 (inparams [GetMethodParameters mc "DeleteKey"]))

    (if tree
	[SetPropertyValue inparams "hDefKey" (cast (box (cdr (assoc tree *hkey-trees*))) "System.UInt32")])
    [SetPropertyValue inparams "sSubKeyName" key-name]

    (let ((outparams (invoke-method mc "DeleteKey"
				     inparams (make-null-object "System.Management.InvokeMethodOptions"))))
      (unbox-from-type-name [GetPropertyValue outparams "ReturnValue"]))))

