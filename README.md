# VbTrickThreading
## Module for working with multithreading in VB6

<p align="center">Hello everyone!</p>

I present the module for working with multithreading in VB6 for Standard EXE projects. This module is based on [this](http://www.vbforums.com/showthread.php?788713-VB6-Multithreading-in-VB6-part-4-multithreading-in-Standart-EXE) solution with some bugfixing and the new functionality is added. The module doesn't require any additional dependencies and type libraries, works as in the IDE (all the functions work in the main thread) as in the compiled form.

[![Watch video](http://img.youtube.com/vi/D1-3PlAoEnk/0.jpg)](http://www.youtube.com/watch?v=D1-3PlAoEnk)

To start working with the module, you need to call the **Initialize** function, which initializes the necessary data (it initializes the critical sections for exclusive access to the heaps of marshalinig and threads, modifies **VBHeader** ([here](http://www.vbforums.com/showthread.php?788713-VB6-Multithreading-in-VB6-part-4-multithreading-in-Standart-EXE) is description), allocates a **TLS** slot for passing the parameters to the thread).

The main function of thread creation is **vbCreateThread**, which is an analog of the [CreateThread](https://msdn.microsoft.com/en-us/library/windows/desktop/ms682453(v=vs.85).aspx) function.

```vb
' // Create a new thread
Public Function vbCreateThread(ByVal lpThreadAttributes As Long, _
                               ByVal dwStackSize As Long, _
                               ByVal lpStartAddress As Long, _
                               ByVal lpParameter As Long, _
                               ByVal dwCreationFlags As Long, _
                               ByRef lpThreadId As Long, _
                               Optional ByVal bIDEInSameThread As Boolean = True) As Long
```

The function creates a thread and calls the function passed in the **lpStartAddress** parameter with the **lpParameter** parameter.
In the IDE, the call is reduced to a simple call by the pointer implemented through [DispCallFunc](url=https://msdn.microsoft.com/en-us/library/windows/desktop/ms221473(v=vs.85).aspx). In the compiled form, this function works differently. Because a thread requires initialization of project-specific data and initialization of the runtime, the parameters passed to **lpStartAddress** and **lpParameter** are temporarily stored into the heap by the **PrepareData** function, and the thread is created in the **ThreadProc** function, which immediately deals with the initialization and calling of the user-defined function with the user parameter. This function creates a copy of the **VBHeader** structure via **CreateVBHeaderCopy** and changes the public variable placement data in the **VbPublicObjectDescriptor.lpPublicBytes**, **VbPublicObjectDescriptor.lpStaticBytes** structures (BTW it wasn't implemented in the previous version) so that global variables are not affected during initialization. Further, **VBDllGetClassObject** calls the **FakeMain** function (whose address is written to the modified **VBHeader** structure). To transfer user parameters, it uses a **TLS** slot (since **Main** function doesn't accept parameters, details [here](http://www.vbforums.com/showthread.php?788713-VB6-Multithreading-in-VB6-part-4-multithreading-in-Standart-EXE)). In **FakeMain**, parameters are directly extracted from **TLS** and a user procedure is called. The return value of the function is also passed back through **TLS**. There is one interesting point related to the copy of the header that wasn't included in the previous version. Because the runtime uses the header after the thread ends (with **DLL_THREAD_DETACH**), we can't release the header in the **ThreadProc** procedure, therefore there will be a memory leak. To prevent the memory leaks, the heap of fixed size is used, the headers aren't cleared until there is a free memory in this heap. As soon as the memory ends (and it's allocated in the **CreateVBHeaderCopy** function), resources are cleared. The first **DWORD** of header actually stores the **ID** of the thread which it was created in and the **FreeUnusedHeaders** function checks all the headers in the heap. If a thread is completed, the memory is freed (although the **ID** can be repeated, but this doesn't play a special role, since in any case there will be a free memory in the heap and if the header isn't freed in one case, it will be released later). Due to the fact that the cleanup process can be run immediately from several threads, access to the cleanup is shared by the critical section **tLockHeap.tWinApiSection** and if some thread is already cleaning up the memory the function will return True which means that the calling thread should little bit waits and the memory will be available.

The another feature of the module is the ability to initialize the runtime and the project and call the callback function. This can be useful for callback functions that can be called in the context of an arbitrary thread (for example, **InternetStatusCallback**). To do this, use the **InitCurrentThreadAndCallFunction** and **InitCurrentThreadAndCallFunctionIDEProc** functions. The first one is used in the compiled application and takes the address of the callback function that will be called after the runtime initialization, as well as the parameter to be passed to this function. The address of the first parameter is passed to the callback procedure to refer to it in the user procedure:

```vb
' // This function is used in compiled form
Public Function CallbackProc( _
                ByVal lThreadId As Long, _
                ByVal sKey As String, _
                ByVal fTimeFromLastTick As Single) As Long
    ' // Init runtime and call CallBackProc_user with VarPtr(lThreadId) parameter
    InitCurrentThreadAndCallFunction AddressOf CallBackProc_user, VarPtr(lThreadId), CallbackProc
End Function

' // Callback function is called by runtime/window proc (in IDE)
Public Function CallBackProc_user( _
                ByRef tParam As tCallbackParams) As Long

End Function
```

**CallBackProc_user** will be called with the initialized runtime.

This function doesn't work in the **IDE** because in the **IDE** everything works in the main thread. For debugging in the **IDE** the function **InitCurrentThreadAndCallFunctionIDEProc** is used which returns the address of the assembler thunk that translates the call to the main thread and calls the user function in the context of the main thread. This function takes the address of the user's callback function and the size of the parameters in bytes. It always passes the address of the first parameter as a parameter of a user-defined function. I'll tell you a little more about the work of this approach in the **IDE**. To translate a call from the calling thread to the main thread it uses a message-only window. This window is created by calling the **InitializeMessageWindow** function. The first call creates a **WindowProc** procedure with the following code:

```asm
    CMP DWORD [ESP+8], WM_ONCALLBACK
    JE SHORT L
    JMP DefWindowProcW
L:  PUSH DWORD PTR SS:[ESP+10]
    CALL DWORD PTR SS:[ESP+10]
    RETN 10
```

As you can see from the code, this procedure "listens" to the **WM_ONCALLBACK** message which contains the parameter **wParam** - the function address, and in the **lParam** parameters. Upon receiving this message it calls this procedure with this parameter, the remaining messages are ignored. This message is sent just by the assembler thunk from the caller thread. Futher, a window is created and the handle of this window and the code heap are stored into the data of the window class. This is used to avoid a memory leak in the **IDE** because if the window class is registered once, then these parameters can be obtained in any debugging session. The callback function is generated in **InitCurrentThreadAndCallFunctionIDEProc**, but first it's checked whether the same callback procedure has already been created (in order to don't create the same thunk). The thunk has the following code:

```asm
LEA EAX, [ESP+4]
PUSH EAX
PUSH pfnCallback
PUSH WM_ONCALLBACK
PUSH hMsgWindow
Call SendMessageW
RETN lParametersSize
```

As you can see from the code, during calling a callback function, the call is transmitted via **SendMessage** to the main thread. The **lParametersSize** parameter is used to correctly restore the stack.

The next feature of the module is the creation of objects in a separate thread, and you can create them as private objects (the method is based on the code of [the NameBasedObjectFactory by firehacker module](http://bbs.vbstreets.ru/viewtopic.php?f=28&t=43201)) as public ones. To create the project classes use the **CreatePrivateObjectByNameInNewThread** function and for **ActiveX**-public classes **CreateActiveXObjectInNewThread** and **CreateActiveXObjectInNewThread2** ones. Before creating instances of the project classes you must first enable marshaling of these objects by calling the **EnablePrivateMarshaling** function. These functions accept the class identifier (**ProgID** / **CLSID** for **ActiveX** and the name for the project classes) and the interface identifier (**IDispatch** / **Object** is used by default). If the function is successfully called a marshaled object and an asynchronous call **ID** are returned. For the compiled version this is the **ID** of thread for **IDE** it's a pointer to the object. Objects are created and "live" in the **ActiveXThreadProc** function. The life of objects is controlled through the reference count (when it is equal to 1  it means only **ActiveXThreadProc** refers to the object and you can delete it and terminate the thread).
You can call the methods either synchronously - just call the method as usual or asynchronously - using the **AsynchDispMethodCall** procedure. This procedure takes an asynchronous call **ID**, a method name, a call type, an object that receives the call notification, a notification method name and the list of parameters. The procedure copies the parameters to the temporary memory, marshals the notification object, and sends the data to the object's thread via **WM_ASYNCH_CALL**. It should be noted that marshaling of parameters isn't supported right now therefore it's necessary to transfer links to objects with care. If you want to marshal an object reference you should use a synchronous method to marshal the objects and then call the asynchronous method. The procedure is returned immediately. In the **ActiveXThreadProc** thread the data is retrieved and a synchronous call is made via **MakeAsynchCall**. Everything is simple, **CallByName** is called for the thread object and **CallByName** for notification. The notification method has the following prototype:

```vb
Public Sub CallBack (ByVal vRet As Variant)
```

, where **vRet** accepts the return value of the method.

The following functions are intended for marshaling: **Marshal**, **Marshal2**, **UnMarshal**, **FreeMarshalData**. The first one creates information about the marshaling (**Proxy**) of the interface and puts it into the stream (**IStream**) that is returned. It accepts the interface identifier in the **pInterface** parameter (**IDispatch** / **Object** by default). The **UnMarshal** function, on the contrary, receives a stream and creates a **Proxy** object based on the information in the stream. Optionally, you can release the thread object. **Marshal2** does the same thing as **Marshal** except that it allows you to create a **Proxy** object many times in different threads. **FreeMarshalData** releases the data and the stream accordingly.
If, for example, you want to transfer a reference to an object between two threads, it is enough to call the **Marshal** / **UnMarshal** pair in the thread which created the object and in the thread that receives the link respectively. In another case, if for example there is the one global object and you need to pass a reference to it to the multiple threads (for example, the logging object), then **Marshal2** is called in the object thread, and **UnMarshal** with the **bReleaseStream** parameter is set to **False** is called in client threads. When the data is no longer needed, **FreeMarshalData** is called.

The **WaitForObjectThreadCompletion** function is designed to wait for the completion of the object thread and receives the **ID** of the asynchronous call. It is desirable to call this function always at the end of the main process because an object thread can somehow interact with the main thread and its objects (for example, if the object thread has a marshal link to the interface of the main thread).

The **SuspendResume** function is designed to suspend/resume the object's thread; **bSuspend** determines whether to sleep or resume the thread.

In addition, there are also several examples in the attacment of working with module:

- **Callback** - the project demonstrates the work with the **callback**-function periodically called in the different threads. Also, there is an additional project of native dll (on VB6) which calls the function periodically in the different threads;
- **JuliaSet** - the **Julia** fractal generation in the several threads (user-defined);
- **CopyProgress** - Copy the folder in a separate thread with the progress of the copy;
- **PublicMarshaling** - Creating public objects (**Dictionary**) in the different threads and calling their methods (synchronously / asynchronously);
- **PrivateMarshaling** - Creating private objects in different threads and calling their methods (synchronously / asynchronously);
- **MarshalUserInterface** - Creating private objects in different threads and calling their methods (synchronously / asynchronously) based on user interfaces (contains tlb and Reg-Free manifest).
- **InitProjectContextDll** - Initialization of the runtime in an ActiveX DLL and call the exported function from the different threads. Setup callbacks to the executable.
- **InternetStatusCallback** - IternetStatusCallback usage in VB6. Async file downloading.

The module is poorly tested so bugs are possible. I would be very glad to any bug-reports, wherever possible I will correct them.
Thank you all for attention!

Best Regards,

The trick.
