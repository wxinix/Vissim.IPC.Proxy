# vissim-ipcp
Vissim in-process COM proxy, BOOST Vissim performance x15 faster!

# About
Vissim is the most widely-used and industry microscopic traffic simulator developed by PTV. It provides a COM-based interface for user-customized applications, invoking Vissim as an out-of-process automation server.

Vissim COM is convenient and powerful, however, it can be slow especially for loop-intensive applications. Albeit the interface is well designed and comprehensive, it ONLY supports out-of-process call site (except its performance-constrained event-based COM scripting, which is in-process COM). When calling Vissim COM interface in a simstep-by-step fashion, the run time performance could possibly become unacceptable.

- The performance hit is mainly due to the COM marshaling from one program to the other. The performance hit can be further exacerbated when using .NET languages such as C#, or scripting langue such as Python or Matlab, because of .NET COM p/invoke , or the additional run-time overhead of the IDispatch interface.

So I made this "hack" to Vissim COM model. 

It implements a zero-overhead proxy that enables Vissim COM as in-process COM. In other words, we no longer have to invoke Vissim as an out-of-process automation server. Better, we now have access to all the functionalities (theoretically) in the same process space as the Vissim host. This will greatly improve the run time performance, which in my case, boosting the speed as much as ~x15 faster. Your mileage may vary, though.

# Vissim COM can be slow, how do we improve it?
This Vissim in-process COM proxy improves Vissim COM performance by making it possible to call Vissim COM interface in the same process space, for example, from inside Driver Model Dll, or from inside Signal Control Dll. Several unconventional yet quite clever "hacking tricks" have been employed (the source code reveals all the details):

- The Vissim Launcher will launch Vissim.exe while injecting a Dll, named VissimComHook.dll, as Operating System hook into Vissim process space

- The VissimComHook.dll will intercept Vissim's calls to Windows Operating System COM runtime (i.e., Ole.dll). This includes first creating a dummy empty COM script - which will "cheat" Vissim to construct an internal IVissim object when Vissim tries to establish a site for active scripting

- As soon as the site of active scripting is obtained ("hijacked") by VissimComHook.dll via hooking ("intercepting") Vissim call to Ole.dll (COM runtime), the site is forwarded to VissimInProcComProxy.dll, and cached there

- VissimInProcComProxy.dll itself is an in-process COM server sharing the same memory space with Vissim. It serves as a proxy of Vissim in-process COM, because it internally holds a copy of the aforementioned active scripting site. What it actually does, is quite simple - it just uses the cached active scripting site to do a simple QueryInterface to get the IUnknown interface. The IUnknown interface, points to Vissim's internal implementation object of IVissim

- The IVissim obtained this way, is a pure in-process handle to Vissim's internal COM object. Once we hunt it down, we can do the normal things we used to do with the standard Vissim COM interfaces. Imagination will be your limit! 

# Usage
This blog article provides a sample on how to use it with C++. http://blog.wupingxin.net/2018/04/12/dissecting-vissim-com-internal-from-inside-out-7-installation-and-a-sample-in-c/

# More readings

- http://blog.wupingxin.net/2018/06/10/advanced-vissim-com-programming-1-vissims-com-thread-model-and-instance-model/
- http://blog.wupingxin.net/2015/05/22/dissecting-vissim-com-internal-1/
- http://blog.wupingxin.net/2015/05/26/dissecting-vissim-com-internal-from-inside-out-2/
- http://blog.wupingxin.net/2015/06/05/dissecting-vissim-com-internal-from-inside-out-3/
- http://blog.wupingxin.net/2017/11/07/dissecting-vissim-com-internal-from-inside-out-4-a-loop-hole-for-a-perfect-solution/
- http://blog.wupingxin.net/2018/04/08/dissecting-vissim-com-internal-from-inside-out-5-a-binary-fix-up-for-vissim-in-process-com/
- http://blog.wupingxin.net/2018/04/10/dissecting-vissim-com-internal-from-inside-out-6-check-out-the-real-horsepower/
- http://blog.wupingxin.net/2018/04/12/dissecting-vissim-com-internal-from-inside-out-7-installation-and-a-sample-in-c/
- http://blog.wupingxin.net/2019/04/29/vissim-com-interface-design-a-pitfall-and-a-caveat/
