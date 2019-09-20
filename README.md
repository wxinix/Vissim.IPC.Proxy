# vissim-ipcp
Vissim in-process COM proxy, BOOST Vissim performance x15 faster!

# About
Vissim is a widely-used microscopic traffic simulator developed by PTV.  It provides a COM-based interface for user-customized applications invoking Vissim as an out-of-process automation server.

Vissim COM is convenient, but slow, especially for loop-intensive applications.  Albeit the interfaces are well designed and comprehensive, it ONLY supports out-of-process call site.  If you need to call Vissim COM interfaces on a simulation step-by-step basis, the run time performance could become unacceptable (at least, to me).

So I made this "hack" to Vissim COM model. It implements a zero-overhead proxy that enables Vissim COM interface as an in-process COM server. In other words, you no longer have to invoke Vissim as an out-of-process automation server in order to access those COM functionalties.  With this proxy, you have access to all the functionalites in the same process space as the Vissim host. This will greatly improve the run time performance,  which in my case,  boost the speed as much as x15 faster. Your mileage may vary.

# Vissim COM is slow, how do we improve it?
This Vissim in-process COM proxy improves Vissim COM performance by calling Vissim COM interfaces in the same process space, rather than invoking Vissim as an out-of-process automation server.

- The Vissim Launcher will launch Vissim.exe while injecting a Dll hook, called VissimComHook.dll into Vissim process space
- The VissimComHook.dll will intercept Vissim's calls to Windows System COM library (i.e., Ole.dll). When we creates a dummy empty COM script, Vissim internally will create IVissim interface via Microsoft IActiveScript interface
-  Cache this interface and transform it to IUnknown and return it upon request
- VissimInProcComProxy.dll is a in-process COM server. It serves as a proxy of Vissim in-process COM since it returns the IUnknown of IVissim it caches

#Usage
Follow the samples in this blog article on the usage. http://blog.wupingxin.net/2018/04/12/dissecting-vissim-com-internal-from-inside-out-7-installation-and-a-sample-in-c/

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
