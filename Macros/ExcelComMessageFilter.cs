using System;
using System.Runtime.InteropServices;

namespace RibbonXmlEditor.Macros
{
    internal static class ExcelComMessageFilter
    {
        public static IDisposable Register()
        {
            var filter = new RetryMessageFilter();
            int result = CoRegisterMessageFilter(filter, out IMessageFilter previousFilter);
            Marshal.ThrowExceptionForHR(result);
            return new Registration(filter, previousFilter);
        }

        [DllImport("ole32.dll")]
        private static extern int CoRegisterMessageFilter(
            [MarshalAs(UnmanagedType.Interface)] IMessageFilter newFilter,
            out IMessageFilter oldFilter);

        [ComVisible(true)]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("00000016-0000-0000-C000-000000000046")]
        private interface IMessageFilter
        {
            [PreserveSig]
            int HandleInComingCall(int callType, IntPtr callerTask, int elapsedTicks, IntPtr interfaceInfo);

            [PreserveSig]
            int RetryRejectedCall(IntPtr calleeTask, int elapsedTicks, int rejectType);

            [PreserveSig]
            int MessagePending(IntPtr calleeTask, int elapsedTicks, int pendingType);
        }

        [ComVisible(true)]
        [ClassInterface(ClassInterfaceType.None)]
        private sealed class RetryMessageFilter : IMessageFilter
        {
            private const int ServerCallRejected = 1;
            private const int ServerCallRetryLater = 2;
            private const int RetryDelayMilliseconds = 250;
            private const int MaximumRetryMilliseconds = 30000;
            private const int PendingMessageWaitAndProcess = 2;

            public int HandleInComingCall(int callType, IntPtr callerTask, int elapsedTicks, IntPtr interfaceInfo)
            {
                return 0;
            }

            public int RetryRejectedCall(IntPtr calleeTask, int elapsedTicks, int rejectType)
            {
                if ((rejectType == ServerCallRejected || rejectType == ServerCallRetryLater)
                    && elapsedTicks < MaximumRetryMilliseconds)
                    return RetryDelayMilliseconds;

                return -1;
            }

            public int MessagePending(IntPtr calleeTask, int elapsedTicks, int pendingType)
            {
                return PendingMessageWaitAndProcess;
            }
        }

        private sealed class Registration : IDisposable
        {
            private readonly IMessageFilter _registeredFilter;
            private readonly IMessageFilter _previousFilter;
            private bool _disposed;

            public Registration(IMessageFilter registeredFilter, IMessageFilter previousFilter)
            {
                _registeredFilter = registeredFilter;
                _previousFilter = previousFilter;
            }

            public void Dispose()
            {
                if (_disposed)
                    return;

                _disposed = true;
                int result = CoRegisterMessageFilter(_previousFilter, out _);
                GC.KeepAlive(_registeredFilter);
                Marshal.ThrowExceptionForHR(result);
            }
        }
    }
}
