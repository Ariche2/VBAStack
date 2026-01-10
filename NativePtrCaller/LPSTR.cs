using System;
using System.Runtime.InteropServices;

namespace NativePtrCaller
{
    internal sealed class LPSTR : IDisposable
    {
        private IntPtr _ptr;
        private bool _disposed;

        public LPSTR(int capacity)
        {
            if (capacity < 0)
                throw new ArgumentOutOfRangeException(nameof(capacity));

            _ptr = Marshal.AllocHGlobal(capacity + 1);
        }

        public static implicit operator IntPtr(LPSTR safeLPSTR)
        {
            if (safeLPSTR == null)
                throw new ArgumentNullException(nameof(safeLPSTR));
            if (safeLPSTR._disposed)
                throw new ObjectDisposedException(nameof(LPSTR));

            return safeLPSTR._ptr;
        }

        public override string ToString()
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(LPSTR));
            if (_ptr == IntPtr.Zero)
                return string.Empty;

            return Marshal.PtrToStringAnsi(_ptr) ?? string.Empty;
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                if (_ptr != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(_ptr);
                    _ptr = IntPtr.Zero;
                }
                _disposed = true;
            }
        }
    }
}
