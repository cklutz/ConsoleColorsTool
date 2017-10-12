using System;
using System.ComponentModel;
using System.IO;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace ConsoleColorsClean
{
    public static class Shortcut
    {
        public static void RmProps(string path)
        {
            var lnk = new CoShellLink();
            var data = (IShellLinkDataList)lnk;
            var file = (IPersistFile)lnk;
            file.Load(path, (int)STGM.Read);
            data.RemoveDataBlock(CoShellLink.NT_CONSOLE_PROPS_SIG);
            file.Save(path, true);
            Marshal.ReleaseComObject(data);
            Marshal.ReleaseComObject(file);
            Marshal.ReleaseComObject(lnk);
        }

        private static void DumpMember<T>(TextWriter tw, NT_CONSOLE_PROPS props, Expression<Func<NT_CONSOLE_PROPS, T>> expression,
            object value = null)
        {
            var name = ((FieldInfo)((MemberExpression)expression.Body).Member).Name;
            var func = expression.Compile();

            if (value != null)
            {
                tw.WriteLine("{0}: {1}", name, value);
            }
            else
            {
                tw.WriteLine("{0}: {1}", name, func(props));
            }
        }

        private static string GetString(ushort[] value)
        {
            int pos = Array.IndexOf(value, '\0');

            byte[] asBytes = new byte[(pos + 1) * sizeof(ushort)];
            Buffer.BlockCopy(value, 0, asBytes, 0, asBytes.Length);
            return Encoding.Unicode.GetString(asBytes);
        }

        public static void DumpConsoleInfo(TextWriter tw, string path)
        {
            var lnk = new CoShellLink();
            var data = (IShellLinkDataList)lnk;
            var file = (IPersistFile)lnk;
            file.Load(path, (int)STGM.Read);

            IntPtr block;
            int rc = data.CopyDataBlock(CoShellLink.NT_CONSOLE_PROPS_SIG, out block);
            if (rc != 0)
                throw new Win32Exception("CopyDataBlock failed. " + rc.ToString("X"));

            int size = Marshal.ReadInt32(block);
            var res = Marshal.PtrToStructure<NT_CONSOLE_PROPS>(block);

            tw.WriteLine("Flags: {0}", GetFlags(path));

            DumpMember(tw, res, p => p.bAutoPosition);
            DumpMember(tw, res, p => p.bFullScreen);
            DumpMember(tw, res, p => p.bHistoryNoDup);
            DumpMember(tw, res, p => p.bInsertMode);
            DumpMember(tw, res, p => p.bQuickEdit);
            DumpMember(tw, res, p => p.nInputBufferSize);
            DumpMember(tw, res, p => p.uCursorSize);
            DumpMember(tw, res, p => p.FaceName, GetString(res.FaceName));
            DumpMember(tw, res, p => p.nFont);
            DumpMember(tw, res, p => p.uFontFamily);
            DumpMember(tw, res, p => p.uFontWeight);
            DumpMember(tw, res, p => p.dwFontSize);
            DumpMember(tw, res, p => p.uHistoryBufferSize);
            DumpMember(tw, res, p => p.uNumberOfHistoryBuffers);
            DumpMember(tw, res, p => p.wFilleAttribute);
            DumpMember(tw, res, p => p.wPopupFillAttribute);
            DumpMember(tw, res, p => p.dwScreenBufferSize);
            DumpMember(tw, res, p => p.dwWindowOrigin);
            DumpMember(tw, res, p => p.dwWindowSize);

            for (var i = 0; i < res.ColorTable.Length; i++)
            {
                tw.WriteLine("ColorTable[{0}]: RGB({1},{2},{3}), #{1:X}{2:X}{3:X}", i,
                    res.ColorTable[i].R, res.ColorTable[i].G, res.ColorTable[i].B);
            }

            Marshal.FreeHGlobal(block);
            Marshal.ReleaseComObject(data);
            Marshal.ReleaseComObject(file);
            Marshal.ReleaseComObject(lnk);
        }

        public static string GetFlags(string path)
        {
            var lnk = new CoShellLink();
            var data = (IShellLinkDataList)lnk;
            var file = (IPersistFile)lnk;
            file.Load(path, (int)STGM.Read);

            var flags = data.GetFlags();

            Marshal.ReleaseComObject(data);
            Marshal.ReleaseComObject(file);
            Marshal.ReleaseComObject(lnk);

            return flags.ToString();
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct COORD
    {
        public short X;
        public short Y;

        public override string ToString()
        {
            return $"({X}, {Y})";
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct NT_CONSOLE_PROPS
    {
        public DATABLOCKHEADER Headedbh;
        public ushort wFilleAttribute;
        public ushort wPopupFillAttribute;
        public COORD dwScreenBufferSize;
        public COORD dwWindowSize;
        public COORD dwWindowOrigin;
        public UInt32 nFont;
        public UInt32 nInputBufferSize;
        public COORD dwFontSize;
        public UInt32 uFontFamily;
        public UInt32 uFontWeight;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 32)]
        public ushort[] FaceName;
        public UInt32 uCursorSize;
        public bool bFullScreen;
        public bool bQuickEdit;
        public bool bInsertMode;
        public bool bAutoPosition;
        public UInt32 uHistoryBufferSize;
        public UInt32 uNumberOfHistoryBuffers;
        public bool bHistoryNoDup;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 16)]
        public COLORREF[] ColorTable;
    }

    [StructLayout(LayoutKind.Explicit, Size = 4)]
    public struct COLORREF
    {
        public COLORREF(byte r, byte g, byte b)
        {
            Value = 0;
            R = r;
            G = g;
            B = b;
        }

        public COLORREF(uint value)
        {
            R = 0;
            G = 0;
            B = 0;
            Value = value & 0x00FFFFFF;
        }

        [FieldOffset(0)]
        public byte R;

        [FieldOffset(1)]
        public byte G;

        [FieldOffset(2)]
        public byte B;

        [FieldOffset(0)]
        public uint Value;

        public override string ToString()
        {
            return string.Format("#{0:x2}{1:x2}{2:x2}", R, G, B);
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct DATABLOCKHEADER
    {
        public int cbSize;
        public uint dwSignature;
    }

    // Indicate conditions for creating and deleting the object and access modes for the object
    [Flags]
    public enum STGM
    {
        Direct = 0x00000000,
        Transacted = 0x00010000,
        Simple = 0x08000000,
        Read = 0x00000000,
        Write = 0x00000001,
        ReadWrite = 0x00000002,
        ShareDenyNone = 0x00000040,
        ShareDenyRead = 0x00000030,
        ShareDenyWrite = 0x00000020,
        ShareExclusive = 0x00000010,
        Priority = 0x00040000,
        DeleteOnRelease = 0x04000000,
        Noscratch = 0x00100000,
        Create = 0x00001000,
        Convert = 0x00020000,
        FailIfThere = 0x00000000,
        NoSnapsHot = 0x00200000,
        DirectSwmr = 0x00400000,
    }

    [Flags]
    public enum SHELL_LINK_DATA_FLAGS : uint
    {
        SLDF_DISABLE_KNOWNFOLDER_RELATIVE_TRACKING = 0x200000,
        SLDF_ENABLE_TARGET_METADATA = 0x80000,
        SLDF_FORCE_NO_LINKINFO = 0x100,
        SLDF_FORCE_NO_LINKTRACK = 0x40000,
        SLDF_FORCE_UNCNAME = 0x10000,
        SLDF_HAS_ARGS = 0x20,
        SLDF_HAS_DARWINID = 0x1000,
        SLDF_HAS_EXP_ICON_SZ = 0x4000,
        SLDF_HAS_EXP_SZ = 0x200,
        SLDF_HAS_ICONLOCATION = 0x40,
        SLDF_HAS_ID_LIST = 1,
        SLDF_HAS_LINK_INFO = 2,
        SLDF_HAS_LOGO3ID = 0x800,
        SLDF_HAS_NAME = 4,
        SLDF_HAS_RELPATH = 8,
        SLDF_HAS_WORKINGDIR = 0x10,
        SLDF_NO_PIDL_ALIAS = 0x8000,
        SLDF_RESERVED = 0x80000000,
        SLDF_RUN_IN_SEPARATE = 0x400,
        SLDF_RUN_WITH_SHIMLAYER = 0x20000,
        SLDF_RUNAS_USER = 0x2000,
        SLDF_UNICODE = 0x80,
        SLDF_VALID = 0x3ff7ff
    }

    [ComImport, ClassInterface(ClassInterfaceType.None), Guid("00021401-0000-0000-C000-000000000046")]
    public class CoShellLink
    {
        public const uint EXP_DARWIN_ID_SIG = 0xa0000006;
        public const uint EXP_LOGO3_ID_SIG = 0xa0000007;
        public const uint EXP_SPECIAL_FOLDER_SIG = 0xa0000005;
        public const uint EXP_SZ_ICON_SIG = 0xa0000007;
        public const uint EXP_SZ_LINK_SIG = 0xa0000001;
        public const uint NT_CONSOLE_PROPS_SIG = 0xa0000002;
        public const uint NT_FE_CONSOLE_PROPS_SIG = 0xa0000004;
    }

    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("45e2b4ae-b1c3-11d0-b92f-00a0c90312e1")]
    interface IShellLinkDataList
    {
        void AddDataBlock(IntPtr pDataBlock);
        [PreserveSig]
        int CopyDataBlock(uint dwSig, out IntPtr ppDataBlock);
        void RemoveDataBlock(uint dwSig);
        SHELL_LINK_DATA_FLAGS GetFlags();
        void SetFlags(SHELL_LINK_DATA_FLAGS dwFlags);
    }
}