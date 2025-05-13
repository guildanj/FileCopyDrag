using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
namespace Downtime_Tool.Packages
{
    public static class FileCopyDrag
    {
        #region "  NativeMethods"
        internal static class NativeMethods
        {
            [DllImport("kernel32.dll")]
            static extern IntPtr GlobalLock(IntPtr hMem);

            [DllImport("ole32.dll", PreserveSig = false)]
            public static extern ILockBytes CreateILockBytesOnHGlobal(IntPtr hGlobal, bool fDeleteOnRelease);

            [DllImport("OLE32.DLL", CharSet = CharSet.Auto, PreserveSig = false)]
            public static extern IntPtr GetHGlobalFromILockBytes(ILockBytes pLockBytes);

            [DllImport("OLE32.DLL", CharSet = CharSet.Unicode, PreserveSig = false)]
            public static extern IStorage StgCreateDocfileOnILockBytes(ILockBytes plkbyt, uint grfMode, uint reserved);

            [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("0000000B-0000-0000-C000-000000000046")]
            public interface IStorage
            {
                [return: MarshalAs(UnmanagedType.Interface)]
                IStream CreateStream([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, [In, MarshalAs(UnmanagedType.U4)] int grfMode, [In, MarshalAs(UnmanagedType.U4)] int reserved1, [In, MarshalAs(UnmanagedType.U4)] int reserved2);
                [return: MarshalAs(UnmanagedType.Interface)]
                IStream OpenStream([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, IntPtr reserved1, [In, MarshalAs(UnmanagedType.U4)] int grfMode, [In, MarshalAs(UnmanagedType.U4)] int reserved2);
                [return: MarshalAs(UnmanagedType.Interface)]
                IStorage CreateStorage([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, [In, MarshalAs(UnmanagedType.U4)] int grfMode, [In, MarshalAs(UnmanagedType.U4)] int reserved1, [In, MarshalAs(UnmanagedType.U4)] int reserved2);
                [return: MarshalAs(UnmanagedType.Interface)]
                IStorage OpenStorage([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, IntPtr pstgPriority, [In, MarshalAs(UnmanagedType.U4)] int grfMode, IntPtr snbExclude, [In, MarshalAs(UnmanagedType.U4)] int reserved);
                void CopyTo(int ciidExclude, [In, MarshalAs(UnmanagedType.LPArray)] Guid[] pIIDExclude, IntPtr snbExclude, [In, MarshalAs(UnmanagedType.Interface)] IStorage stgDest);
                void MoveElementTo([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, [In, MarshalAs(UnmanagedType.Interface)] IStorage stgDest, [In, MarshalAs(UnmanagedType.BStr)] string pwcsNewName, [In, MarshalAs(UnmanagedType.U4)] int grfFlags);
                void Commit(int grfCommitFlags);
                void Revert();
                void EnumElements([In, MarshalAs(UnmanagedType.U4)] int reserved1, IntPtr reserved2, [In, MarshalAs(UnmanagedType.U4)] int reserved3, [MarshalAs(UnmanagedType.Interface)] out object ppVal);
                void DestroyElement([In, MarshalAs(UnmanagedType.BStr)] string pwcsName);
                void RenameElement([In, MarshalAs(UnmanagedType.BStr)] string pwcsOldName, [In, MarshalAs(UnmanagedType.BStr)] string pwcsNewName);
                void SetElementTimes([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, [In] System.Runtime.InteropServices.ComTypes.FILETIME pctime, [In] System.Runtime.InteropServices.ComTypes.FILETIME patime, [In] System.Runtime.InteropServices.ComTypes.FILETIME pmtime);
                void SetClass([In] ref Guid clsid);
                void SetStateBits(int grfStateBits, int grfMask);
                void Stat([Out] out System.Runtime.InteropServices.ComTypes.STATSTG pStatStg, int grfStatFlag);
            }

            [ComImport, Guid("0000000A-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            public interface ILockBytes
            {
                void ReadAt([In, MarshalAs(UnmanagedType.U8)] long ulOffset, [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] byte[] pv, [In, MarshalAs(UnmanagedType.U4)] int cb, [Out, MarshalAs(UnmanagedType.LPArray)] int[] pcbRead);
                void WriteAt([In, MarshalAs(UnmanagedType.U8)] long ulOffset, IntPtr pv, [In, MarshalAs(UnmanagedType.U4)] int cb, [Out, MarshalAs(UnmanagedType.LPArray)] int[] pcbWritten);
                void Flush();
                void SetSize([In, MarshalAs(UnmanagedType.U8)] long cb);
                void LockRegion([In, MarshalAs(UnmanagedType.U8)] long libOffset, [In, MarshalAs(UnmanagedType.U8)] long cb, [In, MarshalAs(UnmanagedType.U4)] int dwLockType);
                void UnlockRegion([In, MarshalAs(UnmanagedType.U8)] long libOffset, [In, MarshalAs(UnmanagedType.U8)] long cb, [In, MarshalAs(UnmanagedType.U4)] int dwLockType);
                void Stat([Out] out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg, [In, MarshalAs(UnmanagedType.U4)] int grfStatFlag);
            }

            [StructLayout(LayoutKind.Sequential)]
            public sealed class POINTL
            {
                public int x;
                public int y;
            }

            [StructLayout(LayoutKind.Sequential)]
            public sealed class SIZEL
            {
                public int cx;
                public int cy;
            }

            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
            public sealed class FILEGROUPDESCRIPTORA
            {
                public uint cItems;
                public required FILEDESCRIPTORA[] fgd;
            }

            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
            public sealed class FILEDESCRIPTORA
            {
                public uint dwFlags;
                public Guid clsid;
                public required SIZEL sizel;
                public required POINTL pointl;
                public uint dwFileAttributes;
                public System.Runtime.InteropServices.ComTypes.FILETIME ftCreationTime;
                public System.Runtime.InteropServices.ComTypes.FILETIME ftLastAccessTime;
                public System.Runtime.InteropServices.ComTypes.FILETIME ftLastWriteTime;
                public uint nFileSizeHigh;
                public uint nFileSizeLow;
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
                public required string cFileName;
            }

            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
            public sealed class FILEGROUPDESCRIPTORW
            {
                public uint cItems;
                public required FILEDESCRIPTORW[] fgd;
            }

            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
            public sealed class FILEDESCRIPTORW
            {
                public uint dwFlags;
                public Guid clsid;
                public required SIZEL sizel;
                public required POINTL pointl;
                public uint dwFileAttributes;
                public System.Runtime.InteropServices.ComTypes.FILETIME ftCreationTime;
                public System.Runtime.InteropServices.ComTypes.FILETIME ftLastAccessTime;
                public System.Runtime.InteropServices.ComTypes.FILETIME ftLastWriteTime;
                public uint nFileSizeHigh;
                public uint nFileSizeLow;
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
                public required string cFileName;
            }
        }
        #endregion

        #region "  FileDescriptorReader" 
        [Flags]
        enum FileDescriptorFlags : uint
        {
            ClsId = 0x00000001,
            SizePoint = 0x00000002,
            Attributes = 0x00000004,
            CreateTime = 0x00000008,
            AccessTime = 0x00000010,
            WritesTime = 0x00000020,
            FileSize = 0x00000040,
            ProgressUI = 0x00004000,
            LinkUI = 0x00008000,
            Unicode = 0x80000000,
        }
        static class FileDescriptorReader
        {
            internal sealed class FileDescriptor
            {
                public FileDescriptorFlags Flags { get; set; }
                public Guid ClassId { get; set; }
                public Size Size { get; set; }
                public Point Point { get; set; }
                public FileAttributes FileAttributes { get; set; }
                public DateTime CreationTime { get; set; }
                public DateTime LastAccessTime { get; set; }
                public DateTime LastWriteTime { get; set; }
                public Int64 FileSize { get; set; }
                public string FileName { get; set; }

                public FileDescriptor(BinaryReader reader)
                {
                    //Flags
                    Flags = (FileDescriptorFlags)reader.ReadUInt32();
                    //ClassID
                    ClassId = new Guid(reader.ReadBytes(16));
                    //Size
                    Size = new Size(reader.ReadInt32(), reader.ReadInt32());
                    //Point
                    Point = new Point(reader.ReadInt32(), reader.ReadInt32());
                    //FileAttributes
                    FileAttributes = (FileAttributes)reader.ReadUInt32();
                    //CreationTime
                    CreationTime = new DateTime(1601, 1, 1).AddTicks(reader.ReadInt64());
                    //LastAccessTime
                    LastAccessTime = new DateTime(1601, 1, 1).AddTicks(reader.ReadInt64());
                    //LastWriteTime
                    LastWriteTime = new DateTime(1601, 1, 1).AddTicks(reader.ReadInt64());
                    //FileSize
                    FileSize = reader.ReadInt64();
                    //FileName
                    byte[] nameBytes = reader.ReadBytes(520);
                    int i = 0;
                    while (i < nameBytes.Length)
                    {
                        if (nameBytes[i] == 0 && nameBytes[i + 1] == 0)
                            break;
                        i++;
                        i++;
                    }
                    FileName = UnicodeEncoding.Unicode.GetString(nameBytes, 0, i);
                }
            }

            public static IEnumerable<FileDescriptor> Read(Stream fileDescriptorStream)
            {
                BinaryReader reader = new(fileDescriptorStream);
                var count = reader.ReadUInt32();

                while (count > 0)
                {
                    FileDescriptor descriptor = new(reader);

                    yield return descriptor;

                    count--;
                }
            }

            public static IEnumerable<string> ReadFileNames(Stream fileDescriptorStream)
            {
                BinaryReader reader = new(fileDescriptorStream);
                var count = reader.ReadUInt32();
                while (count > 0)
                {
                    FileDescriptor descriptor = new(reader);

                    yield return descriptor.FileName;

                    count--;
                }
            }
        }
        #endregion

        #region "  Private Members"
        private static MemoryStream? GetFileContents(System.Windows.Forms.IDataObject dataObject, int index)
        {
            //cast the default IDataObject to a com IDataObject
            System.Runtime.InteropServices.ComTypes.IDataObject comDataObject;
            comDataObject = (System.Runtime.InteropServices.ComTypes.IDataObject)dataObject;

            var Format = System.Windows.Forms.DataFormats.GetFormat("FileContents");
            if (Format == null) return null;

            FORMATETC formatetc = new()
            {
                cfFormat = (short)Format.Id,
                dwAspect = DVASPECT.DVASPECT_CONTENT,
                lindex = index,
                ptd = new IntPtr(0),
                tymed = TYMED.TYMED_ISTREAM | TYMED.TYMED_ISTORAGE | TYMED.TYMED_HGLOBAL
            };


            //create STGMEDIUM to output request results into
            STGMEDIUM medium = new();

            //using the com IDataObject interface get the data using the defined FORMATETC
            comDataObject.GetData(ref formatetc, out medium);

            return medium.tymed switch
            {
                TYMED.TYMED_ISTREAM => GetIStream(medium),
                TYMED.TYMED_ISTORAGE => GetIStorage(medium),
                _ => throw new NotSupportedException("Not Supported:" + medium.tymed.ToString()),
            };
        }

        private static MemoryStream GetIStorage(STGMEDIUM medium)
        {
            NativeMethods.IStorage? iStorage = null;
            NativeMethods.IStorage? iStorage2 = null;
            NativeMethods.ILockBytes? iLockBytes = null;
            System.Runtime.InteropServices.ComTypes.STATSTG iLockBytesStat;
            try
            {
                //marshal the returned pointer to a IStorage object
                iStorage = (NativeMethods.IStorage)Marshal.GetObjectForIUnknown(medium.unionmember);
                Marshal.Release(medium.unionmember);

                //create a ILockBytes (unmanaged byte array) and then create a IStorage using the byte array as a backing store
                iLockBytes = NativeMethods.CreateILockBytesOnHGlobal(IntPtr.Zero, true);
                iStorage2 = NativeMethods.StgCreateDocfileOnILockBytes(iLockBytes, 0x00001012, 0);

                //copy the returned IStorage into the new IStorage
                iStorage.CopyTo(0, null!, IntPtr.Zero, iStorage2);
                iLockBytes.Flush();
                iStorage2.Commit(0);

                //get the STATSTG of the ILockBytes to determine how many bytes were written to it
                iLockBytesStat = new System.Runtime.InteropServices.ComTypes.STATSTG();
                iLockBytes.Stat(out iLockBytesStat, 1);
                int iLockBytesSize = (int)iLockBytesStat.cbSize;

                //read the data from the ILockBytes (unmanaged byte array) into a managed byte array
                byte[] iLockBytesContent = new byte[iLockBytesSize];
                iLockBytes.ReadAt(0, iLockBytesContent, iLockBytesContent.Length, null!);

                //wrapped the managed byte array into a memory stream and return it
                return new MemoryStream(iLockBytesContent);
            }
            finally
            {
                //release all unmanaged objects
                Marshal.ReleaseComObject(iStorage2!);
                Marshal.ReleaseComObject(iLockBytes!);
                Marshal.ReleaseComObject(iStorage!);
            }
        }

        private static MemoryStream GetIStream(STGMEDIUM medium)
        {
            //marshal the returned pointer to a IStream object
            IStream iStream = (IStream)Marshal.GetObjectForIUnknown(medium.unionmember);
            Marshal.Release(medium.unionmember);

            //get the STATSTG of the IStream to determine how many bytes are in it
            var iStreamStat = new System.Runtime.InteropServices.ComTypes.STATSTG();
            iStream.Stat(out iStreamStat, 0);
            int iStreamSize = (int)iStreamStat.cbSize;

            //read the data from the IStream into a managed byte array
            byte[] iStreamContent = new byte[iStreamSize];
            iStream.Read(iStreamContent, iStreamContent.Length, IntPtr.Zero);

            //wrapped the managed byte array into a memory stream
            return new MemoryStream(iStreamContent);
        }
        #endregion

        private static List<string> Cache = [];
        private static bool EnableCache = false;

        public static void ClearCache()
        {
            foreach (var filename in Cache)
            {
                if (File.Exists(filename))
                {
                    try
                    {
                        File.Delete(filename);
                    }
                    catch (Exception)
                    {
                    }
                }
            }
        }

        public static string[]? PasteFromClipboard()
        {
            if (Clipboard.GetDataObject()?.GetData("FileGroupDescriptorW") is MemoryStream descriptor)
            {
                List<string> filenames = [];
                var files = FileDescriptorReader.Read(descriptor);
                var fileIndex = 0;

                foreach (var fileContentFile in files)
                {
                    if ((fileContentFile.FileAttributes & FileAttributes.Directory) == 0)
                    {
                        if (Clipboard.GetDataObject() is System.Windows.Forms.IDataObject dataObject)
                        {
                            var fileData = GetFileContents(dataObject, fileIndex);
                            if (fileData != null)
                            {
                                string filePath = Path.Combine(Path.GetTempPath(), fileContentFile.FileName);
                                using (var fileStream = new FileStream(filePath, FileMode.Create))
                                {
                                    fileData.CopyTo(fileStream);
                                }

                                if (EnableCache) Cache.Add(filePath);
                                filenames.Add(filePath);
                            }
                        }
                    }

                    fileIndex++;
                }

                if (filenames.Count > 0) return filenames.ToArray();
            }

            return null;
        }

        public static string[]? DragDrop(System.Windows.Forms.IDataObject dataObject)
        {
            if (dataObject.GetDataPresent(DataFormats.FileDrop, false))
            {
                if (dataObject.GetData(DataFormats.FileDrop) is string[] filenames && filenames.Length > 0)
                {
                    return filenames;
                }
            }
            else if (dataObject.GetDataPresent("FileGroupDescriptorW"))
            {
                if (dataObject.GetData("FileGroupDescriptorW") is MemoryStream descriptor)
                {
                    List<string> filenames = [];
                    var files = FileDescriptorReader.Read(descriptor);
                    var fileIndex = 0;

                    foreach (var fileContentFile in files)
                    {
                        if ((fileContentFile.FileAttributes & FileAttributes.Directory) == 0)
                        {
                            var fileData = GetFileContents(dataObject, fileIndex);
                            if (fileData != null)
                            {
                                string filePath = Path.Combine(Path.GetTempPath(), fileContentFile.FileName);
                                using (var fileStream = new FileStream(filePath, FileMode.Create))
                                {
                                    fileData.CopyTo(fileStream);
                                }

                                if (EnableCache) Cache.Add(filePath);
                                filenames.Add(filePath);
                            }
                        }

                        fileIndex++;
                    }

                    if (filenames.Count > 0) return filenames.ToArray();
                }
            }

            return null;
        }
    }
}
