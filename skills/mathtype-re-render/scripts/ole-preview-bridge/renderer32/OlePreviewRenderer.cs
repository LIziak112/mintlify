using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

internal static class OlePreviewRenderer
{
    private const string BridgeVersion = "5";
    private const int STGM_READWRITE = 0x00000002;
    private const int STGM_SHARE_EXCLUSIVE = 0x00000010;
    private const int DVASPECT_CONTENT = 1;
    private const int CF_METAFILEPICT = 3;

    private static readonly Guid IID_IDataObject = new Guid("0000010E-0000-0000-C000-000000000046");

    [ComImport]
    [Guid("00000112-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IOleObject
    {
        [PreserveSig]
        int SetClientSite(IntPtr clientSite);

        [PreserveSig]
        int GetClientSite(out IntPtr clientSite);

        [PreserveSig]
        int SetHostNames(
            [MarshalAs(UnmanagedType.LPWStr)] string containerApp,
            [MarshalAs(UnmanagedType.LPWStr)] string containerObject);

        [PreserveSig]
        int Close(uint saveOption);

        [PreserveSig]
        int SetMoniker(uint whichMoniker, IntPtr moniker);

        [PreserveSig]
        int GetMoniker(uint assign, uint whichMoniker, out IntPtr moniker);

        [PreserveSig]
        int InitFromData(IntPtr dataObject, [MarshalAs(UnmanagedType.Bool)] bool creation, uint reserved);

        [PreserveSig]
        int GetClipboardData(uint reserved, out IntPtr dataObject);

        [PreserveSig]
        int DoVerb(
            int verb,
            IntPtr message,
            IntPtr activeSite,
            int index,
            IntPtr parentWindow,
            IntPtr positionRectangle);

        [PreserveSig]
        int EnumVerbs(out IntPtr enumerator);

        [PreserveSig]
        int Update();
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MetaFilePict
    {
        public int MappingMode;
        public int XExt;
        public int YExt;
        public IntPtr MetaFile;
    }

    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int StgOpenStorage(
        string name,
        IntPtr priority,
        int mode,
        IntPtr exclude,
        int reserved,
        out IntPtr storage);

    [DllImport("ole32.dll")]
    private static extern int OleLoad(
        IntPtr storage,
        ref Guid iid,
        IntPtr clientSite,
        out IntPtr result);

    [DllImport("ole32.dll")]
    private static extern int OleRun(IntPtr unknown);

    [DllImport("ole32.dll")]
    private static extern int OleInitialize(IntPtr reserved);

    [DllImport("ole32.dll")]
    private static extern void OleUninitialize();

    [DllImport("ole32.dll")]
    private static extern void ReleaseStgMedium(ref STGMEDIUM medium);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern ushort RegisterClipboardFormat(string format);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClipboardFormatName(uint format, System.Text.StringBuilder name, int maxCount);

    [DllImport("kernel32.dll")]
    private static extern IntPtr GlobalLock(IntPtr memory);

    [DllImport("kernel32.dll")]
    private static extern bool GlobalUnlock(IntPtr memory);

    [DllImport("kernel32.dll")]
    private static extern UIntPtr GlobalSize(IntPtr memory);

    [DllImport("gdi32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr CopyMetaFile(IntPtr source, string fileName);

    [DllImport("gdi32.dll")]
    private static extern bool DeleteMetaFile(IntPtr metaFile);

    private static void ThrowIfFailed(int hr, string operation)
    {
        if (hr < 0)
        {
            throw new COMException(operation + " failed", hr);
        }
    }

    private static STGMEDIUM GetData(IDataObject dataObject, short format, TYMED mediumType)
    {
        var formatEtc = new FORMATETC
        {
            cfFormat = format,
            dwAspect = (DVASPECT)DVASPECT_CONTENT,
            lindex = -1,
            ptd = IntPtr.Zero,
            tymed = mediumType
        };
        STGMEDIUM medium;
        dataObject.GetData(ref formatEtc, out medium);
        return medium;
    }

    private static MetaFilePict ExportWmf(IDataObject dataObject, string outputPath)
    {
        var medium = GetData(dataObject, CF_METAFILEPICT, TYMED.TYMED_MFPICT);
        try
        {
            IntPtr locked = GlobalLock(medium.unionmember);
            if (locked == IntPtr.Zero) throw new InvalidOperationException("GlobalLock(CF_METAFILEPICT) failed");
            MetaFilePict picture;
            try
            {
                picture = (MetaFilePict)Marshal.PtrToStructure(locked, typeof(MetaFilePict));
            }
            finally
            {
                GlobalUnlock(medium.unionmember);
            }

            string rawPath = outputPath + ".raw";
            IntPtr copy = CopyMetaFile(picture.MetaFile, rawPath);
            if (copy == IntPtr.Zero) throw new InvalidOperationException("CopyMetaFile failed");
            DeleteMetaFile(copy);
            byte[] raw = File.ReadAllBytes(rawPath);
            File.Delete(rawPath);
            WritePlaceableWmf(outputPath, raw, picture.XExt, picture.YExt);
            return picture;
        }
        finally
        {
            ReleaseStgMedium(ref medium);
        }
    }

    private static void WriteUInt16(BinaryWriter writer, ushort value)
    {
        writer.Write(value);
    }

    private static void WritePlaceableWmf(string outputPath, byte[] rawWmf, int xExt, int yExt)
    {
        const uint key = 0x9AC6CDD7;
        const ushort unitsPerInch = 2540; // METAFILEPICT extents use 0.01 mm.
        short right = checked((short)Math.Abs(xExt));
        short bottom = checked((short)Math.Abs(yExt));

        using (var headerStream = new MemoryStream())
        using (var writer = new BinaryWriter(headerStream))
        {
            writer.Write(key);
            WriteUInt16(writer, 0); // hmf
            writer.Write((short)0); // left
            writer.Write((short)0); // top
            writer.Write(right);
            writer.Write(bottom);
            WriteUInt16(writer, unitsPerInch);
            writer.Write((uint)0);
            writer.Flush();

            byte[] first20 = headerStream.ToArray();
            ushort checksum = 0;
            for (int i = 0; i < 20; i += 2)
                checksum ^= BitConverter.ToUInt16(first20, i);

            using (var output = File.Create(outputPath))
            {
                output.Write(first20, 0, first20.Length);
                output.Write(BitConverter.GetBytes(checksum), 0, 2);
                output.Write(rawWmf, 0, rawWmf.Length);
            }
        }
    }

    private static string ClipboardFormatName(short format)
    {
        int value = unchecked((ushort)format);
        if (value == CF_METAFILEPICT) return "CF_METAFILEPICT";
        if (value < 0xC000) return "CF_" + value;
        var buffer = new System.Text.StringBuilder(256);
        int length = GetClipboardFormatName((uint)value, buffer, buffer.Capacity);
        return length > 0 ? buffer.ToString() : "REGISTERED_" + value;
    }

    private static void TraceFormats(IDataObject dataObject, Action<string> trace)
    {
        IEnumFORMATETC enumerator = dataObject.EnumFormatEtc(DATADIR.DATADIR_GET);
        var items = new FORMATETC[1];
        int[] fetched = new int[1];
        while (enumerator.Next(1, items, fetched) == 0 && fetched[0] == 1)
        {
            FORMATETC item = items[0];
            trace(
                "format name=" + ClipboardFormatName(item.cfFormat)
                + " id=" + unchecked((ushort)item.cfFormat)
                + " tymed=" + item.tymed
                + " aspect=" + item.dwAspect
                + " lindex=" + item.lindex);
        }
    }

    private static string DecodePath(string value)
    {
        return Encoding.UTF8.GetString(Convert.FromBase64String(value));
    }

    private static string EncodeText(string value)
    {
        return Convert.ToBase64String(Encoding.UTF8.GetBytes(value));
    }

    private static MetaFilePict RenderOne(string inputPath, string outputPath, string dimensionsPath)
    {
        string input = Path.GetFullPath(inputPath);
        string outputWmf = Path.GetFullPath(outputPath);
        string outputDims = Path.GetFullPath(dimensionsPath);
        string outputDirectory = Path.GetDirectoryName(outputWmf);
        string dimensionsDirectory = Path.GetDirectoryName(outputDims);
        if (!File.Exists(input)) throw new FileNotFoundException("OLE input does not exist", input);
        if (!String.IsNullOrEmpty(outputDirectory)) Directory.CreateDirectory(outputDirectory);
        if (!String.IsNullOrEmpty(dimensionsDirectory)) Directory.CreateDirectory(dimensionsDirectory);

        IntPtr storage = IntPtr.Zero;
        IntPtr dataObjectPointer = IntPtr.Zero;
        IDataObject dataObject = null;
        IOleObject oleObject = null;
        try
        {
            int hr = StgOpenStorage(input, IntPtr.Zero, STGM_READWRITE | STGM_SHARE_EXCLUSIVE, IntPtr.Zero, 0, out storage);
            ThrowIfFailed(hr, "StgOpenStorage");

            Guid iid = IID_IDataObject;
            hr = OleLoad(storage, ref iid, IntPtr.Zero, out dataObjectPointer);
            ThrowIfFailed(hr, "OleLoad");

            hr = OleRun(dataObjectPointer);
            ThrowIfFailed(hr, "OleRun");

            dataObject = (IDataObject)Marshal.GetTypedObjectForIUnknown(dataObjectPointer, typeof(IDataObject));
            oleObject = (IOleObject)dataObject;
            hr = oleObject.Update();
            ThrowIfFailed(hr, "IOleObject.Update");
            MetaFilePict picture = ExportWmf(dataObject, outputWmf);
            double widthPt = Math.Abs(picture.XExt) * 72.0 / 2540.0;
            double heightPt = Math.Abs(picture.YExt) * 72.0 / 2540.0;
            string dimensionsJson = "{\n"
                + "  \"source\": \"METAFILEPICT\",\n"
                + "  \"mapping_mode\": " + picture.MappingMode + ",\n"
                + "  \"x_ext_0_01mm\": " + picture.XExt + ",\n"
                + "  \"y_ext_0_01mm\": " + picture.YExt + ",\n"
                + "  \"width_pt\": " + widthPt.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture) + ",\n"
                + "  \"height_pt\": " + heightPt.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture) + ",\n"
                + "  \"baseline_source\": \"preserve_existing_word_position\",\n"
                + "  \"eqndims_clipboard_available\": false\n"
                + "}\n";
            File.WriteAllText(outputDims, dimensionsJson, new UTF8Encoding(false));
            return picture;
        }
        finally
        {
            if (oleObject != null)
            {
                try
                {
                    int closeHr = oleObject.Close(1); // OLECLOSE_NOSAVE
                    ThrowIfFailed(closeHr, "IOleObject.Close");
                }
                catch
                {
                    // Release the COM identity even when the server cannot close cleanly.
                }
            }
            if (dataObject != null && Marshal.IsComObject(dataObject)) Marshal.FinalReleaseComObject(dataObject);
            else if (dataObjectPointer != IntPtr.Zero) Marshal.Release(dataObjectPointer);
            if (storage != IntPtr.Zero) Marshal.Release(storage);
        }
    }

    private static int RunWorker()
    {
        Console.InputEncoding = new UTF8Encoding(false);
        Console.OutputEncoding = new UTF8Encoding(false);
        Console.WriteLine("READY\t" + BridgeVersion);
        Console.Out.Flush();

        string line;
        while ((line = Console.ReadLine()) != null)
        {
            if (line == "QUIT")
            {
                Console.WriteLine("BYE");
                Console.Out.Flush();
                return 0;
            }

            string[] fields = line.Split('\t');
            if (fields.Length != 4 || fields[0] != "RENDER")
            {
                Console.WriteLine("ERR\t" + EncodeText("Invalid worker request"));
                Console.Out.Flush();
                continue;
            }

            try
            {
                MetaFilePict picture = RenderOne(
                    DecodePath(fields[1]), DecodePath(fields[2]), DecodePath(fields[3]));
                Console.WriteLine(
                    "OK\t" + picture.MappingMode
                    + "\t" + picture.XExt
                    + "\t" + picture.YExt);
            }
            catch (Exception ex)
            {
                string message = ex.GetType().FullName + ": " + ex.Message;
                if (ex is COMException)
                    message += " (HRESULT 0x" + ((COMException)ex).ErrorCode.ToString("X8") + ")";
                Console.WriteLine("ERR\t" + EncodeText(message));
            }
            Console.Out.Flush();
        }
        return 0;
    }

    [STAThread]
    public static int Main(string[] args)
    {
        if (args.Length == 1 && args[0] == "--version")
        {
            Console.WriteLine("OlePreviewRenderer " + BridgeVersion);
            return 0;
        }
        if (!((args.Length == 1 && args[0] == "--worker") || args.Length == 3))
        {
            Console.Error.WriteLine("usage: OlePreviewRenderer.exe input.bin output.wmf dimensions.json");
            Console.Error.WriteLine("   or: OlePreviewRenderer.exe --worker");
            return 2;
        }

        int oleHr = OleInitialize(IntPtr.Zero);
        try
        {
            ThrowIfFailed(oleHr, "OleInitialize");
            if (args.Length == 1) return RunWorker();

            MetaFilePict picture = RenderOne(args[0], args[1], args[2]);
            double widthPt = Math.Abs(picture.XExt) * 72.0 / 2540.0;
            double heightPt = Math.Abs(picture.YExt) * 72.0 / 2540.0;
            Console.WriteLine("status=ok");
            Console.WriteLine("wmf=" + Path.GetFullPath(args[1]));
            Console.WriteLine("mapping_mode=" + picture.MappingMode);
            Console.WriteLine("x_ext=" + picture.XExt);
            Console.WriteLine("y_ext=" + picture.YExt);
            Console.WriteLine("width_pt=" + widthPt.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture));
            Console.WriteLine("height_pt=" + heightPt.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture));
            Console.WriteLine("dimensions=" + Path.GetFullPath(args[2]));
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("status=error");
            Console.Error.WriteLine(ex.GetType().FullName);
            Console.Error.WriteLine(ex.Message);
            if (ex is COMException)
                Console.Error.WriteLine("hresult=0x" + ((COMException)ex).ErrorCode.ToString("X8"));
            return 1;
        }
        finally
        {
            if (oleHr >= 0) OleUninitialize();
        }
    }
}
