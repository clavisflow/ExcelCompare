using System.Buffers.Binary;

namespace ExcelCompare.Services;

internal sealed class OleCompoundFile
{
    private const int FreeSector = -1;
    private const int EndOfChain = -2;
    private const int FatSector = -3;
    private const int DifatSector = -4;
    private const int NoStream = -1;

    private readonly byte[] fileBytes;
    private readonly int sectorSize;
    private readonly int miniSectorSize;
    private readonly int miniStreamCutoffSize;
    private readonly int firstDirectorySector;
    private readonly int firstMiniFatSector;
    private readonly int miniFatSectorCount;
    private readonly int fatSectorCount;
    private readonly IReadOnlyList<int> fat;
    private readonly IReadOnlyList<int> miniFat;
    private readonly IReadOnlyList<DirectoryEntry> directoryEntries;
    private readonly byte[] miniStream;
    private readonly Dictionary<string, DirectoryEntry> streams;

    private OleCompoundFile(
        byte[] fileBytes,
        int sectorSize,
        int miniSectorSize,
        int miniStreamCutoffSize,
        int firstDirectorySector,
        int firstMiniFatSector,
        int miniFatSectorCount,
        int fatSectorCount,
        IReadOnlyList<int> fat,
        IReadOnlyList<int> miniFat,
        IReadOnlyList<DirectoryEntry> directoryEntries,
        byte[] miniStream,
        Dictionary<string, DirectoryEntry> streams)
    {
        this.fileBytes = fileBytes;
        this.sectorSize = sectorSize;
        this.miniSectorSize = miniSectorSize;
        this.miniStreamCutoffSize = miniStreamCutoffSize;
        this.firstDirectorySector = firstDirectorySector;
        this.firstMiniFatSector = firstMiniFatSector;
        this.miniFatSectorCount = miniFatSectorCount;
        this.fatSectorCount = fatSectorCount;
        this.fat = fat;
        this.miniFat = miniFat;
        this.directoryEntries = directoryEntries;
        this.miniStream = miniStream;
        this.streams = streams;
    }

    public static OleCompoundFile Open(byte[] fileBytes)
    {
        if (fileBytes.Length < 512 || !IsCompoundFile(fileBytes))
        {
            throw new InvalidDataException("VBA プロジェクトの OLE 構造を読み取れませんでした。");
        }

        var sectorShift = ReadUInt16(fileBytes, 30);
        var miniSectorShift = ReadUInt16(fileBytes, 32);
        var sectorSize = 1 << sectorShift;
        var miniSectorSize = 1 << miniSectorShift;
        var fatSectorCount = ReadInt32(fileBytes, 44);
        var firstDirectorySector = ReadInt32(fileBytes, 48);
        var miniStreamCutoffSize = ReadInt32(fileBytes, 56);
        var firstMiniFatSector = ReadInt32(fileBytes, 60);
        var miniFatSectorCount = ReadInt32(fileBytes, 64);
        var firstDifatSector = ReadInt32(fileBytes, 68);
        var difatSectorCount = ReadInt32(fileBytes, 72);

        var difat = ReadDifat(fileBytes, sectorSize, fatSectorCount, firstDifatSector, difatSectorCount);
        var fat = ReadFat(fileBytes, sectorSize, difat, fatSectorCount);
        var directoryBytes = ReadRegularStream(fileBytes, sectorSize, fat, firstDirectorySector, long.MaxValue);
        var directoryEntries = ReadDirectoryEntries(directoryBytes).ToList();

        var miniFat = firstMiniFatSector >= 0 && miniFatSectorCount > 0
            ? ReadMiniFat(fileBytes, sectorSize, fat, firstMiniFatSector, miniFatSectorCount)
            : [];

        var rootEntry = directoryEntries.Count > 0 ? directoryEntries[0] : null;
        var miniStream = rootEntry is not null && rootEntry.StartSector >= 0 && rootEntry.StreamSize > 0
            ? ReadRegularStream(fileBytes, sectorSize, fat, rootEntry.StartSector, ToSafeStreamSize(rootEntry.StreamSize))
            : [];

        var streams = new Dictionary<string, DirectoryEntry>(StringComparer.OrdinalIgnoreCase);
        if (rootEntry is not null)
        {
            TraverseChildren(directoryEntries, rootEntry.ChildId, string.Empty, streams);
        }

        return new OleCompoundFile(
            fileBytes,
            sectorSize,
            miniSectorSize,
            miniStreamCutoffSize,
            firstDirectorySector,
            firstMiniFatSector,
            miniFatSectorCount,
            fatSectorCount,
            fat,
            miniFat,
            directoryEntries,
            miniStream,
            streams);
    }

    public bool TryReadStream(string path, out byte[] bytes)
    {
        bytes = [];
        var normalizedPath = NormalizePath(path);
        if (!streams.TryGetValue(normalizedPath, out var entry))
        {
            return false;
        }

        bytes = ReadStream(entry);
        return true;
    }

    public IReadOnlyList<string> StreamPaths => streams.Keys.Order(StringComparer.OrdinalIgnoreCase).ToList();

    private byte[] ReadStream(DirectoryEntry entry)
    {
        if (entry.StreamSize <= 0)
        {
            return [];
        }

        var streamSize = ToSafeStreamSize(entry.StreamSize);

        if (streamSize < miniStreamCutoffSize && miniStream.Length > 0 && miniFat.Count > 0)
        {
            return ReadMiniStream(entry.StartSector, streamSize);
        }

        return ReadRegularStream(fileBytes, sectorSize, fat, entry.StartSector, streamSize);
    }

    private byte[] ReadMiniStream(int startMiniSector, long streamSize)
    {
        if (startMiniSector < 0)
        {
            return [];
        }

        using var output = new MemoryStream();
        var sector = startMiniSector;
        var visited = new HashSet<int>();

        while (sector >= 0 && sector < miniFat.Count && visited.Add(sector) && output.Length < streamSize)
        {
            var offset = sector * miniSectorSize;
            if (offset < 0 || offset >= miniStream.Length)
            {
                break;
            }

            var count = (int)Math.Min(miniSectorSize, Math.Min(streamSize - output.Length, miniStream.Length - offset));
            output.Write(miniStream, offset, count);

            var next = miniFat[sector];
            if (next == EndOfChain || next == FreeSector)
            {
                break;
            }

            sector = next;
        }

        return output.ToArray();
    }

    private static bool IsCompoundFile(byte[] bytes)
    {
        ReadOnlySpan<byte> signature = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
        return bytes.AsSpan(0, signature.Length).SequenceEqual(signature);
    }

    private static IReadOnlyList<int> ReadDifat(
        byte[] bytes,
        int sectorSize,
        int fatSectorCount,
        int firstDifatSector,
        int difatSectorCount)
    {
        var difat = new List<int>();

        for (var offset = 76; offset < 512 && difat.Count < fatSectorCount; offset += 4)
        {
            var sector = ReadInt32(bytes, offset);
            if (sector >= 0)
            {
                difat.Add(sector);
            }
        }

        var currentDifatSector = firstDifatSector;
        for (var index = 0; index < difatSectorCount && currentDifatSector >= 0 && difat.Count < fatSectorCount; index++)
        {
            var sectorBytes = ReadSector(bytes, sectorSize, currentDifatSector);
            var entriesPerSector = sectorSize / 4;

            for (var entryIndex = 0; entryIndex < entriesPerSector - 1 && difat.Count < fatSectorCount; entryIndex++)
            {
                var sector = ReadInt32(sectorBytes, entryIndex * 4);
                if (sector >= 0)
                {
                    difat.Add(sector);
                }
            }

            currentDifatSector = ReadInt32(sectorBytes, (entriesPerSector - 1) * 4);
        }

        return difat;
    }

    private static IReadOnlyList<int> ReadFat(byte[] bytes, int sectorSize, IReadOnlyList<int> difat, int fatSectorCount)
    {
        var fat = new List<int>();
        foreach (var fatSector in difat.Take(fatSectorCount))
        {
            if (fatSector < 0)
            {
                continue;
            }

            var sectorBytes = ReadSector(bytes, sectorSize, fatSector);
            for (var offset = 0; offset < sectorBytes.Length; offset += 4)
            {
                fat.Add(ReadInt32(sectorBytes, offset));
            }
        }

        return fat;
    }

    private static IReadOnlyList<int> ReadMiniFat(
        byte[] bytes,
        int sectorSize,
        IReadOnlyList<int> fat,
        int firstMiniFatSector,
        int miniFatSectorCount)
    {
        var miniFatBytes = ReadRegularStream(bytes, sectorSize, fat, firstMiniFatSector, (long)miniFatSectorCount * sectorSize);
        var miniFat = new List<int>();
        for (var offset = 0; offset + 4 <= miniFatBytes.Length; offset += 4)
        {
            miniFat.Add(ReadInt32(miniFatBytes, offset));
        }

        return miniFat;
    }

    private static byte[] ReadRegularStream(
        byte[] bytes,
        int sectorSize,
        IReadOnlyList<int> fat,
        int startSector,
        long streamSize)
    {
        if (startSector < 0)
        {
            return [];
        }

        using var output = new MemoryStream();
        var sector = startSector;
        var visited = new HashSet<int>();

        while (sector >= 0 && sector < fat.Count && visited.Add(sector) && output.Length < streamSize)
        {
            var sectorBytes = ReadSector(bytes, sectorSize, sector);
            var count = (int)Math.Min(sectorBytes.Length, streamSize - output.Length);
            output.Write(sectorBytes, 0, count);

            var next = fat[sector];
            if (next is EndOfChain or FreeSector or FatSector or DifatSector)
            {
                break;
            }

            sector = next;
        }

        return output.ToArray();
    }

    private static byte[] ReadSector(byte[] bytes, int sectorSize, int sector)
    {
        var offset = checked((sector + 1) * sectorSize);
        if (offset < 0 || offset >= bytes.Length)
        {
            return [];
        }

        var count = Math.Min(sectorSize, bytes.Length - offset);
        var sectorBytes = new byte[count];
        Buffer.BlockCopy(bytes, offset, sectorBytes, 0, count);
        return sectorBytes;
    }

    private static IEnumerable<DirectoryEntry> ReadDirectoryEntries(byte[] directoryBytes)
    {
        for (var offset = 0; offset + 128 <= directoryBytes.Length; offset += 128)
        {
            var nameLength = ReadUInt16(directoryBytes, offset + 64);
            if (nameLength < 2 || nameLength > 64)
            {
                continue;
            }

            var nameBytesLength = nameLength - 2;
            var name = System.Text.Encoding.Unicode.GetString(directoryBytes, offset, nameBytesLength);
            var objectType = directoryBytes[offset + 66];
            var leftSiblingId = ReadInt32(directoryBytes, offset + 68);
            var rightSiblingId = ReadInt32(directoryBytes, offset + 72);
            var childId = ReadInt32(directoryBytes, offset + 76);
            var startSector = ReadInt32(directoryBytes, offset + 116);
            var streamSize = ReadUInt64(directoryBytes, offset + 120);

            yield return new DirectoryEntry(
                name,
                objectType,
                leftSiblingId,
                rightSiblingId,
                childId,
                startSector,
                streamSize);
        }
    }

    private static void TraverseChildren(
        IReadOnlyList<DirectoryEntry> directoryEntries,
        int entryId,
        string parentPath,
        IDictionary<string, DirectoryEntry> streams)
    {
        if (entryId is NoStream || entryId < 0 || entryId >= directoryEntries.Count)
        {
            return;
        }

        var entry = directoryEntries[entryId];
        TraverseChildren(directoryEntries, entry.LeftSiblingId, parentPath, streams);

        var path = string.IsNullOrEmpty(parentPath) ? entry.Name : $"{parentPath}/{entry.Name}";
        if (entry.ObjectType == 2)
        {
            streams[NormalizePath(path)] = entry;
        }
        else if (entry.ObjectType is 1 or 5)
        {
            TraverseChildren(directoryEntries, entry.ChildId, path, streams);
        }

        TraverseChildren(directoryEntries, entry.RightSiblingId, parentPath, streams);
    }

    private static string NormalizePath(string path) => path.Replace('\\', '/').Trim('/');

    private static long ToSafeStreamSize(ulong streamSize) =>
        streamSize > long.MaxValue ? long.MaxValue : (long)streamSize;

    private static ushort ReadUInt16(byte[] bytes, int offset) =>
        BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(offset, 2));

    private static int ReadInt32(byte[] bytes, int offset) =>
        BinaryPrimitives.ReadInt32LittleEndian(bytes.AsSpan(offset, 4));

    private static ulong ReadUInt64(byte[] bytes, int offset) =>
        BinaryPrimitives.ReadUInt64LittleEndian(bytes.AsSpan(offset, 8));

    private sealed record DirectoryEntry(
        string Name,
        byte ObjectType,
        int LeftSiblingId,
        int RightSiblingId,
        int ChildId,
        int StartSector,
        ulong StreamSize);
}
