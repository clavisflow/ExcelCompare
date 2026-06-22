using System.Buffers.Binary;
using System.Text;

namespace ExcelCompare.Services;

internal static class VbaProjectSourceReader
{
    public static VbaProjectSource Read(byte[] vbaProjectBytes)
    {
        var compoundFile = OleCompoundFile.Open(vbaProjectBytes);
        if (!compoundFile.TryReadStream("VBA/dir", out var compressedDir))
        {
            throw new InvalidDataException("VBA/dir ストリームが見つかりませんでした。");
        }

        var dirBytes = DecompressVbaContainer(compressedDir);
        var codePage = TryReadCodePage(dirBytes);
        var modules = ReadModuleDirectory(dirBytes)
            .Select(module => ReadModuleSource(compoundFile, module, codePage))
            .Where(module => module is not null)
            .Cast<VbaModuleSource>()
            .ToList();

        return new VbaProjectSource(modules, codePage);
    }

    private static VbaModuleSource? ReadModuleSource(
        OleCompoundFile compoundFile,
        ModuleDirectoryEntry module,
        int? codePage)
    {
        if (string.IsNullOrWhiteSpace(module.StreamName) ||
            !compoundFile.TryReadStream($"VBA/{module.StreamName}", out var streamBytes) ||
            module.TextOffset < 0 ||
            module.TextOffset >= streamBytes.Length)
        {
            return null;
        }

        var compressedSource = streamBytes[module.TextOffset..];
        var sourceBytes = DecompressVbaContainer(compressedSource);
        var sourceText = DecodeSource(sourceBytes, codePage);

        return new VbaModuleSource(
            string.IsNullOrWhiteSpace(module.Name) ? module.StreamName : module.Name,
            module.StreamName,
            sourceText);
    }

    private static IReadOnlyList<ModuleDirectoryEntry> ReadModuleDirectory(byte[] dirBytes)
    {
        var projectModulesOffset = FindProjectModulesOffset(dirBytes);
        if (projectModulesOffset < 0)
        {
            return [];
        }

        var moduleCount = BinaryPrimitives.ReadUInt16LittleEndian(dirBytes.AsSpan(projectModulesOffset + 6, 2));
        var offset = projectModulesOffset + 8;
        var modules = new List<ModuleDirectoryEntry>();

        for (var index = 0; index < moduleCount && offset + 6 <= dirBytes.Length; index++)
        {
            string? moduleName = null;
            string? streamName = null;
            var textOffset = -1;

            while (offset + 6 <= dirBytes.Length)
            {
                var recordId = BinaryPrimitives.ReadUInt16LittleEndian(dirBytes.AsSpan(offset, 2));
                var recordSize = BinaryPrimitives.ReadUInt32LittleEndian(dirBytes.AsSpan(offset + 2, 4));
                offset += 6;

                if (recordSize > int.MaxValue || offset + (int)recordSize > dirBytes.Length)
                {
                    return modules;
                }

                var recordBytes = dirBytes.AsSpan(offset, (int)recordSize);
                switch (recordId)
                {
                    case 0x0019:
                        moduleName = DecodeDirectoryString(recordBytes);
                        break;
                    case 0x001A:
                        streamName = DecodeDirectoryString(recordBytes);
                        break;
                    case 0x0031 when recordBytes.Length >= 4:
                        textOffset = BinaryPrimitives.ReadInt32LittleEndian(recordBytes[..4]);
                        break;
                }

                offset += (int)recordSize;

                if (recordId == 0x002B)
                {
                    modules.Add(new ModuleDirectoryEntry(moduleName, streamName, textOffset));
                    break;
                }
            }
        }

        return modules;
    }

    private static int FindProjectModulesOffset(byte[] dirBytes)
    {
        for (var offset = 0; offset + 8 <= dirBytes.Length; offset++)
        {
            var recordId = BinaryPrimitives.ReadUInt16LittleEndian(dirBytes.AsSpan(offset, 2));
            var recordSize = BinaryPrimitives.ReadUInt32LittleEndian(dirBytes.AsSpan(offset + 2, 4));

            if (recordId != 0x000F || recordSize != 2)
            {
                continue;
            }

            var moduleCount = BinaryPrimitives.ReadUInt16LittleEndian(dirBytes.AsSpan(offset + 6, 2));
            if (moduleCount is > 0 and < 1024)
            {
                return offset;
            }
        }

        return -1;
    }

    private static int? TryReadCodePage(byte[] dirBytes)
    {
        for (var offset = 0; offset + 8 <= dirBytes.Length; offset++)
        {
            var recordId = BinaryPrimitives.ReadUInt16LittleEndian(dirBytes.AsSpan(offset, 2));
            var recordSize = BinaryPrimitives.ReadUInt32LittleEndian(dirBytes.AsSpan(offset + 2, 4));

            if (recordId == 0x0003 && recordSize == 2)
            {
                return BinaryPrimitives.ReadUInt16LittleEndian(dirBytes.AsSpan(offset + 6, 2));
            }
        }

        return null;
    }

    private static string DecodeDirectoryString(ReadOnlySpan<byte> bytes)
    {
        return Encoding.Latin1.GetString(bytes).TrimEnd('\0');
    }

    private static string DecodeSource(byte[] bytes, int? codePage)
    {
        if (codePage is not null)
        {
            try
            {
                var encoding = Encoding.GetEncoding(
                    codePage.Value,
                    EncoderFallback.ExceptionFallback,
                    DecoderFallback.ExceptionFallback);
                return encoding.GetString(bytes);
            }
            catch (DecoderFallbackException)
            {
                // Fall through to tolerant decoding rather than failing the whole comparison.
            }
            catch (ArgumentException)
            {
                // Some VBA projects omit or use a code page unavailable in the runtime.
            }
        }

        foreach (var fallback in new[]
                 {
                     new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true),
                     Encoding.GetEncoding(1200, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback)
                 })
        {
            try
            {
                return fallback.GetString(bytes);
            }
            catch (DecoderFallbackException)
            {
                // Try the next tolerant fallback.
            }
        }

        // Latin1 preserves ASCII keywords even when the project uses an unknown legacy code page.
        return Encoding.Latin1.GetString(bytes);
    }

    private static byte[] DecompressVbaContainer(byte[] compressedContainer)
    {
        if (compressedContainer.Length == 0 || compressedContainer[0] != 0x01)
        {
            throw new InvalidDataException("VBA 圧縮ストリームのシグネチャが不正です。");
        }

        using var output = new MemoryStream();
        var inputOffset = 1;

        while (inputOffset + 2 <= compressedContainer.Length)
        {
            var chunkStart = inputOffset;
            var header = BinaryPrimitives.ReadUInt16LittleEndian(compressedContainer.AsSpan(inputOffset, 2));
            inputOffset += 2;

            var chunkSize = (header & 0x0FFF) + 3;
            var chunkSignature = (header & 0x7000) >> 12;
            var isCompressed = (header & 0x8000) != 0;
            var chunkEnd = Math.Min(chunkStart + chunkSize, compressedContainer.Length);

            if (chunkSignature != 0b011 || chunkEnd < inputOffset)
            {
                break;
            }

            if (!isCompressed)
            {
                var rawCount = Math.Min(4096, chunkEnd - inputOffset);
                output.Write(compressedContainer, inputOffset, rawCount);
                inputOffset = chunkEnd;
                continue;
            }

            var decompressedChunkStart = (int)output.Length;
            while (inputOffset < chunkEnd)
            {
                var flags = compressedContainer[inputOffset++];
                for (var bit = 0; bit < 8 && inputOffset < chunkEnd; bit++)
                {
                    var isCopyToken = (flags & (1 << bit)) != 0;
                    if (!isCopyToken)
                    {
                        output.WriteByte(compressedContainer[inputOffset++]);
                        continue;
                    }

                    if (inputOffset + 2 > chunkEnd)
                    {
                        inputOffset = chunkEnd;
                        break;
                    }

                    var copyToken = BinaryPrimitives.ReadUInt16LittleEndian(compressedContainer.AsSpan(inputOffset, 2));
                    inputOffset += 2;
                    CopyTokenBytes(output, decompressedChunkStart, copyToken);
                }
            }

            inputOffset = chunkEnd;
        }

        return output.ToArray();
    }

    private static void CopyTokenBytes(MemoryStream output, int decompressedChunkStart, ushort copyToken)
    {
        var decompressedCurrent = (int)output.Length;
        var difference = decompressedCurrent - decompressedChunkStart;
        var bitCount = Math.Max(4, CeilingLog2(Math.Max(1, difference)));
        var lengthMask = 0xFFFF >> bitCount;
        var offsetMask = ~lengthMask;
        var copyLength = (copyToken & lengthMask) + 3;
        var copyOffset = ((copyToken & offsetMask) >> (16 - bitCount)) + 1;
        var copySource = decompressedCurrent - copyOffset;

        if (copySource < decompressedChunkStart || copySource < 0)
        {
            return;
        }

        var buffer = output.GetBuffer();
        for (var index = 0; index < copyLength; index++)
        {
            output.WriteByte(buffer[copySource + index]);
            buffer = output.GetBuffer();
        }
    }

    private static int CeilingLog2(int value)
    {
        var result = 0;
        var power = 1;
        while (power < value)
        {
            power <<= 1;
            result++;
        }

        return result;
    }

    private sealed record ModuleDirectoryEntry(string? Name, string? StreamName, int TextOffset);
}

internal sealed record VbaProjectSource(IReadOnlyList<VbaModuleSource> Modules, int? CodePage);

internal sealed record VbaModuleSource(string Name, string StreamName, string SourceText);
