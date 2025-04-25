# This function provides a more controlled and transparent method for creating a bootable ISO.
function New-BootableIsoFile {
<#
.SYNOPSIS
    Creates a bootable ISO file from the specified source directory.
.DESCRIPTION
    This function creates a bootable ISO file from the specified source directory. 
    If the isoPath parameter is not provided, the function generates the ISO file name based on the source directory name. 
    If the isoLabel parameter is not provided, it uses the ISO file name as the label.

.PARAMETER SourceDir
    Specifies the source directory containing the files to be included in the ISO.

.PARAMETER IsoPath
    Specifies the path for the output ISO file. If not provided, it generates the path based on the source directory.

.PARAMETER IsoLabel
    Specifies the label to be assigned to the ISO file. If not provided, it uses the ISO file name as the label.

.EXAMPLE
    New-BootableIsoFile -SourceDir "C:\SourceFiles" -IsoPath "C:\Output\MyIsoFile.iso" -IsoLabel "MyLabel"
    Creates a bootable ISO file named MyIsoFile.iso with the content of C:\SourceFiles and assigns MyLabel as the ISO label.
#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$sourceDir,

        [Parameter(Mandatory=$false)]
        [string]$isoPath,

        [Parameter(Mandatory=$false)]
        [string]$isoLabel
    )

    # Normalize sourceDir to full path
    $sourceDir = (Resolve-Path -LiteralPath $sourceDir).Path

    # Set the default value for isoPath based on the sourceDir
    if (-not $isoPath) {
        $IsoName = Split-Path -Leaf $sourceDir
        $isoPath = Join-Path (Split-Path $sourceDir -Parent) ($IsoName + ".iso")
    }

    # Set the default value for isoLabel based on the isoPath
    if (!$isoLabel) {
        $isoLabel = [System.IO.Path]::GetFileNameWithoutExtension($isoPath)
    }

    # Script block to be run in a separate process
    $scriptBlock = {
    # base on MakeISO by AveYo
    param($sourceDir,$isoPath,$isoLabel)
    $Source = @"
    using System;
    using System.Runtime.InteropServices;
    using System.Runtime.InteropServices.ComTypes;

    public class IsoCreator {
        [DllImport("shlwapi", CharSet = CharSet.Unicode)]
        internal static extern void SHCreateStreamOnFileEx(string f, uint m, uint d, bool b, IStream r, out IStream s);
        public static int Create(string isoPath, ref object sourceStream, int blockSize, int totalBlocks) {
            IStream source = (IStream)sourceStream, isoStream;
            try {
                SHCreateStreamOnFileEx(isoPath, 0x1001, 0x80, true, null, out isoStream);
            }
            catch (Exception) {
                return 1;
            }

            int divider = totalBlocks > 1024 ? 1024 : 1;
            int padding = totalBlocks % divider;
            int blockCount = blockSize * divider;
            int totalBlock = (totalBlocks - padding) / divider;

            if (padding > 0) source.CopyTo(isoStream, padding * blockCount, IntPtr.Zero, IntPtr.Zero);
            while (totalBlock-- > 0) {
                source.CopyTo(isoStream, blockCount, IntPtr.Zero, IntPtr.Zero);
            }

            isoStream.Commit(0);
            return 0;
        }
    }
"@
        Add-Type $Source
        $BOOT = @()
        $bootable = 0
        $mbr_efi = @(0,0xEF)
        $bootfiles = @('boot\etfsboot.com','efi\microsoft\boot\efisys.bin')

        0,1 | ForEach-Object {
            $bootfile = Join-Path $sourceDir $bootfiles[$_]
            if (Test-Path $bootfile -PathType Leaf) {
                $bin = New-Object -ComObject ADODB.Stream
                $bin.Open()
                $bin.Type = 1
                $bin.LoadFromFile($bootfile)

                $opt = New-Object -ComObject IMAPI2FS.BootOptions
                $opt.AssignBootImage($bin.psobject.BaseObject)
                $opt.PlatformId = $mbr_efi[$_]
                $opt.Emulation = 0
                $bootable = 1
                $opt.Manufacturer = 'Microsoft'
                $BOOT += $opt.psobject.BaseObject
            }
        }

        $fsi = New-Object -ComObject IMAPI2FS.MsftFileSystemImage
        $fsi.FileSystemsToCreate = 4
        $fsi.FreeMediaBlocks = 0

        if ($bootable) {
            $fsi.BootImageOptionsArray = $BOOT
        }

        $TREE = $fsi.Root
        $TREE.AddTree($sourceDir,$false)
        $fsi.VolumeName = $isoLabel

        $obj = $fsi.CreateResultImage()
        $ret = [IsoCreator]::Create($isoPath,[ref]$obj.ImageStream,$obj.BlockSize,$obj.TotalBlocks)

        [GC]::Collect()
        return $ret
    }

    $job = sajb -Script $scriptBlock -args $sourceDir,$isoPath,$isoLabel

    $result = rcjb $job -wait -auto
    sleep 2

    if ($result -eq 0) {
        Write-Host -b black -f cyan "The ISO file has been created at $isoPath"
    } else {
        Write-Warning "Creation of the ISO file has failed."
    }
}
