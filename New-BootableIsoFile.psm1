function New-BootableIsoFile {
<#
.SYNOPSIS
    Creates a bootable ISO file from the specified source directory using an optimized approach.
.DESCRIPTION
    This function creates a bootable ISO file from the specified source directory using a streamlined
    approach. It's optimized for large Windows ISOs and includes enhanced error handling and logging.

.PARAMETER SourceDir
    Specifies the source directory containing the files to be included in the ISO.

.PARAMETER IsoPath
    Specifies the path for the output ISO file. If not provided, it generates the path based on the source directory.

.PARAMETER IsoLabel
    Specifies the label to be assigned to the ISO file. If not provided, it uses the ISO file name as the label.

.PARAMETER FileSystem
    Specifies the file system type. Valid values: 'UDF', 'ISO9660', 'Joliet', 'All'. Defaults to 'UDF' for best compatibility with large files.

.PARAMETER Force
    Overwrites the destination ISO file if it already exists.

.PARAMETER Verbose
    Provides detailed progress information during ISO creation.

.EXAMPLE
    New-BootableIsoFile -SourceDir "C:\WindowsISO" -IsoPath "C:\Output\Windows.iso" -IsoLabel "WIN11_X64" -Verbose
    Creates a bootable Windows ISO with detailed progress output.

.EXAMPLE
    New-BootableIsoFile "C:\SourceFiles"
    Creates an ISO using default naming in the parent directory of the source.

.NOTES
    Based on MakeISO by AveYo. Optimized for large Windows ISO creation.
    Requires Windows 7/Server 2008 R2 or newer for IMAPI2.
#>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias("FullName")]
        [string]$SourceDir,

        [Parameter(Mandatory = $false)]
        [string]$IsoPath,

        [Parameter(Mandatory = $false)]
        [string]$IsoLabel,

        [Parameter(Mandatory = $false)]
        [ValidateSet('UDF', 'ISO9660', 'Joliet', 'All')]
        [string]$FileSystem = 'UDF',

        [Parameter(Mandatory = $false)]
        [switch]$Force
    )

    begin {
        # Validate and normalize source directory
        try {
            $SourceDir = (Resolve-Path -LiteralPath $SourceDir -ErrorAction Stop).Path
            if (-not (Test-Path -LiteralPath $SourceDir -PathType Container)) {
                throw "Source path is not a valid directory."
            }
            Write-Verbose "Source directory: $SourceDir"
        }
        catch {
            Write-Error "Invalid source directory '$SourceDir': $($_.Exception.Message)"
            return
        }

        # Set default ISO path if not provided
        if (-not $IsoPath) {
            $isoName = Split-Path -Leaf $SourceDir
            $parentDir = Split-Path $SourceDir -Parent
            if (-not $parentDir) { $parentDir = $SourceDir }
            $IsoPath = Join-Path $parentDir ($isoName + ".iso")
        }

        # Resolve and validate ISO path
        try {
            $isoParentDir = Split-Path -Path $IsoPath -Parent
            if ($isoParentDir -and -not (Test-Path -LiteralPath $isoParentDir -PathType Container)) {
                New-Item -Path $isoParentDir -ItemType Directory -Force | Out-Null
                Write-Verbose "Created parent directory: $isoParentDir"
            }
            Write-Verbose "ISO output path: $IsoPath"
        }
        catch {
            Write-Error "Invalid ISO path '$IsoPath': $($_.Exception.Message)"
            return
        }

        # Set default label if not provided and sanitize it
        if (-not $IsoLabel) {
            $IsoLabel = [System.IO.Path]::GetFileNameWithoutExtension($IsoPath)
        }
        
        # Sanitize ISO label - remove invalid characters and limit length
        $IsoLabel = $IsoLabel -replace '[^\w\s-]', '_'  # Replace special chars with underscore
        $IsoLabel = $IsoLabel -replace '\s+', '_'      # Replace spaces with underscore  
        $IsoLabel = $IsoLabel.Substring(0, [Math]::Min($IsoLabel.Length, 32))  # Limit to 32 chars
        $IsoLabel = $IsoLabel.Trim('_').ToUpper()      # Remove trailing underscores and convert to uppercase

        if ([string]::IsNullOrEmpty($IsoLabel)) {
            $IsoLabel = "DVD-ROM"
        }

        Write-Verbose "ISO label (sanitized): $IsoLabel"

        # Check for existing file
        if ((Test-Path -LiteralPath $IsoPath) -and -not $Force) {
            Write-Error "ISO file already exists at '$IsoPath'. Use -Force to overwrite."
            return
        }

        # Calculate source size for progress reporting
        try {
            $sourceSize = (Get-ChildItem -Path $SourceDir -Recurse -File -ErrorAction SilentlyContinue | 
                          Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
            if ($sourceSize) {
                $sourceSizeGB = [math]::Round($sourceSize / 1GB, 2)
                Write-Verbose "Source directory size: ${sourceSizeGB}GB"
                
                if ($sourceSizeGB -gt 4.7) {
                    Write-Warning "Source size (${sourceSizeGB}GB) exceeds DVD capacity. Consider using Blu-ray or USB for distribution."
                }
            }
        }
        catch {
            Write-Verbose "Could not calculate source directory size."
        }

        # Set file system flag
        $fileSystemFlag = switch ($FileSystem) {
            'ISO9660' { 1 }
            'Joliet'  { 2 }
            'UDF'     { 4 }
            'All'     { 7 } # ISO9660 + Joliet + UDF
            default   { 4 } # Default to UDF
        }
        Write-Verbose "Using file system: $FileSystem (flag: $fileSystemFlag)"
    }

    process {
        if (-not $PSCmdlet.ShouldProcess($IsoPath, "Create ISO from '$SourceDir'")) {
            return
        }

        # Script block for ISO creation job
        $scriptBlock = {
            param($sourceDir, $isoPath, $isoLabel, $fileSystemFlag, $verbosePreference)
            
            # Set verbose preference in job
            $VerbosePreference = $verbosePreference

            # Enhanced C# ISO Creator with better error handling
            $cSharpSource = @"
            using System;
            using System.Runtime.InteropServices;
            using System.Runtime.InteropServices.ComTypes;

            public class IsoCreator {
                [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
                internal static extern int SHCreateStreamOnFileEx(
                    string fileName, 
                    uint grfMode, 
                    uint dwAttributes, 
                    bool fCreate, 
                    IStream pstmTemplate, 
                    out IStream ppstm
                );

                public static int Create(string isoPath, ref object sourceStream, int blockSize, int totalBlocks) {
                    IStream source = null;
                    IStream isoStream = null;
                    
                    try {
                        source = (IStream)sourceStream;
                        
                        // STGM_CREATE | STGM_WRITE | STGM_SHARE_EXCLUSIVE
                        int hr = SHCreateStreamOnFileEx(isoPath, 0x1001, 0x80, true, null, out isoStream);
                        if (hr != 0) {
                            return 1; // Failed to create output stream
                        }

                        // Optimized copying algorithm for large files
                        int divider = totalBlocks > 1024 ? 1024 : 1;
                        int padding = totalBlocks % divider;
                        int blockCount = blockSize * divider;
                        int totalBlock = (totalBlocks - padding) / divider;

                        // Copy any remaining blocks first
                        if (padding > 0) {
                            source.CopyTo(isoStream, (long)padding * blockSize, IntPtr.Zero, IntPtr.Zero);
                        }

                        // Copy main blocks in optimized chunks
                        while (totalBlock-- > 0) {
                            source.CopyTo(isoStream, blockCount, IntPtr.Zero, IntPtr.Zero);
                        }

                        // Commit the stream
                        isoStream.Commit(0);
                        return 0; // Success
                    }
                    catch (OutOfMemoryException) {
                        return 2; // Out of memory
                    }
                    catch (UnauthorizedAccessException) {
                        return 3; // Access denied
                    }
                    catch (Exception) {
                        return 4; // General error
                    }
                    finally {
                        if (isoStream != null) {
                            try { Marshal.ReleaseComObject(isoStream); } catch { }
                        }
                        // Don't release source stream - it's managed by IMAPI
                    }
                }
            }
"@

            try {
                Add-Type -TypeDefinition $cSharpSource -Language CSharp -ErrorAction Stop
                Write-Verbose "C# ISO Creator compiled successfully."
            }
            catch {
                Write-Error "Failed to compile C# helper: $($_.Exception.Message)"
                return 99
            }

            # Initialize COM objects
            $bootOptions = @()
            $isBootable = $false
            $fsi = $null
            $resultImage = $null

            try {
                Write-Verbose "Checking for boot files..."
                
                # Boot file configurations: [path, platformId]
                $bootConfigs = @(
                    @('boot\etfsboot.com', 0),      # BIOS boot
                    @('efi\microsoft\boot\efisys.bin', 0xEF)  # UEFI boot
                )

                # Check for boot files and create boot options
                foreach ($config in $bootConfigs) {
                    $bootFilePath = Join-Path $sourceDir $config[0]
                    $platformId = $config[1]
                    
                    if (Test-Path -LiteralPath $bootFilePath -PathType Leaf) {
                        Write-Verbose "Found boot file: $bootFilePath (Platform: 0x$($platformId.ToString('X')))"
                        
                        try {
                            # Create ADODB stream for boot file
                            $adodbStream = New-Object -ComObject ADODB.Stream
                            $adodbStream.Type = 1  # Binary
                            $adodbStream.Open()
                            $adodbStream.LoadFromFile($bootFilePath)

                            # Create boot options
                            $bootOption = New-Object -ComObject IMAPI2FS.BootOptions
                            $bootOption.AssignBootImage($adodbStream.psobject.BaseObject)
                            $bootOption.PlatformId = $platformId
                            $bootOption.Emulation = 0  # No emulation
                            $bootOption.Manufacturer = 'Microsoft'
                            
                            $bootOptions += $bootOption.psobject.BaseObject
                            $isBootable = $true
                            
                            Write-Verbose "Boot option created for platform 0x$($platformId.ToString('X'))"
                        }
                        catch {
                            Write-Warning "Failed to create boot option for $bootFilePath`: $($_.Exception.Message)"
                        }
                    }
                }

                if (-not $isBootable) {
                    Write-Verbose "No boot files found. Creating non-bootable ISO."
                }

                Write-Verbose "Creating file system image..."
                $fsi = New-Object -ComObject IMAPI2FS.MsftFileSystemImage
                $fsi.FileSystemsToCreate = $fileSystemFlag
                $fsi.FreeMediaBlocks = 0  # Auto-calculate
                
                # Validate and set volume name with additional checks
                try {
                    Write-Verbose "Setting volume name to: $isoLabel"
                    $fsi.VolumeName = $isoLabel
                }
                catch {
                    Write-Warning "Failed to set volume name '$isoLabel'. Using default."
                    $fsi.VolumeName = "WINDOWS_ISO"
                }

                if ($isBootable) {
                    Write-Verbose "Setting boot image options array with $($bootOptions.Count) boot option(s)."
                    try {
                        $fsi.BootImageOptionsArray = $bootOptions
                    }
                    catch {
                        Write-Warning "Failed to set boot options. Creating non-bootable ISO."
                        $isBootable = $false
                    }
                }

                Write-Verbose "Adding directory tree to ISO image..."
                $fsi.Root.AddTree($sourceDir, $false)

                Write-Verbose "Creating result image stream..."
                $resultImage = $fsi.CreateResultImage()
                
                Write-Verbose "Image details - Block Size: $($resultImage.BlockSize), Total Blocks: $($resultImage.TotalBlocks)"
                $estimatedSizeMB = [math]::Round(($resultImage.BlockSize * $resultImage.TotalBlocks) / 1MB, 2)
                Write-Verbose "Estimated ISO size: ${estimatedSizeMB}MB"

                Write-Verbose "Writing ISO file to disk..."
                $result = [IsoCreator]::Create($isoPath, [ref]$resultImage.ImageStream, $resultImage.BlockSize, $resultImage.TotalBlocks)

                # Interpret result codes
                switch ($result) {
                    0 { 
                        Write-Verbose "ISO creation completed successfully."
                        return 0 
                    }
                    1 { 
                        Write-Error "Failed to create output file stream. Check path and permissions."
                        return 1 
                    }
                    2 { 
                        Write-Error "Out of memory during ISO creation."
                        return 2 
                    }
                    3 { 
                        Write-Error "Access denied. Check file permissions and disk space."
                        return 3 
                    }
                    4 { 
                        Write-Error "General error during ISO creation."
                        return 4 
                    }
                    default { 
                        Write-Error "Unknown error code: $result"
                        return $result 
                    }
                }
            }
            catch [System.Runtime.InteropServices.COMException] {
                $hresult = $_.Exception.HResult
                Write-Error "COM Exception (HRESULT: 0x$($hresult.ToString('X8'))): $($_.Exception.Message)"
                
                # Handle specific COM errors
                switch ($hresult) {
                    0xC0AAB132 { return 10 }  # IMAPI resource/memory error
                    0x8007000E { return 11 }  # System out of memory
                    0x80070057 { return 12 }  # Invalid parameter
                    0x80070005 { return 13 }  # Access denied
                    0xC0AAB101 { return 14 }  # Invalid value for parameter
                    default { return 20 }     # Other COM error
                }
            }
            catch {
                Write-Error "Unexpected error during ISO creation: $($_.Exception.Message)"
                return 30
            }
            finally {
                # Clean up COM objects
                if ($resultImage) {
                    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($resultImage) | Out-Null } catch { }
                }
                if ($fsi) {
                    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($fsi) | Out-Null } catch { }
                }
                foreach ($bootOpt in $bootOptions) {
                    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($bootOpt) | Out-Null } catch { }
                }
                
                # Force garbage collection
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers() 
                [System.GC]::Collect()
                
                Write-Verbose "COM objects cleaned up."
            }
        }

        # Execute ISO creation job
        $job = $null
        try {
            Write-Verbose "Starting ISO creation job..."
            $startTime = Get-Date
            
            $jobArgs = @($SourceDir, $IsoPath, $IsoLabel, $fileSystemFlag, $VerbosePreference)
            $job = Start-Job -ScriptBlock $scriptBlock -ArgumentList $jobArgs

            Write-Verbose "Waiting for job completion (Job ID: $($job.Id))..."
            
            # Monitor job with progress updates
            do {
                Start-Sleep -Seconds 5
                $elapsed = (Get-Date) - $startTime
                if ($elapsed.TotalMinutes -gt 0 -and ($elapsed.TotalSeconds % 30) -lt 5) {
                    Write-Verbose "ISO creation in progress... ($([math]::Round($elapsed.TotalMinutes, 1)) minutes elapsed)"
                }
            } while ($job.State -eq 'Running' -and $elapsed.TotalMinutes -lt 60)

            # Handle timeout
            if ($job.State -eq 'Running') {
                Write-Warning "Job is taking longer than expected. Waiting for completion..."
                Wait-Job $job -Timeout 1800 | Out-Null  # 30 minute max timeout
            } else {
                Wait-Job $job | Out-Null
            }

            # Get job results
            $jobOutput = Receive-Job $job
            $finalResult = $jobOutput | Where-Object { $_ -is [int] } | Select-Object -Last 1

            if ($job.State -eq 'Completed' -and $finalResult -eq 0) {
                $elapsedTime = (Get-Date) - $startTime
                Write-Host "✓ ISO file created successfully at: $IsoPath" -ForegroundColor Green
                Write-Verbose "Total creation time: $([math]::Round($elapsedTime.TotalMinutes, 2)) minutes"
                
                # Return file info
                if (Test-Path -LiteralPath $IsoPath) {
                    $isoFile = Get-Item -LiteralPath $IsoPath
                    $isoSizeMB = [math]::Round($isoFile.Length / 1MB, 2)
                    Write-Verbose "Final ISO size: ${isoSizeMB}MB"
                    return $isoFile
                }
            } else {
                # Handle job errors
                $errorMsg = switch ($finalResult) {
                    10 { "IMAPI resource/memory error. Try with a smaller source or more system memory." }
                    11 { "System out of memory. Close other applications and try again." }
                    12 { "Invalid parameter. Check source directory and boot files." }
                    13 { "Access denied. Check file permissions and available disk space." }
                    14 { "Invalid value for parameter. Check ISO label and file paths." }
                    20 { "COM interface error. IMAPI2 may not be properly installed." }
                    30 { "Unexpected error during processing." }
                    $null { "Job failed to return a result code." }
                    default { "ISO creation failed with code: $finalResult" }
                }
                
                Write-Error $errorMsg
                
                # Output job details for troubleshooting
                if ($job.ChildJobs[0].Verbose) {
                    $job.ChildJobs[0].Verbose | ForEach-Object { Write-Verbose "JOB: $($_.Message)" }
                }
                if ($job.ChildJobs[0].Warning) {
                    $job.ChildJobs[0].Warning | ForEach-Object { Write-Warning "JOB: $($_.Message)" }
                }
                if ($job.ChildJobs[0].Error) {
                    $job.ChildJobs[0].Error | ForEach-Object { Write-Error "JOB: $($_.Exception.Message)" }
                }
            }
        }
        catch {
            Write-Error "Failed to manage ISO creation job: $($_.Exception.Message)"
        }
        finally {
            if ($job) {
                Remove-Job $job -Force -ErrorAction SilentlyContinue
            }
        }
    }
}