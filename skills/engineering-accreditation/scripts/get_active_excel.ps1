param([switch]$ListWorkbooks)

# Read-only running-object discovery. This helper never starts Excel or saves a workbook.
function Get-AccreditationExcelInstance {
    if (-not ([System.Management.Automation.PSTypeName]'Accreditation.ExcelRunningObject').Type) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace Accreditation {
    public static class ExcelRunningObject {
        [DllImport("oleaut32.dll", PreserveSig=false)]
        private static extern void GetActiveObject(ref Guid clsid, IntPtr reserved,
            [MarshalAs(UnmanagedType.IUnknown)] out object instance);
        public static object Get() {
            Guid clsid;
            using (var key = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(@"Excel.Application\CLSID", false)) {
                var value = key == null ? null : key.GetValue(null) as string;
                if (!Guid.TryParse(value, out clsid)) {
                    throw new InvalidOperationException("Excel.Application has no valid registered CLSID; no registration or launch attempted.");
                }
            }
            object instance;
            GetActiveObject(ref clsid, IntPtr.Zero, out instance);
            return instance;
        }
    }
}
'@
    }
    return [Accreditation.ExcelRunningObject]::Get()
}

if ($ListWorkbooks) {
    try {
        $excel = Get-AccreditationExcelInstance
        $books = @()
        foreach ($book in $excel.Workbooks) {
            $books += [pscustomobject]@{
                Name = $book.Name
                FullName = $book.FullName
                Saved = [bool]$book.Saved
            }
        }
        [pscustomobject]@{Connected=$true; Workbooks=$books} | ConvertTo-Json -Depth 4 -Compress
    } catch {
        [pscustomobject]@{Connected=$false; Error=$_.Exception.Message} | ConvertTo-Json -Compress
        exit 1
    }
}
