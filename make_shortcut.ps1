# make_shortcut.ps1 — 文書横断検索のタスクバー用ショートカットを作成
# doc_search.pyw / doc_search.ico と同じフォルダに置いて実行する
# (実行は同梱の make_shortcut.bat 経由を推奨)

$AppId   = "Base.DocSearch.Explorer"
$Here    = Split-Path -Parent $MyInvocation.MyCommand.Path
$Script  = Join-Path $Here "doc_search.pyw"
$Icon    = Join-Path $Here "doc_search.ico"
$LnkPath = Join-Path ([Environment]::GetFolderPath("Desktop")) "文書横断検索.lnk"

if (-not (Test-Path $Script)) {
    Write-Host "[エラー] doc_search.pyw が同じフォルダにありません: $Script"
    exit 1
}

# pythonw.exe を探す
$pyw = $null
$cmd = Get-Command pythonw.exe -ErrorAction SilentlyContinue
if ($cmd) { $pyw = $cmd.Source }
if (-not $pyw) {
    $cmd = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($cmd) {
        $cand = Join-Path (Split-Path $cmd.Source) "pythonw.exe"
        if (Test-Path $cand) { $pyw = $cand }
    }
}
if (-not $pyw) {
    Write-Host "[エラー] pythonw.exe が見つかりません。PATHを確認してください。"
    exit 1
}

# 1) 通常のショートカットを作成
$ws  = New-Object -ComObject WScript.Shell
$lnk = $ws.CreateShortcut($LnkPath)
$lnk.TargetPath       = $pyw
$lnk.Arguments        = '"' + $Script + '"'
$lnk.WorkingDirectory = $Here
if (Test-Path $Icon) { $lnk.IconLocation = $Icon }
$lnk.Description      = "文書横断検索"
$lnk.Save()

# 2) ショートカットに AppUserModelID を書き込む
#    (アプリ側の SetCurrentProcessExplicitAppUserModelID と一致させることで、
#     ピン留めアイコンと実行中ウィンドウがタスクバー上で1つに統合される)
$code = @"
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

public class ShortcutAumid {
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct PropertyKey {
        public Guid fmtid; public uint pid;
        public PropertyKey(Guid f, uint p) { fmtid = f; pid = p; }
    }

    [StructLayout(LayoutKind.Explicit)]
    public struct PropVariant {
        [FieldOffset(0)] public ushort vt;
        [FieldOffset(8)] public IntPtr pv;
    }

    [DllImport("ole32.dll")]
    static extern int PropVariantClear(ref PropVariant pvar);

    [ComImport, Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IPropertyStore {
        int GetCount(out uint c);
        int GetAt(uint i, out PropertyKey k);
        int GetValue(ref PropertyKey k, out PropVariant v);
        int SetValue(ref PropertyKey k, ref PropVariant v);
        int Commit();
    }

    public static void Set(string lnkPath, string appId) {
        Type t = Type.GetTypeFromCLSID(
            new Guid("00021401-0000-0000-C000-000000000046"));
        object link = Activator.CreateInstance(t);
        ((IPersistFile)link).Load(lnkPath, 2);
        IPropertyStore ps = (IPropertyStore)link;
        PropertyKey key = new PropertyKey(
            new Guid("9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3"), 5);
        PropVariant v = new PropVariant();
        v.vt = 31;
        v.pv = Marshal.StringToCoTaskMemUni(appId);
        ps.SetValue(ref key, ref v);
        ps.Commit();
        PropVariantClear(ref v);
        ((IPersistFile)link).Save(lnkPath, true);
    }
}
"@
Add-Type -TypeDefinition $code
[ShortcutAumid]::Set($LnkPath, $AppId)

Write-Host "作成しました: $LnkPath"
Write-Host "デスクトップのショートカットを右クリック → 「タスクバーにピン留めする」で完了です。"
Write-Host "(Windows 11 では右クリック → 「その他のオプションを表示」の中にあります)"
