// PowerShell bridge that controls PowerPoint (Windows) through COM.
// Reads one JSON request per line on stdin and writes one JSON response per line on stdout.
// Request: { id, action, data }  Response: { id, ok, result?, error? }
// Keep this script ASCII only (Windows PowerShell 5.1 reads BOM-less .ps1 files as ANSI).

export const POWERPOINT_SCRIPT = String.raw`
$ErrorActionPreference = 'Stop'
$utf8 = New-Object System.Text.UTF8Encoding $false
$stdin = New-Object System.IO.StreamReader([Console]::OpenStandardInput(), $utf8)
$stdout = New-Object System.IO.StreamWriter([Console]::OpenStandardOutput(), $utf8)
$stdout.AutoFlush = $true

Add-Type -Namespace FreeShow -Name Win32 -UsingNamespace System.Text -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
[DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
[DllImport("user32.dll")] public static extern int GetWindowLong(IntPtr hWnd, int nIndex);
[DllImport("user32.dll")] public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
[DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
[DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr hWnd, StringBuilder name, int max);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int max);

// visible PowerPoint slide show window, preferring the one with the file name in its title
public static IntPtr FindSlideShowWindow(string fileName) {
    IntPtr found = IntPtr.Zero;
    EnumWindows(delegate (IntPtr hWnd, IntPtr lParam) {
        StringBuilder className = new StringBuilder(256);
        GetClassName(hWnd, className, 256);
        if (className.ToString() != "screenClass" || !IsWindowVisible(hWnd)) return true;

        StringBuilder title = new StringBuilder(512);
        GetWindowText(hWnd, title, 512);
        if (title.ToString().Contains(fileName)) { found = hWnd; return false; }
        if (found == IntPtr.Zero) found = hWnd;
        return true;
    }, IntPtr.Zero);
    return found;
}
'@

# window positions are physical pixels
[FreeShow.Win32]::SetProcessDPIAware() | Out-Null

$script:app = $null
$script:pres = $null
$script:startedApp = $false
$script:openedByUs = $false
# full path -> presentation opened in the background, ready to start
$script:preloaded = @{}

function Send($obj) {
    $stdout.WriteLine((ConvertTo-Json -InputObject $obj -Compress -Depth 8))
}

# PowerPoint stays open for the life of this bridge; quit on exit only if the bridge started it
function Get-App {
    if ($script:app) { return $script:app }
    try {
        $script:app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    } catch {
        $script:app = New-Object -ComObject PowerPoint.Application
        $script:startedApp = $true
    }
    return $script:app
}

function Find-Presentation($fullPath) {
    # files in a OneDrive-synced folder (e.g. Documents) report a web address as FullName, so match those by file name
    $name = [IO.Path]::GetFileName($fullPath)
    $synced = $null
    foreach ($p in (Get-App).Presentations) {
        if ($p.FullName -eq $fullPath) { return $p }
        if (-not $synced -and $p.FullName -match '^https?://' -and $p.Name -eq $name) { $synced = $p }
    }
    return $synced
}

function Get-View {
    if (-not $script:pres) { return $null }
    try { return $script:pres.SlideShowWindow.View } catch { return $null }
}

function Get-State {
    $view = Get-View
    if (-not $view) { return @{ running = $false } }

    $slides = [int]$script:pres.Slides.Count
    $position = 0
    $clickIndex = 0
    $clickCount = 0
    try { $position = [int]$view.Slide.SlideIndex } catch { $position = $slides }
    try { $clickIndex = [int]$view.GetClickIndex(); $clickCount = [int]$view.GetClickCount() } catch {}

    return @{ running = $true; position = $position; slides = $slides; state = [int]$view.State; clickIndex = $clickIndex; clickCount = $clickCount }
}

function Get-Info {
    $titles = @()
    $notes = @()
    foreach ($slide in $script:pres.Slides) {
        $title = ''
        $note = ''
        try { if ($slide.Shapes.HasTitle) { $title = [string]$slide.Shapes.Title.TextFrame.TextRange.Text } } catch {}
        try { $note = [string]$slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text } catch {}
        $titles += $title
        $notes += $note
    }
    return @{ titles = $titles; notes = $notes }
}

# make the slide show window cover the output and hide it from the taskbar / alt+tab
function Set-ShowWindow($bounds) {
    # SlideShowWindow.HWND is 0 for presentations opened without a window, so look it up
    $hwnd = [IntPtr]::Zero
    for ($i = 0; $i -lt 30; $i++) {
        $hwnd = [FreeShow.Win32]::FindSlideShowWindow([string]$script:pres.Name)
        if ($hwnd -ne [IntPtr]::Zero) { break }
        Start-Sleep -Milliseconds 100
    }
    if ($hwnd -eq [IntPtr]::Zero) { return }

    $GWL_EXSTYLE = -20
    $WS_EX_TOOLWINDOW = 0x80
    $WS_EX_APPWINDOW = 0x40000
    $style = [FreeShow.Win32]::GetWindowLong($hwnd, $GWL_EXSTYLE)
    $style = ($style -bor $WS_EX_TOOLWINDOW) -band (-bnot $WS_EX_APPWINDOW)
    [FreeShow.Win32]::ShowWindow($hwnd, 0) | Out-Null
    [FreeShow.Win32]::SetWindowLong($hwnd, $GWL_EXSTYLE, $style) | Out-Null

    # SWP_NOACTIVATE (0x10) | SWP_SHOWWINDOW (0x40), HWND_TOPMOST = -1
    $flags = 0x50
    if ($bounds) {
        [FreeShow.Win32]::SetWindowPos($hwnd, [IntPtr](-1), [int]$bounds.x, [int]$bounds.y, [int]$bounds.width, [int]$bounds.height, $flags) | Out-Null
    } else {
        # SWP_NOMOVE (0x2) | SWP_NOSIZE (0x1)
        [FreeShow.Win32]::SetWindowPos($hwnd, [IntPtr](-1), 0, 0, 0, 0, ($flags -bor 0x3)) | Out-Null
    }
}

# ends the slide show; a presentation that was preloaded stays open for the next start
function Close-Presentation {
    if ($script:pres) {
        try { $script:pres.SlideShowWindow.View.Exit() } catch {}
        $isPreloaded = $false
        try { $isPreloaded = $script:preloaded.ContainsKey([string]$script:pres.FullName) } catch {}
        if ($script:openedByUs -and -not $isPreloaded) { try { $script:pres.Close() } catch {} }
    }
    $script:pres = $null
    $script:openedByUs = $false
}

function Open-Presentation($data) {
    Close-Presentation

    $fullPath = [IO.Path]::GetFullPath([string]$data.path)
    $script:pres = Find-Presentation $fullPath
    if (-not $script:pres) {
        # ReadOnly = msoTrue, Untitled = msoFalse, WithWindow = msoFalse (no editing window)
        $script:pres = (Get-App).Presentations.Open($fullPath, -1, 0, 0)
        $script:openedByUs = $true
    }

    $settings = $script:pres.SlideShowSettings
    $settings.ShowType = 1 # ppShowTypeSpeaker (full screen)
    $settings.RangeType = 1 # ppShowAll
    try { $settings.ShowPresenterView = 0 } catch {}
    $settings.Run() | Out-Null

    Set-ShowWindow $data.bounds

    return @{ state = (Get-State); info = (Get-Info); slideWidth = [double]$script:pres.PageSetup.SlideWidth; slideHeight = [double]$script:pres.PageSetup.SlideHeight }
}

# keep exactly these files open in the background (without windows), so starting them is instant
function Set-Preload($paths) {
    $wanted = @{}
    foreach ($p in @($paths)) { if ($p) { $wanted[[IO.Path]::GetFullPath([string]$p)] = $true } }

    foreach ($key in @($script:preloaded.Keys)) {
        $p = $script:preloaded[$key]
        # a file without a window is closed by PowerPoint when its slide show ends, and the old reference doesn't tell
        $isOpen = [bool](Find-Presentation $key)
        if ($isOpen -and -not $wanted.ContainsKey($key)) {
            $isShowing = $false
            try { $isShowing = $script:pres -and ($script:pres.FullName -eq $key) } catch {}
            if (-not $isShowing) { try { $p.Close() } catch {} }
        }
        if (-not $isOpen -or -not $wanted.ContainsKey($key)) { $script:preloaded.Remove($key) }
    }

    foreach ($key in @($wanted.Keys)) {
        if ($script:preloaded.ContainsKey($key)) { continue }
        if (-not (Test-Path -LiteralPath $key)) { continue }
        $open = Find-Presentation $key
        if ($open) {
            # open without a window = opened by this bridge (e.g. after its slide show): keep tracking it
            # with a window = opened by the user: leave it alone
            if ($open.Windows.Count -eq 0) { $script:preloaded[$key] = $open }
            continue
        }
        $script:preloaded[$key] = (Get-App).Presentations.Open($key, -1, 0, 0)
    }

    # slide count per requested path, in the same order (0 when it could not be opened)
    $slides = @()
    foreach ($p in @($paths)) {
        $count = 0
        if ($p) { $count = Get-SlideCount ([IO.Path]::GetFullPath([string]$p)) }
        $slides += $count
    }
    return @{ slides = $slides }
}

# PowerPoint rejects calls for a moment after a slide show ends, so retry briefly
function Get-SlideCount($fullPath) {
    for ($i = 0; $i -lt 5; $i++) {
        try { return [int](Find-Presentation $fullPath).Slides.Count } catch { Start-Sleep -Milliseconds 200 }
    }
    return 0
}

function Close-All {
    Close-Presentation
    foreach ($p in @($script:preloaded.Values)) { try { $p.Close() } catch {} }
    $script:preloaded = @{}
    if ($script:app -and $script:startedApp) {
        try { if ($script:app.Presentations.Count -eq 0) { $script:app.Quit() } } catch {}
    }
}

function Invoke-Action($action, $data) {
    if ($action -eq 'open') { return Open-Presentation $data }
    if ($action -eq 'close') { Close-Presentation; return @{ running = $false } }
    if ($action -eq 'state') { return Get-State }

    if ($action -eq 'preload') { return Set-Preload $data.paths }

    if ($action -eq 'thumbnail') {
        $target = Find-Presentation ([IO.Path]::GetFullPath([string]$data.file))
        if (-not $target) { throw 'presentation is not open' }
        $slide = $target.Slides.Item([int]$data.index)
        $width = [int]$data.width
        $height = [int][Math]::Round($width * $target.PageSetup.SlideHeight / $target.PageSetup.SlideWidth)
        $slide.Export([string]$data.path, 'PNG', $width, $height)
        return @{ path = [string]$data.path }
    }

    $view = Get-View
    if (-not $view) { return Get-State }

    switch ($action) {
        'next' {
            $atEnd = $false
            try { $atEnd = ($view.Slide.SlideIndex -ge $script:pres.Slides.Count) -and ($view.GetClickIndex() -ge $view.GetClickCount()) } catch {}
            if (-not $atEnd) { $view.Next() }
        }
        'previous' { $view.Previous() }
        'first' { $view.First() }
        'last' { $view.Last() }
        'goto' { $view.GotoSlide([int]$data.index) }
        'black' { if ($view.State -eq 3) { $view.State = 1 } else { $view.State = 3 } }
        default { throw ('unknown action: ' + $action) }
    }

    # PowerPoint can pause the show when it loses focus, which freezes animations
    try { if ($action -ne 'black' -and $view.State -eq 2) { $view.State = 1 } } catch {}

    return Get-State
}

while ($true) {
    $line = $stdin.ReadLine()
    if ($null -eq $line) { break }
    if (-not $line.Trim()) { continue }

    $request = $null
    try {
        $request = ConvertFrom-Json $line
        $result = Invoke-Action $request.action $request.data
        Send @{ id = $request.id; ok = $true; result = $result }
    } catch {
        $id = $null
        if ($request) { $id = $request.id }
        Send @{ id = $id; ok = $false; error = $_.Exception.Message }
    }
}

Close-All
`
