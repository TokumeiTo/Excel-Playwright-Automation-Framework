param(
    [int]$TotalTests,
    [string]$TempFile
)

Add-Type -AssemblyName PresentationFramework

# Create window
$window = New-Object System.Windows.Window
$window.Title = "Test Progress"
$window.Width = 400
$window.Height = 100
$window.Topmost = $true
$window.WindowStartupLocation = "CenterScreen"   # <-- This centers the window

# Create progress bar
$progressBar = New-Object System.Windows.Controls.ProgressBar
$progressBar.Minimum = 0
$progressBar.Maximum = $TotalTests
$progressBar.Width = 350
$progressBar.Height = 25
$progressBar.Margin = "20,20,20,20"
$progressBar.Value = 0

# StackPanel
$stack = New-Object System.Windows.Controls.StackPanel
$stack.Children.Add($progressBar)
$window.Content = $stack

# Timer updates every 200ms
$timer = [System.Windows.Threading.DispatcherTimer]::new()
$timer.Interval = [TimeSpan]::FromMilliseconds(200)
$timer.Add_Tick({

    if (Test-Path $TempFile) {
        $val = Get-Content $TempFile
        $num = 0  # <-- Initialize before using [ref]
        if ([int]::TryParse($val, [ref]$num)) {
            $progressBar.Value = $num
        }
    }

    if ($progressBar.Value -ge $TotalTests) {
        Set-Content -Path $TempFile -Value "DONE"
        $window.Close()
    }
})

$timer.Start()

# Show window (blocks until closed)
$window.ShowDialog() | Out-Null
