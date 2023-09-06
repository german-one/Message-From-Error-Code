# Copyright (c) Steffen Illhardt
# Licensed under the MIT license.

#region dependencies
Add-Type -AN PresentationFramework, System.Drawing, System.Windows.Forms # The latter only for `Timer` and `DoubleClickTime` in the triple-click handler.
Add-Type API -NS Win32 -MemberDefinition @'
  [DllImport("dwmapi.dll")]
  public static extern int DwmSetWindowAttribute(IntPtr wnd, int attrId, ref int attr, int attrSize);
  [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
  public static extern int FormatMessage(int flags, IntPtr src, int msgId, int langId, ref IntPtr msgPtr, int minLen, IntPtr args);
  [DllImport("kernel32.dll")]
  public static extern IntPtr GetConsoleWindow();
  [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
  public static extern IntPtr GetModuleHandle(string fileName);
  [DllImport("kernel32.dll")]
  public static extern IntPtr LocalFree(IntPtr mem);
  [DllImport("shcore.dll")]
  public static extern int SetProcessDpiAwareness(int value);
  [DllImport("user32.dll")]
  public static extern int ShowWindowAsync(IntPtr wnd, int state);
'@
#endregion dependencies

[void][Win32.API]::ShowWindowAsync([Win32.API]::GetConsoleWindow(), 0 <#SW_HIDE#>) # Try to hide the CLI as we only use the GUI.
try { [void][Win32.API]::SetProcessDpiAwareness(2 <#PROCESS_PER_MONITOR_DPI_AWARE#>) } catch {} # Avoid blurry text.

#region GUI elements
$mainWnd = [Windows.Markup.XamlReader]::Parse(@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Message From Error Code" ResizeMode="CanMinimize"
        SizeToContent="WidthAndHeight"
        FocusManager.FocusedElement="{x:Reference codeTxtBx}">
    <Window.Resources>
        <x:Double x:Key="C0Left">20</x:Double><!-- left canvas padding, X-pos. of the first column of controls -->
        <x:Double x:Key="C0Width">125</x:Double><!-- code -->
        <x:Double x:Key="C1Left">155</x:Double><!-- type -->
        <x:Double x:Key="C1Width">185</x:Double>
        <x:Double x:Key="CMergedWidth">320</x:Double><!-- language and message -->
        <x:Double x:Key="CanvasWidth">360</x:Double>
        <x:Double x:Key="CanvasHalfWidth">180</x:Double>
        <x:Double x:Key="R0Top">5</x:Double><!-- top canvas padding, Y-pos. of the first row of labels -->
        <x:Double x:Key="R1Top">31</x:Double><!-- code and type boxes -->
        <x:Double x:Key="R2Top">61</x:Double><!-- language label -->
        <x:Double x:Key="R3Top">87</x:Double><!-- language box -->
        <x:Double x:Key="R4Top">117</x:Double><!-- message label -->
        <x:Double x:Key="R0To4Height">25</x:Double>
        <x:Double x:Key="R5Top">143</x:Double><!-- message box -->
        <x:Double x:Key="R5Height">95</x:Double>
        <x:Double x:Key="R6Top">248</x:Double><!-- facility -->
        <x:Double x:Key="R7Top">268</x:Double><!-- severty and provider -->
        <x:Double x:Key="R8Top">288</x:Double><!-- status bar -->
        <x:Double x:Key="R6To8Height">20</x:Double>
        <x:Double x:Key="CanvasHeight">308</x:Double>
        <x:Double x:Key="LabelFontSize">12</x:Double>
        <x:Double x:Key="BoxFontSize">14</x:Double><!-- both TextBoxes and ComboBoxes -->
        <Color x:Key="AnthraciteBlack">#0f1012</Color>
        <Color x:Key="AnthraciteShade">#1f2023</Color>
        <Color x:Key="AnthraciteTint">#3f4144</Color>
        <LinearGradientBrush x:Key="AnthraciteGradientBrush" StartPoint="0,0" EndPoint="1,1">
            <GradientStop Color="{StaticResource AnthraciteShade}" Offset="0.0" />
            <GradientStop Color="{StaticResource AnthraciteTint}" Offset="0.5" />
        </LinearGradientBrush>
        <SolidColorBrush x:Key="AnthraciteBlackBrush" Color="{StaticResource AnthraciteBlack}" />
        <SolidColorBrush x:Key="AnthraciteDarkBrush" Color="{StaticResource AnthraciteShade}" />
        <Style x:Key="LabelStyle" TargetType="Label">
            <Setter Property="Height" Value="{StaticResource R0To4Height}" />
            <Setter Property="FontSize" Value="{StaticResource LabelFontSize}" />
            <Setter Property="Foreground" Value="White" />
        </Style>
        <Style x:Key="ParameterBoxStyle" TargetType="{x:Type Control}">
            <Setter Property="Height" Value="{StaticResource R0To4Height}" />
            <Setter Property="FontSize" Value="{StaticResource BoxFontSize}" />
            <Setter Property="VerticalContentAlignment" Value="Center" />
        </Style>
        <Style x:Key="FooterStyle" TargetType="Label">
            <Setter Property="Height" Value="{StaticResource R6To8Height}" />
            <Setter Property="VerticalContentAlignment" Value="Center" />
            <Setter Property="Padding" Value="23,0,23,0" />
        </Style>
    </Window.Resources>
    <Canvas Width="{StaticResource CanvasWidth}" Height="{StaticResource CanvasHeight}"
            Background="{StaticResource AnthraciteGradientBrush}">
        <Label Content="Error Code:"
               Canvas.Left="{StaticResource C0Left}" Canvas.Top="{StaticResource R0Top}"
               Width="{StaticResource C0Width}"
               Style="{StaticResource LabelStyle}" />
        <TextBox x:Name="codeTxtBx"
                 Canvas.Left="{StaticResource C0Left}" Canvas.Top="{StaticResource R1Top}"
                 Width="{StaticResource C0Width}" MaxLength="20"
                 HorizontalContentAlignment="Right" Padding="3,0,3,0"
                 Style="{StaticResource ParameterBoxStyle}">
            <TextBox.ToolTip>
                <ToolTip Content="Enter the Windows error code as&#10; - (sign prefixed) decimal integer&#10; - '0x' prefixed hexadecimal number"
                         Background="White" />
            </TextBox.ToolTip>
            <TextBox.ContextMenu>
                <ContextMenu>
                    <MenuItem Command="ApplicationCommands.Cut" />
                    <MenuItem Command="ApplicationCommands.Copy" />
                    <MenuItem Command="ApplicationCommands.Paste" />
                    <MenuItem Command="ApplicationCommands.SelectAll" />
                </ContextMenu>
            </TextBox.ContextMenu>
        </TextBox>
        <Label Content="Code Type:"
               Canvas.Left="{StaticResource C1Left}" Canvas.Top="{StaticResource R0Top}"
               Width="{StaticResource C1Width}"
               Style="{StaticResource LabelStyle}" />
        <ComboBox x:Name="typeCoBx"
                  Canvas.Left="{StaticResource C1Left}" Canvas.Top="{StaticResource R1Top}"
                  Width="{StaticResource C1Width}"
                  Style="{StaticResource ParameterBoxStyle}"
                  SelectedIndex="0">
            <ComboBoxItem Content="Win32 Error or HRESULT" />
            <ComboBoxItem Content="NTSTATUS" />
            <ComboBoxItem Content="HRESULT from NTSTATUS" />
        </ComboBox>
        <Label Content="Message Language:"
               Canvas.Left="{StaticResource C0Left}" Canvas.Top="{StaticResource R2Top}"
               Width="{StaticResource CMergedWidth}"
               Style="{StaticResource LabelStyle}" />
        <ComboBox x:Name="langCoBx"
                  Canvas.Left="{StaticResource C0Left}" Canvas.Top="{StaticResource R3Top}"
                  Width="{StaticResource CMergedWidth}"
                  Style="{StaticResource ParameterBoxStyle}" />
        <Label Content="Message Text:"
               Canvas.Left="{StaticResource C0Left}" Canvas.Top="{StaticResource R4Top}"
               Width="{StaticResource CMergedWidth}"
               Style="{StaticResource LabelStyle}" />
        <TextBox x:Name="msgTxtBx"
                 Canvas.Left="{StaticResource C0Left}" Canvas.Top="{StaticResource R5Top}"
                 Width="{StaticResource CMergedWidth}" Height="{StaticResource R5Height}"
                 Padding="3,0,3,0"
                 FontSize="{StaticResource BoxFontSize}" Background="WhiteSmoke"
                 IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto">
            <TextBox.ContextMenu>
                <ContextMenu>
                    <MenuItem Command="ApplicationCommands.Copy" />
                    <MenuItem Command="ApplicationCommands.SelectAll" />
                </ContextMenu>
            </TextBox.ContextMenu>
        </TextBox>
        <Label Canvas.Top="{StaticResource R6Top}"
               Width="{StaticResource CanvasWidth}"
               Foreground="White" Background="{StaticResource AnthraciteDarkBrush}"
               Style="{StaticResource FooterStyle}">
            <TextBlock>
                <Hyperlink x:Name="facilityLink" ToolTip="No URI linked.">Facility</Hyperlink>
                <Run x:Name="facilityTxtRun" />
            </TextBlock>
        </Label>
        <Label x:Name="severityLbl"
               Canvas.Top="{StaticResource R7Top}"
               Width="{StaticResource CanvasHalfWidth}"
               Foreground="White" Background="{StaticResource AnthraciteDarkBrush}"
               Style="{StaticResource FooterStyle}" />
        <Label x:Name="providerLbl"
               Canvas.Left="{StaticResource CanvasHalfWidth}" Canvas.Top="{StaticResource R7Top}"
               Width="{StaticResource CanvasHalfWidth}"
               Foreground="White" Background="{StaticResource AnthraciteDarkBrush}"
               Style="{StaticResource FooterStyle}" />
        <Label x:Name="statusBarLbl"
               Canvas.Top="{StaticResource R8Top}"
               Foreground="WhiteSmoke" Background="{StaticResource AnthraciteBlackBrush}"
               Width="{StaticResource CanvasWidth}"
               Style="{StaticResource FooterStyle}" />
    </Canvas>
</Window>
'@)

'codeTxtBx', 'typeCoBx', 'langCoBx', 'msgTxtBx', 'facilityLink', 'facilityTxtRun', 'severityLbl', 'providerLbl', 'statusBarLbl' | ForEach-Object { Set-Variable $_ $mainWnd.FindName($_) }
#endregion GUI elements

#region lookup tables
# The data in this region is taken from "Rust for Windows" source code (https://github.com/microsoft/windows-rs) as this seems to contain the most complete set of those error constants that can be found on the internet.
# Copyright (c) Microsoft Corporation. Apache 2.0 and MIT licensed, see:
# https://github.com/microsoft/windows-rs/blob/master/license-apache-2.0
# https://github.com/microsoft/windows-rs/blob/master/license-mit
# The data has been rearranged and adapted for use in this code. Duplicate values have been removed because hash tables require unique keys.
# Common values are described in the Microsoft docs. Start from there: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-erref/1bc92ddf-b79e-413c-bbaa-99a5281a6c90

# source: https://github.com/microsoft/windows-rs/blob/master/crates/libs/windows/src/Windows/Win32/System/Diagnostics/Debug/mod.rs
$hResFacilities = @{
  0 = 'FACILITY_NULL'
  1 = 'FACILITY_RPC'
  2 = 'FACILITY_DISPATCH'
  3 = 'FACILITY_STORAGE'
  4 = 'FACILITY_ITF'
  7 = 'FACILITY_WIN32'
  8 = 'FACILITY_WINDOWS'
  9 = 'FACILITY_SSPI'
  10 = 'FACILITY_CONTROL'
  11 = 'FACILITY_CERT'
  12 = 'FACILITY_INTERNET'
  13 = 'FACILITY_MEDIASERVER'
  14 = 'FACILITY_MSMQ'
  15 = 'FACILITY_SETUPAPI'
  16 = 'FACILITY_SCARD'
  17 = 'FACILITY_COMPLUS'
  18 = 'FACILITY_AAF'
  19 = 'FACILITY_URT'
  20 = 'FACILITY_ACS'
  21 = 'FACILITY_DPLAY'
  22 = 'FACILITY_UMI'
  23 = 'FACILITY_SXS'
  24 = 'FACILITY_WINDOWS_CE'
  25 = 'FACILITY_HTTP'
  26 = 'FACILITY_USERMODE_COMMONLOG'
  27 = 'FACILITY_WER'
  31 = 'FACILITY_USERMODE_FILTER_MANAGER'
  32 = 'FACILITY_BACKGROUNDCOPY'
  33 = 'FACILITY_WIA'
  34 = 'FACILITY_STATE_MANAGEMENT'
  35 = 'FACILITY_METADIRECTORY'
  36 = 'FACILITY_WINDOWSUPDATE'
  37 = 'FACILITY_DIRECTORYSERVICE'
  38 = 'FACILITY_GRAPHICS'
  39 = 'FACILITY_NAP'
  40 = 'FACILITY_TPM_SERVICES'
  41 = 'FACILITY_TPM_SOFTWARE'
  42 = 'FACILITY_UI'
  43 = 'FACILITY_XAML'
  44 = 'FACILITY_ACTION_QUEUE'
  48 = 'FACILITY_PLA'
  49 = 'FACILITY_FVE'
  50 = 'FACILITY_FWP'
  51 = 'FACILITY_WINRM'
  52 = 'FACILITY_NDIS'
  53 = 'FACILITY_USERMODE_HYPERVISOR'
  54 = 'FACILITY_CMI'
  55 = 'FACILITY_USERMODE_VIRTUALIZATION'
  56 = 'FACILITY_USERMODE_VOLMGR'
  57 = 'FACILITY_BCD'
  58 = 'FACILITY_USERMODE_VHD'
  59 = 'FACILITY_USERMODE_HNS'
  60 = 'FACILITY_SDIAG'
  61 = 'FACILITY_WINPE'
  62 = 'FACILITY_WPN'
  63 = 'FACILITY_WINDOWS_STORE'
  64 = 'FACILITY_INPUT'
  65 = 'FACILITY_QUIC'
  66 = 'FACILITY_EAP'
  70 = 'FACILITY_IORING'
  80 = 'FACILITY_WINDOWS_DEFENDER'
  81 = 'FACILITY_OPC'
  82 = 'FACILITY_XPS'
  83 = 'FACILITY_RAS'
  84 = 'FACILITY_MBN'
  85 = 'FACILITY_EAS'
  98 = 'FACILITY_P2P_INT'
  99 = 'FACILITY_P2P'
  100 = 'FACILITY_DAF'
  101 = 'FACILITY_BLUETOOTH_ATT'
  102 = 'FACILITY_AUDIO'
  103 = 'FACILITY_STATEREPOSITORY'
  109 = 'FACILITY_VISUALCPP'
  112 = 'FACILITY_SCRIPT'
  113 = 'FACILITY_PARSE'
  120 = 'FACILITY_BLB'
  121 = 'FACILITY_BLB_CLI'
  122 = 'FACILITY_WSBAPP'
  128 = 'FACILITY_BLBUI'
  129 = 'FACILITY_USN'
  130 = 'FACILITY_USERMODE_VOLSNAP'
  131 = 'FACILITY_TIERING'
  133 = 'FACILITY_WSB_ONLINE'
  134 = 'FACILITY_ONLINE_ID'
  135 = 'FACILITY_DEVICE_UPDATE_AGENT'
  136 = 'FACILITY_DRVSERVICING'
  153 = 'FACILITY_DLS'
  160 = 'FACILITY_SOS'
  173 = 'FACILITY_OCP_UPDATE_AGENT'
  176 = 'FACILITY_DEBUGGERS'
  208 = 'FACILITY_DELIVERY_OPTIMIZATION'
  231 = 'FACILITY_USERMODE_SPACES'
  232 = 'FACILITY_USER_MODE_SECURITY_CORE'
  234 = 'FACILITY_USERMODE_LICENSING'
  256 = 'FACILITY_SPP'
  257 = 'FACILITY_DEPLOYMENT_SERVICES_SERVER'
  258 = 'FACILITY_DEPLOYMENT_SERVICES_IMAGING'
  259 = 'FACILITY_DEPLOYMENT_SERVICES_MANAGEMENT'
  260 = 'FACILITY_DEPLOYMENT_SERVICES_UTIL'
  261 = 'FACILITY_DEPLOYMENT_SERVICES_BINLSVC'
  263 = 'FACILITY_DEPLOYMENT_SERVICES_PXE'
  264 = 'FACILITY_DEPLOYMENT_SERVICES_TFTP'
  272 = 'FACILITY_DEPLOYMENT_SERVICES_TRANSPORT_MANAGEMENT'
  278 = 'FACILITY_DEPLOYMENT_SERVICES_DRIVER_PROVISIONING'
  289 = 'FACILITY_DEPLOYMENT_SERVICES_MULTICAST_SERVER'
  290 = 'FACILITY_DEPLOYMENT_SERVICES_MULTICAST_CLIENT'
  293 = 'FACILITY_DEPLOYMENT_SERVICES_CONTENT_PROVIDER'
  296 = 'FACILITY_HSP_SERVICES'
  297 = 'FACILITY_HSP_SOFTWARE'
  305 = 'FACILITY_LINGUISTIC_SERVICES'
  885 = 'FACILITY_WEB'
  886 = 'FACILITY_WEB_SOCKET'
  1094 = 'FACILITY_AUDIOSTREAMING'
  1490 = 'FACILITY_TTD'
  1536 = 'FACILITY_ACCELERATOR'
  1793 = 'FACILITY_MOBILE'
  1967 = 'FACILITY_SQLITE'
  1968 = 'FACILITY_SERVICE_FABRIC'
  1989 = 'FACILITY_UTC'
  1996 = 'FACILITY_WMAAECMA'
  2049 = 'FACILITY_WEP'
  2050 = 'FACILITY_SYNCENGINE'
  2168 = 'FACILITY_DIRECTMUSIC'
  2169 = 'FACILITY_DIRECT3D10'
  2170 = 'FACILITY_DXGI'
  2171 = 'FACILITY_DXGI_DDI'
  2172 = 'FACILITY_DIRECT3D11'
  2173 = 'FACILITY_DIRECT3D11_DEBUG'
  2174 = 'FACILITY_DIRECT3D12'
  2175 = 'FACILITY_DIRECT3D12_DEBUG'
  2176 = 'FACILITY_DXCORE'
  2177 = 'FACILITY_PRESENTATION'
  2184 = 'FACILITY_LEAP'
  2185 = 'FACILITY_AUDCLNT'
  2192 = 'FACILITY_WINML'
  2200 = 'FACILITY_WINCODEC_DWRITE_DWM'
  2201 = 'FACILITY_DIRECT2D'
  2304 = 'FACILITY_DEFRAG'
  2305 = 'FACILITY_USERMODE_SDBUS'
  2306 = 'FACILITY_JSCRIPT'
  2339 = 'FACILITY_XBOX'
  2340 = 'FACILITY_GAME'
  2561 = 'FACILITY_PIDGENX'
  2748 = 'FACILITY_PIX'
}

# source: https://github.com/microsoft/windows-rs/blob/master/crates/libs/windows/src/Windows/Win32/Foundation/mod.rs
$ntStatSeverities = @('Success', 'Informational', 'Warning', 'Error')

# source: https://github.com/microsoft/windows-rs/blob/master/crates/libs/windows/src/Windows/Win32/Foundation/mod.rs
$ntStatFacilities = @{
  1 = 'FACILITY_DEBUGGER'
  2 = 'FACILITY_RPC_RUNTIME'
  3 = 'FACILITY_RPC_STUBS'
  4 = 'FACILITY_IO_ERROR_CODE'
  5 = 'FACILITY_MCA_ERROR_CODE'
  6 = 'FACILITY_CODCLASS_ERROR_CODE'
  7 = 'FACILITY_NTWIN32'
  8 = 'FACILITY_NTCERT'
  9 = 'FACILITY_NTSSPI'
  10 = 'FACILITY_TERMINAL_SERVER'
  16 = 'FACILITY_USB_ERROR_CODE'
  17 = 'FACILITY_HID_ERROR_CODE'
  18 = 'FACILITY_FIREWIRE_ERROR_CODE'
  19 = 'FACILITY_CLUSTER_ERROR_CODE'
  20 = 'FACILITY_ACPI_ERROR_CODE'
  21 = 'FACILITY_SXS_ERROR_CODE'
  25 = 'FACILITY_TRANSACTION'
  26 = 'FACILITY_COMMONLOG'
  27 = 'FACILITY_VIDEO'
  28 = 'FACILITY_FILTER_MANAGER'
  29 = 'FACILITY_MONITOR'
  30 = 'FACILITY_GRAPHICS_KERNEL'
  32 = 'FACILITY_DRIVER_FRAMEWORK'
  33 = 'FACILITY_FVE_ERROR_CODE'
  34 = 'FACILITY_FWP_ERROR_CODE'
  35 = 'FACILITY_NDIS_ERROR_CODE'
  36 = 'FACILITY_QUIC_ERROR_CODE'
  41 = 'FACILITY_TPM'
  42 = 'FACILITY_RTPM'
  53 = 'FACILITY_HYPERVISOR'
  54 = 'FACILITY_IPSEC'
  55 = 'FACILITY_VIRTUALIZATION'
  56 = 'FACILITY_VOLMGR'
  57 = 'FACILITY_BCD_ERROR_CODE'
  62 = 'FACILITY_WIN32K_NTUSER'
  63 = 'FACILITY_WIN32K_NTGDI'
  64 = 'FACILITY_RESUME_KEY_FILTER'
  65 = 'FACILITY_RDBSS'
  66 = 'FACILITY_BTH_ATT'
  67 = 'FACILITY_SECUREBOOT'
  68 = 'FACILITY_AUDIO_KERNEL'
  69 = 'FACILITY_VSM'
  70 = 'FACILITY_NT_IORING'
  80 = 'FACILITY_VOLSNAP'
  81 = 'FACILITY_SDBUS'
  92 = 'FACILITY_SHARED_VHDX'
  93 = 'FACILITY_SMB'
  94 = 'FACILITY_XVS'
  153 = 'FACILITY_INTERIX'
  231 = 'FACILITY_SPACES'
  232 = 'FACILITY_SECURITY_CORE'
  233 = 'FACILITY_SYSTEM_INTEGRITY'
  234 = 'FACILITY_LICENSING'
  235 = 'FACILITY_PLATFORM_MANIFEST'
  236 = 'FACILITY_APP_EXEC'
}
#endregion lookup tables

#region language data collection
$langList = [Collections.Generic.SortedList[string, int]]::new()
# In the "System32" folder, look for directories with patterns such as "en-US" ...
Get-ChildItem ([Environment]::SystemDirectory) '??-??' -AD -Force | ForEach-Object {
  # ... that contain an "ntdll.dll.mui" file.
  if (Test-Path "$($_.FullName)\ntdll.dll.mui" -Type Leaf) {
    $cultInfo = [Globalization.CultureInfo]::new($_.Name) # Use the directory names to create `CultureInfo` objects ...
    $langList[$cultInfo.EnglishName] = $cultInfo.LCID # ... from which we collect the `EnglishName` and `LCID` properties into a list.
  }
}
#endregion language data collection

#region GUI event handling
  #region Win11 dark title bar
  $setDarkTitleBar = {
    $Win32_TRUE = 1
    [void][Win32.API]::DwmSetWindowAttribute([Windows.Interop.WindowInteropHelper]::new($this).Handle, 20 <#DWMWA_USE_IMMERSIVE_DARK_MODE#>, [ref]$Win32_TRUE, 4 <#SizeOf($Win32_TRUE)#>)
  }
  #endregion Win11 dark title bar

  #region output
  $ntDll = [Win32.API]::GetModuleHandle('ntdll.dll')
  $updateOutput = {
    $facilityLink.NavigateUri = ''
    $facilityLink.ToolTip = 'No URI linked.'
    $facilityTxtRun.Text = ': '
    $severityLbl.Content = 'Severity: '
    $providerLbl.Content = 'Provider: '
    $msgPtr = [IntPtr]::Zero
    $codeStr = ''
    try {
      $codeStr = $codeTxtBx.Text.Trim()
      $numBase = $(if ($codeStr.StartsWith('0x', 'OrdinalIgnoreCase')) { 16 } else { 10 })
      $codeNum = $converted = [Convert]::ToInt32($codeStr, $numBase)
      if ($typeCoBx.SelectedIndex -eq 0) {
        $flgs = 0x1300 # FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_IGNORE_INSERTS
        $msgSrc = [IntPtr]::Zero
        if (-not ($converted -bAnd 0xFFFF0000)) {
          $codeNum = $converted -bOr 0x80070000
        }

        $facilityLink.NavigateUri = 'https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-erref/0642cb2f-2075-4469-918c-4441e69c548a'
        $facilityLink.ToolTip = 'Common HRESULT facility values.'
        $facility = $hResFacilities[($codeNum -shr 16) -bAnd 0x1FFF] # Based on HRESULT_FACILITY macro. Ensures that both the "X" and "N" bits are 0.
        $severity = $(if ($converted -and ($codeNum -lt 0)) { 'Error' } else { 'Success' })
      }
      else {
        $flgs = 0x0B00 # FORMAT_MESSAGE_FROM_HMODULE | FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_IGNORE_INSERTS
        $msgSrc = $ntDll
        if ($typeCoBx.SelectedIndex -eq 2) {
          $codeNum = $(if ($converted -bAnd 0x10000000) { $converted -bAnd 0xEFFFFFFF } else { 0x10000000 }) # Either remove the "N" bit to get an "NTSTATUS" or explicitly invalidate $codeNum if it wasn't an "HRESULT from NTSTATUS".
        }

        $facilityLink.NavigateUri = 'https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-erref/1714a7aa-8e53-4076-8f8d-75073b780a41#Appendix_A_5'
        $facilityLink.ToolTip = 'Common NTSTATUS facility values.'
        $facility = $ntStatFacilities[($codeNum -shr 16) -bAnd 0x1FFF] # Even though the facility part is only 12 bits, we also check that the "N" bit is 0.
        $severity = $ntStatSeverities[($codeNum -shr 30) -bAnd 3]
      }

      $langId = $(if ($langCoBx.SelectedItem) { $langList[$langCoBx.SelectedItem] } else { 0x0000 }) # If no item is selected, fall back to "neutral" to follow the search order for a suitable language.
      $msgLength = [Win32.API]::FormatMessage($flgs, $msgSrc, $codeNum, $langId, [ref]$msgPtr, 0, [IntPtr]::Zero)
      $facilityTxtRun.Text += $(if ($facility) { $facility } else { 'unknown' })
      $severityLbl.Content += $severity
      $providerLbl.Content += $(if ($codeNum -bAnd 0x20000000) { 'customer' } else { 'Microsoft' })
      if ($msgLength) {
        $msgTxtBx.Text = [Runtime.InteropServices.Marshal]::PtrToStringAuto($msgPtr, $msgLength)
        $statusBarLbl.Content = "Message with code 0x$($converted.ToString('X8'))."
      }
      else {
        $msgTxtBx.Clear()
        $statusBarLbl.Content = "No message with code 0x$($converted.ToString('X8')) found."
      }
    }
    catch {
      $msgTxtBx.Clear()
      $statusBarLbl.Content = $(if ($codeStr) { 'Not convertible to a Windows error code.' } else { '' })
    }
    finally {
      [void][Win32.API]::LocalFree($msgPtr)
    }
  }
  #endregion output

  #region tabstop auto highlighting
  $highlightOnFocusedByTabKey = {
    if ([Windows.Input.Keyboard]::IsKeyDown('Tab')) {
      $this.SelectAll()
    }
  }
  #endregion tabstop auto highlighting

  #region triple-click highlighting
  $script:clickCount = 0
  $script:clickTimer = [Windows.Forms.Timer]::new()
  $script:clickTimer.Interval = [Windows.Forms.SystemInformation]::DoubleClickTime
  $script:clickTimer.Add_Tick({
      [Threading.Monitor]::Enter($this)
      $this.Stop()
      $script:clickCount = 0
      [Threading.Monitor]::Exit($this)
    })

  $highlightOnTripleClick = {
    [Threading.Monitor]::Enter($script:clickTimer)
    $script:clickTimer.Stop()
    ++$script:clickCount
    if ($script:clickCount -lt 3) {
      $this.SelectionLength = 0
      $script:clickTimer.Start()
    }
    else {
      $this.SelectAll()
      $script:clickCount = 0
      $_.Handled = $True
    }

    [Threading.Monitor]::Exit($script:clickTimer)
  }
  #endregion triple-click highlighting

  #region hyperlink navigation
  $execHyperlink = {
    if ($this.NavigateUri) {
      Start-Process $this.NavigateUri
    }
  }
  #endregion hyperlink navigation

$mainWnd.Add_Loaded($setDarkTitleBar)
$codeTxtBx.Add_GotFocus($highlightOnFocusedByTabKey)
$codeTxtBx.Add_PreviewMouseLeftButtonDown($highlightOnTripleClick)
$codeTxtBx.Add_TextChanged($updateOutput)
$typeCoBx.Add_SelectionChanged($updateOutput)
$langCoBx.Add_SelectionChanged($updateOutput)
$msgTxtBx.Add_GotFocus($highlightOnFocusedByTabKey)
$msgTxtBx.Add_PreviewMouseLeftButtonDown($highlightOnTripleClick)
$facilityLink.Add_PreviewMouseLeftButtonDown($execHyperlink)
#endregion GUI event handling

#region GUI dynamic properties
$mainWnd.Icon = [Windows.Interop.Imaging]::CreateBitmapSourceFromHIcon([Drawing.SystemIcons]::Information.Handle, [Windows.Int32Rect]::Empty, [Windows.Media.Imaging.BitmapSizeOptions]::FromEmptyOptions())
$langCoBx.ItemsSource = $langList.Keys
$langCoBx.SelectedItem = [Globalization.CultureInfo]::CurrentUICulture.EnglishName # If this fails, we silently fall back to "neutral" (LCID 0x0000).
#endregion GUI dynamic properties

[void]$mainWnd.ShowDialog()
