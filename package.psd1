@{
    Root = 'c:\Users\aaron.haydon\OneDrive - Ground Control\Documents\MailboxManager\Main.ps1'
    OutputPath = 'C:\Users\aaron.haydon\OneDrive - Ground Control\Documents\MailboxManager\'
    Package = @{
        Enabled = $true
        Obfuscate = $false
        HideConsoleWindow = $false
        DotNetVersion = 'net8.0'
        PowershellVersion = '7.4.1'
        FileVersion = '1.1.0'
        FileDescription = 'Mailbox Manager Application'
        ProductName = 'Mailbox Manager'
        ProductVersion = '1.0'
        Copyright = 'Ground Control Ltd. 2024'
        RequireElevation = $false
        ApplicationIconPath = 'D:\Scripts\Scripts\GUI Projects\Mailbox Manager\Mailbox Manager.ico'
        PackageType = 'Console'
    }
    Bundle = @{
        Enabled = $true
        Modules = $true
        # IgnoredModules = @()
    }
}
        