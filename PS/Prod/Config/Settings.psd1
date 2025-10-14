@{
    # Desktop Management Logon/Logoff Suite - Main Configuration
    # Version 2.0 - PowerShell Implementation
    
    Version = '2.0.0'
    
    # Mapper Service Configuration
    Mapper = @{
        ProductionServer = 'gdpmappercb.nomura.com'
        QAServer = 'gdpmappercbqa.nomura.com'
        ServicePath = 'ClassicMapper.asmx'
        Timeout = 10000  # milliseconds
        EnableHealthCheck = $True
    }
    
    # Inventory Service Configuration
    Inventory = @{
        ProductionServer = 'gdpmappercb.nomura.com'
        QAServer = 'gdpmappercbqa.nomura.com'
        ServicePath = 'ClassicInventory.asmx'
        Timeout = 10000  # milliseconds
        EnableHealthCheck = $True
    }
    
    # Logging Configuration
    Logging = @{
        BasePath = '$env:USERPROFILE\Nomura\GDP\Desktop Management'
        MaxAge = 60  # days
        VerboseDefault = $False
        BufferSize = 100  # lines before flush
        TimestampFormat = 'yyyyMMddHHmmss'
    }
    
    # Registry Configuration
    Registry = @{
        BasePath = 'HKCU:\Software\Nomura\GDP\Desktop Management'
        TrackExecution = $True
        StoreVersion = $True
    }
    
    # Performance Configuration
    Performance = @{
        EnableParallelProcessing = $False  # For future optimization
        MaxConcurrentOperations = 5
        RetryAttempts = 3
        RetryDelayMs = 1000
    }
    
    # Feature Toggles (Quick disable without code changes)
    Features = @{
        EnableInventory = $True
        EnableMapper = $True
        EnablePowerConfig = $True
        EnablePasswordNotification = $True
        EnableRetailFeatures = $True
        EnableLegacyIEZones = $False  # Deprecated
    }
}

