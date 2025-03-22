# WUIntegrate

WUIntegrate is a piece of software that integrates the latest Windows update into a WIM or ISO file.

## Supported image versions
WUIntegrate supports Windows 7 and newer*.
>*Windows 8 is currently not supported.

## Known issues
- WUIntegrate may not find updates for Windows 11.
- WUIntegrate is potentially unstable and may not work correctly.

## How to use
WUIntegrate is supported on Windows 10, build 17763 and later.

WUIntegrate is a command-line tool. In a command prompt as administrator, and provide the necessary arguments:
```WUIntegrate.exe (path to WIM or ISO) (logging enabler boolean)```

> Logging enabler boolean takes ``true`` or ``false`` as arguments. If ``true``, WUIntegrate will output a logfile to ``%TEMP%\wuintegrate.log``.