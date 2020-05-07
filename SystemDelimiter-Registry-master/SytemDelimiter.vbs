Sub SystemDelimiter()
    m_sDecimal = readFromRegistry("HKEY_CURRENT_USER\Control Panel\International\sDecimal", "")
    m_strSystemDelimiter = readFromRegistry("HKEY_CURRENT_USER\Control Panel\International\sList", "")
    m_sAutoFlushCache = readFromRegistry("HKEY_CURRENT_USER\SOFTWARE\CWA\Foundation Classes\Debug Options\Auto Flush Cache", "0")
    m_sShortDate = readFromRegistry("HKEY_CURRENT_USER\Control Panel\International\sShortDate", "")
End Sub
