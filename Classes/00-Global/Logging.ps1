class Logging{


    static trace($e){
        Write-Debug($e.ToString())
        Write-Debug($e.ScriptStackTrace)
    }


    static Debug($msg){
        try{
            # Throw an exception so we can grab a stack trace
            throw [System.Diagnostics.Tracing.EventSourceException]::new
        }
        catch{
            $stack = $_.ScriptStackTrace.split("`n")
            $stack2 = @()
            for($i=1; $i -le $stack.length; $i++){
                $stack2 += $stack[$i]
            }
            $stack = $stack2 -join "`n"

            Write-Debug("-----")
            Write-Debug($msg)
            Write-Debug($stack)
        }
    }
}
