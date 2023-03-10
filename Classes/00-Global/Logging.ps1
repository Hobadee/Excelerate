class Logging{


    <#
    .SYNOPSIS
    Prints out an exception trace when an exception has been caught

    .EXAMPLE
    Call from within a `catch` block using:
    [Logging]::Trace($_)
    #>
    static Exception($e){
        Write-Error($e.ToString())
        Write-Error($e.ScriptStackTrace)
    }


    <#
    #>
    static Debug($msg){
        try{
            # Throw an exception so we can grab a stack trace
            throw [System.Diagnostics.Tracing.EventSourceException]::new
        }
        catch{
            $stack = $_.ScriptStackTrace.split("`n")
            <#
            $stack2 = @()
            for($i=1; $i -le $stack.length; $i++){
                $stack2 += $stack[$i]
            }
            $stack = $stack2 -join "`n"
            #>

            Write-Debug("-----")
            Write-Debug($msg)
            Write-Debug($stack[1])
        }
    }


    <#
    .SYNOPSIS
    Print a debug message with full stack trace information
    #>
    static Trace($msg){
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
