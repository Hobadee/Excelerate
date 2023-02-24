<#
.SYNOPSIS
Class to throw to simulate abstract methods

.DESCRIPTION
We should throw this class as an exception when faking abstract methods

#>
class AbstractException : InvalidOperationException{

    AbstractException() : base("Cannot call an abstract reference."){
        #
    }
}
