:: 
:: This is an example BAT that shows how to use GOMRemote
:: 
:: 
:: Command Line Options:
:: -fullscreen : Makes GOMPlayer fullscreen.
:: -maximize : Maximizes the GOMPlayer window.
:: -exitaftermovie : Exits GOMPlayer after the current movie ends.
:: -exitafterplaylist : Exits GOMPlayer after the current playlist ends.
:: 
:: Created by Agret, alias.zero0297@gmail.com
:: 

@start GOMRemote.exe -fullscreen -exitaftermovie
@"%windir%\clock.avi"