# ControlBartenderViaCom
minimum necessary code to control BarTender (by Seagull Scientific Inc.) via COM a.k.a. ActiveX

## Target Environment
- [BarTender] 10.1 (Automation edition or Enterprise Automation edition)
- .NET Framework 4.5

## How to use
This Class Library controls just Start/End of BarTender and its managed reference, so if you want control formats and more, create a new derived class like below

    Class Hoge : ControlBartenderBase
    {
    	//control something
    }
Then you use this class in using satement

    using (var foo = new Hoge())
    {
    	var bar = foo.StartBartenderAsync();
    	bar.Wait();
    	//do something
    }
Or Dispose() finally

    var foo = new Hoge();
    try {
    	var bar = foo.StartBartenderAsync();
		bar.Wait();
		//do something
	}
	finally {
		foo.Dispose();
	}
 
[BarTender]: http://www.seagullscientific.com/label-software/barcode-label-design-and-printing.aspx