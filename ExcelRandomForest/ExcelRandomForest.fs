namespace ExcelRandomForest
[<AutoOpen>]
module RExcel =
    //open libraries
    open RDotNet
    open RProvider
    open RProvider.``base``
    open RProvider.stats
    open RProvider.randomForest
    open ExcelDna.Integration   

    [<ExcelFunction(Name="randomForestClassification", Description="Random Forest Classification", Category="RFunctions")>]
    let randomForestClassification([<ExcelArgument(Name="Feature value matrix", Description = "Matrix of Feature values for training (float[,])")>] x: float[,], 
                                   [<ExcelArgument(Name="Response vector", Description = "A response vector for training (string[] or integer[])")>](y: obj[]),
                                   [<ExcelArgument(Name="New feature value matrix",Description = "Matrix of Feature values for prediction (float[,])")>](x_new: float[,])) = 
        //obj array -> string
        let y_converted = Array.map (fun z -> string z) y
        //Random forest after convert string to factor
        let rf = R.randomForest(x, R.factor(y_converted))
        //don't forget boxing the string array
        box(R.as_character(R.predict(rf, x_new)).GetValue<string[]>() )
            
    [<ExcelFunction(Name="randomForestRegression", Description="Random Forest Regression", Category="RFunctions")>]
    let randomForestRegression([<ExcelArgument(Name="Feature value matrix", Description = "Matrix of Feature values for training (float[,])")>] x: float[,], 
                               [<ExcelArgument(Name="Response vector", Description = "A response vector for training (float[])")>](y: float[]),
                               [<ExcelArgument(Name="New feature value matrix",Description = "Matrix of Feature values for prediction (float[,])")>](x_new: float[,])) = 
        let rf = R.randomForest(x, y) 
        R.predict(rf, x_new).GetValue<float[]>()

    [<ExcelFunction(Name="rnorm", Description="Random generation for normal distribution", Category="RFunctions")>]
    let rnorm([<ExcelArgument(Name="n", Description = "Number of observations (int)")>] n:int, 
              [<ExcelArgument(Name="mean", Description = "Mean (float)")>] mean:float, 
              [<ExcelArgument(Name="sd", Description = "Standard Deviation (float)")>] sd:float) = 
        R.rnorm(n, mean, sd).GetValue<float[]>()

    [<ExcelFunction(Name = "gsub", Description = "Replace all matched string", Category="RFunctions")>]
    let gsub([<ExcelArgument(Name="pattern", Description = " Regular expression to be matched in the given string (string)")>] pattern:string, 
             [<ExcelArgument(Name="replacement", Description = "Replacement for matched pattern (string)")>] replacement:string, 
             [<ExcelArgument(Name="x", Description = "string where matches are sough (string)")>] x:string) = 
        R.gsub(pattern, replacement, x, false, true).GetValue<string>()
