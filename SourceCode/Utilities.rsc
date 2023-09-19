/*
    Macro that creates an empty using the IDs from the TAZ file
    Input (option array):
        TAZFile
        DataType
        Label
        OutputFile
        Cores
*/
Macro "Create Empty Matrix"(mSpec)
    dm = CreateObject("DataManager")
    {vwTAZ} = dm.AddDataSource("TAZ", {FileName: mSpec.TAZFile, DataType: "DB"})
    vID = GetDataVector(vwTAZ + "|", "TAZID",)
    tazIDs = SortArray(v2a(vID))

    obj = CreateObject("Matrix") 
    obj.SetMatrixOptions({Compressed: 1, DataType: mSpec.DataType, MatrixLabel: mSpec.Label})
    obj.MatrixFileName = mSpec.OutputFile
    opts = {RowIDs: tazIDs, ColIDs: tazIDs, MatrixNames: mSpec.Cores, RowIndexName: "TAZ", ColIndexName: "TAZ"}
    mat = obj.CreateFromArrays(opts)
    dm = null
    Return(mat)
endMacro


/*
    Macro to create Intrazonal file
*/
Macro "Compute Intrazonal Matrix"(Args)
    mSpec = {TAZFile: Args.TAZGeography, DataType: "Short", Label: "Intrazonal", Cores: {"IZ"}, OutputFile: Args.IZMatrix}
    mat = RunMacro("Create Empty Matrix", mSpec)
    
    mIZ = CreateObject("Matrix", mat)
    mIZ.IZ := 0
    v = mIZ.GetVector({Core: 'IZ', Diagonal: 'Row'})
    v = v + 1
    mIZ.SetVector({Core: 'IZ', Vector: v, Diagonal: 'Row'})
    mat = null

    return(true)
endMacro


Macro "NotImplemented"(Args)
    ShowMessage("Not yet implemented")
    Return(1)
endMacro
