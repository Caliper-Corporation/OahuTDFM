/*
Called by flowchart
*/

Macro "Airport" (Args)
    RunMacro("Airport Productions/Attractions", Args)
    RunMacro("Airport Gravity", Args)
    // RunMacro("Airport TOD", Args)
    return(1)
endmacro

/*
CV productions
Attractions are the same as productions
*/

Macro "Airport Productions/Attractions" (Args)

    se_file = Args.DemographicOutputs
    rate_file = Args.[Airport Attraction Rates]

    se = CreateObject("Table", se_file)
    se_vw = se.GetView()
    
    {drive, folder, name, ext} = SplitPath(rate_file)
    RunMacro("Create Sum Product Fields", {
        view: se_vw, factor_file: rate_file,
        field_desc: "CV Productions and Attractions|See " + name + ext + " for details."
    })
EndMacro

/*
Prepares arguments for the "Gravity" macro in the utils.rsc library.
*/

Macro "Airport Gravity" (Args)

    se_file = Args.DemographicOutputs
    param_file = Args.[Airport Gravity Params]
    out_dir = Args.[Output Folder]
    

    RunMacro("Gravity", {
        se_file: se_file,
        skim_file: out_dir + "/skims/HighwaySkimAM.mtx",
        param_file: param_file,
        output_matrix: out_dir + "/airport/air_gravity.mtx"
    })
EndMacro

/*
Split CV productions and attractions into time periods
*/

Macro "Airport TOD" (Args)

    se_file = Args.SE
    rate_file = Args.[CV TOD Rates]

    se_vw = OpenTable("se", "FFB", {se_file})
    {drive, folder, name, ext} = SplitPath(rate_file)
    RunMacro("Create Sum Product Fields", {
        view: se_vw, factor_file: rate_file,
        field_desc: "CV Productions and Attractions by Time of Day|See " + name + ext + " for details."
    })

    CloseView(se_vw)
endmacro
