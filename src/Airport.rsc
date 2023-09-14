/*
Called by flowchart
*/

Macro "Airport Model" (Args)
    RunMacro("Airport Generation", Args)
    RunMacro("Airport Gravity", Args)
    // RunMacro("Airport TOD", Args)
    return(1)
endmacro

/*
Airport generation. Productions are specified as inputs in the AirportTrips
field in the SE data.
*/

Macro "Airport Generation" (Args)

    se_file = Args.DemographicOutputs
    rate_file = Args.[Airport Attraction Rates]
    vis_ratio = Args.[Airport Vis Ratio]

    se = CreateObject("Table", se_file)
    se_vw = se.GetView()

    // Productions
    se.AddField("air_vis_p")
    se.air_vis_p = se.AirportTrips * vis_ratio
    se.AddField("air_res_p")
    se.air_res_p = se.AirportTrips * (1 - vis_ratio)

    // Attractions    
    {drive, folder, name, ext} = SplitPath(rate_file)
    RunMacro("Create Sum Product Fields", {
        view: se_vw, factor_file: rate_file,
        field_desc: "Airport Attractions|See " + name + ext + " for details."
    })
EndMacro

/*
Prepares arguments for the "Gravity" macro in the utils.rsc library.
*/

Macro "Airport Gravity" (Args)

    se_file = Args.DemographicOutputs
    param_file = Args.[Airport Gravity Params]
    out_dir = Args.[Output Folder]
    air_dir = out_dir + "/airport"
    RunMacro("Create Directory", air_dir)
    

    RunMacro("Gravity", {
        se_file: se_file,
        skim_file: out_dir + "/skims/HighwaySkimAM.mtx",
        param_file: param_file,
        output_matrix: air_dir + "/air_gravity.mtx"
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
