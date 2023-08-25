/*

*/

Macro "Visitor Model" (Args)

    RunMacro("Visitor Lodging Locations", Args)
    RunMacro("Visitor Trip Generation", Args)
    RunMacro("Visitor Create MC Features", Args)
    RunMacro("Visitor Calculate MC", Args)
    return(1)
endmacro

/*
Determine visitor home/loding locations
*/

Macro "Visitor Lodging Locations" (Args)

    se_file = Args.DemographicOutputs

    se = CreateObject("Table", se_file)
    se.AddFields({Fields: {
        {FieldName: "visitors_p", Description: "Visitors in the zone on a personal trip to the island"},
        {FieldName: "visitors_b", Description: "Visitors in the zone on a business trip to the island"}
    }})

    se.visitors_b = se.OccupiedHH * Args.[Vis HH Occ Rate] * Args.[Vis HH Business Ratio] +
        se.HR * Args.[Vis Hotel Occ Rate] * Args.[Vis Hotel Business Ratio] + 
        se.RC * Args.[Vis Condo Occ Rate] * Args.[Vis Condo Business Ratio]
    se.visitors_b = se.visitors_b * Args.[Vis Business Party Size] * Args.[Vis Party Calibration Factor]
    se.visitors_p = se.OccupiedHH * Args.[Vis HH Occ Rate] * (1 - Args.[Vis HH Business Ratio]) +
        se.HR * Args.[Vis Hotel Occ Rate] * (1 - Args.[Vis Hotel Business Ratio]) + 
        se.RC * Args.[Vis Condo Occ Rate] * (1 - Args.[Vis Condo Business Ratio])
    se.visitors_p = se.visitors_p * Args.[Vis Personal Party Size] * Args.[Vis Party Calibration Factor]
endmacro

/*
Generate trips by purpose for business and personal visitors
*/

Macro "Visitor Trip Generation" (Args)
    se_file = Args.DemographicOutputs
    rate_file = Args.[Vis Trip Rates]

    se = CreateObject("Table", se_file)
    se_vw = se.GetView()
    {drive, folder, name, ext} = SplitPath(rate_file)
    RunMacro("Create Sum Product Fields", {
        view: se_vw, factor_file: rate_file,
        field_desc: "CV Productions and Attractions|See " + name + ext + " for details."
    })
endmacro

/*
Create any features needed by the visitor mc model
*/

Macro "Visitor Create MC Features" (Args)
    se_file = Args.DemographicOutputs

    se = CreateObject("Table", se_file)
    se.AddField("AreaTypeNum")
    se.AreaTypeNum = if se.AreaType = "Downtown" then 1
        else if se.AreaType = "Urban" then 2
        else if se.AreaType = "Suburban" then 3
        else if se.AreaType = "Rural" then 4
endmacro

/*
Loops over purposes and preps options for the "MC" macro
*/

Macro "Visitor Calculate MC" (Args)

    scen_dir = Args.[Scenario Folder]
    skims_dir = scen_dir + "\\output\\skims\\"
    input_dir = Args.[Input Folder]
    input_mc_dir = input_dir + "/visitors/mc"
    output_dir = Args.[Output Folder] + "/visitors/mc"
    periods = Args.TimePeriods
    mode_table = Args.TransitModeTable
    access_modes = Args.AccessModes
    se_file = Args.DemographicOutputs
    
    // Specify trip purposes and transit modes
    trip_types = {"HBEat", "HBO", "HBRec", "HBShop", "HBW", "NHB"}
    transit_modes = RunMacro("Get Transit Net Def Col Names", mode_table)
    pos = transit_modes.position("all")
    if pos > 0 then transit_modes = ExcludeArrayElements(transit_modes, pos, 1)

    opts = null
    opts.primary_spec = {Name: "bus_skim"}
    for trip_type in trip_types do
        opts.segments = {"personal", "business"}
        opts.trip_type = trip_type
        opts.util_file = input_mc_dir + "/" + trip_type + ".csv"
        nest_file = input_mc_dir + "/" + trip_type + "_nest.csv"
        if GetFileInfo(nest_file) <> null 
            then opts.nest_file = nest_file
            else opts.nest_file = null

        // Set sources
        opts.tables = {
            se: {File: se_file, IDField: "TAZ"}
        }
        for i = 1 to periods.length do
            period = periods[i][1]
            opts.period = period
            opts.matrices = {
                auto_skim: {File: skims_dir + "\\HighwaySkim" + period + ".mtx"},
                walk_skim: {File: skims_dir + "\\WalkSkim.mtx"},
                bike_skim: {File: skims_dir + "\\BikeSkim.mtx"},
                iz_skim: {File: skims_dir + "\\IntraZonal.mtx"}
            }
            // Transit skims depend on which modes are present in the scenario
            for transit_mode in transit_modes do
                source_name = transit_mode + "_skim"
                file_name = skims_dir + "transit\\" + period + "_w_" + transit_mode + ".mtx"
                if GetFileInfo(file_name) <> null then opts.matrices.(source_name) = {File: file_name}
            end

            opts.output_dir = output_dir
            RunMacro("MC", opts)
        end
    end
endmacro