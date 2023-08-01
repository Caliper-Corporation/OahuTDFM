/*

*/

Macro "Visitor Model" (Args)

    RunMacro("Visitor Lodging Locations", Args)
    RunMacro("Visitor Trip Generation", Args)
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