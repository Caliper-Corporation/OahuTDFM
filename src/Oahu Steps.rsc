/*

*/

Macro "Oahu Steps" (Args)

  // Pre-process box

  Args.[Master Folder] = Args.[Base Folder] + "\\Master"
  Args.[Scenario Folder] = Args.[Base Folder] + "\\Scenarios\\base_2022"
  Args.[Input Folder] = Args.[Scenario Folder] + "\\Input"
  Args.[Output Folder] = Args.[Scenario Folder] + "\\Output"
  Args.[Master TAZs] = Args.[Master Folder] + "\\tazs\\master_tazs.dbd"
  Args.TAZGeography = Args.[Input Folder] + "\\tazs\\scenario_tazs.dbd"
  Args.[Master SE] = Args.[Master Folder] + "\\sedata\\base_2020.bin"
  Args.Demographics = Args.[Input Folder] + "\\sedata\\scenario_se.bin"
  Args.DemographicOutputs = Args.[Output Folder] + "\\sedata\\scenario_se.bin"
  Args.[Master Links] = Args.[Master Folder] + "\\networks\\master_links.dbd"
  Args.HighwayInputDatabase = Args.[Input Folder] + "\\networks\\scenario_links.dbd"
  Args.HighwayDatabase = Args.[Output Folder] + "\\networks\\scenario_links.dbd"
  Args.[Master Routes] = Args.[Master Folder] + "\\networks\\master_routes.rts"
  Args.TransitRouteInputs = Args.[Input Folder] + "\\networks\\scenario_routes.rts"
  Args.TransitRoutes = Args.[Output Folder] + "\\networks\\scenario_routes.rts"
  // RunMacro("Create Scenario", Args)

  Args.SEDMarginals = Args.[Output Folder] + "\\Population\\SEDMarginals.bin"
  Args.SizeCurves = Args.[Input Folder] + "\\Population\\disagg_model\\size_curves.csv"
  Args.IncomeCurves = Args.[Input Folder] + "\\Population\\disagg_model\\income_curves.csv"
  Args.WorkerCurves = Args.[Input Folder] + "\\Population\\disagg_model\\worker_curves.csv"
  Args.RegionalMedianIncome = 92600 // Source 2021 ACS Honolulu County
  Args.PUMS_Households = Args.[Input Folder] + "\\Population\\Seed\\HHSeed_PUMS.bin"
  Args.PUMS_Persons = Args.[Input Folder] + "\\Population\\Seed\\PersonSeed_PUMS.bin"
  Args.Households = Args.[Output Folder] + "\\Population\\Households.bin"
  Args.Persons = Args.[Output Folder] + "\\Population\\Persons.bin"
  Args.PopSynTolerance = .001
  Args.[Synthesized Tabulations] = Args.[Output Folder] + "\\Population\\Tabulations.bin"
  // RunMacro("PopulationSynthesis Oahu", Args)

  Args.SpeedCapacityLookup = Args.[Input Folder] + "\\networks\\speed_and_capacity.csv"
  Args.AreaTypes = {
    {AreaType: "Rural", Density: 0, Buffer: 0},
    {AreaType: "Suburban", Density: 1000, Buffer: .5},
    {AreaType: "Urban", Density: 10000, Buffer: .5},
    {AreaType: "Downtown", Density: 25000, Buffer: .25}
  }
  Args.IZMatrix = Args.[Output Folder] + "\\Skims\\IntraZonal.mtx"
  // RunMacro("Network Calculations", Args)

  // Build Networks box

  Args.HighwayNetwork = Args.[Output Folder] + "\\Skims\\highwaynet.net"
  Args.TransitNetwork = Args.[Output Folder] + "\\Skims\\TransitNetwork.tnw"
  RunMacro("BuildNetworks Oahu", Args)

  Args.HighwaySkimAM = Args.[Output Folder] + "\\Skims\\Highway_AM.mtx"
  Args.HighwaySkimPM = Args.[Output Folder] + "\\Skims\\Highway_PM.mtx"
  Args.HighwaySkimOP = Args.[Output Folder] + "\\Skims\\Highway_OP.mtx"
  Args.WalkSkim = Args.[Output Folder] + "\\Skims\\Walk.mtx"
  Args.BikeSkim = Args.[Output Folder] + "\\Skims\\Bike.mtx"
  // RunMacro("HighwayAndTransitSkim Oahu", Args)

  // Vistors
  Args.[Vis Hotel Occ Rate] = .846
  Args.[Vis Condo Occ Rate] = .846
  Args.[Vis HH Occ Rate] = .022
  Args.[Vis Hotel Business Ratio] = .088
  Args.[Vis Condo Business Ratio] = .063
  Args.[Vis HH Business Ratio] = .012
  Args.[Vis Personal Party Size] = 2.64
  Args.[Vis Business Party Size] = 1.74
  Args.[Vis Party Calibration Factor] = 1
  Args.[Vis Trip Rates] = Args.[Input Folder] + "\\visitors\\vis_generation.csv"
  RunMacro("Visitor Model", Args)
  
  return(1)
endmacro