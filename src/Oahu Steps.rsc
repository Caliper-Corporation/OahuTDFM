/*

*/

Macro "Oahu Steps" (Args)

  Args.[Master Folder] = Args.[Base Folder] + "\\Master"
  Args.[Scenario Folder] = Args.[Base Folder] + "\\Scenarios\\base_2022"
  Args.[Input Folder] = Args.[Scenario Folder] + "\\Input"
  Args.[Output Folder] = Args.[Scenario Folder] + "\\Output"
  Args.[Master TAZs] = Args.[Master Folder] + "\\tazs\\master_tazs.dbd"
  Args.TAZGeography = Args.[Input Folder] + "\\tazs\\scenario_tazs.dbd"
  Args.[Master SE] = Args.[Master Folder] + "\\sedata\\base_2020.bin"
  Args.Demographics = Args.[Input Folder] + "\\sedata\\scenario_se.bin"
  Args.[Master Links] = Args.[Master Folder] + "\\networks\\master_links.dbd"
  Args.HighwayInputDatabase = Args.[Input Folder] + "\\networks\\scenario_links.dbd"
  Args.[Master Routes] = Args.[Master Folder] + "\\networks\\master_routes.rts"
  Args.TransitRouteInputs = Args.[Input Folder] + "\\networks\\scenario_routes.rts"
  RunMacro("Create Scenario", Args)

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
  RunMacro("PopulationSynthesis Oahu", Args)

  // RunMacro("Network Calculations", Args)
  return(1)
endmacro