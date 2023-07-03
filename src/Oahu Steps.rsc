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
  Args.Demographics = Args.[Input Folder] + "\\sedata\\scenario_se.dbd"
  Args.[Master Links] = Args.[Master Folder] + "\\networks\\master_links.dbd"
  Args.HighwayInputDatabase = Args.[Input Folder] + "\\networks\\scenario_links.dbd"
  Args.[Master Routes] = Args.[Master Folder] + "\\networks\\master_routes.rts"
  Args.TransitRouteInputs = Args.[Input Folder] + "\\networks\\scenario_routes.rts"

  RunMacro("Create Scenario", Args)
  // RunMacro("PopulationSynthesis Oahu", Args)
  // RunMacro("Network Calculations", Args)
  return(1)
endmacro