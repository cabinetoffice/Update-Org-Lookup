# Read in original org list ####

library(magrittr)

in_location <- "~/Codes/Update-Org-Lookup/inputs/List of Organisations - GCS Data Audit 2022.xlsx"

df_raw <-
  readxl::read_excel(
    in_location,
    "Organisations Table"
  )

out_location <-
  "~/Codes/Update-Org-Lookup/outputs"

wb <- openxlsx::createWorkbook()

# Make changes ####

## v 1.01 - Add an About page and version history ####

version_str <- "1.01"

sheet_name_about = "About"

about <-
  tibble::tibble(
    `ABOUT` = c("This file lists all organisations in government and the departments to which they report. It was originally formulated from the database underlying the following gov.uk page: https://www.gov.uk/government/organisations",
      "This document was made to supplement the GCS Data Audit 2022. Queries should be directed to reshapinggcs@cabinetoffice.gov.uk",
      "The table below shows a version history for this document.")
  )

version_history <- tibble::tibble(
  "Version" = c("1",
                "1.01"),
  "Description and Changes" = c("The initial list sent with the commission.",
                                "Add an 'About' page describing the dataset and version history.")
  )

## v 1.02 - Fill in top-level organisations ####

version_str <- "1.02"

version_history <-
  version_history %>%
  tibble::add_row(
    Version = version_str,
    `Description and Changes` = "Complete entries for top-level organisations - in the original list, top-level organisations had incomplete entries which caused unintuitive behaviour of filters."
    )

df_intermediate <-
  df_raw %>%
  dplyr::mutate(
    `Top-level sponsor organisation` = dplyr::case_when(
      is.na(`Top-level sponsor organisation`) ~ Organisation,
      T ~ `Top-level sponsor organisation`
    ),
    `Top-level sponsor organisation ID (API)` = dplyr::case_when(
      is.na(`Top-level sponsor organisation ID (API)`) ~ `ID (API)`,
      T ~ `Top-level sponsor organisation ID (API)`
    ),
    `Top-level sponsor organisation slug (readable ID)` = dplyr::case_when(
      is.na(`Top-level sponsor organisation slug (readable ID)`) ~ `Slug (readable ID)`,
      T ~ `Top-level sponsor organisation slug (readable ID)`
    ),
    `Top-level sponsor organisation abbreviation` = dplyr::case_when(
      is.na(`Top-level sponsor organisation abbreviation`) ~ Abbreviation,
      T ~ `Top-level sponsor organisation abbreviation`
    )
  )

## v 1.03 - MOD org changes ####

mod_changes_location <-
  "~/Codes/Update-Org-Lookup/inputs/MOD changes - List of Organisations - GCS Data Audit 2022.xlsx"

mod_changes <-
  readxl::read_excel(
    mod_changes_location,
    sheet = 1
  )

orgs_to_remove <-
  mod_changes %>%
  dplyr::filter(remove) %>%
  dplyr::pull(`Slug (readable ID)`)

orgs_to_add <-
  mod_changes %>%
  dplyr::filter(add) %>%
  dplyr::select(-c(add, remove))

df_intermediate <-
  df_intermediate %>%
  dplyr::filter(
    !(`Slug (readable ID)` %in% orgs_to_remove)
  ) %>%
  dplyr::bind_rows(orgs_to_add)

version_str <- "1.03"

version_history <-
  version_history %>%
  tibble::add_row(
    Version = version_str,
    `Description and Changes` = "Change the list of MOD organisations - following discussions with MOD colleagues, MOD will report using the organisational structure defined in the Defence Operating Model rather than that defined on gov.uk. Some MOD organisations have been added/removed to accommodate this."
  )

## v 1.04 - Remove Directly Operated Railways Limited - no longer exists ####

dorl_changes_location <-
  "~/Codes/Update-Org-Lookup/inputs/DORL changes - List of Organisations - GCS Data Audit 2022.xlsx"

dorl_changes <-
  readxl::read_excel(
    dorl_changes_location,
    sheet = 1
  )

orgs_to_remove <-
  dorl_changes %>%
  dplyr::filter(remove) %>%
  dplyr::pull(`Slug (readable ID)`)

df_intermediate <-
  df_intermediate %>%
  dplyr::filter(
    !(`Slug (readable ID)` %in% orgs_to_remove)
  )

version_str <- "1.04"

version_history <-
  version_history %>%
  tibble::add_row(
    Version = version_str,
    `Description and Changes` = "Remove the organisation Directly Operated Railways Limited - it no longer exists."
  )

## v 1.05 - Defra org changes - Remove Defra core departments and add Defra Group ####

defra_changes_location <-
  "~/Codes/Update-Org-Lookup/inputs/Defra changes - List of Organisations - GCS Data Audit 2022.xlsx"

defra_changes <-
  readxl::read_excel(
    defra_changes_location,
    sheet = 1
  )

orgs_to_remove <-
  defra_changes %>%
  dplyr::filter(remove) %>%
  dplyr::pull(`Slug (readable ID)`)

orgs_to_add <-
  defra_changes %>%
  dplyr::filter(add) %>%
  dplyr::select(-c(add, remove))

df_intermediate <-
  df_intermediate %>%
  dplyr::filter(
    !(`Slug (readable ID)` %in% orgs_to_remove)
  ) %>%
  dplyr::bind_rows(orgs_to_add)

version_str <- "1.05"

version_history <-
  version_history %>%
  tibble::add_row(
    Version = version_str,
    `Description and Changes` = "Change the list of Defra organisations and remove duplication of MOD organisations - following discussions with Defra colleagues, Defra will have the Defra Group organisation added and the Defra core department removed. This is to accommodate the Defra Group structure. Defra Group includes the following organisations: Defra department, Environment agency, Natural England, Forestry Commission, Animal, Plant Health Agency and Rural Payments agency. While the Defra department has been removed from the list, the other 5 will remain. This is to give flexibility to Defra when completing the return. People attributed to Defra Group may spend time in one or multiple of the organisations across Defra Group. The duplication of MOD organisations was due to a bug which has now been fixed."
    )

## v 1.06 - MoJ org changes ####

moj_changes_location <-
  "~/Codes/Update-Org-Lookup/inputs/MOJ changes - List of Organisations - GCS Data Audit 2022.xlsx"

moj_changes <-
  readxl::read_excel(
    moj_changes_location,
    sheet = 1
  )

orgs_to_remove <-
  moj_changes %>%
  dplyr::filter(remove) %>%
  dplyr::pull(`Slug (readable ID)`)

orgs_to_add <-
  moj_changes %>%
  dplyr::filter(add) %>%
  dplyr::select(-c(add, remove))

df_intermediate <-
  df_intermediate %>%
  dplyr::filter(
    !(`Slug (readable ID)` %in% orgs_to_remove)
  ) %>%
  dplyr::bind_rows(orgs_to_add)

version_str <- "1.06"

version_history <-
  version_history %>%
  tibble::add_row(
    Version = version_str,
    `Description and Changes` = "Change the list of MoJ organisations - following discussions with MoJ colleagues, many organisations are removed and one is added. Most of the removed organisations are courts of some kind and will report through HMCTS. The added organisation is Assessor of Compensation for Miscarriages of Justice."
  )

## v 1.07 - BEIS org change - add Insolvency Rules Committee to BEIS (previously with MOJ) ####

beis_changes_location <-
  "~/Codes/Update-Org-Lookup/inputs/BEIS changes - List of Organisations - GCS Data Audit 2022.xlsx"

beis_changes <-
  readxl::read_excel(
    beis_changes_location,
    sheet = 1
  )

orgs_to_remove <-
  beis_changes %>%
  dplyr::filter(remove) %>%
  dplyr::pull(`Slug (readable ID)`)

orgs_to_add <-
  beis_changes %>%
  dplyr::filter(add) %>%
  dplyr::select(-c(add, remove))

df_intermediate <-
  df_intermediate %>%
  dplyr::filter(
    !(`Slug (readable ID)` %in% orgs_to_remove)
  ) %>%
  dplyr::bind_rows(orgs_to_add)

version_str <- "1.07"

version_history <-
  version_history %>%
  tibble::add_row(
    Version = version_str,
    `Description and Changes` = "Add IRC to BEIS organisations - following discussions with MoJ colleagues, the Insolvency Rules Committee top level organisation has been moved from MOJ to BEIS."
  )

## v 1.08 - MoJ org change - add Design102 Government Shared Services ####

moj_2_changes_location <-
  "~/Codes/Update-Org-Lookup/inputs/MOJ changes 2 - List of Organisations - GCS Data Audit 2022.xlsx"

moj_2_changes <-
  readxl::read_excel(
    moj_2_changes_location,
    sheet = 1
  )

orgs_to_remove <-
  moj_2_changes %>%
  dplyr::filter(remove) %>%
  dplyr::pull(`Slug (readable ID)`)

orgs_to_add <-
  moj_2_changes %>%
  dplyr::filter(add) %>%
  dplyr::select(-c(add, remove))

df_intermediate <-
  df_intermediate %>%
  dplyr::filter(
    !(`Slug (readable ID)` %in% orgs_to_remove)
  ) %>%
  dplyr::bind_rows(orgs_to_add)

version_str <- "1.08"

version_history <-
  version_history %>%
  tibble::add_row(
    Version = version_str,
    `Description and Changes` = "Add Design 102 to MoJ organisations - following discussions with MoJ colleagues, it is imperative the D102 - being a cross government service - is recorded separately to the MoJ core department."
  )

## v 1.09 - FCDO org change - remove Chevening as they are part of FCDO core department ####

fcdo_changes_location <-
  "~/Codes/Update-Org-Lookup/inputs/FCDO changes - List of Organisations - GCS Data Audit 2022.xlsx"

fcdo_changes <-
  readxl::read_excel(
    fcdo_changes_location,
    sheet = 1
  )

orgs_to_remove <-
  fcdo_changes %>%
  dplyr::filter(remove) %>%
  dplyr::pull(`Slug (readable ID)`)

orgs_to_add <-
  fcdo_changes %>%
  dplyr::filter(add) %>%
  dplyr::select(-c(add, remove))

df_intermediate <-
  df_intermediate %>%
  dplyr::filter(
    !(`Slug (readable ID)` %in% orgs_to_remove)
  ) %>%
  dplyr::bind_rows(orgs_to_add)

version_str <- "1.09"

version_history <-
  version_history %>%
  tibble::add_row(
    Version = version_str,
    `Description and Changes` = "Remove Chevening Scholarship Programme - following discussions with FCDO colleagues Chevening has been removed as it is part of the FCDO core department."
  )

## v 1.10 - Remove BBC World Service from FCDO and attribute it to DCMS ####

fcdo_dcms_changes_location <-
  "~/Codes/Update-Org-Lookup/inputs/FCDO and DCMS changes - List of Organisations - GCS Data Audit 2022 OFFICIAL.xlsx"

fcdo_dcms_changes_location <-
  readxl::read_excel(
    fcdo_dcms_changes_location,
    sheet = 1
  )

orgs_to_remove <-
  fcdo_dcms_changes_location %>%
  dplyr::filter(remove) %>%
  dplyr::pull(`Slug (readable ID)`)

orgs_to_add <-
  fcdo_dcms_changes_location %>%
  dplyr::filter(add) %>%
  dplyr::select(-c(add, remove))

df_intermediate <-
  df_intermediate %>%
  dplyr::filter(
    !(`Slug (readable ID)` %in% orgs_to_remove)
  ) %>%
  dplyr::bind_rows(orgs_to_add)

version_str <- "1.10"

version_history <-
  version_history %>%
  tibble::add_row(
    Version = version_str,
    `Description and Changes` = "Remove BBC World Service from FCDO and attribute it to DCMS  - following discussions with FCDO colleagues BBC World Service has been reallocated to DCMS. FCDO suggest that BBC WS may be included in the main BBC return."
  )

## v 1.11 - Defra changes - remove BCMS, FCERM R&D Programme, FFC and RDPE Network ####

defra_2_changes_location <-
  "~/Codes/Update-Org-Lookup/inputs/Defra changes 2 - List of Organisations - GCS Data Audit 2022 OFFICIAL.xlsx"

defra_2_changes <-
  readxl::read_excel(
    defra_2_changes_location,
    sheet = 1
  )

orgs_to_remove <-
  defra_2_changes %>%
  dplyr::filter(remove) %>%
  dplyr::pull(`Slug (readable ID)`)

orgs_to_add <-
  defra_2_changes %>%
  dplyr::filter(add) %>%
  dplyr::select(-c(add, remove))

df_intermediate <-
  df_intermediate %>%
  dplyr::filter(
    !(`Slug (readable ID)` %in% orgs_to_remove)
  ) %>%
  dplyr::bind_rows(orgs_to_add)

version_str <- "1.11"

version_history <-
  version_history %>%
  tibble::add_row(
    Version = version_str,
    `Description and Changes` = "Remove BCMS, FCERM R&D Programme, FFC and RDPE Network from Defra's organsiations - BCMS has merged with RPA. RDPE is a network and not a body. FCERM R&D Programme is a programme within the Environment Agency. FFC isn’t a Defra public body, but it may sit under the Met Office."
  )

## v 1.12 - DIT changes - remove all DIT ALBs except Trade Remedies Authority ####

dit_changes_location <-
  "~/Codes/Update-Org-Lookup/inputs/DIT changes - List of Organisations v1.11 - GCS Data Audit 2022 OFFICIAL.xlsx"

dit_changes <-
  readxl::read_excel(
    dit_changes_location,
    sheet = 1
  )

orgs_to_remove <-
  dit_changes %>%
  dplyr::filter(remove) %>%
  dplyr::pull(`Slug (readable ID)`)

orgs_to_add <-
  dit_changes %>%
  dplyr::filter(add) %>%
  dplyr::select(-c(add, remove))

df_intermediate <-
  df_intermediate %>%
  dplyr::filter(
    !(`Slug (readable ID)` %in% orgs_to_remove)
  ) %>%
  dplyr::bind_rows(orgs_to_add)

version_str <- "1.12"

version_history <-
  version_history %>%
  tibble::add_row(
    Version = version_str,
    `Description and Changes` = "Remove the following orgs from DIT's ALBS: Export Control Joint Unit, Life Sciences, Office for Investment, UK Defence and Security Exports, UK National Contact Point - Following discussion with DIT colleagues these were identified as all part of the DIT central organisation."
  )

## v 1.13 - Defra changes - Remove Veterinary Products COmmittee and the Science Advisory Council from Defra's ALBs ####

defra_changes_3_location <-
  "~/Codes/Update-Org-Lookup/inputs/Defra changes 3 - List of Organisations - GCS Data Audit 2022 OFFICIAL.xlsx"

defra_changes_3 <-
  readxl::read_excel(
    defra_changes_3_location,
    sheet = 1
  )

orgs_to_remove <-
  defra_changes_3 %>%
  dplyr::filter(remove) %>%
  dplyr::pull(`Slug (readable ID)`)

orgs_to_add <-
  defra_changes_3 %>%
  dplyr::filter(add) %>%
  dplyr::select(-c(add, remove))

df_intermediate <-
  df_intermediate %>%
  dplyr::filter(
    !(`Slug (readable ID)` %in% orgs_to_remove)
  ) %>%
  dplyr::bind_rows(orgs_to_add)

version_str <- "1.13"

version_history <-
  version_history %>%
  tibble::add_row(
    Version = version_str,
    `Description and Changes` = "Remove Veterinary Products Committee and Science Advisory Council - Defra colleagues indicate that VPC is part of the Veterinary Medicines Directorate. The SAC is part of core Defra."
  )

## v 1.14 - Defra changes - Add the Defra core department back in ####

defra_changes_4_location <-
  "~/Codes/Update-Org-Lookup/inputs/Defra changes 4 - List of Organisations - GCS Data Audit 2022.xlsx"

defra_changes_4 <-
  readxl::read_excel(
    defra_changes_4_location,
    sheet = 1
  )

orgs_to_remove <-
  defra_changes_4 %>%
  dplyr::filter(remove) %>%
  dplyr::pull(`Slug (readable ID)`)

orgs_to_add <-
  defra_changes_4 %>%
  dplyr::filter(add) %>%
  dplyr::select(-c(add, remove))

df_intermediate <-
  df_intermediate %>%
  dplyr::filter(
    !(`Slug (readable ID)` %in% orgs_to_remove)
  ) %>%
  dplyr::bind_rows(orgs_to_add)

version_str <- "1.14"

version_history <-
  version_history %>%
  tibble::add_row(
    Version = version_str,
    `Description and Changes` = "Add Defra core department - Following discussion with Defra colleagues, the core department has been re-added in."
  )

## v 1.15 - DCMS changes - Remove the following as they are either public corporations out of scope of the exercise, or small boards that sit within ALBS: BBC, BBC World Service, Channel 4, Historic Royal Palaces, S4C, The Advisory Council on National Records and Archives, The Reviewing Committee on the Export of Works of Art and Objects of Cultural Interest, The Theatres Trust, Treasure Valuation Committee ####

dcms_changes_location <-
  "~/Codes/Update-Org-Lookup/inputs/DCMS changes - List of Organisations - GCS Data Audit 2022.xlsx"

dcms_changes <-
  readxl::read_excel(
    dcms_changes_location,
    sheet = 1
  )

orgs_to_remove <-
  dcms_changes %>%
  dplyr::filter(remove) %>%
  dplyr::pull(`Slug (readable ID)`)

orgs_to_add <-
  dcms_changes %>%
  dplyr::filter(add) %>%
  dplyr::select(-c(add, remove))

df_intermediate <-
  df_intermediate %>%
  dplyr::filter(
    !(`Slug (readable ID)` %in% orgs_to_remove)
  ) %>%
  dplyr::bind_rows(orgs_to_add)

version_str <- "1.15"

version_history <-
  version_history %>%
  tibble::add_row(
    Version = version_str,
    `Description and Changes` = "Remove some DCMS organisations - Following discussion with DCMS colleagues, the following organisations have been removed as they are either public corporations out of scope of the exercise, or small boards that sit within ALBS: BBC, BBC World Service, Channel 4, Historic Royal Palaces, S4C, The Advisory Council on National Records and Archives, The Reviewing Committee on the Export of Works of Art and Objects of Cultural Interest, The Theatres Trust, Treasure Valuation Committee"
  )

## v 1.16 - Devolved Administration changes - Remove the devolved administrations ####

devolved_changes_location <-
  "~/Codes/Update-Org-Lookup/inputs/Devolved Administration changes - GCS Data Audit 2022 OFFICIAL.xlsx"

devolved_changes <-
  readxl::read_excel(
    devolved_changes_location,
    sheet = 1
  )

orgs_to_remove <-
  devolved_changes %>%
  dplyr::filter(remove) %>%
  dplyr::pull(`Slug (readable ID)`)

orgs_to_add <-
  devolved_changes %>%
  dplyr::filter(add) %>%
  dplyr::select(-c(add, remove))

df_intermediate <-
  df_intermediate %>%
  dplyr::filter(
    !(`Slug (readable ID)` %in% orgs_to_remove)
  ) %>%
  dplyr::bind_rows(orgs_to_add)

version_str <- "1.16"

version_history <-
  version_history %>%
  tibble::add_row(
    Version = version_str,
    `Description and Changes` = "Remove the devolved administration - Devolved Administrations are out of scope of the GCS Data Audit."
  )

## v 1.17 - Change abbreviation for Scotland Office - Change the abbreviation used for Scotland Office from OOTSOSSFS and OSSS to SO ####

scotland_changes_location <-
  "~/Codes/Update-Org-Lookup/inputs/Scotland Office changes - GCS Data Audit 2022 OFFICIAL.xlsx"

scotland_changes <-
  readxl::read_excel(
    scotland_changes_location,
    sheet = 1
  )

orgs_to_remove <-
  scotland_changes %>%
  dplyr::filter(remove) %>%
  dplyr::pull(`Slug (readable ID)`)

orgs_to_add <-
  scotland_changes %>%
  dplyr::filter(add) %>%
  dplyr::select(-c(add, remove))

df_intermediate <-
  df_intermediate %>%
  dplyr::filter(
    !(`Slug (readable ID)` %in% orgs_to_remove)
  ) %>%
  dplyr::bind_rows(orgs_to_add)

version_str <- "1.17"

version_history <-
  version_history %>%
  tibble::add_row(
    Version = version_str,
    `Description and Changes` = "Change the abbreviation used for the Scotland Office."
  )

## v 1.18 - Change abbreviation for Home Office - Change the abbreviation used for Scotland Office from Home Office to HO ####

home_office_changes_location <-
  "~/Codes/Update-Org-Lookup/inputs/Home Office changes - GCS Data Audit 2022 OFFICIAL.xlsx"

home_office_changes <-
  readxl::read_excel(
    home_office_changes_location,
    sheet = 1
  )

orgs_to_remove <-
  home_office_changes %>%
  dplyr::filter(remove) %>%
  dplyr::pull(`Slug (readable ID)`)

orgs_to_add <-
  home_office_changes %>%
  dplyr::filter(add) %>%
  dplyr::select(-c(add, remove))

df_intermediate <-
  df_intermediate %>%
  dplyr::filter(
    !(`Slug (readable ID)` %in% orgs_to_remove)
  ) %>%
  dplyr::bind_rows(orgs_to_add)

version_str <- "1.18"

version_history <-
  version_history %>%
  tibble::add_row(
    Version = version_str,
    `Description and Changes` = "Change the abbreviation used for the Home Office."
  )

## v 1.19 - Change abbreviation for Wales Office - Change the abbreviation used for Wales Office from UK Government in Wales to WO ####

wales_office_changes_location <-
  "~/Codes/Update-Org-Lookup/inputs/Wales Office changes - GCS Data Audit 2022 OFFICIAL.xlsx"

wales_office_changes <-
  readxl::read_excel(
    wales_office_changes_location,
    sheet = 1
  )

orgs_to_remove <-
  wales_office_changes %>%
  dplyr::filter(remove) %>%
  dplyr::pull(`Slug (readable ID)`)

orgs_to_add <-
  wales_office_changes %>%
  dplyr::filter(add) %>%
  dplyr::select(-c(add, remove))

df_intermediate <-
  df_intermediate %>%
  dplyr::filter(
    !(`Slug (readable ID)` %in% orgs_to_remove)
  ) %>%
  dplyr::bind_rows(orgs_to_add)

version_str <- "1.19"

version_history <-
  version_history %>%
  tibble::add_row(
    Version = version_str,
    `Description and Changes` = "Change the abbreviation used for the UK Government in Wales."
  )

## v 1.20 - DCMS changes - Remove some DCMS organisations - Following conversations with DCMS colleagues, the following organisations were included with the main DCMS submission and therefore the organisations have been removed: Office for Artificial Intelligence and Centre for Data Ethics and Innovation. ####

dcms_changes_2_location <-
  "~/Codes/Update-Org-Lookup/inputs/DCMS changes 2 - List of Organisations - GCS Data Audit 2022 OFFICIAL.xlsx"

dcms_changes_2 <-
  readxl::read_excel(
    dcms_changes_2_location,
    sheet = 1
  )

orgs_to_remove <-
  dcms_changes_2 %>%
  dplyr::filter(remove) %>%
  dplyr::pull(`Slug (readable ID)`)

orgs_to_add <-
  dcms_changes_2 %>%
  dplyr::filter(add) %>%
  dplyr::select(-c(add, remove))

df_intermediate <-
  df_intermediate %>%
  dplyr::filter(
    !(`Slug (readable ID)` %in% orgs_to_remove)
  ) %>%
  dplyr::bind_rows(orgs_to_add)

version_str <- "1.20"

version_history <-
  version_history %>%
  tibble::add_row(
    Version = version_str,
    `Description and Changes` = "Remove some DCMS organisations - Following conversations with DCMS colleagues, the following organisations were included with the main DCMS submission and therefore the organisations have been removed: Office for Artificial Intelligence and Centre for Data Ethics and Innovation."
  )

## v 1.21 - DHSC changes - Add Healthwatch England to DHSC's list of organisations - Following conversations with DHSC colleagues. ####

dhsc_changes_location <-
  "~/Codes/Update-Org-Lookup/inputs/DHSC changes - List of Organisations - GCS Data Audit 2022 OFFICIAL.xlsx"

dhsc_changes <-
  readxl::read_excel(
    dhsc_changes_location,
    sheet = 1
  )

orgs_to_remove <-
  dhsc_changes %>%
  dplyr::filter(remove) %>%
  dplyr::pull(`Slug (readable ID)`)

orgs_to_add <-
  dhsc_changes %>%
  dplyr::filter(add) %>%
  dplyr::select(-c(add, remove))

df_intermediate <-
  df_intermediate %>%
  dplyr::filter(
    !(`Slug (readable ID)` %in% orgs_to_remove)
  ) %>%
  dplyr::bind_rows(orgs_to_add)

version_str <- "1.21"

version_history <-
  version_history %>%
  tibble::add_row(
    Version = version_str,
    `Description and Changes` = "Add Healthwatch England to DHSC's list of organisations."
  )


# Write latest version ####

df_final <-
  df_intermediate %>%
  dplyr::arrange(Organisation)

sysdatetime <- Sys.time() %>%
  format("%Y-%m-%d_%H-%M-%S_%Z")

sheet_name_table = "Organisations Table"

openxlsx::addWorksheet(wb, sheet_name_table)

openxlsx::writeDataTable(wb, sheet = sheet_name_table, df_final)

openxlsx::addWorksheet(wb, sheet_name_about)

openxlsx::writeData(wb, sheet = sheet_name_about, x = about)

openxlsx::writeData(wb, sheet = sheet_name_about, x = version_history, startRow = nrow(about) + 3)

openxlsx::saveWorkbook(
  wb,
  file = paste0(out_location, "/", sysdatetime, " List of Organisations v", version_str, " - GCS Data Audit 2022 OFFICIAL.xlsx")
)
