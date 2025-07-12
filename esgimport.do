// Change working directory to ESG data folder
cd "C:\Users\FM\OneDrive\ESG"

// Initialize empty tempfile for merging
clear
tempfile master_data
save `master_data', emptyok

// Loop through years from 2024 to 2018
forvalues year = 2024(-1)2015 {
    // Construct filename for each year
    local filename "wind_esg_results_`year'.xlsx"
    
    // Check if file exists
    capture confirm file "`filename'"
    if _rc == 0 {
        // Import Excel file if exists
        import excel "`filename'", sheet("Sheet1") firstrow clear
        
        // Rename first column to stock_code if it's named A
        capture rename A stock_code
        
        // Append to master dataset
        append using `master_data',force
        save `master_data', replace
    }
    else {
        // Display message if file doesn't exist
        di "File `filename' does not exist, skipped."
    }
}

// Clean stock_code in ESG data
replace stock_code = subinstr(stock_code, ".SH", "", .)
replace stock_code = subinstr(stock_code, ".SZ", "", .)
replace stock_code = subinstr(stock_code, ".BJ", "", .)
replace stock_code = regexr(stock_code, "\.[A-Z]+$", "")

// Drop AA column if it exists
drop AA

// Save final merged ESG data
save "esg.dta", replace

* ==============================================
* ESG Variables Cleaning Script
* Purpose: Identify and drop completely empty ESG variables (100% missing)
* Last Updated: 2024-03-20
* ==============================================

* ----------------------------------------------
* Step 1: Define the ESG variable list
* ----------------------------------------------
local esg_var_list ///
ESG_CONCLUSION ESG_RATING_WIND ESG_SCORE_WIND ESG_MGMTSCORE_WIND ///
ESG_EVENTSCORE_WIND ESG_ESCORE_WIND ESG_SSCORE_WIND ESG_GSCORE_WIND ///
ESG_RATING_WIND2 ESG_SCORE_WIND2 ESG_MGMTSCORE_WIND2 ESG_EVENTSCORE_WIND2 ///
ESG_ESCORE_WIND2 ESG_SSCORE_WIND2 ESG_GSCORE_WIND2 ESG_GHG1_WIND ///
ESG_GHGR1_WIND ESG_GHG2_WIND ESG_GHGR2_WIND ESG_GHG12_WIND ESG_GHGR12_WIND ///
ESG_EM_WIND ESG_CR_WIND ESG_RATING_SSI ESG_RATING_FTSERUSSELL ESG_RATING

* ----------------------------------------------
* Step 2: Initialize empty list for variables to drop
* ----------------------------------------------
local esg_vars_to_drop ""

* ----------------------------------------------
* Step 3: Check each ESG variable for complete missingness
* ----------------------------------------------
foreach var of local esg_var_list {
    * Count missing values
    qui count if missing(`var')
    
    * Check if all values are missing
    if r(N) == _N {
        di "ESG Variable `var' is completely empty (100% missing)"
        local esg_vars_to_drop `esg_vars_to_drop' `var'
    }
    else {
        di "ESG Variable `var' has " _N - r(N) " non-missing values (keeping)"
    }
}

* ----------------------------------------------
* Step 4: Drop completely empty ESG variables if any exist
* ----------------------------------------------
if "`esg_vars_to_drop'" != "" {
    di "Dropping the following completely empty ESG variables: `esg_vars_to_drop'"
    drop `esg_vars_to_drop'
    
    * Display remaining ESG variables
    di _n "Remaining ESG variables:"
    *describe `esg_var_list', replace
    *list name in 1/`r(k)', clean noobs
}
else {
    di "No ESG variables are completely empty (none dropped)"
}

* ----------------------------------------------
* Step 5: Save summary of ESG variables
* ----------------------------------------------
*misstable summarize `esg_var_list'

* ==============================================
* END OF ESG VARIABLES CLEANING SCRIPT
* ==============================================
destring ESG_MGMTSCORE_WIND ESG_EVENTSCORE_WIND ESG_MGMTSCORE_WIND2 ESG_EVENTSCORE_WIND2, replace
drop ESG_RATING_WIND ESG_MGMTSCORE_WIND ESG_EVENTSCORE_WIND ESG_RATING_WIND2 ESG_MGMTSCORE_WIND2 ESG_EVENTSCORE_WIND2
save "esg.dta", replace
