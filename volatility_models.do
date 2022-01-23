cls
clear

capture log close

set linesize 90
set scheme s2mono

graph drop _all

set graphics on

// alter the path to where the excel file is
local path = ""

if "`path'" == "" {
	di "Alter the path variable to the directory where the excel is"
	exit
}

cd "`path'"

// dates importing
import excel "dowjones.xlsx", firstrow allstring

// if save = 1 saves all results
local save = 1

// transform string dates to float dates
destring DJ, generate(prices)
drop DJ
destring M3TBILL, generate(m3tbill)
drop M3TBILL

// returns
gen ln_index = ln(prices)
gen returns = (ln_index - ln_index[_n-1]) * 100
drop ln_index
drop if returns == .

// time variable
gen dates = date(DATA, "MDY")
format %td dates
gen time = _n
tsset time

// Create 2-day period, 3-day period, 4-day period and weekly
gen days_of_week = dow(dates)

foreach i of numlist 2 3 4 {
	gen day_period_`i' = .
	replace day_period_`i' = 1 if mod(time, `i') == 0
} 

gen weekly = 1 if days_of_week == 1

gen daily = 1

drop days_of_week

// DJ prices over time
line prices dates, ylabel(0(2000)12000) yline(2000, lpattern(--)) ytitle("Prices") xtitle("Time") title("DJ prices over time") name(dj_prices_time)

// DJ returns over time
line returns dates, ylabel(-8(2)6) ytitle("DJ returns") xtitle("Time") title("DJ returns over time") name(dj_returns_time)

// Dj histogram of returns
histogram returns, normal ytitle("Density") xtitle("DJ returns") title("DJ histogram returns") name(dj_returns_histogram)

// DJ returns correlation
ac returns, ylabel(-0.06(0.02)0.06) xlabel(0(2)20) lags(20) ytitle("Correlation") title("DJ returns correlation") name(dj_returns_correlation)

// DJ returns squared correlation
gen returns_2 = returns ^ 2
ac returns_2, ylabel(0(0.05)0.25) xlabel(0(2)20) lags(20) ytitle("Correlation") title("DJ returns squared correlation") name(dj_returns_squared_correlation)

// Jarque-Bera test
quietly: tabstat returns, stat(n mean sd min max skewness kurtosis) save
matrix jb_matrix_and_statistics = r(StatTotal)'
quietly: sktest returns
local jb = (jb_matrix_and_statistics[1,1] / 6) * (jb_matrix_and_statistics[1,6] ^ 2 + 0.25 * (jb_matrix_and_statistics[1,7] - 3)^2)
local jb_pvalue = r(p_chi2)
matrix jb_test = (`jb',`jb_pvalue')
matrix jb_matrix_and_statistics = jb_matrix_and_statistics, jb_test
matrix colnames jb_matrix_and_statistics = "N" "Mean" "Standard-Deviation" "Min" "Max" "Skewness" "Kurtosis" "Jarque-Bera" "p-value"

// Optimal lag selection for models - AIC optimal lag is 1
varsoc returns time

// GARCH(1,1)
foreach freq in "daily" "day_period_2" "day_period_3" "day_period_4" "weekly" {
	
	preserve
	
	keep if `freq' == 1
	
	gen ln_index = ln(prices)
	replace returns = (ln_index - ln_index[_n - 1]) * 100
	drop ln_index
	drop if returns == .
	
	replace time = _n
	tsset time
	
	quietly: arch returns, arch(1) garch(1) save
	matrix garch_11 = r(table)
	matlist garch_11
	local b = garch_11[1, 1]
	local alfa = garch_11[1, 2]
	local beta = garch_11[1, 3]
	local omega = garch_11[1, 4]
	local b_se = garch_11[2, 1]
	local alfa_se = garch_11[2, 2]
	local beta_se = garch_11[2, 3]
	local omega_se = garch_11[2, 4]
	
	if "`freq'" == `"daily"' {
		// Tabela 6
		matrix garch_11_results = (`b', `b_se' \ `omega', `omega_se' \ `alfa', `alfa_se' \ `beta', `beta_se')
		matrix garch_11_results_periods = (`b' \ `omega' \ `alfa' \ `beta')
	}
	else {
		// Tabela 2
		matrix garch_11_results_periods = garch_11_results_periods, (`b' \ `omega' \ `alfa' \ `beta')
	}
	
	if "`freq'" == `"daily"' {
		predict myVariances_garch, variance
		gen desvio_garch = sqrt(myVariances_garch) * sqrt(252)
		
		line desvio_garch dates, ylabel(5(5)35) ytitle("GARCH(1,1) annual volatility") ttitle("time") name("annual_volatility_garch")
		matrix colnames garch_11_results = "Coeficients" "Standard-error"
		matrix rownames garch_11_results = "Constant" "omega" "alfa" "beta"
	}
	
	restore
}

matrix colnames garch_11_results_periods = "Daily" "2 day period" "3 day period" "4 day period" "Weekly"
matrix rownames garch_11_results_periods = "Constant" "omega" "alfa" "beta"

// TARCH(1,1,1)
quietly: arch returns, arch(1) garch(1) tarch(1) save
matrix tarch_111 = r(table)
local b = tarch_111[1, 1]
local gama = tarch_111[1, 3]
local alfa = tarch_111[1, 2]
local beta = tarch_111[1, 4]
local omega = tarch_111[1, 5]
local b_se = tarch_111[2, 1]
local gama_se = tarch_111[2, 3]
local alfa_se = tarch_111[2, 2]
local beta_se = tarch_111[2, 4]
local omega_se = tarch_111[2, 5]
matrix tarch_111_results = (`b', `b_se' \ `omega', `omega_se' \ `alfa', `alfa_se' \ `gama', `gama_se' \ `beta', `beta_se')
matrix colnames tarch_111_results = "Coeficients" "Standard-error"
matrix rownames tarch_111_results = "Constant" "omega" "alfa" "gama" "beta"
predict myVariances_tarch, variance
gen desvio_tarch_1 = sqrt(myVariances_tarch) * sqrt(252)

line desvio_tarch_1 dates, ylabel(5(5)35) ytitle("TARCH(1,1,1) annual volatility") ttitle("time") name("annual_volatility_tarch")

// GARCH(1,1) - X
quietly: arch returns, arch(1) garch(1) het(L.m3tbill)
matrix garch_11_X = r(table)
local b = garch_11_X[1, 1]
local phi = garch_11_X[1, 2]
local omega = garch_11_X[1, 3]
local alfa = garch_11_X[1, 4]
local beta = garch_11_X[1, 5]
local b_se = garch_11_X[2, 1]
local phi_se = garch_11_X[2, 2]
local omega_se = garch_11_X[2, 3]
local alfa_se = garch_11_X[2, 4]
local beta_se = garch_11_X[2, 5]
matrix garch_11_X_results = (`b', `b_se' \ `omega', `omega_se' \ `alfa', `alfa_se' \ `beta', `beta_se' \ `phi', `phi_se')
matrix colnames garch_11_X_results = "Coeficients" "Standard-error"
matrix rownames garch_11_X_results = "Constant" "omega" "alfa" "beta" "phi"
predict myVariances_garch_X, variance
gen desvio_garch_X = sqrt(myVariances_garch_X) * sqrt(252)

line desvio_garch_X dates, ylabel(5(5)35) ytitle("GARCH(1,1) - X annual volatility") ttitle("time") name("annual_volatility_garch_X")

if `save' == 1 {
	
	capture mkdir graphs
	capture mkdir excels
	
	graph export "graphs/dj_prices_time.png", replace name(dj_prices_time)
	graph export "graphs/dj_returns_time.png", replace name(dj_returns_time)
	graph export "graphs/dj_returns_histogram.png", replace name(dj_returns_histogram)
	graph export "graphs/dj_returns_correlation.png", replace name(dj_returns_correlation)
	graph export "graphs/dj_returns_squared_correlation.png", replace name(dj_returns_squared_correlation)
	graph export "graphs/annual_volatility_garch.png", replace name(annual_volatility_garch)
	graph export "graphs/annual_volatility_tarch.png", replace name(annual_volatility_tarch)
	graph export "graphs/annual_volatility_garch_X.png", replace name(annual_volatility_garch_X)

	putexcel set "excels/jb_matrix_and_statistics.xlsx", replace
	putexcel A1 = matrix(jb_matrix_and_statistics), names
	putexcel set "excels/garch_11_results.xlsx", replace
	putexcel A1 = matrix(garch_11_results), names
	putexcel set "excels/tarch_111_results.xlsx", replace
	putexcel A1 = matrix(tarch_111_results), names
	putexcel set "excels/garch_11_X_results.xlsx", replace
	putexcel A1 = matrix(garch_11_X_results), names
	putexcel set "excels/garch_11_results_periods", replace
	putexcel A1 = matrix(garch_11_results_periods), names
	
}

matlist jb_matrix_and_statistics
matlist garch_11_results
matlist garch_11_results_periods
matlist tarch_111_results
matlist garch_11_X_results
matlist garch_11_results_periods
