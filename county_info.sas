options nodate; 
ods noproctitle; 

libname county "C:\Users\niwi8\OneDrive\Documents\Practicum\opioid_prediction\data\county";

* AHRF cleaning (AHRF was imported as a SAS dataset using a SAS script supplied by mainteners of dataset); 

data ahrf; 
	set opioid.ahrf; 
	
	if f00008 = "Guam" then delete; 
	if f00008 = "Puerto Rico" then delete; 
	if f00008 = "US Virgin Islands" then delete; 
	if f00002 = "02001" then delete; 

	keep county f0874610  f1387610  f0885710  f0453010; 

	county = input(f00002, 8.); 

	rename f0874610 = percent_female_head
		   f1387610 = pop_density
		   f0885710 = total_md_2010
		   f0453010 = pop_2010;  
run;

data county.ahrf;
	retain county; 
	set ahrf; 
	phys = total_md_2010 / pop_2010 * 100000;
	dens_sq = pop_density * pop_density; 
run;  
 
* SAIPE - poverty rate and median income; 

proc import out = saipe
	datafile = "C:\Users\niwi8\OneDrive\Documents\Practicum\opioid_prediction\data\poverty_income_saipe.xls"
	dbms = xls replace; 
	getnames = yes; 
run;

data county.saipe; 
	retain county; 
	set saipe; 

	state_code = input(State_FIPS, 8.); 
	poverty_rate = input(Poverty_Percent_All_Ages, 8.); 
	income = input(Median_Household_Income, 8.); 
	county = (state_code * 1000) + County_FIPS;

	if County_Fips = 0 then delete; 
	if county = 0 then delete;  
	keep county poverty_rate income;  
run;  

* education shares; 

proc import out = education
	datafile = "C:\Users\niwi8\OneDrive\Documents\Practicum\opioid_prediction\data\education_shares.xls"
	dbms = xls replace; 
	getnames = yes; 
run; 

data education; 
	retain county; 
	set education;

	if FIPS_code = "11000" then delete; 
	if state = "PR" then delete; 
	if state = "US" then delete; 

	if Area_name in ("Alaska", "Alabama", "Arkansas", "Arizona", "California", "Colorado", "Connecticut", "Delaware",  
					 "Florida", "Georgia", "Hawaii", "Iowa", "Idaho", "Illinois", "Indiana", "Kansas", "Kentucky",  
					 "Lousiana", "Massachusetts", "Maryland", "Maine", "Mississippi", "Minnesota", "Missouri", "Michigan",  
					 "Montana", "North Carolina", "North Dakota", "Nebraska", "New Hampshire", "New Jersey", "New Mexico", 
					 "Nevada", "New York", "Ohio", "Oklahoma", "Oregon", "Pennsylvania", "Rhode Island", "South Carolina", 
					 "South Dakota", "Tennessee", "Texas", "Utah", "Virginia", "Vermont", "Washington", "Wisconsin", "West Virginia", 
				     "Wyoming")
		then delete;  

	county = input(FIPS_Code, 8.); 

	keep county Percent_of_adults_with_less_tha4 Percent_of_adults_with_a_high_s4 Percent_of_adults_completing_so4 Percent_of_adults_with_a_bachel2; 
	rename Percent_of_adults_with_less_tha4 = percent_hs_drop
		   Percent_of_adults_with_a_high_s4 = percent_hs_grad
		   Percent_of_adults_completing_so4 = percent_some_col
		   Percent_of_adults_with_a_bachel2 = percent_col_grad; 
run; 

data county.education; 
	set education; 
	if county = . then delete; 
run; 

* merging county level info; 

proc sort data = county.ahrf; 
	by county; 
run; 

proc sort data = county.saipe; 
	by county; 
run; 

proc sort data = county.education; 
	by county; 
run; 

data cleaned.county_info; 
	merge county.ahrf county.saipe county.education; 
	by county; 
run; 
