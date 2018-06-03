ods listing close; 
ods pdf file= 'C:\Users\niwi8\OneDrive\Desktop\test.pdf' style=journal2 notoc; 
ods noproctitle; 
ods startpage = never; 
options nodate center; 

title "Final Project"; *creating an overall title called final project; 

*importing data; 

proc import out = info
	datafile = "C:\Users\niwi8\OneDrive\Documents\Spring 2018\SAS\final project\bmi.xlsx"
	dbms = xlsx replace; *importing the sheet info from excel file and naming it info; 
	sheet = "info"; *idicating sheet is called info; 
	getnames = yes; *using first row as variable names; 
run; 

proc import out = bmiwide
	datafile = "C:\Users\niwi8\OneDrive\Documents\Spring 2018\SAS\final project\bmi.xlsx"
	dbms = xlsx replace; *importing sheet bmi into sas and naming it bmiwide; 
	sheet = "BMI"; *indicating the sheet name; 
	getnames = yes; *using first row as variable names; 
run; 

proc import out = drink
	datafile = "C:\Users\niwi8\OneDrive\Documents\Spring 2018\SAS\final project\bmi.xlsx"
	dbms = xlsx replace; *importing she drink into sas and naming it drink; 
	sheet = "drink"; *indicating to pull sheet drink from file; 
	getnames = yes; *using first row as variable names; 
run; 

*transposing from wide to long; 

proc sort data = bmiwide; *sorting bmiwide; 
	by id; *sorting according to id; 
proc sort data = drink; *sorting drink; 
	by id; *sorting according to id; 
run; 

proc transpose data=bmiwide out=bmiLong
			(drop=_LABEL_ rename=(col1=bmi));*transposing bmiwide to long format and naming it bmilong, indicating to drop the label variable
				and to rename col1 to bmi; 
	by id; *transposing according to id; 
	var BMI0 BMI1 BMI2 BMI3 BMI4; *indicating variables that need to be transposed; 
run; 

data bmiLong; *data step on bmilong; 
	set bmiLong; *using data in bmilong; 
		drop _NAME_; *indicating to drop the variable _NAME_; 
	months = input(transtrn(_NAME_, "BMI", ""), 4.); *creating variable called months that takes the number after bmi for each observation and codes that as the month; 
run;

proc transpose data = drink out = drinkLong
			(drop = _LABEL_ rename = (col1 = drink)); *transposing drink to long format and naming it drinklong; 
	by id; *transposing by id; 
	var DRINK0 DRINK1 DRINK2 DRINK3 DRINK4; *indicating variables that need to be transposed; 
run; 

*merging sheets; 

proc sort data = info; *sorting the datasets info, bmilong, and drinklong by id; 
	by id; 
proc sort data = bmiLong; 
	by id; 
proc sort data = drinkLong; 
	by id;
run; 

data bmi; *creating dataset named bmi; 
	retain id months; *using retain to rearrange months to after id; 
	merge info bmiLong drinkLong; *merging datasets info, bmilong, and drinklong; 
		by id; *merging by id; 
	drop _NAME_; *dropping the variable name; 
	if months = 1 then months = 6; *if, else statements to change months to the actual month of follow-up; 
		else if months = 2 then months = 12; 
		else if months = 3 then months = 18; 
		else if months = 4 then months = 24; 
	if bmi = "NA" then bmi = .; *changing missing values in bmi from N/A to .; 
	if drink = "NA" then drink = .; *changing missing values in drink from NA to .; 
	num_bmi = input(bmi, 12.); *creating a variable called num_bmi that is a numeric version of bmi; 
		drop bmi; *dropping variable bmi; 
		rename num_bmi = bmi; *renaming num_bmi to bmi; 
	num_drink = input(drink, 12.); *creating a variable called num_drink that is a numeric version of drink; 
		drop drink; *dropping variable drink; 
		rename num_drink = drink; *renaming num_drink to drink; 
run; 

*creating formats and labels; 

proc format; *creating some formats; 
	value genfmt 0 = "Male"
				 1 = "Female"; *creating format for gender; 
	value difmt  0 = "No"
			     1 = "Yes"; *creating format for diabetes; 
	value hyperfmt 0 = "No"
				   1 = "Yes"; *creating format for hypertension; 
	value trtfmt 1 = "Control"
				 2 = "Laparoscopic"
				 3 = "Roux-en-Y"; *creating format for treatment; 
run; 

data bmi; *editing bmi; 
	set bmi; *using data in bmi; 
	label diabetes = "Diabetic status at baseline"
		  hypertension = "Hypertension status at baseline" 
		  TRT = "treatment"
		  drink = "Number of alcoholic drinks consumed per week" 
		  bmi = "BMI measurement"; * applying labels to variables; 
	format gender genfmt.
		   diabetes difmt.
		   hypertension hyperfmt. 
		   trt trtfmt.; *applying previously created formats to variables I created the formats for...;
run;

*creating extra datasets; 

data bmiwide; *editing bmiwide; 
	merge info bmiwide; *merging info and bmiwide; 
		by id; *merging by id; 
	format trt trtfmt.; *applying treatment format to treatment variable; 
	if bmi0 = "NA" then bmi0 = .; *changing na to .; 
		else if bmi1 = "NA" then bmi1 = .; *changing na to .;
		else if bmi2 = "NA" then bmi2 = .; *changing na to .;
		else if bmi3 = "NA" then bmi3 = .;*changing na to .;
		else if bmi4 = "NA" then bmi4 = .;*changing na to .;
	num_bmi0 = input(bmi0, 12.); *converting to numeric variable; 
		drop bmi0; *dropping bmi0; 
		rename num_bmi0 = bmi0; *renaming numberic to bmi0; 
	num_bmi1 = input(bmi1, 12.); *converting to numeric variable; 
		drop bmi1; *dropping bmi1; 
		rename num_bmi1 = bmi1;*renaming to bmi1; 
	num_bmi2 = input(bmi2, 12.); *converting to numeric variable; 
		drop bmi2; *dropping bmi2; 
		rename num_bmi2 = bmi2; *renaming to bmi2; 
	num_bmi3 = input(bmi3, 12.); *coverting to numeric; 
		drop bmi3; *dropping bmi3; 
		rename num_bmi3 = bmi3; *renaming to bmi3; 
	num_bmi4 = input(bmi4, 12.); *converting to numeric; 
		drop bmi4; *dropping bmi4; 
		rename num_bmi4 = bmi4; *renaming to bmi4; 
run; 

* exploratory analysis tables; 

proc print data = bmi (obs = 10) label; run; *printing first 10 observations of dataset bmi with labels; 

proc sort data = bmi; *sorting bmi; 
	by descending hypertension; *sorting by hypertension in descending order;  
run; 

title "Number of observations across treatment, diabetic status,"; *creating table title; 
title2 "and hypertensive status at baseline"; *creating table title on second line; 
proc tabulate data = bmi order = data
		missing; *tabulate procedure using the data in bmi and to order the data as it is in the dataset; 
	class trt gender diabetes hypertension; *indicating trt, gender, diabetes, and hypertension are categorical variables; 
	table (trt = "" ALL), (gender = "" ALL)*(diabetes = "Diabetes" hypertension = "Hypertension" ALL)*n = '' / 
		box = "" ; *creating a table with treatment in the rows with a total row, gender in the columns with diabetic status and hypertension status in the columns underneath gender; 
					* also requesting sas to print the counts for each cell; 
		where months = 0; *indicating to create table out of observations at baseline; 
	keylabel ALL = "Total"; *changing the label ALL to total; 
run;   

title "Mean and standard deviation of BMI across treatment, diabetic status, and hypertensive status"; *creating table title; 
title2 "at monthly follow-up"; *creating table title on second line; 
proc tabulate data = bmi order = data; *tabulate procedure using data in bmi and order that data as it appears; 
	class trt diabetes hypertension months; *indicating trt, diabetes, hypertension, and months are categorical variables; 
	var bmi; *indicating bmi is a continuous variable; 
	table (trt = "")*(months = ""), (diabetes = "Diabetes" hypertension = "Hypertension" ALL)*(bmi = "")*(mean std); *creating a table with trt and months in the rows with treatment higher 
		in the hierarchy than months, diabetes and hypertension are in the columns with bmi underneath them and printing the mean and std of bmi for each cell;  
	keylabel ALL = "Total"; *changing the label ALL to total; 
run; 

*investigating missing values; 

%macro missing(var); *creating a macro called missing that takes one variable: var; 

proc sort data = bmi; *sorting data in bmi; 
	by &var; *sorting by the variable &var; 
run; 

proc mi data = bmi nimpute = 0; *using a proc mi procedure on dataset bmi; 
	by &var; *indicating to run the procedure by the levels of the to be determined &var; 
	var trt gender diabetes hypertension bmi drink; *indicating the variables to look for missing values in ; 
run; 

*missing values only occur in bmi and drinks and they always occur together. The most missing values occured at the 6 month follow-up period 
and then decline.; 

%mend; 

%missing(months) *running missing on months; 
%missing(diabetes) *running missing on diabetes; 
%missing(hypertension) *running missing on hypertension; 
%missing(trt) *running missing on trt; 

*examing overrall change;

*sample sizes are sufficiently large;  

title "Control"; 
proc ttest data = bmiwide plots = none; *conducting ttest using data in bmiwide; 
	paired bmi0*bmi4; *conducting paired ttest between variables bmi0 and bmi4; 
	where trt = 1; *conducting the test for observations where trt is 1; 
run; 

title "Laparoscopic";
proc ttest data = bmiwide plots = none; *conducting ttest using data in bmiwide;
	paired bmi0*bmi4; *conducting paired ttest between variables bmi0 and bmi4;
	where trt = 2; *conducting the test for observations where trt is 2;
run; 

title "Roux"; 
proc ttest data = bmiwide plots = none; *conducting ttest using data in bmiwide;
	paired bmi0*bmi4; *conducting paired ttest between variables bmi0 and bmi4;
	where trt = 3; *conducting the test for observations where trt is 3;
run; 
	
*only examining last follow-up; 

title "Final project"; 
proc corr data = bmi plots = matrix; *correlation procedure using data in bmi and making matrix plot; 
	var bmi drink; *calculating correlatoin between bmi and drink; 
		where months = 24; *running procedure for observations where month = 24; 
run; 	

proc glm data = bmi; *conducting ANOVA using data in bmi to look if there is any difference at baseline; 
	class trt (ref = "Control"); *indicating trt is a categorical variable and to use control as the reference group; 
	model bmi = trt; *comparing bmi among the levels of trt; 
		where months = 0; *calculating differences at baseline;   
run; 

title "Final project"; *creating title; 
title2 "Modeling Last Observation all Covariates"; *creating second title; 
proc glm data = bmi
		plots = diagnostics; *glm procedure using data in bmi and asking for diagnostic plots; 
	class trt (ref = "Control"); *indicating trt is categorical and to use control as the reference; 
	class gender (ref = "Male"); *indicating gender is categorical and to use male as the reference; 
	class diabetes (ref = "No"); *indicating diabetes is categorical and to use no as the reference; 
	class hypertension (ref = "No"); *indicating hypertension is categorical and to use no as the reference; 
	model bmi = trt gender diabetes hypertension drink / noint solution; *creating linear model of bmi as a function of bmi, gender, diabetes, hypertension, 
			and drink and don't want an intercept and want solutions to effects; 
		where months = 24; *calculating for observations at month 24; 
		lsmeans trt / pdiff adjust = tukey; *indicating to calculate post hoc tests for pairwise differences among trt with a tukey-kramer adjustment; 
	output out = check r = resid p = pred; *creating a dataset called check that contains residuals and predicted values; 
	ods output ModelANOVA = type3_model24; *creating a dataset of ods ouput thats called type3_model24; 
run; quit; 

proc glm data = bmi
		plots = diagnostics; *glm procedure using data in bmi and asking for diagnostic plots; 
	class trt (ref = "Control"); *indicating trt is a categorical variable and that control is the reference group; 
	model bmi = trt / noint solution; *model bmi as a function of trt and asking for solutions to the effects; 
		where months = 24; *conducting analysis for observations at 24 months; 
		output out = check r = resid p = pred; *outputting residuals and predicted values to dataset called check; 
		lsmeans trt / pdiff adjust = tukey; *conducting pariwsie differences using tukey adjustment; 
run; quit; 

*At 24 months, only treatment is signficant and thus the final model only includes treatment. This makes sense as the study was an RCT and the treatment
groups should have been balanced on all other possible explanations. For the sake of covering all bases, I reported on the full models in the poster.; 

proc univariate data = check normal; *univariate procedure using data in check and to assess normality; 
	var resid; *running procedure on residuals; 
	histogram resid / normal kernel; *creating histogram of residuals overlayed with a normal plot and a kernel plot; 
	qqplot resid; *creating qqplot of residuals; 
run; 

			* kolmogorov-Smirnov test: p = 0.133 --> fail to reject the null, residuals are normally distributed; 

proc sgplot data = check; *plotting procedure using data in check; 
	scatter x = pred y = resid; *creating scatter plot of predicted values and residuals to assess homogeneity of variance; 
run; 

			* a spot check of the residuals versus predicted values shows equality of variance; 

%macro normal(var); *creating macro to assess normality of bmi for different variables at 24 months; 

proc univariate data = bmi normal; 
	class &var;  
	var bmi; 
		where months = 24; 
	histogram bmi / normal kernel; 
run; 

%mend; 

%normal(trt) *running normal on trt; 
%normal(gender) *running normal on gender; 
%normal(diabetes) *running normal on diabetes; 
%normal(hypertension) *running normal on hypertension; 
	
*random effects model; 

* data is sufficiently large; 

title "Random effects model"; *creating first title; 
title2 "All variables"; *creating a second title; 
proc mixed data = bmi covtest plots = residualpanel; *mixed effects procedure using data in bmi and want to conduct a covariance test and wnat plots of the resdiuals; 
	class trt (ref = "Control"); *indicating trt is a categorical variable and control is the reference; 
	class gender (ref = "Male"); *idicating gender is a categorical variable and male is the reference; 
	class diabetes (ref = "No"); *indicating diabetes is a categorical variable and no is the reference; 
	class hypertension (ref = "No"); *indicating hypertension is a categorical variable and no is the reference; 
	model bmi = trt | months gender diabetes hypertension drink / solution outpm = pred; *modeling bmi as a function of trt, months, and interaction between 
				trt and months, gender, diabetes, hypertensions, and drinks with predicted values stored in a dataset called pred; 
		random intercept / subject = id type = un; *indicating to treat the intercept as a random effect and that there are multiple observations for id; 
		lsmeans trt / at months = 0 adjust = tukey; *pairwise difference for trt at month 0 with tukey adjustment; 
		lsmeans trt / at months = 6 adjust = tukey; *pairwise difference for trt at month 6 with tukey adjustment;
		lsmeans trt / at months = 12 adjust = tukey; *pairwise difference for trt at month 12 with tukey adjustment;
		lsmeans trt / at months = 18 adjust = tukey; *pairwise difference for trt at month 18 with tukey adjustment;
		lsmeans trt / at months = 24 adjust = tukey; *pairwise difference for trt at month 24 with tukey adjustment;
	ods output  Tests3 = type3_allmonths; *outputting ods tests to a dataset called type3_allmonths; 
run; 

title "Random effects model"; *creating first title; 
title2 "only treatment, months, and drinks"; *creating a second title; 
proc mixed data = bmi covtest; *mixed effects procedure using data in bmi and want to conduct a covariance test and wnat plots of the resdiuals;
	class trt (ref = "Control"); *indicating trt is a categorical variable and control is the reference;
	model bmi = trt | months drink / solution outpm = pred2; *modeling bmi as a function of trt, months, and interaction between trt and months, and drinks; 
		random intercept / subject = id type = un; *indicating to treat the intercept as a random effect and that there are multiple observations for id; 
		lsmeans trt / at months = 0 adjust = tukey; *pairwise difference for trt at month 0 with tukey adjustment;
		lsmeans trt / at months = 6 adjust = tukey; *pairwise difference for trt at month 6 with tukey adjustment;
		lsmeans trt / at months = 12 adjust = tukey; *pairwise difference for trt at month 12 with tukey adjustment;
		lsmeans trt / at months = 18 adjust = tukey; *pairwise difference for trt at month 18 with tukey adjustment;
		lsmeans trt / at months = 24 adjust = tukey; *pairwise difference for trt at month 24 with tukey adjustment;
run; 

*only the interaciton between months and treatment, months, and drinks were significant. Because the interaction was significant the main effect
of treatmnet and months should not be interpreted. It appears alcohol consumption has a negative affect on weight loss. Both surgical treatments appear to 
have a similar effect on weight loss and are probably equivalent. However, while surgical treatments outperformed the non-surgical treatment over 24 months, 
the surgical treatment still resulted in weight loss and should probably be considerd the more cost effective option and maybe more appropriate treatment option 
for less severe cases of obesity; 

proc univariate data = pred2 normal; *univariate procedure using data in pred and to assess normality; 
	var resid; *examinging residuals; 
	histogram resid / normal kernel; *creating histogram of residauls; 
	qqplot resid; *creating qqplot of residuals; 
run; 

			* kolmogorov smirnov test: p = .091 --> fail to reject the null, residuals are normally distributed; 

proc sgplot data = pred2; *making scatter plot for residuals for model; 
	scatter x = pred y = resid; 
run; 

			* homogeneity of variance assumption appears fine on scatter plot; 

*making some tables and plots; 

title "Predicted BMI according to treatment regimine over 24 months"; *creating a title; 
title2 "random-effects model"; *creating another title.... this is so redundant; 
proc sgplot data=pred; *plotting procedure using data in pred; 
	scatter x = months y = pred / group = trt; *creating a scatter with months on the x axis and predicted values on the y, grouped by trt; 
	styleattrs datacontrastcolors=(viro sto bigb)
		datalinepatterns = (dash); *changing group colors and making lines dashed; 
	y2axis display = all; *indicating to plot a y2axis; 
	reg x = months y = pred / group = trt
		legendlabel = "Predicted Fit" name = "series" jitter transparency = 0.6; *plotting a line through scatter plot with a jitter and transparency of 0.6; 
	reg x = months y = pred / group = trt
		legendlabel = "Predicted Fit" name = "series" jitter y2axis; *plotting line again for the y2axis; 
		y2axis display = (nolabel); *not labeling the y2axis; 
	keylegend "series" / location = inside 
		position = bottomleft; *changing the position of the legend; 
	xaxis label="Months"; *changing xaxis label; 
	yaxis label="Predicted BMI"; *changing yaxis label; 
run;

data type3_model24; *editing type3_model24; 
	set type3_model24; *using type3_model24; 
	if HypothesisType = 1 then delete; *deleting type 1 hypothesis tests; 
	model = "Last Follow-up"; *creating a variable called model with all observations getting the label last follow-up; 
	rename source = effect; *renaming variable source to effect; 
run; 

data type3_allmonths; *editing type3_allmonths; 
	set type3_allmonths; *using type3_allmonths; 
	model = "Mixed Effects Model"; *creating variable called model and labeling all observations mixed effects model; 
run; 

data model_summaries; *creating dataset called model_summaries; 
	retain model effect; *moving model effect as the first variable; 
	set type3_model24 type3_allmonths; *combining type3_model24 and type3_allmonths; 
	keep effect Fvalue ProbF model; *only keeping these variables; 
run; 

proc format; 
	 value $effectfmt "TRT" = "Treatment"
				    "drink" = "Number of drinks"
					"months" = "Month of follow-up"
					"months*TRT" = "Month, treatment interaction"; *creating a format for the variable effect; 
run; 

title "Type III analysis for effects in models"; *creating a title; 
proc report data = model_summaries (rename=(ProbF = pbl)); *report procedure using data in model summaries and renaming ProbF to pbl; 
	column model effect Fvalue pbl; *indicating the variables to use are model effect Fvalue and pbl; 
		define model / group order = data 'Model'; *making model a row; 
		define effect / group center order = data'Effect'; *making effect a row; 
		define pb1 / analysis sum 'P-value'; *conducting an analysis on pbl that takes a sum which is just that value; 
		define Fvalue /analysis sum 'F-value'; *conducting an analysis on pbl that takes a sum which is just that value;
	format effect $effectfmt.; *formatting the varaible effect; 
run;

ods pdf close; 












 

