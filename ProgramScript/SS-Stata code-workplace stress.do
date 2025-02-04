/* Load the necessary package for reading .xls files */
ssc install importexcel, replace

/* Read the .xls file from the specified location. This project used data 
collected through administering a structured questionnaire having two modules. 
The first module includes questions regarding employees’ socio-demographic and 
occupational factors like sex, income range, job department and designation. 
As the second module an adapted version of Occupational Stress Index (OSI) 
developed and standardized by Srivastav and Singh, is used to evaluate workplace 
stress of the study population (Srivastava & Singh, 1984).*/

import excel "\dataset.xls", sheet("Sheet1") firstrow clear

/* Start logging the output to a log file */
log using RF_DA_v2.log, replace

/* Rename the specified variable */
rename Sex_cat sex_cat
rename religioncat rel_cat
rename maritalstatuscat ms_cat

/* Encode a string variable into numeric and generate a new variable */
encode Desig, gen(dg_en)
encode Dept, gen(dpt_en)
encode MaritalStatus, gen(ms_en)

/* List the values of the specified variable */
list dg_en
list dg_en, nolab

list dpt_en
list dpt_en, nolab

list ms_en
list ms_en, nolab

/* Sort the data by the specified variable */
sort dg_en, stable
sort dpt_en, stable

/* Generate row totals and means for specified variables */
egen isl = rowtotal(OSI1-OSI46) /* ro= Individual stress level */
egen isl_cat = cut(isl), at(46,96,159,230) icodes

/* Generate the 12 occupational stressor variables as the row total of specified variables */
egen ro = rowtotal( OSI1 OSI13 OSI22 OSI32 OSI37 OSI39 OSI40) /* ro= Role overload */
egen ra = rowtotal( OSI2 OSI23) /* ra= Role ambiguity */
egen rc = rowtotal( OSI3 OSI14 OSI33 OSI38) /* rc= Role conflict */
egen ugp = rowtotal( OSI4 OSI15 OSI24) /* ugp= Unreasonable group & political pressures */
egen rp = rowtotal( OSI5 OSI25) /* rp= Responsibility for persons  */
egen up = rowtotal( OSI6 OSI16 OSI26) /* up= Under participation */
egen pwn = rowtotal( OSI7 OSI27) /* pwn= Powerlessness */
egen ppr = rowtotal( OSI8 OSI17 OSI28 OSI34) /* ppr= Poor peer relations */
egen ini = rowtotal( OSI9 OSI18 OSI29 OSI35) /* ini= Intrinsic impoverishment */
egen ls = rowtotal( OSI10 OSI19 OSI30 OSI45) /* ls= Low status  */
egen swc = rowtotal( OSI12 OSI21 OSI31 OSI36 OSI41 OSI46) /* swc= Strenuous working conditions */
egen upb = rowtotal( OSI11 OSI20 OSI43 OSI44) /* upb= Unprofitability */

/* Generate new variables as the row mean of specified variables */
egen ro_mn = rowmean( OSI1 OSI13 OSI22 OSI32 OSI37 OSI39 OSI40)
egen ra_mn = rowmean( OSI2 OSI23)
egen rc_mn = rowmean( OSI3 OSI14 OSI33 OSI38)
egen ugp_mn = rowmean( OSI4 OSI15 OSI24)
egen rp_mn = rowmean( OSI5 OSI25)
egen up_mn = rowmean( OSI6 OSI16 OSI26)
egen pwn_mn = rowmean( OSI7 OSI27)
egen ppr_mn = rowmean( OSI8 OSI17 OSI28 OSI34)
egen ini_mn = rowmean( OSI9 OSI18 OSI29 OSI35)
egen ls_mn = rowmean( OSI10 OSI19 OSI30 OSI45)
egen swc_mn = rowmean( OSI12 OSI21 OSI31 OSI36 OSI41 OSI46)
egen upb_mn = rowmean( OSI11 OSI20 OSI43 OSI44)

/* Display summary statistics for the specified variables */
summarize ro_mn ra_mn rc_mn ugp_mn rp_mn up_mn pwn_mn ppr_mn ini_mn ls_mn swc_mn upb_mn

/* Calculate pairwise correlations between the specified variables */
pwcorr ro_mn ra_mn rc_mn ugp_mn rp_mn up_mn pwn_mn ppr_mn ini_mn ls_mn swc_mn upb_mn , sig star(.05)

/* Create a table of frequencies for the specified variables */
tab ro_mn isl_cat , chi2

*\ Create a binary variable for each of the three categories of isl_cat */
gen isl_low = 1 if isl_cat == 0
replace isl_low = 0 if isl_cat != 0

gen isl_mod = 1 if isl_cat == 1
replace isl_mod = 0 if isl_cat != 1

gen isl_hg = 1 if isl_cat == 2
replace isl_hg = 0 if isl_cat != 2

*\ Create a binary variable for each of the three categories of dg_cat */
gen dg_up = 1 if dg_cat == 0
replace dg_up = 0 if dg_cat != 0

gen dg_mid = 1 if dg_cat == 1
replace dg_mid = 0 if dg_cat != 1

gen dg_low = 1 if dg_cat == 2
replace dg_low = 0 if dg_cat != 2

/* Make the 12 occupational stressors dichotomous */
foreach var in ro ra rc ugp rp up pwn ppr ini ls upb {
    gen `var'_yn = .  // Initialize the new variable
    replace `var'_yn = 1 if `var'_mn > 3
    replace `var'_yn = 0 if `var'_mn <= 3


/* Create a histogram for the specified variable */
histogram isl

/* Define value labels for employee's designation */
label define dg_en_new 1 "General Manager" 2 "Assist. General Manager" 3 "Senior Manager" 4 "Manager" 5 "Deputy Manager" 6 "Assit. Manager" 7 "Sr. Group Manager" 8 "Brand/Group Manager" 9 "Sr. Executive" 10 "Executive"

/* Create a bar graph for the specified variables */
graph bar (mean) ro_mn ra_mn rc_mn ugp_mn rp_mn up_mn pwn_mn ppr_mn ini_mn ls_mn swc_mn upb_mn

/* Perform stepwise regression with specified variables */
xi: stepwise, pr(.1): regress isl (i.sex_en) (i.dpt_en)

/* Generate a new variable based on a condition */
gen dg_new=1 if Desig=="General Manager"
replace dg_new=2 if Desig=="Assist. General Manager"
replace dg_new=3 if Desig=="Senior Manager"
replace dg_new=4 if Desig=="Manager"
replace dg_new=5 if Desig=="Deputy Manager"
replace dg_new=6 if Desig=="Assit. Manager"
replace dg_new=7 if Desig=="Sr. Group Manager"
replace dg_new=8 if Desig=="Brand/Group Manager"
replace dg_new=9 if Desig=="Sr. Executive"
replace dg_new=10 if Desig=="Executive"

/* Generate a new categorical variable by cutting a continuous variable */
egen dg_cat = cut(dg_new), at(1,4,8,11) icodes

/* Drop the specified variables */
drop isl_cat isl_low isl_mod isl_hg ra rc ra_mn rc_mn ra_yn rc_yn

/* Generate a new categorical variable by cutting individual stress level continuous variable */
egen isl_cat = cut(isl), at(46,96,159,230) icodes

/* Separate the categorical variable into binary indicators */
gen isl_low = 1 if isl_cat == 0
replace isl_low = 0 if isl_cat != 0
gen isl_mod = 1 if isl_cat == 1
replace isl_mod = 0 if isl_cat != 1
gen isl_hg = 1 if isl_cat == 2
replace isl_hg = 0 if isl_cat != 2

/* Generate new variables as the row total of specified variables */
egen ra = rowtotal( OSI2 OSI23)
egen rc = rowtotal( OSI3 OSI14 OSI33 OSI38)

/* Generate new variables as the row mean of specified variables */
egen ra_mn = rowmean( OSI2 OSI23)
egen rc_mn = rowmean( OSI3 OSI14 OSI33 OSI38)

/* Make stressors dichotomous */
gen ra_yn = 1 if ra_mn > 3
replace ra_yn = 0 if ra_mn <= 3
gen rc_yn = 1 if rc_mn > 3
replace rc_yn = 0 if rc_mn <= 3

/* Perform one-way ANOVA for the specified variables */
oneway isl dpt_en, tabulate


/* Calculate pairwise means for the specified variables with multiple comparisons */
pwmean isl, over(dpt_en) mcompare(tukey) effects

/* Perform ANOVA for the specified variables with interaction terms */
anova isl dpt_en##dg_cat
anova isl dg_cat##dpt_en##sex_en

/* Perform t-test for the specified variables by groups */
ttest ugp_mn, by(sex_en)

/* Generate a new variable as the standard deviation of specified variables */
egen vra=sd(ra)

/* Calculate Cronbach's alpha for the specified variables */
alpha OSI1-OSI46

/* Perform Shapiro-Wilk test for normality on the specified variable */
swilk isl

/* Display summary statistics for the specified variable */
summarize Age

/* Create a table of frequencies for the specified variables */
tab Age_cat sex_en rel_en ms_en edu_en js_en ji_en dpt_en dg_cat TWE_cat WECC_cat aso_en

/* Display summary statistics for the specified variable */
summarize isl

/* Create a box plot for the specified variable */
graph box isl

/* Perform regression analysis with specified variables */
regress isl ib2.dpt_en i.dg_cat i.sex_en i.ms_en i.TWE_cat i.WECC_cat i.workhour_cat i.edu_en
regress zisl Age ib2.dpt_en i.dg_cat i.sex_en ib2.ms_en i.TWE_cat i.WECC_cat i.workhour_cat i.edu_en

foreach var in ro ra rc ugp rp up pwn ppr ini ls swc upb {
    regress `var'_mn ib2.dpt_en
    regress `var'_mn i.workhour_cat
    regress `var'_mn i.WECC_cat
    regress `var'_mn i.sex_en
    regress `var'_mn i.dg_cat
}
