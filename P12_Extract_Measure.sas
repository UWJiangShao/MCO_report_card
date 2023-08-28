OPTIONS PS=MAX FORMCHAR="|----|+|---+=|-/\<>*" MLOGIC MPRINT SYMBOLGEN;

* Prepare the dataframe for survey rating;

%LET JOB = P12;

LIBNAME OUT "..\DATA\Temp_Data\";

* macro to import all sheets from one Excel file;
%macro import_excel(excel_address, prog);

	libname ELIB excel &excel_address. mixed=yes;

	proc sql;
		create table excel_sheet as
		select *
		from dictionary.tables
		where libname = "ELIB"
		;

		select memname
		into :sheet_list separated by '*'
		from excel_sheet
		;

		select count(memname)
		into :sheet_n
		from excel_sheet
		;
	quit;

	%put &sheet_list.;
	%put &sheet_n.;

	%macro import_sheet;
		%do i = 1 %to &sheet_n.;
			%let var = %scan(&sheet_list., &i., *);
			%if %sysfunc(find(&var., $)) = 0 %then %do;
				%let sf = %substr(&var, 1, %length(&var.));
				data &prog._&sf.;
					set ELIB."&var."n;
					where PHI_Plan_Code ne ' ';
				run;

				proc contents data=&prog._&sf. varnum;
				run;

				proc sort data=&prog._&sf.;
					by PHI_Plan_Code;
				run;
			%end;
		%end;
	%mend import_sheet;

	%import_sheet;

%mend import_excel;


%macro import_single_sheet(excel_address, sheet_name, prog);
	* import population;
	proc import datafile = &excel_address.
			out=&prog._population
			dbms=excel
			replace;
		sheet = &sheet_name.;
		getnames = Yes;
	run;

	data &prog._population;
		set &prog._population;
		rename 
			PLAN_CD = PHI_Plan_Code
			sample_pool_members = Population
			;
	run;

	proc sort data=&prog._population;
		by PHI_Plan_Code;
	run;

	proc contents data=&prog._population varnum;
	run;
%mend import_single_sheet;

** For CHIP ----------------------------------------------------------------------------------------------;

%import_excel("..\DATA\Raw_Data\CHIP_2022_completes_out_plancode.xlsx", CH);
%import_single_sheet("..\DATA\Raw_Data\CHIP_22_samp_cnts_2203.xlsx", "2022_Counts", CH);


data CHIP_for_analysis;
	merge 
		CH_GCQ
		CH_HPRAT
		CH_HWDC
		CH_PDRAT
		CH_population(keep=PHI_Plan_Code Population)
		;
	by PHI_Plan_Code;
run;

proc export data=CHIP_for_analysis
	outfile="..\DATA\Temp_Data\CH23_out.xlsx"
	dbms=xlsx
	replace
	;
run;


* * For STAR Child ----------------------------------------------------------------------------------------;
* %import_excel("..\DATA\Raw_Data\STAR_Child_22_completes_out_ARC plancode.xlsx", SC);
* %import_single_sheet("..\DATA\Raw_Data\STAR_Child_22_samp_cnts_2203.xlsx", "STAR_Child_2022_Counts", SC);

* data STAR_Child_for_analysis;
* 	merge 
* 		SC_GCQ
* 		SC_HPRAT
* 		SC_HWDC
* 		SC_PDRAT
* 		SC_ROUTC
* 		SC_URGC
* 		SC_population(keep=PHI_Plan_Code Population)
* 		;
* 	by PHI_Plan_Code;
* run;

* proc export data=STAR_Child_for_analysis
* 	outfile="..\DATA\Temp_Data\SC23_out.xlsx"
* 	dbms=xlsx
* 	replace
* 	;
* run;


* * For STAR Adults ---------------------------------------------------------------------------------------;
* %import_excel("..\DATA\Raw_Data\STAR_Adult_2022_completes_out_ARC_plancode_v1.xlsx", SA);
* %import_single_sheet("..\DATA\Raw_Data\STAR_Adult_22_samp_cnts_2203.xlsx", "STAR_Adult_2022_Counts", SA);

* data STAR_Adult_for_analysis;
* 	merge 
* 		SA_ATC
* 		SA_GCQ
* 		SA_GNC
* 		SA_HPRAT
* 		SA_HWDC
* 		SA_PDRAT
* 		SA_population(keep=PHI_Plan_Code Population)
* 		;
* 	by PHI_Plan_Code;
* run;

* proc export data=STAR_Adult_for_analysis
* 	outfile="..\DATA\Temp_Data\SA23_out.xlsx"
* 	dbms=xlsx
* 	replace
* 	;
* run;


* * For STAR PLUS -----------------------------------------------------------------------------------------;
* %import_excel("..\DATA\Raw_Data\STAR_PLUS_2022_Completes_out_ARC.xlsx", SP);
* %import_single_sheet("..\DATA\Raw_Data\STAR_PLUS_22_samp_cnts_2204.xlsx", "Non-BCS", SP);


* DATA STARPLUS_for_analysis;
* 	merge 
* 		SP_ATC
* 		SP_GCQ
* 		SP_GNC
* 		SP_HPRAT
* 		SP_HWDC
* 		SP_PDRAT
* 		SP_population(keep=PHI_Plan_Code Population)
* 		;
* 	by PHI_Plan_Code;
* run;

* proc export data=STARPLUS_for_analysis
* 	outfile="..\DATA\Temp_Data\SP23_out.xlsx"
* 	dbms=xlsx
* 	replace
* 	;
* run;


* * For STAR Kids -----------------------------------------------------------------------------------------;
* %import_excel("..\DATA\Raw_Data\STAR_Kids_2022_merged_out_ARC.xlsx", SK);
* %import_single_sheet("..\DATA\Raw_Data\STAR_Kids_22_samp_cnts_2203.xlsx", "2022_Non-waiver", SK);

* DATA STARKIDS_for_analysis;
* 	MERGE 
* 		SK_ATC
* 		SK_GCQ
* 		SK_GNC
* 		SK_HPRAT
* 		SK_SPECTHER
* 		SK_COORD
* 		SK_BHCOUN
* 		SK_CCCGNI
* 		SK_CCCMEDS
* 		SK_TRTADULT
* 		SK_population(keep=PHI_Plan_Code Population)
* 		;
* 	by PHI_Plan_Code;
* run;

* data STARKIDS_for_analysis;
* 	set STARKIDS_for_analysis;
* 	rename 
* 		CCCGNI = GNI
* 		CCCGNI_den = GNI_den
* 		CCCGNI_stderr = GNI_stderr
* 		CCCGNI_sig = GNI_sig
* 		CCCMeds = APM
* 		CCCMeds_den = APM_den
* 		CCCMeds_stderr = APM_stderr
* 		CCCMeds_sig = APM_sig
* 		trtadult = transit
* 		trtadult_den = transit_den
* 		trtadult_stderr = transit_stderr
* 		trtadult_sig = transit_sig
* 		;
* run;

* proc export data=STARKIDS_for_analysis
* 	outfile="..\DATA\Temp_Data\SK23_out.xlsx"
* 	dbms=xlsx
* 	replace
* 	;
* run;
