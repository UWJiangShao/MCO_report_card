OPTIONS PS=MAX FORMCHAR="|----|+|---+=|-/\<>*" MLOGIC MPRINT SYMBOLGEN noxwait noxsync;

* Formatting results;
* Using DDE technique; 

%LET JOB = P31;


*** ---- import plancode description ---------------------------------------------------------------------;
proc import datafile="..\DATA\Raw_Data\plancode.xlsx"
	dbms=XLSX
	out=plancode
	;
run;

*** --- import quoatas and weighting ---------------------------------------------------------------------;
%macro import_population(excel_address, sheet_name, prog);
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
		* base_weight = eligible_members / sample_members;
	run;

	proc sort data=&prog._population;
		by PHI_SA_Name PHI_Plan_Name;
	run;

	proc contents data=&prog._population varnum;
	run;
%mend import_population;

%import_population("..\DATA\Raw_Data\CHIP_22_samp_cnts_2203.xlsx", "2022_Counts", CH);
%import_population("..\DATA\Raw_Data\STAR_Child_22_samp_cnts_2203.xlsx", "STAR_Child_2022_Counts", SC);
%import_population("..\DATA\Raw_Data\STAR_Adult_22_samp_cnts_2203.xlsx", "STAR_Adult_2022_Counts", SA);
%import_population("..\DATA\Raw_Data\STAR_PLUS_22_samp_cnts_2204.xlsx", "Non-BCS", SP);
%import_population("..\DATA\Raw_Data\STAR_Kids_22_samp_cnts_2203.xlsx", "2022_Non-waiver", SK);



*** ---- CHIP prepare data ------------------------------------------------------------------------------------------------------;
proc import datafile="..\DATA\Temp_Data\CH23_out.xlsx"
	dbms=XLSX
	out=CH23
	;
	sheet="CH23";
run;

proc contents data=CH23 varnum;
run;

proc sql;
	create table CHIP as
	select *
		,avg(GCQ) as mean_GCQ
		,avg(HWDC) as mean_HWDC
		,avg(PDRat) as mean_PDRat
		,avg(HPRat) as mean_HPRat
		from CH23
		;
quit;

data CHIP;
	set CHIP;
	rename PHI_Plan_Code = plancode;

	dffm_GCQ = GCQ - mean_GCQ;
	dffm_HWDC = HWDC - mean_HWDC;
	dffm_PDRat = PDRat - mean_PDRat;
	dffm_HPRat = HPRat - mean_HPRat;
run;

proc freq data=CHIP;
	table plancode;
run;

proc sort data=plancode; by plancode; run;
proc sort data=CHIP; by plancode; run;

data CHIP;
	merge CHIP(in=a) plancode(keep=MCONAME PLANCODE SERVICEAREA);
	by plancode;
	if a;
run;

proc sort data=CHIP; by SERVICEAREA MCONAME; run;


*** ------ STAR Child prepare data ----------------------------------------------------------------------------------------------------------;
proc import datafile="..\DATA\Temp_Data\SC23_out.xlsx"
	dbms=XLSX
	out=SC23
	;
	sheet="SC23";
run;

proc contents data=SC23 varnum;
run;

proc sql;
	create table STAR_Child as
	select *
		,avg(GCQ) as mean_GCQ
		,avg(HWDC) as mean_HWDC
		,avg(PDRat) as mean_PDRat
		,avg(HPRat) as mean_HPRat
		from SC23
		;
quit;

data STAR_Child;
	set STAR_Child;
	rename PHI_Plan_Code = plancode;

	dffm_GCQ = GCQ - mean_GCQ;
	dffm_HWDC = HWDC - mean_HWDC;
	dffm_PDRat = PDRat - mean_PDRat;
	dffm_HPRat = HPRat - mean_HPRat;
run;

proc sort data=STAR_Child; by plancode; run;

data STAR_Child;
	merge STAR_Child(in=a) plancode(keep=MCONAME PLANCODE SERVICEAREA);
	by plancode;
	if a;
run;

proc sort data=STAR_Child; by SERVICEAREA MCONAME; run;

proc print data=STAR_Child noobs;
	var plancode GCQ mean_GCQ dffm_GCQ;
run;


*** ------ STAR Adult prepare data ----------------------------------------------------------------------------------------------------------;
proc import datafile="..\DATA\Temp_Data\SA23_out.xlsx"
	dbms=XLSX
	out=SA23
	;
	sheet="SA23";
run;

proc contents data=SA23 varnum;
run;

proc sql;
	create table STAR_Adult as
	select *
		,avg(AtC) as mean_AtC
		,avg(HWDC) as mean_HWDC
		,avg(PDRat) as mean_PDRat
		,avg(HPRat) as mean_HPRat
		from SA23
		;
quit;

data STAR_Adult;
	set STAR_Adult;
	rename PHI_Plan_Code = plancode;

	dffm_AtC = AtC - mean_AtC;
	dffm_HWDC = HWDC - mean_HWDC;
	dffm_PDRat = PDRat - mean_PDRat;
	dffm_HPRat = HPRat - mean_HPRat;
run;

proc sort data=STAR_Adult; by plancode; run;

data STAR_Adult;
	merge STAR_Adult(in=a) plancode(keep=MCONAME PLANCODE SERVICEAREA);
	by plancode;
	if a;
run;

proc sort data=STAR_Adult; by SERVICEAREA MCONAME; run;


*** ------ STAR + PLUS prepare data ----------------------------------------------------------------------------------------------------------;
proc import datafile="..\DATA\Temp_Data\SP23_out.xlsx"
	dbms=XLSX
	out=SP23
	;
	sheet="SP23";
run;

proc contents data=SP23 varnum;
run;

proc sql;
	create table STARPLUS as
	select *
		,avg(AtC) as mean_AtC
		,avg(HWDC) as mean_HWDC
		,avg(PDRat) as mean_PDRat
		,avg(HPRat) as mean_HPRat
		from SP23
		;
quit;

data STARPLUS;
	set STARPLUS;
	rename PHI_Plan_Code = plancode;

	dffm_AtC = AtC - mean_AtC;
	dffm_HWDC = HWDC - mean_HWDC;
	dffm_PDRat = PDRat - mean_PDRat;
	dffm_HPRat = HPRat - mean_HPRat;
run;

proc sort data=STARPLUS; by plancode; run;

data STARPLUS;
	merge STARPLUS(in=a) plancode(keep=MCONAME PLANCODE SERVICEAREA);
	by plancode;
	if a;
run;

proc sort data=STARPLUS; by SERVICEAREA MCONAME; run;


*** ------ STAR Kids prepare data ----------------------------------------------------------------------------------------------------------;
proc import datafile="..\DATA\Temp_Data\SK23_out.xlsx"
	dbms=XLSX
	out=SK23
	;
	sheet="SK23";
run;

proc contents data=SK23 varnum;
run;

proc sql;
	create table STARKids as
	select *
		,avg(HPRat) as mean_HPRat
		,avg(AtC) as mean_AtC
		,avg(SpecTher) as mean_SpecTher
		,avg(APM) as mean_APM
		,avg(coord) as mean_coord
		,avg(GNI) as mean_GNI
		,avg(transit) as mean_transit
		,avg(BHCoun) as mean_BHCoun
		from SK23
		;
quit;

data STARKids;
	set STARKids;
	rename PHI_Plan_Code = plancode;

	dffm_HPRat = HPRat - mean_HPRat;
	dffm_AtC = AtC - mean_AtC;
	dffm_SpecTher = SpecTher - mean_SpecTher;
	dffm_APM = APM - mean_APM;
	dffm_coord = coord - mean_coord;
	dffm_GNI = GNI - mean_GNI;
	dffm_transit = transit - mean_transit;
	dffm_BHCoun = BHCoun - mean_BHCoun;
run;

proc sort data=STARKids; by plancode; run;

data STARKids;
	merge STARKids(in=a) plancode(keep=MCONAME PLANCODE SERVICEAREA);
	by plancode;
	if a;
run;

proc sort data=STARKids; by SERVICEAREA MCONAME; run;


** ---- Create frequency table for the rating guide ---------------------------------------------;
proc freq data=CHIP nlevels;
	table HPRat_rat /list out=RG_CH_HPRat;
	table GCQ_rat /list out=RG_CH_GCQ;
	table HWDC_rat /list out=RG_CH_HWDC;
	table PDRat_rat /list out=RG_CH_PDRat;
run;

proc freq data=STAR_Child nlevels;
	table HPRat_rat /list out=RG_SC_HPRat;
	table GCQ_rat /list out=RG_SC_GCQ;
	table HWDC_rat /list out=RG_SC_HWDC;
	table PDRat_rat /list out=RG_SC_PDRat;
run;

proc freq data=STAR_Adult nlevels;
	table HPRat_rat /list out=RG_SA_HPRat;
	table AtC_rat /list out=RG_SA_AtC;
	table HWDC_rat /list out=RG_SA_HWDC;
	table PDRat_rat /list out=RG_SA_PDRat;
run;

proc freq data=STARPLUS nlevels;
	table HPRat_rat /list out=RG_SP_HPRat;
	table AtC_rat /list out=RG_SP_AtC;
	table HWDC_rat /list out=RG_SP_HWDC;
	table PDRat_rat /list out=RG_SP_PDRat;
run;

proc freq data=STARKids nlevels;
	table HPRat_rat /list out=RG_SK_HPRat;
	table AtC_rat /list out=RG_SK_AtC;
	table SpecTher_rat /list out=RG_SK_SpecTher;
	table APM_rat /list out=RG_SK_APM;
	table coord_rat /list out=RG_SK_coord;
	table GNI_rat /list out=RG_SK_GNI;
	table transit_rat /list out=RG_SK_transit;
	table BHCoun_rat /list out=RG_SK_BHCoun;
run;


** ---- Exporting using DDE --------------------------------------------------------------------;
filename ddeopen DDE 'Excel|system';

* template file;
x '"C:\Users\liqian\Dropbox (UFL)\Project_QL\2022_MCO_Report_Cards\SFY2022-2023\3. Survey\DATA\Raw_Data\MCO Report Cards Survey Ratings Template.xlsx"';

** ---- fill quotas and weighting ----------------------------------------------------------------------------;
filename CHpop dde "Excel|Quotas and Weighting!r4c10:r35c15" notab;
data _null_;
	set CH_population;
	file CHpop;
	put PHI_Plan_Name '09'x PHI_SA_Name '09'x PLAN_CD '09'x 
		sample_pool_members;
run;

filename SCpop dde "Excel|Quotas and Weighting!r4c18:r47c23" notab;
data _null_;
	set SC_population;
	file SCpop;
	put PHI_Plan_Name '09'x PHI_SA_Name '09'x PLAN_CD '09'x 
		sample_pool_members;
run;

filename SApop dde "Excel|Quotas and Weighting!r4c26:r47c31" notab;
data _null_;
	set SA_population;
	file SApop;
	put PHI_Plan_Name '09'x PHI_SA_Name '09'x PLAN_CD '09'x 
		sample_pool_members;
run;

filename SPpop dde "Excel|Quotas and Weighting!r4c34:r30c39" notab;
data _null_;
	set SP_population;
	file SPpop;
	put PHI_Plan_Name '09'x PHI_SA_Name '09'x PLAN_CD '09'x 
		sample_pool_members;
run;

filename SKpop dde "Excel|Quotas and Weighting!r4c42:r31c47" notab;
data _null_;
	set SK_population;
	file SKpop;
	put PHI_Plan_Name '09'x PHI_SA_Name '09'x PLAN_CD '09'x 
		sample_pool_members;
run;

** ---- fill ratings ----------------------------------------------------------------------------;
filename CHGCQ dde "Excel|CHIP-Getting Care Quickly!r3c1:r34c11" notab;
filename CHHWDC dde "Excel|CHIP-HWDC!r3c1:r34c11" notab;
filename CHPDRat dde "Excel|CHIP-Rate Personal Doctor!r3c1:r34c11" notab;
filename CHHPRat dde "Excel|CHIP-Rate Health Plan!r3c1:r34c11" notab;

data _null_;
	set CHIP;

	file CHGCQ;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		GCQ_den '09'x GCQ '09'x mean_GCQ '09'x dffm_GCQ '09'x
		GCQ_stderr '09'x GCQ_relb '09'x GCQ_sig '09'x GCQ_rat
		;

	file CHHWDC;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		HWDC_den '09'x HWDC '09'x mean_HWDC '09'x dffm_HWDC '09'x
		HWDC_stderr '09'x HWDC_relb '09'x HWDC_sig '09'x HWDC_rat
		;

	file CHPDRat;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		PDRat_den '09'x PDRat '09'x mean_PDRat '09'x dffm_PDRat '09'x
		PDRat_stderr '09'x PDRat_relb '09'x PDRat_sig '09'x PDRat_rat
		;

	file CHHPRat;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		HPRat_den '09'x HPRat '09'x mean_HPRat '09'x dffm_HPRat '09'x
		HPRat_stderr '09'x HPRat_relb '09'x HPRat_sig '09'x HPRat_rat
		;
run;


filename SCHPRat dde "Excel|STAR Child-Rate Health Plan!r3c1:r46c11" notab;
filename SCGCQ dde "Excel|STAR Child-Getting Care Quickly!r3c1:r46c11" notab;
filename SCHWDC dde "Excel|STAR Child-HWDC!r3c1:r46c11" notab;
filename SCPDRat dde "Excel|STAR Child-Rate Personal Doctor!r3c1:r46c11" notab;

data _null_;
	set STAR_Child;

	file SCHPRat;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		HPRat_den '09'x HPRat '09'x mean_HPRat '09'x dffm_HPRat '09'x
		HPRat_stderr '09'x HPRat_relb '09'x HPRat_sig '09'x HPRat_rat
		;

	file SCGCQ;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		GCQ_den '09'x GCQ '09'x mean_GCQ '09'x dffm_GCQ '09'x
		GCQ_stderr '09'x GCQ_relb '09'x GCQ_sig '09'x GCQ_rat
		;

	file SCHWDC;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		HWDC_den '09'x HWDC '09'x mean_HWDC '09'x dffm_HWDC '09'x
		HWDC_stderr '09'x HWDC_relb '09'x HWDC_sig '09'x HWDC_rat
		;

	file SCPDRat;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		PDRat_den '09'x PDRat '09'x mean_PDRat '09'x dffm_PDRat '09'x
		PDRat_stderr '09'x PDRat_relb '09'x PDRat_sig '09'x PDRat_rat
		;
run;

filename SAHPRat dde "Excel|STAR Adult-Rate Health Plan!r3c1:r46c11" notab;
filename SAAtC dde "Excel|STAR Adult-Getting Treat!r3c1:r46c13" notab;
filename SAHWDC dde "Excel|STAR Adult-HWDC!r3c1:r46c11" notab;
filename SAPDRat dde "Excel|STAR Adult-Rate Personal Doctor!r3c1:r46c11" notab;

data _null_;
	set STAR_Adult;

	file SAHPRat;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		HPRat_den '09'x HPRat '09'x mean_HPRat '09'x dffm_HPRat '09'x
		HPRat_stderr '09'x HPRat_relb '09'x HPRat_sig '09'x HPRat_rat
		;

	file SAAtC;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		GCQ '09'x GNC '09'x
		AtC_den '09'x AtC '09'x mean_AtC '09'x dffm_AtC '09'x
		AtC_stderr '09'x AtC_relb '09'x AtC_sig '09'x AtC_rat
		;

	file SAHWDC;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		HWDC_den '09'x HWDC '09'x mean_HWDC '09'x dffm_HWDC '09'x
		HWDC_stderr '09'x HWDC_relb '09'x HWDC_sig '09'x HWDC_rat
		;

	file SAPDRat;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		PDRat_den '09'x PDRat '09'x mean_PDRat '09'x dffm_PDRat '09'x
		PDRat_stderr '09'x PDRat_relb '09'x PDRat_sig '09'x PDRat_rat
		;
run;


filename SPHPRat dde "Excel|STAR+PLUS-Rate Health Plan!r3c1:r29c11" notab;
filename SPAtC dde "Excel|STAR+PLUS-Getting Treat!r3c1:r29c13" notab;
filename SPHWDC dde "Excel|STAR+PLUS-HWDC!r3c1:r29c11" notab;
filename SPPDRat dde "Excel|STAR+PLUS-Rate Personal Doctor!r3c1:r29c11" notab;

data _null_;
	set STARPLUS;

	file SPHPRat;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		HPRat_den '09'x HPRat '09'x mean_HPRat '09'x dffm_HPRat '09'x
		HPRat_stderr '09'x HPRat_relb '09'x HPRat_sig '09'x HPRat_rat
		;

	file SPAtC;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		GCQ '09'x GNC '09'x
		AtC_den '09'x AtC '09'x mean_AtC '09'x dffm_AtC '09'x
		AtC_stderr '09'x AtC_relb '09'x AtC_sig '09'x AtC_rat
		;

	file SPHWDC;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		HWDC_den '09'x HWDC '09'x mean_HWDC '09'x dffm_HWDC '09'x
		HWDC_stderr '09'x HWDC_relb '09'x HWDC_sig '09'x HWDC_rat
		;

	file SPPDRat;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		PDRat_den '09'x PDRat '09'x mean_PDRat '09'x dffm_PDRat '09'x
		PDRat_stderr '09'x PDRat_relb '09'x PDRat_sig '09'x PDRat_rat
		;
run;

filename SKHPRat dde "Excel|STAR Kids-Rate Health Plan!r3c1:r30c11" notab;
filename SKAtC dde "Excel|STAR Kids-Getting Treat!r3c1:r30c13" notab;
filename SKSpec dde "Excel|STAR Kids-Special Therapy!r3c1:r30c11" notab;
filename SKAPM dde "Excel|STAR Kids-Prescriptions!r3c1:r30c11" notab;
filename SKcoord dde "Excel|STAR Kids-Care Coordination!r3c1:r30c11" notab;
filename SKGNI dde "Excel|STAR Kids-Getting Information!r3c1:r30c11" notab;
filename SKtrans dde "Excel|STAR Kids-Transition!r3c1:r30c11" notab;
filename SKBHCoun dde "Excel|STAR Kids-Counseling!r3c1:r30c11" notab;

data _null_;
	set STARKids;

	file SKHPRat;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		HPRat_den '09'x HPRat '09'x mean_HPRat '09'x dffm_HPRat '09'x
		HPRat_stderr '09'x HPRat_relb '09'x HPRat_sig '09'x HPRat_rat
		;

	file SKAtC;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		GCQ '09'x GNC '09'x
		AtC_den '09'x AtC '09'x mean_AtC '09'x dffm_AtC '09'x
		AtC_stderr '09'x AtC_relb '09'x AtC_sig '09'x AtC_rat
		;

	file SKSpec;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		SpecTher_den '09'x SpecTher '09'x mean_SpecTher '09'x dffm_SpecTher '09'x
		SpecTher_stderr '09'x SpecTher_relb '09'x SpecTher_sig '09'x SpecTher_rat
		;

	file SKAPM;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		APM_den '09'x APM '09'x mean_APM '09'x dffm_APM '09'x
		APM_stderr '09'x APM_relb '09'x APM_sig '09'x APM_rat
		;

	file SKcoord;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		coord_den '09'x coord '09'x mean_coord '09'x dffm_coord '09'x
		coord_stderr '09'x coord_relb '09'x coord_sig '09'x coord_rat
		;

	file SKGNI;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		GNI_den '09'x GNI '09'x mean_GNI '09'x dffm_GNI '09'x
		GNI_stderr '09'x GNI_relb '09'x GNI_sig '09'x GNI_rat
		;

	file SKtrans;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		transit_den '09'x transit '09'x mean_transit '09'x dffm_transit '09'x
		transit_stderr '09'x transit_relb '09'x transit_sig '09'x transit_rat
		;

	file SKBHCoun;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		BHCoun_den '09'x BHCoun '09'x mean_BHCoun '09'x dffm_BHCoun '09'x
		BHCoun_stderr '09'x BHCoun_relb '09'x BHCoun_sig '09'x BHCoun_rat
		;
run;


** ---- fill rating guide ----------------------------------------------------------------------------;

%macro fill_rating_guide(freqdata, ratevar, address);
	filename RateGu dde "Excel|&address." notab;
	data _null_;
		set &freqdata.(where=(&ratevar. ne .));
		file RateGu;
		put &ratevar. '09'x count
			;
	run;
%mend fill_rating_guide;

%fill_rating_guide(RG_CH_HPRat, HPRat_rat, CHIP-Rate Health Plan!r4c15:r8C16);
%fill_rating_guide(RG_CH_GCQ, GCQ_rat, CHIP-Getting Care Quickly!r4c15:r8C16);
%fill_rating_guide(RG_CH_HWDC, HWDC_rat, CHIP-HWDC!r4c15:r8C16);
%fill_rating_guide(RG_CH_PDRat, PDRat_rat, CHIP-Rate Personal Doctor!r4c15:r8C16);

%fill_rating_guide(RG_SC_HPRat, HPRat_rat, STAR Child-Rate Health Plan!r4c15:r8C16);
%fill_rating_guide(RG_SC_GCQ, GCQ_rat, STAR Child-Getting Care Quickly!r4c15:r8C16);
%fill_rating_guide(RG_SC_HWDC, HWDC_rat, STAR Child-HWDC!r4c15:r8C16);
%fill_rating_guide(RG_SC_PDRat, PDRat_rat, STAR Child-Rate Personal Doctor!r4c15:r8C16);

%fill_rating_guide(RG_SA_HPRat, HPRat_rat, STAR Adult-Rate Health Plan!r4c15:r8C16);
%fill_rating_guide(RG_SA_AtC, AtC_rat, STAR Adult-Getting Treat!r4c17:r8C18);
%fill_rating_guide(RG_SA_HWDC, HWDC_rat, STAR Adult-HWDC!r4c15:r8C16);
%fill_rating_guide(RG_SA_PDRat, PDRat_rat, STAR Adult-Rate Personal Doctor!r4c15:r8C16);

%fill_rating_guide(RG_SP_HPRat, HPRat_rat, STAR+PLUS-Rate Health Plan!r4c15:r8C16);
%fill_rating_guide(RG_SP_AtC, AtC_rat, STAR+PLUS-Getting Treat!r4c17:r8C18);
%fill_rating_guide(RG_SP_HWDC, HWDC_rat, STAR+PLUS-HWDC!r4c15:r8C16);
%fill_rating_guide(RG_SP_PDRat, PDRat_rat, STAR+PLUS-Rate Personal Doctor!r4c15:r8C16);

%fill_rating_guide(RG_SK_HPRat, HPRat_rat, STAR Kids-Rate Health Plan!r4c15:r8C16);
%fill_rating_guide(RG_SK_AtC, AtC_rat, STAR Kids-Getting Treat!r4c17:r8C18);
%fill_rating_guide(RG_SK_SpecTher, SpecTher_rat, STAR Kids-Special Therapy!r4c15:r8C16);
%fill_rating_guide(RG_SK_APM, APM_rat, STAR Kids-Prescriptions!r4c15:r8C16);
%fill_rating_guide(RG_SK_coord, coord_rat, STAR Kids-Care Coordination!r4c15:r8C16);
%fill_rating_guide(RG_SK_GNI, GNI_rat, STAR Kids-Getting Information!r4c15:r8C16);
%fill_rating_guide(RG_SK_transit, transit_rat, STAR Kids-Transition!r4c15:r8C16);
%fill_rating_guide(RG_SK_BHCoun, BHCoun_rat, STAR Kids-Counseling!r4c15:r8C16);

** ---- save and close excel file ----------------------------------------------------------------------------;
data _null_;
	file ddeopen;
	put '[error(false)]';
	put '[save.as("C:\Users\liqian\Dropbox (UFL)\Project_QL\2022_MCO_Report_Cards\SFY2022-2023\3. Survey\Output\MCO_Report_Cards_Survey_Ratings_Draft_1025.xlsx")]';
	* put '[file.close(false)]';
run;
