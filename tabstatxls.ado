
capture program drop tabstatxls
program define tabstatxls

*! version 0.0.8  TomaHawk  27nov2017
version 14
**TODO LIST
syntax varlist [if] [in] [, by(name) Statistics(str asis) Columns(string) format(str asis) NOTotal Missing  ///
                            xlsfile(str) replace sheet(str) sheetmodify sheetreplace cell(str) caption(str asis) note(str asis) ///
                            wintr1(real 40) wintr2(real 30) intc1(str) intc2(str) intc_size(real 30) resc_size(real 16) rows_size(real 15) ///
                            fontname(str asis) fontsize(real 11) pattern_intc(str asis) ///
                            vardisp(string) bold ///
                            dfs1(string) dfs2(string) dfs3(string) dfs4(string) dfs5(string) dfs6(string) dfs7(string) dfs8(string) dfs9(string) dfs10(string)    ///
                            s1(string) s2(string) s3(string) s4(string) s5(string) s6(string) s7(string) s8(string) s9(string) s10(string)    ///
                                             /* options for excel */ ]
**macro list

mata: mata clear

if "`sheet'"=="" local sheet = "Foglio 1"


if "`columns'" == "" local columns = "variables"

local statistics = subinword("`statistics'","q","p25 p50 p75",1)
if "`statistics'"=="" local statistics = "mean"
local nstat = wordcount("`statistics'")

/***   ???
if "`dfs1'" != "" {
  mata: vec_df = J(0,1,"")
  forvalues i=1(1)`nstat' {
    mata: vec_df = vec_df \ "`dfs`i''"
  }
}
************/


local nvar =  wordcount("`varlist'")

if "`fontname'"=="" {
  local font_flag = 0
  local fontname = "Calibri"
}
else local font_flag = 1

local cell = upper("`cell'")
if "`cell'"=="" local cell A1

if "`s1'" == "" local default_stat = 1
else local default_stat = 0

local n_catvar=0 /** serve per non avere problemi nell'if di mata if ("`columns'" == "..." & "`by'" != "") */
if "`by'" != "" {
  local byvar = "`by'"
  local by = "by(`by')"
  qui fre `byvar'
  local n_catvar = r(r)
}
if "`vardisp'" == "" local vardisp = "varlabel"

qui tabstat `varlist' `if' `in', `by' save s(`statistics') c(`columns') `nototal' `missing'
mata: StatTotal = st_matrix("r(StatTotal)")

mata: desc_catvar = J(1,1,.) /** ser per non avere problemi nell'if di mata if ("`columns'" == "..." & "`by'" != "") */
mata: STAT = J(1,1,.) /** serve per non avere problemi nell'if di mata if ("`columns'" == "..." & "`by'" != "") */
if "`by'" != "" {
  mata: desc_catvar = J(0,1,"")
  mata: STAT = J(0,`nvar',.)
  forvalues i=1(1)`n_catvar' {
    **mata: Stat`i' = st_matrix("r(Stat`i')")
    mata: STAT = STAT \ st_matrix("r(Stat`i')")
    **matrix Stat`i' = r(Stat`i') /** se ci sono + variabili in varlist, la matrice sarà del tipo [nstat x n varlist]  **/
    **local desc_catvar`i' = "`r(name`i')'"
    mata: desc_catvar = desc_catvar \ "`r(name`i')'"
  }
}

if substr("`columns'",1,1) == "s" {
  local cols_int = "`statistics'"
  local rows_int = "`varlist'"
  local columns = "statistics"
  local ncols : word count `statistics'
}

if substr("`columns'",1,1) == "v" {
  local cols_int = "`varlist'"
  local rows_int = "`statistics'"
  local columns = "variables"
  local ncols : word count `varlist'
}

if "`replace'" != "" capture erase "`xlsfile'"

if regexm("`cell'","([0-9]*$)") {
      local tryN = regexs(1)
    }
if regexm("`cell'","(^[A-Z]*)") {
      local tryS=  regexs(1)
    }



if "`columns'" == "statistics" {
  mata: vec_colsint = J(1,0,"")
  if "`s1'" == "" {
    foreach i in `cols_int' {
      if "`i'"=="mean" local ii="Media"
      else if "`i'"=="count" local ii="Numero di osservazioni"
      else if "`i'"=="n" local ii="Numero di osservazioni"
      else if "`i'"=="sum" local ii="Sommatoria"
      else if "`i'"=="max" local ii="Massimo"
      else if "`i'"=="min" local ii="Minimo"
      else if "`i'"=="range" local ii="Massimo - Minimo"
      else if "`i'"=="sd" local ii="Deviazione standard"
      else if "`i'"=="variance" local ii="Varianza"
      else if "`i'"=="cv" local ii="Coefficente di variazione"
      else if "`i'"=="semean" local ii="Errore standard della media"
      else if "`i'"=="skewness" local ii="Simmetria"
      else if "`i'"=="kurtosis" local ii="Curtosi"
      else if "`i'"=="p1" local ii="1° percentile"
      else if "`i'"=="p5" local ii="5° percentile"
      else if "`i'"=="p10" local ii="10° percentile"
      else if "`i'"=="p25" local ii="25° percentile"
      else if "`i'"=="median" local ii="Mediana"
      else if "`i'"=="p50" local ii="50° percentile"
      else if "`i'"=="p75" local ii="75° percentile"
      else if "`i'"=="p90" local ii="90° percentile"
      else if "`i'"=="p95" local ii="95° percentile"
      else if "`i'"=="p99" local ii="99° percentile"
      else if "`i'"=="iqr" local ii="Range interquartile"
    mata: vec_colsint = vec_colsint , "`ii'"
    }

  }
  else {
    local cnt=1
    foreach i in `cols_int' {
      local ii = "`s`cnt''"
      local cnt `++cnt'
    mata: vec_colsint = vec_colsint , "`ii'"
    }
  }

  mata: vec_rowsint = J(0,1,"")
  foreach i in `rows_int' {
    local varlab : variable label `i'
    mata: vec_rowsint = vec_rowsint \ "`varlab'"
  }
}


else { /* variables */
  mata: vec_colsint = J(1,0,"")
  foreach i in `cols_int' {
    local varlab : variable label `i'
    mata: vec_colsint = vec_colsint , "`varlab'"
  }

  mata: vec_rowsint = J(0,1,"")
  if "`s1'" == "" {
    foreach i in `rows_int' {
      if "`i'"=="mean" local ii="Media"
      else if "`i'"=="count" local ii="Numero di osservazioni"
      else if "`i'"=="n" local ii="Numero di osservazioni"
      else if "`i'"=="sum" local ii="Sommatoria"
      else if "`i'"=="max" local ii="Massimo"
      else if "`i'"=="min" local ii="Minimo"
      else if "`i'"=="range" local ii="Massimo - Minimo"
      else if "`i'"=="sd" local ii="Deviazione standard"
      else if "`i'"=="variance" local ii="Varianza"
      else if "`i'"=="cv" local ii="Coefficente di variazione"
      else if "`i'"=="semean" local ii="Errore standard della media"
      else if "`i'"=="skewness" local ii="Simmetria"
      else if "`i'"=="kurtosis" local ii="Curtosi"
      else if "`i'"=="p1" local ii="1° percentile"
      else if "`i'"=="p5" local ii="5° percentile"
      else if "`i'"=="p10" local ii="10° percentile"
      else if "`i'"=="p25" local ii="25° percentile"
      else if "`i'"=="median" local ii="Mediana"
      else if "`i'"=="p50" local ii="50° percentile"
      else if "`i'"=="p75" local ii="75° percentile"
      else if "`i'"=="p90" local ii="90° percentile"
      else if "`i'"=="p95" local ii="95° percentile"
      else if "`i'"=="p99" local ii="99° percentile"
      else if "`i'"=="iqr" local ii="Range interquartile"
    mata: vec_rowsint = vec_rowsint \ "`ii'"
    }
  }

  else {
    local cnt=1
    foreach i in `rows_int' {
      local ii = "`s`cnt''"
      local cnt `++cnt'
    mata: vec_rowsint = vec_rowsint \ "`ii'"
    }
  }
}

**set trace on

local enda "end"
mata

b = xl()
if ("`replace'" != "") b.create_book("`xlsfile'", "`sheet'", "xlsx")
if ("`replace'" == "" & "`sheetreplace'"!="") {
  b.load_book("`xlsfile'")
  b.add_sheet("`sheet'")
  b.clear_sheet("`sheet'")
  b.set_sheet("`sheet'")
}

if ("`replace'" == "" & "`sheetmodify'"!="") {
  b.load_book("`xlsfile'")
  b.set_sheet("`sheet'")
}
b.set_mode("open")
b.set_sheet_gridlines("`sheet'", "off")

Ysp = `tryN'
Xsp = b.get_colnum("`tryS'")

if ("`caption'" != "") {
  b.put_string(Ysp,Xsp,"`caption'")
  b.set_font_bold(Ysp,Xsp,"on")
};

if ("`caption'" != "")  Y0 = Ysp+1;
if ("`caption'" == "") Y0 = Ysp;


if ("`columns'"=="statistics" & "`by'"=="") {
  Y1T = Y0 + 1
  X1 = Xsp + 1
  Y1 = Y1T + 1
  Yn = Y1T + cols(StatTotal)
  Xn = X1 + rows(StatTotal) - 1

  if ("`intc1'" !="") b.put_string(Y1T,Xsp,"`intc1'")
  if ("`intc2'" !="") b.put_string(Y1T,X1,"`intc2'")

  b.put_string(Y1T,X1,vec_colsint)
  b.put_string(Y1,Xsp,vec_rowsint)
  b.put_number(Y1,X1,StatTotal')

  //FORMAT
  //font & dimensione
  rfs = (Ysp,Yn)
  cfs = (Xsp,Xn)
  if (`font_flag' == 1) b.set_font(rfs, cfs, "`fontname'", `fontsize')

  //riga intestazione
  cols = (Xsp,Xn)
  b.set_horizontal_align(Y1T,cols,"center")
  b.set_vertical_align(Y1T,cols,"center")
  if ("`bold'"=="bold") b.set_font_bold(Y1T,cols,"on")
  b.set_row_height(Y1T,Y1T, `intc_size')
  b.set_text_wrap(Y1T,cols,"on")
  if ("`pattern_intc'" != "")  b.set_fill_pattern(Y1T,cols,"solid","`pattern_intc'")

  // colonna intestazione righe
  rows = (Y1T,Yn)
  b.set_vertical_align(rows,Xsp,"center")
  b.set_column_width(Xsp, Xsp, `wintr1')
  b.set_row_height(Y1,Yn, `rows_size')
  b.set_text_wrap(rows,Xsp,"on")

  // larghezza e allineamneto colonne della tabella
  rows = (Y1,Yn)
  cols = (X1,Xn)
  b.set_column_width(X1, Xn, `resc_size')
  b.set_vertical_align(rows,cols,"center")
  b.set_horizontal_align(rows,cols,"center")

  // formato numerico
  //default è number_sep_d2
  if ("`dfs1'"=="") {
    b.set_number_format(rows,cols,"number_sep_d2")
  };

  if ("`dfs1'"!="") {
    coli = X1
    b.set_number_format(rows,coli,"`dfs1'")
  }
  if ("`dfs2'"!="") {
    coli = coli+1
    b.set_number_format(rows,coli,"`dfs2'")
  }
  if ("`dfs3'"!="") {
    coli = coli+1
    b.set_number_format(rows,coli,"`dfs3'")
  }
  if ("`dfs4'"!="") {
    coli = coli+1
    b.set_number_format(rows,coli,"`dfs4'")
  }
  if ("`dfs5'"!="") {
    coli = coli+1
    b.set_number_format(rows,coli,"`dfs5'")
  }
  if ("`dfs6'"!="") {
    coli = coli+1
    b.set_number_format(rows,coli,"`dfs6'")
  }
  if ("`dfs7'"!="") {
    coli = coli+1
    b.set_number_format(rows,coli,"`dfs7'")
  }
  if ("`dfs8'"!="") {
    coli = coli+1
    b.set_number_format(rows,coli,"`dfs8'")
  }
  if ("`dfs9'"!="") {
    coli = coli+1
    b.set_number_format(rows,coli,"`dfs9'")
  }
  if ("`dfs10'"!="") {
    coli = coli+1
    b.set_number_format(rows,coli,"`dfs10'")
  }

  // bordi
  cols = (Xsp,Xn)
  b.set_top_border(Y1T,cols,"medium","black")
  b.set_bottom_border(Y1T,cols,"thin","black")
  b.set_bottom_border(Yn,cols,"medium","black")
};


if ("`columns'"=="variables" & "`by'"=="") {
  Y1T = Y0 + 1
  X1 = Xsp + 1
  Y1 = Y1T + 1
  Yn = Y1T + rows(StatTotal)
  Xn = X1 + cols(StatTotal) - 1
  X1
  b.put_string(Y1T,X1,vec_colsint)
  b.put_string(Y1,Xsp,vec_rowsint)
  b.put_number(Y1,X1,StatTotal)

  //FORMAT
  //font & dimensione
  rfs = (Ysp,Yn)
  cfs = (Xsp,Xn)
  if (`font_flag' == 1) b.set_font(rfs, cfs, "`fontname'", `fontsize')

  //riga intestazione
  cols = (Xsp,Xn)
  b.set_horizontal_align(Y1T,cols,"center")
  b.set_vertical_align(Y1T,cols,"center")
  if ("`bold'"=="bold") b.set_font_bold(Y1T,cols,"on")
  b.set_row_height(Y1T,Y1T, `intc_size')
  b.set_text_wrap(Y1T,cols,"on")
  if ("`pattern_intc'" != "")  b.set_fill_pattern(Y1T,cols,"solid","`pattern_intc'")

  // colonna intestazione righe
  rows = (Y1T,Yn)
  b.set_vertical_align(rows,Xsp,"center")
  b.set_column_width(Xsp, Xsp, `wintr1')
  b.set_row_height(Y1,Yn, `rows_size')
  b.set_text_wrap(rows,Xsp,"on")

  // larghezza e allineamneto colonne della tabella
  rows = (Y1,Yn)
  cols = (X1,Xn)
  b.set_column_width(X1, Xn, `resc_size')
  b.set_vertical_align(rows,cols,"center")
  b.set_horizontal_align(rows,cols,"center")

  // formato numerico
  //default è number_sep_d2
  if ("`dfs1'"=="") {
    b.set_number_format(rows,cols,"number_sep_d2")
  };

  if ("`dfs1'"!="") {
    rowi = Y1
    b.set_number_format(rowi,cols,"`dfs1'")
  }
  if ("`dfs2'"!="") {
    rowi = rowi+1
    b.set_number_format(rowi,cols,"`dfs2'")
  }
  if ("`dfs3'"!="") {
    rowi = rowi+1
    b.set_number_format(rowi,cols,"`dfs3'")
  }
  if ("`dfs4'"!="") {
    rowi = rowi+1
    b.set_number_format(rowi,cols,"`dfs4'")
  }
  if ("`dfs5'"!="") {
    rowi = rowi+1
    b.set_number_format(rowi,cols,"`dfs5'")
  }
  if ("`dfs6'"!="") {
    rowi = rowi+1
    b.set_number_format(rowi,cols,"`dfs6'")
  }
  if ("`dfs7'"!="") {
    rowi = rowi+1
    b.set_number_format(rowi,cols,"`dfs7'")
  }
  if ("`dfs8'"!="") {
    rowi = rowi+1
    b.set_number_format(rowi,cols,"`dfs8'")
  }
  if ("`dfs9'"!="") {
    rowi = rowi+1
    b.set_number_format(rowi,cols,"`dfs9'")
  }
  if ("`dfs10'"!="") {
    rowi = rowi+1
    b.set_number_format(rowi,cols,"`dfs10'")
  }

  // bordi
  cols = (Xsp,Xn)
  b.set_top_border(Y1T,cols,"medium","black")
  b.set_bottom_border(Y1T,cols,"thin","black")
  b.set_bottom_border(Yn,cols,"medium","black")
};




if ("`columns'" == "variables" & "`by'" != "") {
  Xint = Xsp + 1
  X1 = Xint + 1
  Xn = X1 + cols(StatTotal) - 1

  Y1T = Ysp + 1
  Y1 = Y1T + 1
  Yn = Y1T + rows(StatTotal)*`n_catvar'
  if ("`nototal'"=="") Yn = Yn + rows(StatTotal);

  if ("`intc1'" !="") b.put_string(Y1T,Xsp,"`intc1'")
  if ("`intc2'" !="") b.put_string(Y1T,Xint,"`intc2'")

  b.put_string(Y1T,X1,vec_colsint)

  rowi = Y1
  js = 1
  je = js + `nstat' -1
  for (j=1; j<=rows(desc_catvar); j++) {
    b.put_string(rowi,Xsp,desc_catvar[j,.])
    b.put_string(rowi,Xint,vec_rowsint)
    b.put_number(rowi,X1,STAT[js..je,.])
    rowi=rowi + `nstat'
    js = js + `nstat'
    je = je + `nstat'
  }

  if ("`nototal'"=="") {
    b.put_string(rowi,Xsp,"Totale")
    b.put_string(rowi,Xint,vec_rowsint)
    b.put_number(rowi,X1,StatTotal)
  }

  //FORMAT
  //font & dimensione
  rfs = (Ysp,Yn)
  cfs = (Xsp,Xn)
  if (`font_flag' == 1) b.set_font(rfs, cfs, "`fontname'", `fontsize')

  //riga intestazione
  cols = (Xsp,Xn)
  b.set_horizontal_align(Y1T,cols,"center")
  b.set_vertical_align(Y1T,cols,"center")
  if ("`bold'"=="bold") b.set_font_bold(Y1T,cols,"on")
  b.set_row_height(Y1T,Y1T, `intc_size')
  b.set_text_wrap(Y1T,cols,"on")
  if ("`pattern_intc'" != "")  b.set_fill_pattern(Y1T,cols,"solid","`pattern_intc'")

  // colonna intestazione righe
  rows = (Y1T,Yn)
  b.set_vertical_align(rows,Xsp,"center")
  b.set_column_width(Xsp, Xsp, `wintr1')
  b.set_row_height(Y1,Yn, `rows_size')
  b.set_text_wrap(rows,Xsp,"on")

  // colonna delle statistiche
  b.set_vertical_align(rows,Xint,"center")
  b.set_column_width(Xint, Xint, `wintr2')

  // larghezza e allineamneto colonne della tabella
  rows = (Y1,Yn)
  cols = (X1,Xn)
  b.set_column_width(X1, Xn, `resc_size')
  b.set_vertical_align(rows,cols,"center")
  b.set_horizontal_align(rows,cols,"center")
  b.set_column_width(X1, Xn, `resc_size')

  // formato numerico
  //default è number_sep_d2
  if ("`dfs1'"=="") {
    b.set_number_format(rows,cols,"number_sep_d2")
  };

  if ("`dfs1'"!="") {
    rowi = Y1
    rowi_sj = rowi //riga i-esima per la statistica j-esima
    for (j=1; j<=`nstat'; j++) {
      b.set_number_format(rowi_sj,cols,"`dfs1'")
      rowi_sj = rowi_sj + `nstat'
    }
  }
  if ("`dfs2'"!="") {
    rowi = rowi+1
    rowi_sj = rowi
    for (j=1; j<=`nstat'; j++) {
      b.set_number_format(rowi_sj,cols,"`dfs2'")
      rowi_sj = rowi_sj + `nstat'
    }
  }
  if ("`dfs3'"!="") {
    rowi = rowi+1
    rowi_sj = rowi
    for (j=1; j<=`nstat'; j++) {
      b.set_number_format(rowi_sj,cols,"`dfs3'")
      rowi_sj = rowi_sj + `nstat'
    }
  }
  if ("`dfs4'"!="") {
    rowi = rowi+1
    rowi_sj = rowi
    for (j=1; j<=`nstat'; j++) {
      b.set_number_format(rowi_sj,cols,"`dfs4'")
      rowi_sj = rowi_sj + `nstat'
    }
  }
  if ("`dfs5'"!="") {
    rowi = rowi+1
    rowi_sj = rowi
    for (j=1; j<=`nstat'; j++) {
      b.set_number_format(rowi_sj,cols,"`dfs5'")
      rowi_sj = rowi_sj + `nstat'
    }
  }
  if ("`dfs6'"!="") {
    rowi = rowi+1
    rowi_sj = rowi
    for (j=1; j<=`nstat'; j++) {
      b.set_number_format(rowi_sj,cols,"`dfs6'")
      rowi_sj = rowi_sj + `nstat'
    }
  }
  if ("`dfs7'"!="") {
    rowi = rowi+1
    rowi_sj = rowi
    for (j=1; j<=`nstat'; j++) {
      b.set_number_format(rowi_sj,cols,"`dfs7'")
      rowi_sj = rowi_sj + `nstat'
    }
  }
  if ("`dfs8'"!="") {
    rowi = rowi+1
    rowi_sj = rowi
    for (j=1; j<=`nstat'; j++) {
      b.set_number_format(rowi_sj,cols,"`dfs8'")
      rowi_sj = rowi_sj + `nstat'
    }
  }
  if ("`dfs9'"!="") {
    rowi = rowi+1
    rowi_sj = rowi
    for (j=1; j<=`nstat'; j++) {
      b.set_number_format(rowi_sj,cols,"`dfs9'")
      rowi_sj = rowi_sj + `nstat'
    }
  }
  if ("`dfs10'"!="") {
    rowi = rowi+1
    rowi_sj = rowi
    for (j=1; j<=`nstat'; j++) {
      b.set_number_format(rowi_sj,cols,"`dfs10'")
      rowi_sj = rowi_sj + `nstat'
    }
  }

  // bordi
  rowi = Y1 + `nstat' - 1
  cols = (Xint,Xn)
  for (j=1; j<=`n_catvar'; j++) {
    b.set_bottom_border(rowi,cols,"dotted","gray")
    rowi=rowi+`nstat'
  }

  cols = (Xsp,Xn)
  b.set_top_border(Y1T,cols,"medium","black")
  b.set_bottom_border(Y1T,cols,"thin","black")
  b.set_bottom_border(Yn,cols,"medium","black")
};



if ("`columns'"=="statistics" & "`by'"!="") {
  Xint = Xsp + 1
  X1 = Xint + 1
  Xn = X1 + rows(StatTotal) - 1

  Y1T = Ysp + 1
  Y1 = Y1T + 1
  Yn = Y1T + cols(StatTotal)*`n_catvar'
  if ("`nototal'"=="") Yn = Yn + cols(StatTotal)

  if ("`intc1'" !="") b.put_string(Y1T,Xsp,"`intc1'")
  if ("`intc2'" !="") b.put_string(Y1T,Xint,"`intc2'")

  b.put_string(Y1T,X1,vec_colsint)

  rowi = Y1
  js = 1
  je = js + `nstat' -1
  for (j=1; j<=rows(desc_catvar); j++) {
    b.put_string(rowi,Xsp,desc_catvar[j,.])
    b.put_string(rowi,Xint,vec_rowsint)
    b.put_number(rowi,X1,STAT[js..je,.]')
    rowi=rowi + `nvar'
    js = js + `nstat'
    je = je + `nstat'
  }

  if ("`nototal'"=="") {
    b.put_string(rowi,Xsp,"Totale")
    b.put_string(rowi,Xint,vec_rowsint)
    b.put_number(rowi,X1,StatTotal')
  }

  if ("`note'"!="" ) {
    Ynote = Yn+1
    b.put_string(Ynote,Xsp,"`note'")
  }





  //FORMAT
  //font & dimensione
  rfs = (Ysp,Yn)
  cfs = (Xsp,Xn)
  if (`font_flag' == 1) b.set_font(rfs, cfs, "`fontname'", `fontsize')

  //riga intestazione
  cols = (Xsp,Xn)
  b.set_horizontal_align(Y1T,cols,"center")
  b.set_vertical_align(Y1T,cols,"center")
  if ("`bold'"=="bold") b.set_font_bold(Y1T,cols,"on")
  b.set_row_height(Y1T,Y1T, `intc_size')
  b.set_text_wrap(Y1T,cols,"on")
  if ("`pattern_intc'" != "")  b.set_fill_pattern(Y1T,cols,"solid","`pattern_intc'")

  // colonna intestazione righe
  rows = (Y1T,Yn)
  b.set_vertical_align(rows,Xsp,"center")
  b.set_column_width(Xsp, Xsp, `wintr1')
  b.set_row_height(Y1,Yn, `rows_size')
  b.set_text_wrap(rows,Xsp,"on")

  // colonna delle variabili
  b.set_vertical_align(rows,Xint,"center")
  b.set_column_width(Xint, Xint, `wintr2')

  // larghezza e allineamneto colonne della tabella
  rows = (Y1,Yn)
  cols = (X1,Xn)
  b.set_column_width(X1, Xn, `resc_size')
  b.set_vertical_align(rows,cols,"center")
  b.set_horizontal_align(rows,cols,"center")

  // formato numerico
  //default è number_sep_d2
  if ("`dfs1'"=="") {
    b.set_number_format(rows,cols,"number_sep_d2")
  };

  if ("`dfs1'"!="") {
    coli = X1
    b.set_number_format(rows,coli,"`dfs1'")
  }
  if ("`dfs2'"!="") {
    coli = coli+1
    b.set_number_format(rows,coli,"`dfs2'")
  }
  if ("`dfs3'"!="") {
    coli = coli+1
    b.set_number_format(rows,coli,"`dfs3'")
  }
  if ("`dfs4'"!="") {
    coli = coli+1
    b.set_number_format(rows,coli,"`dfs4'")
  }
  if ("`dfs5'"!="") {
    coli = coli+1
    b.set_number_format(rows,coli,"`dfs5'")
  }
  if ("`dfs6'"!="") {
    coli = coli+1
    b.set_number_format(rows,coli,"`dfs6'")
  }
  if ("`dfs7'"!="") {
    coli = coli+1
    b.set_number_format(rows,coli,"`dfs7'")
  }
  if ("`dfs8'"!="") {
    coli = coli+1
    b.set_number_format(rows,coli,"`dfs8'")
  }
  if ("`dfs9'"!="") {
    coli = coli+1
    b.set_number_format(rows,coli,"`dfs9'")
  }
  if ("`dfs10'"!="") {
    coli = coli+1
    b.set_number_format(rows,coli,"`dfs10'")
  }

  // bordi
  rowi = Y1 + `nvar' - 1
  cols = (Xint,Xn)
  for (j=1; j<=`n_catvar'; j++) {
    b.set_bottom_border(rowi,cols,"dotted","gray")
    rowi=rowi+`nvar'
  }

  cols = (Xsp,Xn)
  b.set_top_border(Y1T,cols,"medium","black")
  b.set_bottom_border(Y1T,cols,"thin","black")
  b.set_bottom_border(Yn,cols,"medium","black")
};

 if ("`note'"!="") {
  fontsize_note = `fontsize' - 2
  b.set_font(Ynote, Xsp , "`fontname'", fontsize_note)
 }



b.close_book()
`enda'
di as txt _n `"Apri il file excel:  {ul:{bf:{browse `"`c(pwd)'/`xlsfile'"':`xlsfile'}}} "'

end
exit
