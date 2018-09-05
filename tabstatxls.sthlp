{smcl}
{hline}
help for {hi:tabstatxls}
{hline}

{title:Esportare l'output di tabstat in Microsoft Excel}

{p 8 12 2}
{cmd:tabstatxls} {it:varlist} {ifin} [{cmd:,} {help tabstatxls##tabstatopt:{it:tabstat_options}}]  {help tabstatxls##latexopt:{it:excel_options}}


{title:Description}

{p 4 4 2}{cmd:tabstatxls} permette di esportare in LaTex l'output del comando {cmd:tabstat}. {it:varlist} è la lista delle variabili di cui si vogliono esportare le statistiche.
Il comado usa la classe mata {cmd:xl()} per esportare in Excel 1997/2003 i files di estensione .xls e in Excel 2007/2013 i files di estensione .xlsx. Per funzionare
è richiesta la presenza del comando {cmd:fre}.


{marker tabstatopt}{title:tabstat options}

{p 4 8 2}{cmd:by(varname)}: specifica che le statistiche delle variabili specificate in {it:varlist} devono essere visualizzate condizionando per la variabile specificata in {cmd:by(varname)}

{p 4 8 2}{cmdab:s:tatistics(}{it:statname}{cmd:)}: specifica quali statistiche devono essere visualizzate. Se non si specifica nulla viene calcolata la sola media. Ogni statistica deve essere separata da uno spazio. Le possibili statistiche sono:

{synoptset 17}{...}
{synopt:{space 4}{it:statname}}Definizione{p_end}
{space 4}{synoptline}
{synopt:{space 4}{opt me:an}} media{p_end}
{synopt:{space 4}{opt co:unt}} numero di osservazioni non missing{p_end}
{synopt:{space 4}{opt n}} uguale a {cmd:count}{p_end}
{synopt:{space 4}{opt su:m}} sommatoria{p_end}
{synopt:{space 4}{opt ma:x}} massimo{p_end}
{synopt:{space 4}{opt mi:n}} minimo{p_end}
{synopt:{space 4}{opt r:ange}} range = {opt max} - {opt min}{p_end}
{synopt:{space 4}{opt sd}} standard deviation{p_end}
{synopt:{space 4}{opt v:ariance}} varianza{p_end}
{synopt:{space 4}{opt cv}} coefficiente di variatione ({cmd:sd/mean}){p_end}
{synopt:{space 4}{opt sem:ean}} standard error della media ({cmd:sd/sqrt(n)}){p_end}
{synopt:{space 4}{opt sk:ewness}} simmetria{p_end}
{synopt:{space 4}{opt k:urtosis}} curtosi{p_end}
{synopt:{space 4}{opt p1}} 1° percentile{p_end}
{synopt:{space 4}{opt p5}} 5° percentile{p_end}
{synopt:{space 4}{opt p10}} 10° percentile{p_end}
{synopt:{space 4}{opt p25}} 25° percentile{p_end}
{synopt:{space 4}{opt med:ian}} mediana (equivalente a  {opt p50}){p_end}
{synopt:{space 4}{opt p50}} 50° percentile (equivalente a {opt median}){p_end}
{synopt:{space 4}{opt p75}} 75° percentile{p_end}
{synopt:{space 4}{opt p90}} 90° percentile{p_end}
{synopt:{space 4}{opt p95}} 95° percentile{p_end}
{synopt:{space 4}{opt p99}} 99° percentile{p_end}
{synopt:{space 4}{opt iqr}} range interquartile = {opt p75} - {opt p25}{p_end}
{synopt:{space 4}{opt q}} equivale a specificare {cmd:p25 p50 p75}{p_end}
{space 4}{synoptline}
{p2colreset}{...}

{p 4 8 2}{cmdab:c:olumns(}{cmdab:v:ariables|}{cmdab:s:tatistics)}: specifica cosa deve essere visulizzato in colonna. {cmd:variables} visualizza le variabili di {it:varlist} (opzione di default), {cmd:statistics}
visualizza le statistiche specificate nell'opzione {cmd:statistics({it:statname})}.

{p 4 8 2}{opt f:ormat}{cmd:(%}{it:{help format:fmt}}{cmd:)} specifica il formato generale di visualizzazione delle statistiche. Il formato di default è {cmd:%12.2gc}.

{p 4 8 2}{opt not:otal} non visualizza le statistiche generali; si usa sole se viene specificata l'opzione {opt by(varname)}.

{p 4 8 2}{opt m:issing} visualizza le statistiche anche per i valori missing della variabile {opt by(varname)}.



{marker excelopt}{title:excel options}

{p 4 8 2}{cmd:xlsfile(filename.ext)}: specifica il file .xls o .xlsx (ed eventuale percorso) in cui salvare il codice della tabella. Questa opzione e l'estensione del file sono obbligatori.

{p 4 8 2}{cmd:sheet(sheetname)}: specifica il nome del foglio in cui scrivere l'output. Di default si usa "Foglio 1".

{p 4 8 2}{cmd:replace}: specifica di sovrascrivere il file indicato in {cmd:texfile(filename.ext)}.

{p 4 8 2}{cmd:sheetreplace}: specifica di sovrascrivere il foglio indicato in {cmd:sheet(sheetname)}.

{p 4 8 2}{cmd:sheetmodify}: specifica di modificare il foglio indicato in {cmd:sheet(sheetname)}.

{p 4 8 2}{cmd:cell}: specifica la cella da cui iniziare l'output Di default si usa A1. Usare solo la notazione lettera e numero.

{p 4 8 2}{cmd:caption(string)}: specifica il testo da inserire come titolo della tabella. Di default è vuoto.

{p 4 8 2}{cmd:note(string)}: specifica il testo da inserire come nota a piè di tabella. Di default è vuoto.

{p 4 8 2}{cmd:intc1(string)}: specifica il testo da inserire come descrizione della prima colonna della tabella. In assenza dell'opzione {opt by(varname)} nella prima colonna ci possono essere le
variabili o le statistiche, dipende da cosa specificato nell'opzione {cmd:columns()} e in questo caso di default {cmd:intc1()} è vuoto. Se viene specificata l'opzione {opt by(varname)},
nella prima colonna ci sono i valori della variabile {opt varname} e di default in {cmd:intc1()} c'è la descrizione associata a {opt varname}.

{p 4 8 2}{cmd:intc2(string)}: specifica il testo da inserire come descrizione della seconda colonna della tabella e si applica solo nel caso in cui sia specificata l'opzione {cmd:by(varname)}.
Se l'opzione {cmd:columns()} è {cmd:variables} il default è {cmd:intc2(Statistiche)}, se l'opzione {cmd:columns()} è {cmd:statistics} il default è {cmd:intc2(Variabili)}.

{p 4 8 2}{cmd:wintr1(number)}: specifica la larghezza della prima colonna della tabella. In assenza dell'opzione {opt by(varname)} nella prima colonna ci possono essere le
variabili o le statistiche, se l'opzione {opt by(varname)} è specificata nella prima colonna ci sono i valori della variabile {opt varname}.
Di default il valore è pari a 40.

{p 4 8 2}{cmd:wintr2(number)}: specifica la larghezza della seconda colonna della tabella e si applica solo nel caso in cui sia specificata l'opzione {cmd:by(varname)}.
Se l'opzione {cmd:columns()} è {cmd:variables} nella seconda colonna ci sono le statistiche, se l'opzione {cmd:columns()} è {cmd:statistics} ci sono le variabili.
Di default il valore è pari a 30.

{p 4 8 2}{cmd:intc_size(number)}: specifica l'altezza della prima riga della tabella. La prima riga contiene la descrizione delle variabili o delle statistiche
a seconda di cosa è specificato in {cmd:columns()}. Di default il valore è pari a 30.

{p 4 8 2}{cmd:resc_size(number)}: specifica la larghezza delle colonne del corpo della tabella cioè delle colonne con i risultati delle statistiche specificate in
{cmd:statistics(}{it:statname}{cmd:)}. Di default il valore è 16.

{p 4 8 2}{cmd:nodispstat}: sopprime la dicitura della statistica nelle intestazioni di riga. Questa opzione è abilitata solo quando vengono specificate le opzioni
{cmd:columns(variables)} e {cmd:by(varname)}

{p 4 8 2}{cmd:rows_size(number)}: specifica l'altezza delle righe del corpo della tabella. Di default il valore è 15.

{p 4 8 2}{cmd:fontname(string)}: specifica il font da usare nella tabella. Il default è {cmd:fontname(Calibri)}

{p 4 8 2}{cmd:fontsize(number)}: specifica la dimensione del font usato nella tabella. Il default è 11.

{p 4 8 2}{cmd:bcolor_intc(string)}: specifica il colore di sfondo della prima riga della tabella. I colori possono essere indicati nel formato RGB
all'interno di virgolette ({cmd:pattern_intc("255 255 255")} o usando uno dei colori predefiniti da Stata per l'esportazione in Excel, vedi
 {cmd:{help [M-5] xl():[M-5] xl()}} alla sezione Format colors. Di default non è previsto nessun colore.

{p 4 8 2}{cmd:pattern_intc(string)}: specifica il pattern di riempimento della prima riga della tabella. Vedi {cmd:{help [M-5] xl():[M-5] xl()}} alla sezione Codes for fill pattern styles. Di default non è solid.

{p 4 8 2}{cmd:vardisp(varlabel|varname)}: specifica come visualizzare le variabili specificate in {it:varlist}. {cmd:vardisp(varlabel)} visualizza la descrizione associata a ciascuna variabile,
{cmd:vardisp(varname)} visualizza solo il nome della variabile. {cmd:vardisp(varlabel)} è il default. DA VERIFICARE !!

{p 4 8 2}{cmd:bold}: specifica di formattare in bold la prima riga della tabella (intestazioni delle colonne).

{p 4 8 2}{cmd:s1(string)...s10(string)}: specifica la descrizione delle statistiche indicate nell'opzione {cmd:statistics({it:statname})}. L'ordine deve essere quello di {it:statname}, ovvero
{cmd:s1()} indica la descrizione della prima statistica di {cmd:statistics({it:statname})}, {cmd:s2()} indica la descrizione della seconda statistica di {cmd:statistics({it:statname})} e così via.
Queste sono le descrizioni di default:

{synoptset 17}{...}
{synopt:{space 4}{it:statname}}Descrizione{p_end}
{space 4}{synoptline}
{synopt:{space 4}{opt mean}} Media{p_end}
{synopt:{space 4}{opt count}} Numero di osservazioni{p_end}
{synopt:{space 4}{opt n}} Numero di osservazioni{p_end}
{synopt:{space 4}{opt sum}} Sommatoria{p_end}
{synopt:{space 4}{opt max}} Massimo{p_end}
{synopt:{space 4}{opt min}} Minimo{p_end}
{synopt:{space 4}{opt range}} Massimo - Minimo{p_end}
{synopt:{space 4}{opt sd}} Deviazione standard{p_end}
{synopt:{space 4}{opt variance}} Varianza{p_end}
{synopt:{space 4}{opt cv}} Coefficiente di variatione{p_end}
{synopt:{space 4}{opt semean}} Errore standard della media{p_end}
{synopt:{space 4}{opt skewness}} Simmetria{p_end}
{synopt:{space 4}{opt kurtosis}} Curtosi{p_end}
{synopt:{space 4}{opt p1}} 1° percentile{p_end}
{synopt:{space 4}{opt p5}} 5° percentile{p_end}
{synopt:{space 4}{opt p10}} 10° percentile{p_end}
{synopt:{space 4}{opt p25}} 25° percentile{p_end}
{synopt:{space 4}{opt median}} Mediana{p_end}
{synopt:{space 4}{opt p50}} 50° percentile{p_end}
{synopt:{space 4}{opt p75}} 75° percentile{p_end}
{synopt:{space 4}{opt p90}} 90° percentile{p_end}
{synopt:{space 4}{opt p95}} 95° percentile{p_end}
{synopt:{space 4}{opt p99}} 99° percentile{p_end}
{synopt:{space 4}{opt iqr}} Range interquartile{p_end}
{space 4}{synoptline}
{p2colreset}{...}

{p 4 8 2}{cmd:dfs1(string)...dfs10(string)}: specifica il formato numerico delle statistiche indicate nell'opzione {cmd:statistics({it:statname})}. L'ordine deve essere quello di {it:statname}, ovvero
{cmd:dfs1()} indica il formato della prima statistica di {cmd:statistics({it:statname})}, {cmd:dfs2()} indica il formato della seconda statistica di {cmd:statistics({it:statname})} e così via.
La sintassi del formato è la medesima di Mata per i formati numerici nell'esportazione in Excel (vedi {cmd:{help [M-5] xl():[M-5] xl()}} alla sezione Codes for numeric formats).

{synoptset 30}{...}
{synopt:{space 4}Formato}Esempio{p_end}
{space 4}{synoptline}
{synopt:{space 4}{opt number}}1000{p_end}
{synopt:{space 4}{opt number_d2}}1000.00{p_end}
{synopt:{space 4}{opt number_sep}}100,000{p_end}
{synopt:{space 4}{opt number_sep_d2}}100,000.00{p_end}
{synopt:{space 4}{opt number_sep_negbra}}(1,000){p_end}
{synopt:{space 4}{opt number_sep_negbrared}}(1,000){p_end}
{synopt:{space 4}{opt number_d2_sep_negbra}}(1,000.00){p_end}
{synopt:{space 4}{opt number_d2_sep_negbrared}}(1,000.00){p_end}
{synopt:{space 4}{opt currency_negbra}}($4000){p_end}
{synopt:{space 4}{opt currency_negbrared}}($4000){p_end}
{synopt:{space 4}{opt currency_d2_negbra}}($4000.00){p_end}
{synopt:{space 4}{opt currency_d2_negbrared}}($4000.00){p_end}
{synopt:{space 4}{opt account}}5,000{p_end}
{synopt:{space 4}{opt accountcur}}$     5,000{p_end}
{synopt:{space 4}{opt account_d2}}5,000.00{p_end}
{synopt:{space 4}{opt account_d2_cur}}$  5,000.00{p_end}
{synopt:{space 4}{opt percent}}75%{p_end}
{synopt:{space 4}{opt percent_d2}}75.00%{p_end}
{synopt:{space 4}{opt scientific_d2}}10.00E+1{p_end}
{synopt:{space 4}{opt fraction_onedig}}10 1/2{p_end}
{synopt:{space 4}{opt fraction_twodig}}10 23/95{p_end}
{synopt:{space 4}{opt date}}3/18/2007{p_end}
{synopt:{space 4}{opt date_d_mon_yy}}18-Mar-07{p_end}
{synopt:{space 4}{opt date_d_mon}}18-Mar{p_end}
{synopt:{space 4}{opt date_mon_yy}}Mar-07{p_end}
{synopt:{space 4}{opt time_hmm_AM}}8:30 AM{p_end}
{synopt:{space 4}{opt time_HMMSS_AM}}8:30:00 AM{p_end}
{synopt:{space 4}{opt time_HMM}}8:30{p_end}
{synopt:{space 4}{opt time_HMMSS}}8:30:00{p_end}
{synopt:{space 4}{opt time_MMSS}}30:55{p_end}
{synopt:{space 4}{opt time_H0MMSS}}20:30:55{p_end}
{synopt:{space 4}{opt time_MMSS0}}30:55.0{p_end}
{synopt:{space 4}{opt date_time}}3/18/2007 8:30{p_end}
{synopt:{space 4}{opt text}}this is text{p_end}
{space 4}{synoptline}
{p2colreset}{...}

 {p 4 8 2}É possibile usare anche i formati numerici personalizzati, vedi {cmd:{help [M-5] xl():[M-5] xl()}} alla sezione Custom formatting.



{title:Examples}

{p 4 8 2}{cmd:. tabstatxls price weight mpg rep78, stat(n mean cv q) col(stat) xlsfile(tabstat.xlsx) wintr1(20) replace sheet(sheet1) cell(A1) dfs1(number) dfs2(number_d2) dfs3(number_d2) dfs4(number) dfs5(number) dfs6(number) resc_size(14)}{p_end}

{p 4 8 2}{cmd:. tabstatxls price weight mpg rep78, stat(n mean cv q) col(stat) xlsfile(tabstat.xlsx) wintr1(20) replace sheet(sheet1) cell(A1) dfs1(number) dfs2(number_d2) dfs3(number_d2) dfs4(number) dfs5(number) dfs6(number) resc_size(14)}{p_end}

{p 4 8 2}{cmd:. tabstatxls price weight mpg rep78, stat(n mean cv q) col(v) xlsfile(tabstat.xlsx) sheetmodify sheet(sheet1) cell(I1) dfs1(number) dfs2(number_d2) dfs3(number_d2) dfs4(number) dfs5(number) dfs6(number)}{p_end}

{p 4 8 2}{cmd:. tabstatxls price weight mpg rep78, stat(n mean cv q) col(v) by(foreign) xlsfile(tabstat.xlsx) sheetmodify cell(A80) sheet(esempi excel) wintr1(12) wintr2(23) s1(Nonmissing observations) s2(Mean)}{p_end}
 {cmd:       s3(Coefficient of variation) s4(25th percentile) s5(50th percentile) s6(75th percentile)}

{p 4 8 2}{cmd:. tabstatxls price weight mpg rep78, stat(mean sd min max) c(s) by(foreign) xlsfile(tabstat.xlsx) sheetmodify cell(A40) sheet(esempi excel) wintr1(12) wintr2(20) fontname(Times New Roman) fontsize(9) pattern_intc(silver)}{p_end}



{title:Author}

{p 4 4 2}Nicola Tommasi{break}
nicola.tommasi@univr.it{break}
nicola.tommasi@gmail.com


{title:Acknowledgments}

{p 4 4 2}



{title:References}




{title:Also see}

{p 4 13 2}Online:
{help tabstat},
{help tabstattex} (if installed)
