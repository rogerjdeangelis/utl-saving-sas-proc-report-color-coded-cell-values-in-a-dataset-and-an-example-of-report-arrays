%let pgm=utl-saving-sas-proc-report-color-coded-cell-values-in-a-dataset-and-an-example-of-report-arrays;

%stop_submission;

Saving sas proc report color coded cell values in a dataset and an example of report arrays

SOAPBOX ON
  I am not a fan of complex processing in proc report.
  I like to use report as an enhanced 'proc print'
  The complex report is unable to provide a datasets with the column information.
SOAPBOX OFF

TWO SOLUTIONS

 1 hard coded columns
   A more flexuble datastep language
   Debugging is much easier in a datastep?
   Could make dynamic with array/do_over macros.
   PROVIDES AOUTPUT DTASET WITH COLOR INFOMATION
   YOU CAN ASO CAN GET A DATASET USING OUT= IN OROC REPORT STATEMENT

 2 dynamic report arrays
   ksharp
   https://tinyurl.com/2dbcesx9
   DOES NOT PROVIDE A DATSETS WITH COLOR COLUMNS
   COLS STATEMENT DOES NOT HAVE COLOR COLUMNS

github
https://tinyurl.com/2d6ff6h8
https://github.com/rogerjdeangelis/utl-saving-sas-proc-report-color-coded-cell-values-in-a-dataset-and-an-example-of-report-arrays

PROBLEM
=======

 Create this output
 (note cannot show color so i added the text. Not a fan of html for documentation purposed)

  d:/xls/colorx.xlsx

  ------------------------
  | A1| fx        | TYPE |
  --------------------------------------------+
  [_] |    A     |    B  |   C     |  D       |
  --------------------------------------------|
   1  | TYPE     |_202008|_202009  |_292009   |
   -- |----------+-------+---------+----------|
   2  |  TypeA   |312    |283 GREEN|253  GREEN|
   -- |----------+-------+---------+----------|
   3  |  TypeB   |296    |310  RED |313  RED  |
   -- |----------+-------+---------+----------|
   4  |  TypeC   |545    |553  RED |590  RED  |
   -- |----------+-------+---------+----------|
   5  |  TypeD   |1697    |1784 RED|1813 RED  |
   -- |----------+-------+---------+----------|
   6  |  Sum     |2850   |2930 RED |2969 RED  |
   -- |----------+-------+---------+----------+
  [COLOR]


communities.sas
https://tinyurl.com/ykbnufhh
https://communities.sas.com/t5/SAS-Programming/Conditional-Formatting-using-Proc-Report/m-p/763428#M241783

ksharp
https://tinyurl.com/2dbcesx9
https://communities.sas.com/t5/user/viewprofilepage/user-id/18408


RELATED REPOs
---------------------------------------------------------------------------------------------------
https://github.com/rogerjdeangelis/utl-get-the-color-of-a-cell-in-excel-xlsx
https://github.com/rogerjdeangelis/utl-ods-excel-color-code-every-other-column-in-a-specified-row

github (this technique applies to any proc that can use formchar ie tabulate freq report )
https://tinyurl.com/r3s8f5fy
https://github.com/rogerjdeangelis/utl-creating-a-proc-tabulate-cross-tabulation-sas-dataset
https://github.com/rogerjdeangelis/utl-proc-report-greenbar-or-alternate-shading-of-rows-for-easy-reading

/**************************************************************************************************************************/
/* INPUT               | PROCESS                                             | OUTPUT                                     */
/* =====               | =======                                             | ======                                     */
/*  DATE   TYPE  EXCPT | 1 HARD CODED COLUMNS                                | d:/xls/color.xlsx                          */
/*                     | ====================                                |                                            */
/* 2020_08 TypeA   312 |                                                     | ------------------                         */
/* 2020_08 TypeB   296 | ods exclude all;                                    | | A1| fx   | TYPE |                        */
/* 2020_08 TypeC   545 | ods output observed=                                | ----------------------------------------+  */
/* 2020_08 TypeD  1697 |  havtab (drop=sum);                                 | [_] |   A  |    B  |   C     |  D       |  */
/* 2020_09 TypeA   283 | proc corresp data=have                              | ----------------------------------------|  */
/* 2020_09 TypeB   310 |  observed dim=2;                                    |  1  |TYPE |_2020_08|_2020_09 |_2920_09  |  */
/* 2020_09 TypeC   553 | tables type, date ;                                 |  -- |-----+--------+---------+----------|  */
/* 2020_09 TypeD  1784 | weight excpt;                                       |  2  |TypeA|312     |283 GREEN|253  GREEN|  */
/* 2020_10 TypeA   253 | run;quit;                                           |  -- |-----+--------+---------+----------|  */
/* 2020_10 TypeB   313 | ods select all;                                     |  3  |TypeB|296     |310  RED |313  RED  |  */
/* 2020_10 TypeC   590 |                                                     |  -- |-----+--------+---------+----------|  */
/* 2020_10 TypeD  1813 | Label  _2020_08 _2020_09 _2020_10                   |  4  |TypeC|545     |553  RED |590  RED  |  */
/*                     |                                                     |  -- |-----+--------+---------+----------|  */
/* CONTENTS            | TypeA     312      283      253                     |  5  |TypeD|1697    |1784 RED |1813 RED  |  */
/*  Var    Type Len    | TypeB     296      310      313                     |  -- |-----+--------+---------+----------|  */
/*   DATE   Char 7 Char| TypeC     545      553      590                     |  6  |Sum  |2850    |2930 RED |2969 RED  |  */
/*   TYPE   Char 8     | TypeD    1697     1784     1813                     |  -- |-----+--------+---------+----------+  */
/*   EXCPT  Num  8     | Sum      2850     2930     2969                     | [COLOR]                                    */
/*                     |                                                     |                                            */
/*  data have;         |                                                     |                                            */
/*   input             | data color (                                        |                                            */
/*    date $7.         |  drop=idx clr1 clr2                                 | WORK.COLOR                 COLOR INFO      */
/*    type $           |  rename=(clr3-clr4=c2020_09-c2020_10));             |                          ==============    */
/*    excpt;           |  set havtab;                                        |         2020  2020 2020   C2020   C2020    */
/*  cards4;            |                                                     |  LABEL    08    09   10      09      10    */
/*  2020_08 TypeA 312  |  array mon[*] _numeric_ ;                           |                                            */
/*  2020_08 TypeB 296  |  array clr[%utl_varcount(havtab)] $80 ;             |  TypeA   312   283  253   GREEN   GREEN    */
/*  2020_08 TypeC 545  |   do idx= 2 to dim1(mon);                           |  TypeB   296   310  313   RED     RED      */
/*  2020_08 TypeD 1697 |   select;                                           |  TypeC   545   553  590   RED     RED      */
/*  2020_09 TypeA 283  |    when (mon[idx]=mon[idx-1]) clr[idx+1]="YELLOW";  |  TypeD  1697  1784 1813   RED     RED      */
/*  2020_09 TypeB 310  |    when (mon[idx]>mon[idx-1]) clr[idx+1]="RED";     |  Sum    2850  2930 2969   RED     RED      */
/*  2020_09 TypeC 553  |    when (mon[idx]<mon[idx-1]) clr[idx+1]="GREEN";   |                                            */
/*  2020_09 TypeD 1784 |    otherwise clr[idx+1] = "BLUE";                   |                                            */
/*  2020_10 TypeA 253  |   end;                                              |  FROM REPORT WORK.COLORINFO                */
/*  2020_10 TypeB 313  |  end;                                               |                                            */
/*  2020_10 TypeC 590  | run;quit;                                           |        COLOR INFO                          */
/*  2020_10 TypeD 1813 |                             COLOR INFO              |  LABEL C2020  C2020   _2020 _2020 _2020    */
/*  ;;;;               |         2020  2020 2020   C2020   C2020             |           09     10      08    09    10    */
/*  run;quit;          |  LABEL    08    09   10      09      10             |                                            */
/*                     |                                                     |  TypeA  GREEN  GREEN    312   283   253    */
/*                     |  TypeA   312   283  253   GREEN   GREEN             |  TypeB  RED    RED      296   310   313    */
/*                     |  TypeB   296   310  313   RED     RED               |  TypeC  RED    RED      545   553   590    */
/*                     |  TypeC   545   553  590   RED     RED               |  TypeD  RED    RED     1697  1784  1813    */
/*                     |  TypeD  1697  1784 1813   RED     RED               |  Sum    RED    RED     2850  2930  2969    */
/*                     |  Sum    2850  2930 2969   RED     RED               |                                            */
/*                     |                                                     |                                            */
/*                     | %utlfkil(d:/xls/color.xlsx);                        |                                            */
/*                     | ods excel file="d:/xls/color.xlsx"                  |                                            */
/*                     |     style=journal options(sheet_name="color");      |                                            */
/*                     |                                                     |                                            */
/*                     | proc report data=color nowd out=colorinfo;          |                                            */
/*                     |                                                     |                                            */
/*                     |  columns                                            |                                            */
/*                     |                                                     |                                            */
/*                     |    label                                            |                                            */
/*                     |    c2020_09 c2020_10                                |                                            */
/*                     |    _2020_08 _2020_09 _2020_10;                      |                                            */
/*                     |                                                     |                                            */
/*                     |  define label    / display;                         |                                            */
/*                     |  define _2020_08  / display;                        |                                            */
/*                     |                                                     |                                            */
/*                     |  define c2020_09 / noprint; /*-has to be first-*/   |                                            */
/*                     |  define c2020_10 / noprint; /*-has to be first-*/   |                                            */
/*                     |                                                     |                                            */
/*                     |  define _2020_09  / display;                        |                                            */
/*                     |  define _2020_10  / display;                        |                                            */
/*                     |                                                     |                                            */
/*                     |  compute _2020_09;                                  |                                            */
/*                     |    call define('_2020_09', 'style'                  |                                            */
/*                     |     ,'style=[background=' || c2020_09 || ']');      |                                            */
/*                     |  endcompute;                                        |                                            */
/*                     |                                                     |                                            */
/*                     |  compute _2020_10;                                  |                                            */
/*                     |    call define('_2020_10', 'style'                  |                                            */
/*                     |      ,'style=[background=' || c2020_10 || ']');     |                                            */
/*                     |  endcompute;                                        |                                            */
/*                     |                                                     |                                            */
/*                     | un;quit;                                            |                                            */
/*                     | ods excel close;                                    |                                            */
/*                     |                                                     |                                            */
/*                     | d:/xls/colorx.xlsx                                  |                                            */
/*                     |                                                     |                                            */
/*                     |                                                     |                                            */
/*                     | proc print data=colorinfo heading=vertical;         |                                            */
/*                     | run;quit;                                           |                                            */
/*                     |                                                     |                                            */
/*                     |--------------------------------------------------------------------------------------------------*/
/*                     | 2 DYNAMICREPORT ARRAYS                              |                                            */
/*                     | Not sold on complex analysis in report              | %put array x{*} %do_over(_cs,phrase=_c?_); */
/*                     | ======================================              |                                            */
/*                     |                                                     | array x{*} _c2_ _c3_ _c4_                  */
/*                     | %array(_cs                                          |                                            */
/*                     |  ,values=2-%eval(%utl_unqvar(have,date)+1));        | d:/xls/colorx.xlsx                         */
/*                     |                                                     |                                            */
/*                     | %put array x{*} %do_over(_cs,phrase=_c?_);          | ------------------                         */
/*                     |                                                     | | A1| fx   | TYPE |                        */
/*                     | %utlfkil(d:/xls/colorx.xlsx);                       | ----------------------------------------+  */
/*                     | ods excel file="d:/xls/colorx.xlsx"                 | [_] |   A  |    B  |   C     |  D       |  */
/*                     |  options(sheet_name="color") style=journal;         | ----------------------------------------|  */
/*                     | proc report data=have nowd out=colorx;              |  1  |TYPE |_2020_08|_2020_09 |_2920_09  |  */
/*                     |                                                     |  -- |-----+--------+---------+----------|  */
/*                     |  columns                                            |  2  |TypeA|312     |283 GREEN|253  GREEN|  */
/*                     |     ("type" type)                                   |  -- |-----+--------+---------+----------|  */
/*                     |     date                                            |  3  |TypeB|296     |310  RED |313  RED  |  */
/*                     |     ,excpt;                                         |  -- |-----+--------+---------+----------|  */
/*                     |                                                     |  4  |TypeC|545     |553  RED |590  RED  |  */
/*                     |  define type/group ' ' ;                            |  -- |-----+--------+---------+----------|  */
/*                     |  define date/across order=data ' ' ;                |  5  |TypeD|1697    |1784 RED |1813 RED  |  */
/*                     |  define excpt/analysis sum "";                      |  -- |-----+--------+---------+----------|  */
/*                     |                                                     |  6  |GRAN |2850    |2930     |2969      |  */
/*                     |  compute excpt;                                     |  -- |-----+--------+---------+----------+  */
/*                     |                                                     | [COLORX]                                   */
/*                     |   array x[*] %do_over(_cs,phrase=_c?_);             |                                            */
/*                     |                                                     |                                            */
/*                     |   if missing(_break_) then do;                      |                                            */
/*                     |    do i=2 to dim(x);                                |                                            */
/*                     |                                                     |                                            */
/*                     |     if x{i-1}<x{i} then call define(vname(x{i})     |                                            */
/*                     |        ,'style','style={background=red}');          |                                            */
/*                     |                                                     |                                            */
/*                     |     if x{i-1}=x{i} then call define(vname(x{i})     |                                            */
/*                     |        ,'style','style={background=yellow}');       |                                            */
/*                     |                                                     |                                            */
/*                     |     if x{i-1}>x{i} then call define(vname(x{i})     |                                            */
/*                     |        ,'style','style={background=green}');        |                                            */
/*                     |    end;                                             |                                            */
/*                     |                                                     |                                            */
/*                     |   end;                                              |                                            */
/*                     |                                                     |                                            */
/*                     |  endcomp;                                           |                                            */
/*                     |                                                     |                                            */
/*                     |  compute after;                                     |                                            */
/*                     |    type='GrandTotal';                               |                                            */
/*                     |  endcomp;                                           |                                            */
/*                     |  rbreak after /summarize;                           |                                            */
/*                     |                                                     |                                            */
/*                     |  run;quit;                                          |                                            */
/*                     |  ods excel close;                                   |                                            */
/**************************************************************************************************************************/

/*                   _
(_)_ __  _ __  _   _| |_
| | `_ \| `_ \| | | | __|
| | | | | |_) | |_| | |_
|_|_| |_| .__/ \__,_|\__|
        |_|
*/

data have;
 input
  date $7.
  type $
  excpt;
cards4;
2020_08 TypeA 312
2020_08 TypeB 296
2020_08 TypeC 545
2020_08 TypeD 1697
2020_09 TypeA 283
2020_09 TypeB 310
2020_09 TypeC 553
2020_09 TypeD 1784
2020_10 TypeA 253
2020_10 TypeB 313
2020_10 TypeC 590
2020_10 TypeD 1813
;;;;
run;quit;

/**************************************************************************************************************************/
/*   DATE      TYPE     EXCPT                                                                                             */
/*                                                                                                                        */
/*  2020_08    TypeA      312                                                                                             */
/*  2020_08    TypeB      296                                                                                             */
/*  2020_08    TypeC      545                                                                                             */
/*  2020_08    TypeD     1697                                                                                             */
/*  2020_09    TypeA      283                                                                                             */
/*  2020_09    TypeB      310                                                                                             */
/*  2020_09    TypeC      553                                                                                             */
/*  2020_09    TypeD     1784                                                                                             */
/*  2020_10    TypeA      253                                                                                             */
/*  2020_10    TypeB      313                                                                                             */
/*  2020_10    TypeC      590                                                                                             */
/*  2020_10    TypeD     1813                                                                                             */
/**************************************************************************************************************************/

/*   _                   _                 _          _            _
/ | | |__   __ _ _ __ __| |   ___ ___   __| | ___  __| |  ___ ___ | |_   _ _ __ ___  _ __  ___
| | | `_ \ / _` | `__/ _` |  / __/ _ \ / _` |/ _ \/ _` | / __/ _ \| | | | | `_ ` _ \| `_ \/ __|
| | | | | | (_| | | | (_| | | (_| (_) | (_| |  __/ (_| || (_| (_) | | |_| | | | | | | | | \__ \
|_| |_| |_|\__,_|_|  \__,_|  \___\___/ \__,_|\___|\__,_| \___\___/|_|\__,_|_| |_| |_|_| |_|___/
*/

proc datasets lib=work;
 delete havtab color colorinfo;
run;quit;

ods exclude all;
ods output observed=
 havtab (drop=sum);
proc corresp data=have
 observed dim=2;
tables type, date ;
weight excpt;
run;quit;
ods select all;

/*----
Label  _2020_08 _2020_09 _2020_10

TypeA     312      283      253
TypeB     296      310      313
TypeC     545      553      590
TypeD    1697     1784     1813
Sum      2850     2930     2969
----*/

data color (
 drop=idx clr1 clr2
 rename=(clr3-clr4=c2020_09-c2020_10));
 set havtab;

 array mon[*] _numeric_ ;
 array clr[%utl_varcount(havtab)] $80 ;
  do idx= 2 to dim1(mon);
  select;
   when (mon[idx]=mon[idx-1]) clr[idx+1]="YELLOW";
   when (mon[idx]>mon[idx-1]) clr[idx+1]="RED";
   when (mon[idx]<mon[idx-1]) clr[idx+1]="GREEN";
   otherwise clr[idx+1] = "BLUE";
  end;
 end;
run;quit;

/*----
                            COLOR INFO
        2020  2020 2020   C2020   C2020
 LABEL    08    09   10      09      10

 TypeA   312   283  253   GREEN   GREEN
 TypeB   296   310  313   RED     RED
 TypeC   545   553  590   RED     RED
 TypeD  1697  1784 1813   RED     RED
 Sum    2850  2930 2969   RED     RED
----*/

%utlfkil(d:/xls/color.xlsx);
ods excel file="d:/xls/color.xlsx"
    style=journal options(sheet_name="color");

proc report data=color nowd out=colorinfo;

 columns

   label
   c2020_09 c2020_10
   _2020_08 _2020_09 _2020_10;

 define label    / display;
 define _2020_08  / display;

 define c2020_09 / noprint; /*-has to be first-*/
 define c2020_10 / noprint; /*-has to be first-*/

 define _2020_09  / display;
 define _2020_10  / display;

 compute _2020_09;
   call define('_2020_09', 'style'
    ,'style=[background=' || c2020_09 || ']');
 endcompute;

 compute _2020_10;
   call define('_2020_10', 'style'
     ,'style=[background=' || c2020_10 || ']');
 endcompute;

run;quit;
ods excel close;

proc print data=colorinfo(drop=_break_);
run;quit;


/**************************************************************************************************************************/
/*  d:/xls/color.xlsx                            | WORK.COLORINFO (FROM REPORT)                                           */
/*                                               |                                                                        */
/*  ------------------                           | LABEL    C2020_09    C2020_10    _2020_08    _2020_09    _2020_10      */
/*  | A1| fx   | TYPE |                          |                                                                        */
/*  ----------------------------------------+    | TypeA     GREEN       GREEN         312         283         253        */
/*  [_] |   A  |    B  |   C     |  D       |    | TypeB     RED         RED           296         310         313        */
/*  ----------------------------------------|    | TypeC     RED         RED           545         553         590        */
/*   1  |TYPE |_2020_08|_2020_09 |_2920_09  |    | TypeD     RED         RED          1697        1784        1813        */
/*   -- |-----+--------+---------+----------|    | Sum       RED         RED          2850        2930        2969        */
/*   2  |TypeA|312     |283 GREEN|253  GREEN|    |                                                                        */
/*   -- |-----+--------+---------+----------|    |                                                                        */
/*   3  |TypeB|296     |310  RED |313  RED  |    | WORK.COLOR                                                             */
/*   -- |-----+--------+---------+----------|    |                                                                        */
/*   4  |TypeC|545     |553  RED |590  RED  |    | LABEL    _2020_08    _2020_09    _2020_10    C2020_09    C2020_10      */
/*   -- |-----+--------+---------+----------|    |                                                                        */
/*   5  |TypeD|1697    |1784 RED |1813 RED  |    | TypeA       312         283         253       GREEN       GREEN        */
/*   -- |-----+--------+---------+----------|    | TypeB       296         310         313       RED         RED          */
/*   6  |Sum  |2850    |2930 RED |2969 RED  |    | TypeC       545         553         590       RED         RED          */
/*   -- |-----+--------+---------+----------+    | TypeD      1697        1784        1813       RED         RED          */
/*  [COLOR]                                      | Sum        2850        2930        2969       RED         RED          */
/**************************************************************************************************************************/

/*___        _                             _                                 _
|___ \    __| |_   _ _ __   __ _ _ __ ___ (_) ___  _ __ ___ _ __   ___  _ __| |_    __ _ _ __ _ __ __ _ _   _
  __) |  / _` | | | | `_ \ / _` | `_ ` _ \| |/ __|| `__/ _ \ `_ \ / _ \| `__| __|  / _` | `__| `__/ _` | | | |
 / __/  | (_| | |_| | | | | (_| | | | | | | | (__ | | |  __/ |_) | (_) | |  | |_  | (_| | |  | | | (_| | |_| |
|_____|  \__,_|\__, |_| |_|\__,_|_| |_| |_|_|\___||_|  \___| .__/ \___/|_|   \__|  \__,_|_|  |_|  \__,_|\__, |
               |___/                                       |_|                                          |___/
*/
%arraydelete(_cs);

%array(_cs
 ,values=2-%eval(%utl_unqvar(have,date)+1));

%put array x{*} %do_over(_cs,phrase=_c?_);

%utlfkil(d:/xls/colorx.xlsx);
ods excel file="d:/xls/colorx.xlsx"
 options(sheet_name="color") style=journal;

proc report data=have nowd out=colorx;

 columns
    ("type" type)
    date
    ,excpt;

 define type/group ' ' ;
 define date/across order=data ' ' ;
 define excpt/analysis sum "";

 compute excpt;

  array x[*] %do_over(_cs,phrase=_c?_);

  if missing(_break_) then do;
   do i=2 to dim(x);

    if x{i-1}<x{i} then call define(vname(x{i})
       ,'style','style={background=red}');

    if x{i-1}=x{i} then call define(vname(x{i})
       ,'style','style={background=yellow}');

    if x{i-1}>x{i} then call define(vname(x{i})
       ,'style','style={background=green}');
   end;

  end;

 endcomp;

 compute after;
   type='GrandTotal';
 endcomp;
 rbreak after /summarize;

 run;quit;
 ods excel close;

/**************************************************************************************************************************/
/* %put array x{*} %do_over(_cs,phrase=_c?_);                                                                             */
/*                                                                                                                        */
/* array x{*} _c2_ _c3_ _c4_                                                                                              */
/*                                                                                                                        */
/* d:/xls/colorx.xlsx                                                                                                     */
/*                                                                                                                        */
/* ------------------                                                                                                     */
/* | A1| fx   | TYPE |                                                                                                    */
/* ----------------------------------------+                                                                              */
/* [_] |   A  |    B  |   C     |  D       |                                                                              */
/* ----------------------------------------|                                                                              */
/*  1  |TYPE |_2020_08|_2020_09 |_2920_09  |                                                                              */
/*  -- |-----+--------+---------+----------|                                                                              */
/*  2  |TypeA|312     |283 GREEN|253  GREEN|                                                                              */
/*  -- |-----+--------+---------+----------|                                                                              */
/*  3  |TypeB|296     |310  RED |313  RED  |                                                                              */
/*  -- |-----+--------+---------+----------|                                                                              */
/*  4  |TypeC|545     |553  RED |590  RED  |                                                                              */
/*  -- |-----+--------+---------+----------|                                                                              */
/*  5  |TypeD|1697    |1784 RED |1813 RED  |                                                                              */
/*  -- |-----+--------+---------+----------|                                                                              */
/*  6  |GRAN |2850    |2930     |2969      |                                                                              */
/*  -- |-----+--------+---------+----------+                                                                              */
/* [COLORX]                                                                                                               */
/**************************************************************************************************************************/

/*              _
  ___ _ __   __| |
 / _ \ `_ \ / _` |
|  __/ | | | (_| |
 \___|_| |_|\__,_|

*/
