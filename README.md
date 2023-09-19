# utl-drop-down-to-powershell-and-programatically-create-an-odbc-data-source-for-excel-wps-r-rodbc
utl-drop-down-to-powershell-and-programatically-create-an-odbc-data-source-for-excel-wps-r-rodbc
    %let pgm=utl-drop-down-to-powershell-and-programatically-create-an-odbc-data-source-for-excel-wps-r-rodbc;

    Drop down to powershell and programatically create an odbc data source for execl

    github
    https://tinyurl.com/mrenx557
    https://github.com/rogerjdeangelis/utl-drop-down-to-powershell-and-programatically-create-an-odbc-data-source-for-excel-wps-r-rodbc

    ODBC provides passthru to microsoft sql. We can use Microsoft SQL to manipulate an excel workbook on the excel side.
    No need to convert excel data to datasets or dataframes.

    Powershell drop down macro on end and in

    Problem

       Convert sheet 1 to a R dataframe using passthru to microsoft SQL
       You need an existing excel workbook. I ceate a sample workbook.

       I gave up trying to use just the 'dsn' in python(pyodbc)
       R seems more mature then python, very simple to use the powsershell create DSN in R.
       channel <- odbcConnect("have");

    /*                   _
    (_)_ __  _ __  _   _| |_
    | | `_ \| `_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    */

    /*---- Ceate a workbook                                                  ----*/

    %utlfkil(d:/xls/have.xlsx);

    %utl_submit_wps64x('
      libname xls excel "d:/xls/have.xlsx";
      data xls.have;
       set sashelp.zipcode(obs=6 keep=zip statecode statename);
      run;quit;
      libname xls clear;
    ');

    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /* d:/xls/have.xlsx                                                                                                       */
    /*                                                                                                                        */
    /*     +---------------------+-------------------+                                                                        */
    /*     |  A  |       B       |        C          |                                                                        */
    /*     +---------------------+-------------------+                                                                        */
    /* 1   | ZIP |  STATECODE    |     STATENAME     |                                                                        */
    /*     |-----+---------------|-------------------|                                                                        */
    /* 2   |00501|     NY        |     New York      |                                                                        */
    /*     |-----+---------------+-------------------+                                                                        */
    /* 3   |00544|     NY        |     New York      |                                                                        */
    /*     |-----+---------------|-------------------|                                                                        */
    /* 4   |00601|     PR        |    PUERTO RICO    |                                                                        */
    /*     |-----+---------------+-------------------+                                                                        */
    /* 5   |00602|     PR        |    PUERTO RICO    |                                                                        */
    /*     |-----+---------------+-------------------+                                                                        */
    /* 6   |00603|     PR        |    PUERTO RICO    |                                                                        */
    /*     -------------------------------------------                                                                        */
    /* ...                                                                                                                    */
    /*                                                                                                                        */
    /* [HAVE]                                                                                                                 */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

    /*
     _ __  _ __ ___   ___ ___  ___ ___
    | `_ \| `__/ _ \ / __/ _ \/ __/ __|
    | |_) | | | (_) | (_|  __/\__ \__ \
    | .__/|_|  \___/ \___\___||___/___/
    |_|
    */

    /*---- Be careful you need this single long line.                        ----*/
    /*---- Powersheel has very messy multiple capability with nested quotes  ----*/
    /*---- I gave up on the powershell kingon code to handle nested quotes   ----*/

    %utl_submit_ps64('
    Add-OdbcDsn -Name "have" -DriverName "Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)" -DsnType "User" -Platform "64-bit" -SetPropertyValue "Dbq=d:\xls\have.xlsx";
    Get-OdbcDsn;
    ');

    /*----  This seems to work to create multiline command                   ----*/
    %let longline=%str(Add-OdbcDsn -Name 'have' -DriverName 'Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)'
     -DsnType 'User' -Platform '64-bit' -SetPropertyValue 'Dbq=d:\xls\have.xlsx');

    options ls=255;
    %put &=longline;

    %utl_submit_ps64("
    &longline;
    Get-OdbcDsn;
    ");


    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /*  From the LOG DSN 'HAVE'                                                                                               */
    /*                                                                                                                        */
    /*  Name       : have                                                                                                     */
    /*  DsnType    : User                                                                                                     */
    /*  Platform   : 64-bit                                                                                                   */
    /*  DriverName : Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)                                                   */
    /*  Attribute  : {DBQ, DriverId, ImplicitCommitSync, Threads...}                                                          */
    /*                                                                                                                        */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

    /*___
    |  _ \   _   _ ___  __ _  __ _  ___
    | |_) | | | | / __|/ _` |/ _` |/ _ \
    |  _ <  | |_| \__ \ (_| | (_| |  __/
    |_| \_\  \__,_|___/\__,_|\__, |\___|
                             |___/
    */

    %utl_submit_wps64x('
    proc r;
    submit;
    library(RODBC);
    ch <- odbcConnect("have");
    sqlResult <- sqlQuery(ch, "select * from have";);
    sqlResult;
    odbcClose(ch);
    endsubmit;
    ');

    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /* The WPS System                                                                                                         */
    /*                                                                                                                        */
    /*   ZIP STATECODE   STATENAME                                                                                            */
    /* 1 501        NY    New York                                                                                            */
    /* 2 544        NY    New York                                                                                            */
    /* 3 601        PR Puerto Rico                                                                                            */
    /* 4 602        PR Puerto Rico                                                                                            */
    /* 5 603        PR Puerto Rico                                                                                            */
    /* 6 604        PR Puerto Rico                                                                                            */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

    /*                                _          _ _
     _ __   _____      _____ _ __ ___| |__   ___| | |  _ __ ___   __ _  ___ _ __ ___
    | `_ \ / _ \ \ /\ / / _ \ `__/ __| `_ \ / _ \ | | | `_ ` _ \ / _` |/ __| `__/ _ \
    | |_) | (_) \ V  V /  __/ |  \__ \ | | |  __/ | | | | | | | | (_| | (__| | | (_) |
    | .__/ \___/ \_/\_/ \___|_|  |___/_| |_|\___|_|_| |_| |_| |_|\__,_|\___|_|  \___/
    |_|

    */

    %macro utl_submit_ps64(
          pgm
         ,return=  /* name for the macro variable from Powershell */
         )/des="Semi colon separated set of python commands - drop down to python";


      /*
          %let pgm='Get-Content -Path d:/txt/back.txt | Measure-Object -Line | clip;';
      */

      * write the program to a temporary file;
      filename py_pgm "%sysfunc(pathname(work))/py_pgm.ps1" lrecl=32766 recfm=v;
      data _null_;
        length pgm  $32755 cmd $1024;
        file py_pgm ;
        pgm=&pgm;
        semi=countc(pgm,';');
          do idx=1 to semi;
            cmd=cats(scan(pgm,idx,';'));
            if cmd=:'. ' then
               cmd=trim(substr(cmd,2));
             put cmd $char384.;
             putlog cmd $char384.;
          end;
      run;quit;
      %let _loc=%sysfunc(pathname(py_pgm));
      %put &_loc;
      filename rut pipe  "powershell.exe -executionpolicy bypass -file &_loc ";
      data _null_;
        file print;
        infile rut;
        input;
        put _infile_;
        putlog _infile_;
      run;
      filename rut clear;
      filename py_pgm clear;

      * use the clipboard to create macro variable;
      %if "&return" ^= "" %then %do;
        filename clp clipbrd ;
        data _null_;
         length txt $200;
         infile clp;
         input;
         putlog "*******  " _infile_;
         call symputx("&return",_infile_,"G");
        run;quit;
      %end;

    %mend utl_submit_ps64;

    /*              _
      ___ _ __   __| |
     / _ \ `_ \ / _` |
    |  __/ | | | (_| |
     \___|_| |_|\__,_|

    */
