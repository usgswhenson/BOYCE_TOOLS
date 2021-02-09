@Echo off
set path=%cd%\Code\VirtualEnv\Python27_64\;%path%
:: Check WMIC is available
WMIC.EXE Alias /? >NUL 2>&1 || GOTO s_error

:: Use WMIC to retrieve date and time
FOR /F "skip=1 tokens=1-6" %%G IN ('WMIC Path Win32_LocalTime Get Day^,Hour^,Minute^,Month^,Second^,Year /Format:table') DO (
   IF "%%~L"=="" goto s_done
      Set _yyyy=%%L
      Set _mm=00%%J
      Set _dd=00%%G
      Set _hour=00%%H
      SET _minute=00%%I
)
::s_done

:: Pad digits with leading zeros
      Set _mm=%_mm:~-2%
      Set _dd=%_dd:~-2%
      Set _hour=%_hour:~-2%
      Set _minute=%_minute:~-2%

:: get tag and date
set "_calVar=KCOFEV2playsfrical"
set _date=%_mm%-%_dd%-%_yyyy%
Set _isodate=%_mm%-%_dd%-%_yyyy%%_calvar%


::::##################################################################################

::
::HEADS
echo HEADS
echo. | python.exe .\Code\Scripts\Hob_To_Hydrograph_Excel_Chart.py    .\OUT_TR\HOB.out   .\Simulated_Analysis\4__Heads\Hob_To_Hydrograph_SVIHM.xlsx    Hob_Script_Input          HOB_Data   -999      False     Hydrograph   Correlation
copy /Y .\OUT_TR\Analysis\4__Heads\Hob_To_Hydrograph_SVIHM.xlsx .\OUT_TR\Analysis\4__Heads\Hob_To_Hydrograph_SVIHM_%_isodate%.xlsx

echo HEADS subset
echo. | python.exe .\Code\Scripts\Hob_To_Hydrograph_Excel_Chart.py    .\OUT_TR\HOB.out   .\OUT_TR\Analysis\4__Heads\Hob_To_Hydrograph_SVIHM_subset.xlsx    Hob_Script_Input          HOB_Data   -999      False     Hydrograph   Correlation
copy /Y .\OUT_TR\Analysis\4__Heads\Hob_To_Hydrograph_SVIHM_subset.xlsx .\OUT_TR\Analysis\4__Heads\Hob_To_Hydrograph_SVIHM_subset_%_isodate%.xlsx

::HEAD DIFFERENCES
echo HEAD DIFFERENCES
echo. | python.exe .\Code\Scripts\Hob_To_Hydrograph_Excel_Chart.py    .\5_pst_SA\Observations\SVIHM_DIFF.sim   .\OUT_TR\Analysis\4__Heads\HeadDifferences\Hob_Head_Difference_SVIHM.xlsx    Chart_Input      HOB_Data   -999     False     Hydrograph   Correlation
copy /Y .\OUT_TR\Analysis\4__Heads\HeadDifferences\Hob_Head_Difference_SVIHM.xlsx .\OUT_TR\Analysis\4__Heads\HeadDifferences\Hob_Head_Difference_SVIHM_%_isodate%.xlsx

::STREAMS
echo STREAMS
echo. | python.exe .\Code\Scripts\Hob_To_Hydrograph_Excel_Chart.py    .\5_pst_SA\Observations\SFR_OBSHOB.sim   .\OUT_TR\Analysis\5__Surfacewater\SFR_To_Hydrograph_SVIHM.xlsx    SFR_Script_Input          SFR_Data   -999      False     Hydrograph   Correlation True
copy /Y .\OUT_TR\Analysis\5__Surfacewater\SFR_To_Hydrograph_SVIHM.xlsx .\OUT_TR\Analysis\5__Surfacewater\SFR_To_Hydrograph_SVIHM_%_isodate%.xlsx

::SFR DB DEG
echo DB DEG
python.exe .\Code\Scripts\SFR_DB_SEG.py  .\OUT_TR\Analysis\5__Surfacewater\SFR_SVIHM_DB_SEG.xlsx  Input  YES .\OUT_TR\SFR_out_DB.txt
copy /Y .\OUT_TR\Analysis\5__Surfacewater\SFR_SVIHM_DB_SEG.xlsx .\OUT_TR\Analysis\5__Surfacewater\SFR_SVIHM_DB_SEG_%_isodate%.xlsx

::::ZONE BUDGET
::echo ZONE BUDGET
::..\OUT_TR\Analysis\3__batch\go_zbdgt4_31_auto.bat
::
:::: Upper-Layer-Specific Zonebudget Analysis => All uppermost cells within each layer for all WBS
::..\OUT_TR\Analysis\1__exe\ZoneBudget.exe < .\OUT_TR\Analysis\6__WaterBudget\ZoneBudget\SVIHM_UpLyzone.in > zonbdgt_uply.out
::echo SFR GAINS_LOSSES
::echo. | python.exe .\OUT_TR\Analysis\2__python_Scripts\StreamRechargeAverage.py    .\OUT_TR\sfrout.txt   .\OUT_TR\Analysis\10_GIS\Salinas_SFR.mxd    .\OUT_TR\Analysis\10_GIS\Salinas_SFR.mxd\SFR_cell.shp     NewAverage%_calvar%.bmp
::
::echo SGMApy Budget
::"C:\anaconda3\envs\sgmapy2\python.exe" .\OUT_TR\Analysis\2__python_Scripts\SalinasBudget.py  %_calVar% %_date%    
pause
