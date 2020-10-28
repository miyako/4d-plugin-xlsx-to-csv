/* --------------------------------------------------------------------------------
 #
 #	4DPlugin-xlnt-csv.h
 #	source generated by 4D Plugin Wizard
 #	Project : xlnt-csv
 #	author : miyako
 #	2020/10/02
 #  
 # --------------------------------------------------------------------------------*/

#ifndef PLUGIN_XLNT_CSV_H
#define PLUGIN_XLNT_CSV_H

#include "4DPluginAPI.h"
#include "xlnt/xlnt.hpp"

#include "xlsxio_read.h"
#include "xlsxio_write.h"

#include "ARRAY_TEXT.h"
#include "C_TEXT.h"

#include <time.h>
#include <stdio.h>
#include <stdlib.h>

#if VERSIONWIN
#include <process.h>
#endif

#pragma mark -

void convert_xlsx_to_csv(PA_PluginParameters params);
void convert_xlsx_to_csv_v2(PA_PluginParameters params);
void get_xlsx_sheets(PA_PluginParameters params);

#endif /* PLUGIN_XLNT_CSV_H */
