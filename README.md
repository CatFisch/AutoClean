#AutoCleanProject


##Licence
Copyright 2020 Catharina Fischer and Alexandra Tichauer

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.


##What does the AutoClean-Script do?

The AutoClean-Script applies the "clean_skript_V3" (as a module) to excel files. 


##Requirements for using the AutoCleansktipt

1. The actual Clean-Script has to be named / renamed to "clean_skript_V3" (for its called as module).
2. The AutoClean-Script and actual Clean-Script have to be in the same directory as the files on which the scripts should run.


##How to apply the AutoClean script

#Clean all files in directory:

dir/to/excel/files$ python AutoClean.py


#Clean a list of selected files (with the --table command)

dir/to/excel/files$ python AutoClean.py --table <some_excel_file.xlsx> <another_excel_file.xlsx> <third_file.xlsx> <...>


##Collected Output

All Output files that arise running the script is collected in the same directory in a folder called "CollectedOutput". It can be deleted if not needed.



 
