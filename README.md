# PowerQuery_relative-path-setting
Method to set up relative path query in MS Excel PowerQuery

find any cell and use below formula in Excel to return current the path
=LEFT(CELL("filename",$A$1),FIND("[",CELL("filename",$A$1),1)-1)

![image](https://user-images.githubusercontent.com/117622597/231323028-c147bb83-dc07-42de-86d2-f936b4bc8947.png)

use name namager to name above cell as FILE_PATH

In query editor, open Advanced Editor, and replace the source code with:

let
    FILE_PATH = Excel.CurrentWorkbook(){[Name="FILE_PATH"]}[Content]{0}[Column1],
    FullPathToFile = FILE_PATH & "(filename).xlsx",
    Source = Excel.Workbook(File.Contents(FullPathToFile), null, true),
